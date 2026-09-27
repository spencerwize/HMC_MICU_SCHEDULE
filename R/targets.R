# ─────────────────────────────────────────────────────────────────────────────
# targets.R  —  Compute per-person per-PP shift targets
#
# Returns: named list  person -> named list  pp_name -> list(
#   avail, credited, target, sched_target, pto_needed,
#   off_days, vac_days, cme_days, pp_dates
# )
# ─────────────────────────────────────────────────────────────────────────────

# Map the number of OFF + VAC days requested in a pay period to the target
# reduction (and equivalently the PTO days needed to backfill the shortfall).
# CME days are excluded here — they reduce the base target separately
# (base = 6 - CME) and are credited, not PTO.
#
#   off/vac days   target ↓ / PTO needed
#   ────────────   ─────────────────────
#   0–4                 0
#   5–6                 1
#   7–8                 2
#   9–10                3
#   11–12               4
#   13                  5
#   14                  6
pto_reduction <- function(n_offvac) {
  if (n_offvac <= 4L)  return(0L)
  if (n_offvac <= 6L)  return(1L)
  if (n_offvac <= 8L)  return(2L)
  if (n_offvac <= 10L) return(3L)
  if (n_offvac <= 12L) return(4L)
  if (n_offvac <= 13L) return(5L)
  6L
}

compute_targets <- function(time_off) {
  targets <- setNames(
    lapply(STAFF, function(person) {
      setNames(
        lapply(seq_len(nrow(PAY_PERIODS)), function(i) {
          pp_name  <- PAY_PERIODS$name[i]
          pp_start <- PAY_PERIODS$start[i]
          pp_end   <- PAY_PERIODS$end[i]
          p_dates  <- seq(pp_start, pp_end, by = "day")

          pdata    <- time_off[[person]]

          off_days   <- pdata$date[pdata$type == "off"]
          vac_days   <- pdata$date[pdata$type == "vac"]
          cme_days   <- pdata$date[pdata$type == "cme"]
          pto_days   <- pdata$date[pdata$type == "pto"]
          green_days <- pdata$date[pdata$type == "green"]

          # Intersect with this PP's date range
          off_days   <- off_days[off_days     %in% p_dates]
          vac_days   <- vac_days[vac_days     %in% p_dates]
          cme_days   <- cme_days[cme_days     %in% p_dates]
          pto_days   <- pto_days[pto_days     %in% p_dates]
          green_days <- green_days[green_days %in% p_dates]


          credited <- length(cme_days)
          # NOTE: green_days are deliberately absent from all_off / avail and from
          # n_offvac below. A requested-WORK day is schedulable and must not move
          # sched_target - green governs WHICH days are worked, not HOW MANY.
          all_off  <- unique(c(off_days, vac_days, cme_days, pto_days))
          avail    <- sum(!p_dates %in% all_off)

          # Base target per PP (default 6, optionally overridden per person/PP).
          bt <- BASE_TARGETS[[person]]
          base_target <- if (is.null(bt)) {
            6L
          } else if (is.list(bt)) {
            if (!is.null(bt[[pp_name]])) as.integer(bt[[pp_name]]) else 6L
          } else {
            as.integer(bt)
          }

          # Target adjustment: the number of OFF + VAC days requested in this PP
          # bumps the target down and defines how many PTO days are needed to
          # backfill the resulting shortfall (see pto_reduction()).  CME days are
          # handled separately via `credited` (base = 6 - CME).
          # Explicit PTO days are days off too, so they feed the automatic formula
          # alongside off/vac. The period's PTO is then the LARGER of what was
          # explicitly requested and what the formula would grant on its own -
          # so marking PTO never yields fewer PTO days than today's behaviour,
          # and marking none leaves that behaviour exactly as it was.
          n_offvac   <- length(unique(c(off_days, vac_days, pto_days)))
          pto_auto   <- pto_reduction(n_offvac)
          pto_needed <- max(length(pto_days), pto_auto)

          # `target`       — adjusted shift target before CME credit
          # `sched_target` — actual shifts the solver/fill should assign
          #                  ( = 6 - CME - pto_reduction, floored at 0 )
          target       <- max(0L, base_target - pto_needed)
          sched_target <- max(0L, target - credited)

          # soft_min: scheduling urgency drops once this floor is reached.
          # The person can still receive up to sched_target shifts; they are
          # simply deprioritised relative to people below their own soft_min.
          flex_floor <- FLEX_TARGETS[[person]]
          soft_min   <- if (!is.null(flex_floor)) {
                          max(0L, min(as.integer(flex_floor), sched_target))
                        } else {
                          heavy_off <- (length(off_days) + length(vac_days)) >= 5L
                          floor_val <- if (heavy_off) 4L else DEFAULT_SOFT_MIN
                          max(0L, min(floor_val, sched_target))
                        }

          list(
            pp_name      = pp_name,
            avail        = avail,
            credited     = credited,
            target       = target,
            sched_target = sched_target,
            pto_needed   = pto_needed,
            soft_min     = soft_min,
            off_days     = off_days,
            vac_days     = vac_days,
            cme_days     = cme_days,
            pto_days     = pto_days,
            pto_auto     = pto_auto,      # what the formula alone would grant
            green_days   = green_days,
            pp_dates     = p_dates
          )
        }),
        PAY_PERIODS$name
      )
    }),
    STAFF
  )
  targets
}

#' Summarise targets as a tidy data.frame for display
targets_summary_df <- function(targets) {
  rows <- lapply(STAFF, function(person) {
    lapply(PAY_PERIODS$name, function(pp) {
      info <- targets[[person]][[pp]]
      data.frame(
        person       = person,
        pp           = pp,
        avail        = info$avail,
        credited     = info$credited,
        target       = info$target,
        sched_target = info$sched_target,
        pto_needed   = info$pto_needed,
        soft_min     = info$soft_min,
        green        = length(info$green_days),
        stringsAsFactors = FALSE
      )
    })
  })
  do.call(rbind, unlist(rows, recursive = FALSE))
}

# ─────────────────────────────────────────────────────────────────────────────
# green_supply_report()  —  Will the green-only phase have enough to work with?
#
# The green-only phase can only draw on requested-work days, so a person who
# marks fewer greens than their pay-period target simply cannot be fully
# scheduled there; the shortfall lands in the fill phase as yellow days. This
# reports that BEFORE the solver runs, which is also the second independent
# signal that a work keyword was mistyped (a mistyped green becomes an OFF day,
# so greens go down and the target goes down too).
#
# Returns a data.frame invisibly and messages a human-readable summary.
# ─────────────────────────────────────────────────────────────────────────────
green_supply_report <- function(time_off, targets) {
  rows <- do.call(rbind, lapply(STAFF, function(person) {
    do.call(rbind, lapply(PAY_PERIODS$name, function(pp) {
      info <- targets[[person]][[pp]]
      data.frame(person = person, pp = pp,
                 green = length(info$green_days),
                 sched_target = info$sched_target,
                 short = max(0L, info$sched_target - length(info$green_days)),
                 stringsAsFactors = FALSE)
    }))
  }))

  message("Green supply vs pay-period targets:")
  tot_green  <- sum(rows$green)
  tot_target <- sum(rows$sched_target)
  message(sprintf("  %d green requests vs %d target shifts across %d person-PPs.",
                  tot_green, tot_target, nrow(rows)))

  bad <- rows[rows$short > 0L, ]
  if (nrow(bad) == 0L) {
    message("  Every person has at least as many green days as their PP target.")
  } else {
    message(sprintf("  %d person-PP(s) have FEWER green days than their target -",
                    nrow(bad)))
    message("  those shifts can only be filled on yellow days:")
    for (i in seq_len(nrow(bad)))
      message(sprintf("    %-10s %-5s green=%d target=%d (short %d)",
                      bad$person[i], bad$pp[i], bad$green[i],
                      bad$sched_target[i], bad$short[i]))
  }

  # Schedule-wide floors the green-only phase is NOT allowed to enforce, but the
  # fill phase is. Flagging them here explains where yellow days will come from.
  wknd_all <- all_dates()[is_weekend(all_dates())]
  for (person in STAFF) {
    g <- unlist(lapply(PAY_PERIODS$name,
                       function(pp) as.character(targets[[person]][[pp]]$green_days)))
    n_g   <- length(g)
    n_gwk <- sum(as.Date(g) %in% wknd_all)
    if (n_gwk < MIN_WKND_HARD)
      message(sprintf("  NOTE %-10s only %d green weekend day(s); floor is %d -> expect yellow weekends.",
                      person, n_gwk, MIN_WKND_HARD))
    if (n_g < MIN_NIGHTS_HARD)
      message(sprintf("  NOTE %-10s only %d green day(s) total; night floor is %d.",
                      person, n_g, MIN_NIGHTS_HARD))
  }
  invisible(rows)
}
