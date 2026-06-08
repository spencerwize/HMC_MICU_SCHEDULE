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

          off_days <- pdata$date[pdata$type == "off"]
          vac_days <- pdata$date[pdata$type == "vac"]
          cme_days <- pdata$date[pdata$type == "cme"]

          # Intersect with this PP's date range
          off_days <- off_days[off_days %in% p_dates]
          vac_days <- vac_days[vac_days %in% p_dates]
          cme_days <- cme_days[cme_days %in% p_dates]


          credited <- length(cme_days)
          all_off  <- unique(c(off_days, vac_days, cme_days))
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
          n_offvac   <- length(unique(c(off_days, vac_days)))
          pto_needed <- pto_reduction(n_offvac)

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
        stringsAsFactors = FALSE
      )
    })
  })
  do.call(rbind, unlist(rows, recursive = FALSE))
}

#' Per-PP staffing balance: slot capacity vs shift demand vs available person-days
staffing_balance_df <- function(targets) {
  do.call(rbind, lapply(seq_len(nrow(PAY_PERIODS)), function(i) {
    pp_name  <- PAY_PERIODS$name[i]
    pp_start <- PAY_PERIODS$start[i]
    pp_end   <- PAY_PERIODS$end[i]
    n_days   <- as.integer(pp_end - pp_start + 1L)
    capacity <- n_days * 4L

    demand     <- sum(vapply(STAFF, function(p) targets[[p]][[pp_name]]$sched_target, integer(1L)))
    staff_days <- sum(vapply(STAFF, function(p) targets[[p]][[pp_name]]$avail,        integer(1L)))
    slack      <- staff_days - demand

    data.frame(
      PP         = pp_name,
      Days       = n_days,
      Capacity   = capacity,
      Demand     = demand,
      Staff_Days = staff_days,
      Slack      = slack,
      Status     = if (slack < 0L) "IMPOSSIBLE" else if (slack <= 5L) "TIGHT" else "OK",
      stringsAsFactors = FALSE
    )
  }))
}
