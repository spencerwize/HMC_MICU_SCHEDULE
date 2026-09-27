# ─────────────────────────────────────────────────────────────────────────────
# targets.R  —  Compute per-person per-PP shift targets
#
# Returns: named list  person -> named list  pp_name -> list(
#   avail, credited, target, sched_target,
#   off_days, vac_days, cme_days, pp_dates
# )
# ─────────────────────────────────────────────────────────────────────────────

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

          bt <- BASE_TARGETS[[person]]
          base_target <- if (is.null(bt)) {
            6L
          } else if (is.list(bt)) {
            if (!is.null(bt[[pp_name]])) as.integer(bt[[pp_name]]) else 6L
          } else {
            as.integer(bt)
          }
          # ── Hard target = base − CME − PTO-reduction(requested time off) ──────
          # The target is FIRM (no soft band).  Requested off+vac days in the PP
          # lower the target by a fixed amount, and that SAME amount is the PTO the
          # person needs (tracked for reporting, never scheduled):
          #   off+vac in PP:  ≤4 → 0   5-6 → 1   7-8 → 2   9-10 → 3
          #                   11 → 4   12 → 5    13-14 → 6
          # (The stated ranges overlap at 10; resolved to 3 so the 5-6/7-8/9-10
          #  pairs stay regular.)
          n_requested_off <- length(off_days) + length(vac_days)
          pto_needed <- if (n_requested_off >= 13L) 6L
                        else if (n_requested_off == 12L) 5L
                        else if (n_requested_off == 11L) 4L
                        else if (n_requested_off >=  9L) 3L
                        else if (n_requested_off >=  7L) 2L
                        else if (n_requested_off >=  5L) 1L
                        else 0L

          # base_target is 6 by default (BASE_TARGETS may override per person/PP).
          # Clamp to available work days so the firm target is never unachievable.
          target       <- max(0L, base_target - credited - pto_needed)
          target       <- min(target, avail)
          sched_target <- target

          # Firm target → floor == ceiling.  FLEX staff (e.g. Todd) keep their own
          # lower flexibility floor.
          flex_floor <- FLEX_TARGETS[[person]]
          soft_min   <- if (!is.null(flex_floor))
                          max(0L, min(as.integer(flex_floor), target))
                        else
                          target

          list(
            pp_name      = pp_name,
            avail        = avail,
            credited     = credited,
            target       = target,
            pto_needed   = pto_needed,   # PTO needed this PP (tracked, NOT scheduled)
            sched_target = sched_target,
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
        pto_needed   = info$pto_needed,
        sched_target = info$sched_target,
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
