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
          avail_for_target <- sum(!p_dates %in% unique(c(off_days, cme_days)))
          base   <- min(base_target, avail_for_target + credited)
          target <- if (avail_for_target >= base_target) base_target else base
          sched_target <- max(0L, target - credited)

          tgt0 <- max(0L, target - credited)   # base target = 6 − CME (availability-limited)

          # ── Relaxed band for heavy requested time off (replaces PTO) ──────────
          # Heavy requested time off lowers the FLOOR and shrinks the ceiling.  The
          # person AIMS for the base target but may settle as low as (base − z); no
          # PTO is charged for landing anywhere in that band.  (Todd-style: aim
          # high, OK lower.)
          #   z = 1 when requested (off + vac) days in the PP are in [5, 7]
          #   z = 2 when requested (off + vac) days in the PP are > 7
          #   z = 0 otherwise
          # Resulting band (minus CME):  z0 → 6..6 (firm)   z1 → 5..6   z2 → 4..5
          n_requested_off <- length(off_days) + length(vac_days)
          relaxed_by <- if (n_requested_off > 7L) 2L
                        else if (n_requested_off >= 5L) 1L
                        else 0L

          relaxed_floor <- max(0L, tgt0 - relaxed_by)        # min shifts (soft floor)
          sched_target  <- min(tgt0, relaxed_floor + 1L)     # ceiling: floor+1, capped at base

          # soft_min is the hard floor the LP must meet; sched_target is the ceiling
          # it aims for.  FLEX staff (e.g. Todd) keep their own lower flex floor.
          flex_floor <- FLEX_TARGETS[[person]]
          soft_min   <- if (!is.null(flex_floor))
                          max(0L, min(as.integer(flex_floor), sched_target))
                        else
                          relaxed_floor

          list(
            pp_name      = pp_name,
            avail        = avail,
            credited     = credited,
            target       = target,
            relaxed_by   = relaxed_by,   # shifts removed from target for heavy time off
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
        relaxed_by   = info$relaxed_by,
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
