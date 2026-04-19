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

          # VAC days block scheduling (person can't work them) but do NOT reduce
          # the 6-shift target.  Only hard-OFF days and CME days affect the target.
          # If a person can't reach 6 due to too many VAC days, cleanup_pass will
          # grant PTO credits to cover the shortfall.
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

          # soft_min: scheduling urgency drops once this floor is reached.
          # The person can still receive up to sched_target shifts; they are
          # simply deprioritised relative to people below their own soft_min.
          flex_floor <- FLEX_TARGETS[[person]]
          soft_min   <- if (!is.null(flex_floor))
                          max(0L, min(as.integer(flex_floor), sched_target))
                        else
                          max(0L, min(DEFAULT_SOFT_MIN, sched_target))

          list(
            pp_name      = pp_name,
            avail        = avail,
            credited     = credited,
            target       = target,
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
        sched_target = info$sched_target,
        soft_min     = info$soft_min,
        stringsAsFactors = FALSE
      )
    })
  })
  do.call(rbind, unlist(rows, recursive = FALSE))
}
