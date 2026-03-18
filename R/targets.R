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

          base   <- min(6L, avail + credited)
          target <- if (avail >= 6) 6L else base
          sched_target <- max(0L, target - credited)

          list(
            pp_name      = pp_name,
            avail        = avail,
            credited     = credited,
            target       = target,
            sched_target = sched_target,
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
        stringsAsFactors = FALSE
      )
    })
  })
  do.call(rbind, unlist(rows, recursive = FALSE))
}
