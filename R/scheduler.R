# ─────────────────────────────────────────────────────────────────────────────
# scheduler.R  —  R6 Scheduler class
#
# Hard constraints enforced (never violated):
#   1. APP1 always filled
#   2. Every night staffed
#   3. No day-shift morning after night (d+1 after night d)
#   4. Night recovery: blocked d+1 and d+2 after LAST night of streak (d+3 eligible)
#   5. Max 3 consecutive nights
#   6. Max 4 consecutive working days
#   7. No day-to-night same calendar day
#   8. PP shift targets respected
# ─────────────────────────────────────────────────────────────────────────────

Scheduler <- R6::R6Class("Scheduler",
  public = list(

    time_off      = NULL,   # person -> data.frame(date, type)
    targets       = NULL,   # person -> pp -> list(...)
    dates         = NULL,   # Date vector: all schedule dates

    # schedule: named list  "YYYY-MM-DD" -> list(APP1, APP2, Roaming, Night)
    schedule      = NULL,

    # per-person night dates (Date vectors) and day-shift records
    person_nights = NULL,   # person -> Date vector
    person_shifts = NULL,   # person -> data.frame(date, slot)

    # PP shift counts (includes pre-seeded holidays + CME credited)
    pp_counts     = NULL,   # person -> named int vector (pp -> count)

    # ── Initialise ────────────────────────────────────────────────────────────
    initialize = function(time_off, targets) {
      self$time_off <- time_off
      self$targets  <- targets
      self$dates    <- all_dates()

      empty_slot <- list(APP1 = NA_character_, APP2 = NA_character_,
                         Roaming = NA_character_, Night = NA_character_)
      self$schedule <- setNames(
        lapply(self$dates, function(d) empty_slot),
        as.character(self$dates)
      )

      self$person_nights <- setNames(
        lapply(STAFF, function(p) as.Date(character())), STAFF)

      self$person_shifts <- setNames(
        lapply(STAFF, function(p)
          data.frame(date = as.Date(character()), slot = character(),
                     stringsAsFactors = FALSE)),
        STAFF)

      zero_pp <- setNames(integer(nrow(PAY_PERIODS)), PAY_PERIODS$name)
      self$pp_counts <- setNames(
        lapply(STAFF, function(p) zero_pp), STAFF)
    },

    # ── Low-level getters / setters ───────────────────────────────────────────
    get_slot = function(d, slot) {
      v <- self$schedule[[as.character(d)]][[slot]]
      if (is.null(v)) NA_character_ else v
    },

    set_slot = function(d, slot, person) {
      self$schedule[[as.character(d)]][[slot]] <- person
    },

    # ── Add a shift (updates all tracking state) ──────────────────────────────
    add_shift = function(person, d, slot) {
      self$set_slot(d, slot, person)
      pp <- get_pp(d)
      if (!is.na(pp)) {
        self$pp_counts[[person]][[pp]] <-
          self$pp_counts[[person]][[pp]] + 1L
      }
      if (slot == "Night") {
        self$person_nights[[person]] <-
          sort(c(self$person_nights[[person]], d))
      } else {
        self$person_shifts[[person]] <- rbind(
          self$person_shifts[[person]],
          data.frame(date = d, slot = slot, stringsAsFactors = FALSE)
        )
      }
    },

    # ── Eligibility helpers ───────────────────────────────────────────────────

    #' TRUE if person cannot work at all on date d (off/cme/conference)
    is_blocked = function(person, d) {
      pdata <- self$time_off[[person]]
      if (nrow(pdata) > 0) {
        m <- pdata[pdata$date == d, ]
        if (nrow(m) > 0 && m$type[1] %in% c("off", "cme")) return(TRUE)
      }
      if (person %in% PP13_CONFERENCE &&
          d >= PP13_CONF_START && d <= PP13_CONF_END) return(TRUE)
      FALSE
    },

    #' Night recovery check.
    #' for_night=TRUE  → allow streak continuation (last night == d-1 is OK)
    #' for_night=FALSE → strict: any night within 2 days blocks day shifts
    night_recovery_blocked = function(person, d, for_night = FALSE) {
      past <- self$person_nights[[person]]
      past <- past[past < d]
      if (length(past) == 0L) return(FALSE)
      last_night <- max(past)
      # Streak continuation: if we worked last night, we can keep going
      if (for_night && last_night == d - 1L) return(FALSE)
      d <= last_night + 2L
    },

    had_night_on = function(person, d) {
      d %in% self$person_nights[[person]]
    },

    already_assigned = function(person, d) {
      day_data <- self$schedule[[as.character(d)]]
      any(sapply(SLOTS, function(s) {
        v <- day_data[[s]]
        !is.na(v) && v == person
      }))
    },

    # ── Consecutive-count helpers ─────────────────────────────────────────────

    consec_nights_if_added = function(person, d) {
      nights <- self$person_nights[[person]]
      count  <- 1L
      prev   <- d - 1L
      while (prev %in% nights) { count <- count + 1L; prev <- prev - 1L }
      count
    },

    consec_days_if_added = function(person, d) {
      worked <- c(self$person_shifts[[person]]$date,
                  self$person_nights[[person]])
      count  <- 1L
      prev   <- d - 1L
      while (prev %in% worked) { count <- count + 1L; prev <- prev - 1L }
      count
    },

    # ── Eligibility predicates ────────────────────────────────────────────────

    can_work_night = function(person, d) {
      if (self$is_blocked(person, d))                                 return(FALSE)
      if (self$night_recovery_blocked(person, d, for_night = TRUE))   return(FALSE)
      ds       <- as.character(d)
      day_data <- self$schedule[[ds]]
      # No day-to-night same day
      for (s in DAY_SLOTS) {
        v <- day_data[[s]]
        if (!is.na(v) && v == person) return(FALSE)
      }
      # Not already on night
      if (!is.na(day_data$Night) && day_data$Night == person) return(FALSE)
      # Not pre-assigned to day shift on D+1 or D+2 (holiday guard + recovery)
      for (offset in 1:2) {
        fwd_ds <- as.character(d + offset)
        if (!is.null(self$schedule[[fwd_ds]])) {
          fwd <- self$schedule[[fwd_ds]]
          for (s in DAY_SLOTS) {
            v <- fwd[[s]]
            if (!is.na(v) && v == person) return(FALSE)
          }
        }
      }
      if (self$consec_nights_if_added(person, d) > 3L) return(FALSE)
      if (self$consec_days_if_added(person, d)  > 4L) return(FALSE)
      TRUE
    },

    can_work_day = function(person, d) {
      if (self$is_blocked(person, d))                                 return(FALSE)
      if (self$night_recovery_blocked(person, d, for_night = FALSE))  return(FALSE)
      if (self$already_assigned(person, d))                           return(FALSE)
      if (self$consec_days_if_added(person, d) > 4L)                  return(FALSE)
      # Not assigned night tonight
      v <- self$schedule[[as.character(d)]]$Night
      if (!is.na(v) && v == person) return(FALSE)
      TRUE
    },

    # ── Scoring ───────────────────────────────────────────────────────────────

    #' Night score — lower is preferred.
    #' Returns named list(priority, total, pp) for lexicographic sort.
    night_score = function(person, d) {
      nights        <- self$person_nights[[person]]
      is_cont       <- (d - 1L) %in% nights
      pp            <- get_pp(d)
      pp_nights     <- if (!is.na(pp)) sum(get_pp_vec(nights) == pp, na.rm = TRUE) else 0L
      list(
        priority = if (is_cont) 0L else 1L,
        total    = length(nights),
        pp       = pp_nights
      )
    },

    #' Day score — higher is preferred.
    day_score = function(person, d, slot) {
      pp      <- get_pp(d)
      pp_info <- self$targets[[person]][[pp]]

      days_left  <- sum(pp_info$pp_dates >= d)
      slots_left <- max(0L, pp_info$sched_target - self$pp_counts[[person]][[pp]])
      urgency    <- slots_left / max(days_left, 1L)

      # Cluster score
      worked      <- c(self$person_shifts[[person]]$date,
                       self$person_nights[[person]])
      prev_worked <- (d - 1L) %in% worked
      next_worked <- (d + 1L) %in% worked
      cluster     <- if (prev_worked && next_worked) 3L
                     else if (prev_worked) 2L
                     else if (next_worked) 1L
                     else -3L

      # Roaming balance
      shifts      <- self$person_shifts[[person]]
      n_total     <- nrow(shifts)
      n_roam      <- if (n_total > 0) sum(shifts$slot == "Roaming") else 0L
      roam_score  <- if (slot == "Roaming" && n_total > 0) -n_roam / n_total else 0

      # Weekend fairness
      worked_d    <- shifts$date
      n_wknd      <- if (length(worked_d) > 0)
                       sum(is_weekend(worked_d)) else 0L
      wknd_ratio  <- if (length(worked_d) > 0) n_wknd / length(worked_d) else 0
      wknd_score  <- if (is_weekend(d)) -wknd_ratio else 0

      urgency * 10 + cluster + roam_score + wknd_score
    },

    # ── Algorithm phases ──────────────────────────────────────────────────────

    preseed_holidays = function() {
      for (ds in names(HOLIDAYS)) {
        d <- as.Date(ds)
        if (d < SCHEDULE_START || d > SCHEDULE_END) next
        for (slot in SLOTS) {
          person <- HOLIDAYS[[ds]][[slot]]
          if (!is.null(person)) self$add_shift(person, d, slot)
        }
      }
    },

    schedule_nights = function() {
      unscheduled <- self$dates[
        sapply(self$dates, function(d)
          is.na(self$schedule[[as.character(d)]]$Night))]

      # Sort by number of eligible candidates (hardest = fewest eligible first)
      difficulty <- sapply(unscheduled, function(d)
        sum(sapply(STAFF, function(p) self$can_work_night(p, d))))

      dates_sorted <- unscheduled[order(difficulty, unscheduled)]

      for (d in dates_sorted) {
        # Skip if already filled (holiday pre-seed)
        if (!is.na(self$schedule[[as.character(d)]]$Night)) next

        eligible <- STAFF[sapply(STAFF, function(p) self$can_work_night(p, d))]
        pp       <- get_pp(d)

        if (length(eligible) == 0L) {
          # Force-fill: anyone not blocked by off/cme and not just off last night
          eligible <- STAFF[sapply(STAFF, function(p)
            !self$is_blocked(p, d) && !self$had_night_on(p, d - 1L))]
          if (length(eligible) == 0L) eligible <- STAFF
        }

        # Prefer under-target
        under <- eligible[sapply(eligible, function(p) {
          !is.na(pp) && self$pp_counts[[p]][[pp]] < self$targets[[p]][[pp]]$sched_target
        })]
        pool <- if (length(under) > 0L) under else eligible

        # Sort by night score (lexicographic on priority, total, pp)
        scores    <- lapply(pool, function(p) self$night_score(p, d))
        ord       <- order(
          sapply(scores, `[[`, "priority"),
          sapply(scores, `[[`, "total"),
          sapply(scores, `[[`, "pp")
        )
        chosen <- pool[ord[1L]]
        self$add_shift(chosen, d, "Night")
      }
    },

    schedule_days = function() {
      for (d in self$dates) {
        ds      <- as.character(d)
        pp      <- get_pp(d)
        night_p <- self$schedule[[ds]]$Night

        for (slot in DAY_SLOTS) {
          if (!is.na(self$schedule[[ds]][[slot]])) next  # already filled

          eligible <- STAFF[sapply(STAFF, function(p) {
            self$can_work_day(p, d) &&
              (is.na(night_p) || night_p != p)
          })]
          if (length(eligible) == 0L) next

          under <- eligible[sapply(eligible, function(p) {
            !is.na(pp) &&
              self$pp_counts[[p]][[pp]] < self$targets[[p]][[pp]]$sched_target
          })]
          pool <- if (length(under) > 0L) under else eligible

          scores <- sapply(pool, function(p) self$day_score(p, d, slot))
          chosen <- pool[which.max(scores)]
          self$add_shift(chosen, d, slot)
        }
      }
    },

    cleanup_pass = function() {
      for (i in seq_len(nrow(PAY_PERIODS))) {
        pp_name  <- PAY_PERIODS$name[i]
        p_dates  <- seq(PAY_PERIODS$start[i], PAY_PERIODS$end[i], by = "day")

        for (person in STAFF) {
          pp_info  <- self$targets[[person]][[pp_name]]
          current  <- self$pp_counts[[person]][[pp_name]]
          needed   <- pp_info$sched_target - current
          if (needed <= 0L) next

          for (d in p_dates) {
            if (needed <= 0L) break
            if (!self$can_work_day(person, d)) next
            ds <- as.character(d)
            v  <- self$schedule[[ds]]$Night
            if (!is.na(v) && v == person) next

            for (slot in DAY_SLOTS) {
              if (is.na(self$schedule[[ds]][[slot]])) {
                self$add_shift(person, d, slot)
                needed <- needed - 1L
                break
              }
            }
          }
        }
      }
    },

    force_fill_app1 = function() {
      for (d in self$dates) {
        ds <- as.character(d)
        if (!is.na(self$schedule[[ds]]$APP1)) next
        pp <- get_pp(d)

        night_p <- self$schedule[[ds]]$Night
        app2_p  <- self$schedule[[ds]]$APP2
        roam_p  <- self$schedule[[ds]]$Roaming

        eligible <- STAFF[sapply(STAFF, function(p) {
          !self$is_blocked(p, d) &&
          (is.na(night_p) || night_p != p) &&
          (is.na(app2_p)  || app2_p  != p) &&
          (is.na(roam_p)  || roam_p  != p) &&
          !self$had_night_on(p, d - 1L)
        })]
        if (length(eligible) == 0L) {
          eligible <- STAFF[sapply(STAFF, function(p)
            is.na(night_p) || night_p != p)]
        }
        if (length(eligible) == 0L) eligible <- STAFF

        # Prefer least-loaded person in this PP
        counts <- sapply(eligible, function(p)
          if (!is.na(pp)) self$pp_counts[[p]][[pp]] else 0L)
        chosen <- eligible[which.min(counts)]
        self$add_shift(chosen, d, "APP1")
      }
    },

    # ── Main entry point ──────────────────────────────────────────────────────
    run = function() {
      message("  Pre-seeding holidays...")
      self$preseed_holidays()
      message("  Scheduling nights (hardest-first)...")
      self$schedule_nights()
      message("  Scheduling day shifts (forward fill)...")
      self$schedule_days()
      message("  Cleanup pass...")
      self$cleanup_pass()
      message("  Force-filling APP1...")
      self$force_fill_app1()
      message("  Done.")
      invisible(self)
    },

    # ── Export ────────────────────────────────────────────────────────────────

    #' Convert internal state to a flat data.frame (one row per date)
    to_dataframe = function() {
      df <- do.call(rbind, lapply(self$dates, function(d) {
        ds  <- as.character(d)
        day <- self$schedule[[ds]]
        data.frame(
          date    = d,
          day_name = weekdays(d, abbreviate = TRUE),
          pp      = get_pp(d),
          APP1    = ifelse(is.na(day$APP1),    "", day$APP1),
          APP2    = ifelse(is.na(day$APP2),    "", day$APP2),
          Roaming = ifelse(is.na(day$Roaming), "", day$Roaming),
          Night   = ifelse(is.na(day$Night),   "", day$Night),
          stringsAsFactors = FALSE
        )
      }))
      rownames(df) <- NULL
      df
    },

    #' Full grid: one row per (date × person) with their role and display metadata
    to_person_grid = function(time_off, targets) {
      all_d <- self$dates
      rows <- lapply(all_d, function(d) {
        ds      <- as.character(d)
        day_s   <- self$schedule[[ds]]
        pp      <- get_pp(d)
        lapply(STAFF, function(person) {
          role  <- NA_character_
          for (s in SLOTS) {
            v <- day_s[[s]]
            if (!is.na(v) && v == person) {
              role <- if (s == "Night") "Night" else
                      if (s == "APP1")  "APP1"  else
                      if (s == "APP2")  "APP2"  else "Roam"
              break
            }
          }
          # If not working, check time-off type
          if (is.na(role)) {
            pdata <- time_off[[person]]
            m     <- pdata[pdata$date == d, ]
            typ   <- if (nrow(m) > 0) m$type[1] else NA_character_
            # Conference override
            if (person %in% PP13_CONFERENCE &&
                d >= PP13_CONF_START && d <= PP13_CONF_END) typ <- "cme"
            if (!is.na(typ)) {
              role <- if (typ == "cme") "CME" else
                      if (typ == "off") "OFF" else {
                        # vac: charge PTO if avail < 6 in this PP
                        pp_info <- targets[[person]][[pp]]
                        if (!is.null(pp_info) && !is.na(pp) && pp_info$avail < 6L)
                          "PTO" else "VAC"
                      }
            }
          }
          is_hol <- d %in% HOLIDAY_DATES
          data.frame(
            date       = d,
            day_name   = weekdays(d, abbreviate = TRUE),
            pp         = pp,
            person     = person,
            role       = ifelse(is.na(role), "", role),
            is_holiday = is_hol,
            is_weekend = is_weekend(d),
            stringsAsFactors = FALSE
          )
        })
      })
      do.call(rbind, unlist(rows, recursive = FALSE))
    }
  )
)
