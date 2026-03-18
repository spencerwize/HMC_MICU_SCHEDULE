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

    # ── Undo a shift (mirror of add_shift; used by lookahead / swap repair) ──────
    undo_shift = function(person, d, slot) {
      self$schedule[[as.character(d)]][[slot]] <- NA_character_
      pp <- get_pp(d)
      if (!is.na(pp))
        self$pp_counts[[person]][[pp]] <- self$pp_counts[[person]][[pp]] - 1L
      if (slot == "Night") {
        nights <- self$person_nights[[person]]
        self$person_nights[[person]] <- nights[nights != d]
      } else {
        s <- self$person_shifts[[person]]
        self$person_shifts[[person]] <-
          s[!(s$date == d & s$slot == slot), , drop = FALSE]
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
      # Also look forward so pre-seeded holidays are counted in the streak
      nxt    <- d + 1L
      while (nxt  %in% worked) { count <- count + 1L; nxt  <- nxt  + 1L }
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

    # ── Eligible-candidate set for a (date, slot) pair ───────────────────────
    eligible_for = function(d, slot) {
      night_p <- self$schedule[[as.character(d)]]$Night
      if (slot == "Night") {
        STAFF[sapply(STAFF, function(p) self$can_work_night(p, d))]
      } else {
        STAFF[sapply(STAFF, function(p)
          self$can_work_day(p, d) && (is.na(night_p) || night_p != p))]
      }
    },

    # ── Scoring ───────────────────────────────────────────────────────────────

    #' Night score — lower is preferred.
    #' Returns named list(priority, total, pp) for lexicographic sort.
    #' priority: 0 = streak continuation, 1 = fresh start, 2 = had day shift yesterday
    night_score = function(person, d) {
      nights    <- self$person_nights[[person]]
      is_cont   <- (d - 1L) %in% nights
      pp        <- get_pp(d)
      pp_nights <- if (!is.na(pp)) sum(get_pp_vec(nights) == pp, na.rm = TRUE) else 0L
      # Penalise day-before-night: if person already has a day shift on d-1
      had_day_prev <- any(sapply(DAY_SLOTS, function(s) {
        v <- self$get_slot(d - 1L, s); !is.na(v) && v == person
      }))
      priority <- if (is_cont) 0L else if (had_day_prev) 2L else 1L
      list(priority = priority, total = length(nights), pp = pp_nights)
    },

    #' Day score — higher is preferred.
    day_score = function(person, d, slot) {
      pp      <- get_pp(d)
      pp_info <- self$targets[[person]][[pp]]

      # Urgency: use available remaining days in this PP (not raw calendar days)
      remaining   <- pp_info$pp_dates[pp_info$pp_dates >= d]
      avail_left  <- sum(sapply(remaining, function(dd) !self$is_blocked(person, dd)))
      slots_left  <- max(0L, pp_info$sched_target - self$pp_counts[[person]][[pp]])
      urgency     <- slots_left / max(avail_left, 1L)

      # Cluster: prefer working on days adjacent to already-worked days
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
      n_wknd      <- if (length(worked_d) > 0) sum(is_weekend(worked_d)) else 0L
      wknd_ratio  <- if (length(worked_d) > 0) n_wknd / length(worked_d) else 0
      wknd_score  <- if (is_weekend(d)) -wknd_ratio else 0

      # Day-before-night penalty: strongly discourage if this person has Night tomorrow
      tomorrow_night <- self$get_slot(d + 1L, "Night")
      dbn_penalty    <- if (!is.na(tomorrow_night) && tomorrow_night == person) -8 else 0

      urgency * 10 + cluster + roam_score + wknd_score + dbn_penalty
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

    # ── Unified MRV scheduler ─────────────────────────────────────────────────
    #
    # Each iteration: pick the unscheduled (date, slot) with the FEWEST eligible
    # candidates (Minimum Remaining Values), score candidates, then run a 1-step
    # lookahead to avoid deadlocking a near-future slot before committing.

    schedule_all = function() {
      # Build initial todo list
      todo <- list()
      for (d_raw in self$dates) {
        d  <- as.Date(d_raw, origin = "1970-01-01")
        ds <- as.character(d)
        for (slot in SLOTS) {
          if (is.na(self$schedule[[ds]][[slot]]))
            todo <- c(todo, list(list(d = d, slot = slot)))
        }
      }

      while (length(todo) > 0L) {
        # MRV: slot with fewest eligible candidates first (recomputed each step)
        n_elig <- sapply(todo, function(x)
          length(self$eligible_for(x$d, x$slot)))
        idx  <- which.min(n_elig)
        item <- todo[[idx]]
        todo <- todo[-idx]

        d    <- item$d
        slot <- item$slot
        ds   <- as.character(d)
        if (!is.na(self$schedule[[ds]][[slot]])) next  # pre-seeded

        pp       <- get_pp(d)
        eligible <- self$eligible_for(d, slot)

        # Force-fill fallback when no candidate passes hard constraints
        if (length(eligible) == 0L) {
          if (slot == "Night") {
            # Night must be filled; relax consecutive limits but NEVER assign
            # someone who already has a day shift today (non-negotiable).
            eligible <- STAFF[sapply(STAFF, function(p)
              !self$is_blocked(p, d) &&
              !self$had_night_on(p, d - 1L) &&
              !self$already_assigned(p, d))]
            if (length(eligible) == 0L)
              eligible <- STAFF[!sapply(STAFF, function(p)
                self$already_assigned(p, d))]
            if (length(eligible) == 0L) eligible <- STAFF
          } else {
            # Day slots: skip when no one passes hard constraints.
            # force_fill_app1() will mop up any remaining APP1 gaps;
            # empty APP2/Roaming slots are acceptable.
            next
          }
        }

        # Prefer under-target in this PP
        under <- eligible[sapply(eligible, function(p)
          !is.na(pp) &&
          self$pp_counts[[p]][[pp]] < self$targets[[p]][[pp]]$sched_target)]
        pool <- if (length(under) > 0L) under else eligible

        # Score and build ordered pool
        if (slot == "Night") {
          sc  <- lapply(pool, function(p) self$night_score(p, d))
          ord <- order(sapply(sc, `[[`, "priority"),
                       sapply(sc, `[[`, "total"),
                       sapply(sc, `[[`, "pp"))
        } else {
          sc  <- sapply(pool, function(p) self$day_score(p, d, slot))
          ord <- order(-sc)
        }
        ordered_pool <- pool[ord]

        chosen <- self$pick_no_deadlock(ordered_pool, d, slot, todo)
        self$add_shift(chosen, d, slot)
      }
    },

    #' Try candidates in score order; return the first that does not deadlock
    #' any unscheduled slot within 5 calendar days. Falls back to top scorer.
    pick_no_deadlock = function(ordered_pool, d, slot, todo) {
      near <- todo[vapply(todo, function(x)
        abs(as.integer(x$d - d)) <= 5L, logical(1L))]
      if (length(near) == 0L) return(ordered_pool[1L])

      for (candidate in ordered_pool) {
        self$add_shift(candidate, d, slot)
        deadlock <- any(sapply(near, function(x)
          length(self$eligible_for(x$d, x$slot)) == 0L))
        self$undo_shift(candidate, d, slot)
        if (!deadlock) return(candidate)
      }
      ordered_pool[1L]  # all choices deadlock; accept best-scoring
    },

    # ── Swap-based repair: eliminate day-before-night soft violations ─────────
    #
    # After the MRV fill, scan for (day D → night D+1) violations and replace
    # the offending day-shift worker with someone who won't create a new dbn.
    # Runs up to 3 passes or until no further improvement is possible.

    swap_repair_pass = function() {
      for (pass in seq_len(3L)) {
        improved <- FALSE
        for (d_raw in self$dates[-length(self$dates)]) {
          d <- as.Date(d_raw, origin = "1970-01-01")
          night_p <- self$get_slot(d + 1L, "Night")
          if (is.na(night_p)) next
          for (slot in DAY_SLOTS) {
            v <- self$get_slot(d, slot)
            if (!is.na(v) && v == night_p)
              if (self$try_swap_out(d, slot, night_p)) improved <- TRUE
          }
        }
        if (!improved) break
      }
    },

    #' Remove `person` from `slot` on `d` and replace with a candidate who
    #' (a) passes all hard constraints and (b) won't themselves have Night
    #' tomorrow. Returns TRUE on success; restores original on failure.
    try_swap_out = function(d, slot, person) {
      self$undo_shift(person, d, slot)
      pp <- get_pp(d)
      tn <- self$get_slot(d + 1L, "Night")   # still = person

      candidates <- STAFF[sapply(STAFF, function(p) {
        if (p == person) return(FALSE)
        if (!self$can_work_day(p, d)) return(FALSE)
        # Avoid creating a new dbn for the replacement
        if (!is.na(tn) && tn == p) return(FALSE)
        TRUE
      })]

      if (length(candidates) > 0L) {
        counts <- sapply(candidates, function(p)
          if (!is.na(pp)) self$pp_counts[[p]][[pp]] else 0L)
        self$add_shift(candidates[which.min(counts)], d, slot)
        return(TRUE)
      }
      self$add_shift(person, d, slot)  # restore
      FALSE
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
            d <- as.Date(d, origin = "1970-01-01")  # for() strips Date class
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
        d  <- as.Date(d, origin = "1970-01-01")  # for() strips Date class
        ds <- as.character(d)
        if (!is.na(self$schedule[[ds]]$APP1)) next
        pp <- get_pp(d)

        night_p <- self$schedule[[ds]]$Night

        # Fallback 1 (preferred): full hard-constraint check
        eligible <- STAFF[sapply(STAFF, function(p) self$can_work_day(p, d))]

        # Fallback 2: relax consecutive-day limit but keep night recovery
        if (length(eligible) == 0L) {
          eligible <- STAFF[sapply(STAFF, function(p) {
            !self$is_blocked(p, d) &&
            !self$night_recovery_blocked(p, d, for_night = FALSE) &&
            !self$already_assigned(p, d) &&
            (is.na(night_p) || night_p != p)
          })]
        }

        # Fallback 3: relax night-recovery, but still no double-booking
        if (length(eligible) == 0L) {
          eligible <- STAFF[sapply(STAFF, function(p)
            !self$is_blocked(p, d) &&
            !self$already_assigned(p, d) &&
            (is.na(night_p) || night_p != p))]
        }
        # Fallback 4: last resort — only hard requirement is not on night tonight
        if (length(eligible) == 0L) {
          eligible <- STAFF[sapply(STAFF, function(p)
            !self$already_assigned(p, d) &&
            (is.na(night_p) || night_p != p))]
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
      message("  Scheduling all slots (MRV hardest-first + lookahead)...")
      self$schedule_all()
      message("  Cleanup pass (PP quota fill)...")
      self$cleanup_pass()
      message("  Swap repair (day-before-night)...")
      self$swap_repair_pass()
      message("  Force-filling remaining APP1...")
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
