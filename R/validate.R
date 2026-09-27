# ─────────────────────────────────────────────────────────────────────────────
# validate.R  —  Hard-constraint checker
#
# Returns a list with:
#   $errors   character vector of constraint violations
#   $warnings character vector of soft-constraint flags
# ─────────────────────────────────────────────────────────────────────────────

# `partial = TRUE` validates a GREEN-ONLY (phase 1) schedule. That phase is
# incomplete by construction - slots go unfilled where too few people volunteered,
# and every per-person FLOOR is deliberately not enforced (see the phase gates in
# scheduler_lp.R). So coverage and floor checks are demoted to warnings, while
# every CEILING and safety rule is still checked exactly as in a full schedule.
validate_schedule <- function(sched_obj, time_off, targets, partial = FALSE) {
  errors   <- character()
  warnings <- character()
  dates    <- sched_obj$dates

  # Coverage/floor findings: hard errors in a finished schedule, expected
  # outcomes in a partial one.
  add_gap <- function(msg) {
    if (partial) warnings <<- c(warnings, msg) else errors <<- c(errors, msg)
  }

  # for() strips Date class from Date vectors; always re-cast inside loops
  as_date <- function(x) as.Date(x, origin = "1970-01-01")

  get <- function(d, slot) {
    v <- sched_obj$schedule[[as.character(as_date(d))]][[slot]]
    if (is.null(v)) NA_character_ else v
  }

  # ── 1. APP1 always filled ──────────────────────────────────────────────────
  for (d in dates) {
    d <- as_date(d)
    if (is.na(get(d, "APP1"))) {
      add_gap(sprintf("%s: APP1 not filled", as.character(d)))
    }
  }

  # ── 1b. APP2 always filled ────────────────────────────────────────────────
  # C2b makes this a hard equality in the ILP, but it was never checked here.
  for (d in dates) {
    d <- as_date(d)
    if (is.na(get(d, "APP2"))) {
      add_gap(sprintf("%s: APP2 not filled", as.character(d)))
    }
  }

  # ── 2. Night always filled ────────────────────────────────────────────────
  for (d in dates) {
    d <- as_date(d)
    if (is.na(get(d, "Night"))) {
      add_gap(sprintf("%s: Night not filled", as.character(d)))
    }
  }

  for (person in STAFF) {
    nights <- sort(sched_obj$person_nights[[person]])

    # ── 3. Night → no day shift next morning ─────────────────────────────────
    for (nd in nights) {
      nd     <- as_date(nd)
      next_d <- nd + 1L
      if (next_d %in% dates) {
        for (s in DAY_SLOTS) {
          v <- get(next_d, s)
          if (!is.na(v) && v == person) {
            errors <- c(errors, sprintf(
              "%s: day shift on %s after night on %s",
              person, as.character(next_d), as.character(nd)))
          }
        }
      }
    }

    # ── 5. Max 3 consecutive nights ────────────────────────────────────────
    night_set <- as.character(nights)
    for (nd in nights) {
      nd <- as_date(nd)
      if (all(as.character(nd - (1:3)) %in% night_set)) {
        errors <- c(errors, sprintf(
          "%s: 4+ consecutive nights ending %s", person, as.character(nd)))
      }
    }

    # ── 6. Max 4 consecutive working days ─────────────────────────────────
    worked <- sort(c(sched_obj$person_shifts[[person]]$date, nights))
    worked_set <- as.character(worked)
    for (wd in worked) {
      wd <- as_date(wd)
      if (all(as.character(wd - (1:4)) %in% worked_set)) {
        errors <- c(errors, sprintf(
          "%s: 5+ consecutive working days ending %s",
          person, as.character(wd)))
      }
    }

    # ── 7. Day-to-night same calendar day ─────────────────────────────────
    for (d in dates) {
      d <- as_date(d)
      has_day   <- any(sapply(DAY_SLOTS, function(s) {
        v <- get(d, s); !is.na(v) && v == person
      }))
      has_night <- !is.na(get(d, "Night")) && get(d, "Night") == person
      if (has_day && has_night) {
        errors <- c(errors, sprintf(
          "%s: both day and night on %s", person, as.character(d)))
      }
    }
  }

  # ── Double-booking (one person, multiple slots same day) ─────────────────
  for (d in dates) {
    d <- as_date(d)
    assigned <- Filter(Negate(is.na), sapply(SLOTS, function(s) get(d, s)))
    dups <- assigned[duplicated(assigned)]
    if (length(dups) > 0L) {
      errors <- c(errors, sprintf(
        "%s: double-booked: %s", as.character(d),
        paste(unique(dups), collapse = ", ")))
    }
  }

  # ── C10c: 8-day density cap (max 6 shifts in any 8-day window) ───────────
  for (person in STAFF) {
    all_worked <- sort(c(sched_obj$person_shifts[[person]]$date,
                         sched_obj$person_nights[[person]]))
    for (i in seq_along(dates)) {
      d_start <- as_date(dates[i])
      d_end   <- d_start + 7L
      if (d_end > as_date(dates[length(dates)])) break
      n_in_window <- sum(all_worked >= d_start & all_worked <= d_end)
      if (n_in_window > 6L) {
        errors <- c(errors, sprintf(
          "%s: %d shifts in 8-day window %s–%s (max 6)",
          person, n_in_window, as.character(d_start), as.character(d_end)))
      }
    }
  }

  # ── C9b: nights must be stacked (no Night-Empty-Night) ────────────────────
  for (person in STAFF) {
    nights <- sort(as.integer(sched_obj$person_nights[[person]]))
    if (length(nights) < 2L) next
    for (i in 2:length(nights)) {
      if (nights[i] - nights[i - 1L] == 2L)
        errors <- c(errors, sprintf(
          "%s: nights on %s and %s with %s off between (nights must be stacked)",
          person,
          as.character(as_date(nights[i - 1L])),
          as.character(as_date(nights[i])),
          as.character(as_date(nights[i - 1L] + 1L))))
    }
  }

  # ── C11b: no night stretch starting in consecutive PPs ───────────────────
  # This rule is dropped by the relaxation cascade when the per-person night
  # FLOOR cannot otherwise be met (C11b caps a person at one stretch start per
  # two adjacent pay periods, i.e. four across seven - too few to reach 8
  # nights). When the tier that produced this schedule dropped it, a violation
  # is an accepted trade-off, not a defect, so it is reported as a warning.
  c11b_relaxed <- isTRUE(!is.null(sched_obj$tier_used$relaxed) &&
                         identical(sched_obj$tier_used$relaxed$c11b, FALSE))
  add_c11b_finding <- function(msg) {
    if (c11b_relaxed) warnings <<- c(warnings, paste0("[C11b relaxed] ", msg))
    else              errors   <<- c(errors, msg)
  }
  for (person in STAFF) {
    nights <- sort(sched_obj$person_nights[[person]])
    if (length(nights) == 0L) next
    # Identify stretch-start dates: night on d with no night on d-1
    stretch_starts <- vapply(nights, function(nd) {
      nd <- as_date(nd)
      !(nd - 1L) %in% nights
    }, logical(1L))
    start_dates <- nights[stretch_starts]
    start_pps   <- vapply(start_dates, function(d) get_pp(as_date(d)), character(1L))
    for (k in seq_len(nrow(PAY_PERIODS) - 1L)) {
      pp_k  <- PAY_PERIODS$name[k]
      pp_k1 <- PAY_PERIODS$name[k + 1L]
      n_starts <- sum(start_pps %in% c(pp_k, pp_k1), na.rm = TRUE)
      if (n_starts > 1L) {
        add_c11b_finding(sprintf(
          "%s: night stretches start in consecutive PPs %s and %s",
          person, pp_k, pp_k1))
      }
    }
  }

  # ── C11d: per-person night band ───────────────────────────────────────────
  # Absolute: no relaxation tier widens it. If the schedule cannot fit all nights
  # inside the band, nights are left UNSTAFFED instead - so a violation here is a
  # real bug, not a capacity symptom.
  for (person in STAFF) {
    n_nights <- length(sched_obj$person_nights[[person]])
    if (n_nights < MIN_NIGHTS_HARD)
      add_gap(sprintf("%s: only %d night shifts (min %d)",
                      person, n_nights, MIN_NIGHTS_HARD))
    else if (n_nights > MAX_NIGHTS_HARD)
      errors <- c(errors, sprintf("%s: %d night shifts exceeds max %d",
                                  person, n_nights, MAX_NIGHTS_HARD))
  }

  # ── C5b: no night the day before an off / vacation / CME day ──────────────
  for (person in STAFF) {
    pdata <- time_off[[person]]
    if (is.null(pdata) || nrow(pdata) == 0) next
    blocked_next <- pdata$date[pdata$type %in% BLOCKED_TYPES]
    if (!length(blocked_next)) next
    for (nd in sched_obj$person_nights[[person]]) {
      nd <- as_date(nd)
      if ((nd + 1L) %in% blocked_next) {
        typ <- pdata$type[pdata$date == (nd + 1L)][1]
        errors <- c(errors, sprintf(
          "%s: night on %s but %s is %s", person, as.character(nd),
          as.character(nd + 1L), toupper(typ)))
      }
    }
  }

  # ── C12: holiday pre-seeds must survive intact ────────────────────────────
  # HOLIDAYS assigns specific people to specific holiday slots; C12 fixes them
  # with lb = ub = 1. The greedy aesthetic pass used to swap them away silently,
  # which nothing here caught. Checked now so it cannot regress.
  for (ds in names(HOLIDAYS)) {
    d_hol <- as.Date(ds)
    if (d_hol < SCHEDULE_START || d_hol > SCHEDULE_END) next
    for (sl in names(HOLIDAYS[[ds]])) {
      want <- HOLIDAYS[[ds]][[sl]]
      got  <- get(d_hol, sl)
      # A designated person who requested the day off is legitimately skipped.
      pdata <- time_off[[want]]
      excused <- !is.null(pdata) && nrow(pdata) > 0 &&
        any(pdata$date == d_hol & pdata$type %in% BLOCKED_TYPES)
      if (excused) next
      if (is.na(got)) {
        add_gap(sprintf("%s: holiday %s slot empty (assigned to %s)", ds, sl, want))
      } else if (got != want) {
        errors <- c(errors, sprintf(
          "%s: holiday %s assigned to %s but %s is scheduled", ds, sl, want, got))
      }
    }
  }

  # ── C11c: weekend hard bounds ─────────────────────────────────────────────
  # Weekend = Friday night through Sunday night; see is_weekend_shift().
  # Must match the ILP's definition or this reports violations that are not real.
  for (person in STAFF) {
    sh     <- sched_obj$person_shifts[[person]]
    nights <- sched_obj$person_nights[[person]]
    n_wknd <- sum(is_weekend_shift(sh$date, sh$slot)) +
              sum(is_weekend_shift(nights, "Night"))
    if (n_wknd < MIN_WKND_HARD) {
      add_gap(sprintf(
        "%s: only %d weekend shifts (min %d)", person, n_wknd, MIN_WKND_HARD))
    } else if (n_wknd > MAX_WKND_HARD) {
      errors <- c(errors, sprintf(
        "%s: %d weekend shifts exceeds max %d", person, n_wknd, MAX_WKND_HARD))
    }
  }

  # ── Soft: day-before-night flag ───────────────────────────────────────────
  for (d in dates[-length(dates)]) {
    d <- as_date(d)
    tomorrow_night <- get(d + 1L, "Night")
    if (!is.na(tomorrow_night)) {
      for (s in DAY_SLOTS) {
        v <- get(d, s)
        if (!is.na(v) && v == tomorrow_night) {
          warnings <- c(warnings, sprintf(
            "%s: %s works day on %s then night on %s (soft buffer violation)",
            tomorrow_night, s, as.character(d), as.character(d + 1L)))
        }
      }
    }
  }

  # ── PP target summary ─────────────────────────────────────────────────────
  for (person in STAFF) {
    for (pp_name in PAY_PERIODS$name) {
      info    <- targets[[person]][[pp_name]]
      actual  <- sched_obj$pp_counts[[person]][[pp_name]]
      cme     <- info$credited
      sched_t <- info$sched_target   # shifts still needed beyond CME
      full_t  <- info$target         # full base target (before CME subtraction)
      if (actual < sched_t) {
        cme_note <- if (cme > 0L)
          sprintf(" + %d CME = %d / full target %d", cme, actual + cme, full_t)
        else
          sprintf(" / target %d", full_t)
        warnings <- c(warnings, sprintf(
          "%s %s: scheduled %d%s",
          person, pp_name, actual, cme_note))
      }
    }
  }

  list(errors = errors, warnings = warnings)
}

#' Print validation results to console
print_validation <- function(result) {
  if (length(result$errors) == 0L) {
    message("  ✓ No hard constraint violations.")
  } else {
    message(sprintf("  ✗ %d ERRORS:", length(result$errors)))
    for (e in result$errors) message("    ERROR: ", e)
  }
  if (length(result$warnings) > 0L) {
    message(sprintf("  ⚠ %d warnings:", length(result$warnings)))
    for (w in result$warnings) message("    WARN:  ", w)
  }
}

# ─────────────────────────────────────────────────────────────────────────────
# green_summary()  —  How much of the schedule landed on requested-work days?
#
# The headline percentage is meaningless on its own: if people mark 60% of all
# available days green, a scheduler that ignores green entirely would still hit
# roughly 60%. So the chance baseline is reported alongside it, and the number
# that matters is the LIFT over that baseline.
#
# Returns invisibly: list(overall, by_person, by_pp)
# ─────────────────────────────────────────────────────────────────────────────
green_summary <- function(sched_obj, time_off, targets, quiet = FALSE) {
  green_key <- setNames(lapply(STAFF, function(p) {
    df <- time_off[[p]]
    if (is.null(df) || !nrow(df)) character(0)
    else as.character(df$date[df$type == "green"])
  }), STAFF)
  blocked_key <- setNames(lapply(STAFF, function(p) {
    df <- time_off[[p]]
    if (is.null(df) || !nrow(df)) character(0)
    else as.character(df$date[df$type %in% BLOCKED_TYPES])
  }), STAFF)

  all_ds <- as.character(sched_obj$dates)

  worked <- setNames(lapply(STAFF, function(p) {
    c(as.character(sched_obj$person_shifts[[p]]$date),
      as.character(sched_obj$person_nights[[p]]))
  }), STAFF)

  by_person <- do.call(rbind, lapply(STAFF, function(p) {
    w  <- worked[[p]]
    wg <- sum(w %in% green_key[[p]])
    data.frame(
      person       = p,
      green_avail  = length(green_key[[p]]),
      total        = length(w),
      worked_green = wg,
      worked_yellow= length(w) - wg,
      green_unused = length(setdiff(green_key[[p]], w)),
      pct_green    = if (length(w)) 100 * wg / length(w) else NA_real_,
      # Chance baseline: green days as a share of everything this person COULD
      # have been scheduled on.
      baseline_pct = 100 * length(green_key[[p]]) /
                     max(1L, length(all_ds) - length(blocked_key[[p]])),
      stringsAsFactors = FALSE)
  }))

  tot   <- sum(by_person$total)
  tot_g <- sum(by_person$worked_green)
  tot_y <- sum(by_person$worked_yellow)
  base  <- 100 * sum(by_person$green_avail) /
           max(1L, sum(length(all_ds) - lengths(blocked_key)))
  worst <- by_person$person[which.max(by_person$worked_yellow)]

  by_pp <- do.call(rbind, lapply(STAFF, function(p) {
    do.call(rbind, lapply(PAY_PERIODS$name, function(pp) {
      pd <- as.character(pp_dates(pp))
      w  <- worked[[p]][worked[[p]] %in% pd]
      data.frame(person = p, pp = pp,
                 worked = length(w),
                 worked_green = sum(w %in% green_key[[p]]),
                 sched_target = as.integer(targets[[p]][[pp]]$sched_target),
                 stringsAsFactors = FALSE)
    }))
  }))

  overall <- list(total = tot, green = tot_g, yellow = tot_y,
                  pct_green = if (tot) 100 * tot_g / tot else NA_real_,
                  baseline_pct = base,
                  max_yellow = max(by_person$worked_yellow),
                  max_yellow_person = worst)

  if (!quiet) {
    message("Requested-work (green) fill:")
    message(sprintf("  %d of %d shifts on requested days = %.1f%%  (chance baseline %.1f%%, lift %+.1f pts)",
                    tot_g, tot, overall$pct_green, base, overall$pct_green - base))
    message(sprintf("  %d shift(s) on yellow days; worst-off person: %s with %d",
                    tot_y, worst, overall$max_yellow))
    message("  person      green  total  onGreen  onYellow  unused   pct")
    for (i in seq_len(nrow(by_person)))
      message(sprintf("  %-10s %5d  %5d  %7d  %8d  %6d  %4.0f%%",
                      by_person$person[i], by_person$green_avail[i],
                      by_person$total[i], by_person$worked_green[i],
                      by_person$worked_yellow[i], by_person$green_unused[i],
                      by_person$pct_green[i]))
  }
  invisible(list(overall = overall, by_person = by_person, by_pp = by_pp))
}
