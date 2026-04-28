# ─────────────────────────────────────────────────────────────────────────────
# validate.R  —  Hard-constraint checker
#
# Returns a list with:
#   $errors   character vector of constraint violations
#   $warnings character vector of soft-constraint flags
# ─────────────────────────────────────────────────────────────────────────────

validate_schedule <- function(sched_obj, time_off, targets) {
  errors   <- character()
  warnings <- character()
  dates    <- sched_obj$dates

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
      errors <- c(errors, sprintf("%s: APP1 not filled", as.character(d)))
    }
  }

  # ── 2. Night always filled ────────────────────────────────────────────────
  for (d in dates) {
    d <- as_date(d)
    if (is.na(get(d, "Night"))) {
      errors <- c(errors, sprintf("%s: Night not filled", as.character(d)))
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

  # ── C11b: no night stretch starting in consecutive PPs ───────────────────
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
        errors <- c(errors, sprintf(
          "%s: night stretches start in consecutive PPs %s and %s",
          person, pp_k, pp_k1))
      }
    }
  }

  # ── C11c: weekend hard bounds ─────────────────────────────────────────────
  wknd_dates <- dates[weekdays(as_date(dates)) %in% c("Saturday", "Sunday")]
  for (person in STAFF) {
    sh     <- sched_obj$person_shifts[[person]]
    nights <- sched_obj$person_nights[[person]]
    worked <- sort(c(sh$date, nights))
    n_wknd <- sum(worked %in% wknd_dates)
    if (n_wknd < MIN_WKND_HARD) {
      errors <- c(errors, sprintf(
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
