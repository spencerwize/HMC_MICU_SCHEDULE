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

    # ── 3 & 4. Night recovery ─────────────────────────────────────────────
    for (nd in nights) {
      nd <- as_date(nd)
      for (offset in 1:2) {
        check_d <- nd + offset
        if (!check_d %in% dates) next
        for (s in DAY_SLOTS) {
          v <- get(check_d, s)
          if (!is.na(v) && v == person) {
            errors <- c(errors, sprintf(
              "%s: day shift on %s within %d day(s) of night on %s",
              person, as.character(check_d), offset, as.character(nd)))
          }
        }
      }
      # Constraint 3: no day shift the morning after (d+1)
      next_d <- nd + 1L
      if (next_d %in% dates) {
        for (s in DAY_SLOTS) {
          v <- get(next_d, s)
          if (!is.na(v) && v == person) {
            errors <- c(errors, sprintf(
              "%s: day shift on %s after night on %s (constraint 3)",
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
      actual <- sched_obj$pp_counts[[person]][[pp_name]]
      target <- targets[[person]][[pp_name]]$sched_target
      cme    <- targets[[person]][[pp_name]]$credited
      if (actual < target) {
        warnings <- c(warnings, sprintf(
          "%s %s: scheduled %d / target %d (CME credited: %d)",
          person, pp_name, actual, target, cme))
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
