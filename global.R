# ─────────────────────────────────────────────────────────────────────────────
# global.R  —  Loaded by both Shiny (app startup) and run_schedule.R
# ─────────────────────────────────────────────────────────────────────────────

suppressPackageStartupMessages({
  library(R6)
  library(openxlsx)
  library(dplyr)
  library(tidyr)
  library(highs)
  library(Matrix)
  library(shiny)
  library(bslib)
  library(reactable)
  library(plotly)
  library(shinyWidgets)
  library(htmltools)
  library(shinyjs)
})

# Source all R module files
r_files <- c(
  "R/constants.R",
  "R/roles.R",
  "R/parse_time_off.R",
  "R/targets.R",
  "R/scheduler_lp.R",
  "R/validate.R",
  "R/excel_output.R"
)
for (f in r_files) source(f, local = FALSE)

# ── Google Sheets source ──────────────────────────────────────────────────────
TIMEOFF_GSHEET_URL <- "https://docs.google.com/spreadsheets/d/1dXiGuI0Ri6ogt3RqJT_EJQAFnU1i5ctb9W9kKRQS8oE/edit"

# Known tab names in the Google Sheet — pre-populates the sheet dropdown so
# it works immediately without API auth. Add/remove names to match the workbook.
TIMEOFF_SHEETS <- c("Oct26-Jan31")

# Per-sheet schedule configurations.
# Each entry maps a sheet tab name to the date range, pay periods, holiday
# pre-seeds, and calendar month choices that apply to that period.
SHEET_CONFIGS <- list(
  # ── Oct 26, 2026 – Jan 31, 2027 ────────────────────────────────────────────
  # Matches the "Oct26-Jan31" tab of the Red/Yellow/Green REQUESTS sheet:
  # 98 contiguous days = exactly 7 fourteen-day pay periods from Oct 26.
  "Oct26-Jan31" = list(
    schedule_start = as.Date("2026-10-26"),
    schedule_end   = as.Date("2027-01-31"),
    pay_periods = data.frame(
      name  = c("PP1","PP2","PP3","PP4","PP5","PP6","PP7"),
      start = as.Date(c("2026-10-26","2026-11-09","2026-11-23","2026-12-07",
                        "2026-12-21","2027-01-04","2027-01-18")),
      end   = as.Date(c("2026-11-08","2026-11-22","2026-12-06","2026-12-20",
                        "2027-01-03","2027-01-17","2027-01-31")),
      stringsAsFactors = FALSE
    ),
    # No holiday pre-seeds: holiday coverage is expressed in the request sheet
    # itself (people mark Red/Green on those dates), so the solver honours it the
    # same way it honours any other request. The dates below are still
    # highlighted on the calendar for readability.
    holidays = list(),
    holiday_dates = as.Date(c("2026-11-26", "2026-12-25", "2027-01-01")),
    holiday_names = c("2026-11-26" = "Thanksgiving",
                      "2026-12-25" = "Christmas",
                      "2027-01-01" = "New Year's Day"),
    cal_months = c("October 2026"  = "2026-10", "November 2026" = "2026-11",
                   "December 2026" = "2026-12", "January 2027"  = "2027-01")
  ),
  "April13-July19" = list(
    schedule_start = as.Date("2026-04-13"),
    schedule_end   = as.Date("2026-07-19"),
    pay_periods = data.frame(
      name  = c("PP8","PP9","PP10","PP11","PP12","PP13","PP14"),
      start = as.Date(c("2026-04-13","2026-04-27","2026-05-11",
                        "2026-05-25","2026-06-08","2026-06-22","2026-07-06")),
      end   = as.Date(c("2026-04-26","2026-05-10","2026-05-24",
                        "2026-06-07","2026-06-21","2026-07-05","2026-07-19")),
      stringsAsFactors = FALSE
    ),
    holidays = list(
      "2026-05-25" = list(APP1 = "Hayden",  APP2 = "Todd",
                          Roaming = "Radha", Night = "Isabel"),
      "2026-06-19" = list(APP1 = "Mandie",  APP2 = "Caroline",
                          Roaming = "Radha", Night = "Isabel"),
      "2026-07-04" = list(APP1 = "Kristin", APP2 = "John",
                          Roaming = "Caroline", Night = "Mandie")
    ),
    holiday_names = c("2026-05-25" = "Memorial Day",
                      "2026-06-19" = "Juneteenth",
                      "2026-07-04" = "July 4th"),
    cal_months = c("April 2026" = "2026-04", "May 2026"  = "2026-05",
                   "June 2026"  = "2026-06", "July 2026" = "2026-07")
  ),
  "July20-Oct25" = list(
    schedule_start = as.Date("2026-07-20"),
    schedule_end   = as.Date("2026-10-25"),
    pay_periods = data.frame(
      name  = c("PP15","PP16","PP17","PP18","PP19","PP20","PP21"),
      start = as.Date(c("2026-07-20","2026-08-03","2026-08-17",
                        "2026-08-31","2026-09-14","2026-09-28","2026-10-12")),
      end   = as.Date(c("2026-08-02","2026-08-16","2026-08-30",
                        "2026-09-13","2026-09-27","2026-10-11","2026-10-25")),
      stringsAsFactors = FALSE
    ),
    # Pre-seed Labor Day (Sep 7) and the weekend prior (Sep 5-6) with the
    # same crew so they can all go out of town if they want.
    # holiday_dates controls which dates get the yellow holiday highlight —
    # only Labor Day itself, not the regular weekend days.
    holidays = list(
      "2026-09-05" = list(APP1 = "Katie",   APP2 = "Maureen",
                          Roaming = "Kristin", Night = "Hayden"),
      "2026-09-06" = list(APP1 = "Katie",   APP2 = "Maureen",
                          Roaming = "Kristin", Night = "Hayden"),
      "2026-09-07" = list(APP1 = "Katie",   APP2 = "Maureen",
                          Roaming = "Kristin", Night = "Hayden")
    ),
    holiday_dates = as.Date("2026-09-07"),   # only Labor Day colored yellow
    holiday_names = c("2026-09-07" = "Labor Day"),
    cal_months = c("July 2026"      = "2026-07", "August 2026"    = "2026-08",
                   "September 2026" = "2026-09", "October 2026"   = "2026-10")
  )
)

# Legacy env-var fallback (used by run_schedule.R)
TIMEOFF_DEFAULT_SOURCE <- Sys.getenv("TIMEOFF_SOURCE",
                                     unset = TIMEOFF_GSHEET_URL)

# ── Run the full pipeline and return a named list ─────────────────────────────
# `mode` selects how requested-work ("green") days are honoured:
#   "green"  - GREEN-ONLY phase only (DEFAULT). Partial by design: unfilled slots
#              where too few people volunteered, people left below target.
#              Inspect with sched$holes_df(), then complete it with mode = "fill".
#   "fill"   - green-only phase followed by the completion phase, which keeps as
#              many of its assignments as it can (hard pins first, falling back
#              to a strong preference) and minimises yellow-day work.
#   "single" - one solve over all days, preferring green via GREEN_WORK_BONUS.
#              No pinning, so the solver optimises globally.
#   `sheet`      - tab of the request workbook. Also selects the SHEET_CONFIGS
#                  entry (date range, pay periods, holidays). Defaults to the
#                  first entry of TIMEOFF_SHEETS. Getting this wrong silently
#                  schedules the wrong date window, so it is applied here rather
#                  than left to the caller.
#   `time_limit` - solver budget in SECONDS per solve. This model does not
#                  converge on real data, so the budget is effectively a quality
#                  dial: measured on the Oct26-Jan31 sheet, 600s left 27 shifts
#                  unfilled, 900s left 14, and 3600s left 7. Default
#                  SOLVER_TIME_LIMIT (3600).
run_pipeline <- function(path          = TIMEOFF_DEFAULT_SOURCE,
                         sheet         = NULL,
                         time_limit    = SOLVER_TIME_LIMIT,
                         verbose       = TRUE,
                         two_stage     = TRUE,
                         greedy_stage2 = TRUE,
                         mode          = c("green", "fill", "single")) {
  mode <- match.arg(mode)

  # ── Apply the sheet's schedule configuration ─────────────────────────────
  if (is.null(sheet) && length(TIMEOFF_SHEETS)) sheet <- TIMEOFF_SHEETS[1]
  cfg <- if (!is.null(sheet)) SHEET_CONFIGS[[sheet]] else NULL
  if (!is.null(cfg)) {
    SCHEDULE_START <<- cfg$schedule_start
    SCHEDULE_END   <<- cfg$schedule_end
    PAY_PERIODS    <<- cfg$pay_periods
    HOLIDAYS       <<- cfg$holidays
    HOLIDAY_DATES  <<- if (!is.null(cfg$holiday_dates)) cfg$holiday_dates
                       else as.Date(names(cfg$holidays))
    HOLIDAY_NAMES  <<- cfg$holiday_names
    if (verbose)
      message(sprintf("Sheet '%s': %s to %s (%d days, %d pay periods)",
                      sheet, SCHEDULE_START, SCHEDULE_END,
                      as.integer(SCHEDULE_END - SCHEDULE_START) + 1L,
                      nrow(PAY_PERIODS)))
  } else if (!is.null(sheet)) {
    warning(sprintf("No SHEET_CONFIGS entry for '%s' - using the date range currently set.", sheet))
  }

  if (verbose) message("Parsing time-off data...")
  time_off <- parse_time_off(path, sheet = sheet)

  if (verbose) message("Computing per-PP targets...")
  targets  <- compute_targets(time_off)
  if (verbose) green_supply_report(time_off, targets)

  if (verbose) message(sprintf("Solver budget: %g seconds per solve.", time_limit))
  sched <- SchedulerLP$new(time_off, targets)
  if (mode == "single") {
    if (verbose) message("Running ILP scheduler (single solve, green-preferring)...")
    sched$run(two_stage = two_stage, greedy_stage2 = greedy_stage2, phase = 2L,
              time_limit = time_limit)
  } else {
    if (verbose) message("Phase 1: building the green-only schedule...")
    sched$run_green(two_stage = two_stage, greedy_stage2 = greedy_stage2,
                    time_limit = time_limit)
    if (mode == "fill") {
      if (verbose) message("Phase 2: filling the remainder...")
      sched$run_fill(two_stage = two_stage, greedy_stage2 = greedy_stage2,
                     time_limit = time_limit)
    }
  }

  partial <- (mode == "green")
  if (verbose) message("Validating...")
  validation <- validate_schedule(sched, time_off, targets, partial = partial)
  if (verbose) print_validation(validation)

  green <- if (verbose) green_summary(sched, time_off, targets)
           else green_summary(sched, time_off, targets, quiet = TRUE)

  list(
    sched      = sched,
    time_off   = time_off,
    targets    = targets,
    validation = validation,
    green      = green,
    holes      = sched$holes_df(),
    mode       = mode,
    df         = sched$to_dataframe(),
    grid       = sched$to_person_grid(time_off, targets)
  )
}
