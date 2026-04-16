# ─────────────────────────────────────────────────────────────────────────────
# global.R  —  Loaded by both Shiny (app startup) and run_schedule.R
# ─────────────────────────────────────────────────────────────────────────────

suppressPackageStartupMessages({
  library(R6)
  library(tidyxl)
  library(openxlsx)
  library(lubridate)
  library(dplyr)
  library(tidyr)
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
  "R/parse_time_off.R",
  "R/targets.R",
  "R/scheduler.R",
  "R/scheduler_lp.R",
  "R/validate.R",
  "R/excel_output.R"
)
for (f in r_files) source(f, local = FALSE)

# ── Google Sheets source ──────────────────────────────────────────────────────
TIMEOFF_GSHEET_URL <- "https://docs.google.com/spreadsheets/d/1pjme5ne-O7XdM4aDfliJo65QjcdjbHebk87jmhIX3VI/edit"

# Known tab names in the Google Sheet — pre-populates the sheet dropdown so
# it works immediately without API auth. Add/remove names to match the workbook.
TIMEOFF_SHEETS <- c("April13-July19", "July20-Oct25")

# Per-sheet schedule configurations.
# Each entry maps a sheet tab name to the date range, pay periods, holiday
# pre-seeds, and calendar month choices that apply to that period.
SHEET_CONFIGS <- list(
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
run_pipeline <- function(path    = TIMEOFF_DEFAULT_SOURCE,
                         verbose = TRUE) {
  if (verbose) message("Parsing time-off data...")
  time_off <- parse_time_off(path)

  if (verbose) message("Computing per-PP targets...")
  targets  <- compute_targets(time_off)

  if (verbose) message("Running scheduler...")
  sched    <- Scheduler$new(time_off, targets)
  sched$run()

  if (verbose) message("Validating...")
  validation <- validate_schedule(sched, time_off, targets)
  if (verbose) print_validation(validation)

  list(
    sched      = sched,
    time_off   = time_off,
    targets    = targets,
    validation = validation,
    df         = sched$to_dataframe(),
    grid       = sched$to_person_grid(time_off, targets)
  )
}
