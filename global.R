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
  "R/validate.R",
  "R/excel_output.R"
)
for (f in r_files) source(f, local = FALSE)

# ── Google Sheets source ──────────────────────────────────────────────────────
TIMEOFF_GSHEET_URL <- "https://docs.google.com/spreadsheets/d/1pjme5ne-O7XdM4aDfliJo65QjcdjbHebk87jmhIX3VI/edit"

# Known tab names in the Google Sheet — pre-populates the sheet dropdown so
# it works immediately without API auth. Add/remove names to match the workbook.
TIMEOFF_SHEETS <- c("Present")

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
