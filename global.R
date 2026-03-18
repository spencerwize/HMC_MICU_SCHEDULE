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

# ── Default time-off source ───────────────────────────────────────────────────
# Set env var TIMEOFF_SOURCE to a Google Sheets URL (or bare sheet ID) to pull
# live data without any file upload.  Falls back to the local XLSX if unset.
TIMEOFF_DEFAULT_SOURCE <- Sys.getenv("TIMEOFF_SOURCE",
                                     unset = "Time_Off_Requests.xlsx")

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
