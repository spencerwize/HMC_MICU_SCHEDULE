#!/usr/bin/env Rscript
# ─────────────────────────────────────────────────────────────────────────────
# run_schedule.R  —  Standalone script (no Shiny needed)
#
# Usage:
#   Rscript run_schedule.R
#   Rscript run_schedule.R --input Time_Off_Requests.xlsx --output MySchedule.xlsx
#   Rscript run_schedule.R --time-limit 1800 --mode green
#
# Flags:
#   --input       path or Google Sheet URL (default: the configured sheet)
#   --output      .xlsx to write
#   --sheet       workbook tab (default: first of TIMEOFF_SHEETS)
#   --mode        green (default) | fill | single
#   --time-limit  solver seconds per solve (default 3600)
# ─────────────────────────────────────────────────────────────────────────────

args <- commandArgs(trailingOnly = TRUE)

get_arg <- function(flag, default) {
  idx <- which(args == flag)
  if (length(idx) > 0 && idx < length(args)) args[idx + 1] else default
}

# NA -> resolved after global.R loads, to TIMEOFF_DEFAULT_SOURCE (the configured
# Google Sheet). The constants are not available until global.R is sourced below.
xlsx_input  <- get_arg("--input",  NA_character_)
output_path <- get_arg("--output", "MICU_APP_Schedule_R_2026.xlsx")
# --mode green  : green-only phase only (partial; holes reported)   DEFAULT
#        fill   : green-only phase, then completion
#        single : one green-preferring solve, no pinning
mode        <- get_arg("--mode",   "green")
# Tab of the request workbook; also picks the date range / pay periods.
sheet       <- get_arg("--sheet",  NA_character_)
# Solver budget in SECONDS per solve. This model does not converge on real data,
# so more time means a better schedule: measured 600s -> 27 shifts unfilled,
# 900s -> 14, 3600s -> 7.
time_limit  <- as.numeric(get_arg("--time-limit", "3600"))

# Load everything
source("global.R")

if (is.na(xlsx_input)) xlsx_input <- TIMEOFF_DEFAULT_SOURCE

cat("─────────────────────────────────────────────────────────\n")
cat("HMC MICU APP Shift Schedule Builder  (R version)\n")
cat("Apr 13 – Jul 19, 2026  |  PP8 – PP14\n")
cat("─────────────────────────────────────────────────────────\n\n")

result <- run_pipeline(path = xlsx_input, verbose = TRUE, mode = mode,
                       sheet = if (is.na(sheet)) NULL else sheet,
                       time_limit = time_limit)

cat("\nQuick shift summary:\n")
cat(sprintf("  %-10s  %4s  %6s  %6s  %s\n",
            "Person", "Days", "Nights", "Roaming", "Total"))
for (person in STAFF) {
  nights <- length(result$sched$person_nights[[person]])
  days   <- nrow(result$sched$person_shifts[[person]])
  roam   <- sum(result$sched$person_shifts[[person]]$slot == "Roaming")
  cat(sprintf("  %-10s  %4d  %6d  %6d  %d\n",
              person, days, nights, roam, days + nights))
}

if (nrow(result$holes) > 0) {
  hs <- result$holes[result$holes$kind == "slot", ]
  tg <- result$holes[result$holes$kind == "target", ]
  cat(sprintf("\nRemaining gaps: %d unfilled slot(s), %d person-PP(s) below target\n",
              nrow(hs), nrow(tg)))
  if (nrow(hs) > 0) {
    cat("  Unfilled slots:\n")
    for (i in seq_len(min(nrow(hs), 15L)))
      cat(sprintf("    %s  %s\n", hs$date[i], hs$slot[i]))
    if (nrow(hs) > 15L) cat(sprintf("    ... and %d more\n", nrow(hs) - 15L))
  }
}

cat("\nBuilding Excel output...\n")
build_excel(result$sched, result$time_off, result$targets, output_path)

cat("\nDone →", output_path, "\n")
