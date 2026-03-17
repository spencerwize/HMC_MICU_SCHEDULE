#!/usr/bin/env Rscript
# ─────────────────────────────────────────────────────────────────────────────
# run_schedule.R  —  Standalone script (no Shiny needed)
#
# Usage:
#   Rscript run_schedule.R
#   Rscript run_schedule.R --input Time_Off_Requests.xlsx --output MySchedule.xlsx
# ─────────────────────────────────────────────────────────────────────────────

args <- commandArgs(trailingOnly = TRUE)

get_arg <- function(flag, default) {
  idx <- which(args == flag)
  if (length(idx) > 0 && idx < length(args)) args[idx + 1] else default
}

xlsx_input  <- get_arg("--input",  "Time_Off_Requests.xlsx")
output_path <- get_arg("--output", "MICU_APP_Schedule_R_2026.xlsx")

# Load everything
source("global.R")

cat("─────────────────────────────────────────────────────────\n")
cat("HMC MICU APP Shift Schedule Builder  (R version)\n")
cat("Apr 13 – Jul 19, 2026  |  PP8 – PP14\n")
cat("─────────────────────────────────────────────────────────\n\n")

result <- run_pipeline(xlsx_path = xlsx_input, verbose = TRUE)

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

cat("\nBuilding Excel output...\n")
build_excel(result$sched, result$time_off, result$targets, output_path)

cat("\nDone →", output_path, "\n")
