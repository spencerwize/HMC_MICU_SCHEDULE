# HMC MICU APP Schedule — Project Context

## What this is
A Shiny app that builds a 12-hour rotating shift schedule for ~10–12 APP staff
at HMC MICU covering a ~14-week date range (two sheet configs: Apr 13 – Jul 19
and Jul 20 – Oct 25, 2026). Time-off data comes from a Google Sheet (with a
local `Time_Off_Requests.xlsx` fallback).

## Stack
- R / Shiny (`global.R`, `ui.R`, `server.R`)
- Core logic in `R/` subdirectory
- Solver: HiGHS MIP via the `highs` CRAN package (sparse `Matrix` constraints)
- Output: interactive Shiny UI + downloadable Excel workbook (`openxlsx`)

## Key files
| File | Purpose |
|---|---|
| `R/constants.R` | STAFF, PAY_PERIODS, HOLIDAYS, solver knobs, palette, helper fns |
| `R/parse_time_off.R` | Reads Google Sheet / XLSX / CSV → named list of time-off per person |
| `R/targets.R` | Computes per-person per-PP shift targets (sched_target, soft_min) |
| `R/scheduler_lp.R` | ILP scheduler (R6 `SchedulerLP`): model build, relaxation tiers, greedy aesthetic pass |
| `R/validate.R` | Hard-constraint checker |
| `R/excel_output.R` | Builds formatted Excel workbook |
| `global.R` | Loads packages, sources R/ files, defines SHEET_CONFIGS + run_pipeline() |
| `server.R` | Shiny server — `SchedulerLP$new(...)$run()` on Generate |
| `run_schedule.R` | Standalone CLI runner (no Shiny) |

## How the solve works
1. **Two-stage lexicographic solve** (default): Stage 1 is a fairness-only ILP
   (aesthetic auxiliaries suppressed); Stage 2 is a greedy local-search swap
   pass that improves run lengths while preserving every person's night /
   weekend / total / per-PP counts exactly.
2. **Relaxation cascade**: 7 tiers — night band 8–12 → 7–12 → 6–11, then 4
   "nuclear" fallbacks that progressively relax PP caps, C8b, C11b, C11c,
   per-PP minimums, and finally the fairness spread. The safety rules
   (C7/C7b/C8/C9/C10/C10b/C10c) are built unconditionally and never relaxed.
3. Nights are never hard-required: an unstaffed night costs
   `UNSTAFFED_NIGHT_PEN` in the objective instead.

## Solver performance notes
- The LP relaxation sits well above the integer optimum, so the MIP-gap stop
  criteria rarely trigger — `SOLVER_TIME_LIMIT` (constants.R) is the effective
  stopping rule and HiGHS returns its best incumbent when it hits.
- Constraint triplets are accumulated in lists and `unlist()`ed once
  (`add_con` in `build_and_solve`) — do not revert to `c(vec, ...)` appends.
- `count_solutions()` re-solves the ILP per enumeration step and passes its
  own short `time_limit_override` (default 120 s per solve).

## Git branch
Main branch for PRs: `claude/app-shift-schedule-Kk4RJ`
