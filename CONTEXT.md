# HMC MICU APP Schedule — Project Context

## What this is
A Shiny app (`/home/user/HMC_MICU_SCHEDULE`) that builds a 12-hour rotating
shift schedule for 10 APP staff at HMC MICU covering a ~97-day date range
(e.g. Apr 13 – Jul 19, 2026). Time-off data comes from a Google Sheet.

## Stack
- R / Shiny (`global.R`, `ui.R`, `server.R`)
- Core logic in `R/` subdirectory
- Output: interactive Shiny UI + downloadable Excel workbook

## Key files
| File | Purpose |
|---|---|
| `R/constants.R` | STAFF, PAY_PERIODS, HOLIDAYS, color palette, helper fns |
| `R/parse_time_off.R` | Reads Google Sheet / XLSX / CSV → named list of time-off per person |
| `R/targets.R` | Computes per-person per-PP shift targets (sched_target, soft_min) |
| `R/scheduler.R` | Greedy MRV scheduler (R6 `Scheduler` class) |
| `R/scheduler_lp.R` | **New** ILP-based scheduler (R6 `SchedulerLP` class, uses `lpSolveAPI`) |
| `R/validate.R` | Hard-constraint checker |
| `R/excel_output.R` | Builds formatted Excel workbook |
| `global.R` | Loads packages, sources R/ files, defines SHEET_CONFIGS |
| `server.R` | Shiny server — pipeline uses `SchedulerLP$new(...) + run()` |
| `install_packages.R` | One-time package installer (includes `lpSolveAPI`) |

## Schedule rules enforced
- 4 slots per day: APP1, APP2, Roaming (day shifts), Night
- APP1 and Night **must** be filled every day
- 6-shift target per person per pay period (sched_target can be lower if off/CME)
- Hard: no day shift the morning after a night shift (24h ban)
- Hard: 2-day day-shift recovery after last night in a streak
- Hard: ≤ 3 consecutive nights
- Hard: ≤ 4 consecutive working days
- Hard: ≤ 12 total night shifts across the entire schedule (MAX_NIGHTS_TOTAL)
- Holidays pre-seeded from `HOLIDAYS` constant in `constants.R`

## ILP scheduler (`R/scheduler_lp.R`) — current state
Replaces the greedy MRV multi-pass approach with a single ILP solve.

**Variables:**
- `x[p,d,s]` binary — person p assigned to slot s on day d (3,880 vars)
- `z[p,d,s]` continuous [0,1] — day-before-night linearisation (2,880 vars)

**Constraints (~12,100 rows):**
- C1 slot uniqueness, C2 APP1=1, C3 Night=1, C4 no double-booking
- C5 availability (set as bounds), C6 PP cap, C7/C8 night recovery
- C9 ≤3 consec nights, C10 ≤4 consec work days, C11 night total cap
- C12 holiday pre-seeds (fixed bounds), C13 DBN linearisation lower bound

**Objective (maximise):**
`4×Night + 4×APP1 + 2×APP2 + 2×Roaming − 8×z[day-before-night]`

**Solver:** `lpSolveAPI` (lp_solve 5.5), 5-min timeout, falls back to greedy
`Scheduler` if not installed or infeasible.

## Git branch
`claude/app-shift-schedule-Kk4RJ`  
Latest commit: `d2e62a6` — "Replace greedy scheduler with ILP-based SchedulerLP"

## What the user asked last
"Can you not just build a single script for the ILP model using everything
you know about this problem?" — suggesting they may want a **standalone**
`Rscript`-runnable ILP script (no Shiny, no R6 class) that is easier to
read, run, and debug independently of the app.

## Possible next steps
1. Write a standalone `run_schedule_ilp.R` script (flat, no R6) that solves
   the ILP and prints/exports results — good for debugging solver performance
2. Tune solver performance (warm-start from greedy, tighten LP relaxation)
3. Add a UI toggle between ILP and greedy so the user can compare
