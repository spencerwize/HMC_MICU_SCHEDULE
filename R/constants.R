# ─────────────────────────────────────────────────────────────────────────────
# constants.R  —  HMC MICU APP Schedule  Apr 13 – Jul 19, 2026
# ─────────────────────────────────────────────────────────────────────────────

# Startup default — overwritten dynamically by parse_time_off() from the
# sheet's column headers. Edit this only if you need a fallback when the
# sheet is unavailable at app startup (e.g. for the cal_person picker).
STAFF <- c("Katie", "John", "Hayden", "Todd", "Caroline",
           "Isabel", "Kristin", "Mandie", "Maureen", "Radha")

# Column index (1-based) in Time_Off_Requests.xlsx
STAFF_COL_MAP <- c(
  Todd = 4, Mandie = 5, Isabel = 7, Radha = 8,
  Maureen = 9, Katie = 10, Hayden = 11, Caroline = 12,
  John = 13, Kristin = 14
)

SCHEDULE_START <- as.Date("2026-04-13")
SCHEDULE_END   <- as.Date("2026-07-19")

PAY_PERIODS <- data.frame(
  name  = c("PP8","PP9","PP10","PP11","PP12","PP13","PP14"),
  start = as.Date(c("2026-04-13","2026-04-27","2026-05-11",
                    "2026-05-25","2026-06-08","2026-06-22","2026-07-06")),
  end   = as.Date(c("2026-04-26","2026-05-10","2026-05-24",
                    "2026-06-07","2026-06-21","2026-07-05","2026-07-19")),
  stringsAsFactors = FALSE
)

# Fixed holiday pre-seeds  (date string -> named list slot -> person)
HOLIDAYS <- list(
  "2026-05-25" = list(APP1 = "Hayden",  APP2 = "Todd",     Roaming = "Radha",    Night = "Isabel"),
  "2026-06-19" = list(APP1 = "Mandie",  APP2 = "Caroline", Roaming = "Radha",    Night = "Isabel"),
  "2026-07-04" = list(APP1 = "Kristin", APP2 = "John",     Roaming = "Caroline", Night = "Mandie")
)
HOLIDAY_DATES <- as.Date(names(HOLIDAYS))

# Startup default — overwritten by server.R from SHEET_CONFIGS
HOLIDAY_NAMES <- c("2026-05-25" = "Memorial Day",
                   "2026-06-19" = "Juneteenth",
                   "2026-07-04" = "July 4th")


SLOTS     <- c("APP1", "APP2", "Roaming", "Night")
DAY_SLOTS <- c("APP1", "APP2", "Roaming")

# Hard cap on total night shifts across the ENTIRE schedule (all PPs combined).
MAX_NIGHTS_TOTAL <- 12L

# Per-person soft minimum shifts per pay period.
# People listed here are deprioritised once they reach their soft floor but can
# still receive up to 6 shifts if slots are easy to fill.  Everyone else has
# soft_min == sched_target (i.e., the scheduler treats their target as firm).
FLEX_TARGETS <- list(
  Todd = 4L
)

# ILP solver wall-clock budget per TIER (seconds).
# This is a safety *ceiling*, not a target — on real data the MIP gap below stops
# the solver well before it.  But the full model (tier 1) is combinatorially hard
# to find a FIRST integer-feasible point for, so the ceiling must stay generous
# enough that tier 1 isn't abandoned for a weaker tier.  30 min is a 24× cut from
# the old 12h value — it bounds a runaway tier without sacrificing tier 1.
#   • The real speed levers are SOLVER_THREADS (HiGHS defaults to 1 thread!), the
#     tightened constraints, and mip_heuristic_effort (set in scheduler_lp.R).
#   • If hard instances still abandon tier 1 here, raise this; if real runs always
#     stop on the gap in a few minutes, you can safely lower it.
SOLVER_TIME_LIMIT <- 60*60*2
# Stop early when best integer solution is within this fraction of the LP bound.
# The LP relaxation is inherently ~8-9% above the integer optimum for this problem
# (fractional person-days + night-spread auxiliaries all relax to 1.0), so a gap
# of 0.09 accepts the integer optimum without wasting time chasing the LP ceiling.
SOLVER_MIP_GAP   <- 0.09

# Number of threads HiGHS may use for parallel branch-and-bound.
# 0 = autodetect (parallel::detectCores()); a positive integer pins the count.
SOLVER_THREADS <- 0L

# ── Fairness-polish phase (see SchedulerLP$run) ──────────────────────────────
# After the cascade finds a feasible tier, a second "polish" solve re-optimises
# only the fairness/streak/isolation terms with total coverage floored at the
# value already achieved.  Because coverage is dominated by hard equalities, the
# 9% coverage gap above can leave the small fairness coefficients loose; this
# phase tightens them.  It is always feasible (the phase-1 solution satisfies the
# floor) so it can only improve, never worsen, the schedule.
FAIRNESS_POLISH           <- TRUE
SOLVER_POLISH_TIME_LIMIT  <- 60      # seconds — short; coverage is already fixed
SOLVER_POLISH_MIP_GAP     <- 0.01    # tight — fairness-only objective has a small range
# Soft-minimum total shift counts per person across the FULL schedule.
# The solver penalises falling below these floors in the objective but they
# are not hard constraints — availability/vacation may prevent reaching them.
MIN_NIGHTS_SOFT_TOTAL <- 9L   # total night shifts per person
MIN_WKND_SOFT_TOTAL   <- 11L  # total Saturday + Sunday shifts per person

# Hard per-person bounds on total night shifts across the full schedule.
MIN_NIGHTS_HARD <- 8L
MAX_NIGHTS_HARD <- 12L

# Hard per-person bounds on total weekend (Sat+Sun) shifts across the full schedule.
MIN_WKND_HARD <- 7L
MAX_WKND_HARD <- 11L

# Default soft minimum shifts per PP for anyone not listed in FLEX_TARGETS.
# Scheduler will try to reach sched_target (6) but only hard-enforces this floor.
DEFAULT_SOFT_MIN <- 5L

# Per-person base shift target overrides (default 6 for everyone not listed).
# Value can be a single integer (applies to all PPs) or a named list of
# PP-specific overrides — unlisted PPs fall back to 6.
BASE_TARGETS <- list()

VAC_KEYWORDS <- c("vac", "hawaii", "galapagos", "trip", "vacation", "travel")

# ── Excel / UI color palette (ARGB hex strings) ───────────────────────────────
CLR_GREEN      <- "FF92D050"  # day shift
CLR_BLUE       <- "FFBDD7EE"  # night shift
CLR_PEACH      <- "FFFFD966"  # vacation
CLR_PINK       <- "FFFF99CC"  # PTO
CLR_ORANGE     <- "FFFF6D01"  # CME (fill)
CLR_YELLOW_HL  <- "FFFFFF99"  # holiday
CLR_LIGHT_RED  <- "FFFFC7CE"  # off day
CLR_WHITE      <- "FFFFFFFF"
CLR_HEADER     <- "FF203864"  # dark navy header
CLR_HEADER2    <- "FF2E75B6"  # medium blue sub-header
CLR_GRAY       <- "FFD9D9D9"
CLR_WEEKEND    <- "FFF2F2F2"

# CSS-friendly hex (no alpha prefix) for Shiny / reactable
UI_CLR <- list(
  day_shift  = "#92D050",
  night      = "#BDD7EE",
  vacation   = "#FFD966",
  pto        = "#FF99CC",
  cme        = "#FF6D01",
  holiday    = "#FFFF99",
  off        = "#FFC7CE",
  weekend_bg = "#F2F2F2",
  header     = "#203864",
  header2    = "#2E75B6",
  gray       = "#D9D9D9",
  white      = "#FFFFFF",
  dbn_border = "#FF6D01"   # day-before-night orange dashed
)

# ── Helper functions ──────────────────────────────────────────────────────────

#' Return PP name for a given date, or NA_character_ if out of range
get_pp <- function(d) {
  if (is.null(d) || length(d) == 0 || is.na(d)) return(NA_character_)
  for (i in seq_len(nrow(PAY_PERIODS))) {
    if (d >= PAY_PERIODS$start[i] && d <= PAY_PERIODS$end[i])
      return(PAY_PERIODS$name[i])
  }
  NA_character_
}

#' Vectorised get_pp
get_pp_vec <- function(dates) {
  vapply(dates, get_pp, character(1))
}

#' Dates in a given PP
pp_dates <- function(pp_name) {
  row <- PAY_PERIODS[PAY_PERIODS$name == pp_name, ]
  if (nrow(row) == 0) return(as.Date(character()))
  seq(row$start, row$end, by = "day")
}

is_weekend <- function(d) {
  weekdays(d) %in% c("Saturday", "Sunday")
}

all_dates <- function() {
  seq(SCHEDULE_START, SCHEDULE_END, by = "day")
}
