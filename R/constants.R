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

# Per-person soft minimum shifts per pay period.
# People listed here are deprioritised once they reach their soft floor but can
# still receive up to 6 shifts if slots are easy to fill.  Everyone else has
# soft_min == sched_target (i.e., the scheduler treats their target as firm).
FLEX_TARGETS <- list(
  Todd = 4L
)

# ILP solver wall-clock budget per solve (seconds).
# NOTE: the LP relaxation of this model sits well above the integer optimum, so
# the MIP gap criteria below often never trigger — this time limit is the
# EFFECTIVE stopping rule for most runs. HiGHS returns the best incumbent found
# when the limit is hit. Incumbent quality typically plateaus well within an
# hour; raise this only if validation shows the solution is still improving.
SOLVER_TIME_LIMIT <- 60*60
# Stop early when best integer solution is within this fraction of the LP bound.
# The LP relaxation is inherently ~8-9% above the integer optimum for this problem
# (fractional person-days + night-spread auxiliaries all relax to 1.0), so a gap
# of 0.09 accepts the integer optimum without wasting time chasing the LP ceiling.
SOLVER_MIP_GAP   <- 0.1
# Stage-1 (fairness) gap for the two-stage lexicographic solve. Tighter than the
# overall gap so the fairness target we pin in Stage 2 is genuinely near-optimal.
SOLVER_MIP_GAP_STAGE1 <- 0.02

# ABSOLUTE gap for the green-only phase. That phase's objective is dominated by
# the unfilled-slot penalties (100 per missing APP1/Night, 80 per APP2), so with
# holes present the objective is NEGATIVE and a *relative* gap is meaningless -
# HiGHS reports things like "objective = -1172.9 (max: -916.4)" and a 2% relative
# criterion never triggers sensibly. An absolute gap is directly interpretable:
# 40 means "within half a missing shift of proven optimal".
SOLVER_MIP_ABS_GAP_GREEN <- 40
# Roaming (APP3) fill weight in the objective. The Roaming slot is optional; this
# positive weight makes the ILP itself fill feasible Roaming slots (up to PP caps
# and density caps) instead of relying on a post-solve greedy pass. Kept below the
# fairness-spread coefficients so it shapes coverage without distorting fairness.
ROAM_FILL_WEIGHT <- 1.5
# Number of solver threads. HiGHS defaults to 1 thread; using all-but-one core
# lets it parallelise the branch-and-bound search and find a first incumbent
# faster on this large MIP. Falls back to 1 if core detection fails.
SOLVER_THREADS   <- tryCatch(max(1L, parallel::detectCores() - 1L),
                             error = function(e) 1L)
# How hard HiGHS works on primal feasibility heuristics before/while branching
# (0–1, default 0.05). Higher = finds a first incumbent faster on this large,
# hard-to-seed MIP — the closest available substitute for a warm start.
SOLVER_HEURISTIC_EFFORT <- 0.4
# Soft-minimum total shift counts per person across the FULL schedule.
# The solver penalises falling below these floors in the objective but they
# are not hard constraints — availability/vacation may prevent reaching them.
MIN_NIGHTS_SOFT_TOTAL <- 9L   # total night shifts per person
MIN_WKND_SOFT_TOTAL   <- 11L  # total Saturday + Sunday shifts per person

# Hard per-person bounds on total night shifts across the full schedule.
MIN_NIGHTS_HARD <- 8L
MAX_NIGHTS_HARD <- 12L

# Hard per-person bounds on total weekend (Sat+Sun) shifts across the full schedule.
MIN_WKND_HARD <- 8L
MAX_WKND_HARD <- 13L

# Penalty per UNSTAFFED night in the objective. Nights are never hard-required —
# any night may be left empty — but this heavy penalty makes the solver staff
# every night it feasibly can, only leaving one empty when no eligible person can
# legally cover it. Set well above the fairness/aesthetic coefficients so coverage
# always wins over distribution niceties.
UNSTAFFED_NIGHT_PEN <- 100

# ── Green-first scheduling weights ────────────────────────────────────────────
# GREEN-ONLY PHASE. APP1/APP2 relax from hard equalities to soft coverage, so an
# unfilled day slot needs a penalty. Both sit far above every fairness
# coefficient (max 5) so coverage is settled before distribution; APP1 is priced
# above APP2 so that when only one can be staffed, APP1 wins.
UNFILLED_APP1_PEN <- 100   # matches UNSTAFFED_NIGHT_PEN - the primary day slot
UNFILLED_APP2_PEN <- 80

# FILL PHASE. Bonus for placing a shift on a day the person asked to work.
# Must sit strictly between ROAM_FILL_WEIGHT (1.5, so it actually bites) and the
# tightest fairness term - the total-shift spread, effective price 4 (so it can
# never buy a green day by widening a spread). See the ladder documented at the
# objective in scheduler_lp.R.
GREEN_WORK_BONUS <- 3.0

# Penalty on the WORST person's yellow-day count (I_MAX_YELLOW). No matching
# +min term, so the effective price is 4, not 8: above the weekend-spread
# coefficient (3) so a flatter yellow distribution wins ties, below the
# night-spread coefficient (5) so it never outranks night fairness.
MAX_YELLOW_PEN <- 4.0

# Greedy Stage-2 swap: value of moving one shift onto a requested-work day,
# scored against run_pts = c(-10, 3, 8, 5). At 6, a net green day buys a
# 3-run -> 2-run reshape (-5) but cannot buy 3-run -> 2-run + solo (-15) nor
# block a solo -> 2-run repair (+13). Above ~13 the pass starts manufacturing
# solo shifts to chase green days.
GREEN_SWAP_PTS <- 6

# ── Weekend pairing ───────────────────────────────────────────────────────────
# Penalty for working exactly ONE of a Saturday/Sunday pair. Working a whole
# weekend costs one weekend; working a lone Saturday and, another week, a lone
# Sunday costs two - the same shift count feels like twice the weekends.
#
# Priced between ROAM_FILL_WEIGHT (1.5) and the total-shift spread (4), so it
# outranks optional Roaming placement but never buys a split weekend at the cost
# of fairness. Split weekends stay possible where coverage demands them.
WEEKEND_SPLIT_PEN <- 2.5

# Fill phase, SOFT-pin fallback. The fill phase first tries to hold every
# green-only assignment as a hard pin. That is not always possible: the
# green-only phase optimises green fill without knowing which shifts the fill
# phase will still need to place, and the adjacency rules (C7/C7b/C8/C8b) plus
# hard APP1/APP2 coverage can leave no completion. Rather than fail, the fill
# phase re-runs with the pins as a strong objective PREFERENCE instead.
#
# 6.0 stacks on top of GREEN_WORK_BONUS (every green-only assignment is by
# definition on a green day), so keeping one is worth 9.0 - above every other
# term except the night spread (10) and an unstaffed night (100). Fairness still
# wins, which is deliberate: a schedule should not become lopsided to preserve a
# provisional assignment.
PIN_KEEP_BONUS <- 6.0

# Default soft minimum shifts per PP for anyone not listed in FLEX_TARGETS.
# Scheduler will try to reach sched_target (6) but only hard-enforces this floor.
DEFAULT_SOFT_MIN <- 5L

# Per-person base shift target overrides (default 6 for everyone not listed).
# Value can be a single integer (applies to all PPs) or a named list of
# PP-specific overrides — unlisted PPs fall back to 6.
BASE_TARGETS <- list()

VAC_KEYWORDS <- c("vac", "hawaii", "galapagos", "trip", "vacation", "travel")

# ── Requested-WORK ("green") day keywords ─────────────────────────────────────
# A cell matching one of these marks a day the person WANTS to work. Checked
# AFTER the CME and VAC patterns, so "work trip" still classifies as vacation.
#
# "yes"/"y" are deliberately absent: in a sheet whose entire history is "time
# off", a bare "yes" most plausibly means "yes, I want this day off".
WORK_KEYWORDS <- c("w", "work", "works", "working", "green",
                   "prefer", "available", "avail")

# ── Explicit requested-OFF keywords ───────────────────────────────────────────
# Not used for classification (any unrecognised non-blank cell already falls
# through to "off"). These exist so the parser can tell a DELIBERATE off-request
# from an unrecognised value — likely a typo — and warn about the latter.
# "red" is the sheet's own vocabulary (see the Rules tab: Red = definitely
# cannot work). Listing it explicitly means it classifies deliberately rather
# than by fallthrough, so it is not reported as an unrecognised value.
OFF_KEYWORDS <- c("off", "no", "unavail", "unavailable", "ooo", "out",
                  "x", "red")

# ── Requested-PTO keywords ────────────────────────────────────────────────────
# A cell marked PTO is a day the person will not work (blocked, like Red) that
# ALSO counts as a shift toward their pay-period target, like CME. It feeds the
# per-period PTO total as  max(explicit PTO days, pto_reduction(off+vac+pto)),
# so marking PTO can never yield fewer PTO days than the automatic formula
# would have granted anyway. See compute_targets().
PTO_KEYWORDS <- c("pto", "paid time off")

# ── Which request types remove a day from scheduling entirely ─────────────────
# Single source of truth for "blocked". Everything that tests availability -
# the ILP bounds, the greedy pass, validation, the Excel export - reads this.
BLOCKED_TYPES <- c("off", "vac", "cme", "pto")

# ── Staff who appear in the WORKBOOK but not in the solve ─────────────────────
# New hires who have not started: they get a column on the Schedule sheet, a row
# in the Summary, and an entry in every dropdown, so shifts can be assigned to
# them by hand as they onboard. They are NOT scheduled by the solver and are not
# expected to be in the request sheet.
#
# Their pay-period target is 0 - they owe nothing yet - so an empty period reads
# "0/0" rather than showing as a shortfall. Assigning a shift makes it "1/0",
# flagged blue (over target), which is the cue to give them a real target once
# they actually start.
#
# Anyone here who later appears in the request sheet is picked up from the sheet
# instead and silently dropped from this list, so a name can be left here across
# the transition without creating a duplicate column.
PENDING_STAFF <- c("Matt", "Jen", "Kennedy", "Preet", "Rose", "Katia")

# ── Explicitly-neutral ("yellow") keywords ────────────────────────────────────
# The sheet's Rules tab defines Yellow as "Avoid Assigning Work If Possible" -
# i.e. schedulable, just not requested. That is exactly what a BLANK cell means,
# so these classify to NA (no row recorded) rather than to a blocked day.
#
# Getting this wrong is expensive: without it "Yellow" falls through to "off"
# and every yellow day becomes a hard block, which is the opposite of "avoid if
# possible".
NEUTRAL_KEYWORDS <- c("yellow", "maybe", "prefer not", "rather not")

# ── What an EMPTY cell means ──────────────────────────────────────────────────
# The sheet's Rules tab defines Green / Red / Yellow / CME but says nothing about
# a blank cell. Two defensible readings:
#   "green"   - no constraint expressed, so the person is available to work.
#   "neutral" - same as Yellow: schedulable, but avoided when possible.
# Set to "green": a blank is an absence of any restriction, not a soft refusal.
#
# Consequence worth knowing: someone who has not filled in the sheet at all
# becomes fully available rather than fully avoided. parse_time_off() reports a
# per-person blank count so that is visible rather than silent.
BLANK_CELL_MEANS <- "green"

# ── Excel / UI color palette (ARGB hex strings) ───────────────────────────────
CLR_GREEN      <- "FF92D050"  # day shift
CLR_BLUE       <- "FFBDD7EE"  # night shift
CLR_PEACH      <- "FFFFD966"  # vacation
CLR_PINK       <- "FFFF99CC"  # PTO
CLR_ORANGE     <- "FFFF6D01"  # CME (fill)
CLR_YELLOW_HL  <- "FFFFFF99"  # holiday
CLR_LIGHT_RED  <- "FFFFC7CE"  # off day
CLR_GREEN_REQ  <- "FF2E7D32"  # requested-work ("green") day — OUTLINE only.
                              # Deliberately not a fill: CLR_GREEN (#92D050) is
                              # already "day shift" and CLR_YELLOW_HL (#FFFF99)
                              # is already "holiday", so a literal green/yellow
                              # fill would collide. The outline composes with the
                              # role fill, showing who works AND whether they
                              # asked for the day in a single cell.
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
  dbn_border = "#FF6D01",  # day-before-night orange dashed
  green_req  = "#2E7D32",  # requested-work day outline (see CLR_GREEN_REQ)
  hole       = "#FFC7CE"   # unfilled slot after the green-only phase
)

# ── Helper functions ──────────────────────────────────────────────────────────

#' Vectorised PP lookup: PP name per date, NA_character_ where out of range.
#' findInterval against the sorted PP starts, then bounds-check against the
#' matched PP's end (PPs are contiguous but the schedule may extend past them).
get_pp_vec <- function(dates) {
  dates <- as.Date(dates, origin = "1970-01-01")
  idx   <- findInterval(as.numeric(dates), as.numeric(PAY_PERIODS$start))
  out   <- rep(NA_character_, length(dates))
  hit   <- !is.na(idx) & idx >= 1L
  hit[hit] <- dates[hit] <= PAY_PERIODS$end[idx[hit]]
  out[hit] <- PAY_PERIODS$name[idx[hit]]
  out
}

#' Return PP name for a single date, or NA_character_ if out of range
get_pp <- function(d) {
  if (is.null(d) || length(d) == 0 || is.na(d)) return(NA_character_)
  get_pp_vec(d)[1L]
}

#' Dates in a given PP
pp_dates <- function(pp_name) {
  row <- PAY_PERIODS[PAY_PERIODS$name == pp_name, ]
  if (nrow(row) == 0) return(as.Date(character()))
  seq(row$start, row$end, by = "day")
}

# Calendar shading only: which DATES fall on a weekend.
is_weekend <- function(d) {
  weekdays(d) %in% c("Saturday", "Sunday")
}

# Which SHIFTS count as weekend work. The weekend block runs Friday night
# through Sunday night:
#     Friday    - Night shift only  (a Friday DAY shift is a weekday shift)
#     Saturday  - every shift
#     Sunday    - every shift
# This is deliberately (date, slot)-dependent, not date-only: counting a Friday
# day shift as weekend work would overstate the load, and NOT counting the
# Friday night would understate it.
#
# `slot` accepts the slot names in SLOTS. Vectorised over both arguments.
is_weekend_shift <- function(d, slot) {
  wd <- weekdays(as.Date(d, origin = "1970-01-01"))
  (wd == "Friday" & slot == "Night") | (wd %in% c("Saturday", "Sunday"))
}

all_dates <- function() {
  seq(SCHEDULE_START, SCHEDULE_END, by = "day")
}
