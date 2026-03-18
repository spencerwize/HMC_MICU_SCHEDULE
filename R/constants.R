# ─────────────────────────────────────────────────────────────────────────────
# constants.R  —  HMC MICU APP Schedule  Apr 13 – Jul 19, 2026
# ─────────────────────────────────────────────────────────────────────────────

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


SLOTS     <- c("APP1", "APP2", "Roaming", "Night")
DAY_SLOTS <- c("APP1", "APP2", "Roaming")

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
