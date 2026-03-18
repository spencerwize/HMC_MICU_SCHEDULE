# ─────────────────────────────────────────────────────────────────────────────
# new_parse_time_off.R  —  Text-value parser for Google Sheets / simple XLSX
#
# Replaces the color-coded parse_time_off.R with a plain-text approach:
# staff just type a value in their column; no color coding needed.
#
# ── Expected Google Sheet layout ─────────────────────────────────────────────
#
#   Row 1  (header):  Date | Katie | John | Hayden | Todd | Caroline | ...
#   Row 2+:           <date> | <value> | ...
#
#   • Column order doesn't matter — columns are matched by header name.
#   • The date column must be named "Date" (case-insensitive).
#   • Dates can be in any of: YYYY-MM-DD, MM/DD/YYYY, MM/DD/YY, M/D/YY.
#   • Blank cell → no time-off entry (person is available).
#
# ── Cell values (case-insensitive, leading/trailing whitespace stripped) ──────
#
#   CME / credited day:   "cme", "conf", "conference"
#   Vacation:             "vac", "vacation", or any VAC_KEYWORDS match
#   Off (non-credited):   "off", "x", or any other non-blank, non-CME value
#
# ── To switch the app to this parser ─────────────────────────────────────────
#   In global.R, change:
#     source("R/parse_time_off.R", local = FALSE)
#   to:
#     source("R/new_parse_time_off.R", local = FALSE)
#   and rename new_parse_time_off → parse_time_off (or alias it below).
#
# ── Google Sheets export steps ───────────────────────────────────────────────
#   File → Download → Microsoft Excel (.xlsx)   ← preferred
#   File → Download → Comma Separated Values    ← also supported (pass .csv)
#
# Returns: named list  person → data.frame(date = Date, type = chr)
#   type values: "off" | "vac" | "cme"
# ─────────────────────────────────────────────────────────────────────────────

new_parse_time_off <- function(path,
                               sheet     = NULL,
                               skip_rows = 0L) {

  empty_df <- function()
    data.frame(date = as.Date(character()), type = character(),
               stringsAsFactors = FALSE)
  result <- setNames(lapply(STAFF, function(p) empty_df()), STAFF)

  if (!file.exists(path)) {
    message("WARNING: ", path, " not found — proceeding with no time-off data.")
    return(result)
  }

  ext <- tolower(tools::file_ext(path))

  # ── Read raw data ────────────────────────────────────────────────────────
  if (ext == "csv") {
    raw <- read.csv(path, header = TRUE, stringsAsFactors = FALSE,
                    check.names = FALSE, skip = skip_rows, na.strings = c("", "NA"))
  } else {
    # xlsx / xls — use openxlsx (already a project dependency)
    sh <- if (!is.null(sheet)) sheet else 1L
    raw <- tryCatch(
      openxlsx::read.xlsx(path, sheet = sh, startRow = 1L + skip_rows,
                          colNames = TRUE, detectDates = TRUE,
                          na.strings = c("", "NA")),
      error = function(e) {
        message("WARNING: could not read '", path, "': ", conditionMessage(e))
        NULL
      }
    )
    if (is.null(raw)) return(result)
  }

  if (nrow(raw) == 0) {
    message("WARNING: time-off sheet is empty.")
    return(result)
  }

  # ── Locate the Date column ───────────────────────────────────────────────
  hdr      <- colnames(raw)
  date_col <- which(tolower(trimws(hdr)) == "date")
  if (length(date_col) == 0) {
    # Fall back: first column that looks like dates
    date_col <- which(sapply(raw, function(col) {
      non_na <- col[!is.na(col)]
      length(non_na) > 0 && (
        inherits(non_na, "Date") ||
        all(grepl("^\\d{1,4}[/-]\\d{1,2}[/-]\\d{2,4}$",
                  as.character(non_na)))
      )
    }))[1]
  }
  if (is.na(date_col) || length(date_col) == 0) {
    message("WARNING: no 'Date' column found in time-off sheet.")
    return(result)
  }

  # ── Parse the date column ────────────────────────────────────────────────
  date_vec <- raw[[date_col]]
  if (inherits(date_vec, "Date")) {
    dates <- date_vec
  } else if (is.numeric(date_vec)) {
    # openxlsx sometimes returns Excel serial numbers
    dates <- as.Date(date_vec, origin = "1899-12-30")
  } else {
    dates <- parse_date_flexibly(as.character(date_vec))
  }

  # Keep only rows within the schedule window
  in_window <- !is.na(dates) &
    dates >= SCHEDULE_START &
    dates <= SCHEDULE_END
  if (!any(in_window)) {
    message("WARNING: no dates in time-off sheet fall within the schedule window ",
            format(SCHEDULE_START, "%b %d"), " – ", format(SCHEDULE_END, "%b %d"), ".")
    return(result)
  }
  dates <- dates[in_window]
  raw   <- raw[in_window, , drop = FALSE]

  # ── Match staff columns by name ──────────────────────────────────────────
  hdr_clean <- tolower(trimws(hdr))
  for (person in STAFF) {
    col_idx <- which(hdr_clean == tolower(person))
    if (length(col_idx) == 0) {
      message("NOTE: no column found for '", person, "' in time-off sheet.")
      next
    }
    col_idx <- col_idx[1]
    vals    <- as.character(raw[[col_idx]])

    for (i in seq_along(dates)) {
      type <- classify_cell(vals[i])
      if (!is.na(type)) {
        result[[person]] <- rbind(result[[person]],
          data.frame(date = dates[i], type = type,
                     stringsAsFactors = FALSE))
      }
    }
  }

  # Ensure Date class is preserved after rbind
  for (p in STAFF)
    result[[p]]$date <- as.Date(result[[p]]$date, origin = "1970-01-01")

  result
}

# ── Helpers ──────────────────────────────────────────────────────────────────

#' Classify a single cell text value → "cme" | "vac" | "off" | NA
classify_cell <- function(val) {
  if (is.na(val) || !nzchar(trimws(val))) return(NA_character_)
  v <- tolower(trimws(val))
  if (v %in% c("cme", "conf", "conference")) return("cme")
  if (v %in% c("vac", "vacation") ||
      any(sapply(VAC_KEYWORDS, function(kw) grepl(kw, v, fixed = TRUE))))
    return("vac")
  "off"   # anything else non-blank → regular off day
}

#' Try several common date formats; return Date vector (NA where unparseable)
parse_date_flexibly <- function(x) {
  fmts <- c("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%m-%d-%Y", "%m-%d-%y",
            "%d/%m/%Y", "%B %d, %Y", "%b %d, %Y")
  out <- rep(NA_real_, length(x))
  for (fmt in fmts) {
    still_na <- is.na(out)
    if (!any(still_na)) break
    parsed <- suppressWarnings(as.Date(x[still_na], format = fmt))
    out[still_na] <- as.numeric(parsed)
  }
  as.Date(out, origin = "1970-01-01")
}
