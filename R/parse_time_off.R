# ─────────────────────────────────────────────────────────────────────────────
# parse_time_off.R  —  Text-value parser supporting Google Sheets,
#                       local XLSX, and local CSV
#
# ── Google Sheet layout ───────────────────────────────────────────────────────
#
#   Row 1 (header):  Date | Katie | John | Hayden | Todd | Caroline | ...
#   Row 2+:          <date> | <value> | <value> | ...
#
#   • Column order doesn't matter — columns are matched by header name.
#   • The date column must be named "Date" (case-insensitive).
#   • Blank cell → available to work (no entry recorded).
#
# ── Cell values (case-insensitive) ───────────────────────────────────────────
#
#   "cme", "conf", "conference"          → cme  (credited, no shift target)
#   "vac", "vacation", or VAC_KEYWORDS   → vac  (may become PTO)
#   anything else non-blank              → off  (plain off day)
#
# ── Auth for private Google Sheets ───────────────────────────────────────────
#
#   Option A — Service account (recommended for deployed apps):
#     Set env var GS_SERVICE_ACCOUNT_JSON to the path of the JSON key file,
#     or paste the entire JSON string into the env var directly.
#     Shinyapps.io: add the env var in App → Settings → Environment Variables.
#
#   Option B — Sheet published to web (no credentials needed):
#     In Google Sheets: File → Share → Publish to web → Sheet → CSV
#     Pass the resulting URL as `path`.  No GS_SERVICE_ACCOUNT_JSON needed.
#
# Returns: named list  person → data.frame(date = Date, type = chr)
#   type values: "off" | "vac" | "cme"
# ─────────────────────────────────────────────────────────────────────────────

parse_time_off <- function(path, sheet = NULL) {

  empty_df <- function()
    data.frame(date = as.Date(character()), type = character(),
               stringsAsFactors = FALSE)
  result <- setNames(lapply(STAFF, function(p) empty_df()), STAFF)

  raw <- read_timeoff_source(path, sheet)
  if (is.null(raw) || nrow(raw) == 0) return(result)

  # ── Locate Date column ────────────────────────────────────────────────────
  hdr      <- colnames(raw)
  date_col <- which(tolower(trimws(hdr)) == "date")
  if (length(date_col) == 0)
    date_col <- which(vapply(raw, looks_like_dates, logical(1L)))[1]
  if (length(date_col) == 0 || is.na(date_col)) {
    message("WARNING: no 'Date' column found in time-off source.")
    return(result)
  }

  # ── Parse dates ────────────────────────────────────────────────────────────
  dates <- coerce_dates(raw[[date_col]])

  in_window <- !is.na(dates) &
    dates >= SCHEDULE_START & dates <= SCHEDULE_END
  if (!any(in_window)) {
    message("WARNING: no rows in time-off source fall within ",
            format(SCHEDULE_START, "%b %d"), " – ",
            format(SCHEDULE_END,   "%b %d"), ".")
    return(result)
  }
  dates <- dates[in_window]
  raw   <- raw[in_window, , drop = FALSE]

  # ── Match staff columns by name ───────────────────────────────────────────
  hdr_lc <- tolower(trimws(hdr))
  for (person in STAFF) {
    col_idx <- which(hdr_lc == tolower(person))
    if (length(col_idx) == 0) {
      message("NOTE: no column for '", person, "' in time-off source.")
      next
    }
    vals <- as.character(raw[[col_idx[1L]]])
    for (i in seq_along(dates)) {
      type <- classify_cell(vals[i])
      if (!is.na(type))
        result[[person]] <- rbind(result[[person]],
          data.frame(date = dates[i], type = type, stringsAsFactors = FALSE))
    }
  }

  for (p in STAFF)
    result[[p]]$date <- as.Date(result[[p]]$date, origin = "1970-01-01")

  result
}

# ── Source dispatcher ────────────────────────────────────────────────────────

read_timeoff_source <- function(path, sheet) {
  if (is_sheets_url(path)) return(read_from_gsheets(path, sheet))
  if (!file.exists(path)) {
    message("WARNING: '", path, "' not found — proceeding with no time-off data.")
    return(NULL)
  }
  ext <- tolower(tools::file_ext(path))
  if (ext == "csv") {
    read.csv(path, header = TRUE, stringsAsFactors = FALSE,
             check.names = FALSE, na.strings = c("", "NA"))
  } else {
    sh <- if (!is.null(sheet)) sheet else 1L
    tryCatch(
      openxlsx::read.xlsx(path, sheet = sh, colNames = TRUE,
                          detectDates = TRUE, na.strings = c("", "NA")),
      error = function(e) {
        message("WARNING: could not read '", path, "': ", conditionMessage(e))
        NULL
      }
    )
  }
}

# ── Google Sheets reader ─────────────────────────────────────────────────────

is_sheets_url <- function(x) {
  grepl("docs\\.google\\.com/spreadsheets", x) ||
  grepl("^1[A-Za-z0-9_-]{20,}$", x)   # raw sheet ID
}

read_from_gsheets <- function(url_or_id, sheet) {
  if (!requireNamespace("googlesheets4", quietly = TRUE))
    stop("Install the 'googlesheets4' package to read Google Sheets directly:\n",
         "  install.packages('googlesheets4')")

  gs4_auth_auto()

  sh <- if (!is.null(sheet)) sheet else 1L
  message("Fetching time-off data from Google Sheets...")
  tryCatch({
    df <- googlesheets4::read_sheet(url_or_id, sheet = sh,
                                    col_types = "c")   # everything as character
    as.data.frame(df, stringsAsFactors = FALSE)
  }, error = function(e) {
    message("ERROR reading Google Sheet: ", conditionMessage(e))
    NULL
  })
}

#' Authenticate with googlesheets4.
#' Tries (in order):
#'   1. GS_SERVICE_ACCOUNT_JSON env var  (path to JSON file, or JSON string)
#'   2. GOOGLE_APPLICATION_CREDENTIALS env var  (standard ADC path)
#'   3. gs4_deauth()  — no auth, works for publicly published sheets
gs4_auth_auto <- function() {
  svc_json <- Sys.getenv("GS_SERVICE_ACCOUNT_JSON", unset = "")
  adc_path <- Sys.getenv("GOOGLE_APPLICATION_CREDENTIALS", unset = "")

  if (nzchar(svc_json)) {
    cred <- if (file.exists(svc_json)) svc_json else
              jsonlite::fromJSON(svc_json)
    googlesheets4::gs4_auth(path = cred)
    return(invisible(NULL))
  }
  if (nzchar(adc_path) && file.exists(adc_path)) {
    googlesheets4::gs4_auth(path = adc_path)
    return(invisible(NULL))
  }
  googlesheets4::gs4_deauth()
}

# ── Cell-level helpers ───────────────────────────────────────────────────────

#' "cme" | "vac" | "off" | NA
classify_cell <- function(val) {
  if (is.na(val) || !nzchar(trimws(val))) return(NA_character_)
  v <- tolower(trimws(val))
  if (v %in% c("cme", "conf", "conference"))              return("cme")
  if (v %in% c("vac", "vacation") ||
      any(vapply(VAC_KEYWORDS,
                 function(kw) grepl(kw, v, fixed = TRUE),
                 logical(1L))))                            return("vac")
  "off"
}

looks_like_dates <- function(col) {
  non_na <- col[!is.na(col)]
  length(non_na) > 0 && (
    inherits(non_na, "Date") ||
    all(grepl("^\\d{1,4}[/-]\\d{1,2}[/-]\\d{2,4}$", as.character(non_na)))
  )
}

coerce_dates <- function(x) {
  if (inherits(x, "Date"))   return(x)
  if (is.numeric(x))         return(as.Date(x, origin = "1899-12-30"))
  fmts <- c("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y",
            "%m-%d-%Y", "%m-%d-%y", "%B %d, %Y", "%b %d, %Y")
  out <- rep(NA_real_, length(x))
  for (fmt in fmts) {
    need <- is.na(out)
    if (!any(need)) break
    parsed      <- suppressWarnings(as.Date(as.character(x)[need], format = fmt))
    out[need]   <- as.numeric(parsed)
  }
  as.Date(out, origin = "1970-01-01")
}
