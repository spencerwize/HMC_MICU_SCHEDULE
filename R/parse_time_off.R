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
#   • Blank cell → neutral 'yellow' day: schedulable, not requested.
#
# ── Cell values (case-insensitive) ───────────────────────────────────────────
#
#   "cme", "conf", "conference"          → cme  (credited, no shift target)
#   "vac", "vacation", or VAC_KEYWORDS   → vac  (may become PTO)
#   "w", "work", "green", or WORK_KEYWORDS → green (REQUESTED WORK day)
#   anything else non-blank                 → off   (plain off day)
#
#   Order matters: CME beats VAC beats WORK, so "CME trip" is cme and
#   "work trip" is vac. Unrecognised values become "off" (the safe default)
#   and are reported as a typo warning at parse time.
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
#   type values: "off" | "vac" | "cme" | "green"
#   A date with NO row is a neutral "yellow" day.
# ─────────────────────────────────────────────────────────────────────────────

parse_time_off <- function(path, sheet = NULL) {

  empty_df <- function()
    data.frame(date = as.Date(character()), type = character(),
               stringsAsFactors = FALSE)

  raw <- read_timeoff_source(path, sheet)

  # If the remote source failed, fall back to the local XLSX
  if ((is.null(raw) || nrow(raw) == 0) &&
      (is_sheets_url(path) || is_sheets_csv_url(path))) {
    local_xlsx <- "Time_Off_Requests.xlsx"
    if (file.exists(local_xlsx)) {
      message("Remote source unavailable — falling back to ", local_xlsx)
      raw <- read_timeoff_source(local_xlsx, sheet)
    }
  }

  if (is.null(raw) || nrow(raw) == 0)
    return(setNames(lapply(STAFF, function(p) empty_df()), STAFF))

  # ── Locate Date column ────────────────────────────────────────────────────
  hdr      <- colnames(raw)
  date_col <- which(tolower(trimws(hdr)) == "date")
  if (length(date_col) == 0)
    date_col <- which(vapply(raw, looks_like_dates, logical(1L)))[1]
  if (length(date_col) == 0 || is.na(date_col)) {
    message("WARNING: no 'Date' column found in time-off source.")
    return(setNames(lapply(STAFF, function(p) empty_df()), STAFF))
  }

  # ── Detect staff dynamically from column headers ──────────────────────────
  # Every non-date, non-empty column is a staff member. This updates the
  # global STAFF so the scheduler, validator, and UI all stay in sync with
  # whoever is actually listed in the sheet — no code changes needed when
  # people are added or removed.
  detected <- trimws(hdr[-date_col])
  # Drop columns that are not people:
  #   • blank headers
  #   • tibble/readr name-repair placeholders ("...1", "...2", ...) - the sheet
  #     has an unnamed first column holding day names, which would otherwise be
  #     detected as an 11th staff member whose every cell reads as a day OFF
  #   • explicit day/weekday label columns
  detected <- detected[nzchar(detected)]
  detected <- detected[!grepl("^[.]{3}[0-9]+$", detected)]
  detected <- detected[!tolower(detected) %in%
                         c("day", "days", "weekday", "day of week", "dow", "notes")]
  if (length(detected) > 0) {
    STAFF <<- detected
    message("Staff detected from sheet (", length(STAFF), "): ",
            paste(STAFF, collapse = ", "))
  }
  result <- setNames(lapply(STAFF, function(p) empty_df()), STAFF)

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
  unrecognised <- character(0)   # non-blank cells matching no known keyword
  blank_count  <- setNames(integer(length(STAFF)), STAFF)  # empty cells per person
  for (person in STAFF) {
    col_idx <- which(hdr_lc == tolower(person))
    # Fallback: match any header that starts with the person's first name
    if (length(col_idx) == 0)
      col_idx <- which(startsWith(hdr_lc, tolower(person)))
    if (length(col_idx) == 0) {
      message("NOTE: no column for '", person, "' in time-off source.")
      next
    }
    vals  <- as.character(raw[[col_idx[1L]]])
    types <- vapply(vals, classify_cell, character(1L), USE.NAMES = FALSE)
    # Track values that fell through to "off" without matching any keyword.
    # Classification is unchanged; this only feeds the typo warning below.
    odd <- vals[vapply(vals, is_unrecognised_cell, logical(1L), USE.NAMES = FALSE)]
    if (length(odd)) unrecognised <- c(unrecognised, trimws(odd))
    blank_count[[person]] <- sum(is.na(vals) | !nzchar(trimws(vals)))
    keep  <- !is.na(types)
    if (any(keep))
      result[[person]] <- data.frame(date = dates[keep], type = types[keep],
                                     stringsAsFactors = FALSE)
  }

  for (p in STAFF)
    result[[p]]$date <- as.Date(result[[p]]$date, origin = "1970-01-01")

  # ── Diagnostic summary ────────────────────────────────────────────────────
  message("Time-off parse summary (",
          format(SCHEDULE_START, "%b %d"), " \u2013 ",
          format(SCHEDULE_END,   "%b %d"), "):")
  for (p in STAFF) {
    df      <- result[[p]]
    n_off   <- sum(df$type == "off",   na.rm = TRUE)
    n_vac   <- sum(df$type == "vac",   na.rm = TRUE)
    n_cme   <- sum(df$type == "cme",   na.rm = TRUE)
    n_green <- sum(df$type == "green", na.rm = TRUE)
    n_yel   <- sum(df$type == "yellow", na.rm = TRUE)
    n_pto   <- sum(df$type == "pto",    na.rm = TRUE)
    nb <- if (p %in% names(blank_count)) blank_count[[p]] else 0L
    message(sprintf("  %-10s  off=%d  vac=%d  cme=%d  pto=%d  green=%d  yellow=%d%s",
                    p, n_off, n_vac, n_cme, n_pto, n_green, n_yel,
                    if (nb > 0L)
                      sprintf("   (%d blank -> %s)", nb,
                              if (identical(BLANK_CELL_MEANS, "green")) "green" else "neutral")
                    else ""))
  }
  # A mostly-empty column is worth calling out: with BLANK_CELL_MEANS = "green"
  # it reads as full availability, which may simply mean the person has not
  # filled the sheet in yet.
  heavy <- names(blank_count)[blank_count >= 0.5 * sum(in_window)]
  if (identical(BLANK_CELL_MEANS, "green") && length(heavy))
    message(sprintf("NOTE: %s left over half the sheet blank; those days count as available.",
                    paste(heavy, collapse = ", ")))

  # ── Typo warning ──────────────────────────────────────────────────────────
  # Any non-blank cell matching no known keyword was classified "off" (the safe
  # default). That is silent by design, but a mistyped WORK keyword now costs
  # double: it removes the day from the green-only phase AND reduces the
  # person's pay-period target via pto_reduction(). So make it visible.
  if (length(unrecognised)) {
    tb  <- sort(table(tolower(unrecognised)), decreasing = TRUE)
    top <- utils::head(tb, 8L)
    message(sprintf(
      "WARNING: %d cell value(s) matched no known keyword and were read as OFF: %s",
      length(unrecognised),
      paste(sprintf("'%s' x%d", names(top), as.integer(top)), collapse = ", ")))
    message("         Check for typos - a mistyped work keyword becomes a day OFF.")
  }

  # ── Green supply per calendar date ────────────────────────────────────────
  # Predicts where the green-only phase will leave holes, before spending an
  # hour in the solver. Each day needs 2 people for APP1+APP2, 3 to also cover
  # the night.
  green_by_date <- table(unlist(lapply(STAFF, function(p) {
    df <- result[[p]]
    as.character(df$date[df$type == "green"])
  })))
  all_ds  <- as.character(all_dates())
  n_green <- as.integer(green_by_date[all_ds]); n_green[is.na(n_green)] <- 0L
  if (sum(n_green) == 0L) {
    message("NOTE: no requested-work (green) days found. ",
            "The green-only phase will produce an empty schedule.")
  } else {
    short2 <- sum(n_green < 2L); short3 <- sum(n_green < 3L)
    message(sprintf(
      "Green supply: %d green day-requests over %d dates (mean %.1f people/day).",
      sum(n_green), length(all_ds), mean(n_green)))
    if (short2 > 0L)
      message(sprintf(
        "  %d date(s) have <2 green volunteers - APP1/APP2 cannot both be filled.",
        short2))
    if (short3 > 0L)
      message(sprintf(
        "  %d date(s) have <3 green volunteers - the night slot may go unfilled.",
        short3))
  }

  result
}

# ── Source dispatcher ────────────────────────────────────────────────────────

read_timeoff_source <- function(path, sheet) {
  # "Publish to web" CSV or /export?format=csv — no auth needed
  if (is_sheets_csv_url(path)) {
    message("Fetching CSV from Google Sheets public URL...")
    return(tryCatch(
      read.csv(path, header = TRUE, stringsAsFactors = FALSE,
               check.names = FALSE, na.strings = c("", "NA")),
      error = function(e) {
        message("ERROR fetching CSV: ", conditionMessage(e)); NULL
      }
    ))
  }
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

#' TRUE for Google Sheets "Publish to web" CSV or /export?format=csv URLs
#' (no auth needed — routed to read.csv instead of the API)
is_sheets_csv_url <- function(x) {
  grepl("docs\\.google\\.com/spreadsheets", x) &&
  (grepl("[?&]format=csv", x) || grepl("pub\\?.*output=csv", x))
}

is_sheets_url <- function(x) {
  (grepl("docs\\.google\\.com/spreadsheets", x) && !is_sheets_csv_url(x)) ||
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
    msg <- conditionMessage(e)
    message("ERROR reading Google Sheet: ", msg)
    if (grepl("403|PERMISSION_DENIED|permission", msg, ignore.case = TRUE)) {
      message(
        "  \u2192 The sheet is private and no credentials are configured.\n",
        "  \u2192 Quick fix: File \u2192 Share \u2192 Publish to web \u2192 CSV,\n",
        "    then set TIMEOFF_SOURCE to that URL (ends with pub?output=csv).\n",
        "  \u2192 Or set GS_SERVICE_ACCOUNT_JSON to your service-account key path."
      )
    }
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

#' "cme" | "vac" | "green" | "off" | NA
#'
#' Classification rules (evaluated in order; first match wins):
#'   1. Empty / whitespace / NA  -> NA  (neutral "yellow" day, no entry recorded)
#'   2. Contains a CME keyword as a whole word -> "cme"
#'      Keywords: cme, conf, conference
#'   3. Contains a VAC keyword as a whole word -> "vac"
#'      Keywords: VAC_KEYWORDS constant
#'   4. Contains a WORK keyword as a whole word -> "green" (requested WORK day)
#'      Keywords: WORK_KEYWORDS constant
#'   5. Anything else non-blank -> "off"
#'
#' Word-boundary (\b) matching is used throughout, and the ORDER matters:
#'   - CME before VAC  : "CME trip" resolves to cme, not vac.
#'   - VAC before WORK : "work trip" resolves to vac, not green. A work trip is
#'     time away from the unit, so blocking the day is the safe reading.
#'
#' Punctuation is stripped to spaces first so "W.", "Work?" and "y/n" tokenise
#' into whole words that \b can match.
#'
#' NOTE on the final fallthrough: an unrecognised non-blank value becomes "off",
#' NOT "green". This is deliberate. Treating an unrecognised value as available
#' risks scheduling someone on a day they said they could not work - an
#' operational incident; the reverse merely costs a shift. parse_time_off()
#' warns about every value that lands here so typos stay visible.
classify_cell <- function(val) {
  # Empty cell: governed by BLANK_CELL_MEANS ("green" = available, "neutral" =
  # avoid if possible, same as Yellow).
  blank_type <- if (identical(BLANK_CELL_MEANS, "green")) "green" else "yellow"
  if (is.na(val) || !nzchar(trimws(val))) return(blank_type)
  v <- tolower(trimws(val))
  v <- gsub("[[:punct:]]+", " ", v)   # "W." / "Work?" -> whole-word matchable
  v <- trimws(gsub("[[:space:]]+", " ", v))
  if (!nzchar(v)) return(blank_type)      # cell held only punctuation
  # CME check - word-boundary regex so "CME trip" still resolves to cme
  if (grepl("\\b(cme|conf|conference)\\b", v)) return("cme")
  # PTO check - a requested paid-time-off day. Blocked like Red, credited
  # toward the pay-period target like CME (see compute_targets()).
  pto_pat <- paste0("\\b(", paste(unique(PTO_KEYWORDS), collapse = "|"), ")\\b")
  if (grepl(pto_pat, v))                       return("pto")
  # NEUTRAL check - "Yellow" and friends mean schedulable-but-not-requested,
  # which is exactly what a blank cell means. Must come before the OFF
  # fallthrough or every yellow day would be hard-blocked.
  neu_pat <- paste0("\\b(", paste(unique(NEUTRAL_KEYWORDS), collapse = "|"), ")\\b")
  if (grepl(neu_pat, v))                       return("yellow")
  # VAC check - word-boundary regex on VAC_KEYWORDS
  vac_pat <- paste0("\\b(", paste(unique(VAC_KEYWORDS), collapse = "|"), ")\\b")
  if (grepl(vac_pat, v))                       return("vac")
  # WORK check - requested-work ("green") day
  work_pat <- paste0("\\b(", paste(unique(WORK_KEYWORDS), collapse = "|"), ")\\b")
  if (grepl(work_pat, v))                      return("green")
  "off"
}

#' TRUE when `val` is a non-blank cell matching NO known keyword, so it fell
#' through to "off". Used only for the typo warning - classification is unchanged.
is_unrecognised_cell <- function(val) {
  if (is.na(val) || !nzchar(trimws(val))) return(FALSE)
  v <- tolower(trimws(val))
  v <- trimws(gsub("[[:space:]]+", " ", gsub("[[:punct:]]+", " ", v)))
  if (!nzchar(v)) return(FALSE)
  pat <- paste0("\\b(",
                paste(unique(c("cme", "conf", "conference", VAC_KEYWORDS,
                               WORK_KEYWORDS, OFF_KEYWORDS, NEUTRAL_KEYWORDS,
                               PTO_KEYWORDS)),
                      collapse = "|"), ")\\b")
  !grepl(pat, v)
}

looks_like_dates <- function(col) {
  non_na <- col[!is.na(col)]
  length(non_na) > 0 && (
    inherits(non_na, "Date") ||
    all(grepl("^\\d{1,4}[/-]\\d{1,2}[/-]\\d{2,4}$", as.character(non_na)))
  )
}

coerce_dates <- function(x) {
  if (inherits(x, "Date"))                  return(x)
  if (inherits(x, c("POSIXct","POSIXlt"))) return(as.Date(x))
  if (is.numeric(x))                        return(as.Date(x, origin = "1899-12-30"))
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
