# ─────────────────────────────────────────────────────────────────────────────
# parse_time_off.R  —  Read Time_Off_Requests.xlsx using tidyxl
#
# Returns: named list  person -> data.frame(date=Date, type=chr)
#   type values: "off"  plain off day   (red fill, no vac keyword)
#                "vac"  vacation day    (red fill + vac keyword)
#                "cme"  credited day    (orange fill + yellow border)
# ─────────────────────────────────────────────────────────────────────────────

parse_time_off <- function(xlsx_path) {
  # Initialise empty result
  empty_df <- function() {
    data.frame(date = as.Date(character()), type = character(),
               stringsAsFactors = FALSE)
  }
  result <- setNames(lapply(STAFF, function(p) empty_df()), STAFF)

  if (!file.exists(xlsx_path)) {
    message("WARNING: ", xlsx_path, " not found — proceeding with no time-off data.")
    return(result)
  }

  # ── Load cells & formats ───────────────────────────────────────────────────
  sheet_name <- "2026 Full Year"
  avail_sheets <- tidyxl::xlsx_sheet_names(xlsx_path)
  if (!sheet_name %in% avail_sheets) {
    sheet_name <- avail_sheets[1]
    message("WARNING: sheet '2026 Full Year' not found; using '", sheet_name, "'")
  }

  cells   <- tidyxl::xlsx_cells(xlsx_path, sheets = sheet_name)
  formats <- tidyxl::xlsx_formats(xlsx_path)

  # Pull fill fg color and border colors from formats by local_format_id
  fg_colors    <- formats$local$fill$patternFill$fgColor$rgb
  border_left  <- formats$local$border$left$color$rgb
  border_right <- formats$local$border$right$color$rgb
  border_top   <- formats$local$border$top$color$rgb
  border_bot   <- formats$local$border$bottom$color$rgb

  safe_color <- function(vec, idx) {
    if (is.null(vec) || idx < 1 || idx > length(vec)) return(NA_character_)
    toupper(as.character(vec[idx]))
  }

  is_red_fill <- function(fmt_id) {
    clr <- safe_color(fg_colors, fmt_id)
    !is.na(clr) && clr %in% c("FFF4CCCC", "F4CCCC", "FFFF0000",
                               "FFFFE0E0", "FFFFC0C0", "FFFF9999")
  }

  is_orange_fill <- function(fmt_id) {
    clr <- safe_color(fg_colors, fmt_id)
    !is.na(clr) && clr == "FFFF6D01"
  }

  has_yellow_border <- function(fmt_id) {
    YELLOW <- c("FFFFFF00", "FFFF00", "00FFFF00")
    any(sapply(list(border_left, border_right, border_top, border_bot),
               function(vec) safe_color(vec, fmt_id) %in% YELLOW))
  }

  # ── Find date rows (scan col 1 for Date values) ────────────────────────────
  col1 <- cells[cells$col == 1, ]
  date_rows <- col1[!is.na(col1$date), c("row", "date")]

  # Keep only rows within schedule window
  date_rows <- date_rows[
    !is.na(date_rows$date) &
      date_rows$date >= SCHEDULE_START &
      date_rows$date <= SCHEDULE_END, ]

  if (nrow(date_rows) == 0) {
    message("WARNING: No date rows found in time-off sheet for schedule window.")
    return(result)
  }

  # Build a row-indexed lookup of cells for fast access
  cells_by_row <- split(cells, cells$row)

  # ── Parse each date row ────────────────────────────────────────────────────
  for (k in seq_len(nrow(date_rows))) {
    row_date <- date_rows$date[k]
    row_num  <- date_rows$row[k]

    row_cells <- cells_by_row[[as.character(row_num)]]
    if (is.null(row_cells)) next

    for (person in names(STAFF_COL_MAP)) {
      col_idx <- STAFF_COL_MAP[person]
      cell    <- row_cells[row_cells$col == col_idx, ]
      if (nrow(cell) == 0) next

      fmt_id <- cell$local_format_id[1]
      val    <- tolower(as.character(cell$character[1]))
      if (is.na(val)) val <- ""

      # CME: orange fill + yellow border on any side
      if (is_orange_fill(fmt_id) && has_yellow_border(fmt_id)) {
        result[[person]] <- rbind(result[[person]],
          data.frame(date = row_date, type = "cme", stringsAsFactors = FALSE))
        next
      }

      # Off / Vac: red-ish fill
      if (is_red_fill(fmt_id)) {
        is_vac <- any(sapply(VAC_KEYWORDS, function(kw) grepl(kw, val, fixed = TRUE)))
        type   <- if (is_vac) "vac" else "off"
        result[[person]] <- rbind(result[[person]],
          data.frame(date = row_date, type = type, stringsAsFactors = FALSE))
      }
    }
  }

  # Ensure date column is Date class (rbind can coerce)
  for (p in STAFF) {
    result[[p]]$date <- as.Date(result[[p]]$date, origin = "1970-01-01")
  }

  result
}
