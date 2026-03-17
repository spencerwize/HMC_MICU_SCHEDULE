# ─────────────────────────────────────────────────────────────────────────────
# excel_output.R  —  Build 4-sheet Excel workbook with openxlsx
#
# Sheet order:  1. Calendar   monthly Apr–Jul view + staff dropdown
#               2. Summary    PP-by-PP metrics table
#               3. Schedule   full day-by-day grid with colour coding
#               4. _Lookup    date × person role matrix (support data)
# ─────────────────────────────────────────────────────────────────────────────

build_excel <- function(sched_obj, time_off, targets, output_path) {

  wb    <- createWorkbook()
  all_d <- sched_obj$dates

  # ── Style factory ─────────────────────────────────────────────────────────
  mk <- function(fg = NULL, bold = FALSE, size = 10,
                 font_color = "black", halign = "center",
                 border = NULL, border_color = "black",
                 wrap = FALSE) {
    args <- list(fontSize = size, fontColour = font_color,
                 halign = halign, valign = "center",
                 wrapText = wrap, fontName = "Calibri")
    if (!is.null(fg))     args$fgFill         <- fg
    if (bold)             args$textDecoration  <- "bold"
    if (!is.null(border)) {
      args$border       <- border
      args$borderColour <- border_color
    }
    do.call(createStyle, args)
  }

  role_fill <- function(role, is_holiday = FALSE) {
    if (is_holiday && role %in% c("APP1","APP2","Roam","Night"))
      return("#FFFF99")
    switch(role,
      APP1  = "#92D050", APP2 = "#92D050", Roam = "#92D050",
      Night = "#BDD7EE",
      VAC   = "#FFD966",
      PTO   = "#FF99CC",
      CME   = "#FF6D01",
      OFF   = "#FFC7CE",
      NULL
    )
  }

  # ── Person-role lookup (computed once) ────────────────────────────────────
  person_role <- function(person, d) {
    ds    <- as.character(d)
    day_s <- sched_obj$schedule[[ds]]
    for (s in SLOTS) {
      v <- day_s[[s]]
      if (!is.na(v) && v == person)
        return(if (s == "Night") "Night" else
               if (s == "APP1") "APP1"  else
               if (s == "APP2") "APP2"  else "Roam")
    }
    pdata <- time_off[[person]]
    m     <- pdata[pdata$date == d, ]
    typ   <- if (nrow(m) > 0) m$type[1] else NA_character_
    if (person %in% PP13_CONFERENCE &&
        d >= PP13_CONF_START && d <= PP13_CONF_END) typ <- "cme"
    if (!is.na(typ)) {
      pp_now  <- get_pp(d)
      pp_info <- if (!is.na(pp_now)) targets[[person]][[pp_now]] else NULL
      return(switch(typ,
        cme = "CME",
        off = "OFF",
        vac = if (!is.null(pp_info) && pp_info$avail < 6L) "PTO" else "VAC",
        ""
      ))
    }
    ""
  }

  role_lookup <- setNames(
    lapply(all_d, function(d)
      setNames(sapply(STAFF, function(p) person_role(p, d)), STAFF)),
    as.character(all_d)
  )

  # ══════════════════════════════════════════════════════════════════════════
  # SHEET 1: Calendar  (added first so it opens by default)
  # ══════════════════════════════════════════════════════════════════════════
  addWorksheet(wb, "Calendar")

  writeData(wb, "Calendar", x = "HMC MICU APP Schedule  2026",
    startRow = 1, startCol = 1, colNames = FALSE)
  addStyle(wb, "Calendar",
    mk(bold = TRUE, size = 14, halign = "left"), rows = 1, cols = 1)

  writeData(wb, "Calendar", x = "Select Staff:",
    startRow = 2, startCol = 1, colNames = FALSE)
  addStyle(wb, "Calendar",
    mk(bold = TRUE, halign = "left"), rows = 2, cols = 1)

  dataValidation(wb, "Calendar", col = 3, rows = 2, type = "list",
    value = paste0('"', paste(STAFF, collapse = ","), '"'))
  writeData(wb, "Calendar", x = STAFF[1],
    startRow = 2, startCol = 3, colNames = FALSE)

  # Note: static xlsx can't auto-respond to the C2 dropdown without VBA.
  # The calendar cells show the day number; dynamic per-person filtering
  # is available in the Shiny app.  The _Lookup sheet has full role data.
  writeData(wb, "Calendar",
    x = "(Role lookup available in _Lookup sheet; dynamic view in Shiny app)",
    startRow = 2, startCol = 5, colNames = FALSE)
  addStyle(wb, "Calendar",
    mk(font_color = "#888888", size = 9, halign = "left"),
    rows = 2, cols = 5)

  months_list <- list(
    list(year = 2026L, month = 4L, label = "April 2026"),
    list(year = 2026L, month = 5L, label = "May 2026"),
    list(year = 2026L, month = 6L, label = "June 2026"),
    list(year = 2026L, month = 7L, label = "July 2026")
  )
  dow_lbl <- c("Sun","Mon","Tue","Wed","Thu","Fri","Sat")

  cal_row <- 4L
  for (mo in months_list) {
    first_d <- as.Date(sprintf("%d-%02d-01", mo$year, mo$month))
    last_d  <- if (mo$month == 12L)
      as.Date(sprintf("%d-01-01", mo$year + 1L)) - 1L else
      as.Date(sprintf("%d-%02d-01", mo$year, mo$month + 1L)) - 1L

    # Month label spanning cols 1–7
    mergeCells(wb, "Calendar", cols = 1:7, rows = cal_row)
    writeData(wb, "Calendar", x = mo$label,
      startRow = cal_row, startCol = 1, colNames = FALSE)
    addStyle(wb, "Calendar",
      mk(fg = "#203864", bold = TRUE, font_color = "white", size = 12),
      rows = cal_row, cols = 1:7)
    setRowHeights(wb, "Calendar", rows = cal_row, heights = 22)
    cal_row <- cal_row + 1L

    # Weekday header
    for (ci in seq_along(dow_lbl)) {
      writeData(wb, "Calendar", x = dow_lbl[ci],
        startRow = cal_row, startCol = ci, colNames = FALSE)
      addStyle(wb, "Calendar",
        mk(fg = "#2E75B6", bold = TRUE, font_color = "white",
           border = "All", border_color = "white"),
        rows = cal_row, cols = ci)
    }
    cal_row <- cal_row + 1L

    start_dow <- as.integer(format(first_d, "%w"))  # 0=Sun
    week_row  <- cal_row
    col       <- start_dow + 1L
    cur       <- first_d

    while (cur <= last_d) {
      if (col > 7L) { col <- 1L; week_row <- week_row + 1L }

      in_sched <- (cur >= SCHEDULE_START && cur <= SCHEDULE_END)
      day_num  <- as.integer(format(cur, "%d"))

      writeData(wb, "Calendar", x = day_num,
        startRow = week_row, startCol = col, colNames = FALSE)

      bg_clr <- if (!in_sched) "#DDDDDD" else
                if (is_weekend(cur)) "#F2F2F2" else "white"
      is_hol <- cur %in% HOLIDAY_DATES
      if (is_hol) bg_clr <- "#FFFF99"

      addStyle(wb, "Calendar",
        mk(fg = bg_clr, halign = "center",
           border = "All", border_color = "#CCCCCC", size = 10),
        rows = week_row, cols = col)

      cur <- cur + 1L
      col <- col + 1L
    }
    cal_row <- week_row + 2L
  }

  setColWidths(wb, "Calendar", cols = 1:7, widths = rep(12, 7))

  # ══════════════════════════════════════════════════════════════════════════
  # SHEET 2: Summary
  # ══════════════════════════════════════════════════════════════════════════
  addWorksheet(wb, "Summary")

  mergeCells(wb, "Summary", cols = 1:12, rows = 1)
  writeData(wb, "Summary",
    x = "HMC MICU APP Schedule — Apr 13 – Jul 19, 2026",
    startRow = 1, startCol = 1, colNames = FALSE)
  addStyle(wb, "Summary",
    mk(fg = "#203864", bold = TRUE, font_color = "white", size = 14),
    rows = 1, cols = 1:12)
  setRowHeights(wb, "Summary", rows = 1, heights = 28)

  pp_hdrs <- c("Person", PAY_PERIODS$name,
               "Total Day", "Total Night", "Total Roam")
  writeData(wb, "Summary", x = as.data.frame(t(pp_hdrs)),
    startRow = 3, startCol = 1, colNames = FALSE)
  addStyle(wb, "Summary",
    mk(fg = "#2E75B6", bold = TRUE, font_color = "white",
       border = "All", border_color = "white"),
    rows = 3, cols = seq_along(pp_hdrs))

  for (ri in seq_along(STAFF)) {
    person <- STAFF[ri]
    row    <- ri + 3L

    writeData(wb, "Summary", x = person,
      startRow = row, startCol = 1, colNames = FALSE)
    addStyle(wb, "Summary",
      mk(bold = TRUE, halign = "left", border = "All", border_color = "#CCCCCC"),
      rows = row, cols = 1)

    for (ci in seq_along(PAY_PERIODS$name)) {
      pp_name <- PAY_PERIODS$name[ci]
      actual  <- sched_obj$pp_counts[[person]][[pp_name]]
      tgt     <- targets[[person]][[pp_name]]$sched_target
      cme     <- targets[[person]][[pp_name]]$credited
      txt     <- sprintf("%d / %d  (CME:%d)", actual, tgt, cme)
      col     <- ci + 1L

      writeData(wb, "Summary", x = txt,
        startRow = row, startCol = col, colNames = FALSE)
      bg <- if (actual < tgt) "#FFCCCC" else
            if (actual > tgt) "#CCFFCC" else "#FFFFFF"
      addStyle(wb, "Summary",
        mk(fg = bg, border = "All", border_color = "#CCCCCC"),
        rows = row, cols = col)
    }

    total_day   <- nrow(sched_obj$person_shifts[[person]])
    total_night <- length(sched_obj$person_nights[[person]])
    total_roam  <- sum(sched_obj$person_shifts[[person]]$slot == "Roaming")
    nc <- length(PAY_PERIODS$name) + 2L

    for (j in 0:2) {
      val <- c(total_day, total_night, total_roam)[j + 1L]
      writeData(wb, "Summary", x = val,
        startRow = row, startCol = nc + j, colNames = FALSE)
      addStyle(wb, "Summary",
        mk(border = "All", border_color = "#CCCCCC"),
        rows = row, cols = nc + j)
    }
  }

  note_row <- length(STAFF) + 5L
  mergeCells(wb, "Summary", cols = 1:12, rows = note_row)
  writeData(wb, "Summary",
    x = paste("NOTE PP13 (Jun 22–Jul 5):",
              "John/Todd/Mandie/Maureen at conferences.",
              "Only 6 active staff × 6 shifts = 36 available vs 56 slots —",
              "20 slots structurally empty."),
    startRow = note_row, startCol = 1, colNames = FALSE)
  addStyle(wb, "Summary",
    mk(fg = "#FFF2CC", bold = TRUE, halign = "left", size = 9, wrap = TRUE),
    rows = note_row, cols = 1)
  setRowHeights(wb, "Summary", rows = note_row, heights = 32)

  setColWidths(wb, "Summary",
    cols   = seq_along(pp_hdrs),
    widths = c(12, rep(16, length(PAY_PERIODS$name)), 11, 12, 11))

  # ══════════════════════════════════════════════════════════════════════════
  # SHEET 3: Schedule  (full day-by-day grid)
  # ══════════════════════════════════════════════════════════════════════════
  addWorksheet(wb, "Schedule")

  hdr <- c("Date", "Day", "PP", "APP1", "APP2", "Roaming", "Night", STAFF)
  writeData(wb, "Schedule", x = as.data.frame(t(hdr)),
    startRow = 1, startCol = 1, colNames = FALSE)
  addStyle(wb, "Schedule",
    mk(fg = "#203864", bold = TRUE, font_color = "white",
       border = "Bottom", border_color = "white"),
    rows = 1, cols = seq_along(hdr))

  for (ri in seq_along(all_d)) {
    d    <- all_d[ri]
    ds   <- as.character(d)
    row  <- ri + 1L
    pp   <- get_pp(d)
    day  <- sched_obj$schedule[[ds]]
    is_h <- d %in% HOLIDAY_DATES
    is_w <- is_weekend(d)
    bg   <- if (is_w) "#F2F2F2" else "white"

    writeData(wb, "Schedule",
      x = data.frame(
        Date    = format(d, "%m/%d/%Y"),
        Day     = weekdays(d, abbreviate = TRUE),
        PP      = ifelse(is.na(pp), "", pp),
        APP1    = ifelse(is.na(day$APP1),    "", day$APP1),
        APP2    = ifelse(is.na(day$APP2),    "", day$APP2),
        Roaming = ifelse(is.na(day$Roaming), "", day$Roaming),
        Night   = ifelse(is.na(day$Night),   "", day$Night),
        stringsAsFactors = FALSE
      ),
      startRow = row, startCol = 1, colNames = FALSE)

    addStyle(wb, "Schedule",
      mk(fg = bg, halign = "left"), rows = row, cols = 1:3)
    addStyle(wb, "Schedule",
      mk(fg = if (is_h) "#FFFF99" else "#D6E4BC", halign = "center"),
      rows = row, cols = 4:6)
    addStyle(wb, "Schedule",
      mk(fg = if (is_h) "#FFFF99" else "#BDD7EE", halign = "center"),
      rows = row, cols = 7)

    for (ci in seq_along(STAFF)) {
      person <- STAFF[ci]
      col    <- 7L + ci
      role   <- role_lookup[[ds]][person]
      if (is.na(role)) role <- ""

      writeData(wb, "Schedule", x = role,
        startRow = row, startCol = col, colNames = FALSE)

      bg_r <- role_fill(role, is_h)
      sty  <- if (!is.null(bg_r))
                mk(fg = bg_r, size = 9,
                   border = "All", border_color = "#BBBBBB")
              else
                mk(fg = bg, size = 9)
      addStyle(wb, "Schedule", sty, rows = row, cols = col)
    }
  }

  setColWidths(wb, "Schedule",
    cols   = seq_along(hdr),
    widths = c(12, 5, 6, 9, 9, 9, 9, rep(9, length(STAFF))))
  freezePane(wb, "Schedule", firstRow = TRUE)

  # ══════════════════════════════════════════════════════════════════════════
  # SHEET 4: _Lookup  (date × person role matrix)
  # ══════════════════════════════════════════════════════════════════════════
  addWorksheet(wb, "_Lookup")

  lk_hdr <- c("Date", STAFF)
  writeData(wb, "_Lookup", x = as.data.frame(t(lk_hdr)),
    startRow = 1, startCol = 1, colNames = FALSE)
  addStyle(wb, "_Lookup",
    mk(fg = "#203864", bold = TRUE, font_color = "white"),
    rows = 1, cols = seq_along(lk_hdr))

  for (ri in seq_along(all_d)) {
    d   <- all_d[ri]
    row <- ri + 1L
    roles_row <- c(format(d, "%m/%d/%Y"),
                   sapply(STAFF, function(p) {
                     r <- role_lookup[[as.character(d)]][p]
                     if (is.na(r)) "" else r
                   }))
    writeData(wb, "_Lookup", x = as.data.frame(t(roles_row)),
      startRow = row, startCol = 1, colNames = FALSE)
  }
  setColWidths(wb, "_Lookup",
    cols   = seq_along(lk_hdr),
    widths = c(12, rep(9, length(STAFF))))

  # ── Save ──────────────────────────────────────────────────────────────────
  saveWorkbook(wb, output_path, overwrite = TRUE)
  message("  Saved: ", output_path)
  invisible(output_path)
}
