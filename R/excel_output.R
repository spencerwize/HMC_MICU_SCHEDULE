# ─────────────────────────────────────────────────────────────────────────────
# excel_output.R  —  Build 3-sheet Excel workbook matching reference format
#
# Sheet 1: Calendar   monthly Apr–Jul view (two rows/week: date + assignment)
# Sheet 2: Summary    overview stats + PP-by-PP detail + staffing rules
# Sheet 3: Schedule   two rows/day (day shift + night shift) + PP headers
# ─────────────────────────────────────────────────────────────────────────────

build_excel <- function(sched_obj, time_off, targets, output_path) {

  wb    <- createWorkbook()
  all_d <- sched_obj$dates

  # ── Color palette ──────────────────────────────────────────────────────────
  C_NAVY     <- "#1F3864"
  C_BLUE     <- "#2E75B6"
  C_BLUE_LT  <- "#D6E4F3"
  C_NIGHT    <- "#D6E4F0"
  C_LAVENDER <- "#EEF1FF"
  C_GREEN    <- "#E2EFDA"
  C_YELLOW   <- "#FFF2CC"
  C_PEACH    <- "#FCE4D6"
  C_PINK     <- "#F4CCCC"
  C_ORANGE   <- "#FF6D01"
  C_CREAM    <- "#FFFBF0"
  C_GRAY_LT  <- "#F2F2F2"

  F_WHITE    <- "#FFFFFF"
  F_NAVY     <- "#1F3864"
  F_BLUE     <- "#1A5276"
  F_GOLD     <- "#7F6000"
  F_RED      <- "#8B0000"
  F_BROWN    <- "#833C00"
  F_GRAY     <- "#404040"
  F_LGRAY    <- "#CCCCCC"

  # ── Style factory ──────────────────────────────────────────────────────────
  mk <- function(fg = NULL, bold = FALSE, size = 10,
                 font_color = "#000000", halign = "center", valign = "center",
                 border = NULL, border_color = "#CCCCCC",
                 border_style = "thin", wrap = FALSE) {
    args <- list(fontSize = size, fontColour = font_color,
                 halign = halign, valign = valign,
                 wrapText = wrap, fontName = "Calibri")
    if (!is.null(fg))   args$fgFill        <- fg
    if (bold)           args$textDecoration <- "bold"
    if (!is.null(border)) {
      args$border       <- if (border == "All") c("top","bottom","left","right")
                           else tolower(border)
      args$borderColour <- border_color
      args$borderStyle  <- border_style
    }
    do.call(createStyle, args)
  }

  # ── Role helpers ───────────────────────────────────────────────────────────
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
        ""))
    }
    ""
  }

  role_bg <- function(role, is_holiday = FALSE) {
    if (is_holiday && role %in% c("APP1","APP2","Roam","Night"))
      return(C_YELLOW)
    switch(role,
      APP1 = C_GREEN, APP2 = C_GREEN, Roam = C_GREEN,
      Night = C_NIGHT,
      VAC  = C_PEACH,  PTO = "#FF9999",
      CME  = C_ORANGE, OFF = C_PINK,
      NULL)
  }

  role_fc <- function(role) {
    switch(role,
      APP1 = F_BLUE, APP2 = F_BLUE, Roam = F_BLUE,
      Night = F_NAVY,
      VAC = F_BROWN, PTO = F_RED,
      CME = F_WHITE, OFF = F_RED,
      "#000000")
  }

  # Precompute role lookup
  role_lookup <- setNames(
    lapply(all_d, function(d)
      setNames(sapply(STAFF, function(p) person_role(p, d)), STAFF)),
    as.character(all_d))

  # ── Day-before-night violation set ─────────────────────────────────────────
  dbn_set <- list()
  for (i in seq_along(all_d[-length(all_d)])) {
    d       <- as.Date(all_d[i], origin = "1970-01-01")
    night_p <- sched_obj$schedule[[as.character(d + 1L)]]$Night
    if (is.na(night_p)) next
    for (s in DAY_SLOTS) {
      v <- sched_obj$schedule[[as.character(d)]][[s]]
      if (!is.na(v) && v == night_p)
        dbn_set[[length(dbn_set) + 1L]] <- list(date = d, person = night_p)
    }
  }
  is_dbn <- function(d, person)
    any(vapply(dbn_set, function(x) x$date == d && x$person == person, logical(1L)))

  N_STAFF <- length(STAFF)
  N_PP    <- nrow(PAY_PERIODS)

  # ════════════════════════════════════════════════════════════════════════════
  # SHEET 1 · Calendar
  # ════════════════════════════════════════════════════════════════════════════
  addWorksheet(wb, "Calendar")
  setColWidths(wb, "Calendar", cols = 1,   widths = 2.0)
  setColWidths(wb, "Calendar", cols = 2:8, widths = rep(13.0, 7))
  setColWidths(wb, "Calendar", cols = 9,   widths = 2.0)

  # Title
  mergeCells(wb, "Calendar", cols = 2:8, rows = 1)
  writeData(wb, "Calendar",
    x = "APP Staff Schedule \u00B7 Apr 13 \u2013 Jul 19, 2026",
    startRow = 1, startCol = 2, colNames = FALSE)
  addStyle(wb, "Calendar",
    mk(fg = C_NAVY, bold = TRUE, size = 14, font_color = F_WHITE),
    rows = 1, cols = 2:8)
  setRowHeights(wb, "Calendar", rows = 1, heights = 28)

  # Staff selector row
  writeData(wb, "Calendar", x = "Staff member:",
    startRow = 2, startCol = 2, colNames = FALSE)
  addStyle(wb, "Calendar",
    mk(fg = C_BLUE_LT, bold = TRUE, font_color = F_NAVY, halign = "right"),
    rows = 2, cols = 2)
  mergeCells(wb, "Calendar", cols = 3:5, rows = 2)
  writeData(wb, "Calendar", x = STAFF[1],
    startRow = 2, startCol = 3, colNames = FALSE)
  addStyle(wb, "Calendar",
    mk(fg = C_BLUE_LT, font_color = F_NAVY, halign = "left"),
    rows = 2, cols = 3:5)
  dataValidation(wb, "Calendar", col = 3, rows = 2, type = "list",
    value = paste0('"', paste(STAFF, collapse = ","), '"'))
  mergeCells(wb, "Calendar", cols = 6:8, rows = 2)
  writeData(wb, "Calendar",
    x = "(Dynamic view available in the Shiny app)",
    startRow = 2, startCol = 6, colNames = FALSE)
  addStyle(wb, "Calendar",
    mk(fg = C_BLUE_LT, font_color = "#888888", size = 8, halign = "left"),
    rows = 2, cols = 6:8)
  setRowHeights(wb, "Calendar", rows = 3, heights = 6)

  months_list <- list(
    list(year = 2026L, month = 4L, label = "April 2026"),
    list(year = 2026L, month = 5L, label = "May 2026"),
    list(year = 2026L, month = 6L, label = "June 2026"),
    list(year = 2026L, month = 7L, label = "July 2026"))
  dow_lbl <- c("Sun","Mon","Tue","Wed","Thu","Fri","Sat")

  cal_row <- 4L

  for (mo in months_list) {
    first_d <- as.Date(sprintf("%d-%02d-01", mo$year, mo$month))
    last_d  <- as.Date(sprintf("%d-%02d-01",
      mo$year + (mo$month == 12L), (mo$month %% 12L) + 1L)) - 1L

    # Month header
    mergeCells(wb, "Calendar", cols = 2:8, rows = cal_row)
    writeData(wb, "Calendar", x = mo$label,
      startRow = cal_row, startCol = 2, colNames = FALSE)
    addStyle(wb, "Calendar",
      mk(fg = C_BLUE_LT, bold = TRUE, font_color = F_NAVY, size = 12),
      rows = cal_row, cols = 2:8)
    setRowHeights(wb, "Calendar", rows = cal_row, heights = 20)
    cal_row <- cal_row + 1L

    # Day-of-week header
    for (ci in seq_along(dow_lbl)) {
      writeData(wb, "Calendar", x = dow_lbl[ci],
        startRow = cal_row, startCol = ci + 1L, colNames = FALSE)
      addStyle(wb, "Calendar",
        mk(fg = C_NAVY, bold = TRUE, font_color = F_WHITE,
           border = "All", border_color = F_WHITE),
        rows = cal_row, cols = ci + 1L)
    }
    setRowHeights(wb, "Calendar", rows = cal_row, heights = 18)
    cal_row <- cal_row + 1L

    start_dow     <- as.integer(format(first_d, "%w"))
    wk_date_row   <- cal_row
    wk_role_row   <- cal_row + 1L
    col           <- start_dow + 2L
    cur           <- first_d

    while (cur <= last_d) {
      if (col > 8L) {
        col         <- 2L
        wk_date_row <- wk_date_row + 2L
        wk_role_row <- wk_role_row + 2L
      }

      in_sched <- cur >= SCHEDULE_START && cur <= SCHEDULE_END
      is_hol   <- cur %in% HOLIDAY_DATES
      is_wknd  <- as.integer(format(cur, "%w")) %in% c(0L, 6L)

      bg_d <- if (!in_sched) "#F7F7F7"
              else if (is_hol) C_YELLOW
              else if (is_wknd) C_LAVENDER
              else "#FFFFFF"
      fc_d <- if (!in_sched) F_LGRAY else if (is_hol) F_GOLD else "#555555"

      # Date number cell
      writeData(wb, "Calendar", x = as.integer(format(cur, "%d")),
        startRow = wk_date_row, startCol = col, colNames = FALSE)
      addStyle(wb, "Calendar",
        mk(fg = bg_d, font_color = fc_d, size = 9,
           border = "All", border_color = "#DDDDDD"),
        rows = wk_date_row, cols = col)

      # Assignment cell (shows first staff member)
      role_val <- if (in_sched) {
        rv <- role_lookup[[as.character(cur)]][[STAFF[1]]]
        if (is.na(rv)) "" else rv
      } else ""
      cell_bg <- if (in_sched) { rb <- role_bg(role_val, is_hol); if (is.null(rb)) bg_d else rb }
                 else bg_d
      cell_fc <- if (nchar(role_val) > 0) role_fc(role_val) else fc_d

      writeData(wb, "Calendar", x = role_val,
        startRow = wk_role_row, startCol = col, colNames = FALSE)
      addStyle(wb, "Calendar",
        mk(fg = cell_bg, bold = nchar(role_val) > 0, font_color = cell_fc,
           size = 9, border = "All", border_color = "#DDDDDD"),
        rows = wk_role_row, cols = col)

      cur <- cur + 1L
      col <- col + 1L
    }

    setRowHeights(wb, "Calendar",
      rows = seq(cal_row, wk_date_row, by = 2L),
      heights = rep(15, length(seq(cal_row, wk_date_row, by = 2L))))
    setRowHeights(wb, "Calendar",
      rows = seq(cal_row + 1L, wk_role_row, by = 2L),
      heights = rep(18, length(seq(cal_row + 1L, wk_role_row, by = 2L))))

    cal_row <- wk_date_row + 3L  # pair + spacer
  }

  # Legend
  mergeCells(wb, "Calendar", cols = 2:8, rows = cal_row)
  writeData(wb, "Calendar",
    x = paste0("\u25A0 Working (APP1/APP2/Roam) = green  ",
               "\u25A0 Night = blue  \u25A0 OFF = pink  ",
               "\u25A0 VAC = peach  \u25A0 PTO = pink-red  ",
               "\u25A0 CME = orange  \u25A0 Holiday = yellow"),
    startRow = cal_row, startCol = 2, colNames = FALSE)
  addStyle(wb, "Calendar",
    mk(fg = C_CREAM, font_color = "#888888", size = 8, halign = "left"),
    rows = cal_row, cols = 2:8)

  # ════════════════════════════════════════════════════════════════════════════
  # SHEET 2 · Summary
  # ════════════════════════════════════════════════════════════════════════════
  addWorksheet(wb, "Summary")

  SUM_COLS <- 1L + N_STAFF

  sec_hdr <- function(row, label) {
    mergeCells(wb, "Summary", cols = 1:SUM_COLS, rows = row)
    writeData(wb, "Summary", x = label,
      startRow = row, startCol = 1, colNames = FALSE)
    addStyle(wb, "Summary",
      mk(fg = C_BLUE, bold = TRUE, font_color = F_WHITE,
         halign = "left", size = 11),
      rows = row, cols = 1:SUM_COLS)
    setRowHeights(wb, "Summary", rows = row, heights = 20)
  }

  staff_hdr <- function(row) {
    addStyle(wb, "Summary",
      mk(fg = C_BLUE, bold = TRUE, font_color = F_WHITE,
         border = "All", border_color = F_WHITE),
      rows = row, cols = 1)
    for (ci in seq_along(STAFF)) {
      writeData(wb, "Summary", x = STAFF[ci],
        startRow = row, startCol = 1L + ci, colNames = FALSE)
      addStyle(wb, "Summary",
        mk(fg = C_BLUE, bold = TRUE, font_color = F_WHITE,
           border = "All", border_color = F_WHITE),
        rows = row, cols = 1L + ci)
    }
    setRowHeights(wb, "Summary", rows = row, heights = 18)
  }

  # Precompute stats
  pstats <- lapply(STAFF, function(person) {
    shifts  <- sched_obj$person_shifts[[person]]
    nights  <- sched_obj$person_nights[[person]]
    n_cred  <- sum(sapply(PAY_PERIODS$name, function(pp)
      targets[[person]][[pp]]$credited))
    pdata   <- time_off[[person]]
    n_vac   <- sum(pdata$type == "vac", na.rm = TRUE)
    n_pto   <- 0L
    n_bump  <- 0L
    for (ppn in PAY_PERIODS$name) {
      ppi    <- targets[[person]][[ppn]]
      pp_row <- PAY_PERIODS[PAY_PERIODS$name == ppn, ]
      vac_in_pp <- pdata$type == "vac" &
        pdata$date >= pp_row$start & pdata$date <= pp_row$end
      if (ppi$avail < 6L) n_pto <- n_pto + sum(vac_in_pp)
      actual <- sched_obj$pp_counts[[person]][[ppn]]
      if (actual < ppi$sched_target)
        n_bump <- n_bump + (ppi$sched_target - actual)
    }
    list(
      n_sched  = nrow(shifts) + length(nights),
      n_day    = nrow(shifts),
      n_night  = length(nights),
      n_roam   = sum(shifts$slot == "Roaming"),
      n_wknd   = sum(is_weekend(shifts$date)) + sum(is_weekend(nights)),
      n_cred   = n_cred,
      n_total  = nrow(shifts) + length(nights) + n_cred,
      n_vac    = n_vac,
      n_pto    = n_pto,
      n_bump   = n_bump)
  })
  names(pstats) <- STAFF

  srow <- 1L

  # Title
  mergeCells(wb, "Summary", cols = 1:SUM_COLS, rows = srow)
  writeData(wb, "Summary",
    x = "SCHEDULE SUMMARY \u00B7 Apr 13 \u2013 Jul 19, 2026",
    startRow = srow, startCol = 1, colNames = FALSE)
  addStyle(wb, "Summary",
    mk(fg = C_NAVY, bold = TRUE, font_color = F_WHITE,
       size = 14, halign = "left"),
    rows = srow, cols = 1:SUM_COLS)
  setRowHeights(wb, "Summary", rows = srow, heights = 28)
  srow <- srow + 1L

  # ── Overview section ──────────────────────────────────────────────────────
  sec_hdr(srow, "Overview"); srow <- srow + 1L
  staff_hdr(srow);           srow <- srow + 1L

  ovr_rows <- list(
    list("Scheduled Shifts",         "n_sched",  C_GRAY_LT, FALSE),
    list("Credited Days (CME/Conf)", "n_cred",   "#FFFFFF",  TRUE),
    list("Total incl. Credited",     "n_total",  C_GRAY_LT, FALSE),
    list("Night Shifts",             "n_night",  "#FFFFFF",  FALSE),
    list("Roaming APP Shifts",       "n_roam",   C_GRAY_LT, FALSE),
    list("Weekend Shifts",           "n_wknd",   "#FFFFFF",  FALSE),
    list("Vacation Days",            "n_vac",    C_PEACH,    FALSE),
    list("PTO Days",                 "n_pto",    "#FFFFFF",  FALSE),
    list("Shortfall (shift-days)",   "n_bump",   "#FFFFFF",  FALSE))

  for (ov in ovr_rows) {
    label <- ov[[1]]; key <- ov[[2]]; row_bg <- ov[[3]]; cred_flag <- ov[[4]]
    writeData(wb, "Summary", x = label,
      startRow = srow, startCol = 1, colNames = FALSE)
    addStyle(wb, "Summary",
      mk(fg = row_bg, bold = TRUE, font_color = F_NAVY,
         halign = "left", border = "All", border_color = "#DDDDDD"),
      rows = srow, cols = 1)
    for (ci in seq_along(STAFF)) {
      val    <- pstats[[STAFF[ci]]][[key]]
      cbg    <- row_bg
      cfc    <- F_NAVY
      cbold  <- FALSE
      if (cred_flag  && val > 0) { cbg <- C_ORANGE; cfc <- F_WHITE; cbold <- TRUE }
      if (key == "n_pto"  && val > 0) { cbg <- "#FF9999"; cfc <- F_RED;  cbold <- TRUE }
      if (key == "n_bump" && val > 0) { cbg <- "#FFE0E0"; cfc <- "#C00000"; cbold <- TRUE }
      if (key == "n_vac"  && val > 0) { cbg <- C_PEACH;  cfc <- F_BROWN }
      writeData(wb, "Summary", x = val,
        startRow = srow, startCol = 1L + ci, colNames = FALSE)
      addStyle(wb, "Summary",
        mk(fg = cbg, bold = cbold, font_color = cfc,
           border = "All", border_color = "#DDDDDD"),
        rows = srow, cols = 1L + ci)
    }
    setRowHeights(wb, "Summary", rows = srow, heights = 16)
    srow <- srow + 1L
  }
  srow <- srow + 1L  # spacer

  # ── Pay Period Detail section ─────────────────────────────────────────────
  sec_hdr(srow, "Pay Period Detail"); srow <- srow + 1L
  staff_hdr(srow);                    srow <- srow + 1L

  for (i in seq_len(N_PP)) {
    ppn    <- PAY_PERIODS$name[i]
    pp_lbl <- sprintf("%s  (%s \u2013 %s)", ppn,
      format(PAY_PERIODS$start[i], "%b %d"),
      format(PAY_PERIODS$end[i],   "%b %d"))
    writeData(wb, "Summary", x = pp_lbl,
      startRow = srow, startCol = 1, colNames = FALSE)
    addStyle(wb, "Summary",
      mk(fg = C_BLUE_LT, bold = TRUE, font_color = F_NAVY,
         halign = "left", border = "All", border_color = "#DDDDDD"),
      rows = srow, cols = 1)
    for (ci in seq_along(STAFF)) {
      person  <- STAFF[ci]
      actual  <- sched_obj$pp_counts[[person]][[ppn]]
      ppi     <- targets[[person]][[ppn]]
      n_vac_p <- sum(time_off[[person]]$type == "vac" &
        time_off[[person]]$date >= PAY_PERIODS$start[i] &
        time_off[[person]]$date <= PAY_PERIODS$end[i])
      txt <- sprintf("%d/%d", actual, ppi$sched_target)
      if (ppi$credited > 0) txt <- paste0(txt, sprintf(" \u00B7%dc", ppi$credited))
      if (n_vac_p  > 0)     txt <- paste0(txt, sprintf(" \u00B7%dv", n_vac_p))
      cbg <- if (ppi$credited > 0)      C_ORANGE
             else if (actual < ppi$sched_target) "#FFE0E0"
             else if (n_vac_p  > 0)     C_PEACH
             else                       C_GREEN
      cfc  <- if (ppi$credited > 0) F_WHITE else F_NAVY
      cbold <- actual < ppi$sched_target
      writeData(wb, "Summary", x = txt,
        startRow = srow, startCol = 1L + ci, colNames = FALSE)
      addStyle(wb, "Summary",
        mk(fg = cbg, bold = cbold, font_color = cfc,
           border = "All", border_color = "#DDDDDD", size = 9),
        rows = srow, cols = 1L + ci)
    }
    setRowHeights(wb, "Summary", rows = srow, heights = 16)
    srow <- srow + 1L
  }
  srow <- srow + 1L  # spacer

  # ── Staffing & Rules section ──────────────────────────────────────────────
  sec_hdr(srow, "Staffing & Rules"); srow <- srow + 1L

  n_dbn       <- length(dbn_set)
  zero_app    <- sum(sapply(all_d, function(d)
    is.na(sched_obj$schedule[[as.character(d)]]$APP1)))
  n_unstaffed <- sum(sapply(all_d, function(d)
    is.na(sched_obj$schedule[[as.character(d)]]$Night)))

  rules <- list(
    list("Zero-APP Days",
         if (zero_app == 0) "\u2713 None" else paste(zero_app, "days missing APP1")),
    list("Unstaffed Nights",
         if (n_unstaffed == 0) "\u2713 None" else paste(n_unstaffed, "nights unstaffed")),
    list("Day\u2192Night Buffer Violations",
         sprintf("%d occurrence%s \u2014 flagged with dashed orange border in Schedule",
                 n_dbn, if (n_dbn == 1L) "" else "s")),
    list("Night Recovery",
         "After last night of streak: blocked D+1 and D+2; eligible again D+3"),
    list("Max Consecutive Nights",    "3"),
    list("Max Consecutive Work Days", "4"),
    list("PP13 Capacity Note",
         paste("John / Todd / Mandie / Maureen at conferences.",
               "Only 6 active staff \u00D7 6 slots = 36 available vs 56 total \u2014",
               "20 slots structurally empty.")),
    list("PTO Logic",
         "Vacation in a PP where available days < 6 counts as PTO (target-reducing)"))

  for (rl in rules) {
    mergeCells(wb, "Summary", cols = 2:SUM_COLS, rows = srow)
    writeData(wb, "Summary", x = rl[[1]],
      startRow = srow, startCol = 1, colNames = FALSE)
    writeData(wb, "Summary", x = rl[[2]],
      startRow = srow, startCol = 2, colNames = FALSE)
    addStyle(wb, "Summary",
      mk(fg = "#DBEDFF", bold = TRUE, font_color = F_NAVY,
         halign = "left", border = "All", border_color = "#DDDDDD", size = 9),
      rows = srow, cols = 1)
    addStyle(wb, "Summary",
      mk(fg = C_CREAM, font_color = "#333333",
         halign = "left", border = "All", border_color = "#DDDDDD",
         size = 9, wrap = TRUE),
      rows = srow, cols = 2:SUM_COLS)
    setRowHeights(wb, "Summary", rows = srow, heights = 18)
    srow <- srow + 1L
  }

  setColWidths(wb, "Summary",
    cols   = 1:SUM_COLS,
    widths = c(26, rep(12, N_STAFF)))
  freezePane(wb, "Summary", firstRow = TRUE)

  # ════════════════════════════════════════════════════════════════════════════
  # SHEET 3 · Schedule  (two rows per calendar day)
  # ════════════════════════════════════════════════════════════════════════════
  addWorksheet(wb, "Schedule")

  # Col layout: Date | Day | PP | Shift | APP1 | APP2 | Roaming | Night | [staff...] | _key_
  N_HDR  <- 8L
  N_COLS <- N_HDR + N_STAFF + 1L

  hdr <- c("Date","Day","PP","Shift",
           "APP 1","APP 2","Roaming APP","Night Shift",
           STAFF, "_key_")
  writeData(wb, "Schedule", x = as.data.frame(t(hdr)),
    startRow = 1, startCol = 1, colNames = FALSE)
  addStyle(wb, "Schedule",
    mk(fg = C_NAVY, bold = TRUE, font_color = F_WHITE,
       border = "Bottom", border_color = F_WHITE),
    rows = 1, cols = seq_along(hdr))
  setRowHeights(wb, "Schedule", rows = 1, heights = 20)
  freezePane(wb, "Schedule", firstRow = TRUE)

  schr  <- 2L
  prev_pp <- ""

  for (d_raw in all_d) {
    d   <- as.Date(d_raw, origin = "1970-01-01")
    ds  <- as.character(d)
    pp  <- get_pp(d)
    day <- sched_obj$schedule[[ds]]
    is_h <- d %in% HOLIDAY_DATES
    is_w <- is_weekend(d)

    # PP header row at each pay-period boundary
    if (!is.na(pp) && pp != prev_pp) {
      pp_i   <- which(PAY_PERIODS$name == pp)
      pp_hdr <- sprintf("%s   %s \u2013 %s", pp,
        format(PAY_PERIODS$start[pp_i], "%b %d"),
        format(PAY_PERIODS$end[pp_i],   "%b %d"))
      mergeCells(wb, "Schedule", cols = 1:4, rows = schr)
      writeData(wb, "Schedule", x = pp_hdr,
        startRow = schr, startCol = 1, colNames = FALSE)
      addStyle(wb, "Schedule",
        mk(fg = C_BLUE_LT, bold = TRUE, font_color = F_NAVY,
           border = "All", border_color = "#BBBBBB"),
        rows = schr, cols = 1:4)
      # Per-staff targets
      for (ci in seq_along(STAFF)) {
        person  <- STAFF[ci]
        ppi     <- targets[[person]][[pp]]
        n_vac_p <- sum(time_off[[person]]$type == "vac" &
          time_off[[person]]$date >= PAY_PERIODS$start[pp_i] &
          time_off[[person]]$date <= PAY_PERIODS$end[pp_i])
        lbl <- sprintf("T:%d", ppi$sched_target)
        if (ppi$credited > 0) lbl <- paste0(lbl, sprintf("/C:%d", ppi$credited))
        if (n_vac_p > 0)      lbl <- paste0(lbl, sprintf(" %dv", n_vac_p))
        writeData(wb, "Schedule", x = lbl,
          startRow = schr, startCol = N_HDR + ci, colNames = FALSE)
        addStyle(wb, "Schedule",
          mk(fg = C_BLUE_LT, font_color = F_NAVY, size = 8,
             border = "All", border_color = "#BBBBBB"),
          rows = schr, cols = N_HDR + ci)
      }
      # Fill structural and key cols of PP header
      for (col_fill in c(5:N_HDR, N_COLS)) {
        addStyle(wb, "Schedule",
          mk(fg = C_BLUE_LT, border = "All", border_color = "#BBBBBB"),
          rows = schr, cols = col_fill)
      }
      setRowHeights(wb, "Schedule", rows = schr, heights = 16)
      schr    <- schr + 1L
      prev_pp <- pp
    }

    date_str <- format(d, "%m/%d/%y")
    day_lbl  <- weekdays(d, abbreviate = TRUE)
    pp_lbl   <- ifelse(is.na(pp), "", pp)
    app1  <- ifelse(is.na(day$APP1),    "", day$APP1)
    app2  <- ifelse(is.na(day$APP2),    "", day$APP2)
    roam  <- ifelse(is.na(day$Roaming), "", day$Roaming)
    night <- ifelse(is.na(day$Night),   "", day$Night)

    bg_day   <- if (is_h) C_YELLOW else if (is_w) C_LAVENDER else "#FFFFFF"
    bg_night <- C_NIGHT

    # ── Day-shift row ────────────────────────────────────────────────────────
    dr <- schr
    writeData(wb, "Schedule",
      x = data.frame(Date = date_str, Day = day_lbl, PP = pp_lbl,
                     Shift = "6:30AM\u20136:30PM",
                     APP1 = app1, APP2 = app2, Roaming = roam, Night = "",
                     stringsAsFactors = FALSE),
      startRow = dr, startCol = 1, colNames = FALSE)
    writeData(wb, "Schedule",
      x = sprintf("%s|6:30AM-6:30PM", date_str),
      startRow = dr, startCol = N_COLS, colNames = FALSE)

    # Date / Day / PP cols
    addStyle(wb, "Schedule",
      mk(fg = bg_day, bold = TRUE, font_color = F_NAVY, halign = "left",
         border = "All", border_color = "#DDDDDD", size = 9),
      rows = dr, cols = 1:3)
    # Shift label
    addStyle(wb, "Schedule",
      mk(fg = bg_day, font_color = F_GRAY, halign = "left",
         border = "All", border_color = "#DDDDDD", size = 9),
      rows = dr, cols = 4)
    # APP1 / APP2 / Roaming
    for (j in 1:3) {
      val <- c(app1, app2, roam)[j]
      cbg <- if (nchar(val) > 0) (if (is_h) C_YELLOW else C_GREEN) else bg_day
      cfc <- if (nchar(val) > 0) F_BLUE else F_GRAY
      addStyle(wb, "Schedule",
        mk(fg = cbg, bold = nchar(val) > 0, font_color = cfc,
           border = "All", border_color = "#DDDDDD", size = 9),
        rows = dr, cols = 4L + j)
    }
    # Night col (empty in day row)
    addStyle(wb, "Schedule",
      mk(fg = bg_day, border = "All", border_color = "#DDDDDD", size = 9),
      rows = dr, cols = N_HDR)

    # Per-staff cols
    for (ci in seq_along(STAFF)) {
      person <- STAFF[ci]
      col    <- N_HDR + ci
      role   <- role_lookup[[ds]][[person]]
      if (is.na(role)) role <- ""
      show   <- if (role == "Night") "" else role  # night shown in night row
      cbg    <- { rb <- role_bg(show, is_h); if (is.null(rb)) bg_day else rb }
      cfc    <- if (nchar(show) > 0) role_fc(show) else F_GRAY
      bdr_c  <- "#DDDDDD"
      bdr_s  <- "thin"
      if (nchar(show) > 0 && is_dbn(d, person)) {
        bdr_c <- C_ORANGE; bdr_s <- "dashed"
      }
      writeData(wb, "Schedule", x = show,
        startRow = dr, startCol = col, colNames = FALSE)
      addStyle(wb, "Schedule",
        mk(fg = cbg, bold = nchar(show) > 0, font_color = cfc,
           border = "All", border_color = bdr_c, border_style = bdr_s, size = 9),
        rows = dr, cols = col)
    }
    addStyle(wb, "Schedule",
      mk(fg = "#F7F7F7", font_color = F_LGRAY, size = 7, halign = "left"),
      rows = dr, cols = N_COLS)
    setRowHeights(wb, "Schedule", rows = dr, heights = 15)

    # ── Night-shift row ──────────────────────────────────────────────────────
    nr <- schr + 1L
    writeData(wb, "Schedule",
      x = data.frame(Date = date_str, Day = "", PP = "",
                     Shift = "6:30PM\u20136:30AM",
                     APP1 = "", APP2 = "", Roaming = "", Night = night,
                     stringsAsFactors = FALSE),
      startRow = nr, startCol = 1, colNames = FALSE)
    writeData(wb, "Schedule",
      x = sprintf("%s|6:30PM-6:30AM", date_str),
      startRow = nr, startCol = N_COLS, colNames = FALSE)

    addStyle(wb, "Schedule",
      mk(fg = bg_night, bold = TRUE, font_color = F_NAVY, halign = "left",
         border = "All", border_color = "#BBBBBB", size = 9),
      rows = nr, cols = 1:3)
    addStyle(wb, "Schedule",
      mk(fg = bg_night, font_color = F_GRAY, halign = "left",
         border = "All", border_color = "#BBBBBB", size = 9),
      rows = nr, cols = 4)
    addStyle(wb, "Schedule",
      mk(fg = bg_night, border = "All", border_color = "#BBBBBB", size = 9),
      rows = nr, cols = 5:7)
    # Night cell
    cbg_n <- if (nchar(night) > 0) (if (is_h) C_YELLOW else C_NIGHT) else bg_night
    addStyle(wb, "Schedule",
      mk(fg = cbg_n, bold = nchar(night) > 0,
         font_color = if (nchar(night) > 0) F_NAVY else F_GRAY,
         border = "All", border_color = "#BBBBBB", size = 9),
      rows = nr, cols = N_HDR)

    # Per-staff cols (night row: show "Night" only for the night person)
    for (ci in seq_along(STAFF)) {
      person <- STAFF[ci]
      col    <- N_HDR + ci
      show   <- if (!is.na(day$Night) && day$Night == person) "Night" else ""
      # Also show off/vac/cme status in night row if not the night worker
      if (nchar(show) == 0L) {
        role <- role_lookup[[ds]][[person]]
        if (is.na(role)) role <- ""
        if (role %in% c("OFF","VAC","PTO","CME")) show <- role
      }
      cbg <- { rb <- role_bg(show); if (is.null(rb)) bg_night else rb }
      cfc <- if (nchar(show) > 0) role_fc(show) else F_GRAY
      writeData(wb, "Schedule", x = show,
        startRow = nr, startCol = col, colNames = FALSE)
      addStyle(wb, "Schedule",
        mk(fg = cbg, bold = nchar(show) > 0, font_color = cfc,
           border = "All", border_color = "#BBBBBB", size = 9),
        rows = nr, cols = col)
    }
    addStyle(wb, "Schedule",
      mk(fg = C_NIGHT, font_color = F_LGRAY, size = 7, halign = "left"),
      rows = nr, cols = N_COLS)
    setRowHeights(wb, "Schedule", rows = nr, heights = 15)

    schr <- schr + 2L
  }

  setColWidths(wb, "Schedule",
    cols   = seq_len(N_COLS),
    widths = c(11, 5, 5, 17, 13, 9, 13, 13, rep(9, N_STAFF), 18))

  # ── Save ───────────────────────────────────────────────────────────────────
  saveWorkbook(wb, output_path, overwrite = TRUE)
  message("  Saved: ", output_path)
  invisible(output_path)
}
