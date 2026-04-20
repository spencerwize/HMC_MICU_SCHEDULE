# ─────────────────────────────────────────────────────────────────────────────
# excel_output.R  —  Build 3-sheet Excel workbook matching reference format
#
# Sheet 1: Calendar   monthly view (two rows/week: date + assignment)
# Sheet 2: Summary    overview stats + PP-by-PP detail + staffing rules
# Sheet 3: Schedule   one row/day (day + night merged) + PP sub-headers
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
        return(if (s == "Night")  "Night" else
               if (s == "APP1")  "APP1"   else
               if (s == "APP2")  "APP2"   else "APP 3")
    }
    pdata <- time_off[[person]]
    m     <- pdata[pdata$date == d, ]
    typ   <- if (nrow(m) > 0) m$type[1] else NA_character_
    if (!is.na(typ)) {
      return(switch(typ, cme = "CME", off = "OFF", vac = "VAC", ""))
    }
    ""
  }

  role_bg <- function(role, is_holiday = FALSE) {
    if (is_holiday && role %in% c("APP1","APP2","APP 3","Night"))
      return(C_YELLOW)
    switch(role,
      APP1 = C_GREEN, APP2 = C_GREEN, "APP 3" = C_GREEN,
      Night = C_NIGHT,
      VAC  = C_PEACH,
      CME  = C_ORANGE, OFF = C_PINK,
      NULL)
  }

  role_fc <- function(role) {
    switch(role,
      APP1 = F_BLUE, APP2 = F_BLUE, "APP 3" = F_BLUE,
      Night = F_NAVY,
      VAC = F_BROWN,
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

  # ── Pre-compute Schedule row numbers (for Calendar formulas) ───────────────
  # HOLIDAY_NAMES comes from the global set by server.R when the sheet config
  # is applied — no hardcoding needed here.

  sched_row_map <- local({
    rm   <- list()
    cr   <- 2L   # row 1 = col header; row 2 = first PP header
    prev <- ""
    for (d_raw in all_d) {
      d  <- as.Date(d_raw, origin = "1970-01-01")
      pp <- get_pp(d)
      if (!is.na(pp) && pp != prev) { cr <- cr + 1L; prev <- pp }
      rm[[as.character(d)]] <- cr
      cr <- cr + 1L   # one row per day (day + night merged)
    }
    rm
  })

  # Calendar formula helpers — column range is derived from N_STAFF so that
  # adding/removing staff doesn't break the formulas.
  # Schedule sheet: cols 1-7 are fixed headers (Date/Day/PP/APP1/APP2/APP3/Night);
  # staff columns start at col N_HDR+1 (col 8 = H for N_HDR=7).
  col_letter <- function(n) {
    if (n <= 26L) LETTERS[n]
    else paste0(LETTERS[(n - 1L) %/% 26L], LETTERS[(n - 1L) %% 26L + 1L])
  }
  SCHED_STAFF_START <- N_HDR + 1L               # col 8 = H (for N_HDR=7)
  SCHED_STAFF_END   <- N_HDR + N_STAFF           # dynamic (Q for 10 staff, etc.)
  S_LTR  <- col_letter(SCHED_STAFF_START)        # "H"
  E_LTR  <- col_letter(SCHED_STAFF_END)          # dynamic

  # Single row per day — formulas reference just one row in the Schedule sheet.
  cal_role_formula <- function(row) {
    D <- sprintf("Schedule!$%s$%d:$%s$%d", S_LTR, row, E_LTR, row)
    M <- sprintf("MATCH($C$2,Schedule!$%s$1:$%s$1,0)", S_LTR, E_LTR)
    sprintf('IFERROR(IF(INDEX(%s,1,%s)<>"",INDEX(%s,1,%s),""),"")', D,M,D,M)
  }

  cal_hol_formula <- function(row, hol_name) {
    D <- sprintf("Schedule!$%s$%d:$%s$%d", S_LTR, row, E_LTR, row)
    M <- sprintf("MATCH($C$2,Schedule!$%s$1:$%s$1,0)", S_LTR, E_LTR)
    sprintf('"%s  "&IFERROR(IF(INDEX(%s,1,%s)<>"",INDEX(%s,1,%s),""),"")',
      hol_name, D,M,D,M)
  }

  # ════════════════════════════════════════════════════════════════════════════
  # SHEET 1 · Calendar
  # ════════════════════════════════════════════════════════════════════════════
  addWorksheet(wb, "Calendar")
  setColWidths(wb, "Calendar", cols = 1,   widths = 2.0)
  setColWidths(wb, "Calendar", cols = 2:8, widths = rep(13.0, 7))
  setColWidths(wb, "Calendar", cols = 9,   widths = 2.0)

  # Row 1: Title
  mergeCells(wb, "Calendar", cols = 2:8, rows = 1)
  writeData(wb, "Calendar",
    x = sprintf("APP Staff Schedule \u00B7 %s \u2013 %s",
                format(SCHEDULE_START, "%b %d"),
                format(SCHEDULE_END,   "%b %d, %Y")),
    startRow = 1, startCol = 2, colNames = FALSE)
  addStyle(wb, "Calendar",
    mk(fg = C_NAVY, bold = TRUE, size = 13, font_color = F_WHITE),
    rows = 1, cols = 2:8)
  setRowHeights(wb, "Calendar", rows = 1, heights = 27.75)

  # Row 2: Staff selector
  writeData(wb, "Calendar", x = "Staff member:",
    startRow = 2, startCol = 2, colNames = FALSE)
  addStyle(wb, "Calendar",
    mk(fg = "#DAE3F3", bold = TRUE, font_color = F_NAVY, halign = "right"),
    rows = 2, cols = 2)
  mergeCells(wb, "Calendar", cols = 3:5, rows = 2)
  writeData(wb, "Calendar", x = STAFF[1],
    startRow = 2, startCol = 3, colNames = FALSE)
  addStyle(wb, "Calendar",
    mk(fg = "#DAE3F3", bold = TRUE, font_color = F_NAVY, size = 12),
    rows = 2, cols = 3:5)
  dataValidation(wb, "Calendar", col = 3, rows = 2, type = "list",
    value = paste0('"', paste(STAFF, collapse = ","), '"'))
  addStyle(wb, "Calendar",
    mk(fg = "#DAE3F3", font_color = "#888888", size = 8, halign = "left"),
    rows = 2, cols = 6:8)
  setRowHeights(wb, "Calendar", rows = c(2L, 3L), heights = c(25.5, 6.0))

  months_list <- local({
    y <- as.integer(format(SCHEDULE_START, "%Y"))
    m <- as.integer(format(SCHEDULE_START, "%m"))
    y_end <- as.integer(format(SCHEDULE_END, "%Y"))
    m_end <- as.integer(format(SCHEDULE_END, "%m"))
    out <- list()
    repeat {
      out <- c(out, list(list(year = y, month = m,
        label = format(as.Date(sprintf("%d-%02d-01", y, m)), "%B %Y"))))
      if (y == y_end && m == m_end) break
      m <- m + 1L; if (m > 12L) { m <- 1L; y <- y + 1L }
    }
    out
  })
  dow_lbl <- c("Sun","Mon","Tue","Wed","Thu","Fri","Sat")

  cal_row <- 4L

  for (mo in months_list) {
    first_d <- as.Date(sprintf("%d-%02d-01", mo$year, mo$month))
    last_d  <- as.Date(sprintf("%d-%02d-01",
      mo$year + (mo$month == 12L), (mo$month %% 12L) + 1L)) - 1L

    # Month header row
    mergeCells(wb, "Calendar", cols = 2:8, rows = cal_row)
    writeData(wb, "Calendar", x = mo$label,
      startRow = cal_row, startCol = 2, colNames = FALSE)
    addStyle(wb, "Calendar",
      mk(fg = C_BLUE_LT, bold = TRUE, font_color = F_NAVY, size = 11),
      rows = cal_row, cols = 2:8)
    setRowHeights(wb, "Calendar", rows = cal_row, heights = 19.5)
    cal_row <- cal_row + 1L

    # DOW header row
    for (ci in seq_along(dow_lbl)) {
      writeData(wb, "Calendar", x = dow_lbl[ci],
        startRow = cal_row, startCol = ci + 1L, colNames = FALSE)
      addStyle(wb, "Calendar",
        mk(fg = "#203864", bold = TRUE, font_color = F_WHITE, size = 9,
           border = "All", border_color = F_NAVY),
        rows = cal_row, cols = ci + 1L)
    }
    setRowHeights(wb, "Calendar", rows = cal_row, heights = 15.75)
    cal_row <- cal_row + 1L

    # Iterate weeks (Sunday-anchored)
    first_sunday <- first_d - as.integer(format(first_d, "%w"))
    cur_sunday   <- first_sunday

    while (cur_sunday <= last_d) {
      date_row <- cal_row
      role_row <- cal_row + 1L

      for (dow in 0:6) {
        cur <- cur_sunday + as.integer(dow)
        col <- dow + 2L  # Sun=2(B) … Sat=8(H)
        is_wknd   <- dow %in% c(0L, 6L)
        in_month  <- cur >= first_d  && cur <= last_d
        in_sched  <- cur >= SCHEDULE_START && cur <= SCHEDULE_END
        is_hol    <- cur %in% HOLIDAY_DATES

        if (!in_month || !in_sched) {
          # Out-of-month or pre-schedule — style empty gray cells
          bg <- if (!in_month || !in_sched && !in_month) "#F7F7F7"
                else if (is_wknd) C_LAVENDER else "#FFFFFF"
          if (!in_month) bg <- "#F7F7F7"
          fc_num <- if (in_month && !in_sched) F_LGRAY else F_LGRAY
          if (in_month && !in_sched) {
            # In month but before/after schedule: show date grayed out
            writeData(wb, "Calendar", x = as.integer(format(cur, "%d")),
              startRow = date_row, startCol = col, colNames = FALSE)
          }
          addStyle(wb, "Calendar",
            mk(fg = "#F7F7F7", font_color = F_LGRAY, size = 8,
               halign = "left", valign = "center",
               border = "All", border_color = "#D0D0D0"),
            rows = date_row, cols = col)
          addStyle(wb, "Calendar",
            mk(fg = "#F7F7F7", size = 10, wrap = TRUE,
               border = "All", border_color = "#D0D0D0"),
            rows = role_row, cols = col)
        } else {
          # In-schedule date
          bg_d  <- if (is_hol) C_YELLOW else if (is_wknd) C_LAVENDER else "#FFFFFF"
          fc_d  <- if (is_hol) F_GOLD else "#555555"
          bg_r  <- bg_d  # role row background = same as date row

          # Date number cell
          writeData(wb, "Calendar", x = as.integer(format(cur, "%d")),
            startRow = date_row, startCol = col, colNames = FALSE)
          addStyle(wb, "Calendar",
            mk(fg = bg_d, font_color = fc_d, size = 8,
               halign = "left", valign = "center",
               border = "All", border_color = "#D0D0D0"),
            rows = date_row, cols = col)

          # Role cell — Excel formula (dynamic, responds to C2 dropdown)
          ds  <- as.character(cur)
          row <- sched_row_map[[ds]]
          hn  <- HOLIDAY_NAMES[format(cur, "%Y-%m-%d")]
          fml <- if (is_hol && !is.na(hn)) {
            cal_hol_formula(row, hn)
          } else {
            cal_role_formula(row)
          }
          writeFormula(wb, "Calendar", x = fml,
            startRow = role_row, startCol = col)
          addStyle(wb, "Calendar",
            mk(fg = bg_r, bold = TRUE, font_color = F_BLUE, size = 10,
               halign = "center", valign = "center",
               border = "All", border_color = "#D0D0D0", wrap = TRUE),
            rows = role_row, cols = col)
        }
      }

      setRowHeights(wb, "Calendar", rows = date_row, heights = 12.75)
      setRowHeights(wb, "Calendar", rows = role_row, heights = 21.75)
      cal_row    <- cal_row + 2L
      cur_sunday <- cur_sunday + 7L
    }

    # Inter-month spacer
    setRowHeights(wb, "Calendar", rows = cal_row, heights = 7.5)
    cal_row <- cal_row + 1L
  }

  # Legend rows (matches reference rows 59-60)
  mergeCells(wb, "Calendar", cols = 2:8, rows = cal_row)
  writeData(wb, "Calendar",
    x = paste0("APP1 / APP2 / APP 3 = day shift  \u00B7  Night = night shift  ",
               "\u00B7  VAC = vacation  \u00B7  CME = conference  \u00B7  OFF = not scheduled"),
    startRow = cal_row, startCol = 2, colNames = FALSE)
  addStyle(wb, "Calendar",
    mk(fg = "#FFFFFF", font_color = "#888888", size = 8, halign = "left"),
    rows = cal_row, cols = 2:8)
  setRowHeights(wb, "Calendar", rows = cal_row, heights = 13.5)
  cal_row <- cal_row + 1L
  mergeCells(wb, "Calendar", cols = 2:8, rows = cal_row)
  writeData(wb, "Calendar",
    x = "Select staff member in cell C2 to change the view",
    startRow = cal_row, startCol = 2, colNames = FALSE)
  addStyle(wb, "Calendar",
    mk(fg = "#FFFFFF", font_color = "#888888", size = 8, halign = "left"),
    rows = cal_row, cols = 2:8)
  setRowHeights(wb, "Calendar", rows = cal_row, heights = 15.75)

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
    x = sprintf("SCHEDULE SUMMARY \u00B7 %s \u2013 %s",
                format(SCHEDULE_START, "%b %d"),
                format(SCHEDULE_END,   "%b %d, %Y")),
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
    list("APP 3 Shifts",             "n_roam",   C_GRAY_LT, FALSE),
    list("Weekend Shifts",           "n_wknd",   "#FFFFFF",  FALSE),
    list("Vacation Days",            "n_vac",    C_PEACH,    FALSE),
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

  # Col layout: Date | Day | PP | APP1 | APP2 | APP3 | Night | [staff...] | _key_
  # Day and night are merged into one row per calendar day.
  N_HDR  <- 7L
  N_COLS <- N_HDR + N_STAFF + 1L

  hdr <- c("Date","Day","PP",
           "APP 1","APP 2","APP 3","Night Shift",
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
      # Fill slot cols and key col of PP header
      for (col_fill in c(4:N_HDR, N_COLS)) {
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

    bg_day <- if (is_h) C_YELLOW else if (is_w) C_LAVENDER else "#FFFFFF"

    # ── Single combined row (day + night) ────────────────────────────────────
    dr <- schr
    writeData(wb, "Schedule",
      x = data.frame(Date = date_str, Day = day_lbl, PP = pp_lbl,
                     APP1 = app1, APP2 = app2, APP3 = roam, Night = night,
                     stringsAsFactors = FALSE),
      startRow = dr, startCol = 1, colNames = FALSE)
    writeData(wb, "Schedule",
      x = date_str, startRow = dr, startCol = N_COLS, colNames = FALSE)

    # Date / Day / PP cols
    addStyle(wb, "Schedule",
      mk(fg = bg_day, bold = TRUE, font_color = F_NAVY, halign = "left",
         border = "All", border_color = "#DDDDDD", size = 9),
      rows = dr, cols = 1:3)

    # APP1 / APP2 / APP3 slot cols (4-6)
    day_vals <- c(app1, app2, roam)
    for (j in 1:3) {
      val     <- day_vals[j]
      is_app3 <- j == 3L
      cbg <- if (nchar(val) > 0) (if (is_h) C_YELLOW else C_GREEN) else
             if (is_app3) C_PEACH else bg_day
      cfc <- if (nchar(val) > 0) F_BLUE else F_GRAY
      addStyle(wb, "Schedule",
        mk(fg = cbg, bold = nchar(val) > 0, font_color = cfc,
           border = "All", border_color = "#DDDDDD", size = 9),
        rows = dr, cols = 3L + j)
    }

    # Night slot col (7)
    cbg_n <- if (nchar(night) > 0) (if (is_h) C_YELLOW else C_NIGHT) else bg_day
    addStyle(wb, "Schedule",
      mk(fg = cbg_n, bold = nchar(night) > 0,
         font_color = if (nchar(night) > 0) F_NAVY else F_GRAY,
         border = "All", border_color = "#DDDDDD", size = 9),
      rows = dr, cols = N_HDR)

    # Per-staff cols — show full role (day, night, or time-off) in one cell
    for (ci in seq_along(STAFF)) {
      person <- STAFF[ci]
      col    <- N_HDR + ci
      role   <- role_lookup[[ds]][[person]]
      if (is.na(role)) role <- ""
      is_night_role <- role == "Night"
      cbg <- {
        rb <- role_bg(role, is_h)
        if (is.null(rb)) (if (is_night_role) C_NIGHT else bg_day) else rb
      }
      cfc   <- if (nchar(role) > 0) role_fc(role) else F_GRAY
      bdr_c <- "#DDDDDD"; bdr_s <- "thin"
      if (role %in% c("APP1","APP2","APP 3") && is_dbn(d, person)) {
        bdr_c <- C_ORANGE; bdr_s <- "dashed"
      }
      writeData(wb, "Schedule", x = role,
        startRow = dr, startCol = col, colNames = FALSE)
      addStyle(wb, "Schedule",
        mk(fg = cbg, bold = nchar(role) > 0, font_color = cfc,
           border = "All", border_color = bdr_c, border_style = bdr_s, size = 9),
        rows = dr, cols = col)
    }
    addStyle(wb, "Schedule",
      mk(fg = "#F7F7F7", font_color = F_LGRAY, size = 7, halign = "left"),
      rows = dr, cols = N_COLS)
    setRowHeights(wb, "Schedule", rows = dr, heights = 18)

    schr <- schr + 1L
  }

  setColWidths(wb, "Schedule",
    cols   = seq_len(N_COLS),
    widths = c(14, 7, 7, 16, 14, 16, 16, rep(13, N_STAFF), 20))

  # ── Save ───────────────────────────────────────────────────────────────────
  saveWorkbook(wb, output_path, overwrite = TRUE)
  message("  Saved: ", output_path)
  invisible(output_path)
}
