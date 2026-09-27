# ─────────────────────────────────────────────────────────────────────────────
# excel_output.R  —  Build 3-sheet Excel workbook matching reference format
#
# Sheet 1: Calendar   monthly view (two rows/week: date + assignment)
# Sheet 2: Summary    overview stats + PP-by-PP detail + staffing rules
# Sheet 3: Schedule   one row/day (day + night merged) + PP sub-headers
# ─────────────────────────────────────────────────────────────────────────────

build_excel <- function(sched_obj, time_off, targets, output_path) {

  # ── Adopt the schedule frame carried by the ARGUMENTS ─────────────────────
  # STAFF / PAY_PERIODS / SCHEDULE_START / SCHEDULE_END are mutable globals that
  # run_pipeline() rewrites per sheet, and re-sourcing global.R resets them to
  # the constants.R defaults. Either leaves an existing `res` mismatched against
  # them, which used to fail deep inside a vapply ("result is length 0") or, in
  # older code, quietly render the wrong pay periods.
  #
  # Everything needed is already carried by the arguments, so derive the frame
  # from them. The globals are set (not merely shadowed) because helpers defined
  # in constants.R - get_pp(), get_pp_vec(), is_weekend_shift() - read the global
  # PAY_PERIODS, and are called from here. They are restored on exit, so calling
  # build_excel() has no lasting effect on the session.
  .saved <- list(STAFF = STAFF, PAY_PERIODS = PAY_PERIODS,
                 SCHEDULE_START = SCHEDULE_START, SCHEDULE_END = SCHEDULE_END)
  on.exit({
    STAFF          <<- .saved$STAFF
    PAY_PERIODS    <<- .saved$PAY_PERIODS
    SCHEDULE_START <<- .saved$SCHEDULE_START
    SCHEDULE_END   <<- .saved$SCHEDULE_END
  }, add = TRUE)

  if (!length(targets) || !length(sched_obj$dates))
    stop("build_excel(): `targets` or `sched_obj$dates` is empty - nothing to render.")

  STAFF          <<- names(targets)
  SCHEDULE_START <<- min(sched_obj$dates)
  SCHEDULE_END   <<- max(sched_obj$dates)
  PAY_PERIODS    <<- local({
    pps <- names(targets[[STAFF[1L]]])
    do.call(rbind, lapply(pps, function(pp) {
      pd <- targets[[STAFF[1L]]][[pp]]$pp_dates
      if (is.null(pd) || !length(pd))
        stop("build_excel(): targets for pay period '", pp, "' carry no pp_dates.")
      data.frame(name = pp, start = min(pd), end = max(pd), stringsAsFactors = FALSE)
    }))
  })

  # ── Staff who exist in the workbook but not in the solve ──────────────────
  # PENDING_STAFF (new hires) get columns, Summary rows and dropdown entries so
  # they can be assigned by hand. They carry no request-sheet data and no solver
  # output, so synthesise empty records for them here. `targets` and `time_off`
  # are local copies, and the per-person scheduler lookups are read through
  # accessors below rather than mutated - sched_obj is an R6 reference object and
  # writing to it would leak back into the caller's result.
  pending <- setdiff(if (exists("PENDING_STAFF")) PENDING_STAFF else character(0), STAFF)
  if (length(pending)) {
    zero_pp <- setNames(lapply(PAY_PERIODS$name, function(pp) {
      i <- match(pp, PAY_PERIODS$name)
      list(pp_name = pp, avail = 0L, credited = 0L, target = 0L,
           sched_target = 0L, pto_needed = 0L, pto_auto = 0L, soft_min = 0L,
           off_days = as.Date(character()), vac_days = as.Date(character()),
           cme_days = as.Date(character()), pto_days = as.Date(character()),
           green_days = as.Date(character()),
           pp_dates = seq(PAY_PERIODS$start[i], PAY_PERIODS$end[i], by = "day"))
    }), PAY_PERIODS$name)
    for (pp_new in pending) {
      targets[[pp_new]]  <- zero_pp
      time_off[[pp_new]] <- data.frame(date = as.Date(character()),
                                       type = character(), stringsAsFactors = FALSE)
    }
    STAFF <<- c(STAFF, pending)
    message(sprintf("  Including %d pending staff (no solver data): %s",
                    length(pending), paste(pending, collapse = ", ")))
  }
  # Scheduler lookups that must tolerate a person the solver never saw.
  ps_shifts <- function(p) { v <- sched_obj$person_shifts[[p]]
    if (is.null(v)) data.frame(date = as.Date(character()), slot = character(),
                               stringsAsFactors = FALSE) else v }
  ps_nights <- function(p) { v <- sched_obj$person_nights[[p]]
    if (is.null(v)) as.Date(character()) else v }
  ps_ppcount <- function(p, pp) { v <- sched_obj$pp_counts[[p]]
    if (is.null(v) || is.null(v[[pp]])) 0L else as.integer(v[[pp]]) }

  wb    <- createWorkbook()
  all_d <- sched_obj$dates


  # ── Color palette ──────────────────────────────────────────────────────────
  C_NAVY     <- "#1F3864"
  C_BLUE     <- "#2E75B6"
  C_BLUE_LT  <- "#D6E4F3"
  C_NIGHT    <- "#D6E4F0"
  C_FC3      <- "#E4DFEC"   # FC3 service - distinct from MICU day/night
  F_FC3      <- "#5B3A8E"
  C_LAVENDER <- "#EEF1FF"
  C_GREEN    <- "#E2EFDA"
  C_YELLOW   <- "#FFF2CC"
  C_PEACH    <- "#FCE4D6"
  C_PINK     <- "#F4CCCC"
  C_PTO      <- "#FF99CC"
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
  # Delegates to role_of() in R/roles.R — shared with the Schedule Grid and
  # the Calendar tab so the three views cannot drift apart again.
  person_role <- function(person, d)
    role_of(person, d, sched_obj$schedule, time_off, sched_obj$granted_pto)

  role_bg <- function(role, is_holiday = FALSE) {
    if (is_holiday && role %in% c("Day","Night"))
      return(C_YELLOW)
    switch(role,
      Day = C_GREEN,
      Night = C_NIGHT,
      CME  = C_ORANGE, OFF = C_PINK, PTO = C_PTO,
      # "Yellow" = the person asked to avoid this day. Pale amber fill so it
      # reads as a soft flag, clearly distinct from OFF (pink) and from a
      # blank Green day, which carries no fill at all.
      Yellow = C_PEACH,
      NULL)
  }

  role_fc <- function(role) {
    switch(role,
      Day = F_BLUE,
      Night = F_NAVY,
      CME = F_WHITE, OFF = F_RED, PTO = F_RED,
      Yellow = "#8A6D00",                    # dark amber on the peach fill
      "#000000")
  }


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
  N_STAFF <- length(STAFF)
  # ── Schedule sheet column layout (single source of truth) ─────────────────
  #   A Date | B Day | C PP | D E F  MICU day | G MICU night | H I J  FC3 |
  #   [staff...] | _key_
  MICU_COL_FIRST <- 4L
  MICU_COL_LAST  <- 7L                      # D-G
  FC3_COL_FIRST  <- 8L
  FC3_COL_LAST   <- 10L                     # H-J
  N_HDR   <- FC3_COL_LAST                   # header cols before the staff block
  N_PP    <- nrow(PAY_PERIODS)

  # ── Pre-compute Schedule row numbers (for Calendar formulas) ───────────────
  sched_row_map <- local({
    rm   <- list()
    cr   <- 3L   # rows 1-2 = group banner + col header; row 3 = first PP header
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

  # The Schedule sheet carries a group banner on row 1 (MICU / FC3) and the real
  # column headers on row 2, so every MATCH against the staff-name header row
  # must target row 2.
  SCHED_HDR_ROW <- 2L

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

  # Helpers reused by both Summary live formulas and Schedule slot formulas.
  # person_pc(ci) returns the Schedule-sheet column letter for the ci-th staff member.
  person_pc     <- function(ci) col_letter(SCHED_STAFF_START + ci - 1L)
  MAX_SCHED_ROW <- 200L   # row ceiling for COUNTIFS / SUMPRODUCT on Schedule sheet

  # Single row per day — formulas reference just one row in the Schedule sheet.
  cal_role_formula <- function(row) {
    D <- sprintf("Schedule!$%s$%d:$%s$%d", S_LTR, row, E_LTR, row)
    M <- sprintf("MATCH($C$2,Schedule!$%s$2:$%s$2,0)", S_LTR, E_LTR)
    sprintf('IFERROR(IF(INDEX(%s,1,%s)<>"",INDEX(%s,1,%s),""),"")', D,M,D,M)
  }

  cal_hol_formula <- function(row, hol_name) {
    D <- sprintf("Schedule!$%s$%d:$%s$%d", S_LTR, row, E_LTR, row)
    M <- sprintf("MATCH($C$2,Schedule!$%s$2:$%s$2,0)", S_LTR, E_LTR)
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

  # ── "Remaining to fill" block ───────────────────────────────────────────────
  # Per pay period, for whoever is selected in the C2 dropdown: how many shifts
  # they are still short of their target. Entirely formula-driven, so it updates
  # with the dropdown like the calendar below it.
  #
  # Target lookup uses an inline array constant per pay period ({6,5,6,...}, one
  # entry per staff member in STAFF order) indexed by the same MATCH the calendar
  # formulas use. That keeps the block self-contained - no helper cells to hide,
  # and nothing to break if rows shift.
  cal_row <- 4L
  rem_first_row <- cal_row + 2L
  mergeCells(wb, "Calendar", cols = 2:8, rows = cal_row)
  writeData(wb, "Calendar", x = "REMAINING TO FILL (this person, by pay period)",
    startRow = cal_row, startCol = 2, colNames = FALSE)
  addStyle(wb, "Calendar",
    mk(fg = C_BLUE, bold = TRUE, size = 11, font_color = F_WHITE),
    rows = cal_row, cols = 2:8)
  setRowHeights(wb, "Calendar", rows = cal_row, heights = 19.5)
  cal_row <- cal_row + 1L

  rem_hdr <- c("Pay period", "Dates", "Target", "Scheduled", "Remaining",
               "CME", "PTO")
  for (ci in seq_along(rem_hdr)) {
    writeData(wb, "Calendar", x = rem_hdr[ci],
      startRow = cal_row, startCol = ci + 1L, colNames = FALSE)
    addStyle(wb, "Calendar",
      mk(fg = C_GRAY_LT, bold = TRUE, size = 9, halign = "center",
         border = "All", border_color = "#BFBFBF"),
      rows = cal_row, cols = ci + 1L)
  }
  cal_row <- cal_row + 1L

  MATCH_P <- sprintf("MATCH($C$2,Schedule!$%s$2:$%s$2,0)", S_LTR, E_LTR)
  SCOL    <- sprintf("Schedule!$%s$1:$%s$%d", S_LTR, E_LTR, MAX_SCHED_ROW)

  for (ppi in seq_len(nrow(PAY_PERIODS))) {
    pp_name <- PAY_PERIODS$name[ppi]
    pdates  <- seq(PAY_PERIODS$start[ppi], PAY_PERIODS$end[ppi], by = "day")
    pdates  <- pdates[pdates >= SCHEDULE_START & pdates <= SCHEDULE_END]
    if (!length(pdates)) next
    rws <- unlist(sched_row_map[as.character(pdates)])
    r1  <- min(rws); r2 <- max(rws)

    # Per-person target for this PP, in STAFF order, as an Excel array constant.
    # CME comes from the request sheet and is fixed, so it rides as an array
    # constant. PTO is LIVE - see below - because granting a PTO day on the
    # Schedule sheet has to move these numbers.
    cme_vec <- vapply(STAFF, function(p)
      as.numeric(length(targets[[p]][[pp_name]]$cme_days)), numeric(1L))
    # The automatic pto_reduction() figure acts as a floor under the typed count.
    pto_auto_vec <- vapply(STAFF, function(p)
      as.numeric(targets[[p]][[pp_name]]$pto_auto), numeric(1L))
    # BASE target (normally 6). The live target is base - CME - PTO, so granting
    # PTO lowers the shifts owed rather than leaving a stale number here.
    base_vec <- vapply(STAFF, function(p) {
      ti <- targets[[p]][[pp_name]]
      as.numeric(ti$sched_target + ti$credited + ti$pto_needed)
    }, numeric(1L))
    cme_arr  <- paste0("{", paste(cme_vec,      collapse = ","), "}")
    ptoa_arr <- paste0("{", paste(pto_auto_vec, collapse = ","), "}")
    base_arr <- paste0("{", paste(base_vec,     collapse = ","), "}")

    rng   <- sprintf("INDEX(%s,%d,%s):INDEX(%s,%d,%s)", SCOL, r1, MATCH_P, SCOL, r2, MATCH_P)
    # A "scheduled" cell is one holding a work role; OFF / CME / Yellow / blank are not.
    # FC3 counts as a worked shift, so it belongs in the pay-period total here
    # exactly as it does in the Summary.
    cnt   <- sprintf(
      'COUNTIF(%s,"Day")+COUNTIF(%s,"Night")+COUNTIF(%s,"FC3")+COUNTIF(%s,"FC3 Night")',
      rng, rng, rng, rng)
    # PTO actually granted this period: days typed "PTO" in the person's column,
    # floored at the automatic figure.
    f_pto <- sprintf('IFERROR(MAX(COUNTIF(%s,"PTO"),INDEX(%s,1,%s)),0)',
                     rng, ptoa_arr, MATCH_P)
    f_cme <- sprintf("IFERROR(INDEX(%s,1,%s),0)", cme_arr, MATCH_P)
    # Shifts owed = base - CME - PTO, so an extra PTO day lowers the target.
    f_tgt <- sprintf("MAX(0,IFERROR(INDEX(%s,1,%s),0)-%s-%s)",
                     base_arr, MATCH_P, f_cme, f_pto)
    f_cnt <- sprintf("IFERROR(%s,0)", cnt)
    f_rem <- sprintf("MAX(0,%s-%s)", f_tgt, f_cnt)

    writeData(wb, "Calendar", x = pp_name, startRow = cal_row, startCol = 2, colNames = FALSE)
    writeData(wb, "Calendar",
      x = sprintf("%s \u2013 %s", format(min(pdates), "%b %d"), format(max(pdates), "%b %d")),
      startRow = cal_row, startCol = 3, colNames = FALSE)
    writeFormula(wb, "Calendar", x = f_tgt, startRow = cal_row, startCol = 4)
    writeFormula(wb, "Calendar", x = f_cnt, startRow = cal_row, startCol = 5)
    writeFormula(wb, "Calendar", x = f_rem, startRow = cal_row, startCol = 6)
    writeFormula(wb, "Calendar", x = f_cme, startRow = cal_row, startCol = 7)
    writeFormula(wb, "Calendar", x = f_pto, startRow = cal_row, startCol = 8)

    addStyle(wb, "Calendar", mk(bold = TRUE, size = 10, halign = "left",
      border = "All", border_color = "#D0D0D0"), rows = cal_row, cols = 2)
    addStyle(wb, "Calendar", mk(size = 9, font_color = "#666666", halign = "left",
      border = "All", border_color = "#D0D0D0"), rows = cal_row, cols = 3)
    for (cc in 4:5)
      addStyle(wb, "Calendar", mk(size = 10, halign = "center",
        border = "All", border_color = "#D0D0D0"), rows = cal_row, cols = cc)
    addStyle(wb, "Calendar", mk(bold = TRUE, size = 10, halign = "center",
      border = "All", border_color = "#D0D0D0"), rows = cal_row, cols = 6)
    addStyle(wb, "Calendar", mk(fg = C_ORANGE, size = 10, halign = "center",
      font_color = F_WHITE, border = "All", border_color = "#D0D0D0"),
      rows = cal_row, cols = 7)
    addStyle(wb, "Calendar", mk(fg = C_PEACH, size = 10, halign = "center",
      font_color = F_BROWN, border = "All", border_color = "#D0D0D0"),
      rows = cal_row, cols = 8)
    cal_row <- cal_row + 1L
  }

  # Totals row
  rem_last_row <- cal_row - 1L
  writeData(wb, "Calendar", x = "TOTAL", startRow = cal_row, startCol = 2, colNames = FALSE)
  for (cc in 4:8)
    writeFormula(wb, "Calendar",
      x = sprintf("SUM(%s%d:%s%d)", LETTERS[cc], rem_first_row, LETTERS[cc], rem_last_row),
      startRow = cal_row, startCol = cc)
  for (cc in 2:8)
    addStyle(wb, "Calendar",
      mk(fg = C_GRAY_LT, bold = TRUE, size = 10,
         halign = if (cc <= 3) "left" else "center",
         border = "All", border_color = "#BFBFBF"),
      rows = cal_row, cols = cc)
  # Highlight any pay period still short.
  conditionalFormatting(wb, "Calendar",
    cols = 6, rows = rem_first_row:rem_last_row,
    rule = ">0", style = mk(fg = C_PINK, bold = TRUE, font_color = F_RED,
                            halign = "center", border = "All",
                            border_color = "#D0D0D0"))
  cal_row <- cal_row + 2L
  cal_months_start <- cal_row   # CF applies from here down, not over the block above

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
          # ── Understaffed-day marker (schedule-level, not person-level) ────
          # A fully staffed day is 3 day workers (APP1 + APP2 + APP 3) and 1
          # night. Anything less gets a dotted border - subtle enough not to
          # clutter the grid, distinct enough to scan for. A missing NIGHT is
          # the more serious gap, so it reads red; a thin day crew reads amber.
          # LIVE via conditional formatting: both rules read the Schedule sheet's
          # slot block for this date, so filling a slot there clears the marker
          # here. Red (missing night) is added first so it takes priority when
          # both apply.
          addStyle(wb, "Calendar",
            mk(fg = bg_r, bold = TRUE, font_color = F_BLUE, size = 10,
               halign = "center", valign = "center",
               border = "All", border_color = "#D0D0D0", wrap = TRUE),
            rows = role_row, cols = col)
          for (rr in c(role_row, date_row)) {
            conditionalFormatting(wb, "Calendar", cols = col, rows = rr,
              rule = sprintf('Schedule!$G$%d=""', row),
              style = createStyle(border = "TopBottomLeftRight",
                                  borderColour = "#C00000", borderStyle = "dotted",
                                  fontColour = "#C00000"))
            conditionalFormatting(wb, "Calendar", cols = col, rows = rr,
              rule = sprintf('COUNTA(Schedule!$D$%d:$G$%d)<4', row, row),
              style = createStyle(border = "TopBottomLeftRight",
                                  borderColour = "#E8A33D", borderStyle = "dotted",
                                  fontColour = "#E8A33D"))
          }
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
               "\u00B7  CME = conference  \u00B7  OFF = vacation / day off"),
    startRow = cal_row, startCol = 2, colNames = FALSE)
  addStyle(wb, "Calendar",
    mk(fg = "#FFFFFF", font_color = "#888888", size = 8, halign = "left"),
    rows = cal_row, cols = 2:8)
  setRowHeights(wb, "Calendar", rows = cal_row, heights = 13.5)
  cal_row <- cal_row + 1L
  mergeCells(wb, "Calendar", cols = 2:8, rows = cal_row)
  writeData(wb, "Calendar",
    x = paste0("Dotted amber cell = fewer than 3 day workers that day  ·  ",
               "dotted red = night shift unfilled"),
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

  # Conditional formatting for role cells — fires on formula result so updates
  # dynamically when the staff dropdown in C2 changes.
  # CF only overrides fill/font; the static dotted borders that mark understaffed
  # days are preserved, since these CF styles carry no border definition.
  # The range starts below the "Remaining to fill" block so its text is untouched.
  cf_end <- cal_row
  conditionalFormatting(wb, "Calendar", cols = 2:8, rows = cal_months_start:cf_end,
    type = "contains", rule = "Day",
    style = createStyle(fgFill = C_GREEN, fontColour = F_BLUE,
                        textDecoration = "bold", halign = "center"))
  conditionalFormatting(wb, "Calendar", cols = 2:8, rows = cal_months_start:cf_end,
    type = "contains", rule = "Night",
    style = createStyle(fgFill = C_NIGHT, fontColour = F_NAVY,
                        textDecoration = "bold", halign = "center"))
  conditionalFormatting(wb, "Calendar", cols = 2:8, rows = cal_months_start:cf_end,
    type = "contains", rule = "FC3",
    style = createStyle(fgFill = C_FC3, fontColour = F_FC3,
                        textDecoration = "bold", halign = "center"))
  conditionalFormatting(wb, "Calendar", cols = 2:8, rows = cal_months_start:cf_end,
    type = "contains", rule = "OFF",
    style = createStyle(fgFill = C_PINK, fontColour = F_RED,
                        textDecoration = "bold", halign = "center"))
  # CME / PTO / Yellow were missing here, so those days fell through to the bare
  # row background while the Schedule sheet coloured them. Same palette as
  # role_bg()/role_fc() so the two views agree.
  conditionalFormatting(wb, "Calendar", cols = 2:8, rows = cal_months_start:cf_end,
    type = "contains", rule = "CME",
    style = createStyle(fgFill = C_ORANGE, fontColour = F_WHITE,
                        textDecoration = "bold", halign = "center"))
  conditionalFormatting(wb, "Calendar", cols = 2:8, rows = cal_months_start:cf_end,
    type = "contains", rule = "PTO",
    style = createStyle(fgFill = C_PTO, fontColour = F_RED,
                        textDecoration = "bold", halign = "center"))
  conditionalFormatting(wb, "Calendar", cols = 2:8, rows = cal_months_start:cf_end,
    type = "contains", rule = "Yellow",
    style = createStyle(fgFill = C_PEACH, fontColour = "#8A6D00",
                        halign = "center"))

  # ════════════════════════════════════════════════════════════════════════════
  # SHEET 2 · Summary
  # ════════════════════════════════════════════════════════════════════════════
  addWorksheet(wb, "Summary")

  SUM_COLS       <- 1L + N_STAFF
  PP_EMPTY_COL   <- SUM_COLS + 1L          # "Empty shifts" col on Pay Period Detail
  HLPR_COL_START <- SUM_COLS + 2L          # first hidden helper column (col 13 for 10 staff)
  # "n_roam" (APP 3 count) intentionally absent: team assignment is decided by
  # hand, so the counter carried no decision value on this sheet.
  LIVE_KEYS      <- c("n_sched", "n_total", "n_night", "n_wknd", "n_pto", "n_bump")

  # `ncol` lets a section span past the staff columns (the Pay Period Detail
  # table carries one extra column for empty shifts).
  sec_hdr <- function(row, label, ncol = SUM_COLS) {
    mergeCells(wb, "Summary", cols = 1:ncol, rows = row)
    writeData(wb, "Summary", x = label,
      startRow = row, startCol = 1, colNames = FALSE)
    addStyle(wb, "Summary",
      mk(fg = C_BLUE, bold = TRUE, font_color = F_WHITE,
         halign = "left", size = 11),
      rows = row, cols = 1:ncol)
    setRowHeights(wb, "Summary", rows = row, heights = 20)
  }

  staff_hdr <- function(row, extra = NULL) {
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
    if (!is.null(extra)) {
      writeData(wb, "Summary", x = extra,
        startRow = row, startCol = PP_EMPTY_COL, colNames = FALSE)
      addStyle(wb, "Summary",
        mk(fg = C_BLUE, bold = TRUE, font_color = F_WHITE,
           border = "All", border_color = F_WHITE, size = 9, wrap = TRUE),
        rows = row, cols = PP_EMPTY_COL)
    }
    setRowHeights(wb, "Summary", rows = row, heights = 18)
  }

  # Precompute stats
  pstats <- lapply(STAFF, function(person) {
    shifts  <- ps_shifts(person)
    nights  <- ps_nights(person)
    n_cred  <- sum(sapply(PAY_PERIODS$name, function(pp)
      targets[[person]][[pp]]$credited))
    pdata   <- time_off[[person]]
    n_vac   <- sum(pdata$type == "vac", na.rm = TRUE)
    # Explicit PTO days are requested days off too; count them so the
    # "Req. Off Days" row reflects every day the person will not work.
    n_off   <- sum(pdata$type %in% c("off", "pto"), na.rm = TRUE)
    n_pto   <- 0L
    n_bump  <- 0L
    for (ppn in PAY_PERIODS$name) {
      ppi    <- targets[[person]][[ppn]]
      # PTO per PP = max(explicitly requested PTO days, formula on off/vac/pto).
      n_pto  <- n_pto + (if (is.null(ppi$pto_needed)) 0L else ppi$pto_needed)
      actual <- ps_ppcount(person, ppn)
      if (actual < ppi$sched_target)
        n_bump <- n_bump + (ppi$sched_target - actual)
    }
    list(
      n_sched  = nrow(shifts) + length(nights),
      n_day    = nrow(shifts),
      n_night  = length(nights),
      n_roam   = sum(shifts$slot == "Roaming"),
      n_wknd   = sum(is_weekend_shift(shifts$date, shifts$slot)) +
                 sum(is_weekend_shift(nights, "Night")),
      n_cred   = n_cred,
      n_total  = nrow(shifts) + length(nights) + n_cred,
      n_reqoff = n_vac + n_off,
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
    list("Weekend Shifts",           "n_wknd",   "#FFFFFF",  FALSE),
    list("Req. Off Days",             "n_reqoff", C_PEACH,    FALSE),
    list("PTO Needed",                "n_pto",    "#FFFFFF",  FALSE),
    list("Shortfall (shift-days)",   "n_bump",   "#FFFFFF",  FALSE))

  for (ov in ovr_rows) {
    label <- ov[[1]]; key <- ov[[2]]; row_bg <- ov[[3]]; cred_flag <- ov[[4]]
    writeData(wb, "Summary", x = label,
      startRow = srow, startCol = 1, colNames = FALSE)
    addStyle(wb, "Summary",
      mk(fg = row_bg, bold = TRUE, font_color = F_NAVY,
         halign = "left", border = "All", border_color = "#DDDDDD"),
      rows = srow, cols = 1)
    if (key == "n_bump") bump_srow <- srow
    for (ci in seq_along(STAFF)) {
      val   <- pstats[[STAFF[ci]]][[key]]
      cbg   <- row_bg; cfc <- F_NAVY; cbold <- FALSE
      if (cred_flag        && val > 0) { cbg <- C_ORANGE;  cfc <- F_WHITE;    cbold <- TRUE }
      if (key == "n_pto"   && val > 0) { cbg <- "#FF9999"; cfc <- F_RED;      cbold <- TRUE }
      if (key == "n_bump"  && val > 0) { cbg <- "#FFE0E0"; cfc <- "#C00000";  cbold <- TRUE }
      if (key == "n_reqoff"&& val > 0) { cbg <- C_PEACH;   cfc <- F_BROWN }

      if (key %in% LIVE_KEYS) {
        pc  <- person_pc(ci)
        # Shifts worked = Day + Night cells in this person's Schedule column,
        # which are themselves live off the slot columns.
        n_sched_fml <- sprintf(
          paste0('COUNTIF(Schedule!$%1$s:$%1$s,"Day")+COUNTIF(Schedule!$%1$s:$%1$s,"Night")',
                 '+COUNTIF(Schedule!$%1$s:$%1$s,"FC3")',
                 '+COUNTIF(Schedule!$%1$s:$%1$s,"FC3 Night")'), pc)
        fml <- if (key == "n_sched") {
          n_sched_fml
        } else if (key == "n_total") {
          # + credited CME days, which are an input and stay static.
          sprintf('%s+%d', n_sched_fml, as.integer(pstats[[STAFF[ci]]]$n_cred))
        } else if (key == "n_pto") {
          # Live: per pay period, max(PTO days typed on the Schedule sheet,
          # the automatic pto_reduction() figure), summed across the schedule.
          # Granting a PTO day on the Schedule sheet moves this immediately.
          paste(vapply(seq_len(N_PP), function(k) {
            ppn_k <- PAY_PERIODS$name[k]
            sprintf('MAX(COUNTIFS(Schedule!$C$2:$C$%1$d,"%2$s",Schedule!$%3$s$2:$%3$s$%1$d,"PTO"),%4$d)',
                    MAX_SCHED_ROW, ppn_k, pc,
                    as.integer(targets[[STAFF[ci]]][[ppn_k]]$pto_auto))
          }, character(1L)), collapse = "+")
        } else if (key == "n_night") {
          sprintf('COUNTIF(Schedule!$%1$s:$%1$s,"Night")', pc)
        } else if (key == "n_wknd") {
          # Weekend = Friday NIGHT plus every shift on Sat/Sun. Two terms: the
          # Sat/Sun block (Day or Night), and Friday nights only.
          sprintf(paste0(
            'SUMPRODUCT(((Schedule!$B$2:$B$%1$d="Sat")+(Schedule!$B$2:$B$%1$d="Sun"))',
            '*((Schedule!$%2$s$2:$%2$s$%1$d="Day")',
            '+(Schedule!$%2$s$2:$%2$s$%1$d="Night")',
            '+(Schedule!$%2$s$2:$%2$s$%1$d="FC3")',
            '+(Schedule!$%2$s$2:$%2$s$%1$d="FC3 Night")))',
            '+SUMPRODUCT((Schedule!$B$2:$B$%1$d="Fri")',
            '*(Schedule!$%2$s$2:$%2$s$%1$d="Night"))'),
            MAX_SCHED_ROW, pc)
        } else {
          # n_bump: compare per-PP targets (rows 1:N_PP) vs live COUNTIFS actuals (rows N_PP+1:2*N_PP)
          hc <- col_letter(HLPR_COL_START + ci - 1L)
          sprintf('SUMPRODUCT((%1$s$1:%1$s$%2$d>%1$s$%3$d:%1$s$%4$d)*(%1$s$1:%1$s$%2$d-%1$s$%3$d:%1$s$%4$d))',
            hc, N_PP, N_PP + 1L, 2L * N_PP)
        }
        writeFormula(wb, "Summary", x = fml, startRow = srow, startCol = 1L + ci)
        cbg <- if (key == "n_bump") "#FFFAF0" else "#FFFFFF"
        cbold <- FALSE; cfc <- F_NAVY
        # n_pto used to be styled from its static value; as a live cell the
        # highlight has to be conditional on the computed result.
        if (key == "n_pto")
          conditionalFormatting(wb, "Summary", cols = 1L + ci, rows = srow,
            rule = ">0",
            style = createStyle(fgFill = "#FF9999", fontColour = F_RED,
                                textDecoration = "bold"))
      } else {
        writeData(wb, "Summary", x = val, startRow = srow, startCol = 1L + ci, colNames = FALSE)
      }
      addStyle(wb, "Summary",
        mk(fg = cbg, bold = cbold, font_color = cfc,
           border = "All", border_color = "#DDDDDD"),
        rows = srow, cols = 1L + ci)
    }
    setRowHeights(wb, "Summary", rows = srow, heights = 16)
    srow <- srow + 1L
  }
  srow <- srow + 1L  # spacer

  # Conditional formatting for shortfall row: red when live formula > 0
  if (exists("bump_srow")) {
    conditionalFormatting(wb, "Summary",
      cols  = 2L:SUM_COLS, rows = bump_srow,
      type  = "expression",
      rule  = sprintf("B%d>0", bump_srow),
      style = createStyle(fgFill = "#FFE0E0", fontColour = "#C00000",
                          textDecoration = "bold"))
  }

  # ── Pay Period Detail section ─────────────────────────────────────────────
  sec_hdr(srow, "Pay Period Detail", ncol = PP_EMPTY_COL); srow <- srow + 1L
  staff_hdr(srow, extra = "Empty shifts");                 srow <- srow + 1L

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
      actual  <- ps_ppcount(person, ppn)
      ppi     <- targets[[person]][[ppn]]
      n_vac_p <- sum(time_off[[person]]$type == "vac" &
        time_off[[person]]$date >= PAY_PERIODS$start[i] &
        time_off[[person]]$date <= PAY_PERIODS$end[i])
      # Denominator is the BASE pay-period target (normally 6), not sched_target.
      # sched_target is already net of CME and PTO, which made the denominator
      # vary (5, 4 ...) and read as though the target itself had moved.
      # Reconstructed rather than re-derived from BASE_TARGETS so it cannot drift
      # from targets.R:  sched_target = base - pto_needed - credited.
      base_tgt <- ppi$sched_target + ppi$credited + ppi$pto_needed
      # Numerator counts everything that fills the period: shifts worked, CME
      # days credited, and PTO days. So "6/6" always means the pay period is
      # complete, however it was made up.
      # ── LIVE cell: "<worked+CME+PTO>/<base>  N CME, N PTO, N vac" ──────────
      # Worked shifts come from the hidden helper (a COUNTIFS over the slot
      # columns for this pay period), so slot edits flow straight through.
      #
      # PTO is live too: typing "PTO" over a person's cell on the Schedule sheet
      # for a date in this period grants them a day, and the cell moves e.g.
      # "5/6  1 PTO" -> "6/6  2 PTO". Granting PTO raises the numerator because
      # a PTO day counts toward the period, exactly as CME does.
      #
      # It is max(typed PTO days, automatic PTO) so the pto_reduction() formula
      # still applies as a floor - the same rule compute_targets() uses.
      hc  <- col_letter(HLPR_COL_START + ci - 1L)
      pc  <- person_pc(ci)
      P   <- sprintf('MAX(COUNTIFS(Schedule!$C$2:$C$%1$d,"%2$s",Schedule!$%3$s$2:$%3$s$%1$d,"PTO"),%4$d)',
                     MAX_SCHED_ROW, ppn, pc, as.integer(ppi$pto_auto))
      # CME and vacation are inputs from the request sheet and stay static.
      static_parts <- character(0)
      if (ppi$credited > 0) static_parts <- c(static_parts, sprintf("%d CME", ppi$credited))
      if (n_vac_p      > 0) static_parts <- c(static_parts, sprintf("%d vac", n_vac_p))
      static_txt <- paste(static_parts, collapse = ", ")
      # Suffix has to be built in-formula so the PTO count can vary.
      lead <- if (nzchar(static_txt)) sprintf('"  %s, "', static_txt) else '"  "'
      none <- if (nzchar(static_txt)) sprintf('"  %s"',   static_txt) else '""'
      suffix_expr <- sprintf('IF(%1$s>0,%2$s&%1$s&" PTO",%3$s)', P, lead, none)
      fml <- sprintf('(%s%d+%d+%s)&"/%d"&%s',
                     hc, N_PP + i, as.integer(ppi$credited), P,
                     base_tgt, suffix_expr)
      writeFormula(wb, "Summary", x = fml, startRow = srow, startCol = 1L + ci)
      # Colour answers one question only: is this pay period filled? Applied as
      # conditional formatting on the live count, so it tracks edits too. All
      # complete periods share one fill regardless of HOW they were filled.
      addStyle(wb, "Summary",
        mk(fg = C_GREEN, font_color = F_NAVY,
           border = "All", border_color = "#DDDDDD", size = 9),
        rows = srow, cols = 1L + ci)
      live_expr <- sprintf('%s%d+%d+%s', hc, N_PP + i, as.integer(ppi$credited), P)
      # Short of the period target: red.
      conditionalFormatting(wb, "Summary", cols = 1L + ci, rows = srow,
        rule = sprintf('%s<%d', live_expr, base_tgt),
        # Dark red on a stronger pink: the old pale pink sat too close to the
        # pale green of a complete period to read at a glance.
        style = createStyle(fgFill = "#FFC7CE", fontColour = "#9C0006",
                            textDecoration = "bold"))
      # OVER the period target: blue. Reachable now that PTO can be granted from
      # the Schedule sheet - granting a day to someone already at 6/6 pushes them
      # to 7/6 until a shift is freed up. Worth seeing rather than reading as
      # "complete".
      conditionalFormatting(wb, "Summary", cols = 1L + ci, rows = srow,
        rule = sprintf('%s>%d', live_expr, base_tgt),
        style = createStyle(fgFill = "#D6E4F3", fontColour = "#0066CC",
                            textDecoration = "bold"))
    }
    # ── Empty shifts in this pay period ───────────────────────────────────
    # A fully staffed day is 3 day workers (APP1 + APP2 + APP 3) plus a night,
    # so this counts every one of those four slots left unfilled across the pay
    # period - the same standard the calendar's dotted borders use.
    # LIVE: blank cells in the slot block D:G across this pay period's rows.
    pp_days <- all_d[all_d >= PAY_PERIODS$start[i] & all_d <= PAY_PERIODS$end[i]]
    prow    <- unlist(sched_row_map[as.character(pp_days)])
    writeFormula(wb, "Summary",
      x = sprintf('COUNTBLANK(Schedule!$D$%d:$G$%d)', min(prow), max(prow)),
      startRow = srow, startCol = PP_EMPTY_COL)
    addStyle(wb, "Summary",
      mk(fg = C_GREEN, font_color = F_NAVY,
         border = "All", border_color = "#DDDDDD", size = 9),
      rows = srow, cols = PP_EMPTY_COL)
    conditionalFormatting(wb, "Summary", cols = PP_EMPTY_COL, rows = srow,
      rule = ">0",
      style = createStyle(fgFill = "#FFE0E0", fontColour = "#C00000",
                          textDecoration = "bold"))

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
         paste("PTO needed per PP is set by OFF+VAC days requested:",
               "5-6→1, 7-8→2, 9-10→3, 11-12→4, 13→5, 14→6.",
               "Target = 6 - CME - PTO. PTO is a count only — not pinned to specific days.")))

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

  # ── Hidden helper block for live shortfall ────────────────────────────────
  # Rows 1:N_PP       → static sched_target per person per PP
  # Rows (N_PP+1):2*N_PP → COUNTIFS actual shifts per person per PP from Schedule
  # These are referenced by the n_bump live formula in the Overview section.
  for (ci in seq_along(STAFF)) {
    person <- STAFF[ci]
    pc     <- person_pc(ci)
    hcn    <- HLPR_COL_START + ci - 1L
    for (k in seq_len(N_PP)) {
      writeData(wb, "Summary",
        x = as.integer(targets[[person]][[PAY_PERIODS$name[k]]]$sched_target),
        startRow = k, startCol = hcn, colNames = FALSE)
    }
    # Count this person's NAME in the slot columns D:G for the pay period -
    # NOT "Day"/"Night" in their own column. The person columns are formulas
    # that (in their Yellow branch) read THIS helper cell; counting them here
    # would close a reference loop and Excel would flag a circular reference.
    # The slot columns are plain input, so reading them is cycle-free.
    for (k in seq_len(N_PP)) {
      ppn <- PAY_PERIODS$name[k]
      one <- function(slot_col) sprintf(
        'COUNTIFS(Schedule!$C$2:$C$%1$d,"%2$s",Schedule!$%3$s$2:$%3$s$%1$d,"%4$s")',
        MAX_SCHED_ROW, ppn, slot_col, person)
      # Every slot column, MICU (D:G) and FC3 (H:J) alike - an FC3 shift counts
      # toward the pay-period total exactly like a MICU one.
      slot_cols <- vapply(MICU_COL_FIRST:FC3_COL_LAST, col_letter, character(1L))
      writeFormula(wb, "Summary",
        x = paste(vapply(slot_cols, one, character(1L)), collapse = "+"),
        startRow = N_PP + k, startCol = hcn)
    }
  }
  setColWidths(wb, "Summary",
    cols   = seq(HLPR_COL_START, HLPR_COL_START + N_STAFF - 1L),
    widths = rep(8, N_STAFF),
    hidden = TRUE)

  # ════════════════════════════════════════════════════════════════════════════
  # SHEET 3 · Schedule  (two rows per calendar day)
  # ════════════════════════════════════════════════════════════════════════════
  addWorksheet(wb, "Schedule")

  # Col layout:
  #   A Date | B Day | C PP | D E F  MICU day | G MICU night | H I J  FC3 |
  #   [staff...] | _key_
  # Row 1 is a group banner (MICU over D:G, FC3 over H:J); row 2 holds the real
  # column headers. Day and night are merged into one row per calendar day.
  #
  # FC3 is a separate service, staffed BY HAND - the solver never writes those
  # columns. A shift there counts exactly like any other shift for the person:
  # it shows in their column, their pay-period total and their weekend count.
  # What it does NOT do is create a hole: the understaffed markers, the Empty
  # shifts column and Requested Changes all stay on the MICU block (D:G), since
  # an empty FC3 slot is the normal case, not a gap.
  N_COLS <- N_HDR + N_STAFF + 1L

  # Row 1: group banner
  mergeCells(wb, "Schedule", cols = MICU_COL_FIRST:MICU_COL_LAST, rows = 1)
  writeData(wb, "Schedule", x = "MICU", startRow = 1, startCol = MICU_COL_FIRST,
            colNames = FALSE)
  addStyle(wb, "Schedule",
    mk(fg = C_NAVY, bold = TRUE, size = 11, font_color = F_WHITE),
    rows = 1, cols = MICU_COL_FIRST:MICU_COL_LAST)
  mergeCells(wb, "Schedule", cols = FC3_COL_FIRST:FC3_COL_LAST, rows = 1)
  writeData(wb, "Schedule", x = "FC3", startRow = 1, startCol = FC3_COL_FIRST,
            colNames = FALSE)
  addStyle(wb, "Schedule",
    mk(fg = F_FC3, bold = TRUE, size = 11, font_color = F_WHITE),
    rows = 1, cols = FC3_COL_FIRST:FC3_COL_LAST)
  addStyle(wb, "Schedule", mk(fg = C_NAVY), rows = 1,
    cols = setdiff(seq_len(N_COLS), MICU_COL_FIRST:FC3_COL_LAST))
  setRowHeights(wb, "Schedule", rows = 1, heights = 18)

  # Row 2: column headers
  hdr <- c("Date","Day","PP",
           # Slot identity is decided by hand, so all three MICU day columns
           # read "DAY"; only day-vs-night is a real distinction here.
           "DAY","DAY","DAY","NIGHT",
           "FC3_D","FC3_D","FC3_N",
           STAFF, "_key_")
  writeData(wb, "Schedule", x = as.data.frame(t(hdr)),
    startRow = SCHED_HDR_ROW, startCol = 1, colNames = FALSE)
  addStyle(wb, "Schedule",
    mk(fg = C_NAVY, bold = TRUE, font_color = F_WHITE,
       border = "Bottom", border_color = F_WHITE),
    rows = SCHED_HDR_ROW, cols = seq_along(hdr))
  setRowHeights(wb, "Schedule", rows = SCHED_HDR_ROW, heights = 20)
  freezePane(wb, "Schedule", firstActiveRow = SCHED_HDR_ROW + 1L)

  # ── Section rules around the shift blocks ──────────────────────────────────
  # The sheet reads as bands: Date/Day/PP on the left, WHO is on each shift in
  # D-J, and the per-person view on the right. A thick edge boxes MICU and FC3
  # separately. Applied after the per-cell styling below, since openxlsx styles
  # overwrite rather than merge - see the addStyle calls near the end.
  SHIFT_COL_FIRST <- MICU_COL_FIRST
  SHIFT_COL_LAST  <- FC3_COL_LAST

  schr  <- SCHED_HDR_ROW + 1L
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
                     stringsAsFactors = FALSE),
      startRow = dr, startCol = 1, colNames = FALSE)
    # Written statically rather than by reverse lookup: the person cells now all
    # read "Day", so MATCH("Day", ...) could not tell APP1 from APP2 from APP3.
    # The occupants are already known here, so no formula is needed.
    writeData(wb, "Schedule", x = data.frame(a = app1, b = app2, c = roam, d = night,
                                             stringsAsFactors = FALSE),
              startRow = dr, startCol = 4L, colNames = FALSE)
    # _key_: a real Date (not text) so other sheets can MATCH on a typed date
    # without locale-dependent TEXT() formatting. Used by "Requested Changes".
    writeData(wb, "Schedule",
      x = d, startRow = dr, startCol = N_COLS, colNames = FALSE)

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

    # Night slot col — MICU_COL_LAST (G), NOT N_HDR. N_HDR moved from 7 to 10
    # when the FC3 block was added, which silently sent this styling to the last
    # FC3 column and left Night unstyled.
    cbg_n <- if (nchar(night) > 0) (if (is_h) C_YELLOW else C_NIGHT) else bg_day
    addStyle(wb, "Schedule",
      mk(fg = cbg_n, bold = nchar(night) > 0,
         font_color = if (nchar(night) > 0) F_NAVY else F_GRAY,
         border = "All", border_color = "#DDDDDD", size = 9),
      rows = dr, cols = MICU_COL_LAST)

    # FC3 slot cols (H:J) — filled by hand, so they start empty. Ground tint
    # only; the fill when a name is typed comes from conditional formatting.
    for (cc in FC3_COL_FIRST:FC3_COL_LAST)
      addStyle(wb, "Schedule",
        mk(fg = bg_day, font_color = F_FC3, size = 9,
           border = "All", border_color = "#DDDDDD"),
        rows = dr, cols = cc)

    # ── Per-staff cols: LIVE formulas driven by the slot columns D-G ──────────
    # The slot columns are the editable input. Each person cell derives itself:
    #   name appears in D/E/F  -> "Day"
    #   name appears in G      -> "Night"
    #   otherwise              -> that person's time-off marker for the day
    #                             (OFF / CME / PTO / Yellow / blank), baked in
    #                             as a literal because time-off does not change
    #                             by editing the schedule.
    # So typing a name into an empty slot updates their column, and through it
    # the Calendar and every Summary count, without re-running anything.
    #
    # "Yellow" is shown only while it is actionable: the day is not yet fully
    # staffed (COUNTA of D:G < 4) AND the person is still short in that pay
    # period - read live from the Summary helper block, where row k holds the
    # target and row N_PP+k the running count.
    k_pp <- if (is.na(pp)) NA_integer_ else match(pp, PAY_PERIODS$name)
    for (ci in seq_along(STAFF)) {
      person <- STAFF[ci]
      col    <- N_HDR + ci
      pc     <- person_pc(ci)
      # Time-off-only role: what the cell shows when the person is NOT on a slot.
      off_role <- role_of(person, d, list(), time_off)
      tail <- if (identical(off_role, "Yellow") && !is.na(k_pp)) {
        hc <- col_letter(HLPR_COL_START + ci - 1L)
        sprintf('IF(AND(COUNTA($D%1$d:$G%1$d)<4,Summary!$%2$s$%3$d<Summary!$%2$s$%4$d),"Yellow","")',
                dr, hc, N_PP + k_pp, k_pp)
      } else if (nzchar(off_role) && !identical(off_role, "Yellow")) {
        sprintf('"%s"', off_role)
      } else {
        '""'
      }
      # MICU day (D:F) -> "Day"; MICU night (G) -> "Night";
      # FC3 day (H:I) -> "FC3"; FC3 night (J) -> "FC3 Night".
      # Header row is 2, so the name to match lives at <col>$2.
      #
      # NOTE: "FC3 Night" is a DISTINCT token from "Night". Every COUNTIF that
      # tallies shifts has to list it explicitly - COUNTIF matches whole cell
      # values, so "FC3 Night" is not caught by "FC3" or by "Night". That also
      # means it correctly stays OUT of the MICU night count.
      fml <- sprintf(
        paste0('IF(OR($D%1$d=%2$s$2,$E%1$d=%2$s$2,$F%1$d=%2$s$2),"Day",',
               'IF($G%1$d=%2$s$2,"Night",',
               'IF(OR($H%1$d=%2$s$2,$I%1$d=%2$s$2),"FC3",',
               'IF($J%1$d=%2$s$2,"FC3 Night",%3$s))))'),
        dr, pc, tail)
      writeFormula(wb, "Schedule", x = fml, startRow = dr, startCol = col)
      # Base style only (weekend/holiday ground + border). Role colours are
      # conditional formatting applied after the loop, since the value is live.
      addStyle(wb, "Schedule",
        mk(fg = bg_day, size = 9, halign = "center",
           border = "All", border_color = "#DDDDDD"),
        rows = dr, cols = col)
    }
    addStyle(wb, "Schedule",
      mk(fg = "#F7F7F7", font_color = F_LGRAY, size = 7, halign = "left"),
      rows = dr, cols = N_COLS)
    setRowHeights(wb, "Schedule", rows = dr, heights = 18)

    schr <- schr + 1L
  }

  # ── Slot columns D-G are the EDITABLE input ────────────────────────────────
  # Dropdown of staff names so a slot is picked, not typed (a typo would match
  # no one's column and silently vanish from every count).
  .last_row <- schr - 1L
  dataValidation(wb, "Schedule", cols = SHIFT_COL_FIRST:SHIFT_COL_LAST,
    rows = (SCHED_HDR_ROW + 1L):.last_row, type = "list", allowBlank = TRUE,
    value = paste0('"', paste(STAFF, collapse = ","), '"'))
  # Live fill: a name in a day slot reads green, in the night slot blue, so a
  # freshly typed name is coloured immediately. Empty slots keep the row ground
  # set above, which is what makes a hole visible.
  conditionalFormatting(wb, "Schedule",
    cols = MICU_COL_FIRST:(MICU_COL_LAST - 1L), rows = (SCHED_HDR_ROW + 1L):.last_row,
    type = "notBlanks",
    style = createStyle(fgFill = C_GREEN, fontColour = F_BLUE, textDecoration = "bold"))
  conditionalFormatting(wb, "Schedule",
    cols = MICU_COL_LAST, rows = (SCHED_HDR_ROW + 1L):.last_row,
    type = "notBlanks",
    style = createStyle(fgFill = C_NIGHT, fontColour = F_NAVY, textDecoration = "bold"))
  conditionalFormatting(wb, "Schedule",
    cols = FC3_COL_FIRST:FC3_COL_LAST, rows = (SCHED_HDR_ROW + 1L):.last_row,
    type = "notBlanks",
    style = createStyle(fgFill = C_FC3, fontColour = F_FC3, textDecoration = "bold"))
  # Holiday rows LAST so they win. openxlsx assigns conditional-format priority
  # in reverse order of addition - the rule added last gets priority 1, which is
  # the one Excel applies. Added before the fill rules above, these would have
  # been outranked and a staffed holiday would repaint green/blue.
  for (hd in HOLIDAY_DATES) {
    hr <- sched_row_map[[as.character(as.Date(hd, origin = "1970-01-01"))]]
    if (is.null(hr)) next
    conditionalFormatting(wb, "Schedule",
      cols = MICU_COL_FIRST:MICU_COL_LAST, rows = hr, type = "notBlanks",
      style = createStyle(fgFill = C_YELLOW, fontColour = F_GOLD,
                          textDecoration = "bold"))
  }

  # ── Per-person columns: role colours as conditional formatting ─────────────
  # The cell values are formulas now, so their colour must follow the computed
  # text. "contains" rules; none of these tokens is a substring of another.
  .pcols <- (N_HDR + 1L):(N_HDR + N_STAFF)
  .cf <- function(token, fill, font, bold = TRUE)
    conditionalFormatting(wb, "Schedule", cols = .pcols,
      rows = (SCHED_HDR_ROW + 1L):.last_row,
      type = "contains", rule = token,
      style = createStyle(fgFill = fill, fontColour = font,
                          textDecoration = if (bold) "bold" else NULL,
                          halign = "center"))
  .cf("Day",    C_GREEN,  F_BLUE)
  .cf("Night",  C_NIGHT,  F_NAVY)
  # AFTER "Night" on purpose: "FC3 Night" contains both tokens, and openxlsx
  # assigns conditional-format priority in reverse order of addition, so the
  # later rule wins. This keeps an FC3 night reading as FC3, not as a MICU night.
  .cf("FC3",    C_FC3,    F_FC3)
  .cf("OFF",    C_PINK,   F_RED)
  .cf("CME",    C_ORANGE, F_WHITE)
  .cf("PTO",    C_PTO,    F_RED)
  .cf("Yellow", C_PEACH,  "#8A6D00", bold = FALSE)

  # Day-before-night flag (a day shift immediately followed by a night): one
  # relative formula rule per column covers every row. Dashed orange border.
  for (ci in seq_along(STAFF)) {
    pc <- person_pc(ci)
    conditionalFormatting(wb, "Schedule", cols = N_HDR + ci,
      rows = (SCHED_HDR_ROW + 1L):(.last_row - 1L),
      rule = sprintf('AND($%1$s%2$d="Day",$%1$s%3$d="Night")', pc,
                     SCHED_HDR_ROW + 1L, SCHED_HDR_ROW + 2L),
      style = createStyle(border = "TopBottomLeftRight", borderColour = C_ORANGE,
                          borderStyle = "dashed"))
  }

  # ── Box the shift block (cols D-G) ─────────────────────────────────────────
  # Applied LAST and with stack = TRUE: openxlsx replaces a cell's style wholesale
  # otherwise, which would strip the fills and fonts set per cell above. Stacking
  # merges the border in and leaves the rest intact.
  for (edge in list(c(MICU_COL_FIRST, MICU_COL_LAST), c(FC3_COL_FIRST, FC3_COL_LAST))) {
    addStyle(wb, "Schedule",
      createStyle(border = "left", borderColour = "#000000", borderStyle = "medium"),
      rows = 1:.last_row, cols = edge[1], gridExpand = TRUE, stack = TRUE)
    addStyle(wb, "Schedule",
      createStyle(border = "right", borderColour = "#000000", borderStyle = "medium"),
      rows = 1:.last_row, cols = edge[2], gridExpand = TRUE, stack = TRUE)
    addStyle(wb, "Schedule",
      createStyle(border = "top", borderColour = "#000000", borderStyle = "medium"),
      rows = 1, cols = edge[1]:edge[2], gridExpand = TRUE, stack = TRUE)
    addStyle(wb, "Schedule",
      createStyle(border = "bottom", borderColour = "#000000", borderStyle = "medium"),
      rows = .last_row, cols = edge[1]:edge[2], gridExpand = TRUE, stack = TRUE)
  }

  setColWidths(wb, "Schedule",
    cols   = seq_len(N_COLS),
    widths = c(14, 7, 7, 16, 14, 16, 16, 14, 14, 14, rep(13, N_STAFF), 20))

  # ════════════════════════════════════════════════════════════════════════════
  # SHEET 4 · Requested Changes
  # ════════════════════════════════════════════════════════════════════════════
  # A change log keyed by date. Type a date in column A and PP# and Open Shifts
  # fill in by formula from the Schedule sheet - and stay LIVE, so once the
  # requested change is made on the Schedule sheet the row's Open Shifts drops
  # to "none" and it reads as resolved. Provider Add is a dropdown; Notes is
  # free text. Pre-seeded with every date that currently has an open mandatory
  # slot, followed by blank rows for ad-hoc entries.
  addWorksheet(wb, "Requested Changes")
  RC <- "Requested Changes"
  KL <- col_letter(N_COLS)                       # Schedule _key_ (date) column
  N_BLANK_ROWS <- 30L

  mergeCells(wb, RC, cols = 1:5, rows = 1)
  writeData(wb, RC, x = sprintf("Requested Changes · %s – %s",
                                format(SCHEDULE_START, "%b %d"),
                                format(SCHEDULE_END,   "%b %d, %Y")),
            startRow = 1, startCol = 1, colNames = FALSE)
  addStyle(wb, RC, mk(fg = C_NAVY, bold = TRUE, size = 13, font_color = F_WHITE,
                      halign = "left"), rows = 1, cols = 1:5)
  setRowHeights(wb, RC, rows = 1, heights = 27.75)

  rc_hdr <- c("Date", "PP#", "Open Shifts", "Provider Add", "Notes")
  writeData(wb, RC, x = as.data.frame(t(rc_hdr)), startRow = 2, startCol = 1,
            colNames = FALSE)
  addStyle(wb, RC, mk(fg = C_BLUE, bold = TRUE, font_color = F_WHITE,
                      border = "All", border_color = F_WHITE), rows = 2, cols = 1:5)
  freezePane(wb, RC, firstActiveRow = 3)

  # Dates that currently have an open mandatory slot (APP1 / APP2 / Night).
  seed_dates <- all_d[vapply(all_d, function(dd) {
    day_s <- sched_obj$schedule[[as.character(as.Date(dd, origin = "1970-01-01"))]]
    any(vapply(c("APP1", "APP2", "Night"), function(sl) {
      v <- day_s[[sl]]; length(v) != 1L || is.na(v)
    }, logical(1L)))
  }, logical(1L))]
  seed_dates <- as.Date(seed_dates, origin = "1970-01-01")

  rc_first <- 3L
  rc_last  <- rc_first + length(seed_dates) + N_BLANK_ROWS - 1L
  for (r in rc_first:rc_last) {
    i <- r - rc_first + 1L
    if (i <= length(seed_dates))
      writeData(wb, RC, x = seed_dates[i], startRow = r, startCol = 1, colNames = FALSE)
    # Row of the Schedule sheet holding this date (0 when not found).
    M  <- sprintf('MATCH($A%d,Schedule!$%s$1:$%s$%d,0)', r, KL, KL, MAX_SCHED_ROW)
    nd <- sprintf('(3-COUNTA(INDEX(Schedule!$D$1:$F$%d,%s,0)))', MAX_SCHED_ROW, M)
    nn <- sprintf('IF(INDEX(Schedule!$G$1:$G$%d,%s)="",1,0)', MAX_SCHED_ROW, M)
    writeFormula(wb, RC, startRow = r, startCol = 2,
      x = sprintf('IF($A%d="","",IFERROR(INDEX(Schedule!$C$1:$C$%d,%s),"not in schedule"))',
                  r, MAX_SCHED_ROW, M))
    writeFormula(wb, RC, startRow = r, startCol = 3,
      x = sprintf(paste0(
        'IF($A%1$d="","",IFERROR(IF(%2$s+%3$s=0,"none",',
        'IF(%2$s>0,%2$s&" day","")&IF(AND(%2$s>0,%3$s>0),", ","")&IF(%3$s>0,"night","")),',
        '"not in schedule"))'), r, nd, nn))
    addStyle(wb, RC, mk(halign = "center", border = "All", border_color = "#DDDDDD",
                        size = 10), rows = r, cols = 1:5)
    addStyle(wb, RC, createStyle(numFmt = "mm/dd/yy", halign = "center",
                                 border = "TopBottomLeftRight", borderColour = "#DDDDDD"),
             rows = r, cols = 1)
    addStyle(wb, RC, mk(halign = "left", border = "All", border_color = "#DDDDDD",
                        size = 10, wrap = TRUE), rows = r, cols = 5)
  }
  dataValidation(wb, RC, cols = 1, rows = rc_first:rc_last, type = "date",
    operator = "between", value = c(SCHEDULE_START, SCHEDULE_END))
  dataValidation(wb, RC, cols = 4, rows = rc_first:rc_last, type = "list",
    allowBlank = TRUE, value = paste0('"', paste(STAFF, collapse = ","), '"'))
  # Open Shifts: amber while something is open, green once it reads "none".
  conditionalFormatting(wb, RC, cols = 3, rows = rc_first:rc_last,
    type = "contains", rule = "none",
    style = createStyle(fgFill = C_GREEN, fontColour = F_NAVY))
  conditionalFormatting(wb, RC, cols = 3, rows = rc_first:rc_last,
    rule = sprintf('AND($C%1$d<>"",$C%1$d<>"none")', rc_first),
    style = createStyle(fgFill = "#FFF2CC", fontColour = "#7F6000",
                        textDecoration = "bold"))
  setColWidths(wb, RC, cols = 1:5, widths = c(12, 8, 16, 16, 48))

  # ── Force a full recalculation when Excel opens the file ──────────────────
  # openxlsx writes formulas with NO cached result and no calcChain. Excel will
  # usually recompute them, but conditional formatting that depends on those
  # results is not reliably re-evaluated - which left the Pay Period Detail
  # highlights not firing. fullCalcOnLoad makes Excel recompute everything
  # (values and conditional formats) the moment the workbook opens.
  wb$workbook$calcPr <- '<calcPr calcId="171027" fullCalcOnLoad="1"/>'

  # ── Save ───────────────────────────────────────────────────────────────────
  saveWorkbook(wb, output_path, overwrite = TRUE)
  message("  Saved: ", output_path)
  invisible(output_path)
}
