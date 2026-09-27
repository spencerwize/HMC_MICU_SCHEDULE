# ─────────────────────────────────────────────────────────────────────────────
# server.R  —  Shiny Server
# ─────────────────────────────────────────────────────────────────────────────

server <- function(input, output, session) {

  # ── Populate sheet dropdown on startup ────────────────────────────────────
  # Runs once; isolate() prevents googlesheets4 auth internals from creating
  # a reactive dependency that would re-trigger this observer later.
  observe({
    sheet_names <- isolate(tryCatch({
      gs4_auth_auto()
      googlesheets4::sheet_names(TIMEOFF_GSHEET_URL)
    }, error = function(e) {
      message("Could not fetch sheet names from API — using default list.")
      TIMEOFF_SHEETS
    }))
    
    # setdiff (not x[-which(...)]) — the latter empties the vector when no
    # 'Rules' tab exists, because x[-integer(0)] selects nothing.
    sheet_names <- setdiff(sheet_names, c("Rules", "rules"))
    
    # Omit `selected` so the user's current choice (or the ui.R default) is kept
    updateSelectInput(session, "sheet_select", choices = sheet_names)
  })

  # ── Update UI controls when the sheet selection changes ────────────────────
  # Keeps the calendar month picker and PP checkboxes in sync with the chosen
  # date range without requiring the user to regenerate first.
  observeEvent(input$sheet_select, {
    cfg <- SHEET_CONFIGS[[input$sheet_select]]
    if (!is.null(cfg)) {
      updateSelectInput(session, "cal_month",
        choices  = cfg$cal_months,
        selected = cfg$cal_months[1]
      )
      updateCheckboxGroupInput(session, "grid_pp",
        choices  = cfg$pay_periods$name,
        selected = cfg$pay_periods$name
      )
    }
  }, ignoreInit = TRUE)

  # ── Prior schedule grid ────────────────────────────────────────────────────
  # Renders a 7-day × nPerson grid of select inputs (—/Day/Night).
  # Re-renders when the sheet changes so the date labels stay correct.
  output$prior_schedule_ui <- renderUI({
    cfg <- SHEET_CONFIGS[[input$sheet_select]]
    if (is.null(cfg)) return(tags$p(class = "text-muted small", "Select a sheet first."))

    prior_dates  <- seq(cfg$schedule_start - 7L, cfg$schedule_start - 1L, by = "day")
    date_labels  <- format(prior_dates, "%b %d (%a)")

    make_id <- function(person, d)
      sprintf("prior_%s_%s", gsub("[^A-Za-z0-9]", "_", person), format(d, "%Y%m%d"))

    header_row <- tags$tr(
      tags$th(style = "min-width:90px;", ""),
      lapply(date_labels, function(lbl)
        tags$th(class = "text-center", style = "min-width:80px; font-size:0.75rem;",
                HTML(gsub(" \\(", "<br/>(", lbl))))
    )
    person_rows <- lapply(STAFF, function(person) {
      cells <- lapply(prior_dates, function(d) {
        tags$td(class = "p-1",
          selectInput(make_id(person, d), label = NULL, width = "75px",
            choices  = c("—" = "", "Day" = "day", "Night" = "night"),
            selected = "")
        )
      })
      tags$tr(
        tags$td(class = "align-middle fw-semibold pe-2",
                style = "font-size:0.82rem; white-space:nowrap;", person),
        cells
      )
    })

    tags$table(
      class = "table table-sm table-bordered align-middle mb-0",
      style = "font-size:0.8rem;",
      tags$thead(class = "table-light", header_row),
      tags$tbody(person_rows)
    )
  })

  # ── Schedule result store ──────────────────────────────────────────────────
  # Using reactiveVal + observeEvent (not eventReactive) so the result is only
  # ever set by an explicit button click; it cannot be re-triggered by reactive
  # invalidation from googlesheets4 auth internals or any other side-effect.
  pipeline <- reactiveVal(NULL)

  # Manual picks made in the holes worklist. Empty data.frame, same shape as
  # pinned_df(), so it concatenates directly onto the green-only assignments.
  manual_picks <- reactiveVal(
    data.frame(person = character(), date = as.Date(character()),
               slot = character(), stringsAsFactors = FALSE))

  # ── Shared setup for both phases ──────────────────────────────────────────
  # Applies the selected sheet's constants, parses time-off, computes targets
  # and reads the prior-schedule grid. Used by the green-only phase; the fill
  # phase reuses the scheduler object the green phase produced.
  prepare_inputs <- function(sheet_key) {
    cfg <- SHEET_CONFIGS[[sheet_key]]
    if (!is.null(cfg)) {
      SCHEDULE_START <<- cfg$schedule_start
      SCHEDULE_END   <<- cfg$schedule_end
      PAY_PERIODS    <<- cfg$pay_periods
      HOLIDAYS       <<- cfg$holidays
      HOLIDAY_DATES  <<- if (!is.null(cfg$holiday_dates)) cfg$holiday_dates
                         else as.Date(names(cfg$holidays))
      HOLIDAY_NAMES  <<- cfg$holiday_names
    }
    selected_sheet <- if (nzchar(sheet_key)) sheet_key else NULL
    time_off <- parse_time_off(TIMEOFF_GSHEET_URL, sheet = selected_sheet)
    targets  <- compute_targets(time_off)
    green_supply_report(time_off, targets)

    ps_start    <- if (!is.null(cfg)) cfg$schedule_start else SCHEDULE_START
    prior_dates <- seq(ps_start - 7L, ps_start - 1L, by = "day")
    make_id     <- function(person, d)
      sprintf("prior_%s_%s", gsub("[^A-Za-z0-9]", "_", person), format(d, "%Y%m%d"))
    ps_list <- setNames(lapply(STAFF, function(person) {
      rows <- Filter(Negate(is.null), lapply(prior_dates, function(d) {
        val <- input[[make_id(person, d)]]
        if (!is.null(val) && nzchar(val))
          data.frame(date = d, type = val, stringsAsFactors = FALSE)
      }))
      if (length(rows) > 0L) do.call(rbind, rows) else NULL
    }), STAFF)
    prior_schedule <- if (any(vapply(ps_list, function(df) !is.null(df), logical(1L))))
      ps_list else NULL

    list(time_off = time_off, targets = targets, prior_schedule = prior_schedule)
  }

  # Package a solved scheduler into the reactive the rest of the UI reads.
  store_result <- function(sched, time_off, targets, phase) {
    validation <- validate_schedule(sched, time_off, targets, partial = (phase == 1L))
    updatePickerInput(session, "cal_person", choices = STAFF, selected = STAFF[1])
    pipeline(list(
      sched      = sched,
      time_off   = time_off,
      targets    = targets,
      validation = validation,
      tier_used  = sched$tier_used,
      phase      = phase,
      holes      = sched$holes_df(),
      green      = green_summary(sched, time_off, targets, quiet = TRUE),
      pins_kept  = sched$pins_kept,
      df         = sched$to_dataframe(),
      grid       = sched$to_person_grid(time_off, targets)
    ))
  }

  # ── Phase 1: build the green-only schedule ────────────────────────────────
  observeEvent(input$run_btn, {
    shinyjs::disable("run_btn"); shinyjs::disable("fill_btn")
    on.exit({ shinyjs::enable("run_btn"); shinyjs::enable("fill_btn") }, add = TRUE)

    withProgress(message = "Building the green-only schedule…", value = 0, {
      setProgress(0.1, detail = "Parsing requests…")
      inp <- prepare_inputs(input$sheet_select)

      setProgress(0.35, detail = "Solving on requested-work days only…")
      sched <- SchedulerLP$new(inp$time_off, inp$targets,
                               prior_schedule = inp$prior_schedule)
      sched$run_green()

      setProgress(0.95, detail = "Validating…")
      store_result(sched, inp$time_off, inp$targets, phase = 1L)
      setProgress(1.0, detail = "Done.")
    })
  })

  # ── Phase 2: fill the remainder ───────────────────────────────────────────
  observeEvent(input$fill_btn, {
    p <- pipeline()
    if (is.null(p) || is.null(p$sched)) {
      showNotification("Build the green schedule first.", type = "warning")
      return(invisible(NULL))
    }
    shinyjs::disable("run_btn"); shinyjs::disable("fill_btn")
    on.exit({ shinyjs::enable("run_btn"); shinyjs::enable("fill_btn") }, add = TRUE)

    withProgress(message = "Filling the remainder…", value = 0, {
      setProgress(0.2, detail = "Holding the requested-work assignments…")
      sched <- p$sched
      # Any manual worklist picks are held alongside the green-only assignments.
      pins <- rbind(sched$pinned_df(), manual_picks())
      pins <- pins[!duplicated(paste(pins$person, pins$date, pins$slot)), , drop = FALSE]

      setProgress(0.4, detail = "Solving…")
      sched$run_fill(pinned = pins)

      setProgress(0.95, detail = "Validating…")
      store_result(sched, p$time_off, p$targets, phase = 2L)
      setProgress(1.0, detail = "Done.")
    })
  })

  # ── About text — reflects selected sheet's date range ─────────────────────
  output$about_schedule_info <- renderUI({
    cfg <- SHEET_CONFIGS[[input$sheet_select]]
    if (is.null(cfg)) {
      p("Select a sheet and click Generate Schedule.")
    } else {
      pp_names <- cfg$pay_periods$name
      p(sprintf(
        "This tool builds a 12-hour rotating shift schedule for 10 APP staff covering %s – %s (%s–%s).",
        format(cfg$schedule_start, "%B %d, %Y"),
        format(cfg$schedule_end,   "%B %d, %Y"),
        pp_names[1], pp_names[length(pp_names)]
      ))
    }
  })

  # Flag for conditionalPanel
  output$schedule_ready <- reactive({ !is.null(pipeline()) })
  outputOptions(output, "schedule_ready", suspendWhenHidden = FALSE)

  # ── Setup tab: stat cards ──────────────────────────────────────────────────
  output$stat_days <- renderText({
    req(pipeline())
    as.character(length(pipeline()$sched$dates))
  })

  output$stat_shifts <- renderText({
    req(pipeline())
    p  <- pipeline()
    n  <- sum(sapply(STAFF, function(x) {
      nrow(p$sched$person_shifts[[x]]) + length(p$sched$person_nights[[x]])
    }))
    as.character(n)
  })

  output$stat_nights <- renderUI({
    req(pipeline())
    p  <- pipeline()
    ct <- vapply(STAFF, function(x) length(p$sched$person_nights[[x]]), integer(1L))
    mx <- max(ct); mn <- min(ct)
    mxp <- STAFF[which.max(ct)]; mnp <- STAFF[which.min(ct)]
    tagList(
      tags$p(class = "text-muted small mb-2 fw-semibold", "Night Shifts"),
      tags$div(class = "d-flex justify-content-between",
        tags$span("Max:"), tags$span(sprintf("%d  (%s)", mx, mxp), class = "text-primary fw-bold")),
      tags$div(class = "d-flex justify-content-between",
        tags$span("Min:"), tags$span(sprintf("%d  (%s)", mn, mnp), class = "text-primary fw-bold"))
    )
  })

  output$stat_weekends <- renderUI({
    req(pipeline())
    p      <- pipeline()
    all_d  <- as.Date(p$sched$dates, origin = "1970-01-01")
    # Weekend = Friday night through Sunday night (is_weekend_shift), matching
    # the ILP. Night shifts live in person_nights, not person_shifts.
    ct <- vapply(STAFF, function(x) {
      sh  <- p$sched$person_shifts[[x]]
      nts <- p$sched$person_nights[[x]]
      as.integer(
        (if (nrow(sh)) sum(is_weekend_shift(sh$date, sh$slot)) else 0L) +
        (if (length(nts)) sum(is_weekend_shift(nts, "Night"))  else 0L))
    }, integer(1L))
    mx <- max(ct); mn <- min(ct)
    mxp <- STAFF[which.max(ct)]; mnp <- STAFF[which.min(ct)]
    tagList(
      tags$p(class = "text-muted small mb-2 fw-semibold", "Weekend Shifts"),
      tags$div(class = "d-flex justify-content-between",
        tags$span("Max:"), tags$span(sprintf("%d  (%s)", mx, mxp), class = "text-primary fw-bold")),
      tags$div(class = "d-flex justify-content-between",
        tags$span("Min:"), tags$span(sprintf("%d  (%s)", mn, mnp), class = "text-primary fw-bold"))
    )
  })

  output$stat_coverage <- renderUI({
    req(pipeline())
    p      <- pipeline()
    all_ds <- as.character(as.Date(p$sched$dates, origin = "1970-01-01"))
    sched  <- p$sched$schedule
    empty  <- function(x) is.null(x) || is.na(x) || x == ""
    unstaffed <- sum(vapply(all_ds, function(d) empty(sched[[d]]$Night),   logical(1L)))
    no_roam   <- sum(vapply(all_ds, function(d) empty(sched[[d]]$Roaming), logical(1L)))
    tagList(
      tags$p(class = "text-muted small mb-2 fw-semibold", "Coverage Gaps"),
      tags$div(class = "d-flex justify-content-between",
        tags$span("Unstaffed nights:"),
        tags$span(as.character(unstaffed),
          class = if (unstaffed == 0L) "text-success fw-bold" else "text-danger fw-bold")),
      tags$div(class = "d-flex justify-content-between",
        tags$span("No APP3 days:"),
        tags$span(as.character(no_roam),
          class = if (no_roam == 0L) "text-success fw-bold" else "text-warning fw-bold"))
    )
  })

  output$validation_ui <- renderUI({
    req(pipeline())
    warns <- pipeline()$validation$warnings
    if (length(warns) == 0) return(NULL)
    tags$div(class = "alert alert-warning",
      tags$strong(sprintf("%d warning(s):", length(warns))),
      tags$ul(lapply(warns[seq_len(min(20, length(warns)))], tags$li)),
      if (length(warns) > 20)
        tags$li(sprintf("… and %d more", length(warns) - 20))
    )
  })

  # ── Download Excel ────────────────────────────────────────────────────────
  output$dl_excel <- downloadHandler(
    filename = function() {
      paste0("MICU_APP_Schedule_", format(Sys.Date(), "%Y%m%d"), ".xlsx")
    },
    content = function(file) {
      p <- pipeline()
      build_excel(p$sched, p$time_off, p$targets, file)
    }
  )

  # ── Calendar tab ──────────────────────────────────────────────────────────
  output$calendar_ui <- renderUI({
    req(pipeline())
    person <- input$cal_person
    ym     <- input$cal_month

    p      <- pipeline()
    year   <- as.integer(substr(ym, 1, 4))
    month  <- as.integer(substr(ym, 6, 7))

    first_d <- as.Date(sprintf("%d-%02d-01", year, month))
    if (month == 12L) {
      last_d <- as.Date(sprintf("%d-01-01", year + 1L)) - 1L
    } else {
      last_d <- as.Date(sprintf("%d-%02d-01", year, month + 1L)) - 1L
    }

    start_dow <- as.integer(format(first_d, "%w"))  # 0=Sun
    dow_labels <- c("Sun","Mon","Tue","Wed","Thu","Fri","Sat")

    # Build grid cells
    n_cells <- start_dow + as.integer(last_d - first_d) + 1L
    n_rows  <- ceiling(n_cells / 7L)

    grid_cells <- vector("list", n_rows * 7L)
    for (i in seq_len(n_rows * 7L)) grid_cells[[i]] <- tags$td(style = "background:#f8f8f8;")

    idx  <- start_dow + 1L
    cur  <- first_d
    while (cur <= last_d) {
      in_sched <- (cur >= SCHEDULE_START && cur <= SCHEDULE_END)

      role  <- ""
      bg    <- "#FFFFFF"
      color <- "#000"

      if (in_sched) {
        # Role logic lives in R/roles.R — shared with the Schedule Grid and
        # the Excel export so the three views cannot drift apart again.
        role <- role_of(person, cur, p$sched$schedule, p$time_off,
                        p$sched$granted_pto)

        is_hol <- cur %in% HOLIDAY_DATES
        bg <- switch(role,
          Day     = if (is_hol) "#FFFF99" else "#92D050",
          Night   = if (is_hol) "#FFFF99" else "#BDD7EE",
          CME     = "#FF6D01",
          OFF     = "#FFC7CE",
          PTO     = "#FF99CC",
          Yellow  = "#FFD966",
          if (is_weekend(cur)) "#F2F2F2" else "#FFFFFF"
        )
        color <- if (role == "CME") "#FFFFFF" else "#000000"
        # A requested-work day is marked with an OUTLINE, not a fill: the fill is
        # already carrying the role, and #92D050/#FFFF99 are taken by day-shift
        # and holiday. Solid when the request was granted, dashed when it was not.
        # Outline ONLY a shift worked on a day the person marked Yellow. Green is
        # the default state for most days now (blank counts as Green), so
        # outlining every green shift would mark nearly the whole calendar; a
        # shift landing on a Yellow day is the exception worth flagging.
        worked <- role %in% WORK_ROLES
        border <- if (worked && !is_green_day(person, cur, p$time_off))
                    "2px solid #B8860B"
                  else "1px solid #ddd"
      } else {
        bg <- "#EEEEEE"
        border <- "1px solid #ddd"
      }

      day_num <- as.integer(format(cur, "%d"))
      role_lbl <- if (nzchar(role)) tags$div(
        style = "font-size:10px; font-weight:bold; margin-top:2px;", role
      ) else NULL

      grid_cells[[idx]] <- tags$td(
        style = sprintf(
          "background:%s; color:%s; padding:6px 4px; text-align:center;
           border:%s; min-width:60px; height:56px;
           vertical-align:top; font-size:13px;", bg, color, border),
        tags$div(style = "font-weight:600;", day_num),
        role_lbl
      )
      idx <- idx + 1L
      cur <- cur + 1L
    }

    # Build table rows
    trs <- lapply(seq_len(n_rows), function(r) {
      start <- (r - 1L) * 7L + 1L
      cells <- grid_cells[start:(start + 6L)]
      tags$tr(cells)
    })

    header_tr <- tags$tr(
      lapply(dow_labels, function(dow)
        tags$th(dow, style = "background:#2E75B6; color:white;
                 text-align:center; padding:6px; width:60px;"))
    )

    card(
      card_header(
        sprintf("%s — %s",
                format(first_d, "%B %Y"),
                person)
      ),
      card_body(
        tags$table(
          class = "table table-bordered mb-0",
          style = "border-collapse:collapse; width:100%;",
          tags$thead(header_tr),
          tags$tbody(trs)
        )
      )
    )
  })

  # ── Schedule Grid tab ──────────────────────────────────────────────────────
  output$schedule_table <- renderReactable({
    req(pipeline())
    p    <- pipeline()
    grid <- p$grid

    # Filter by selected PPs
    grid <- grid[grid$pp %in% input$grid_pp, ]

    # Pivot wide: date x person
    role_colors <- c(
      Day     = "#92D050",
      Night   = "#BDD7EE",
      CME     = "#FF6D01",
      OFF     = "#FFC7CE",
      PTO     = "#FF99CC",
      Yellow  = "#FFD966"      # marked "avoid if possible"
    )

    # `wants` rides alongside `role` so a WORKED green day can keep its role fill
    # and still be outlined. Encoded into the cell value as a trailing marker,
    # then stripped for display — reactable colDefs see one value per cell.
    grid$role_mark <- ifelse(grid$wants & nzchar(grid$role),
                             paste0(grid$role, "*"), grid$role)
    wide <- grid %>%
      select(date, day_name, pp, person, role_mark, is_holiday, is_weekend) %>%
      tidyr::pivot_wider(
        id_cols     = c(date, day_name, pp, is_holiday, is_weekend),
        names_from  = person,
        values_from = role_mark
      ) %>%
      arrange(date)

    # Flag days where no one is assigned to APP 3
    staff_present <- STAFF[STAFF %in% names(wide)]
    wide$app3_open <- apply(
      wide[, staff_present, drop = FALSE], 1,
      function(row) sum(sub("[*]$", "", row) == "Day", na.rm = TRUE) < 3L
    )

    # Make cell colour helper
    show_off   <- isTRUE(input$grid_show_off)
    show_green <- isTRUE(input$grid_show_green)
    # "APP1*" -> role "APP1" plus a requested-work marker.
    base_role  <- function(v) sub("[*]$", "", v)
    is_wanted  <- function(v) grepl("[*]$", v)
    make_col <- function(person_name) {
      colDef(
        name   = person_name,
        width  = 72,
        style  = function(value) {
          if (is.null(value) || is.na(value) || !nzchar(value))
            return(list(background = "#FAFAFA"))
          r <- base_role(value); w <- is_wanted(value)
          if (!show_off   && r %in% c("OFF", "VAC", "CME"))
            return(list(background = "#FAFAFA"))
          if (!show_green && r == "Yellow")
            return(list(background = "#FAFAFA"))
          bg <- role_colors[r]
          if (is.na(bg)) bg <- "#FAFAFA"
          st <- list(background = bg, fontWeight = "bold",
                     fontSize = "11px", textAlign = "center")
          # Outline ONLY a shift worked on a Yellow day - the exception worth
          # seeing. Green is now the default state for most days, so outlining
          # every green shift would mark almost the whole grid.
          if (!w && r %in% WORK_ROLES)
            st$boxShadow <- "inset 0 0 0 2px #B8860B"
          st
        },
        cell   = function(value) {
          if (is.null(value) || is.na(value)) return("")
          r <- base_role(value)
          if (!show_off   && r %in% c("OFF", "VAC", "CME")) return("")
          if (!show_green && r == "Yellow") return("")
          r
        }
      )
    }

    person_cols <- setNames(lapply(STAFF, make_col), STAFF)

    app3_col <- colDef(
      name  = "APP3",
      width = 46,
      style = function(value) {
        if (isTRUE(value))
          list(background = "#FCE4D6", textAlign = "center")
        else
          list(background = "#E2EFDA", textAlign = "center")
      },
      cell = function(value) if (isTRUE(value)) "\u2205" else "\u2713"
    )

    date_col <- colDef(
      name = "Date",
      width = 90,
      style = function(value) list(fontWeight = "bold"),
      cell  = function(value) format(as.Date(value), "%m/%d")
    )

    reactable(
      wide,
      columns = c(
        list(
          date      = date_col,
          day_name  = colDef(name = "Day",  width = 40),
          pp        = colDef(name = "PP",   width = 45),
          app3_open = app3_col,
          is_holiday = colDef(show = FALSE),
          is_weekend = colDef(show = FALSE)
        ),
        person_cols
      ),
      rowStyle = function(index) {
        row <- wide[index, ]
        if (isTRUE(row$is_holiday)) return(list(border = "2px solid #FFA500"))
        if (isTRUE(row$is_weekend)) return(list(background = "#F5F5F5"))
        list()
      },
      striped         = FALSE,
      highlight       = TRUE,
      bordered        = TRUE,
      compact         = input$grid_compact,
      searchable      = FALSE,
      pagination      = FALSE,
      defaultPageSize = nrow(wide),
      height          = 700,
      theme = reactableTheme(
        headerStyle = list(background = "#2E75B6", color = "white",
                           fontWeight = "bold")
      )
    )
  })

  # ── Summary charts ────────────────────────────────────────────────────────
  output$chart_pp_shifts <- renderPlotly({
    req(pipeline())
    p <- pipeline()

    df <- do.call(rbind, lapply(STAFF, function(person) {
      lapply(PAY_PERIODS$name, function(pp) {
        data.frame(
          person   = person,
          pp       = pp,
          actual   = p$sched$pp_counts[[person]][[pp]],
          target   = p$targets[[person]][[pp]]$sched_target,
          stringsAsFactors = FALSE
        )
      })
    })) %>% bind_rows()

    plot_ly(df, x = ~pp, y = ~actual, color = ~person,
            type = "bar", text = ~actual, textposition = "inside") %>%
      layout(
        barmode = "group",
        xaxis   = list(title = "Pay Period"),
        yaxis   = list(title = "Shifts Scheduled", range = c(0, 8)),
        legend  = list(orientation = "h", x = 0, y = -0.25),
        shapes  = list(
          list(type = "line", x0 = -0.5, x1 = length(PAY_PERIODS$name) - 0.5,
               y0 = 6, y1 = 6,
               line = list(color = "red", width = 1.5, dash = "dot"))
        )
      ) %>%
      config(displayModeBar = FALSE)
  })

  output$chart_nights <- renderPlotly({
    req(pipeline())
    p <- pipeline()
    df <- data.frame(
      person = STAFF,
      nights = sapply(STAFF, function(x) length(p$sched$person_nights[[x]])),
      stringsAsFactors = FALSE
    )
    plot_ly(df, x = ~person, y = ~nights, type = "bar",
            marker = list(color = "#BDD7EE",
                          line = list(color = "#2E75B6", width = 1.5))) %>%
      layout(
        xaxis = list(title = "", tickangle = -30),
        yaxis = list(title = "Night Shifts"),
        showlegend = FALSE
      ) %>%
      config(displayModeBar = FALSE)
  })

  output$chart_roaming <- renderPlotly({
    req(pipeline())
    p         <- pipeline()
    all_dates <- as.Date(p$sched$dates, origin = "1970-01-01")
    df <- data.frame(
      person  = STAFF,
      weekend = sapply(STAFF, function(x) {
        sh  <- p$sched$person_shifts[[x]]
        nts <- p$sched$person_nights[[x]]
        (if (nrow(sh)) sum(is_weekend_shift(sh$date, sh$slot)) else 0L) +
        (if (length(nts)) sum(is_weekend_shift(nts, "Night"))  else 0L)
      }),
      stringsAsFactors = FALSE
    )
    plot_ly(df, x = ~person, y = ~weekend, type = "bar",
            marker = list(color = "#2E75B6",
                          line = list(color = "#1A4D8C", width = 1.5))) %>%
      layout(
        xaxis = list(title = "", tickangle = -30),
        yaxis = list(title = "Weekend Shifts",
                     range = list(0, max(20, max(df$weekend) + 1))),
        showlegend = FALSE,
        shapes = list(
          list(type = "line", x0 = -0.5, x1 = nrow(df) - 0.5,
               y0 = MIN_WKND_HARD, y1 = MIN_WKND_HARD,
               line = list(color = "red", dash = "dot", width = 1.5)),
          list(type = "line", x0 = -0.5, x1 = nrow(df) - 0.5,
               y0 = MAX_WKND_HARD, y1 = MAX_WKND_HARD,
               line = list(color = "red", dash = "dot", width = 1.5))
        )
      ) %>%
      config(displayModeBar = FALSE)
  })

  output$summary_table <- renderReactable({
    req(pipeline())
    p   <- pipeline()
    tdf <- targets_summary_df(p$targets)

    # Join in actual counts
    actual_df <- do.call(rbind, lapply(STAFF, function(person) {
      lapply(PAY_PERIODS$name, function(pp) {
        data.frame(
          person = person,
          pp     = pp,
          actual = p$sched$pp_counts[[person]][[pp]],
          stringsAsFactors = FALSE
        )
      })
    })) %>% bind_rows()

    # Compute PTO granted per person-PP from sched$granted_pto
    pto_df <- do.call(rbind, lapply(STAFF, function(person) {
      pto_dates <- p$sched$granted_pto[[person]]
      lapply(PAY_PERIODS$name, function(pp) {
        pp_row <- PAY_PERIODS[PAY_PERIODS$name == pp, ]
        n <- if (length(pto_dates) == 0L) 0L else
          sum(pto_dates >= pp_row$start & pto_dates <= pp_row$end)
        data.frame(person = person, pp = pp, pto_granted = n,
                   stringsAsFactors = FALSE)
      })
    })) %>% bind_rows()

    df <- left_join(tdf, actual_df, by = c("person", "pp")) %>%
      left_join(pto_df, by = c("person", "pp")) %>%
      mutate(status = case_when(
        actual < soft_min     ~ "Below minimum",
        actual < sched_target ~ "Under",
        actual > sched_target ~ "Over",
        TRUE                  ~ "On target"
      ))

    reactable(df,
      columns = list(
        person       = colDef(name = "Person",       width = 90),
        pp           = colDef(name = "PP",           width = 55),
        avail        = colDef(name = "Avail Days",   width = 85),
        credited     = colDef(name = "CME Credited", width = 100),
        target       = colDef(name = "Target",       width = 70),
        sched_target = colDef(name = "Sched Target", width = 100),
        soft_min     = colDef(name = "Min Floor",    width = 80),
        pto_granted  = colDef(name = "PTO Credited", width = 100,
          cell = function(v) if (!is.na(v) && v > 0L) as.character(v) else "—"),
        actual       = colDef(name = "Actual",       width = 70),
        status       = colDef(name = "Status",       width = 115,
          style = function(value) {
            list(
              color = switch(value,
                "Below minimum" = "#990000",
                "Under"         = "#CC6600",
                "Over"          = "#0066CC",
                "On target"     = "#009900",
                "#000"),
              fontWeight = "bold"
            )
          })
      ),
      groupBy    = "person",
      striped    = TRUE,
      highlight  = TRUE,
      bordered   = TRUE,
      searchable = TRUE,
      theme      = reactableTheme(
        headerStyle = list(background = "#2E75B6", color = "white",
                           fontWeight = "bold")
      )
    )
  })

  # ── Count viable schedules (no-good cut enumeration) ──────────────────────
  count_sol_result <- reactiveVal(NULL)

  observeEvent(input$count_sol_btn, {
    req(pipeline())
    sched <- pipeline()$sched
    lim   <- as.integer(input$count_sol_limit)
    count_sol_result(NULL)
    withProgress(message = sprintf("Counting schedules (up to %d)…", lim), value = 0.1, {
      n <- sched$count_solutions(max_count = lim)
    })
    count_sol_result(list(n = n, lim = lim))
  })

  output$count_sol_ui <- renderUI({
    res <- count_sol_result()
    if (is.null(res)) return(NULL)
    if (res$n >= res$lim) {
      msg <- sprintf(
        "Found at least %d distinct feasible schedules (limit reached — there may be more).",
        res$n)
      cls <- "alert alert-info mt-2"
    } else {
      msg <- sprintf("Found exactly %d distinct feasible schedule(s) satisfying all active constraints.", res$n)
      cls <- "alert alert-success mt-2"
    }
    tags$div(class = cls, tags$strong(msg))
  })
}
