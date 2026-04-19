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

  # ── Schedule result store ──────────────────────────────────────────────────
  # Using reactiveVal + observeEvent (not eventReactive) so the result is only
  # ever set by an explicit button click; it cannot be re-triggered by reactive
  # invalidation from googlesheets4 auth internals or any other side-effect.
  pipeline <- reactiveVal(NULL)

  observeEvent(input$run_btn, {
    shinyjs::disable("run_btn")
    on.exit(shinyjs::enable("run_btn"), add = TRUE)

    # Apply sheet-specific constants before any pipeline step so that
    # SCHEDULE_START / SCHEDULE_END / PAY_PERIODS / HOLIDAYS reflect the
    # currently selected sheet rather than the April–July defaults.
    cfg <- SHEET_CONFIGS[[input$sheet_select]]
    if (!is.null(cfg)) {
      SCHEDULE_START <<- cfg$schedule_start
      SCHEDULE_END   <<- cfg$schedule_end
      PAY_PERIODS    <<- cfg$pay_periods
      HOLIDAYS       <<- cfg$holidays
      # holiday_dates may be explicitly narrower than names(HOLIDAYS) when
      # some entries are pre-seeded weekend shifts that shouldn't be highlighted
      HOLIDAY_DATES  <<- if (!is.null(cfg$holiday_dates)) cfg$holiday_dates
                         else as.Date(names(cfg$holidays))
      HOLIDAY_NAMES  <<- cfg$holiday_names
    }

    withProgress(message = "Building schedule…", value = 0, {
      setProgress(0.1, detail = "Parsing time-off data…")
      selected_sheet <- if (nzchar(input$sheet_select)) input$sheet_select else NULL
      time_off <- parse_time_off(TIMEOFF_GSHEET_URL, sheet = selected_sheet)

      setProgress(0.25, detail = "Computing targets…")
      targets <- compute_targets(time_off)

      setProgress(0.35, detail = "Building and solving schedule (ILP)…")
      sched <- SchedulerLP$new(time_off, targets)
      sched$run(n_candidates = as.integer(input$n_candidates))

      setProgress(0.95, detail = "Validating…")
      validation <- validate_schedule(sched, time_off, targets)

      setProgress(1.0, detail = "Done.")

      # Refresh the cal_person picker to match whoever is in the sheet
      updatePickerInput(session, "cal_person",
        choices  = STAFF,
        selected = STAFF[1]
      )

      pipeline(list(
        sched      = sched,
        time_off   = time_off,
        targets    = targets,
        validation = validation,
        tier_used  = sched$tier_used,
        df         = sched$to_dataframe(),
        grid       = sched$to_person_grid(time_off, targets)
      ))
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

  output$stat_errors <- renderText({
    req(pipeline())
    as.character(length(pipeline()$validation$errors))
  })

  output$stat_tier <- renderUI({
    req(pipeline())
    tu  <- pipeline()$tier_used
    idx <- if (is.null(tu)) NA_integer_ else tu$index
    lbl <- if (is.null(tu)) "—" else sprintf("%d / 18", idx)
    cls <- if (is.na(idx))    "text-secondary"
           else if (idx <= 2) "text-success"
           else if (idx <= 5) "text-warning"
           else               "text-danger"
    tagList(
      h2(lbl, class = paste(cls, "mb-0")),
      if (!is.null(tu))
        tags$small(class = "text-muted", tu$label)
    )
  })

  output$validation_ui <- renderUI({
    req(pipeline())
    v <- pipeline()$validation
    errs  <- v$errors
    warns <- v$warnings

    err_ui <- if (length(errs) == 0) {
      tags$div(class = "alert alert-success",
        icon("check-circle"), " No hard constraint violations.")
    } else {
      tags$div(class = "alert alert-danger",
        tags$strong(sprintf("%d constraint error(s):", length(errs))),
        tags$ul(lapply(errs, tags$li))
      )
    }

    warn_ui <- if (length(warns) == 0) NULL else {
      tags$div(class = "alert alert-warning mt-2",
        tags$strong(sprintf("%d warning(s):", length(warns))),
        tags$ul(lapply(warns[seq_len(min(20, length(warns)))], tags$li)),
        if (length(warns) > 20)
          tags$li(sprintf("… and %d more", length(warns) - 20))
      )
    }
    tagList(err_ui, warn_ui)
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
      ds      <- as.character(cur)
      in_sched <- (cur >= SCHEDULE_START && cur <= SCHEDULE_END)

      role  <- ""
      bg    <- "#FFFFFF"
      color <- "#000"

      if (in_sched) {
        # Look up role
        day_s <- p$sched$schedule[[ds]]
        for (s in SLOTS) {
          v <- day_s[[s]]
          if (!is.na(v) && v == person) {
            role <- if (s == "Night") "Night" else
                    if (s == "APP1")  "APP1"  else
                    if (s == "APP2")  "APP2"  else "APP 3"
            break
          }
        }
        if (role == "") {
          pdata <- p$time_off[[person]]
          m     <- pdata[pdata$date == cur, ]
          typ   <- if (nrow(m) > 0) m$type[1] else NA_character_
          if (!is.na(typ)) {
            pp_now  <- get_pp(cur)
            pp_info <- p$targets[[person]][[pp_now]]
            role <- switch(typ,
              cme = "CME",
              off = "OFF",
              vac = "VAC",
              ""
            )
          }
        }

        is_hol <- cur %in% HOLIDAY_DATES
        bg <- switch(role,
          APP1    = if (is_hol) "#FFFF99" else "#92D050",
          APP2    = if (is_hol) "#FFFF99" else "#92D050",
          "APP 3" = if (is_hol) "#FFFF99" else "#92D050",
          Night   = if (is_hol) "#FFFF99" else "#BDD7EE",
          VAC     = "#FFD966",
          CME     = "#FF6D01",
          OFF     = "#FFC7CE",
          if (is_weekend(cur)) "#F2F2F2" else "#FFFFFF"
        )
        color <- if (role == "CME") "#FFFFFF" else "#000000"
      } else {
        bg <- "#EEEEEE"
      }

      day_num <- as.integer(format(cur, "%d"))
      role_lbl <- if (nzchar(role)) tags$div(
        style = "font-size:10px; font-weight:bold; margin-top:2px;", role
      ) else NULL

      grid_cells[[idx]] <- tags$td(
        style = sprintf(
          "background:%s; color:%s; padding:6px 4px; text-align:center;
           border:1px solid #ddd; min-width:60px; height:56px;
           vertical-align:top; font-size:13px;", bg, color),
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
      APP1    = "#92D050", APP2 = "#92D050", "APP 3" = "#92D050",
      Night   = "#BDD7EE",
      VAC     = "#FFD966", CME  = "#FF6D01",
      OFF     = "#FFC7CE"
    )

    wide <- grid %>%
      select(date, day_name, pp, person, role, is_holiday, is_weekend) %>%
      tidyr::pivot_wider(
        id_cols     = c(date, day_name, pp, is_holiday, is_weekend),
        names_from  = person,
        values_from = role
      ) %>%
      arrange(date)

    # Make cell colour helper
    make_col <- function(person_name) {
      colDef(
        name   = person_name,
        width  = 72,
        style  = function(value) {
          if (is.null(value) || is.na(value) || !nzchar(value))
            return(list(background = "#FAFAFA"))
          bg <- role_colors[value]
          if (is.na(bg)) bg <- "#FAFAFA"
          list(background = bg, fontWeight = "bold",
               fontSize = "11px", textAlign = "center")
        },
        cell   = function(value) {
          if (is.null(value) || is.na(value)) "" else value
        }
      )
    }

    person_cols <- setNames(lapply(STAFF, make_col), STAFF)

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
          date     = date_col,
          day_name = colDef(name = "Day", width = 40),
          pp       = colDef(name = "PP",  width = 45),
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
    p <- pipeline()
    df <- data.frame(
      person = STAFF,
      roam   = sapply(STAFF, function(x)
        sum(p$sched$person_shifts[[x]]$slot == "Roaming")),
      stringsAsFactors = FALSE
    )
    plot_ly(df, x = ~person, y = ~roam, type = "bar",
            marker = list(color = "#92D050",
                          line = list(color = "#5A9E2F", width = 1.5))) %>%
      layout(
        xaxis = list(title = "", tickangle = -30),
        yaxis = list(title = "Roaming Shifts"),
        showlegend = FALSE
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

    df <- left_join(tdf, actual_df, by = c("person", "pp")) %>%
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
