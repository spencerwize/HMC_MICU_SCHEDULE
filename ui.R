# ─────────────────────────────────────────────────────────────────────────────
# ui.R  —  Shiny UI  (bslib Bootstrap 5 layout)
# ─────────────────────────────────────────────────────────────────────────────

# Helper must be defined before page_navbar() evaluates the calls below
legend_chip <- function(color, label) {
  tags$div(class = "d-flex align-items-center mb-1",
    tags$div(style = sprintf(
      "width:14px;height:14px;background:%s;border:1px solid #ccc;
       border-radius:2px;margin-right:6px;flex-shrink:0;", color)),
    tags$small(label)
  )
}

ui <- page_navbar(
  title = tags$span(
    tags$img(src = "favicon.png", height = "20px",
             style = "margin-right:6px; vertical-align:middle;",
             onerror = "this.style.display='none'"),
    "HMC MICU APP Schedule"
  ),
  theme = bs_theme(
    bootswatch  = "flatly",
    primary     = "#203864",
    secondary   = "#2E75B6",
    base_font   = font_google("Inter"),
    heading_font = font_google("Inter")
  ),
  bg      = "#203864",
  inverse = TRUE,
  id      = "main_nav",

  # ── 1. Setup tab ──────────────────────────────────────────────────────────
  nav_panel(
    "Setup",
    icon = icon("gear"),
    layout_sidebar(
      sidebar = sidebar(
        width = 280,
        bg    = "#f5f7fa",
        tags$h6("Time-Off Data", class = "text-muted mt-2"),
        selectInput("sheet_select", "Sheet:",
          choices  = TIMEOFF_SHEETS,
          selected = TIMEOFF_SHEETS[1]
        ),
        tags$small(class = "text-muted",
          "Select the pay-period sheet from the Google Sheet."
        ),
        hr(),
        actionButton("run_btn", "Generate Schedule",
          icon  = icon("play-circle"),
          class = "btn-primary w-100",
          width = "100%"
        ),
        br(), br(),
        conditionalPanel(
          "output.schedule_ready",
          downloadButton("dl_excel", "Download Excel",
            class = "btn-success w-100"
          )
        )
      ),

      # Main content
      fluidRow(
        column(12,
          h4("About This App"),
          uiOutput("about_schedule_info"),
          tags$ul(
            tags$li(strong("Day shift:"), " 6:30 AM – 6:30 PM  (APP1, APP2, Roaming APP)"),
            tags$li(strong("Night shift:"), " 6:30 PM – 6:30 AM  (1 slot)"),
            tags$li("Target: 6 shifts per person per pay period"),
            tags$li("Fixed holiday assignments pre-seeded")
          ),
          hr(),
          fluidRow(
            column(4,
              card(class = "text-center",
                card_body(
                  h2(textOutput("stat_days"),   class = "text-primary mb-0"),
                  p("Schedule Days",             class = "text-muted small")
                )
              )
            ),
            column(4,
              card(class = "text-center",
                card_body(
                  h2(textOutput("stat_shifts"), class = "text-primary mb-0"),
                  p("Total Shifts Scheduled",    class = "text-muted small")
                )
              )
            ),
            column(4,
              card(class = "text-center",
                card_body(
                  h2(textOutput("stat_errors"), class = "text-danger mb-0"),
                  p("Hard Constraint Errors",    class = "text-muted small")
                )
              )
            )
          ),
          br(),
          conditionalPanel(
            "output.schedule_ready",
            card(
              card_header("Validation Results"),
              card_body(uiOutput("validation_ui"))
            )
          )
        )
      )
    )
  ),

  # ── 2. Calendar tab ───────────────────────────────────────────────────────
  nav_panel(
    "Calendar",
    icon = icon("calendar"),
    layout_sidebar(
      sidebar = sidebar(
        width = 220,
        bg = "#f5f7fa",
        tags$h6("Options", class = "text-muted mt-2"),
        pickerInput("cal_person", "Staff Member:",
          choices  = STAFF,
          selected = STAFF[1],
          options  = list(`live-search` = TRUE)
        ),
        selectInput("cal_month", "Month:",
          choices = c("April 2026" = "2026-04",
                      "May 2026"   = "2026-05",
                      "June 2026"  = "2026-06",
                      "July 2026"  = "2026-07"),
          selected = "2026-04"
        ),
        hr(),
        tags$small(class = "text-muted", "Color legend:"),
        tags$div(class = "mt-2",
          legend_chip("#92D050", "Day (APP1/APP2/Roam)"),
          legend_chip("#BDD7EE", "Night"),
          legend_chip("#FFFF99", "Holiday"),
          legend_chip("#FFD966", "Vacation"),
          legend_chip("#FF99CC", "PTO"),
          legend_chip("#FF6D01", "CME / Conference"),
          legend_chip("#FFC7CE", "Off"),
          legend_chip("#F2F2F2", "Weekend (no shift)")
        )
      ),
      conditionalPanel(
        "output.schedule_ready",
        uiOutput("calendar_ui")
      ),
      conditionalPanel(
        "!output.schedule_ready",
        div(class = "text-center text-muted mt-5",
          icon("calendar-times", style = "font-size:3em"),
          h5("Generate a schedule first on the Setup tab.")
        )
      )
    )
  ),

  # ── 3. Schedule Grid tab ──────────────────────────────────────────────────
  nav_panel(
    "Schedule Grid",
    icon = icon("table"),
    layout_sidebar(
      sidebar = sidebar(
        width = 220,
        bg = "#f5f7fa",
        tags$h6("Filters", class = "text-muted mt-2"),
        checkboxGroupInput("grid_pp", "Pay Period:",
          choices  = PAY_PERIODS$name,
          selected = PAY_PERIODS$name,
          inline   = FALSE
        ),
        hr(),
        checkboxInput("grid_show_off", "Show OFF/VAC/CME rows",  value = TRUE),
        checkboxInput("grid_compact",  "Compact (roles only)",   value = FALSE)
      ),
      conditionalPanel(
        "output.schedule_ready",
        card(
          card_header("Full Schedule Grid  — one row per day"),
          card_body(
            reactableOutput("schedule_table", height = "700px")
          )
        )
      ),
      conditionalPanel(
        "!output.schedule_ready",
        div(class = "text-center text-muted mt-5",
          icon("table", style = "font-size:3em"),
          h5("Generate a schedule first on the Setup tab.")
        )
      )
    )
  ),

  # ── 4. Summary & Analytics tab ────────────────────────────────────────────
  nav_panel(
    "Summary",
    icon = icon("chart-bar"),
    conditionalPanel(
      "output.schedule_ready",
      fluidRow(
        column(12,
          h4("Shift Distribution by Pay Period"),
          plotlyOutput("chart_pp_shifts", height = "350px")
        )
      ),
      fluidRow(
        column(6,
          h4("Night Shifts per Person"),
          plotlyOutput("chart_nights", height = "320px")
        ),
        column(6,
          h4("Roaming Shifts per Person"),
          plotlyOutput("chart_roaming", height = "320px")
        )
      ),
      fluidRow(
        column(12,
          h4("PP Target vs Actual"),
          reactableOutput("summary_table")
        )
      )
    ),
    conditionalPanel(
      "!output.schedule_ready",
      div(class = "text-center text-muted mt-5",
        icon("chart-bar", style = "font-size:3em"),
        h5("Generate a schedule first on the Setup tab.")
      )
    )
  ),

  # ── Footer ────────────────────────────────────────────────────────────────
  nav_spacer(),
  nav_item(
    tags$span(class = "navbar-text text-white-50 small",
      "HMC MICU  |  Apr–Jul 2026")
  ),

  # Enable shinyjs
  header = useShinyjs()
)

