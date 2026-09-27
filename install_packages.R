# ─────────────────────────────────────────────────────────────────────────────
# install_packages.R  —  Run once to install all required R packages
# ─────────────────────────────────────────────────────────────────────────────

pkgs <- c(
  # Core logic
  "R6",          # Reference classes for SchedulerLP
  "openxlsx",    # Excel output with rich formatting
  "dplyr",       # Data manipulation
  "tidyr",       # Data reshaping

  # ILP scheduler (SchedulerLP)
  "highs",       # HiGHS MIP solver (fast, modern, open-source)
  "Matrix",      # sparse matrix for HiGHS constraint matrix

  # Request-sheet input
  "googlesheets4", # Reads the Red/Yellow/Green request sheet (public; no auth)

  # Shiny app
  "shiny",       # Web app framework
  "bslib",       # Bootstrap 5 themes + layouts
  "reactable",   # Interactive tables with per-cell styling
  "plotly",      # Interactive charts
  "shinyWidgets",# Enhanced input widgets (pickerInput, etc.)
  "htmltools",   # HTML helpers
  "shinyjs"      # JavaScript helpers (show/hide, etc.)
)

missing <- pkgs[!pkgs %in% rownames(installed.packages())]

if (length(missing) == 0) {
  cat("All packages already installed.\n")
} else {
  cat("Installing:", paste(missing, collapse = ", "), "\n")
  install.packages(missing, repos = "https://cran.rstudio.com/")
  cat("Done.\n")
}
