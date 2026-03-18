# ─────────────────────────────────────────────────────────────────────────────
# app.R  —  Shiny entry point
# Run with:  Rscript -e "shiny::runApp('.')"
#        or: shiny::runApp('/home/user/HMC_MICU_SCHEDULE')
# ─────────────────────────────────────────────────────────────────────────────

source("global.R")
source("ui.R")
source("server.R")

shinyApp(ui = ui, server = server)
