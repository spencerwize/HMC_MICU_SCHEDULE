source("global.R")

# res <- run_pipeline(
#   sheet      = "Oct26-Jan31",   # also sets date range + pay periods
#   time_limit = 3600*24,            # seconds PER SOLVE
#   mode       = "green"          # green-only phase
# )

# res_name <- paste0('res_',Sys.Date(),'.RDS')
# saveRDS(res, res_name)

res <- readRDS('res_2026-09-18.RDS')
build_excel(res$sched, res$time_off, res$targets, "MICU_Schedule_Final.xlsx")
