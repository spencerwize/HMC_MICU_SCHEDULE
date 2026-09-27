# ─────────────────────────────────────────────────────────────────────────────
# roles.R  —  Single source of truth for the per-person, per-day role string.
#
# This logic used to be duplicated (and had drifted) in three places:
#   R/scheduler_lp.R  to_person_grid()   — Schedule Grid tab
#   server.R          calendar renderer  — Calendar tab
#   R/excel_output.R  person_role()      — Excel export
# All three now delegate here.
# ─────────────────────────────────────────────────────────────────────────────

# Role vocabulary:
#   "Day" | "Night"                      - assigned to a day or night shift
#   "CME"                                — requested CME/conference day (blocked)
#   "OFF"                                — requested day off or vacation (blocked)
#   "Yellow"                             - marked Yellow: avoid assigning if possible
#   ""                                   - marked Green but not assigned that day
#
# `granted_pto` is vestigial (never populated — see scheduler_lp.R populate step);
# the argument is retained so the three call sites stay behaviour-identical.
role_of <- function(person, d, schedule, time_off, granted_pto = NULL) {
  day_s <- schedule[[as.character(d)]]
  if (!is.null(day_s)) {
    for (s in SLOTS) {
      v <- day_s[[s]]
      if (length(v) == 1L && !is.na(v) && v == person)
        # All three day slots (APP1 / APP2 / Roaming) render as "Day". The slot
        # DISTINCTION still exists in the model and in the Schedule sheet's own
        # APP1/APP2/APP3 columns - this is purely how a person's own cell reads.
        return(if (s == "Night") "Night" else "Day")
    }
  }
  if (!is.null(granted_pto) && d %in% granted_pto[[person]]) return("PTO")

  pdata <- time_off[[person]]
  if (is.null(pdata) || nrow(pdata) == 0L) return("")
  m <- pdata[pdata$date == d, ]
  if (nrow(m) == 0L) return("")
  switch(m$type[1],
         cme    = "CME",
         off    = "OFF",
         vac    = "OFF",
         pto    = "PTO",      # requested paid time off - blocked, credited
         yellow = "Yellow",   # schedulable, but avoid if possible
         green  = "",         # available and not scheduled - deliberately blank
         "")
}

# TRUE when `person` marked `d` as a requested WORK day (green).
# Kept separate from role_of() because a *worked* green day keeps its slot role
# ("APP1", "Night", …) — the green request is an orthogonal attribute that the
# display layers render as an outline on top of the role fill.
is_green_day <- function(person, d, time_off) {
  pdata <- time_off[[person]]
  if (is.null(pdata) || nrow(pdata) == 0L) return(FALSE)
  any(pdata$date == d & pdata$type == "green")
}

# The four scheduling slots a role string can name, in display order.
WORK_ROLES <- c("Day", "Night")
