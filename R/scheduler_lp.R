# ─────────────────────────────────────────────────────────────────────────────
# scheduler_lp.R  —  Integer Linear Program (ILP) based MICU APP scheduler
#
# WHY ILP?
# ────────
# The greedy MRV heuristic (scheduler.R) makes locally-optimal decisions one
# slot at a time, then applies repair passes to clean up soft-constraint
# violations.  An ILP encodes every scheduling rule as a linear constraint and
# lets a branch-and-bound solver find the globally best assignment in one shot:
#
#   • Provably optimal (or best solution within the time budget)
#   • Every constraint is satisfied simultaneously — no repair passes needed
#   • Soft preferences (day-before-night avoidance, shift fairness) are part of
#     the objective function, not afterthoughts
#   • Naturally handles all pay-period interactions at once
#
# VARIABLE LAYOUT
# ───────────────
# Primary (binary):
#   x[p, d, s]   1 if person p is assigned to slot s on schedule day d
#   Flat 1-based index:   (p-1)*nD*4 + (d-1)*4 + s
#   Slot order:           APP1=1, APP2=2, Roaming=3, Night=4
#
# Auxiliary (continuous [0,1]):
#   z[p, d, s]   linearised indicator for day-before-night:
#                z = x[p,d,s_day] × x[p,d+1,Night]
#   Flat 1-based index (offset by nX = nP*nD*4):
#                nX + (p-1)*(nD-1)*3 + (d-1)*3 + (s_day-1) + 1
#
# CONSTRAINTS
# ───────────
#   C1  Slot uniqueness:   Σ_p x[p,d,s]    ≤ 1           for all d,s
#   C2  APP1 coverage:     Σ_p x[p,d,APP1] = 1           for all d
#   C3  Night coverage:    Σ_p x[p,d,Night] = 1          for all d
#   C4  No double-book:    Σ_s x[p,d,s]    ≤ 1           for all p,d
#   C5  Availability:      upper bound = 0 for blocked (p,d) pairs
#   C6  PP shift cap:      Σ_{d∈PP,s} x[p,d,s] ≤ sched_target[p,PP]
#   C7  Night→Day 1d:      x[p,d,Night] + x[p,d+1,s] ≤ 1   s∈DAY_SLOTS
#   C8  Night recovery 2d: x[p,d,Night] + x[p,d+2,s] ≤ 1   s∈DAY_SLOTS
#   C9  Max 3 consec Nts:  Σ_{k=0}^3 x[p,d+k,Night]    ≤ 3
#   C10 Max 4 consec work: Σ_{k=0}^4 Σ_s x[p,d+k,s]   ≤ 4
#   C11 Night total cap:   Σ_d x[p,d,Night] ≤ MAX_NIGHTS_TOTAL
#   C12 Holiday pre-seed:  lb = ub = 1 for pre-assigned (p,d,s)
#   C13 DBN lower bound:   z[p,d,s] ≥ x[p,d,s] + x[p,d+1,Night] - 1
#
# OBJECTIVE (maximise)
# ────────────────────
#   Σ_{p,d,s}  w[s] × x[p,d,s]   − 8 × Σ_{p,d,s_day} z[p,d,s]
#   where w[APP1]=4, w[Night]=4, w[APP2]=2, w[Roaming]=2
#
# FALLBACK
# ────────
# If lpSolveAPI is not installed, the solver times out without a solution, or
# the ILP is declared infeasible, SchedulerLP automatically falls back to the
# greedy Scheduler and copies its state so the rest of the app is unaffected.
# ─────────────────────────────────────────────────────────────────────────────

SchedulerLP <- R6::R6Class("SchedulerLP",

  public = list(

    # ── Fields (mirror Scheduler public interface) ────────────────────────────
    time_off      = NULL,   # person -> data.frame(date, type)
    targets       = NULL,   # person -> pp -> list(...)
    dates         = NULL,   # Date vector: all schedule dates

    schedule      = NULL,   # "YYYY-MM-DD" -> list(APP1,APP2,Roaming,Night)
    person_nights = NULL,   # person -> Date vector
    person_shifts = NULL,   # person -> data.frame(date, slot)
    pp_counts     = NULL,   # person -> named int vector (pp -> count)
    granted_pto   = NULL,   # person -> Date vector (empty for ILP path)

    # ── Constructor ───────────────────────────────────────────────────────────
    initialize = function(time_off, targets) {
      self$time_off    <- time_off
      self$targets     <- targets
      self$dates       <- all_dates()
      self$granted_pto <- setNames(
        lapply(STAFF, function(p) as.Date(character())), STAFF)

      empty_slot <- list(APP1    = NA_character_, APP2    = NA_character_,
                         Roaming = NA_character_, Night   = NA_character_)
      self$schedule <- setNames(
        lapply(self$dates, function(d) empty_slot),
        as.character(self$dates))

      self$person_nights <- setNames(
        lapply(STAFF, function(p) as.Date(character())), STAFF)

      self$person_shifts <- setNames(
        lapply(STAFF, function(p)
          data.frame(date = as.Date(character()), slot = character(),
                     stringsAsFactors = FALSE)),
        STAFF)

      zero_pp <- setNames(integer(nrow(PAY_PERIODS)), PAY_PERIODS$name)
      self$pp_counts <- setNames(
        lapply(STAFF, function(p) zero_pp), STAFF)
    },

    # ── Main entry point ──────────────────────────────────────────────────────
    run = function() {
      if (!requireNamespace("lpSolveAPI", quietly = TRUE)) {
        message("lpSolveAPI not installed — falling back to greedy scheduler.")
        message("  Install with:  install.packages('lpSolveAPI')")
        return(private$run_greedy_fallback())
      }

      message("  Building ILP model...")
      result <- private$build_and_solve()

      if (is.null(result)) {
        message("  ILP returned no solution — falling back to greedy scheduler.")
        return(private$run_greedy_fallback())
      }

      message("  Translating ILP solution to schedule...")
      private$populate_from_solution(result)
      message("  Done.")
      invisible(self)
    },

    # ── Export (identical to Scheduler) ───────────────────────────────────────
    to_dataframe = function() {
      df <- do.call(rbind, lapply(self$dates, function(d) {
        ds  <- as.character(d)
        day <- self$schedule[[ds]]
        data.frame(
          date     = d,
          day_name = weekdays(d, abbreviate = TRUE),
          pp       = get_pp(d),
          APP1     = ifelse(is.na(day$APP1),    "", day$APP1),
          APP2     = ifelse(is.na(day$APP2),    "", day$APP2),
          Roaming  = ifelse(is.na(day$Roaming), "", day$Roaming),
          Night    = ifelse(is.na(day$Night),   "", day$Night),
          stringsAsFactors = FALSE
        )
      }))
      rownames(df) <- NULL
      df
    },

    to_person_grid = function(time_off, targets) {
      rows <- lapply(self$dates, function(d) {
        ds    <- as.character(d)
        day_s <- self$schedule[[ds]]
        pp    <- get_pp(d)
        lapply(STAFF, function(person) {
          role <- NA_character_
          for (s in SLOTS) {
            v <- day_s[[s]]
            if (!is.na(v) && v == person) {
              role <- if (s == "Night")   "Night"
                      else if (s == "APP1") "APP1"
                      else if (s == "APP2") "APP2"
                      else                  "APP 3"
              break
            }
          }
          if (is.na(role)) {
            pdata <- time_off[[person]]
            m     <- pdata[pdata$date == d, ]
            if (nrow(m) > 0)
              role <- switch(m$type[1],
                cme = "CME", off = "OFF", vac = "VAC", NA_character_)
          }
          data.frame(
            date       = d,
            day_name   = weekdays(d, abbreviate = TRUE),
            pp         = pp,
            person     = person,
            role       = ifelse(is.na(role), "", role),
            is_holiday = d %in% HOLIDAY_DATES,
            is_weekend = is_weekend(d),
            stringsAsFactors = FALSE
          )
        })
      })
      do.call(rbind, unlist(rows, recursive = FALSE))
    }
  ),

  # ── Private implementation ─────────────────────────────────────────────────
  private = list(

    # ── Build and solve the ILP ────────────────────────────────────────────────
    build_and_solve = function() {

      requireNamespace("lpSolveAPI", quietly = TRUE)

      dates_vec <- as.Date(self$dates, origin = "1970-01-01")
      nP  <- length(STAFF)
      nD  <- length(dates_vec)
      nS  <- 4L            # slots: APP1, APP2, Roaming, Night
      nDS <- 3L            # day slots: APP1=1, APP2=2, Roaming=3

      # Slot constants (1-based)
      S_APP1    <- 1L
      S_APP2    <- 2L
      S_ROAM    <- 3L
      S_NIGHT   <- 4L
      DAY_S     <- c(S_APP1, S_APP2, S_ROAM)

      # ── Variable index helpers ───────────────────────────────────────────────
      # Primary x[p, d, s]  binary; s ∈ 1..4
      nX  <- nP * nD * nS
      xidx <- function(p, d, s) (p - 1L) * nD * nS + (d - 1L) * nS + s

      # Auxiliary z[p, d, s_day]  continuous [0,1]; s_day ∈ 1..3 (day slots)
      # Represents day-before-night occurrence: x[p,d,s_day] × x[p,d+1,Night]
      nZ  <- nP * (nD - 1L) * nDS
      nV  <- nX + nZ
      zidx <- function(p, d, s_day)
        nX + (p - 1L) * (nD - 1L) * nDS + (d - 1L) * nDS + s_day

      # ── Create LP model ──────────────────────────────────────────────────────
      model <- lpSolveAPI::make.lp(nrow = 0L, ncol = nV)
      lpSolveAPI::lp.control(model,
        sense   = "max",
        timeout = 300L,      # 5-minute wall-clock limit
        epsel   = 1e-6,
        epsb    = 1e-6,
        epsd    = 1e-6,
        presolve = "rows"
      )

      # ── Objective ────────────────────────────────────────────────────────────
      # w[APP1]=4, w[Night]=4, w[APP2]=2, w[Roaming]=2; z penalty = −8
      obj <- numeric(nV)
      for (p in seq_len(nP)) {
        for (d in seq_len(nD)) {
          obj[xidx(p, d, S_APP1)]  <- 4
          obj[xidx(p, d, S_APP2)]  <- 2
          obj[xidx(p, d, S_ROAM)]  <- 2
          obj[xidx(p, d, S_NIGHT)] <- 4
        }
        for (d in seq_len(nD - 1L)) {
          for (sd in seq_len(nDS))
            obj[zidx(p, d, sd)] <- -8
        }
      }
      lpSolveAPI::set.objfn(model, obj)

      # ── C1: Slot uniqueness — Σ_p x[p,d,s] ≤ 1 ──────────────────────────────
      for (d in seq_len(nD)) {
        for (s in seq_len(nS)) {
          cols <- vapply(seq_len(nP), function(p) xidx(p, d, s), integer(1L))
          lpSolveAPI::add.constraint(model,
            xt = rep(1, nP), type = "<=", rhs = 1, indices = cols)
        }
      }

      # ── C2: APP1 must always be filled — Σ_p x[p,d,APP1] = 1 ────────────────
      for (d in seq_len(nD)) {
        cols <- vapply(seq_len(nP), function(p) xidx(p, d, S_APP1), integer(1L))
        lpSolveAPI::add.constraint(model,
          xt = rep(1, nP), type = "=", rhs = 1, indices = cols)
      }

      # ── C3: Night must always be filled — Σ_p x[p,d,Night] = 1 ──────────────
      for (d in seq_len(nD)) {
        cols <- vapply(seq_len(nP), function(p) xidx(p, d, S_NIGHT), integer(1L))
        lpSolveAPI::add.constraint(model,
          xt = rep(1, nP), type = "=", rhs = 1, indices = cols)
      }

      # ── C4: No double-booking — Σ_s x[p,d,s] ≤ 1 ────────────────────────────
      for (p in seq_len(nP)) {
        for (d in seq_len(nD)) {
          cols <- vapply(seq_len(nS), function(s) xidx(p, d, s), integer(1L))
          lpSolveAPI::add.constraint(model,
            xt = rep(1, nS), type = "<=", rhs = 1, indices = cols)
        }
      }

      # ── C5: Availability — upper-bound blocked (p,d) variables to 0 ──────────
      # VAC days block scheduling (person cannot work), same as off/cme.
      for (pi in seq_len(nP)) {
        person <- STAFF[pi]
        pdata  <- self$time_off[[person]]
        for (di in seq_len(nD)) {
          d <- dates_vec[di]
          blocked <- nrow(pdata) > 0 &&
            any(pdata$date == d & pdata$type %in% c("off", "vac", "cme"))
          if (blocked) {
            for (s in seq_len(nS))
              lpSolveAPI::set.bounds(model,
                upper = 0, columns = xidx(pi, di, s))
          }
        }
      }

      # ── C6: PP shift cap — Σ_{d∈PP,s} x[p,d,s] ≤ sched_target[p,PP] ─────────
      for (pi in seq_len(nP)) {
        person <- STAFF[pi]
        for (ppi in seq_len(nrow(PAY_PERIODS))) {
          pp_name <- PAY_PERIODS$name[ppi]
          pp_d    <- seq(PAY_PERIODS$start[ppi], PAY_PERIODS$end[ppi], by = "day")
          di_pp   <- which(dates_vec %in% pp_d)
          cap     <- self$targets[[person]][[pp_name]]$sched_target
          if (length(di_pp) == 0 || cap <= 0) next
          cols <- as.integer(unlist(lapply(di_pp, function(di)
            vapply(seq_len(nS), function(s) xidx(pi, di, s), integer(1L)))))
          lpSolveAPI::add.constraint(model,
            xt = rep(1, length(cols)), type = "<=", rhs = cap, indices = cols)
        }
      }

      # ── C7: Night-then-day ban (1 day) — x[p,d,Night]+x[p,d+1,s] ≤ 1 ────────
      for (pi in seq_len(nP)) {
        for (di in seq_len(nD - 1L)) {
          ni <- xidx(pi, di, S_NIGHT)
          for (s in DAY_S) {
            lpSolveAPI::add.constraint(model,
              xt = c(1, 1), type = "<=", rhs = 1,
              indices = c(ni, xidx(pi, di + 1L, s)))
          }
        }
      }

      # ── C8: Night recovery (2 days) — x[p,d,Night]+x[p,d+2,s] ≤ 1 ───────────
      for (pi in seq_len(nP)) {
        for (di in seq_len(nD - 2L)) {
          ni <- xidx(pi, di, S_NIGHT)
          for (s in DAY_S) {
            lpSolveAPI::add.constraint(model,
              xt = c(1, 1), type = "<=", rhs = 1,
              indices = c(ni, xidx(pi, di + 2L, s)))
          }
        }
      }

      # ── C9: Max 3 consecutive nights — Σ_{k=0}^3 x[p,d+k,Night] ≤ 3 ─────────
      for (pi in seq_len(nP)) {
        for (di in seq_len(nD - 3L)) {
          cols <- vapply(0:3, function(k) xidx(pi, di + k, S_NIGHT), integer(1L))
          lpSolveAPI::add.constraint(model,
            xt = rep(1, 4L), type = "<=", rhs = 3, indices = cols)
        }
      }

      # ── C10: Max 4 consecutive working days — Σ_{k=0}^4 Σ_s x[p,d+k,s] ≤ 4 ──
      for (pi in seq_len(nP)) {
        for (di in seq_len(nD - 4L)) {
          cols <- as.integer(unlist(lapply(0:4, function(k)
            vapply(seq_len(nS), function(s) xidx(pi, di + k, s), integer(1L)))))
          lpSolveAPI::add.constraint(model,
            xt = rep(1, 20L), type = "<=", rhs = 4, indices = cols)
        }
      }

      # ── C11: Total nights per person ≤ MAX_NIGHTS_TOTAL ──────────────────────
      for (pi in seq_len(nP)) {
        cols <- vapply(seq_len(nD), function(di)
          xidx(pi, di, S_NIGHT), integer(1L))
        lpSolveAPI::add.constraint(model,
          xt = rep(1, nD), type = "<=", rhs = MAX_NIGHTS_TOTAL, indices = cols)
      }

      # ── C12: Holiday pre-seeds — fix x[p,d,s] = 1 ────────────────────────────
      # Skip pre-seed if the person is marked off or vac that day (Rule 1).
      slot_idx <- c(APP1 = S_APP1, APP2 = S_APP2, Roaming = S_ROAM, Night = S_NIGHT)
      for (ds in names(HOLIDAYS)) {
        d_hol <- as.Date(ds)
        if (d_hol < SCHEDULE_START || d_hol > SCHEDULE_END) next
        di <- which(dates_vec == d_hol)
        if (length(di) == 0L) next
        for (slot_name in names(HOLIDAYS[[ds]])) {
          person <- HOLIDAYS[[ds]][[slot_name]]
          pi     <- which(STAFF == person)
          si     <- slot_idx[[slot_name]]
          if (length(pi) == 0L || is.null(si) || is.na(si)) next
          pdata  <- self$time_off[[person]]
          if (nrow(pdata) > 0 &&
              any(pdata$date == d_hol & pdata$type %in% c("off", "vac"))) next
          lpSolveAPI::set.bounds(model,
            lower = 1, upper = 1, columns = xidx(pi[[1L]], di[[1L]], si))
        }
      }

      # ── C13: Day-before-night lower bound — z[p,d,s] ≥ x[p,d,s]+x[p,d+1,Nt]-1
      # Combined with z ≥ 0 and z ≤ 1 (from bounds), this forces z = 1 exactly
      # when both x[p,d,s_day]=1 and x[p,d+1,Night]=1, and z = 0 otherwise.
      for (pi in seq_len(nP)) {
        for (di in seq_len(nD - 1L)) {
          ni <- xidx(pi, di + 1L, S_NIGHT)
          for (sd in seq_len(nDS)) {
            zi <- zidx(pi, di, sd)
            xi <- xidx(pi, di, sd)
            # z ≥ x + x_night - 1
            lpSolveAPI::add.constraint(model,
              xt = c(-1, 1, 1), type = ">=", rhs = -1L,
              indices = c(zi, xi, ni))
          }
        }
      }

      # ── Set variable types ───────────────────────────────────────────────────
      # x variables are binary; z variables are continuous [0,1]
      lpSolveAPI::set.type(model, columns = seq_len(nX), type = "binary")
      # z upper bound = 1 (lower = 0 is the default)
      lpSolveAPI::set.bounds(model,
        lower = rep(0, nZ), upper = rep(1, nZ),
        columns = nX + seq_len(nZ))

      # ── Report model size and solve ──────────────────────────────────────────
      message(sprintf(
        "  ILP: %d vars (%d binary, %d continuous)",
        nV, nX, nZ))

      status <- solve(model)

      # 0 = OPTIMAL, 1 = SUBOPTIMAL (timeout with feasible sol), others = failure
      if (status == 0L) {
        message("  Solver: OPTIMAL solution found.")
      } else if (status == 1L) {
        message("  Solver: SUBOPTIMAL — returning best solution found within timeout.")
      } else {
        message(sprintf("  Solver status %d (2=infeasible, 3=unbounded, 5=error).",
                        status))
        return(NULL)
      }

      sol <- lpSolveAPI::get.variables(model)
      obj_val <- lpSolveAPI::get.objective(model)
      message(sprintf("  Objective value: %.1f", obj_val))

      list(sol       = sol,
           nP        = nP,
           nD        = nD,
           nS        = nS,
           xidx      = xidx,
           dates_vec = dates_vec)
    },

    # ── Translate binary solution vector into schedule data structures ──────────
    populate_from_solution = function(res) {
      sol       <- res$sol
      nP        <- res$nP
      nD        <- res$nD
      xidx      <- res$xidx
      dates_vec <- res$dates_vec

      SLOT_NAMES <- c("APP1", "APP2", "Roaming", "Night")

      for (pi in seq_len(nP)) {
        person <- STAFF[pi]
        for (di in seq_len(nD)) {
          d  <- dates_vec[di]
          ds <- as.character(d)
          pp <- get_pp(d)
          for (si in 1:4) {
            vi <- xidx(pi, di, si)
            if (!is.na(sol[vi]) && round(sol[vi]) == 1L) {
              slot <- SLOT_NAMES[si]
              self$schedule[[ds]][[slot]] <- person
              if (!is.na(pp))
                self$pp_counts[[person]][[pp]] <-
                  self$pp_counts[[person]][[pp]] + 1L
              if (slot == "Night") {
                self$person_nights[[person]] <-
                  sort(c(self$person_nights[[person]], d))
              } else {
                self$person_shifts[[person]] <- rbind(
                  self$person_shifts[[person]],
                  data.frame(date = d, slot = slot, stringsAsFactors = FALSE))
              }
            }
          }
        }
      }

      # Diagnostic summary (mirrors parse_time_off style)
      message("Schedule summary (LP solution):")
      for (p in STAFF) {
        ns <- length(self$person_nights[[p]])
        ds <- nrow(self$person_shifts[[p]])
        message(sprintf("  %-10s  day=%d  night=%d  total=%d",
                        p, ds, ns, ds + ns))
      }
    },

    # ── Greedy fallback (copies Scheduler state into self) ──────────────────────
    run_greedy_fallback = function() {
      message("  Pre-seeding holidays...")
      gs <- Scheduler$new(self$time_off, self$targets)
      gs$preseed_holidays()
      message("  Scheduling all slots (MRV hardest-first + lookahead)...")
      gs$schedule_all()
      message("  Cleanup pass...")
      gs$cleanup_pass()
      message("  Swap repair pass...")
      gs$swap_repair_pass()
      message("  Force-filling APP1...")
      gs$force_fill_app1()

      self$schedule      <- gs$schedule
      self$person_nights <- gs$person_nights
      self$person_shifts <- gs$person_shifts
      self$pp_counts     <- gs$pp_counts
      self$granted_pto   <- gs$granted_pto
      invisible(self)
    }
  )
)
