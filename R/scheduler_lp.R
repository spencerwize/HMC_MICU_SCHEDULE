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
#   • Soft preferences (shift fairness) are part of the objective function,
#     not afterthoughts
#   • Naturally handles all pay-period interactions at once
#
# SOLVER: HiGHS (via the `highs` R package on CRAN)
# ─────────────────────────────────────────────────
# HiGHS is a state-of-the-art open-source MIP solver, typically 5–50× faster
# than lpSolve on this class of problem.  Constraints are accumulated as
# triplet lists and assembled into a single sparse matrix before solving.
#
# VARIABLE LAYOUT
# ───────────────
# Primary (binary):
#   x[p, d, s]   1 if person p is assigned to slot s on schedule day d
#   Flat 1-based index:   (p-1)*nD*4 + (d-1)*4 + s
#   Slot order:           APP1=1, APP2=2, Roaming=3, Night=4
#
# Fairness auxiliaries (continuous [0, nD]):
#   6 variables encoding per-metric max/min across all staff, used to push
#   min–max spread into the objective.
#   I_MAX_NIGHTS, I_MIN_NIGHTS, I_MAX_TOTAL, I_MIN_TOTAL, I_MAX_WKND, I_MIN_WKND
#
# Work auxiliaries (continuous [0, 1]):
#   work[p, d]  ≡  Σ_s x[p,d,s]   (integer automatically — sum of binaries, ub 1;
#   the ub also enforces no-double-booking, so no separate constraint is needed)
#   Flat 1-based index:   nX + nF + (p-1)*nD + d
#   Reduces C10 from 20 to 5 coefficients per sliding window.
#
# CONSTRAINTS
# ───────────
#   C1    Slot uniqueness:    Σ_p x[p,d,s]     ≤ 1   for all d, s ∈ {Roaming, Night}
#                             (APP1/APP2 covered by the C2/C2b equalities)
#   C2    APP1 coverage:      Σ_p x[p,d,APP1]  = 1           for all d
#   C2b   APP2 coverage:      Σ_p x[p,d,APP2]  = 1           for all d
#   C3    Night coverage:     uns[d] + Σ_p x[p,d,Night] ≥ 1  (soft, uns penalised)
#   C5    Availability:       ub = 0 for blocked (p,d) pairs
#   C6    PP shift cap:       Σ_{d∈PP,s} x[p,d,s] ≤ target   for all p,PP
#   C7    Night→Day 1d ban:   x[p,d,Night] + x[p,d+1,s] ≤ 1  s∈DAY_SLOTS
#   C7b   Night→Day 2d ban:   x[p,d,Night] + x[p,d+2,s] ≤ 1  s∈DAY_SLOTS  (2 rest days after last night)
#   C8    Day→Night 1d gap:   x[p,d,s] + x[p,d+1,Night] ≤ 1  s∈DAY_SLOTS  [LOCKED — never relaxed]
#   C9    Max 3 consec Nts:   Σ_{k=0}^3 x[p,d+k,Night]    ≤ 3  [LOCKED — never relaxed]
#   C10   Max 4 consec work:  Σ_{k=0}^4 work[p,d+k]       ≤ 4  [LOCKED — never relaxed]
#   C10b  Max 2 consec wknd: work[p,wday_k]+work[p,wday_{k+1}]+work[p,wday_{k+2}] ≤ 2  (Sat & Sun)
#   C10c  Density caps: ≤5 in 7/8-day windows; ≤6 in 9-day; ≤7 in 10/11-day  [LOCKED — never relaxed]
#   C11b  No night stretch in consecutive PPs: Σ_{d∈PP_k∪PP_{k+1}} s[p,d] ≤ 1  (s = stretch-start)
#   C11c  Weekend hard bounds: MIN_WKND_HARD ≤ Σ_{d∈Sat/Sun} work[p,d] ≤ MAX_WKND_HARD
#   C11d  Night hard bounds:   night_min_hard ≤ Σ_d x[p,d,Night] ≤ night_max_hard
#                             (subsumes the old C11 MAX_NIGHTS_TOTAL cap)
#   C_ns  Night soft-min:    ns_short[p] + Σ_d x[p,d,Night] ≥ MIN_NIGHTS_SOFT_TOTAL
#   C_ws  Weekend soft-min:  ws_short[p] + Σ_{d∈Sat/Sun} work[p,d] ≥ MIN_WKND_SOFT_TOTAL
#   C12   Holiday pre-seed:   lb = ub = 1 for pre-assigned (p,d,s)
#   C14   Fairness bounds:    Σ x ≤ max_M,  Σ x ≥ min_M  per person per metric
#   Run length is shaped ONLY by the objective: 3-runs heavily favoured (+3.0),
#   4 and 2 mildly favoured, solo days greatly penalised (-8.0).
#   C_w   Work definition:    work[p,d] = Σ_s x[p,d,s]
#
# OBJECTIVE (maximise)
# ────────────────────
#   Σ_{p,d,s}  w[s] × x[p,d,s]
#   + ε × (min_nights − max_nights + min_total − max_total + ...)
#   where w[APP1]=4, w[Night]=4, w[APP2]=2, w[Roaming]=2 (or 0 when relaxed)
#
# RELAXATION CASCADE
# ──────────────────
# Tries 7 tiers in order: 3 band tiers (night band 8–12 → 7–12 → 6–11), then 4
# nuclear fallbacks (at the loosest 6–11 band). Nights may be unstaffed (heavily
# penalised in the objective, never hard). C8/C9/C10/C10c are locked on in all of
# them. If all tiers fail, diagnostics print and execution stops.
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
    tier_used      = NULL,   # list(index, label) of relaxation tier that found a solution
    phase          = NULL,   # 1 = green-only (partial), 2 = filled. NULL until a run.
    pins_kept      = NULL,   # list(kept, total, mode) after run_fill()
    prior_schedule = NULL,  # person -> data.frame(date, type) for days before schedule start

    # ── Constructor ───────────────────────────────────────────────────────────
    initialize = function(time_off, targets, prior_schedule = NULL) {
      self$time_off       <- time_off
      self$targets        <- targets
      self$prior_schedule <- prior_schedule
      self$dates       <- all_dates()
      self$granted_pto <- setNames(
        lapply(STAFF, function(p) as.Date(character())), STAFF)

      private$reset_state()
    },

    # ── Unfilled slots and unmet floors ───────────────────────────────────────
    # One row per (date, slot) nobody covers, plus one row per person short of a
    # per-PP target or a schedule-wide floor. This single structure feeds the
    # holes worklist UI, the fill phase, and the summary stats.
    #
    # `soft` rows are advisory in the green-only phase (that phase enforces no
    # floors at all); the fill phase turns every one of them into a hard rule.
    holes_df = function() {
      rows <- list()

      # (a) uncovered slots. Roaming is optional by design, so it is not a hole.
      for (ds in names(self$schedule)) {
        day <- self$schedule[[ds]]
        for (sl in c("APP1", "APP2", "Night")) {
          v <- day[[sl]]
          if (length(v) != 1L || is.na(v))
            rows[[length(rows) + 1L]] <- data.frame(
              kind = "slot", date = as.Date(ds), slot = sl, person = NA_character_,
              pp = get_pp(as.Date(ds)), have = 0L, need = 1L, short = 1L,
              stringsAsFactors = FALSE)
        }
      }

      # (b) people below their pay-period target
      for (person in STAFF) {
        for (pp in PAY_PERIODS$name) {
          info <- self$targets[[person]][[pp]]
          have <- as.integer(self$pp_counts[[person]][[pp]])
          if (have < info$sched_target)
            rows[[length(rows) + 1L]] <- data.frame(
              kind = "target", date = as.Date(NA), slot = NA_character_,
              person = person, pp = pp, have = have,
              need = as.integer(info$sched_target),
              short = as.integer(info$sched_target) - have,
              stringsAsFactors = FALSE)
        }
      }

      # (c) schedule-wide floors the green-only phase is not allowed to enforce
      for (person in STAFF) {
        sh      <- self$person_shifts[[person]]
        nts     <- self$person_nights[[person]]
        worked  <- c(sh$date, nts)
        # Weekend load under the Fri-night..Sun definition.
        n_wknd  <- sum(is_weekend_shift(sh$date, sh$slot)) +
                   sum(is_weekend_shift(nts, "Night"))
        n_night <- length(self$person_nights[[person]])
        if (n_wknd < MIN_WKND_HARD)
          rows[[length(rows) + 1L]] <- data.frame(
            kind = "weekend_floor", date = as.Date(NA), slot = NA_character_,
            person = person, pp = NA_character_, have = as.integer(n_wknd),
            need = as.integer(MIN_WKND_HARD),
            short = as.integer(MIN_WKND_HARD - n_wknd), stringsAsFactors = FALSE)
        if (n_night < MIN_NIGHTS_HARD)
          rows[[length(rows) + 1L]] <- data.frame(
            kind = "night_floor", date = as.Date(NA), slot = NA_character_,
            person = person, pp = NA_character_, have = as.integer(n_night),
            need = as.integer(MIN_NIGHTS_HARD),
            short = as.integer(MIN_NIGHTS_HARD - n_night), stringsAsFactors = FALSE)
      }

      if (!length(rows))
        return(data.frame(kind = character(), date = as.Date(character()),
                          slot = character(), person = character(), pp = character(),
                          have = integer(), need = integer(), short = integer(),
                          stringsAsFactors = FALSE))
      out <- do.call(rbind, rows); rownames(out) <- NULL; out
    },

    # ── Current assignments as pin rows for the fill phase ────────────────────
    pinned_df = function() {
      rows <- list()
      for (ds in names(self$schedule)) {
        day <- self$schedule[[ds]]
        for (sl in SLOTS) {
          v <- day[[sl]]
          if (length(v) == 1L && !is.na(v))
            rows[[length(rows) + 1L]] <- data.frame(
              person = v, date = as.Date(ds), slot = sl, stringsAsFactors = FALSE)
        }
      }
      if (!length(rows))
        return(data.frame(person = character(), date = as.Date(character()),
                          slot = character(), stringsAsFactors = FALSE))
      out <- do.call(rbind, rows); rownames(out) <- NULL; out
    },

    # ── Eligible people for an unfilled (date, slot) ───────────────────────────
    # Checks every per-person hard rule via private$person_feasible(), so the
    # worklist dropdown can only offer legal choices. Returns a data.frame
    # ordered green-first, then by how far below target the person is.
    eligible_for = function(d, slot) {
      d  <- as.Date(d)
      ds <- as.character(d)
      taken <- unlist(lapply(SLOTS, function(sl) {
        v <- self$schedule[[ds]][[sl]]; if (length(v) == 1L && !is.na(v)) v else NULL
      }))
      pp <- get_pp(d)
      out <- lapply(STAFF, function(person) {
        if (person %in% taken)                return(NULL)   # already works that day
        if (private$is_blocked(person, d))    return(NULL)   # requested off
        vec <- private$assignment_vec(person)
        vec[ds] <- slot
        if (!private$person_feasible(person, vec, focus = d)) return(NULL)
        have <- if (is.na(pp)) NA_integer_ else as.integer(self$pp_counts[[person]][[pp]])
        need <- if (is.na(pp)) NA_integer_ else as.integer(self$targets[[person]][[pp]]$sched_target)
        data.frame(person = person,
                   green  = private$is_green(person, d),
                   have   = have, need = need,
                   short  = if (is.na(pp)) 0L else max(0L, need - have),
                   stringsAsFactors = FALSE)
      })
      out <- do.call(rbind, out[!vapply(out, is.null, logical(1L))])
      if (is.null(out) || !nrow(out))
        return(data.frame(person = character(), green = logical(), have = integer(),
                          need = integer(), short = integer(), stringsAsFactors = FALSE))
      out <- out[order(-out$green, -out$short, out$person), ]
      rownames(out) <- NULL; out
    },

    # ── Main entry point ──────────────────────────────────────────────────────
    # two_stage = TRUE  → lexicographic solve: Stage 1 (ILP) optimises fairness
    #                     only. Stage 2 then optimises run-length/aesthetics while
    #                     holding Stage 1's per-person counts fixed:
    #                       greedy_stage2 = TRUE  → fast greedy local-search pass
    #                                               (preserves counts EXACTLY; the
    #                                               only solver call is Stage 1).
    #                       greedy_stage2 = FALSE → second ILP pinned to Stage 1's
    #                                               counts (±1); for A/B comparison.
    #                     Roaming is filled inside the ILP, so the post-solve greedy
    #                     roaming top-up is skipped.
    # two_stage = FALSE → legacy single-stage solve (joint objective) + greedy
    #                     roaming fill. Kept for A/B benchmarking.
    # ── GREEN-ONLY phase ──────────────────────────────────────────────────────
    # Best schedule obtainable using ONLY requested-work days. Partial by design:
    # a date with fewer than two green volunteers cannot be fully covered, and a
    # person whose green days are scarce or crowded ends up below target. Inspect
    # the result with holes_df(), then complete it with run_fill().
    run_green = function(two_stage = TRUE, greedy_stage2 = TRUE, time_limit = NULL) {
      self$run(two_stage = two_stage, greedy_stage2 = greedy_stage2, phase = 1L,
               time_limit = time_limit)
    },

    # ── FILL phase ────────────────────────────────────────────────────────────
    # Completes a green-only schedule. `pinned` (default: everything the green-only
    # phase assigned, via pinned_df()) is forced, yellow days are unlocked, hard
    # coverage and every per-person floor are restored, and yellow work is
    # minimised - both in total and, via I_MAX_YELLOW, for the worst-off person.
    run_fill = function(pinned = NULL, two_stage = TRUE, greedy_stage2 = TRUE,
                        time_limit = NULL) {
      if (is.null(pinned)) pinned <- self$pinned_df()
      before <- pinned

      ok <- tryCatch({
        self$run(two_stage = two_stage, greedy_stage2 = greedy_stage2,
                 phase = 2L, pinned = pinned, pin_mode = "hard",
                 time_limit = time_limit)
        TRUE
      }, error = function(e) FALSE)

      if (!ok) {
        # Holding every green-only assignment is not always possible. That phase
        # maximises green fill without knowing which shifts the fill phase still
        # has to place, and hard APP1/APP2 coverage plus the night-adjacency rules
        # (C7/C7b/C8/C8b) can leave no completion at any relaxation tier.
        # Rather than fail, re-run with the assignments as a strong PREFERENCE.
        message("")
        message("  The green-only assignments cannot all be held while still")
        message("  covering every shift. Re-running with them as a strong")
        message("  preference instead of a hard requirement.")
        self$run(two_stage = two_stage, greedy_stage2 = greedy_stage2,
                 phase = 2L, pinned = pinned, pin_mode = "soft",
                 time_limit = time_limit)
      }

      # Report honestly how much of the green-only schedule survived.
      if (!is.null(before) && nrow(before) > 0L) {
        after <- self$pinned_df()
        kept  <- length(intersect(
          paste(before$person, as.character(before$date), before$slot),
          paste(after$person,  as.character(after$date),  after$slot)))
        self$pins_kept <- list(kept = kept, total = nrow(before),
                               mode = if (ok) "hard" else "soft")
        message(sprintf("  Kept %d of %d green-only assignments (%.0f%%).",
                        kept, nrow(before), 100 * kept / nrow(before)))
      }
      invisible(self)
    },

    # `time_limit` - seconds per solve; NULL uses SOLVER_TIME_LIMIT. Passed
    # through rather than set globally so concurrent runs cannot interfere.
    run = function(two_stage = TRUE, greedy_stage2 = TRUE,
                   phase = 2L, pinned = NULL, pin_mode = "hard",
                   time_limit = NULL) {
      if (!requireNamespace("highs", quietly = TRUE))
        stop("highs package is not installed. Run: install.packages('highs')")

      green_only <- (phase == 1L)
      message(if (green_only)
                "  Phase 1 (GREEN-ONLY): requested-work days only; holes allowed."
              else
                "  Phase 2 (FILL): yellow days unlocked; coverage and floors enforced.")

      # A fresh solve must not accumulate onto a previous phase's schedule.
      private$reset_state()

      tiers <- private$RELAX_TIERS
      nT    <- length(tiers)

      for (ti in seq_len(nT)) {
        t <- tiers[[ti]]
        message(sprintf("  [Tier %d/%d] %s", ti, nT, t$label))
        args <- private$build_args(t)
        phase_args <- list(phase = phase, pinned = pinned, pin_mode = pin_mode,
                           time_limit_override = time_limit)

        # Stage 1 (fairness) — or the full legacy model when two_stage = FALSE.
        r1 <- do.call(private$build_and_solve,
                      c(args, phase_args,
                        list(stage = if (two_stage) 1L else 0L)))
        if (is.null(r1)) next

        result <- r1
        if (two_stage && !greedy_stage2) {
          # Stage 2 via a second ILP (A/B path).
          ft <- private$extract_fairness_targets(r1)
          r2 <- do.call(private$build_and_solve,
                        c(args, phase_args,
                          list(stage = 2L, fix_targets = ft, fix_band = 1L)))
          if (!is.null(r2)) {
            result <- r2
          } else {
            message("  Stage 2 ILP returned no solution — keeping Stage 1 result.")
          }
        }

        self$phase     <- phase
        # Carry the tier's relaxation flags so validate_schedule() can tell a
        # genuine violation from a rule this tier was permitted to drop.
        self$tier_used <- list(index = ti, label = t$label, relaxed = t)
        if (green_only && ti > 1L)
          message("  NOTE: the green-only phase needed a relaxation tier. ",
                  "It enforces no per-person floors, so a binding constraint here ",
                  "is a CEILING - check the parser did not read work keywords as OFF.")
        # PTO is reported as a per-PP count (targets$pto_needed) in the summary
        # sheet, not pinned onto specific calendar days — so granted_pto is left
        # empty and OFF/VAC days render as themselves on the schedule grid.
        message("  Populating solution…")
        private$populate_from_solution(result)
        # Two-stage: the ILP fills Roaming itself (ROAM_FILL_WEIGHT), so the greedy
        # roaming top-up is skipped (it would push counts past the fixed fairness).
        if (two_stage && greedy_stage2) {
          message("  Stage 2: greedy aesthetic pass…")
          private$greedy_aesthetic_pass(
            green_only = green_only,
            pinned = if (identical(pin_mode, "hard")) pinned else NULL)
        } else if (!two_stage) {
          private$fill_roaming_pass()
        }
        message("  Done.")
        return(invisible(self))
      }

      private$report_and_stop()
    },

    # ── Enumerate distinct feasible solutions via no-good cut iteration ───────
    # Adds a cut excluding each previously found x-assignment, then re-solves.
    # Returns integer count (lower bound; stops at max_count or infeasibility).
    # Each iteration gets its own SHORT time limit (per_solve_secs): the goal is
    # "does another feasible schedule exist?", not optimality — without this cap
    # every iteration would inherit the full SOLVER_TIME_LIMIT budget.
    count_solutions = function(max_count = 20L, per_solve_secs = 120) {
      if (is.null(self$tier_used)) stop("Call run() before count_solutions().")
      t     <- private$RELAX_TIERS[[self$tier_used$index]]
      found <- list()
      count <- 0L
      repeat {
        res <- do.call(private$build_and_solve,
                       c(private$build_args(t),
                         list(extra_nogo = found, stage = 0L,
                              time_limit_override = per_solve_secs)))
        if (is.null(res)) break
        count <- count + 1L
        found <- c(found, list(res$sol[seq_len(res$nX)]))
        if (count >= max_count) break
      }
      count
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
        pp <- get_pp(d)
        lapply(STAFF, function(person) {
          # Role logic lives in R/roles.R — shared with the Calendar tab and
          # the Excel export so the three views cannot drift apart again.
          role <- role_of(person, d, self$schedule, time_off, self$granted_pto)
          data.frame(
            date       = d,
            day_name   = weekdays(d, abbreviate = TRUE),
            pp         = pp,
            person     = person,
            role       = role,
            # `wants` is orthogonal to `role`: a WORKED green day keeps its slot
            # role ("APP1", "Night", …) and is flagged here so the display can
            # outline it. An UNWORKED green day gets role "WANT".
            wants      = is_green_day(person, d, time_off),
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

    # ── Clear all derived schedule state ──────────────────────────────────────
    # Called from initialize() and again at the top of every run(), so a fill
    # phase starts from an empty grid and rebuilds from its own solution rather
    # than accumulating onto the green-only phase's counts.
    reset_state = function() {
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
      self$pp_counts <- setNames(lapply(STAFF, function(p) zero_pp), STAFF)
      invisible(self)
    },

    # ── One person's assignments as a named vector: date-string -> slot ───────
    # The shape private$person_feasible() expects. Mirrors how the greedy pass
    # builds `asgn`, but reads from self$schedule so callers outside the pass
    # (the holes worklist) can use it too.
    assignment_vec = function(person) {
      v <- character(0)
      for (ds in names(self$schedule)) {
        day <- self$schedule[[ds]]
        for (sl in SLOTS) {
          val <- day[[sl]]
          if (length(val) == 1L && !is.na(val) && val == person) { v[ds] <- sl; break }
        }
      }
      v
    },

    # ── Availability helpers ──────────────────────────────────────────────────
    # NOTE: "green" is deliberately NOT blocked here — a requested WORK day is
    # fully schedulable. Only off/vac/cme remove a day from consideration.
    is_blocked = function(person, d) {
      pdata <- self$time_off[[person]]
      nrow(pdata) > 0 && any(pdata$date == d & pdata$type %in% BLOCKED_TYPES)
    },

    # TRUE when `person` marked `d` as a requested WORK day (green).
    is_green = function(person, d) {
      pdata <- self$time_off[[person]]
      nrow(pdata) > 0 && any(pdata$date == d & pdata$type == "green")
    },

    # Pay-period INDEX (integer row of PAY_PERIODS) for a date, NA if outside.
    # Distinct from get_pp() in constants.R, which returns the PP *name*.
    pp_index = function(d) {
      i <- which(PAY_PERIODS$start <= d & PAY_PERIODS$end >= d)
      if (length(i)) i[1] else NA_integer_
    },

    # ── Full per-person hard-feasibility check ────────────────────────────────
    # `vec` is a named character vector  date-string -> slot, i.e. one person's
    # complete set of assignments. Re-checks by hand every per-person hard rule
    # the ILP enforces: C5 availability, C9 (<=3 consecutive nights), C10 (<=4
    # consecutive work days), C7/C7b/C8/C8b (night<->day separation), C10c
    # (density caps), C10b (no 3 consecutive same weekday), C11b (no night
    # stretch start in consecutive pay periods).
    #
    # `focus` (a Date, typically a newly added day) scopes the expensive density
    # scan to windows covering it.
    #
    # Extracted from greedy_aesthetic_pass() so the holes worklist can offer
    # only legally-eligible people for an unfilled slot.
    person_feasible = function(person, vec, focus = NULL) {
      if (length(vec) == 0L) return(TRUE)
      dts  <- as.Date(names(vec)); slts <- unname(vec)
      if (any(vapply(dts, function(d) private$is_blocked(person, d), logical(1L)))) return(FALSE)
      ints   <- sort(as.integer(dts))
      nights <- sort(as.integer(dts[slts == "Night"]))
      days_  <- sort(as.integer(dts[slts %in% DAY_SLOTS]))
      # C5b: no night the day before an off / vacation / CME day. A night runs
      # into the next morning, so it collides with whatever that day is for.
      # Checked here as well as in the ILP, or the greedy pass could swap a night
      # onto such a day and produce a schedule the ILP would reject.
      if (length(nights)) {
        pdata <- self$time_off[[person]]
        if (!is.null(pdata) && nrow(pdata) > 0) {
          blocked_next <- pdata$date[pdata$type %in% BLOCKED_TYPES]
          if (length(blocked_next) &&
              any((as.Date(nights, origin = "1970-01-01") + 1L) %in% blocked_next))
            return(FALSE)
        }
      }
      # C9: <=3 consecutive nights
      if (length(nights) > 1L) { r <- 1L
        for (i in 2:length(nights)) { if (nights[i] == nights[i-1] + 1L) { r <- r + 1L; if (r > 3L) return(FALSE) } else r <- 1L } }
      # C9b: nights must be stacked - a gap of exactly one day between two
      # nights is banned (Night-Empty-Night). Checked here too, or the greedy
      # pass could swap a night into the gap position.
      if (length(nights) > 1L)
        for (i in 2:length(nights))
          if (nights[i] - nights[i - 1L] == 2L) return(FALSE)
      # C10: <=4 consecutive work days
      if (length(ints) > 1L) { r <- 1L
        for (i in 2:length(ints)) { if (ints[i] == ints[i-1] + 1L) { r <- r + 1L; if (r > 4L) return(FALSE) } else r <- 1L } }
      # C7/C7b/C8/C8b: a night and a day shift cannot be within 2 days (either order)
      if (length(nights) && length(days_))
        for (n in nights) if (any(abs(days_ - n) <= 2L)) return(FALSE)
      # C10c: density caps
      dens <- list(c(7L,5L), c(8L,5L), c(9L,6L), c(10L,7L), c(11L,7L))
      if (length(ints)) {
        if (is.null(focus)) { lo <- min(ints); hi <- max(ints) }
        else { fi <- as.integer(focus); lo <- fi; hi <- fi }
        for (wb in dens) { W <- wb[1]; B <- wb[2]
          starts <- if (is.null(focus)) lo:hi else (fi - W + 1L):fi
          for (st in starts) if (sum(ints >= st & ints <= st + W - 1L) > B) return(FALSE) }
      }
      # C10b: no 3 consecutive same-weekday (Sat or Sun) worked
      for (wd in c("Saturday", "Sunday")) {
        wdd <- sort(as.integer(dts[weekdays(dts) == wd]))
        for (s in wdd) if ((s + 7L) %in% wdd && (s + 14L) %in% wdd) return(FALSE)
      }
      # C11b: no night-stretch-start in consecutive pay periods
      if (length(nights)) {
        starts <- nights[!((nights - 1L) %in% nights)]
        pps <- sort(unique(vapply(starts, function(x) private$pp_index(as.Date(x, origin = "1970-01-01")),
                                  integer(1L))))
        pps <- pps[!is.na(pps)]
        if (length(pps) >= 2L && any(diff(pps) == 1L)) return(FALSE)
      }
      TRUE
    },

    # ── Map a relaxation tier to the build_and_solve argument list ─────────────
    # Single source of truth for the per-tier args, shared by run() (Stage 1 and
    # Stage 2) and count_solutions() via do.call().
    build_args = function(t) {
      list(
        pp_cap_reduction = t$pp_red,
        add_c8b          = t$c8b,
        add_c11b         = t$c11b,
        add_c11c         = t$c11c,
        night_min_hard   = t$night_min_hard,
        night_max_hard   = t$night_max_hard,
        add_c14          = t$c14,
        add_c_min        = t$c_min
      )
    },

    # ── Extract per-person fairness counts from a (Stage-1) solution ───────────
    # Returns list(night, weekend, total), each a length-nP integer vector, used
    # to pin Stage 2 to Stage 1's fairness via the C_fix constraints.
    extract_fairness_targets = function(r1) {
      sol <- r1$sol; nP <- r1$nP; nD <- r1$nD; nS <- r1$nS; xidx <- r1$xidx
      # Must match the weekend definition used by C11c / C14 / C_ws, or the
      # Stage-2 pin would fix a count the Stage-2 model does not measure.
      fri_di    <- which(weekdays(r1$dates_vec) == "Friday")
      satsun_di <- which(weekdays(r1$dates_vec) %in% c("Saturday", "Sunday"))
      night <- weekend <- total <- integer(nP)
      for (pi in seq_len(nP)) {
        block       <- ((pi - 1L) * nD * nS + 1L):(pi * nD * nS)
        night[pi]   <- as.integer(sum(round(sol[xidx(pi, seq_len(nD), 4L)])))
        total[pi]   <- as.integer(sum(round(sol[block])))
        wcols <- c(
          if (length(fri_di))    xidx(pi, fri_di, 4L) else integer(0),
          if (length(satsun_di)) xidx(pi, rep(satsun_di, each = nS),
                                      rep.int(seq_len(nS), length(satsun_di)))
          else integer(0))
        weekend[pi] <- if (!length(wcols)) 0L else as.integer(sum(round(sol[wcols])))
      }
      list(night = night, weekend = weekend, total = total)
    },

    # ── Relaxation cascade: 3 band tiers + 4 nuclear fallbacks ────────────────
    # Nights are never hard-required (any night may be unstaffed, penalised
    # heavily in the objective). The per-person night band starts strict at 8–12
    # and loosens 8–12 → 7–12 → 6–11 before any other rule is touched. Run length
    # is shaped by the objective (no min-run constraints). The nuclear fallbacks
    # (at band 6–11) relax PP cap, C8b, C11b, C11c, per-PP minimums, and (last)
    # fairness spread. The locked safety rules (C7/C7b/C8/C9/C10/C10b/C10c) are
    # built unconditionally in build_and_solve and never appear here.
    RELAX_TIERS = {
      mk <- function(label, night_min, night_max,
                     pp_red = 0L, c_min = TRUE, c8b = TRUE,
                     c11b = TRUE, c11c = TRUE, c14 = TRUE) {
        list(label = label, pp_red = pp_red, c_min = c_min, c8b = c8b,
             c11b = c11b, c11c = c11c, c14 = c14,
             night_min_hard = night_min, night_max_hard = night_max)
      }
      list(
        # ── Primary + night-band loosening ─────────────────────────────────────
        # Nights may be left unstaffed (never hard-required) but each empty night
        # costs UNSTAFFED_NIGHT_PEN in the objective, so the solver staffs every
        # night it feasibly can. The per-person night band starts STRICT at 8–12;
        # if that is infeasible it loosens to 7–12, then 6–11, before any other
        # rule is relaxed.
        mk("Band 8–12", night_min = 8L, night_max = 12L),
        mk("Band 7–12", night_min = 7L, night_max = 12L),
        mk("Band 6–11", night_min = 6L, night_max = 11L),
        # ── Nuclear fallbacks (at the loosest 6–11 band) ───────────────────────
        # Reaching here means even the 6–11 band could not satisfy the other
        # rules, so these relax: PP cap, C8b, C11b, C11c, per-PP minimums, and
        # (last) the fairness spread. If none solve, diagnostics are reported.
        mk("Relax PP cap by 1",
           night_min = 6L, night_max = 11L, pp_red = 1L),
        mk("Drop C8b + C11b (2-day day→night gap + PP-stretch rule)",
           night_min = 6L, night_max = 11L, pp_red = 1L,
           c8b = FALSE, c11b = FALSE),
        mk("Drop C8b + C11b + per-PP minimums + weekend bounds",
           night_min = 6L, night_max = 11L, pp_red = 1L,
           c_min = FALSE, c8b = FALSE, c11b = FALSE, c11c = FALSE),
        mk("Locked safety rules only (also drop fairness spread)",
           night_min = 6L, night_max = 11L,
           c_min = FALSE, c8b = FALSE, c11b = FALSE, c11c = FALSE, c14 = FALSE)
      )
    },

    # ── Build and solve the ILP (HiGHS) ──────────────────────────────────────
    # The locked safety rules — C7/C7b (night→day rest), C8 (day→night gap),
    # C9 (≤3 consecutive nights), C10 (≤4 consecutive work days), C10b, C10c
    # (density caps) — are built unconditionally: no tier can relax them. If a
    # schedule is infeasible under them, it fails rather than violating them.
    build_and_solve = function(
      pp_cap_reduction = 0L,
      add_c8b          = TRUE,          # Day→Night 2d gap: x[p,d,s] + x[p,d+2,Night] ≤ 1
      add_c11b         = TRUE,          # no night stretch in consecutive PPs
      add_c11c         = TRUE,          # weekend hard bounds: MIN_WKND_HARD ≤ total ≤ MAX_WKND_HARD
      night_min_hard   = MIN_NIGHTS_HARD, # lower hard bound on per-person night total (NULL = no bound)
      night_max_hard   = MAX_NIGHTS_HARD, # upper hard bound on per-person night total (NULL = no bound)
      add_c14          = TRUE,          # fairness min/max spread bounds
      add_c_min        = TRUE,          # per-PP soft_min floors (FLEX_TARGETS aware)
      extra_nogo       = list(),        # previously found x-vectors to exclude (solution enumeration)
      stage            = 0L,            # 0 = single-stage (legacy); 1 = fairness only; 2 = aesthetic w/ fairness pinned
      fix_targets      = NULL,          # Stage 2: list(night=<nP>, weekend=<nP>, total=<nP>) per-person counts from Stage 1
      fix_band         = 1L,            # Stage 2: ± tolerance on the fairness fix (0 = exact equality)
      time_limit_override = NULL,       # seconds; NULL = SOLVER_TIME_LIMIT
      # ── Green-first phase ───────────────────────────────────────────────────
      #   1 = GREEN-ONLY. Only requested-work days are schedulable; APP1/APP2
      #       coverage becomes SOFT (holes allowed) and every per-person FLOOR
      #       is dropped, because the output is partial by construction.
      #   2 = FILL. Yellow days unlocked, hard coverage and all floors restored,
      #       Phase 1's assignments pinned via `pinned`, yellow work minimised.
      # NB: this is orthogonal to `stage` (the fairness/aesthetics split).
      phase            = 2L,
      pinned           = NULL,          # Phase 2: data.frame(person, date, slot) to keep
      # "hard" = force pinned assignments (lb = 1). "soft" = reward them in the
      # objective instead, so the solver keeps as many as it feasibly can.
      pin_mode         = "hard"
    ) {

      dates_vec <- as.Date(self$dates, origin = "1970-01-01")
      nP <- length(STAFF)
      nD <- length(dates_vec)
      nS <- 4L

      S_APP1 <- 1L; S_APP2 <- 2L; S_ROAM <- 3L; S_NIGHT <- 4L
      DAY_S  <- c(S_APP1, S_APP2, S_ROAM)

      green_only <- (phase == 1L)

      # ── Day-tier matrices (nP x nD logicals) ────────────────────────────────
      # Built once here instead of rescanning each person's time-off data.frame
      # inside the nP x nD constraint loops (that was ~1k linear scans per build).
      #   blocked = off / vac / cme  -> never schedulable
      #   green   = requested WORK   -> the only pool Phase 1 may draw on
      #   yellow  = neither          -> schedulable in Phase 2 only, and minimised
      date_keys   <- as.character(dates_vec)
      blocked_mat <- matrix(FALSE, nP, nD)
      green_mat   <- matrix(FALSE, nP, nD)
      for (pi in seq_len(nP)) {
        pdata <- self$time_off[[STAFF[pi]]]
        if (is.null(pdata) || nrow(pdata) == 0L) next
        j  <- match(as.character(pdata$date), date_keys)
        ok <- !is.na(j)
        blocked_mat[pi, j[ok & pdata$type %in% BLOCKED_TYPES]] <- TRUE
        green_mat[pi,   j[ok & pdata$type == "green"]]                  <- TRUE
      }
      # A day cannot be both blocked and green; blocked wins if the sheet says both.
      green_mat[blocked_mat] <- FALSE
      yellow_mat <- !green_mat & !blocked_mat

      # Days this phase may NOT use at all.
      unavail_mat <- if (green_only) (blocked_mat | yellow_mat) else blocked_mat

      # ── Prior-schedule precomputation ─────────────────────────────────────────
      sched_day0   <- dates_vec[1L] - 1L          # last date before schedule
      ps_has_prior <- !is.null(self$prior_schedule)
      if (ps_has_prior) {
        .ps_nights <- lapply(seq_len(nP), function(pi) {
          ps <- self$prior_schedule[[STAFF[pi]]]
          if (is.null(ps) || nrow(ps) == 0L) as.Date(character())
          else as.Date(ps$date[ps$type == "night"])
        })
        .ps_days <- lapply(seq_len(nP), function(pi) {
          ps <- self$prior_schedule[[STAFF[pi]]]
          if (is.null(ps) || nrow(ps) == 0L) as.Date(character())
          else as.Date(ps$date[ps$type == "day"])
        })
        .ps_all <- lapply(seq_len(nP), function(pi)
          c(.ps_nights[[pi]], .ps_days[[pi]]))
      } else {
        .ps_nights <- lapply(seq_len(nP), function(pi) as.Date(character()))
        .ps_days   <- lapply(seq_len(nP), function(pi) as.Date(character()))
        .ps_all    <- lapply(seq_len(nP), function(pi) as.Date(character()))
      }

      # x[p,d,s] binary
      nX   <- nP * nD * nS
      xidx <- function(p, d, s) (p - 1L) * nD * nS + (d - 1L) * nS + s

      # Fairness auxiliaries: 6 continuous [0, nD].
      # (Roaming max/min were dropped — they carried zero objective weight, so
      # they were dead columns + 2 constraint rows per person.)
      #
      # I_MAX_YELLOW is the 7th: the largest number of YELLOW (neither requested
      # off nor requested to work) days any one person works. Penalised in the
      # objective so the fill phase spreads unavoidable yellow work instead of
      # dumping it on whoever is cheapest to move. There is deliberately no
      # I_MIN_YELLOW - a min-side term would REWARD loading someone with yellow.
      #
      # Allocated unconditionally: every downstream offset (widx, ns3off, nV, ub,
      # types) is computed FROM nF, so they all adapt on their own.
      nF           <- 7L
      I_MAX_NIGHTS <- nX + 1L;  I_MIN_NIGHTS <- nX + 2L
      I_MAX_TOTAL  <- nX + 3L;  I_MIN_TOTAL  <- nX + 4L
      I_MAX_WKND   <- nX + 5L;  I_MIN_WKND   <- nX + 6L
      I_MAX_YELLOW <- nX + 7L

      # work[p,d] continuous [0,1] — equals Σ_s x[p,d,s]. Integer automatically
      # (sum of binaries); its ub of 1 doubles as the no-double-book rule.
      nW   <- nP * nD
      widx <- function(p, d) nX + nF + (p - 1L) * nD + d

      # Streak auxiliaries (all continuous [0,1]):
      #   n3/n4 bias night assignments toward 3-packs (3 > 2 > 4)
      #   w3/w4 bias work-day runs toward 3-blocks (3 > 4 > 2 > 1)
      # These are AESTHETIC: suppressed in Stage 1 (fairness-only) by zero-width
      # allocation so they neither enlarge the model nor influence the search.
      # Gated here at definition (NOT later) so every downstream offset and nV
      # recompute correctly — offsets are interleaved with the counts below.
      aesthetic_on <- (stage != 1L)
      nNS3  <- if (aesthetic_on && nD >= 3L) nP * (nD - 2L) else 0L
      nNS4  <- if (aesthetic_on && nD >= 4L) nP * (nD - 3L) else 0L
      nWS3  <- if (aesthetic_on && nD >= 3L) nP * (nD - 2L) else 0L
      nWS4  <- if (aesthetic_on && nD >= 4L) nP * (nD - 3L) else 0L

      ns3off <- nX + nF + nW
      n3idx  <- function(p, d) ns3off + (p - 1L) * (nD - 2L) + d
      ns4off <- ns3off + nNS3
      n4idx  <- function(p, d) ns4off + (p - 1L) * (nD - 3L) + d
      ws3off <- ns4off + nNS4
      w3idx  <- function(p, d) ws3off + (p - 1L) * (nD - 2L) + d
      ws4off <- ws3off + nWS3
      w4idx  <- function(p, d) ws4off + (p - 1L) * (nD - 3L) + d

      # Isolation-cap auxiliaries: iso[p,d] ∈ [0,1] continuous,
      # constrained to equal 1 when work[p,d]=1 and both neighbours=0.
      # Continuous is safe: the C15b lower bounds are integer expressions of
      # work[], and the -8 objective weight pins iso to that (integral) bound.
      nISO   <- if (aesthetic_on && nD >= 2L) nP * nD else 0L
      isoff  <- ws4off + nWS4
      isoidx <- function(p, d) isoff + (p - 1L) * nD + d

      # Soft-minimum shortfall variables (continuous >= 0):
      #   ns_short[p] = max(0, MIN_NIGHTS_SOFT_TOTAL  - actual nights for p)
      #   ws_short[p] = max(0, MIN_WKND_SOFT_TOTAL    - actual Sat+Sun shifts for p)
      # Each is penalised in the objective; the solver treats them as soft floors.
      nNSSHORT  <- nP
      nsshoff   <- isoff + nISO
      nsshidx   <- function(p) nsshoff + p
      nWSSHORT  <- nP
      wsshoff   <- nsshoff + nNSSHORT
      wsshidx   <- function(p) wsshoff + p

      # Night stretch-start auxiliaries: ss[p,d] ∈ [0,1] continuous.
      # Forced equal to 1 iff x[p,d,Night]=1 and x[p,d-1,Night]=0.
      nSS    <- nP * nD
      ssoff  <- wsshoff + nWSSHORT
      ssidx  <- function(p, d) ssoff + (p - 1L) * nD + d

      # Unstaffed-night indicators: uns[d] ∈ [0,1] continuous, forced to 1 when no
      # one covers the Night slot on day d (uns[d] + Σ_p x[p,d,Night] ≥ 1). Each is
      # penalised −UNSTAFFED_NIGHT_PEN in the objective so the solver staffs every
      # night it feasibly can. Always allocated (operational, not aesthetic).
      nUNS   <- nD
      unsoff <- ssoff + nSS
      unsidx <- function(d) unsoff + d

      # Unfilled DAY-slot indicators: uf1[d] / uf2[d] in [0,1], forced to 1 when
      # nobody covers APP1 / APP2 on day d. These exist ONLY in the green-only
      # phase, where C2/C2b relax from hard equalities to the same soft-coverage
      # form the Night slot has always used (C3 + uns[d]).
      #
      # GATING POLARITY WARNING: this is the OPPOSITE of `aesthetic_on` above.
      # Aesthetic blocks are suppressed in stage 1 (the production path); these
      # are allocated ONLY in phase 1. Do not copy that guard.
      # Weekend-split indicators: wsp[p,w] in [0,1] for each Sat/Sun pair w,
      # forced to 1 when the person works exactly one of the two days.
      # Penalised in the objective so whole weekends are preferred over the same
      # number of shifts scattered across twice as many weekends.
      #
      # Allocated unconditionally - NOT gated on aesthetic_on. This must be live
      # in stage 1, which is the production path.
      sat_di  <- which(weekdays(dates_vec) == "Saturday")
      wk_pair <- sat_di[(sat_di + 1L) <= nD]          # Saturdays with a Sunday after
      nWK     <- length(wk_pair)
      nWSP    <- nP * nWK
      wspoff  <- 0L                                    # set after the earlier blocks
      wspidx  <- function(p, w) wspoff + (p - 1L) * nWK + w

      nUF1   <- if (green_only) nD else 0L
      nUF2   <- if (green_only) nD else 0L
      uf1off <- unsoff + nUNS
      uf1idx <- function(d) uf1off + d
      uf2off <- uf1off + nUF1
      uf2idx <- function(d) uf2off + d
      wspoff <- uf2off + nUF2

      nV <- nX + nF + nW + nNS3 + nNS4 + nWS3 + nWS4 + nISO +
            nNSSHORT + nWSSHORT + nSS + nUNS + nUF1 + nUF2 + nWSP

      # Variable bounds and types (x is the only integer block)
      lb    <- numeric(nV)
      ub    <- c(rep(1, nX), rep(as.double(nD), nF), rep(1, nW),
                 rep(1, nNS3), rep(1, nNS4), rep(1, nWS3), rep(1, nWS4),
                 rep(1, nISO),
                 rep(as.double(MIN_NIGHTS_SOFT_TOTAL), nNSSHORT),
                 rep(as.double(MIN_WKND_SOFT_TOTAL),   nWSSHORT),
                 rep(1, nSS), rep(1, nUNS), rep(1, nUF1), rep(1, nUF2),
                 rep(1, nWSP))
      types <- c(rep("I", nX), rep("C", nV - nX))
      stopifnot(length(ub) == nV, length(lb) == nV, length(types) == nV)

      # Objective (maximise) — x weights assigned by slot across all (p,d) at
      # once (xidx is pure arithmetic, so a vector of x-indices per slot works).
      obj <- numeric(nV)
      all_pd <- rep(seq_len(nP), each = nD)          # person of each (p,d) pair
      all_dd <- rep(seq_len(nD), times = nP)         # day    of each (p,d) pair
      obj[xidx(all_pd, all_dd, S_APP1)]  <- 4
      obj[xidx(all_pd, all_dd, S_APP2)]  <- 0.1
      # Roaming (APP3) is optional; this weight makes the ILP fill feasible
      # Roaming slots itself (up to PP/density caps) instead of a post-solve
      # greedy. Kept below the fairness-spread coefficients so it shapes
      # coverage without driving the fairness distribution.
      obj[xidx(all_pd, all_dd, S_ROAM)]  <- ROAM_FILL_WEIGHT
      obj[xidx(all_pd, all_dd, S_NIGHT)] <- 0.1
      # Fairness is the primary objective — large coefficients drive the solver
      # to minimise the night/weekend/total spread, not just fill slots.
      obj[I_MAX_NIGHTS] <- -5.0;  obj[I_MIN_NIGHTS] <- +5.0
      obj[I_MAX_TOTAL]  <- -2.0;  obj[I_MIN_TOTAL]  <- +2.0
      obj[I_MAX_WKND]   <- -3.0;  obj[I_MIN_WKND]   <- +3.0

      # ── Weekend pairing ─────────────────────────────────────────────────────
      # Each split weekend (working Saturday or Sunday but not both) costs
      # WEEKEND_SPLIT_PEN. Two lone days across two weekends therefore cost more
      # than the same two shifts on one weekend, which is what makes a schedule
      # "feel" like fewer weekends worked for the same shift count.
      if (nWSP > 0L)
        obj[(wspoff + 1L):(wspoff + nWSP)] <- -WEEKEND_SPLIT_PEN

      # ── Green / yellow day preference (FILL phase only) ─────────────────────
      # In the green-only phase every schedulable day is green, so a bonus there
      # would be a constant; the preference only means something once yellow days
      # are unlocked.
      #
      # A POSITIVE bonus on green days, not a negative penalty on yellow ones.
      # They are not equivalent: a person's shift count is not fixed (C6 caps,
      # C_min floors, with slack between), so a yellow penalty would suppress
      # total shifts and - at any weight above ROAM_FILL_WEIGHT (1.5) - make an
      # optional Roaming slot objective-NEGATIVE on a yellow day, regressing
      # coverage. It would also drag the objective toward zero, where the
      # relative MIP gap stops being meaningful.
      #
      # Weight ladder this must sit inside (effective price per unit):
      #   unstaffed night 100 | night spread 10 | night shortfall 8
      #   weekend spread 6    | TOTAL SPREAD 4  | weekend shortfall 4
      #   Roaming fill 1.5    | APP2/Night slot 0.1
      # So the usable band is (1.5, 4): above Roaming so it actually bites, below
      # the tightest fairness term so it can never buy a green day by widening a
      # spread. GREEN_WORK_BONUS = 3.0 sits in the middle.
      #
      # The coefficient rides on work[p,d] rather than the four x[p,d,s]: C_work
      # pins work to their sum as an EQUALITY, so this is exactly equivalent to
      # adding it to each slot, with a quarter as many coefficients and no
      # LP-relaxation leakage.
      if (!green_only) {
        green_flag <- as.vector(t(green_mat))   # person-major/day-minor: matches widx
        obj[widx(all_pd, all_dd)] <- GREEN_WORK_BONUS * green_flag
        obj[I_MAX_YELLOW]         <- -MAX_YELLOW_PEN
      }

      # Run-length preferences (the ONLY thing shaping run length now that the
      # min-run hard constraints are gone): heavily favour 3-in-a-row, mildly
      # favour 4 and 2, greatly penalise solo (length-1) shifts.
      #   Work days: 3=+3.0, 4=+1.0, 2=0 (neutral), 1=-8.0
      #   Nights:    3=+3.0, 2=0, 4=-5.0 (4-night run infeasible via C9; dead weight)
      if (nNS3 > 0L) obj[(ns3off + 1L):(ns3off + nNS3)] <- +3.0
      if (nNS4 > 0L) obj[(ns4off + 1L):(ns4off + nNS4)] <- -5.0
      if (nWS3 > 0L) obj[(ws3off + 1L):(ws3off + nWS3)] <- +3.0
      if (nWS4 > 0L) obj[(ws4off + 1L):(ws4off + nWS4)] <- +1.0
      # Isolated single (solo) shifts: greatly penalised, but allowed (not banned).
      if (nISO > 0L) obj[(isoff + 1L):(isoff + nISO)]   <- -8.0
      # Soft-minimum penalties: penalise each unit a person falls below the
      # schedule-wide night/weekend floor.  Penalty > base assignment weight so
      # the solver strongly prefers reaching the minimum before going above it.
      NIGHTS_SHORT_PEN <- 8.0   # per night below MIN_NIGHTS_SOFT_TOTAL
      WKND_SHORT_PEN   <- 4.0   # per Sat/Sun shift below MIN_WKND_SOFT_TOTAL
      for (p in seq_len(nP)) {
        obj[nsshidx(p)] <- -NIGHTS_SHORT_PEN
        obj[wsshidx(p)] <- -WKND_SHORT_PEN
      }
      # Heavy penalty per unstaffed night (uns[d] = 1 when the Night slot is empty).
      for (d in seq_len(nD)) obj[unsidx(d)] <- -UNSTAFFED_NIGHT_PEN
      # Green-only phase: same treatment for the two day slots, which are hard
      # equalities in the fill phase but must be allowed to go unfilled here.
      # Both penalties sit far above every fairness coefficient (max 5), so the
      # phase fills everything green availability permits before shaping anything.
      # APP1 is priced above APP2 so that when only one can be staffed, APP1 wins.
      if (green_only) {
        for (d in seq_len(nD)) {
          obj[uf1idx(d)] <- -UNFILLED_APP1_PEN
          obj[uf2idx(d)] <- -UNFILLED_APP2_PEN
        }
      }

      # Constraint accumulator (triplet form → sparseMatrix).
      # Per-constraint index/coeff chunks are collected in LISTS and unlist()ed
      # once at assembly. The previous `c(ri, ...)` append recopied the full
      # triplet vectors on every one of the ~25k add_con calls — O(n²), billions
      # of element copies per model build.
      n_con   <- 0L
      ri_l    <- list(); ci_l <- list(); vi_l <- list()
      con_lhs <- double(0); con_rhs <- double(0)

      add_con <- function(indices, coeffs, type, rhs_val) {
        n_con <<- n_con + 1L
        ri_l[[n_con]] <<- rep.int(n_con, length(indices))
        ci_l[[n_con]] <<- as.integer(indices)
        vi_l[[n_con]] <<- as.double(coeffs)
        if (type == "<=") {
          con_lhs[[n_con]] <<- -Inf;    con_rhs[[n_con]] <<- rhs_val
        } else if (type == ">=") {
          con_lhs[[n_con]] <<- rhs_val; con_rhs[[n_con]] <<- Inf
        } else {
          con_lhs[[n_con]] <<- rhs_val; con_rhs[[n_con]] <<- rhs_val
        }
      }

      # All-slot x columns for one person over a set of day indices — used by
      # the per-PP, fairness, and fix constraints below.
      x_slots <- function(pi, dis)
        xidx(pi, rep(dis, each = nS), rep.int(seq_len(nS), length(dis)))

      # ── Weekend shift columns ───────────────────────────────────────────────
      # The weekend block is Friday NIGHT through Sunday night, so weekend work
      # is a set of (day, slot) pairs, not a set of days: a Friday DAY shift is
      # weekday work, a Friday NIGHT shift is not. Everything that counts weekend
      # load (C11c bounds, C14 fairness, C_ws soft minimum, the Stage-2 fairness
      # pin) goes through this one definition. See is_weekend_shift() in
      # constants.R for the shared, display-side version.
      fri_di    <- which(weekdays(dates_vec) == "Friday")
      satsun_di <- which(weekdays(dates_vec) %in% c("Saturday", "Sunday"))
      wknd_cols_of <- function(pi)
        c(if (length(fri_di))    xidx(pi, fri_di, S_NIGHT) else integer(0),
          if (length(satsun_di)) x_slots(pi, satsun_di)    else integer(0))

      # ── C1: Slot uniqueness — Σ_p x[p,d,s] ≤ 1 ──────────────────────────────
      # Roaming and Night always need an explicit ≤1 row. APP1/APP2 normally do
      # NOT, because the C2/C2b equalities below pin them exactly.
      #
      # CRITICAL: in the green-only phase those equalities relax to ≥-style soft
      # coverage, which no longer bounds the slot from ABOVE — so APP1/APP2 must
      # join this loop or two people can share one slot on the same day.
      uniq_slots <- if (green_only) c(S_APP1, S_APP2, S_ROAM, S_NIGHT)
                    else            c(S_ROAM, S_NIGHT)
      for (d in seq_len(nD)) {
        for (s in uniq_slots) {
          add_con(xidx(seq_len(nP), d, s), rep(1, nP), "<=", 1)
        }
      }

      if (!green_only) {
        # ── C2: APP1 always filled — Σ_p x[p,d,APP1] = 1 ──────────────────────
        for (d in seq_len(nD))
          add_con(xidx(seq_len(nP), d, S_APP1), rep(1, nP), "=", 1)

        # ── C2b: APP2 always filled — Σ_p x[p,d,APP2] = 1 ─────────────────────
        for (d in seq_len(nD))
          add_con(xidx(seq_len(nP), d, S_APP2), rep(1, nP), "=", 1)
      } else {
        # ── C2/C2b (SOFT, green-only phase) ───────────────────────────────────
        # Exactly the C3 night-coverage form: uf[d] + Σ_p x[p,d,slot] ≥ 1 forces
        # the unfilled indicator to 1 when nobody covers the slot. Paired with the
        # ≤1 rows added above, and penalised -UNFILLED_APP{1,2}_PEN in the
        # objective. A day with fewer than two green volunteers simply cannot be
        # fully covered, and that hole is the intended output of this phase.
        for (d in seq_len(nD)) {
          add_con(c(uf1idx(d), xidx(seq_len(nP), d, S_APP1)),
                  rep(1, nP + 1L), ">=", 1)
          add_con(c(uf2idx(d), xidx(seq_len(nP), d, S_APP2)),
                  rep(1, nP + 1L), ">=", 1)
        }
      }

      # ── C3: Night coverage (SOFT) — nights may be unstaffed, penalised heavily ─
      # Nights are never hard-required. uns[d] + Σ_p x[p,d,Night] ≥ 1 forces the
      # unstaffed indicator uns[d] to 1 when no one covers the night (Σ = 0). C1
      # already caps Σ_p x[p,d,Night] ≤ 1 (≤ one person per night), and the
      # objective penalty −UNSTAFFED_NIGHT_PEN·uns[d] drives the solver to staff
      # every night it feasibly can.
      for (d in seq_len(nD)) {
        add_con(c(unsidx(d), xidx(seq_len(nP), d, S_NIGHT)),
                rep(1, nP + 1L), ">=", 1)
      }

      # (The old C4 no-double-book rows are gone: work[p,d] = Σ_s x[p,d,s] with
      #  ub(work) = 1 enforces the same thing through C_work below.)

      # ── C_work: work[p,d] = Σ_s x[p,d,s] ────────────────────────────────────
      for (pi in seq_len(nP)) {
        for (di in seq_len(nD)) {
          add_con(c(xidx(pi, di, seq_len(nS)), widx(pi, di)),
                  c(rep(1L, nS), -1L), "=", 0L)
        }
      }

      # ── C5: Availability — set ub = 0 for unavailable (p,d) ─────────────────
      # `unavail_mat` is blocked (off/vac/cme) in the fill phase, and
      # blocked-OR-yellow in the green-only phase. Enforcing green-only through
      # variable BOUNDS rather than a constraint row is both the cheapest form
      # and the one that cannot be subtly violated.
      for (pi in seq_len(nP)) {
        di_bad <- which(unavail_mat[pi, ])
        if (!length(di_bad)) next
        ub[xidx(pi, rep(di_bad, each = nS), rep.int(seq_len(nS), length(di_bad)))] <- 0
      }

      # NOTE on Roaming in the green-only phase: it IS filled here, so people get
      # their full pay-period allotment on days they asked for. Within a single
      # solve this is safe - the ILP optimises the whole horizon at once, and the
      # unfilled-slot penalties (100/80) dwarf ROAM_FILL_WEIGHT (1.5), so it will
      # never spend capacity on an optional Roaming shift that a mandatory slot
      # needs.
      #
      # It does make a subsequent FILL phase harder: C6 caps shifts per pay
      # period, so Roaming here consumes capacity the fill phase would otherwise
      # use to staff the holes this phase leaves behind. That is an accepted
      # trade-off for a green-only workflow.

      # ── C5b: No night shift the day before an off, vacation or CME day ───────
      for (pi in seq_len(nP)) {
        person <- STAFF[pi]
        pdata  <- self$time_off[[person]]
        if (nrow(pdata) == 0) next
        for (di in seq_len(nD - 1L)) {
          d_next <- dates_vec[di + 1L]
          # CME included: a night runs into the following morning, so a night
          # before a CME day would collide with the conference day itself.
          if (any(pdata$date == d_next & pdata$type %in% BLOCKED_TYPES))
            ub[xidx(pi, di, S_NIGHT)] <- 0
        }
      }

      # ── C5c: Prior-schedule boundary upper bounds ─────────────────────────────
      # Block assignments that would violate transition rules when a person worked
      # on the days immediately before the schedule starts.
      if (ps_has_prior) {
        for (pi in seq_len(nP)) {
          pre1 <- sched_day0
          pre2 <- sched_day0 - 1L
          had_night_pre1 <- pre1 %in% .ps_nights[[pi]]
          had_day_pre1   <- pre1 %in% .ps_days[[pi]]
          had_day_pre2   <- pre2 %in% .ps_days[[pi]]

          # C7/C7b boundary: Night on pre-day-1 → no Day slots on sched days 1 & 2
          if (had_night_pre1) {
            for (s in DAY_S) ub[xidx(pi, 1L, s)] <- 0
            if (nD >= 2L) for (s in DAY_S) ub[xidx(pi, 2L, s)] <- 0
          }
          # C8 boundary: Day on pre-day-1 → no Night on sched day 1
          if (had_day_pre1)
            ub[xidx(pi, 1L, S_NIGHT)] <- 0
          # C8b boundary: Day on pre-day-1 OR pre-day-2 → no Night on sched day 2
          if ((had_day_pre1 || had_day_pre2) && nD >= 2L)
            ub[xidx(pi, 2L, S_NIGHT)] <- 0
          # C11b boundary: Night on pre-day-1 → ss[p,1] = 0 (person continues existing run)
          if (had_night_pre1)
            ub[ssidx(pi, 1L)] <- 0
          # C15b boundary: Work on pre-day-1 → iso[p,1] = 0 (day 1 is anchored, not isolated)
          if ((pre1 %in% .ps_all[[pi]]) && nISO > 0L)
            ub[isoidx(pi, 1L)] <- 0
        }
      }

      # ── C6: PP shift cap — Σ_{d∈PP,s} x[p,d,s] ≤ sched_target - pp_cap_reduction ──
      for (pi in seq_len(nP)) {
        person  <- STAFF[pi]
        for (ppi in seq_len(nrow(PAY_PERIODS))) {
          pp_name <- PAY_PERIODS$name[ppi]
          pp_d    <- seq(PAY_PERIODS$start[ppi], PAY_PERIODS$end[ppi], by = "day")
          di_pp   <- which(dates_vec %in% pp_d)
          cap <- max(0L, self$targets[[person]][[pp_name]]$sched_target - pp_cap_reduction)
          if (length(di_pp) == 0L || cap <= 0L) next
          cols <- x_slots(pi, di_pp)
          add_con(cols, rep(1, length(cols)), "<=", cap)
        }
      }

      # ── C_min: Soft minimum shifts per PP for all staff ─────────────────────────
      # Uses soft_min (5 for non-flex, FLEX value for Todd) capped at the C6 ceiling.
      # PHASE GATE: skipped in the green-only phase. `sm` is capped by t_info$avail,
      # which counts green AND yellow days, so a person with 3 green days in a PP
      # would still be forced to 5 shifts -> infeasible. A floor is a statement
      # about a FINISHED schedule, and the green-only phase does not produce one.
      if (add_c_min && !green_only) {
        for (pi in seq_len(nP)) {
          person <- STAFF[pi]
          for (ppi in seq_len(nrow(PAY_PERIODS))) {
            pp_name <- PAY_PERIODS$name[ppi]
            pp_d    <- seq(PAY_PERIODS$start[ppi], PAY_PERIODS$end[ppi], by = "day")
            di_pp   <- which(dates_vec %in% pp_d)
            t_info  <- self$targets[[person]][[pp_name]]
            cap_eff <- max(0L, t_info$sched_target - pp_cap_reduction)
            sm <- min(t_info$soft_min, cap_eff, t_info$avail)
            if (sm <= 0L || length(di_pp) == 0L) next
            cols <- x_slots(pi, di_pp)
            add_con(cols, rep(1, length(cols)), ">=", sm)
          }
        }
      }

      # ── C7: Night→Day ban 1d — x[p,d,Night] + x[p,d+1,s] ≤ 1 ──────────────
      for (pi in seq_len(nP)) {
        for (di in seq_len(nD - 1L)) {
          ni <- xidx(pi, di, S_NIGHT)
          for (s in DAY_S)
            add_con(c(ni, xidx(pi, di + 1L, s)), c(1, 1), "<=", 1)
        }
      }

      # ── C7b: Night→Day ban 2d — x[p,d,Night] + x[p,d+2,s] ≤ 1 ─────────────
      # Enforces two full rest days between the end of any night shift and the
      # next day-shift assignment (eliminates Night-rest-APP patterns).
      for (pi in seq_len(nP)) {
        for (di in seq_len(nD - 2L)) {
          ni <- xidx(pi, di, S_NIGHT)
          for (s in DAY_S)
            add_con(c(ni, xidx(pi, di + 2L, s)), c(1, 1), "<=", 1)
        }
      }

      # ── C8: Day→Night gap 1d — x[p,d,s] + x[p,d+1,Night] ≤ 1  [LOCKED] ─────
      for (pi in seq_len(nP)) {
        for (di in seq_len(nD - 1L)) {
          ni <- xidx(pi, di + 1L, S_NIGHT)
          for (s in DAY_S)
            add_con(c(xidx(pi, di, s), ni), c(1, 1), "<=", 1)
        }
      }

      # ── C8b: Day→Night 2d gap — x[p,d,s] + x[p,d+2,Night] ≤ 1 ─────────────
      if (add_c8b) {
        for (pi in seq_len(nP)) {
          for (di in seq_len(nD - 2L)) {
            ni <- xidx(pi, di + 2L, S_NIGHT)
            for (s in DAY_S)
              add_con(c(xidx(pi, di, s), ni), c(1, 1), "<=", 1)
          }
        }
      }

      # ── C9: Max 3 consecutive nights — Σ_{k=0}^3 x[p,d+k,Night] ≤ 3  [LOCKED]
      # Hard cap, no exceptions: a 4-night run is infeasible (not merely penalised).
      for (pi in seq_len(nP)) {
        for (di in seq_len(nD - 3L)) {
          add_con(xidx(pi, di + 0:3, S_NIGHT), rep(1, 4L), "<=", 3)
        }
      }

      # ── C9b: Nights must be stacked — no single-day gap  [LOCKED] ────────────
      #   x[p,d,N] + x[p,d+2,N] - x[p,d+1,N] <= 1
      # Bans the Night-Empty-Night pattern: if a person works nights on d and
      # d+2, then d+1 must be a night too, making it one block rather than two
      # sleep-cycle flips a day apart.
      #
      # Reachable without this rule: C7/C7b already stop a DAY shift on the two
      # days after a night, so the day after a night is either another night or
      # empty - and nothing previously stopped a night resuming the day after
      # that. Truth table:
      #   N N N -> 1+1-1 = 1  ok (one 3-night block)
      #   N _ N -> 1+1-0 = 2  banned
      #   N _ _ -> 1+0-0 = 1  ok
      #   _ _ N -> 0+1-0 = 1  ok
      for (pi in seq_len(nP)) {
        for (di in seq_len(nD - 2L)) {
          add_con(xidx(pi, di + c(0L, 2L, 1L), S_NIGHT), c(1, 1, -1), "<=", 1)
        }
      }

      # ── C9bx: C9b across the prior-schedule boundary ─────────────────────────
      # A night on the last day before the schedule, with day 1 empty, must not
      # be followed by a night on day 2.
      if (ps_has_prior && nD >= 2L) {
        for (pi in seq_len(nP)) {
          if (sched_day0 %in% .ps_nights[[pi]])
            add_con(xidx(pi, c(2L, 1L), S_NIGHT), c(1, -1), "<=", 0)
        }
      }

      # ── C9x: Cross-boundary max-3-consecutive nights ─────────────────────────
      # For windows that span prior-schedule nights into sched days 1..3.
      # k prior days in window → sched nights in remaining (4-k) days ≤ 3 - pre_count.
      if (ps_has_prior) {
        for (pi in seq_len(nP)) {
          for (k in 1L:min(3L, nD)) {
            pre_dates <- seq(sched_day0 - k + 1L, sched_day0, by = 1L)
            pre_count <- as.integer(sum(pre_dates %in% .ps_nights[[pi]]))
            if (pre_count == 0L) next
            j_max <- 4L - k
            if (j_max < 1L) next
            cols <- xidx(pi, seq_len(j_max), S_NIGHT)
            add_con(cols, rep(1L, length(cols)), "<=", max(0L, 3L - pre_count))
          }
        }
      }

      # ── C10: Max 4 consecutive working days — Σ_{k=0}^4 work[p,d+k] ≤ 4  [LOCKED]
      for (pi in seq_len(nP)) {
        for (di in seq_len(nD - 4L)) {
          add_con(widx(pi, di + 0:4), rep(1L, 5L), "<=", 4L)
        }
      }

      # ── C10x: Cross-boundary max-4-consecutive work days ─────────────────────
      # Mirrors C9x but for any shift type (work[p,d]).
      if (ps_has_prior) {
        for (pi in seq_len(nP)) {
          for (k in 1L:min(4L, nD)) {
            pre_dates <- seq(sched_day0 - k + 1L, sched_day0, by = 1L)
            pre_count <- as.integer(sum(pre_dates %in% .ps_all[[pi]]))
            if (pre_count == 0L) next
            j_max <- 5L - k
            if (j_max < 1L) next
            add_con(widx(pi, seq_len(j_max)),
                    rep(1L, j_max), "<=", max(0L, 4L - pre_count))
          }
        }
      }

      # ── C10b: No 3 consecutive Saturday or Sunday shifts per person ───────────
      # "Consecutive" means the k-th, (k+1)-th, (k+2)-th occurrence of that
      # weekday in the schedule window — each 7 days apart.
      for (wday in c("Saturday", "Sunday")) {
        wday_idx <- which(weekdays(dates_vec) == wday)
        if (length(wday_idx) >= 3L) {
          for (pi in seq_len(nP)) {
            for (k in seq_len(length(wday_idx) - 2L)) {
              add_con(widx(pi, wday_idx[k:(k + 2L)]), rep(1L, 3L), "<=", 2L)
            }
          }
        }
      }

      # ── C10c: Multi-window density caps  [LOCKED] ────────────────────────────
      # ≤5 in any 7-day window, ≤5 in any 8-day window,
      # ≤6 in any 9-day window, ≤7 in any 10/11-day window.
      density_windows <- list(
        list(W = 7L,  B = 5L),
        list(W = 8L,  B = 5L),
        list(W = 9L,  B = 6L),
        list(W = 10L, B = 7L),
        list(W = 11L, B = 7L)
      )
      for (dw in density_windows) {
        W <- dw$W; B <- dw$B
        if (nD < W) next
        for (pi in seq_len(nP)) {
          for (di in seq_len(nD - W + 1L)) {
            add_con(widx(pi, di + 0:(W - 1L)), rep(1L, W), "<=", B)
          }
        }
        if (ps_has_prior && nD >= 1L) {
          for (pi in seq_len(nP)) {
            for (k in 1L:min(W - 1L, nD)) {
              pre_dates <- seq(sched_day0 - k + 1L, sched_day0, by = 1L)
              pre_count <- as.integer(sum(pre_dates %in% .ps_all[[pi]]))
              if (pre_count == 0L) next
              j_max <- W - k
              if (j_max < 1L) next
              add_con(widx(pi, seq_len(j_max)),
                      rep(1L, j_max), "<=", max(0L, B - pre_count))
            }
          }
        }
      }

      # (The old C11 blanket cap Σ nights ≤ MAX_NIGHTS_TOTAL is gone — C11d
      #  below always carries a night_max_hard ≤ MAX_NIGHTS_TOTAL, so the row
      #  was redundant in every tier.)

      # ── C11b: No night stretch in consecutive PPs ────────────────────────────
      # ss[p,d] is a stretch-start indicator: 1 iff person p starts a new night
      # stretch on day d (x[p,d,Night]=1 AND x[p,d-1,Night]=0).
      # Definition constraints (hold for all d; treat x[p,0,Night] = 0):
      #   ss[p,d] ≤ x[p,d,N]
      #   ss[p,d] ≤ 1 - x[p,d-1,N]   (= 1 when d=1, implicit via ub)
      #   ss[p,d] ≥ x[p,d,N] - x[p,d-1,N]   (= x[p,d,N] when d=1)
      # PP-consecutive rule: Σ_{d∈PP_k ∪ PP_{k+1}} ss[p,d] ≤ 1  for k=1..nPP-1
      if (add_c11b) {
        for (pi in seq_len(nP)) {
          for (di in seq_len(nD)) {
            si <- ssidx(pi, di)
            ni <- xidx(pi, di, S_NIGHT)
            add_con(c(si, ni), c(1, -1), "<=", 0)          # ss <= x[d,N]
            if (di == 1L) {
              # Skip when person continued from a prior-night run (ub[ss]=0 set in C5c).
              if (!ps_has_prior || !(sched_day0 %in% .ps_nights[[pi]]))
                add_con(c(si, ni), c(1, -1), ">=", 0)      # ss >= x[1,N]
            } else {
              ni_p <- xidx(pi, di - 1L, S_NIGHT)
              add_con(c(si, ni, ni_p), c(1, -1, 1), ">=", 0) # ss >= x[d,N]-x[d-1,N]
              add_con(c(si, ni_p),     c(1,  1),    "<=", 1)  # ss <= 1-x[d-1,N]
            }
          }
        }
        for (pi in seq_len(nP)) {
          for (k in seq_len(nrow(PAY_PERIODS) - 1L)) {
            dk_idx  <- which(dates_vec >= PAY_PERIODS$start[k]   & dates_vec <= PAY_PERIODS$end[k])
            dk1_idx <- which(dates_vec >= PAY_PERIODS$start[k+1L] & dates_vec <= PAY_PERIODS$end[k+1L])
            cols    <- ssidx(pi, c(dk_idx, dk1_idx))
            add_con(cols, rep(1L, length(cols)), "<=", 1L)
          }
        }
      }

      # ── C11c: Weekend hard bounds — MIN_WKND_HARD ≤ Σ_{d∈Sat/Sun} work[p,d] ≤ MAX_WKND_HARD
      if (add_c11c) {
        if (length(fri_di) + length(satsun_di) > 0L) {
          for (pi in seq_len(nP)) {
            cols <- wknd_cols_of(pi)
            # PHASE GATE: the LOWER bound is a floor, dropped in the green-only
            # phase. Measured: enforcing it there makes the phase infeasible at
            # tier 1 (a person's green weekend days are simply too few), so it
            # fell through to a heavily relaxed tier and produced a far worse
            # schedule. The fill phase enforces it instead. The ceiling always
            # applies.
            if (!green_only)
              add_con(cols, rep(1L, length(cols)), ">=", MIN_WKND_HARD)
            add_con(cols, rep(1L, length(cols)), "<=", MAX_WKND_HARD)
          }
        }
      }

      # ── C11d: Night hard bounds — night_min_hard ≤ Σ_d x[p,d,Night] ≤ night_max_hard
      if (!is.null(night_min_hard) || !is.null(night_max_hard)) {
        for (pi in seq_len(nP)) {
          cols <- xidx(pi, seq_len(nD), S_NIGHT)
          # PHASE GATE: same as C11c. The night FLOOR is unsatisfiable on green
          # days alone - a 3-night stretch needs 3 consecutive green days, which
          # scattered requests rarely provide. Enforcing it in the green-only
          # phase was measured to make tier 1 infeasible.
          if (!is.null(night_min_hard) && !green_only)
            add_con(cols, rep(1L, nD), ">=", night_min_hard)
          if (!is.null(night_max_hard))
            add_con(cols, rep(1L, nD), "<=", night_max_hard)
        }
      }

      # ── C12: Holiday pre-seeds — fix x[p,d,s] = 1 ────────────────────────────
      # GREEN-ONLY BEHAVIOUR: a holiday pin is applied only when the day is
      # actually green for that person. Otherwise the slot is RESERVED - ub = 0
      # for everyone - and left as a hole for the fill phase to pin.
      #
      # Forcing the pin regardless (the obvious alternative) breaks C11b. The
      # holidays assign the same person two nights in ADJACENT pay periods
      # (2026-05-25 in PP11 and 2026-06-19 in PP12). C11b allows at most one
      # night-stretch start across any two adjacent PPs, and the only escape is
      # to extend a stretch backwards into the previous PP - which needs the day
      # before to be available. The fill phase has yellow days for that; the
      # green-only phase does not, so the pins are simply deferred.
      #
      # Reserving matters: if the green-only phase gave a holiday slot to someone
      # ELSE, the fill phase would pin two people into one slot and C2 would make
      # the model infeasible.
      slot_idx <- c(APP1 = S_APP1, APP2 = S_APP2, Roaming = S_ROAM, Night = S_NIGHT)
      n_defer <- 0L
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
              any(pdata$date == d_hol & pdata$type %in% BLOCKED_TYPES)) next
          if (green_only && !green_mat[pi[[1L]], di[[1L]]]) {
            # Defer to the fill phase and reserve the slot for the named person.
            ub[xidx(seq_len(nP), di[[1L]], si)] <- 0
            n_defer <- n_defer + 1L
            next
          }
          idx    <- xidx(pi[[1L]], di[[1L]], si)
          lb[idx] <- 1
          ub[idx] <- 1
        }
      }
      if (green_only && n_defer > 0L)
        message(sprintf(
          "    %d holiday slot(s) deferred to the fill phase (not green for the named person).",
          n_defer))

      # ── C14: Fairness min/max bounds ─────────────────────────────────────────
      if (add_c14) {
        for (pi in seq_len(nP)) {
          night_cols <- xidx(pi, seq_len(nD), S_NIGHT)
          all_cols   <- x_slots(pi, seq_len(nD))

          add_con(c(night_cols, I_MAX_NIGHTS), c(rep(1, nD), -1L),      "<=", 0)
          add_con(c(night_cols, I_MIN_NIGHTS), c(rep(1, nD), -1L),      ">=", 0)
          add_con(c(all_cols,   I_MAX_TOTAL),  c(rep(1, nD * nS), -1L), "<=", 0)
          add_con(c(all_cols,   I_MIN_TOTAL),  c(rep(1, nD * nS), -1L), ">=", 0)

          if (length(fri_di) + length(satsun_di) > 0) {
            wknd_cols <- wknd_cols_of(pi)
            add_con(c(wknd_cols, I_MAX_WKND), c(rep(1, length(wknd_cols)), -1L), "<=", 0)
            add_con(c(wknd_cols, I_MIN_WKND), c(rep(1, length(wknd_cols)), -1L), ">=", 0)
          }
        }
      }

      # ── C_fix: Stage-2 fairness pin ─────────────────────────────────────────
      # Lock each person's night / total / weekend shift counts to the Stage-1
      # fairness solution (within ±fix_band). Stage 2 then optimises aesthetics
      # only among schedules that match Stage 1's fairness. Always feasible: the
      # Stage-1 assignment itself satisfies these by construction.
      if (stage == 2L && !is.null(fix_targets)) {
        add_fix <- function(cols, tgt) {
          if (fix_band <= 0L) {
            add_con(cols, rep(1L, length(cols)), "=", tgt)
          } else {
            add_con(cols, rep(1L, length(cols)), "<=", tgt + fix_band)
            add_con(cols, rep(1L, length(cols)), ">=", max(0L, tgt - fix_band))
          }
        }
        for (pi in seq_len(nP)) {
          add_fix(xidx(pi, seq_len(nD), S_NIGHT), fix_targets$night[pi])
          add_fix(x_slots(pi, seq_len(nD)),       fix_targets$total[pi])
          if (length(fri_di) + length(satsun_di) > 0)
            add_fix(wknd_cols_of(pi), fix_targets$weekend[pi])
        }
      }

      # ── C15b: Isolation definition (objective penalty on isolated shifts) ────
      # iso[p,d] is forced to 1 when work[p,d]=1 and both neighbours=0.
      #   iso[p,d] >= work[p,d] - work[p,d-1] - work[p,d+1]  (lower bound)
      #   iso[p,d] <= work[p,d]                                (zero when not working)
      # The -8.0 objective penalty discourages isolated shifts (no hard cap).
      if (nISO > 0L) {
        for (pi in seq_len(nP)) {
          # Left boundary (d=1, no left neighbour)
          # Skip when person worked pre-day-1 (ub[iso[p,1]]=0 set in C5c; LB trivially satisfied).
          if (!ps_has_prior || !(sched_day0 %in% .ps_all[[pi]]))
            add_con(c(isoidx(pi, 1L), widx(pi, 1L), widx(pi, 2L)),
                    c(-1L, 1L, -1L), "<=", 0L)
          # Interior days
          if (nD >= 3L) {
            for (di in 2L:(nD - 1L)) {
              add_con(c(isoidx(pi, di), widx(pi, di - 1L), widx(pi, di), widx(pi, di + 1L)),
                      c(-1L, -1L, 1L, -1L), "<=", 0L)
            }
          }
          # Right boundary (d=nD, no right neighbour)
          add_con(c(isoidx(pi, nD), widx(pi, nD), widx(pi, nD - 1L)),
                  c(-1L, 1L, -1L), "<=", 0L)
          # Upper bound: iso <= work
          for (di in seq_len(nD))
            add_con(c(isoidx(pi, di), widx(pi, di)), c(1L, -1L), "<=", 0L)
        }
      }

      # ── C_n3: Night 3-block — n3[p,d] ≥ Σ_{k=0}^2 x[p,d+k,N] − 2 ────────────
      if (nNS3 > 0L) {
        for (pi in seq_len(nP)) {
          for (di in seq_len(nD - 2L)) {
            add_con(c(xidx(pi, di, S_NIGHT), xidx(pi, di + 1L, S_NIGHT),
                      xidx(pi, di + 2L, S_NIGHT), n3idx(pi, di)),
                    c(1, 1, 1, -1), "<=", 2)
          }
        }
      }

      # ── C_n4: Night 4-block — n4[p,d] ≥ Σ_{k=0}^3 x[p,d+k,N] − 3 ────────────
      if (nNS4 > 0L) {
        for (pi in seq_len(nP)) {
          for (di in seq_len(nD - 3L)) {
            add_con(c(xidx(pi, di, S_NIGHT), xidx(pi, di + 1L, S_NIGHT),
                      xidx(pi, di + 2L, S_NIGHT), xidx(pi, di + 3L, S_NIGHT),
                      n4idx(pi, di)),
                    c(1, 1, 1, 1, -1), "<=", 3)
          }
        }
      }

      # ── C_w3: Work 3-block — w3[p,d] ≥ Σ_{k=0}^2 work[p,d+k] − 2 ───────────
      if (nWS3 > 0L) {
        for (pi in seq_len(nP)) {
          for (di in seq_len(nD - 2L)) {
            add_con(c(widx(pi, di), widx(pi, di + 1L), widx(pi, di + 2L),
                      w3idx(pi, di)),
                    c(1, 1, 1, -1), "<=", 2)
          }
        }
      }

      # ── C_w4: Work 4-block — w4[p,d] ≥ Σ_{k=0}^3 work[p,d+k] − 3 ───────────
      if (nWS4 > 0L) {
        for (pi in seq_len(nP)) {
          for (di in seq_len(nD - 3L)) {
            add_con(c(widx(pi, di), widx(pi, di + 1L), widx(pi, di + 2L),
                      widx(pi, di + 3L), w4idx(pi, di)),
                    c(1, 1, 1, 1, -1), "<=", 3)
          }
        }
      }


      # ── C_ns: Night soft-minimum — ns_short[p] + Σ_d x[p,d,Night] ≥ MIN ────────
      # ns_short[p] = max(0, MIN_NIGHTS_SOFT_TOTAL - actual nights).
      # Combined with lb=0 and ub=MIN, this forces the solver to "pay" a penalty
      # for every night below the soft floor.
      for (pi in seq_len(nP)) {
        add_con(c(nsshidx(pi), xidx(pi, seq_len(nD), S_NIGHT)),
                rep(1L, nD + 1L), ">=", MIN_NIGHTS_SOFT_TOTAL)
      }

      # ── C_ws: Weekend soft-minimum — ws_short[p] + Σ_{d=Sat/Sun} work[p,d] ≥ MIN
      for (pi in seq_len(nP)) {
        wknd_cols <- wknd_cols_of(pi)
        add_con(c(wsshidx(pi), wknd_cols),
                rep(1L, length(wknd_cols) + 1L), ">=", MIN_WKND_SOFT_TOTAL)
      }

      # ── C_ymax: Yellow-day fairness — Σ_{d in yellow(p)} work[p,d] ≤ I_MAX_YELLOW
      # Mirrors C14's I_MAX_NIGHTS linking. I_MAX_YELLOW carries -MAX_YELLOW_PEN in
      # the objective, so maximising drives it down to the largest per-person
      # yellow count and the solver flattens the WORST person's share rather than
      # dumping every unavoidable yellow day on whoever is cheapest to move.
      #
      # NOT gated on aesthetic_on - this must be live in stage 1, which is the
      # production path. Only the green-only phase skips it (no yellow days exist
      # there, so every row would be vacuous).
      if (!green_only) {
        for (pi in seq_len(nP)) {
          yd <- which(yellow_mat[pi, ])
          if (!length(yd)) next
          ycols <- widx(pi, yd)
          add_con(c(ycols, I_MAX_YELLOW), c(rep(1, length(ycols)), -1L), "<=", 0)
        }
      }

      # ── C_wsp: weekend-split linearisation ────────────────────────────────────
      # wsp[p,w] >= work[p,sat] - work[p,sun]  and  >= work[p,sun] - work[p,sat],
      # i.e. wsp >= |sat - sun|. With a negative objective weight the maximiser
      # pushes wsp down to that bound, so it equals 1 exactly when the person
      # works one day of the pair and not the other.
      #
      # NOT gated on aesthetic_on - stage 1 is the production path.
      if (nWK > 0L) {
        for (pi in seq_len(nP)) {
          for (w in seq_len(nWK)) {
            sd <- wk_pair[w]; ud <- sd + 1L
            add_con(c(wspidx(pi, w), widx(pi, sd), widx(pi, ud)),
                    c(1, -1,  1), ">=", 0)
            add_con(c(wspidx(pi, w), widx(pi, sd), widx(pi, ud)),
                    c(1,  1, -1), ">=", 0)
          }
        }
      }

      # ── C_pin: force the green-only phase's assignments (FILL phase) ──────────
      # Same lb/ub mechanism C12 uses for holidays. A shift granted on a requested
      # day is never taken back by the fill phase, which is what makes the
      # two-button workflow explainable: what you saw in phase 1 is what you keep.
      if (!is.null(pinned) && nrow(pinned) > 0L) {
        slot_num <- c(APP1 = S_APP1, APP2 = S_APP2, Roaming = S_ROAM, Night = S_NIGHT)
        n_pin <- 0L
        for (i in seq_len(nrow(pinned))) {
          pi_ <- which(STAFF == pinned$person[i])
          di_ <- which(dates_vec == as.Date(pinned$date[i]))
          si_ <- slot_num[[as.character(pinned$slot[i])]]
          if (!length(pi_) || !length(di_) || is.null(si_) || is.na(si_)) next
          idx <- xidx(pi_[1L], di_[1L], si_)
          if (identical(pin_mode, "soft")) {
            # Preference, not a constraint: reward keeping this exact assignment.
            obj[idx] <- obj[idx] + PIN_KEEP_BONUS
          } else {
            lb[idx] <- 1; ub[idx] <- 1
          }
          n_pin <- n_pin + 1L
        }
        message(sprintf("    %s %d assignment(s) from the green-only phase.",
                        if (identical(pin_mode, "soft")) "Preferring" else "Pinned", n_pin))
      }

      # ── Extra no-good cuts (solution enumeration) ─────────────────────────────
      for (x_prev in extra_nogo) {
        on_idx <- which(x_prev > 0.5)
        if (length(on_idx) > 0L)
          add_con(as.integer(on_idx), rep(1, length(on_idx)), "<=", length(on_idx) - 1L)
      }

      # ── Assemble and solve ────────────────────────────────────────────────────
      nCont <- nV - nX
      stage_lbl <- switch(as.character(stage),
                          "1" = " [Stage 1: fairness]",
                          "2" = " [Stage 2: aesthetics]",
                          "")
      message(sprintf("  ILP%s: %d binary + %d continuous, %d constraints (%d thread%s)",
                      stage_lbl, nX, nCont, n_con, SOLVER_THREADS,
                      if (SOLVER_THREADS == 1L) "" else "s"))

      A <- Matrix::sparseMatrix(i = unlist(ri_l), j = unlist(ci_l),
                                x = unlist(vi_l), dims = c(n_con, nV))

      result <- highs::highs_solve(
        L       = obj,
        lower   = lb,
        upper   = ub,
        A       = A,
        lhs     = con_lhs,
        rhs     = con_rhs,
        types   = types,
        maximum = TRUE,
        control = highs::highs_control(
                    # Every solve is capped: SOLVER_TIME_LIMIT by default, or the
                    # caller's override (count_solutions passes a short one).
                    time_limit           = if (!is.null(time_limit_override))
                                             time_limit_override else SOLVER_TIME_LIMIT,
                    # Stage 1 (fairness) solves to a tighter gap so the pinned
                    # fairness target is near-optimal; other stages use the looser gap.
                    # Green-only phase: stop on an ABSOLUTE gap. Its objective goes
                    # negative once any slot is unfilled, where a relative gap is
                    # not a meaningful stopping rule (see SOLVER_MIP_ABS_GAP_GREEN).
                    mip_rel_gap          = if (green_only) 0
                                           else if (stage == 1L) SOLVER_MIP_GAP_STAGE1
                                           else SOLVER_MIP_GAP,
                    mip_abs_gap          = if (green_only) SOLVER_MIP_ABS_GAP_GREEN else 1e-6,
                    threads              = SOLVER_THREADS,
                    parallel             = "on",
                    # Spend more effort on feasibility heuristics up front so a
                    # first incumbent appears sooner (substitute for a warm start).
                    mip_heuristic_effort = SOLVER_HEURISTIC_EFFORT,
                    presolve             = "on",
                    # This model is highly symmetric (interchangeable person-days);
                    # symmetry detection prunes large swaths of the search tree.
                    mip_detect_symmetry  = "on")
      )

      sol <- result$primal_solution
      if (is.null(sol)) {
        message(sprintf("  Solver: %s", result$status_message))
        return(NULL)
      }
      # Reject LP-relaxation pseudo-solutions (returned when solver times out before
      # finding any integer feasible point — all x values are fractional, round to 0).
      # A valid schedule always has at least one x=1 because C2 mandates APP1 daily.
      if (sum(round(sol[seq_len(nX)])) == 0L) {
        message(sprintf("  Solver: %s (no integer solution found — skipping tier)",
                        result$status_message))
        return(NULL)
      }
      message(sprintf("  Solver: %s  objective = %.1f%s",
                      result$status_message,
                      if (!is.null(result$objective_value)) result$objective_value else NA_real_,
                      {
                        ub <- result$info$mip_dual_bound
                        if (!is.null(ub) && is.finite(ub)) sprintf("  (max: %.1f)", ub) else ""
                      }))

      list(sol = sol, nP = nP, nD = nD, nS = nS, nX = nX, xidx = xidx, dates_vec = dates_vec)
    },

    # ── Translate binary solution vector into schedule data structures ──────────
    # Vectorised: decode person/day/slot straight from the flat indices of the
    # x variables set to 1 (no per-variable loop, no rbind-in-loop).
    populate_from_solution = function(res) {
      sol       <- res$sol
      nP        <- res$nP
      nD        <- res$nD
      nS        <- res$nS
      dates_vec <- res$dates_vec

      SLOT_NAMES <- c("APP1", "APP2", "Roaming", "Night")

      xs  <- sol[seq_len(res$nX)]
      xs[is.na(xs)] <- 0
      on  <- which(round(xs) == 1L)
      # Invert the flat index: idx = (p-1)*nD*nS + (d-1)*nS + s
      s_i <- (on - 1L) %% nS + 1L
      d_i <- ((on - 1L) %/% nS) %% nD + 1L
      p_i <- (on - 1L) %/% (nD * nS) + 1L
      pps <- get_pp_vec(dates_vec)

      for (k in seq_along(on)) {
        person <- STAFF[p_i[k]]
        ds     <- as.character(dates_vec[d_i[k]])
        self$schedule[[ds]][[SLOT_NAMES[s_i[k]]]] <- person
        pp <- pps[d_i[k]]
        if (!is.na(pp))
          self$pp_counts[[person]][[pp]] <- self$pp_counts[[person]][[pp]] + 1L
      }
      for (pi in seq_len(nP)) {
        person  <- STAFF[pi]
        mine    <- p_i == pi
        n_days  <- sort(dates_vec[d_i[mine & s_i == 4L]])
        day_sel <- which(mine & s_i != 4L)
        day_sel <- day_sel[order(dates_vec[d_i[day_sel]])]
        self$person_nights[[person]] <- n_days
        self$person_shifts[[person]] <- data.frame(
          date = dates_vec[d_i[day_sel]],
          slot = SLOT_NAMES[s_i[day_sel]],
          stringsAsFactors = FALSE)
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

    # ── Diagnostics: run after all tiers fail, then stop with error ────────────
    report_and_stop = function() {
      dates_vec <- as.Date(self$dates, origin = "1970-01-01")
      nP        <- length(STAFF)
      nD        <- length(dates_vec)

      issues <- character(0)

      # 1. Per-day: count available staff; flag days where < 2 are free
      for (di in seq_len(nD)) {
        d       <- dates_vec[di]
        n_avail <- sum(vapply(STAFF, function(p) !private$is_blocked(p, d), logical(1L)))
        if (n_avail < 3L) {
          issues <- c(issues, sprintf(
            "[COVERAGE] %s (%s): only %d/%d staff available — need >= 3 for APP1+APP2+Night",
            format(d), weekdays(d, abbreviate = TRUE), n_avail, nP))
        }
      }

      # 2. Per-person per-PP: flag where sched_target > available days
      for (person in STAFF) {
        for (ppi in seq_len(nrow(PAY_PERIODS))) {
          pp_name  <- PAY_PERIODS$name[ppi]
          pp_d     <- seq(PAY_PERIODS$start[ppi], PAY_PERIODS$end[ppi], by = "day")
          pp_d     <- pp_d[pp_d %in% dates_vec]
          if (length(pp_d) == 0L) next
          n_blocked <- sum(vapply(pp_d, function(d) private$is_blocked(person, d), logical(1L)))
          n_avail   <- length(pp_d) - n_blocked
          target    <- self$targets[[person]][[pp_name]]$sched_target
          if (!is.null(target) && target > n_avail) {
            issues <- c(issues, sprintf(
              "[TARGET]   %s / %s: sched_target=%d but only %d/%d days free (blocked=%d)",
              person, pp_name, target, n_avail, length(pp_d), n_blocked))
          }
        }
      }

      # 3. Per-person: flag anyone with no window of >= 2 consecutive free days
      for (person in STAFF) {
        free    <- vapply(dates_vec, function(d) !private$is_blocked(person, d), logical(1L))
        max_run <- 0L; run <- 0L
        for (f in free) {
          if (f) { run <- run + 1L; if (run > max_run) max_run <- run } else run <- 0L
        }
        if (max_run < 2L) {
          issues <- c(issues, sprintf(
            "[NIGHTS]   %s: no window of 2+ consecutive free days — night runs impossible",
            person))
        }
      }

      message("\n  ─── ILP SCHEDULING DIAGNOSTICS ────────────────────────────")
      if (length(issues) == 0L) {
        message("  No obvious availability or target issues detected.")
        message(sprintf("  The model may simply be too complex for the %g-second time budget.", SOLVER_TIME_LIMIT))
        message("  Consider loosening the night bands, loosening PP targets, or")
        message("  removing time-off entries that conflict with coverage requirements.")
      } else {
        message(sprintf("  %d potential issue(s) found:\n", length(issues)))
        for (iss in issues) message(sprintf("  %s", iss))
      }
      message("  ────────────────────────────────────────────────────────────\n")

      stop("ILP scheduling failed after all relaxation tiers — see diagnostics above.",
           call. = FALSE)
    },

    # ── Greedy aesthetic pass (Stage-2 replacement) ────────────────────────────
    # Local hill-climb over the Stage-1 schedule that maximises a run-length score
    # (solo = -10, 2-run = +3, 3-run = +8, 4-run = +5). It relocates a shift from a
    # short run (length 1-2) onto a day adjacent to one of that person's other
    # shifts via a same-slot / same-weekend-status / same-pay-period swap with
    # whoever holds the target day+slot, accepting only swaps that strictly raise
    # the total score. Every swap preserves each person's night, weekend, total and
    # per-PP counts EXACTLY (Stage-1 fairness untouched) and is rejected unless both
    # people remain feasible against all per-person hard rules.
    greedy_aesthetic_pass = function(max_passes = 6L, green_only = FALSE,
                                     pinned = NULL) {
      dates_vec  <- sort(as.Date(names(self$schedule)))
      slot_names <- c("APP1", "APP2", "Roaming", "Night")

      # NULL/NA/length-safe slot occupant lookup.
      occ <- function(day, s) { v <- day[[s]]; if (length(v) != 1L || is.na(v)) NA_character_ else v }

      # Per-person assignment: named char vector  date-string -> slot.
      # NB: iterate Date vectors by INDEX — `for (d in dates_vec)` strips the Date
      # class (d becomes a raw numeric), breaking as.character(d).
      asgn <- setNames(lapply(STAFF, function(p) {
        v <- character(0)
        for (di in seq_along(dates_vec)) {
          d <- dates_vec[di]; ds <- as.character(d); day <- self$schedule[[ds]]
          for (s in slot_names) if (identical(occ(day, s), p)) { v[ds] <- s; break }
        }
        v
      }), STAFF)

      # Run-length aesthetic score: decompose a person's consecutive worked-day
      # runs and score each run (mirrors the former ILP aesthetic objective):
      #   length 1 (solo) = -10 ; 2 = +3 ; 3 = +8 ; 4 = +5  (runs >4 barred by C10)
      run_pts   <- c(-10, 3, 8, 5)
      run_score <- function(vec) {
        if (length(vec) == 0L) return(0)
        ints <- sort(as.integer(as.Date(names(vec)))); n <- length(ints)
        if (n == 0L) return(0)
        total <- 0; rl <- 1L
        for (i in seq_len(n)[-1]) {
          if (ints[i] == ints[i - 1] + 1L) rl <- rl + 1L
          else { total <- total + (if (rl <= 4L) run_pts[rl] else 5); rl <- 1L }
        }
        total + (if (rl <= 4L) run_pts[rl] else 5)
      }
      # Move sources: dates in a short run (length 1 or 2) — the shifts worth
      # relocating to grow a longer (higher-scoring) run.
      short_dates <- function(vec) {
        if (length(vec) == 0L) return(as.Date(character()))
        dts <- as.Date(names(vec)); ord <- order(dts); si <- as.integer(dts[ord])
        keep <- logical(length(si)); i <- 1L
        while (i <= length(si)) {
          j <- i; while (j < length(si) && si[j + 1L] == si[j] + 1L) j <- j + 1L
          if ((j - i + 1L) <= 2L) keep[i:j] <- TRUE
          i <- j + 1L
        }
        dts[ord][keep]
      }
      # Weekend status is a property of (date, SLOT) now - Friday night counts,
      # Friday day does not. Swaps are same-slot, so the slot is carried in.
      is_wknd   <- function(d, s) is_weekend_shift(d, s)
      pp_of     <- function(d) private$pp_index(d)

      # ── Requested-work ("green") days, as date-string sets per person ────────
      green_key <- setNames(lapply(STAFF, function(p) {
        df <- self$time_off[[p]]
        if (is.null(df) || !nrow(df)) character(0)
        else as.character(df$date[df$type == "green"])
      }), STAFF)

      # In the green-only phase EVERY schedulable day is green, so a green term
      # would be a constant and only the run-length score matters. In the fill
      # phase it decides whether a swap is worth making.
      green_score <- function(P, vec) {
        if (green_only || !length(vec)) return(0)
        GREEN_SWAP_PTS * sum(names(vec) %in% green_key[[P]])
      }
      total_score <- function(P, vec) run_score(vec) + green_score(P, vec)

      # ── Days this pass may not move a shift ONTO ─────────────────────────────
      # The green-only phase must never place anyone on a non-green day, so its
      # swap gate is stricter than plain availability.
      cannot_work <- function(person, d)
        private$is_blocked(person, d) ||
        (green_only && !private$is_green(person, d))

      # ── Assignments this pass must not disturb ──────────────────────────────
      # (a) HOLIDAY pre-seeds. C12 fixes these with lb = ub = 1: they are a
      #     management input, not a solver choice. The pass used to swap them
      #     away silently - a genuine bug, invisible because nothing re-checked
      #     holidays afterwards. It also made the post-pass schedule ILP-
      #     INFEASIBLE, which the two-phase fill surfaced: pinning a schedule
      #     whose holiday slots had been reassigned contradicts C12.
      # (b) Assignments the fill phase inherited from the green-only phase.
      pin_key <- character(0)
      for (ds in names(HOLIDAYS)) {
        d_hol <- as.Date(ds)
        if (d_hol < SCHEDULE_START || d_hol > SCHEDULE_END) next
        for (sl in names(HOLIDAYS[[ds]]))
          pin_key <- c(pin_key, paste(HOLIDAYS[[ds]][[sl]], ds, sl))
      }
      if (!is.null(pinned) && nrow(pinned) > 0L)
        pin_key <- c(pin_key, paste(pinned$person, as.character(pinned$date), pinned$slot))
      pin_key   <- unique(pin_key)
      is_pinned <- function(person, dk, s)
        length(pin_key) > 0L && paste(person, dk, s) %in% pin_key

      # A holiday SLOT is frozen regardless of who currently holds it: swapping a
      # different person into it would displace the designated one just as badly.
      hol_slot_key <- character(0)
      for (ds in names(HOLIDAYS)) {
        d_hol <- as.Date(ds)
        if (d_hol < SCHEDULE_START || d_hol > SCHEDULE_END) next
        for (sl in names(HOLIDAYS[[ds]]))
          hol_slot_key <- c(hol_slot_key, paste(ds, sl))
      }
      is_holiday_slot <- function(dk, s) paste(dk, s) %in% hol_slot_key


      # ── Shared candidate filter ─────────────────────────────────────────────
      # Every swap stays SAME-SLOT, SAME-WEEKEND-STATUS and SAME-PAY-PERIOD, which
      # is what preserves each person's night / weekend / total / per-PP counts
      # EXACTLY and leaves Stage-1 fairness untouched.
      filter_cands <- function(cands, d, vecP, s) {
        if (!length(cands)) return(cands)
        keep <- cands %in% dates_vec & cands != d &
                !(as.character(cands) %in% names(vecP)) &
                (is_wknd(cands, s) == is_wknd(d, s)) &
                vapply(seq_along(cands),
                       function(j) identical(pp_of(cands[j]), pp_of(d)), logical(1L))
        cands[keep]
      }

      # ── Shared gate-and-commit ──────────────────────────────────────────────
      # Tries to move P's shift on `d` to one of `cands` by swapping with whoever
      # holds that day+slot. Returns TRUE on the first accepted swap.
      try_swap <- function(P, d, cands) {
        vecP <- asgn[[P]]; dk <- as.character(d)
        if (!(dk %in% names(vecP))) return(FALSE)
        s <- unname(vecP[dk]); if (is.na(s)) return(FALSE)
        if (is_pinned(P, dk, s) || is_holiday_slot(dk, s)) return(FALSE)
        for (cj in seq_along(cands)) {                # index - preserve Date class
          t <- cands[cj]; tk <- as.character(t)
          Q  <- occ(self$schedule[[tk]], s)
          if (is.na(Q) || Q == P) next
          if (is_pinned(Q, tk, s) || is_holiday_slot(tk, s)) next
          if (cannot_work(Q, d) || cannot_work(P, t)) next
          if (dk %in% names(asgn[[Q]])) next          # Q already works d
          vP <- c(vecP[names(vecP) != dk], setNames(s, tk))
          vQ <- c(asgn[[Q]][names(asgn[[Q]]) != tk], setNames(s, dk))
          # Accept only a strict improvement in combined run-length + green score.
          if (total_score(P, vP) + total_score(Q, vQ) <=
              total_score(P, vecP) + total_score(Q, asgn[[Q]])) next
          if (!private$person_feasible(P, vP, focus = t) ||
              !private$person_feasible(Q, vQ, focus = d)) next
          # commit
          self$schedule[[dk]][[s]] <- Q
          self$schedule[[tk]][[s]] <- P
          asgn[[P]] <<- vP; asgn[[Q]] <<- vQ
          return(TRUE)
        }
        FALSE
      }

      swaps_green <- 0L; swaps_run <- 0L
      for (pass in seq_len(max_passes)) {
        improved <- FALSE

        # ── Phase B: chase requested-work days (fill phase only) ──────────────
        # The run-length generator below cannot find these moves: it only ever
        # considers shifts ALREADY in a run of <=2 as sources, and only days
        # ADJACENT to another of P's shifts as targets. A yellow shift parked
        # mid-3-run can never relocate to a non-adjacent green day. So without
        # this generator the green term is purely defensive.
        if (!green_only) {
          for (P in STAFF) {
            want <- setdiff(green_key[[P]], names(asgn[[P]]))   # green, not worked
            if (!length(want)) next
            src <- as.Date(setdiff(names(asgn[[P]]), green_key[[P]]))  # worked, not green
            for (idi in seq_along(src)) {
              d <- src[idi]
              s_src <- unname(asgn[[P]][as.character(d)])
              if (is.na(s_src)) next
              cands <- filter_cands(as.Date(want), d, asgn[[P]], s_src)
              if (!length(cands)) next
              if (try_swap(P, d, cands)) { swaps_green <- swaps_green + 1L; improved <- TRUE }
            }
          }
        }

        # ── Phase A: grow short runs (unchanged behaviour) ────────────────────
        for (P in STAFF) {
          src_ds <- short_dates(asgn[[P]])
          for (idi in seq_along(src_ds)) {           # index - preserve Date class
            d <- src_ds[idi]; vecP <- asgn[[P]]
            others <- as.Date(names(vecP)); others <- others[others != d]
            if (length(others) == 0L) next
            cand_int <- unique(unlist(lapply(seq_along(others),
                          function(j) as.integer(c(others[j] - 1, others[j] + 1)))))
            s_src <- unname(vecP[as.character(d)])
            if (is.na(s_src)) next
            cands <- filter_cands(as.Date(cand_int, origin = "1970-01-01"), d, vecP, s_src)
            if (!length(cands)) next
            if (try_swap(P, d, cands)) { swaps_run <- swaps_run + 1L; improved <- TRUE }
          }
        }
        if (!improved) break
      }
      private$rebuild_derived()
      message(sprintf(
        "  Greedy pass: %d swap(s) applied (%d chasing requested-work days, %d run-length).",
        swaps_green + swaps_run, swaps_green, swaps_run))
    },

    # Rebuild person_shifts / person_nights / pp_counts from self$schedule
    # (after greedy swaps mutate the slot assignments in place).
    rebuild_derived = function() {
      self$person_nights <- setNames(lapply(STAFF, function(p) as.Date(character())), STAFF)
      self$person_shifts <- setNames(lapply(STAFF, function(p)
        data.frame(date = as.Date(character()), slot = character(), stringsAsFactors = FALSE)), STAFF)
      zero_pp <- setNames(integer(nrow(PAY_PERIODS)), PAY_PERIODS$name)
      self$pp_counts <- setNames(lapply(STAFF, function(p) zero_pp), STAFF)
      for (ds in names(self$schedule)) {
        d <- as.Date(ds); pp <- get_pp(d); day <- self$schedule[[ds]]
        for (s in c("APP1", "APP2", "Roaming", "Night")) {
          p <- day[[s]]; if (length(p) != 1L || is.na(p)) next
          if (!is.na(pp)) self$pp_counts[[p]][[pp]] <- self$pp_counts[[p]][[pp]] + 1L
          if (s == "Night") self$person_nights[[p]] <- sort(c(self$person_nights[[p]], d))
          else self$person_shifts[[p]] <- rbind(self$person_shifts[[p]],
                 data.frame(date = d, slot = s, stringsAsFactors = FALSE))
        }
      }
    },

    # ── Phase 2: greedy APP3 fill-in for under-scheduled staff ─────────────────
    # Runs after the ILP solution is committed.  Assigns empty Roaming slots to
    # people who are below their sched_target for that PP.  All hard constraints
    # (availability, no double-booking, C7 night→day ban, C10 max-4-consec, PP
    # cap) are respected; run-length shaping is ignored (aesthetics only).
    fill_roaming_pass = function() {
      # Use explicit index iteration to guarantee Date class is preserved on each d
      dates_vec <- sort(as.Date(names(self$schedule)))

      is_working_p <- function(person, d) {
        day <- self$schedule[[format(d, "%Y-%m-%d")]]
        if (is.null(day)) return(FALSE)
        isTRUE(person %in% c(day$APP1, day$APP2, day$Night, day$Roaming))
      }

      had_night_recent <- function(person, d) {
        # C7/C7b: block if person worked Night on d-1 or d-2
        for (k in 1:2) {
          ps <- format(d - k, "%Y-%m-%d")
          if (isTRUE(ps %in% names(self$schedule)) &&
              isTRUE(self$schedule[[ps]]$Night == person))
            return(TRUE)
        }
        FALSE
      }

      # C10: adding d would create a run of 5+ consecutive work days
      would_exceed_consec <- function(person, d, max_consec = 4L) {
        work_set <- as.Date(c(
          self$person_shifts[[person]]$date,
          self$person_nights[[person]]
        ))
        if (length(work_set) == 0L) return(FALSE)
        run <- 1L
        k   <- 1L
        while (k <= max_consec && isTRUE((d - k) %in% work_set)) {
          run <- run + 1L; k <- k + 1L
        }
        k <- 1L
        while (k <= max_consec && isTRUE((d + k) %in% work_set)) {
          run <- run + 1L; k <- k + 1L
        }
        run > max_consec
      }

      n_filled <- 0L

      for (di in seq_along(dates_vec)) {
        d   <- dates_vec[di]   # [i] preserves Date class
        ds  <- format(d, "%Y-%m-%d")
        day <- self$schedule[[ds]]
        if (is.null(day) || !is.na(day$Roaming)) next

        pp <- get_pp(d)
        if (is.na(pp)) next

        candidates <- character(0)
        deficits   <- integer(0)

        for (person in STAFF) {
          tgt <- self$targets[[person]][[pp]]
          if (is.null(tgt)) next
          deficit <- tgt$sched_target - self$pp_counts[[person]][[pp]]
          if (!isTRUE(deficit > 0L)) next
          if (private$is_blocked(person, d)) next
          if (is_working_p(person, d)) next
          if (had_night_recent(person, d)) next
          if (would_exceed_consec(person, d)) next

          shifts_df      <- self$person_shifts[[person]]
          n_nights       <- length(self$person_nights[[person]])
          n_weekends     <- if (nrow(shifts_df) == 0L) 0L else
                             sum(is_weekend_shift(shifts_df$date, shifts_df$slot))
          # Primary: PP deficit; secondary: total nights (compensate heavy
          # night workers); tertiary: weekend shifts worked
          score <- deficit * 100L + n_nights * 10L + n_weekends

          candidates <- c(candidates, person)
          deficits   <- c(deficits, score)
        }

        if (length(candidates) == 0L) next

        best <- candidates[which.max(deficits)]
        self$schedule[[ds]]$Roaming           <- best
        self$pp_counts[[best]][[pp]]          <- self$pp_counts[[best]][[pp]] + 1L
        self$person_shifts[[best]] <- rbind(
          self$person_shifts[[best]],
          data.frame(date = d, slot = "Roaming", stringsAsFactors = FALSE))
        n_filled <- n_filled + 1L
      }

      message(sprintf("  Roaming fill pass: %d additional APP3 shift(s) assigned.", n_filled))
    }
  )
)
