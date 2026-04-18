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
#   8 variables encoding per-metric max/min across all staff, used to push
#   min–max spread into the objective (tiny ε coefficients, tie-breaking only).
#   I_MAX_NIGHTS, I_MIN_NIGHTS, I_MAX_TOTAL, I_MIN_TOTAL,
#   I_MAX_WKND,   I_MIN_WKND,   I_MAX_ROAM,  I_MIN_ROAM
#
# Work auxiliaries (continuous [0, 1]):
#   work[p, d]  ≡  Σ_s x[p,d,s]   (auto-integer via C4 no-double-book)
#   Flat 1-based index:   nX + nF + (p-1)*nD + d
#   Reduces C10 from 20 to 5 coefficients per sliding window.
#
# CONSTRAINTS
# ───────────
#   C1    Slot uniqueness:    Σ_p x[p,d,s]     ≤ 1           for all d,s
#   C2    APP1 coverage:      Σ_p x[p,d,APP1]  = 1           for all d
#   C3    Night coverage:     Σ_p x[p,d,Night] = 1  (or ≤1)  for all d
#   C4    No double-book:     Σ_s x[p,d,s]     ≤ 1           for all p,d
#   C5    Availability:       ub = 0 for blocked (p,d) pairs
#   C6    PP shift cap:       Σ_{d∈PP,s} x[p,d,s] ≤ target   for all p,PP
#   C7    Night→Day 1d ban:   x[p,d,Night] + x[p,d+1,s] ≤ 1  s∈DAY_SLOTS
#   C8    Day→Night 1d gap:   x[p,d,s] + x[p,d+1,Night] ≤ 1  s∈DAY_SLOTS
#   C9    Max 4 consec Nts:   Σ_{k=0}^4 x[p,d+k,Night]    ≤ 4
#   C10   Max 4 consec work:  Σ_{k=0}^4 work[p,d+k]       ≤ 4
#   C11   Night total cap:    Σ_d x[p,d,Night] ≤ MAX_NIGHTS_TOTAL
#   C12   Holiday pre-seed:   lb = ub = 1 for pre-assigned (p,d,s)
#   C13   Min 2 consec Nts:   isolated night shifts forbidden
#   C14   Fairness bounds:    Σ x ≤ max_M,  Σ x ≥ min_M  per person per metric
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
# Tries 9 tiers in order (full model → hard coverage only).
# If all tiers fail, diagnostics are printed and execution stops.
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
    tier_used     = NULL,   # list(index, label) of relaxation tier that found a solution

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
    run = function(n_candidates = 10L) {
      if (!requireNamespace("highs", quietly = TRUE))
        stop("highs package is not installed. Run: install.packages('highs')")

      tiers <- private$RELAX_TIERS
      nT    <- length(tiers)

      # Helper to call build_and_solve with a tier's parameters
      solve_tier <- function(t, extra_nogo = list()) {
        private$build_and_solve(
          night_required   = t$night_req,
          roam_in_obj      = t$roam_obj,
          pp_cap_reduction = t$pp_red,
          add_c8           = t$c8,
          add_c9           = t$c9,
          add_c10          = t$c10,
          add_c13          = t$c13,
          add_c14          = t$c14,
          add_c15          = t$c15,
          add_c16          = t$c16,
          add_c_min        = t$c_min,
          allow_pto        = t$allow_pto,
          extra_nogo       = extra_nogo
        )
      }

      for (ti in seq_len(nT)) {
        t      <- tiers[[ti]]
        message(sprintf("  [Tier %d/%d] %s", ti, nT, t$label))
        result <- solve_tier(t)
        if (is.null(result)) next

        self$tier_used <- list(index = ti, label = t$label)

        # Collect up to n_candidates distinct solutions at this tier, then pick best
        candidates <- list(result)
        x_found    <- list(result$sol[seq_len(result$nX)])

        if (n_candidates > 1L) {
          message(sprintf("  Collecting up to %d candidates at tier %d…", n_candidates, ti))
          for (ci in seq_len(n_candidates - 1L)) {
            cand <- solve_tier(t, extra_nogo = x_found)
            if (is.null(cand)) break
            candidates <- c(candidates, list(cand))
            x_found    <- c(x_found, list(cand$sol[seq_len(cand$nX)]))
          }
        }

        scores   <- vapply(candidates, private$score_solution, numeric(1L))
        best_idx <- which.max(scores)
        message(sprintf("  %d candidate(s). Scores: [%s]  → best #%d (%.1f)",
                        length(candidates),
                        paste(round(scores, 1L), collapse = ", "),
                        best_idx, scores[best_idx]))

        message("  Translating ILP solution to schedule…")
        private$populate_from_solution(candidates[[best_idx]])
        message("  Done.")
        return(invisible(self))
      }

      private$report_and_stop()
    },

    # ── Enumerate distinct feasible solutions via no-good cut iteration ───────
    # Adds a cut excluding each previously found x-assignment, then re-solves.
    # Returns integer count (lower bound; stops at max_count or infeasibility).
    count_solutions = function(max_count = 20L) {
      if (is.null(self$tier_used)) stop("Call run() before count_solutions().")
      t     <- private$RELAX_TIERS[[self$tier_used$index]]
      found <- list()
      count <- 0L
      repeat {
        res <- private$build_and_solve(
          night_required   = t$night_req,
          roam_in_obj      = t$roam_obj,
          pp_cap_reduction = t$pp_red,
          add_c8           = t$c8,
          add_c9           = t$c9,
          add_c10          = t$c10,
          add_c13          = t$c13,
          add_c14          = t$c14,
          add_c15          = t$c15,
          add_c16          = t$c16,
          add_c_min        = t$c_min,
          allow_pto        = t$allow_pto,
          extra_nogo       = found
        )
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

    # ── Score a candidate solution (higher = better) ──────────────────────────
    # Works directly on the raw result list from build_and_solve — no need to
    # populate schedule data structures.  Four criteria:
    #   1. Target attainment  — quadratic penalty for each missed shift per PP
    #   2. Night fairness     — penalty proportional to SD of night counts
    #   3. Work-run quality   — reward 3/4-day blocks, penalise 1/2-day runs
    #   4. Night-pack quality — reward 3-night packs, penalise 1/4-night runs
    score_solution = function(res) {
      sol       <- res$sol
      nP        <- res$nP
      nD        <- res$nD
      xidx      <- res$xidx
      dates_vec <- res$dates_vec
      S_NIGHT   <- 4L
      score     <- 0.0

      streak_lengths <- function(v) {
        r <- rle(as.integer(v))
        r$lengths[r$values == 1L]
      }

      # ---- 1. Target attainment -----------------------------------------------
      for (pi in seq_len(nP)) {
        person <- STAFF[pi]
        for (ppi in seq_len(nrow(PAY_PERIODS))) {
          pp_d  <- seq(PAY_PERIODS$start[ppi], PAY_PERIODS$end[ppi], by = "day")
          di_pp <- which(dates_vec %in% pp_d)
          if (length(di_pp) == 0L) next
          actual <- sum(vapply(di_pp, function(di)
            sum(vapply(1:4, function(s) round(sol[xidx(pi, di, s)]), numeric(1L))),
            numeric(1L)))
          sm        <- self$targets[[person]][[PAY_PERIODS$name[ppi]]]$sched_target
          shortfall <- max(0L, sm - actual)
          score     <- score - 10.0 * shortfall^2
        }
      }

      # ---- 2. Night fairness --------------------------------------------------
      nights <- vapply(seq_len(nP), function(pi)
        sum(vapply(seq_len(nD), function(di)
          round(sol[xidx(pi, di, S_NIGHT)]), numeric(1L))),
        numeric(1L))
      if (nP > 1L) score <- score - 5.0 * sd(nights)

      # ---- 3 & 4. Run / pack quality ------------------------------------------
      run_w   <- c("1" = -2.0, "2" = -0.5, "3" = 2.0, "4" =  1.0)
      night_w <- c("1" = -0.5, "2" =  0.3, "3" = 1.0, "4" = -0.5)

      for (pi in seq_len(nP)) {
        # Work run quality (all slots)
        work <- vapply(seq_len(nD), function(di)
          as.integer(any(vapply(1:4, function(s) round(sol[xidx(pi, di, s)]) == 1L,
                                logical(1L)))),
          integer(1L))
        for (L in streak_lengths(work))
          score <- score + run_w[min(L, 4L)]

        # Night pack quality
        nv <- vapply(seq_len(nD), function(di)
          as.integer(round(sol[xidx(pi, di, S_NIGHT)])), integer(1L))
        for (L in streak_lengths(nv))
          score <- score + night_w[min(L, 4L)]
      }

      score
    },

    # Relaxation cascade: solver tries each tier in order, stopping at the
    # first feasible solution.  Each tier adds more flexibility than the last.
    RELAX_TIERS = list(
      list(label="Full model — all rules",
           night_req=TRUE,  roam_obj=TRUE,  pp_red=0L, allow_pto=FALSE, c_min=TRUE,
           c8=TRUE,  c9=TRUE,  c10=TRUE,  c13=TRUE,  c14=TRUE,  c15=TRUE,  c16=TRUE),
      list(label="Run >= 2 days (pairs allowed), APP3 required",
           night_req=TRUE,  roam_obj=TRUE,  pp_red=0L, allow_pto=FALSE, c_min=TRUE,
           c8=TRUE,  c9=TRUE,  c10=TRUE,  c13=TRUE,  c14=TRUE,  c15=TRUE,  c16=FALSE),
      list(label="Run >= 2 days, APP3/Roaming may be unstaffed",
           night_req=TRUE,  roam_obj=FALSE, pp_red=0L, allow_pto=FALSE, c_min=TRUE,
           c8=TRUE,  c9=TRUE,  c10=TRUE,  c13=TRUE,  c14=TRUE,  c15=TRUE,  c16=FALSE),
      list(label="Any run length, APP3 required",
           night_req=TRUE,  roam_obj=TRUE,  pp_red=0L, allow_pto=FALSE, c_min=TRUE,
           c8=TRUE,  c9=TRUE,  c10=TRUE,  c13=TRUE,  c14=TRUE,  c15=FALSE, c16=FALSE),
      list(label="Night shift may be unstaffed",
           night_req=FALSE, roam_obj=FALSE, pp_red=0L, allow_pto=FALSE, c_min=TRUE,
           c8=TRUE,  c9=TRUE,  c10=TRUE,  c13=TRUE,  c14=TRUE,  c15=FALSE, c16=FALSE),
      list(label="PP cap reduced by 1",
           night_req=FALSE, roam_obj=FALSE, pp_red=1L, allow_pto=FALSE, c_min=TRUE,
           c8=TRUE,  c9=TRUE,  c10=TRUE,  c13=TRUE,  c14=TRUE,  c15=FALSE, c16=FALSE),
      list(label="Drop C13 (min 2 consecutive nights)",
           night_req=FALSE, roam_obj=FALSE, pp_red=1L, allow_pto=FALSE, c_min=TRUE,
           c8=TRUE,  c9=TRUE,  c10=TRUE,  c13=FALSE, c14=TRUE,  c15=FALSE, c16=FALSE),
      list(label="Drop C8 (day -> night gap)",
           night_req=FALSE, roam_obj=FALSE, pp_red=1L, allow_pto=FALSE, c_min=TRUE,
           c8=FALSE, c9=TRUE,  c10=TRUE,  c13=TRUE,  c14=TRUE,  c15=FALSE, c16=FALSE),
      list(label="Drop C8 + C13 / drop shift minimums",
           night_req=FALSE, roam_obj=FALSE, pp_red=1L, allow_pto=FALSE, c_min=FALSE,
           c8=FALSE, c9=TRUE,  c10=TRUE,  c13=FALSE, c14=TRUE,  c15=FALSE, c16=FALSE),
      list(label="Drop C8 + C13 + C10",
           night_req=FALSE, roam_obj=FALSE, pp_red=1L, allow_pto=FALSE, c_min=FALSE,
           c8=FALSE, c9=TRUE,  c10=FALSE, c13=FALSE, c14=TRUE,  c15=FALSE, c16=FALSE),
      list(label="PTO credits for staff with >4 off/vac days in PP",
           night_req=FALSE, roam_obj=FALSE, pp_red=0L, allow_pto=TRUE,  c_min=FALSE,
           c8=FALSE, c9=TRUE,  c10=FALSE, c13=FALSE, c14=TRUE,  c15=FALSE, c16=FALSE),
      list(label="Hard coverage only (C1-C7, C11, C12)",
           night_req=TRUE,  roam_obj=TRUE,  pp_red=0L, allow_pto=TRUE,  c_min=FALSE,
           c8=FALSE, c9=FALSE, c10=FALSE, c13=FALSE, c14=FALSE, c15=FALSE, c16=FALSE)
    ),

    # ── Build and solve the ILP (HiGHS) ──────────────────────────────────────
    build_and_solve = function(
      night_required   = TRUE,
      roam_in_obj      = TRUE,
      pp_cap_reduction = 0L,
      add_c8           = TRUE,
      add_c9           = TRUE,
      add_c10          = TRUE,
      add_c13          = TRUE,
      add_c14          = TRUE,
      add_c15          = TRUE,   # no isolated single shifts
      add_c16          = TRUE,   # no isolated 2-day blocks of day shifts (min run 3)
      add_c_min        = TRUE,   # enforce soft_min floor for FLEX_TARGETS staff (e.g. Todd >= 4)
      allow_pto        = FALSE,  # credit off/vac days as PTO when blocked_days > 4
      extra_nogo       = list()  # previously found x-vectors to exclude (solution enumeration)
    ) {

      dates_vec <- as.Date(self$dates, origin = "1970-01-01")
      nP <- length(STAFF)
      nD <- length(dates_vec)
      nS <- 4L

      S_APP1 <- 1L; S_APP2 <- 2L; S_ROAM <- 3L; S_NIGHT <- 4L
      DAY_S  <- c(S_APP1, S_APP2, S_ROAM)

      # x[p,d,s] binary
      nX   <- nP * nD * nS
      xidx <- function(p, d, s) (p - 1L) * nD * nS + (d - 1L) * nS + s

      # Fairness auxiliaries: 8 continuous [0, nD]
      nF           <- 8L
      I_MAX_NIGHTS <- nX + 1L;  I_MIN_NIGHTS <- nX + 2L
      I_MAX_TOTAL  <- nX + 3L;  I_MIN_TOTAL  <- nX + 4L
      I_MAX_WKND   <- nX + 5L;  I_MIN_WKND   <- nX + 6L
      I_MAX_ROAM   <- nX + 7L;  I_MIN_ROAM   <- nX + 8L

      # work[p,d] continuous [0,1] — equals Σ_s x[p,d,s], auto-integer via C4
      nW   <- nP * nD
      widx <- function(p, d) nX + nF + (p - 1L) * nD + d

      # Streak auxiliaries (all continuous [0,1]):
      #   n3/n4 bias night assignments toward 3-packs (3 > 2 > 4)
      #   w3/w4 bias work-day runs toward 3-blocks (3 > 4 > 2 > 1)
      nNS3  <- if (nD >= 3L) nP * (nD - 2L) else 0L
      nNS4  <- if (nD >= 4L) nP * (nD - 3L) else 0L
      nWS3  <- if (nD >= 3L) nP * (nD - 2L) else 0L
      nWS4  <- if (nD >= 4L) nP * (nD - 3L) else 0L

      ns3off <- nX + nF + nW
      n3idx  <- function(p, d) ns3off + (p - 1L) * (nD - 2L) + d
      ns4off <- ns3off + nNS3
      n4idx  <- function(p, d) ns4off + (p - 1L) * (nD - 3L) + d
      ws3off <- ns4off + nNS4
      w3idx  <- function(p, d) ws3off + (p - 1L) * (nD - 2L) + d
      ws4off <- ws3off + nWS3
      w4idx  <- function(p, d) ws4off + (p - 1L) * (nD - 3L) + d

      nV <- nX + nF + nW + nNS3 + nNS4 + nWS3 + nWS4

      # Variable bounds and types
      lb    <- numeric(nV)
      ub    <- c(rep(1, nX), rep(as.double(nD), nF), rep(1, nW),
                 rep(1, nNS3), rep(1, nNS4), rep(1, nWS3), rep(1, nWS4))
      types <- c(rep("I", nX), rep("C", nF + nW + nNS3 + nNS4 + nWS3 + nWS4))

      # Objective (maximise)
      roam_w <- if (roam_in_obj) 2 else 0
      obj    <- numeric(nV)
      for (p in seq_len(nP)) {
        for (d in seq_len(nD)) {
          obj[xidx(p, d, S_APP1)]  <- 4
          obj[xidx(p, d, S_APP2)]  <- 2
          obj[xidx(p, d, S_ROAM)]  <- roam_w
          obj[xidx(p, d, S_NIGHT)] <- 4
        }
      }
      obj[I_MAX_NIGHTS] <- -0.010;  obj[I_MIN_NIGHTS] <- +0.010
      obj[I_MAX_TOTAL]  <- -0.005;  obj[I_MIN_TOTAL]  <- +0.005
      obj[I_MAX_WKND]   <- -0.003;  obj[I_MIN_WKND]   <- +0.003
      obj[I_MAX_ROAM]   <- -0.003;  obj[I_MIN_ROAM]   <- +0.003

      # Streak preference bonuses/penalties (3-packs most preferred)
      # Night: reward 3-consecutive (+0.50), penalise 4-consecutive (-0.60)
      #   Net per streak: 2-nights=0, 3-nights=+0.50, 4-nights=+0.40
      # Work: reward 3-consecutive (+0.30), penalise 4-consecutive (-0.40)
      #   Net per streak: 1-day=0, 2-days=0, 3-days=+0.30, 4-days=+0.20
      if (nNS3 > 0L)
        for (p in seq_len(nP)) for (d in seq_len(nD - 2L)) obj[n3idx(p, d)] <- +0.50
      if (nNS4 > 0L)
        for (p in seq_len(nP)) for (d in seq_len(nD - 3L)) obj[n4idx(p, d)] <- -0.60
      if (nWS3 > 0L)
        for (p in seq_len(nP)) for (d in seq_len(nD - 2L)) obj[w3idx(p, d)] <- +0.30
      if (nWS4 > 0L)
        for (p in seq_len(nP)) for (d in seq_len(nD - 3L)) obj[w4idx(p, d)] <- -0.40

      # Constraint accumulator (triplet form → sparseMatrix)
      n_con   <- 0L
      ri      <- integer(0); ci <- integer(0); vi <- double(0)
      con_lhs <- double(0);  con_rhs <- double(0)

      add_con <- function(indices, coeffs, type, rhs_val) {
        n_con <<- n_con + 1L
        ri      <<- c(ri, rep(n_con, length(indices)))
        ci      <<- c(ci, as.integer(indices))
        vi      <<- c(vi, as.double(coeffs))
        if (type == "<=") {
          con_lhs <<- c(con_lhs, -Inf);    con_rhs <<- c(con_rhs, rhs_val)
        } else if (type == ">=") {
          con_lhs <<- c(con_lhs, rhs_val); con_rhs <<- c(con_rhs, Inf)
        } else {
          con_lhs <<- c(con_lhs, rhs_val); con_rhs <<- c(con_rhs, rhs_val)
        }
      }

      # ── C1: Slot uniqueness — Σ_p x[p,d,s] ≤ 1 ──────────────────────────────
      for (d in seq_len(nD)) {
        for (s in seq_len(nS)) {
          cols <- vapply(seq_len(nP), function(p) xidx(p, d, s), integer(1L))
          add_con(cols, rep(1, nP), "<=", 1)
        }
      }

      # ── C2: APP1 always filled — Σ_p x[p,d,APP1] = 1 ────────────────────────
      for (d in seq_len(nD)) {
        cols <- vapply(seq_len(nP), function(p) xidx(p, d, S_APP1), integer(1L))
        add_con(cols, rep(1, nP), "=", 1)
      }

      # ── C3: Night coverage ────────────────────────────────────────────────────
      c3_type <- if (night_required) "=" else "<="
      for (d in seq_len(nD)) {
        cols <- vapply(seq_len(nP), function(p) xidx(p, d, S_NIGHT), integer(1L))
        add_con(cols, rep(1, nP), c3_type, 1)
      }

      # ── C4: No double-booking — Σ_s x[p,d,s] ≤ 1 ────────────────────────────
      for (p in seq_len(nP)) {
        for (d in seq_len(nD)) {
          cols <- vapply(seq_len(nS), function(s) xidx(p, d, s), integer(1L))
          add_con(cols, rep(1, nS), "<=", 1)
        }
      }

      # ── C_work: work[p,d] = Σ_s x[p,d,s] ────────────────────────────────────
      for (pi in seq_len(nP)) {
        for (di in seq_len(nD)) {
          x_cols <- vapply(seq_len(nS), function(s) xidx(pi, di, s), integer(1L))
          add_con(c(x_cols, widx(pi, di)), c(rep(1L, nS), -1L), "=", 0L)
        }
      }

      # ── C5: Availability — set ub = 0 for blocked (p,d) ─────────────────────
      for (pi in seq_len(nP)) {
        person <- STAFF[pi]
        pdata  <- self$time_off[[person]]
        for (di in seq_len(nD)) {
          d <- dates_vec[di]
          blocked <- nrow(pdata) > 0 &&
            any(pdata$date == d & pdata$type %in% c("off", "vac", "cme"))
          if (blocked) {
            for (s in seq_len(nS)) ub[xidx(pi, di, s)] <- 0
          }
        }
      }

      # ── C5b: No night shift the day before an off or vacation day ────────────
      for (pi in seq_len(nP)) {
        person <- STAFF[pi]
        pdata  <- self$time_off[[person]]
        if (nrow(pdata) == 0) next
        for (di in seq_len(nD - 1L)) {
          d_next <- dates_vec[di + 1L]
          if (any(pdata$date == d_next & pdata$type %in% c("off", "vac")))
            ub[xidx(pi, di, S_NIGHT)] <- 0
        }
      }

      # ── C6: PP shift cap — Σ_{d∈PP,s} x[p,d,s] ≤ sched_target[p,PP] ────────
      for (pi in seq_len(nP)) {
        person <- STAFF[pi]
        for (ppi in seq_len(nrow(PAY_PERIODS))) {
          pp_name <- PAY_PERIODS$name[ppi]
          pp_d    <- seq(PAY_PERIODS$start[ppi], PAY_PERIODS$end[ppi], by = "day")
          di_pp   <- which(dates_vec %in% pp_d)
          pto_credit <- 0L
          if (allow_pto && length(di_pp) > 0L) {
            pdata    <- self$time_off[[person]]
            pp_dates <- pp_d[pp_d %in% dates_vec]
            n_offvac <- if (nrow(pdata) == 0L) 0L else
              sum(vapply(pp_dates, function(d)
                any(pdata$date == d & pdata$type %in% c("off", "vac")),
                logical(1L)))
            if (n_offvac > 4L) pto_credit <- n_offvac - 4L
          }
          cap <- max(0L, self$targets[[person]][[pp_name]]$sched_target -
                          pp_cap_reduction - pto_credit)
          if (length(di_pp) == 0 || cap <= 0) next
          cols <- as.integer(unlist(lapply(di_pp, function(di)
            vapply(seq_len(nS), function(s) xidx(pi, di, s), integer(1L)))))
          add_con(cols, rep(1, length(cols)), "<=", cap)
        }
      }

      # ── C_min: Minimum shifts per PP for all staff ──────────────────────────────
      # Enforces Σ x[p,PP] ≥ soft_min for every person.
      # Non-FLEX staff: soft_min = sched_target (must hit full target).
      # FLEX_TARGETS (e.g. Todd): soft_min = FLEX_TARGETS value (e.g. 4).
      # Skipped when person has fewer available days than their soft_min.
      if (add_c_min) {
        for (pi in seq_len(nP)) {
          person <- STAFF[pi]
          for (ppi in seq_len(nrow(PAY_PERIODS))) {
            pp_name <- PAY_PERIODS$name[ppi]
            pp_d    <- seq(PAY_PERIODS$start[ppi], PAY_PERIODS$end[ppi], by = "day")
            di_pp   <- which(dates_vec %in% pp_d)
            t_info  <- self$targets[[person]][[pp_name]]
            sm      <- t_info$soft_min
            if (sm <= 0L || length(di_pp) == 0L || t_info$avail < sm) next
            cols <- as.integer(unlist(lapply(di_pp, function(di)
              vapply(seq_len(nS), function(s) xidx(pi, di, s), integer(1L)))))
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

      # ── C8: Day→Night gap 1d — x[p,d,s] + x[p,d+1,Night] ≤ 1 ──────────────
      if (add_c8) {
        for (pi in seq_len(nP)) {
          for (di in seq_len(nD - 1L)) {
            ni <- xidx(pi, di + 1L, S_NIGHT)
            for (s in DAY_S)
              add_con(c(xidx(pi, di, s), ni), c(1, 1), "<=", 1)
          }
        }
      }

      # ── C9: Max 4 consecutive nights — Σ_{k=0}^4 x[p,d+k,Night] ≤ 4 ────────
      if (add_c9) {
        for (pi in seq_len(nP)) {
          for (di in seq_len(nD - 4L)) {
            cols <- vapply(0:4, function(k) xidx(pi, di + k, S_NIGHT), integer(1L))
            add_con(cols, rep(1, 5L), "<=", 4)
          }
        }
      }

      # ── C10: Max 4 consecutive working days — Σ_{k=0}^4 work[p,d+k] ≤ 4 ─────
      if (add_c10) {
        for (pi in seq_len(nP)) {
          for (di in seq_len(nD - 4L)) {
            cols <- vapply(0:4, function(k) widx(pi, di + k), integer(1L))
            add_con(cols, rep(1L, 5L), "<=", 4L)
          }
        }
      }

      # ── C11: Total nights per person ≤ MAX_NIGHTS_TOTAL ──────────────────────
      for (pi in seq_len(nP)) {
        cols <- vapply(seq_len(nD), function(di) xidx(pi, di, S_NIGHT), integer(1L))
        add_con(cols, rep(1, nD), "<=", MAX_NIGHTS_TOTAL)
      }

      # ── C12: Holiday pre-seeds — fix x[p,d,s] = 1 ────────────────────────────
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
          idx    <- xidx(pi[[1L]], di[[1L]], si)
          lb[idx] <- 1
          ub[idx] <- 1
        }
      }

      # ── C13: Min 2 consecutive nights ────────────────────────────────────────
      # Boundary d=1:  x[p,1,Nt] ≤ x[p,2,Nt]
      # Interior:      x[p,d-1,Nt] + x[p,d+1,Nt] ≥ x[p,d,Nt]
      # Boundary d=nD: x[p,nD,Nt] ≤ x[p,nD-1,Nt]
      if (add_c13) {
        for (pi in seq_len(nP)) {
          add_con(c(xidx(pi, 1L, S_NIGHT), xidx(pi, 2L, S_NIGHT)),
                  c(1, -1), "<=", 0L)
          if (nD >= 3L) {
            for (di in 2L:(nD - 1L)) {
              add_con(c(xidx(pi, di - 1L, S_NIGHT),
                        xidx(pi, di + 1L, S_NIGHT),
                        xidx(pi, di,      S_NIGHT)),
                      c(1, 1, -1), ">=", 0L)
            }
          }
          add_con(c(xidx(pi, nD, S_NIGHT), xidx(pi, nD - 1L, S_NIGHT)),
                  c(1, -1), "<=", 0L)
        }
      }

      # ── C14: Fairness min/max bounds ─────────────────────────────────────────
      weekend_di <- which(weekdays(dates_vec) %in% c("Saturday", "Sunday"))
      if (add_c14) {
        for (pi in seq_len(nP)) {
          night_cols <- vapply(seq_len(nD), function(di)
            xidx(pi, di, S_NIGHT), integer(1L))
          all_cols   <- as.integer(unlist(lapply(seq_len(nD), function(di)
            vapply(seq_len(nS), function(s) xidx(pi, di, s), integer(1L)))))
          roam_cols  <- vapply(seq_len(nD), function(di)
            xidx(pi, di, S_ROAM), integer(1L))

          add_con(c(night_cols, I_MAX_NIGHTS), c(rep(1, nD), -1L),      "<=", 0)
          add_con(c(night_cols, I_MIN_NIGHTS), c(rep(1, nD), -1L),      ">=", 0)
          add_con(c(all_cols,   I_MAX_TOTAL),  c(rep(1, nD * nS), -1L), "<=", 0)
          add_con(c(all_cols,   I_MIN_TOTAL),  c(rep(1, nD * nS), -1L), ">=", 0)
          add_con(c(roam_cols,  I_MAX_ROAM),   c(rep(1, nD), -1L),      "<=", 0)
          add_con(c(roam_cols,  I_MIN_ROAM),   c(rep(1, nD), -1L),      ">=", 0)

          if (length(weekend_di) > 0) {
            wknd_cols <- as.integer(unlist(lapply(weekend_di, function(di)
              vapply(seq_len(nS), function(s) xidx(pi, di, s), integer(1L)))))
            add_con(c(wknd_cols, I_MAX_WKND), c(rep(1, length(wknd_cols)), -1L), "<=", 0)
            add_con(c(wknd_cols, I_MIN_WKND), c(rep(1, length(wknd_cols)), -1L), ">=", 0)
          }
        }
      }

      # ── C15: No isolated single shifts — work[p,d] ≤ work[p,d-1] + work[p,d+1] ─
      # Boundary d=1:  work[p,1] ≤ work[p,2]
      # Interior:      work[p,d-1] + work[p,d+1] ≥ work[p,d]
      # Boundary d=nD: work[p,nD] ≤ work[p,nD-1]
      if (add_c15) {
        for (pi in seq_len(nP)) {
          add_con(c(widx(pi, 1L), widx(pi, 2L)), c(1, -1), "<=", 0L)
          if (nD >= 3L) {
            for (di in 2L:(nD - 1L)) {
              add_con(c(widx(pi, di - 1L), widx(pi, di + 1L), widx(pi, di)),
                      c(1, 1, -1), ">=", 0L)
            }
          }
          add_con(c(widx(pi, nD), widx(pi, nD - 1L)), c(1, -1), "<=", 0L)
        }
      }

      # ── C16: No isolated 2-day blocks of day shifts ──────────────────────────
      # D(p,d) = Σ_{s∈DAY_S} x[p,d,s].  Constraint per overlapping pair:
      # D(p,d) + D(p,d+1) ≤ 1 + D(p,d-1) + D(p,d+2)
      # Together with C15 this enforces day-shift runs of length ≥ 3.
      if (add_c16 && nD >= 3L) {
        dcols <- function(pi, di) vapply(DAY_S, function(s) xidx(pi, di, s), integer(1L))
        for (pi in seq_len(nP)) {
          # Left boundary (d=1, no left neighbour)
          add_con(c(dcols(pi, 1L), dcols(pi, 2L), dcols(pi, 3L)),
                  c(rep(1, 3), rep(1, 3), rep(-1, 3)), "<=", 1)
          # Interior
          if (nD >= 4L) {
            for (di in 2L:(nD - 2L)) {
              add_con(c(dcols(pi, di), dcols(pi, di + 1L),
                        dcols(pi, di - 1L), dcols(pi, di + 2L)),
                      c(rep(1, 3), rep(1, 3), rep(-1, 3), rep(-1, 3)), "<=", 1)
            }
          }
          # Right boundary (d=nD-1, no right neighbour)
          add_con(c(dcols(pi, nD - 1L), dcols(pi, nD), dcols(pi, nD - 2L)),
                  c(rep(1, 3), rep(1, 3), rep(-1, 3)), "<=", 1)
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

      # ── Extra no-good cuts (solution enumeration) ─────────────────────────────
      for (x_prev in extra_nogo) {
        on_idx <- which(x_prev > 0.5)
        if (length(on_idx) > 0L)
          add_con(as.integer(on_idx), rep(1, length(on_idx)), "<=", length(on_idx) - 1L)
      }

      # ── Assemble and solve ────────────────────────────────────────────────────
      nCont <- nF + nW + nNS3 + nNS4 + nWS3 + nWS4
      message(sprintf("  ILP: %d binary + %d continuous, %d constraints",
                      nX, nCont, n_con))

      A <- Matrix::sparseMatrix(i = ri, j = ci, x = vi, dims = c(n_con, nV))

      result <- highs::highs_solve(
        L       = obj,
        lower   = lb,
        upper   = ub,
        A       = A,
        lhs     = con_lhs,
        rhs     = con_rhs,
        types   = types,
        maximum = TRUE,
        control = highs::highs_control(time_limit = 60)
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
      message(sprintf("  Solver: %s  objective = %.1f",
                      result$status_message,
                      if (!is.null(result$objective_value)) result$objective_value else NA_real_))

      list(sol = sol, nP = nP, nD = nD, nX = nX, xidx = xidx, dates_vec = dates_vec)
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

    # ── Diagnostics: run after all tiers fail, then stop with error ────────────
    report_and_stop = function() {
      dates_vec <- as.Date(self$dates, origin = "1970-01-01")
      nP        <- length(STAFF)
      nD        <- length(dates_vec)

      is_blocked <- function(person, d) {
        pdata <- self$time_off[[person]]
        nrow(pdata) > 0 &&
          any(pdata$date == d & pdata$type %in% c("off", "vac", "cme"))
      }

      issues <- character(0)

      # 1. Per-day: count available staff; flag days where < 2 are free
      for (di in seq_len(nD)) {
        d       <- dates_vec[di]
        n_avail <- sum(vapply(STAFF, function(p) !is_blocked(p, d), logical(1L)))
        if (n_avail < 2L) {
          issues <- c(issues, sprintf(
            "[COVERAGE] %s (%s): only %d/%d staff available — need >= 2 for APP1+Night",
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
          n_blocked <- sum(vapply(pp_d, function(d) is_blocked(person, d), logical(1L)))
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
        free    <- vapply(dates_vec, function(d) !is_blocked(person, d), logical(1L))
        max_run <- 0L; run <- 0L
        for (f in free) {
          if (f) { run <- run + 1L; if (run > max_run) max_run <- run } else run <- 0L
        }
        if (max_run < 2L) {
          issues <- c(issues, sprintf(
            "[NIGHTS]   %s: no window of 2+ consecutive free days — C13 unsatisfiable",
            person))
        }
      }

      message("\n  ─── ILP SCHEDULING DIAGNOSTICS ────────────────────────────")
      if (length(issues) == 0L) {
        message("  No obvious availability or target issues detected.")
        message("  The model may simply be too complex for the 2-minute time budget.")
        message("  Consider reducing MAX_NIGHTS_TOTAL, loosening PP targets, or")
        message("  removing time-off entries that conflict with coverage requirements.")
      } else {
        message(sprintf("  %d potential issue(s) found:\n", length(issues)))
        for (iss in issues) message(sprintf("  %s", iss))
      }
      message("  ────────────────────────────────────────────────────────────\n")

      stop("ILP scheduling failed after all relaxation tiers — see diagnostics above.",
           call. = FALSE)
    }
  )
)
