# RDK Audit Playbook

**Purpose.** Documented, repeatable 5-phase audit framework for the RDK financial model. Re-runnable any time the model changes materially, NAIC publishes new factors, or a regulatory milestone approaches.

**When to re-run.**
- **Always** (automated, pre-commit): Phase 1.
- **Annually** or on kernel version bump: Phases 1–5 in full.
- **Before regulatory filing**: Phases 1–5 with trust-level = filing, plus external actuary review.
- **After NAIC publishes new P&C RBC Forecasting Instructions** (usually September): Phase 3 refresh.
- **After a major structural change** to any formula tab: Phases 1, 2, 4 at minimum.

## The 8 categories of findings

Every finding gets classified into one of these and assigned a severity:

| # | Category | Concern |
|---|---|---|
| 1 | Correctness | Does the math match the authoritative source? |
| 2 | Completeness | Is every required input captured and every output produced? |
| 3 | Consistency | Are similar things done the same way across tabs? |
| 4 | Robustness | Does it degrade gracefully at edge inputs? |
| 5 | Auditability | Could an external reviewer trace any number to its source in ≤ 3 clicks? |
| 6 | Operational safety | Will it survive save/load, Setup, concurrent edits? |
| 7 | Compliance | Does it meet the external standards it claims (NAIC, GAAP, SAP)? |
| 8 | Governance | Is institutional knowledge current (docs, patterns, bug log)? |

**Severity levels.**
- **CRITICAL**: wrong answers under realistic inputs. Fix before next use.
- **MAJOR**: structural weakness that will bite. Fix in current audit cycle.
- **MINOR**: paper cut or cosmetic. Fix opportunistically.
- **NOTE**: informational, no action required.

## The 5 phases

### Phase 1 — Static bug-pattern scanner (automated)

**Command.** `python3 scripts/audit_scan.py`

**What it detects.** Four bug classes surfaced during the 2026-04-18/19 RBC audit:
1. `{ROWID:X}` in quarterly formula where X is a single-cell static input → empty at Q2+
2. `{REF:Tab!RowID}` where RowID is undefined → `#REF!`
3. Duplicate `(row, col)` collisions → RowID cache indexes only one
4. Signed-convention mirrors without ABS or leading-minus → sign inversion bugs

**Expected output.** Zero findings = pass. Any non-zero finding blocks merge.

**Effort.** < 1 second runtime. Should be wired as a pre-commit hook and pre-packaging gate.

**Extending.** When a new bug class is discovered, add a `check_<class>()` function following the existing pattern. The `DYNAMIC_TABS`, `DYNAMIC_ROWID_PATTERNS`, and `SIGNED_PATTERNS` constants at the top of the script are the false-positive filters.

### Phase 2 — Tab-by-tab financial-logic review (manual + Agent-assisted)

**Command.** Delegate to an Explore agent with the prompt template in `docs/AUDIT_PLAYBOOK_phase2_prompt.md` (TODO: extract from this doc).

**What it checks per tab.** Six verifications:
1. Accounting identities close (A = L + E; NI → RE; Cash Start + Δ = Cash End)
2. Sign conventions consistent per tab semantics
3. Cross-tab tie-outs (UWEX = Σ UW Program Detail, Net Income ties, etc.)
4. Aggregation consistency (per-period → annual, per-program → company)
5. Algebraic sanity (no obvious div/0, circular refs, over-simplifications)
6. Opaqueness (formulas too compressed to audit)

**Priority order (effort allocation).**
- RBC Capital Model (30 min) — cross-reference to NAIC docs
- Balance Sheet (1 hr) — identity + sign conventions
- Income Statement (1 hr) — line tie-outs + NI closure
- Cash Flow Statement (45 min) — indirect method closes
- UW Exec Summary (30 min) — program aggregates
- UW Program Detail (30 min) — per-program waterfalls
- Investments, Sales Funnel, Revenue Summary, Expense Summary (15 min each)
- Cover Page, Dashboard, User Guide (cosmetic, 5 min)

**Expected output.** Per-tab verdict (CLEAN/MINOR/MAJOR/CRITICAL) + findings with severity, location, suggested fix.

**Effort.** ~4 hours for a thorough first pass. ~1.5 hours for a re-audit.

### Phase 3 — NAIC 2025 factor cross-check (manual, semi-automated)

**Command.**
1. Fetch the current NAIC P&C RBC Forecasting & Instructions (usually updated September). Web-fetch URL: `https://content.naic.org/...` (varies by year).
2. Open `config/rbc_lob_factors.csv` and the `RBC Capital Model` LOB Factor Library (rows 12-26).
3. Diff the 15 LOBs × 6 factor columns against the NAIC tables.

**What it checks.** Full 15 × 6 matrix agreement, not spot-check. Industry RBC, Development Ratio, Investment Income Adjustment for reserves (Lines 4 & 8 of PR017) and the corresponding premium factors (Lines 4 & 7 of PR018).

**Expected output.** Factor drift report: which LOBs changed, by how much, what downstream impact.

**Effort.** ~2 hours once the NAIC doc is in hand. The 2026-04-18 session confirmed 5 of 6 columns match NAIC 2025; R5 InvAdj column was stale.

### Phase 4 — Regression assertions in `KernelTests` (manual)

**Command.** Extend `engine/KernelTests.bas` with `Public Sub TestAuditAssertions()` that runs inside Excel at model-run time. Runs automatically via `KernelTests.RunAllTests`.

**What to assert.**
- Balance Sheet identity: `BS_CHECK` row = 0 at every period
- Income Statement closure: Net Income = sum of components (sanity)
- Cash Flow closure: Opening + Operating + Investing + Financing = Ending
- Quarterly-to-annual: sum of Q1..Q4 = Y Total for flow items; Q4 = Y Total for balance items
- UWEX-to-UW-Program-Detail tie-outs: aggregate = sum of programs (within tolerance)
- PRNG determinism: same seed → same output
- Workspace round-trip: save → load → save produces identical manifest fingerprint
- RBC covariance: Total RBC = R0 + SQRT(R1² + R2² + R3² + R4² + R5² + Rcat²)
- Corp sub-allocation: `SUM($C$103:$C$108)` = 100%

**Expected output.** Failures logged to `Error Log` tab with severity.

**Effort.** ~3 hours for initial implementation. Pays compounding dividends once in place.

### Phase 5 — Consolidated audit report

**Deliverable.** `docs/AUDIT_YYYY-MM-DD.md` structured as:
1. Executive verdict (one paragraph: CLEAN / CONDITIONAL / NEEDS-REWORK)
2. Top-5 risks with severity + effort to close
3. Per-category findings table (sortable by severity, by tab)
4. Per-tab assessment
5. Known-and-accepted (documented deferrals)
6. Regression assertion recommendations
7. Follow-up audit cadence trigger conditions

**Every CRITICAL finding** gets a numbered `BUG-nnn` entry in `data/bug_log.csv` so it routes through the existing institutional-knowledge pipeline.

**Effort.** ~1 hour after Phases 1-4 complete.

## How to trigger a full audit

In a Claude Code session:

> "Run the RDK audit per `docs/AUDIT_PLAYBOOK.md`."

The model should:
1. Run Phase 1 scanner, attach findings.
2. Launch an Explore agent for Phase 2 using the scope above.
3. Fetch current NAIC P&C RBC Instructions for Phase 3.
4. Skim `engine/KernelTests.bas` for Phase 4 coverage; flag missing assertions.
5. Produce `docs/AUDIT_YYYY-MM-DD.md` + append to `data/bug_log.csv` as needed.
6. Commit and push, per the session's usual git discipline.

Expect ~10-12 hours across multiple turns for a full audit. The scanner (Phase 1) and report (Phase 5) are fast; Phases 2-4 are the work.

## Interpreting results

**All-green audit.** Zero CRITICAL + zero MAJOR. Ship.

**Yellow audit.** CRITICAL = 0, MAJOR > 0. Ship with accepted-deferral list; close MAJOR before next regulatory milestone.

**Red audit.** CRITICAL > 0. Do not ship. Fix, then re-run Phase 1 + the specific category that surfaced the CRITICAL.

**Historical comparison.** Each audit report should diff against the prior one — which bugs closed, which new bugs found, which deferrals moved between buckets. Good hygiene: audit reports get committed to `docs/` and never deleted.

## Audit roles

For a single-person project (current Slingshot state):
- **Auditor**: Claude Code, working from this playbook
- **Reviewer**: you, reading the report
- **Fix owner**: whoever has capacity

For a team project or regulatory filing:
- **Auditor**: Claude Code or internal actuarial staff
- **Reviewer**: credentialed P&C actuary (this audit is *necessary* but not *sufficient* for regulatory sign-off)
- **Fix owner**: designated per-finding

## Evolution

This playbook should evolve. When:
- A new bug class is discovered in production → add a Phase 1 check for it.
- A new tab is added to the model → extend Phase 2 priority list.
- NAIC changes methodology → revise Phase 3 procedure.
- A new cross-tab invariant is identified → add to Phase 4 assertions.

Treat the playbook as a living document. Append a "Change log" section at the bottom tracking revisions.

## Change log

- **2026-04-19** (Claude Code): Initial version. Based on the RBC tab bug-hunting sessions of 2026-04-18/19. Phase 1 scanner committed. Phase 2 template in place. Phases 3-5 scoped but not fully automated.
