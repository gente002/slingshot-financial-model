# RDK Developer Flywheel

**Version:** 1.2
**Date:** March 2026
**Status:** Canonical process document. All RDK development follows this flywheel.

---

## The Three God Rules

These are foundational design rules that apply to every system, every module, and every workflow in the RDK. They are not guidelines -- they are constraints. Every design decision must satisfy all three.

### God Rule #1: Never Get Stuck

Use the information you have and make the best assumption available to get to the downstream structure and shape so we can report the same every time.

**Design implications:**
- Always produce output, even with incomplete data
- Always document what was assumed vs what was known
- Every assumption is reversible -- when real data arrives, it replaces the assumption
- The downstream shape is fixed -- upstream variability is resolved before it reaches the output
- Quality issues quarantine rows, they don't halt the pipeline
- Applies to: data ingestion, model computation, formula tabs, reporting, configuration loading

### God Rule #2: Never Lose Control

Users always have a sanctioned path to manually override any value at any layer -- temporarily, with tracking, without breaking the system.

**Design implications:**
- Any value can be overridden: raw data, computed output, system-generated assumptions, configuration-driven defaults
- Overrides are tracked: who, when, why, what was the original value, what was changed
- Overrides persist until explicitly replaced by corrected source data or removed by the user -- they never silently expire
- Overrides that persist beyond a configurable threshold (e.g., 90 days) are flagged for investigation, not auto-removed
- The system doesn't distinguish between source data, automated assumptions, and manual overrides at the output layer -- all produce the same shape
- Override granularity: individual record overrides, bulk overrides (apply one correction to many records), and aggregate overrides (top-side adjustment that allocates down to detail)
- Override hygiene: recurring overrides should be converted to permanent rules or configuration changes
- Applies to: formula tabs (blue input cells), data pipelines (adjustment log), assumptions (user-entered vs system-generated), raw data corrections, any workflow with hard deadlines

### God Rule #3: Never Work Alone

Preview, annotate, converge, confirm. Every data workflow is a multi-party collaboration with real-time feedback, threaded discussion, and progressive convergence toward truth.

**Design implications:**
- Any value at any granularity (cell, row, column, block of rows/columns, or entire dataset) can be tagged with comments and conversation threads
- Changes can be previewed before committing -- including the downstream impact on all consuming systems
- Partners exchange feedback in real-time without waiting for full pipeline reruns
- Data quality issues are grouped into patterns, discussed collaboratively, and resolved incrementally
- Every value is tagged as "actual" (from source) or "assumption" (automated or manual) -- the materiality of assumptions on key metrics is measurable at all times
- Materiality thresholds are configurable per metric (e.g., 0.1% tolerance on premium, 1% on reserves)
- The system tracks convergence: what percentage of values are actual vs assumed, and how that improves over the close cycle
- Applies to: data ingestion workflows, close cycles, partner data exchanges, assumption resolution, any process where two or more parties need to align on data

---

## 1. The Developer Triangle

Three actors, distinct roles, no overlap:

```
                    Claude Online (Opus)
                   /  Architect + Auditor  \
                  /                          \
                 /    STEP 1: Spec            \
                /     STEP 2: Adversarial      \
               /      STEP 5: Review + Capture  \
              /                                   \
    Ethan ─────────────────────────────────────── Claude Code
    Human Validator                                 Autonomous Builder
    STEP 4: Validate                               STEP 3: Build
```

### Claude Online (Opus)
- Produces phase build prompts, config tables, contracts, and specifications.
- Triages adversarial review findings. Makes architectural decisions.
- Reviews Claude Code output for architectural correctness.
- Updates institutional knowledge (anti-patterns, patterns, bug log).
- Generates session handoffs between phases.
- NEVER writes final implementation code. NEVER runs VBA. NEVER tests in Excel.

### ChatGPT (GPT-4o)
- Adversarial reviewer. Reviews phase deliverables before build.
- Evaluates against the stated use case (zero-dependency rapid prototyping).
- Grades findings P0-P3. Scores across 12 dimensions.
- Multiple rounds per phase until convergence.
- NEVER builds. NEVER makes architectural decisions. NEVER triages its own findings.

### Claude Code
- Autonomous builder. Receives reviewed specs and builds.
- Writes .bas files, generates config CSVs, builds Setup scripts, packages ZIPs.
- Iterates locally: write → run → see error → fix → re-run → confirm green.
- Runs KernelLint before packaging (from Phase 4 onward).
- NEVER makes architectural decisions. NEVER modifies the spec. NEVER skips a validation gate.

### Ethan (Human)
- Domain expert and final validator.
- Imports deliverables into Excel. Walks validation gates.
- Provides human judgment: does this look right? Does the UX make sense? Does the domain logic produce reasonable results?
- Makes arbiter decisions when surfaced by Claude Online.
- NEVER writes VBA directly (edits go to .bas files on disk per B.7 disk authority rule).
- Runs Setup.bat to bootstrap workbook from .bas files and config CSVs.

---

## 2. The Five-Step Phase Cycle

Every phase in the build roadmap follows these 5 steps. No steps are skipped. The cycle repeats for each phase (1 through 6).

### Step 1: Spec (Claude Online)

**Input:** Prior phase completion + build roadmap + spec notes.

**Claude Online produces:**
- Phase build prompt (using the appropriate template from Prompts 01-07)
- Config CSV files for this phase (populated for config_sample/, headers-only for config_blank/)
- Any new contracts required (e.g., config table catalog at Phase 4→5A, extension execution contract at Phase 5B→6)
- Domain module specifications (if this phase touches Domain code)
- Exact validation gate checklist with expected values where applicable
- Anti-pattern and pattern reminders relevant to this phase

**Output:** A complete deliverable package that Claude Code could build from without asking questions.

**Quality bar:** If Claude Code would need to make an inference or assumption, the spec is not ready. Every column name, every function signature, every expected value must be explicit.

**Validation gate design (PT-028):** When designing validation gates, classify each check as **automatable** or **manual-only**:
- **Automatable:** Value comparisons, config section existence, module existence, math function outputs, data row counts, column lookups — anything with a deterministic expected value. These MUST be specified with exact expected values so Claude Code can implement them as Tier 5 smoke tests.
- **Manual-only:** Visual rendering (charts, formatting), interactive behavior (button clicks, toggle actions), PDF output inspection, extension deactivation/reactivation (requires re-setup), error handling scenarios (requires deliberate breakage).

The spec should tag each gate check so Claude Code knows which to automate and which to leave as manual walkthrough steps. This dramatically reduces validation time (Phase 6A: 40-50 min manual reduced to 20-25 min with automation).

### Step 2: Adversarial Review (ChatGPT, triaged by Claude Online)

**Input:** Step 1 deliverables (phase build prompt, config CSVs, contracts, validation gates).

**Process:**
1. Ethan sends Step 1 deliverables to ChatGPT with a phase-specific adversarial review prompt.
2. ChatGPT reviews and produces findings graded P0-P3.
3. Ethan sends findings to Claude Online.
4. Claude Online triages: accept (fix), reject (with rationale), or surface to Ethan for arbiter decision.
5. Claude Online updates the Step 1 deliverables with accepted fixes.
6. If any P0 or P1 findings were accepted: send updated deliverables back to ChatGPT for another round.
7. Repeat until: zero P0, zero P1, and the reviewer confirms the deliverables are build-ready.

**Convergence criteria:**
- No P0 findings.
- No P1 findings (or all P1s explicitly accepted as known limitations with documented rationale).
- Reviewer score for the phase deliverables is 8+/10 on Implementability.
- Claude Online confirms: "This is ready for Claude Code."

**Phase-specific review prompt template:**

```
You are reviewing the Phase [N] build deliverables for the Phronex RDK.
This is a zero-dependency rapid prototyping tool (Excel + VBA + CSV + PowerShell).

ATTACHED:
- Phase [N] build prompt (the instructions Claude Code will receive)
- Config CSV files for this phase
- [Any contracts or specifications]
- Validation gate checklist

YOUR TASKS:
1. Can Claude Code build this without making inferences? Flag every ambiguity.
2. Are the config CSVs internally consistent and consistent with the ColumnRegistry?
3. Are the validation gates testable and deterministic?
4. Are there edge cases the build prompt doesn't address?
5. Does anything violate the blessed stack (Excel + VBA + CSV + PowerShell)?
6. Grade each finding P0-P3. Recommend fixes within the blessed stack.
7. Score Implementability 1-10: could a capable VBA developer build this 
   from these materials alone?

Findings requiring technologies outside the blessed stack are INVALID.
```

**Output:** Reviewed, converged deliverables ready for Claude Code.

### Step 3: Build (Claude Code)

**Input:** Reviewed Step 2 deliverables.

**Claude Code executes:**
1. Reads the phase build prompt completely before writing any code.
2. Reads anti_patterns.csv and patterns.csv. Internalizes prevention rules.
3. Reads all config CSVs relevant to this phase.
4. Builds kernel modules per spec. Follows established patterns (PT-001 through PT-010).
5. Builds domain modules per spec (if this phase touches Domain code).
6. Self-checks against anti-patterns before packaging (AP-06, AP-07, AP-08, AP-18, AP-34, AP-35 at minimum).
7. Runs KernelLint if available (Phase 4+).
8. Verifies Sub/Function balance, CRLF line endings, no non-ASCII, module sizes under 50KB.
9. Adds Tier 5 smoke tests to KernelTests.bas for all automatable validation gate checks (PT-028).
10. Verifies Setup.bat runs clean before packaging.
11. Packages deliverable as flat ZIP with date-time stamp.

**Build rules:**
- Use ColIndex() for ALL column references. No magic numbers.
- Use InputValue() for ALL input reads. No hardcoded row/column positions.
- Arrays passed ByRef. Pre-size all arrays. No ReDim Preserve in loops.
- Handle division by zero with If/Then, not IIf with CDbl.
- Log errors via config.LogError(). Never swallow errors silently.
- Atomic file writes: temp → verify → rename.
- CRLF line endings on all .bas files.

**Output:** ZIP containing all .bas files, config CSVs, Setup scripts (if modified), and a brief delivery summary.

### Step 4: Validate (Ethan)

**Input:** Claude Code's ZIP deliverable.

**Ethan executes:**
1. Unzip to working directory.
2. Double-click Setup.bat (if this is Phase 1 or if bootstrap changed).
3. Open the .xlsm workbook.
4. Run `RunProjections` to populate output tabs.
5. Run `RunTests` to execute all automated checks (Tiers 1-5). If any FAIL, stop and report.
6. Run `RunLint` to verify code quality. If any ERROR, stop and report.
7. Walk manual-only validation gate checkboxes (visual, interactive, PDF).
8. If ALL gates pass: phase is complete. Move to Step 5.
9. If ANY gate fails: send failure report to Claude Online for triage.

**Automation-first validation (established Phase 6A):** From Phase 6A onward, validation walkthroughs should only include steps that require human judgment. All deterministic checks (value comparisons, config presence, module existence, math outputs) are covered by Tier 5 smoke tests in KernelTests.bas and verified by `RunTests`. The walkthrough document should explicitly state which checks are automated and which are manual-only. This reduces validation time by ~50% and prevents regression in previously-verified functionality.

**Failure report format:**

```
PHASE [N] VALIDATION FAILURE

Gate: [checkbox text]
Expected: [what should have happened]
Actual: [what happened]
Error (if any): [exact VBA error or Excel behavior]
Screenshot: [if visual issue]
diagnostic_dump.txt: [attached if available, Phase 4+]
```

**If failures exist:** Claude Online triages → produces a fix spec → Claude Code fixes → Ethan re-validates. This inner loop repeats until all gates are green.

**Output:** All validation gates green, or failure report for triage.

### Step 5: Review + Capture (Claude Online)

**Input:** Ethan's validation results (all green) + Claude Code's deliverable.

**Claude Online executes:**
1. Reviews Claude Code's output for architectural correctness (if any questions arose during validation).
2. Updates anti_patterns.csv with any new anti-patterns discovered during build or validation.
3. Updates patterns.csv with any new reusable patterns discovered.
4. Updates bug_log.csv with any bugs found and fixed during the build cycle.
5. Verifies the canonical ledger counts are still accurate (module count, file count, etc.).
6. **Roadmap maintenance (§5.6):** Renumber phases to sequential integers reflecting actual build order. Update `docs/RDK_Phase_Roadmap_v2.0.md` and CLAUDE.md phase table. Archive the old roadmap version. This keeps the roadmap clean as phases are added, split, collapsed, or reordered.
7. Produces the session handoff for the next phase (using Prompt 06 template).
8. Declares phase complete.

**Output:** Updated knowledge base artifacts + phase completion declaration + next phase handoff.

---

## 3. Phase Transition Gates

Between certain phases, contracts must be published before the next phase begins. These are hard gates — not guidelines.

| Transition | Required Contract | Responsible |
|-----------|------------------|-------------|
| Phase 4 → 5A | Config table catalog for Phase 5A/5B tables | Claude Online (Step 1 of Phase 5A) |
| Phase 5B → 6 | Extension execution contract (ordering, failure, data handoff) | Claude Online (Step 1 of Phase 6) |

---

## 4. Adversarial Review Escalation Rules

Not every phase needs the same depth of adversarial review.

| Phase | Review Depth | Rationale |
|-------|-------------|-----------|
| Phase 1: Foundation | FULL (2-3 rounds) | This is the kernel core. Errors here propagate everywhere. |
| Phase 2: Persistence | STANDARD (1-2 rounds) | Important but builds on Phase 1 foundations. |
| Phase 3: Testing | STANDARD (1-2 rounds) | The test framework tests itself — lower specification risk. |
| Phase 4: Observability | LIGHT (1 round) | Diagnostic tooling. Lower architectural risk. |
| Phase 5A: Presentation | STANDARD (1-2 rounds) | New config tables with complex rendering behavior. |
| Phase 5B: Output | STANDARD (1-2 rounds) | Power Pivot integration is the highest-risk item. |
| Phase 6: Extensions | PER-EXTENSION (1 round each for first 3, then batch) | First extensions prove the mechanism; later ones follow the pattern. |

**Escalation triggers (any of these → add a review round):**
- Claude Online is uncertain about a config table schema.
- A new contract is being published for the first time (config catalog, extension contract).
- The phase touches a locked arbiter decision in a way that might change its interpretation.
- Ethan requests deeper review.

---

## 5. Pre-Delivery Checklist (Claude Code)

Before packaging any ZIP, Claude Code verifies:

### Gate 1: File Integrity
- [ ] All .bas files have CRLF line endings
- [ ] No non-ASCII characters in any .bas file
- [ ] ZIP is flat (no wrapper folder)
- [ ] Filename includes date-time stamp with seconds

### Gate 2: VBA Structural
- [ ] Sub/Function count = End Sub/End Function count in every module
- [ ] No IIf() with CDbl() on Variant
- [ ] No compressed syntax (Then: or : Else)
- [ ] No strings starting with =, +, -, @ written to cells via .Value
- [ ] No ReDim Preserve inside loops
- [ ] No sized-array Dim after executable code
- [ ] No two-char variable names starting with i/m/f + uppercase
- [ ] All column references use ColIndex() — no magic numbers
- [ ] All input reads use InputValue() — no hardcoded positions

### Gate 3: Module Size
- [ ] Every module under 50KB

### Gate 4: Anti-Pattern Scan
- [ ] KernelLint clean (Phase 4+), or manual AP check (Phases 1-3)

### Gate 5: Delivery Summary
- [ ] Brief summary included: what was built, what changed, what to test

---

## 6. Knowledge Capture Artifacts

Three machine-readable files maintained across all phases:

### anti_patterns.csv
Append-only. Every AP entry with: ID, Rule, RootCause, PreventionCheck, TriggeringBug, DateAdded.
Claude Code reads at build start. Claude Online appends after each phase.

### patterns.csv
Append-only. Every reusable code pattern with: ID, Name, Category, Description, ExampleModule.
Claude Code follows established patterns. Claude Online appends when new patterns emerge.

### bug_log.csv
Append-only. Every bug with: ID, Description, RootCause, Fix, AntiPattern, Phase, DateFixed.
Claude Online maintains after each validation cycle.

---

## 7. Session Continuity

When a Claude Online session fills its context window:
1. Claude Online generates a handoff using Prompt 06 template.
2. The handoff includes: current phase, step within phase, validation gate status, recent bugs, recent anti-patterns, and 4 verification questions.
3. New session reads the handoff + spec notes + knowledge base artifacts.
4. New session answers the 4 verification questions before proceeding.

When a Claude Code session completes a build:
1. Claude Code packages the ZIP with delivery summary.
2. If issues were encountered: Claude Code documents them in the delivery summary.
3. No handoff needed between Claude Code sessions — each build is self-contained from the reviewed spec.

---

## 8. Process Diagram

```
┌─────────────────────────────────────────────────────────────┐
│                     PHASE N CYCLE                           │
│                                                             │
│  ┌──────────────┐                                           │
│  │  STEP 1:     │  Claude Online produces:                  │
│  │  SPEC        │  - Phase build prompt                     │
│  │              │  - Config CSVs                            │
│  │              │  - Contracts (if gate requires)           │
│  │              │  - Validation gates                       │
│  └──────┬───────┘                                           │
│         │                                                   │
│         ▼                                                   │
│  ┌──────────────┐     ┌──────────────┐                      │
│  │  STEP 2:     │────▶│  ChatGPT     │                      │
│  │  ADVERSARIAL │     │  Reviews     │                      │
│  │  REVIEW      │◀────│  Findings    │                      │
│  │              │     └──────────────┘                      │
│  │  Claude Online                                           │
│  │  triages findings                                        │
│  │  Updates deliverables                                    │
│  │  Repeats until converged                                 │
│  └──────┬───────┘                                           │
│         │  (0 P0, 0 P1, Implementability 8+/10)             │
│         ▼                                                   │
│  ┌──────────────┐                                           │
│  │  STEP 3:     │  Claude Code:                             │
│  │  BUILD       │  - Reads spec + anti-patterns + patterns  │
│  │              │  - Builds autonomously                    │
│  │              │  - Write → test → fix → re-test           │
│  │              │  - Pre-delivery checklist                 │
│  │              │  - Packages ZIP                           │
│  └──────┬───────┘                                           │
│         │                                                   │
│         ▼                                                   │
│  ┌──────────────┐                                           │
│  │  STEP 4:     │  Ethan:                                   │
│  │  VALIDATE    │  - Imports into Excel                     │
│  │              │  - Walks validation gates                 │
│  │              │  - Reports results                        │
│  └──────┬───────┘                                           │
│         │                                                   │
│         ├── ALL GREEN ──▶ Step 5                             │
│         │                                                   │
│         └── FAILURES ──▶ Claude Online triages              │
│                          Claude Code fixes                  │
│                          Ethan re-validates                 │
│                          (inner loop until green)           │
│         │                                                   │
│         ▼                                                   │
│  ┌──────────────┐                                           │
│  │  STEP 5:     │  Claude Online:                           │
│  │  REVIEW +    │  - Reviews architecture                   │
│  │  CAPTURE     │  - Updates anti_patterns.csv              │
│  │              │  - Updates patterns.csv                   │
│  │              │  - Updates bug_log.csv                    │
│  │              │  - Produces next phase handoff            │
│  │              │  - Declares phase complete                │
│  └──────────────┘                                           │
│                                                             │
│  ═══════════════════════════════════════════════════════     │
│  PHASE N COMPLETE → Begin Phase N+1 Step 1                  │
│  (Check phase transition gates first)                       │
└─────────────────────────────────────────────────────────────┘
```

---

## 9. Anti-Pattern Quick Reference (Top 10 for Build)

These are the anti-patterns most likely to be encountered during build. Claude Code should check against these before every delivery.

| AP | Rule | Why It Matters |
|----|------|---------------|
| AP-06 | No non-ASCII in .bas files | VBA import fails silently |
| AP-07 | No strings starting with =,+,-,@ in .Value | Excel interprets as formula |
| AP-08 | Column indices via ColIndex() only | Magic numbers cause schema drift bugs |
| AP-18 | No ReDim Preserve in loops | 100-1000x performance penalty |
| AP-34 | No sized-array Dim after executable code | VBA compile error 10 |
| AP-35 | No 2-char vars starting with i/m/f + uppercase | Conflicts with module-level arrays |
| AP-42 | Domain code never stores cumulative values | Kernel derives cumulative views |
| AP-43 | Every Domain*.bas implements all 4 contract functions | Missing function = FATAL at bootstrap |
| AP-44 | Never overwrite a locked file | Create collision-safe suffix instead |
| AP-45 | No pipeline step may depend on in-memory state without fallback | Config sheet is the fallback source of truth (MBP) |
| AP-46 | Every error message must include manual bypass instructions | Users cannot recover without knowing what to fix and where to resume (MBP) |
| AP-47 | Display mode toggle must never change Detail tab data | Detail tab is single source of truth; display is a view layer concern (MBP) |
| AP-48 | Kill zombie Excel processes before COM automation | Prior failed runs leave hidden Excel holding COM/file locks |
| AP-49 | Never hide Excel during macro-heavy COM automation | Hidden Excel swallows error dialogs; script appears to hang |
| AP-50 | All cell writes of strings starting with = must set NumberFormat=@ | Extends AP-07 to all code paths, not just config loader |
| AP-51 | Bootstrap must ensure required kernel tabs exist regardless of tab_registry | Blank config has no tab entries; downstream code crashes on missing sheets |
| AP-52 | VBA Immediate Window calls must use Application.Run with workbook qualifier | PERSONAL.XLSB context causes Error 424 on unqualified module calls |
| AP-53 | Pipeline entry points must show MsgBox for all outcomes | Silent success/failure leaves users with no feedback |
| AP-54 | Bootstrap zombie-kill and /automation must not corrupt COM add-in state | Killing Excel mid-run corrupts S&P DPAPI storage; /automation + Quit writes bad HKCU stubs |
| PT-001 | Array batch write to Excel | Cell-by-cell is 100-1000x slower |

---

## 10. Design Standards

### DS-01: Run Model Pipeline (Current Architecture -- ABANDONED formula-first)

The RDK uses a VBA-driven computation pipeline. The user clicks Run Model to execute the domain engine. This is the proven, working architecture.

**Formula-First UDF approach was attempted and abandoned (April 2026).** The approach used VBA UDF functions (RDK_Value) called from Excel formula cells, with an in-memory compute cache (KernelComputeCache). It failed due to fundamental VBA UDF performance constraints: Excel throttles worksheet reads in UDF context (~10-100x slower), volatile UDFs trigger 600+ VBA round-trips per recalc, and bootstrap COM automation crashes when domain engine runs in UDF context. 13 bugs logged (BUG-136 through BUG-148), none fully resolved. Do not retry without solving the VBA UDF sandbox constraints first. Consider COM add-in, Office.js, or Python xlwings for a future attempt.

The current architecture works reliably. Formula tabs downstream of QuarterlySummary auto-recalculate when QS changes. The user clicks Run Model once after changing inputs.

DELETED: The following text was the abandoned DS-01 formula-first spec. Retained here as a historical note. Do not implement. The workbook computes via UDF-powered Excel formulas that call the same domain computation code as Run Model. No approximations. No simplified assumptions. Exact match between formula output and VBA detail output.

**Architecture:** UDFs (VBA functions callable from Excel cells) serve as the API layer. The domain engine computes in memory, results are cached, and formula cells read from the cache. Changing any input auto-triggers recomputation via Excel's volatile recalculation.

```
User edits input cell
  -> Excel recalculates volatile UDFs
  -> First UDF call: hash inputs, recompute if changed (debounce)
  -> Domain engine runs full computation in memory
  -> Results cached in VBA Dictionary
  -> All UDF cells read from cache (instant O(1) lookup)
  -> Downstream formula tabs reference UWEX RowIDs (unchanged)
```

**An optional "Generate Detail" button** runs the full pipeline, writes Detail tab + CSV + Triangles, and reconciles against the cached formula values. If they diverge, the model flags it.

| Layer | Purpose | Authority | Required? |
|-------|---------|-----------|-----------|
| Formula model (UDF-driven) | Financial statements, auto-compute | Single source of truth | YES |
| VBA detail generation | Monthly CSV, triangles, audit artifacts | Must reconcile to formula model | OPTIONAL |
| Reconciliation | Compare formula cache vs Detail/QS values | Zero-tolerance validation | Required when detail generated |

**Rules:**
1. Every model ships with both `config_{name}` (Run Model) and `config_{name}_formula` (auto-compute)
2. Both configs use the SAME domain computation code — no approximations, no shortcuts
3. The formula config uses `RDK_Value()` UDFs that call the domain engine via `KernelComputeCache`
4. The Run Model config uses the traditional pipeline (Detail -> QuarterlySummary -> formula tabs)
5. Reconciliation must be zero-tolerance: if formula and detail diverge, it's a bug, not an acceptable difference
6. No "lite" or "simplified" versions — if calculations differ, fix them until they match

**Compute cache architecture (KernelComputeCache.bas):**
- `RDK_Value(rowID, quarterIndex)` — volatile UDF, reads from cache
- `RDK_Compute()` — volatile UDF, triggers computation sentinel
- `QIdx(colNum)` — helper UDF, converts Excel column to quarter index
- Timer-based debounce: only recomputes if >50ms since last compute AND input hash changed
- Domain dispatch via config: `Application.Run domMod & "_FormulaAPI.ComputeToCache"`

**Why this pattern:**
- Single source of computation code — domain engine runs identically in both modes
- No dual-maintenance burden — same math, different output target (cache vs Detail tab)
- Formula model recalculates on input change — no button click needed
- Detail generation is voluntary — for CSV export, triangles, and audit
- Reconciliation proves correctness — if formula cache and Detail/QS don't match, there's a bug

### DS-02: Config-Driven Workbook Identity

Every model config includes a `WorkbookName` setting in branding_config.csv. Setup.bat reads this to name the output .xlsm file. This allows multiple model variants to coexist in the same workbook/ directory without filename clashes.

### DS-04: Auto-Refresh with Input Hash Debounce

UDFs are marked `Application.Volatile` so Excel recalculates them on every calc cycle. To avoid redundant computation:

1. First UDF call per cycle checks `Timer - m_lastCompute > 0.05` (50ms debounce)
2. If triggered, computes a hash of sentinel input cells (~10ms)
3. If hash differs from cached hash, runs full domain computation (~500ms)
4. Updates cache + hash + timer
5. All subsequent UDF calls in the same cycle skip steps 1-4 (debounce window)

This means: one computation per committed input change, hundreds of instant cache lookups. User experience is identical to pure Excel formulas — change a cell, everything updates.

### DS-03: Authorship Fingerprinting

Every RDK deliverable includes:
- Copyright header in all .bas files
- LICENSE.txt at repo root
- Build fingerprint constants in KernelConstants (KERNEL_AUTHOR, KERNEL_BUILD_ID)
- VeryHidden fingerprint sheet stamped at bootstrap
- SHA-256 hash of author name embedded in 4 independent hidden locations

New .bas files must include the copyright header. StampFingerprint runs automatically at bootstrap.

---

## Appendix A: Phase 1 Build Findings

### A.1 Bugs Discovered and Fixed

| Bug | Description | Root Cause | Fix | Anti-Pattern |
|-----|-------------|------------|-----|--------------|
| BUG-001 | .bas files had LF-only line endings | Write tool on Windows/bash outputs LF; VBA import requires CRLF | CRLF conversion pass in Setup | AP-06 |
| BUG-002 | Config sheet columns used magic numbers 1-9 | Build prompt did not define constants for config-sheet layout | Added CREG_COL_*, ISCH_COL_*, GCFG_COL_*, TREG_COL_* to KernelConstants | AP-08 |
| BUG-003 | Summary tab used magic numbers for layout | Summary column positions not defined as constants | Added SUMMARY_COL_* constants | AP-08 |
| BUG-004 | Setup hung indefinitely at BootstrapWorkbook | Zombie Excel process from prior failed run; hidden Excel swallowed dialogs | Zombie cleanup at startup + Visible=True during bootstrap | AP-48, AP-49 |
| BUG-005 | Error 1004 writing `=== ATTRIBUTES ===` | String starting with `=` interpreted as formula by Excel | Set NumberFormat="@" before writing section headers | AP-07, AP-50 |
| BUG-006 | Bootstrap crash with blank config (Error 9: Subscript out of range) | CreateTabsFromRegistry relied on tab_registry data rows; blank config has none | Added EnsureSheet fallback for all required kernel tabs | AP-51 |
| BUG-007 | ResumeFrom Error 424: Object required from Immediate Window | Immediate Window ran in PERSONAL.XLSB context, not RDK_Model | Use Application.Run with workbook-qualified name | AP-52 |
| BUG-008 | RunProjections validation failure produced no visible feedback | Only logged to ErrorLog; no MsgBox shown to user | Added MsgBox for validation failure, success, and unhandled errors | AP-53 |

### A.2 Error Patterns

These recurring error patterns were observed during Phase 1 and should be watched for in future phases:

| Pattern | Description | Mitigation |
|---------|-------------|------------|
| **Silent COM hang** | Excel COM automation appears frozen with no error output. Caused by invisible dialog boxes or zombie processes. | Always run Excel visible during automation. Kill stale processes at startup (PT-011). |
| **Formula injection in display strings** | Any string starting with `=`, `+`, `-`, `@` written to a cell via `.Value` is interpreted as a formula. AP-07 was only enforced in the config loader, not in bootstrap display code. | Enforce AP-47 globally: every `.Value` write of a user-facing string must be preceded by `NumberFormat = "@"` if the string could start with a dangerous character. |
| **Magic number creep** | Even with ColIndex() enforced for data columns, layout constants (header rows, start columns, config sheet column positions) were initially hardcoded. | Extend AP-08 to cover ALL numeric positions, not just data column indices. Define constants for every row/column position. |
| **Config-dependent structural assumptions** | Bootstrap assumed tab_registry would always have data rows to create required tabs. Blank config (headers only) caused cascading failures in downstream steps. | Structural requirements (which tabs must exist) must be hardcoded as fallbacks, not solely derived from config data. Config drives customization; structure must be guaranteed. |
| **VBA project context mismatch** | When PERSONAL.XLSB or other workbooks are open, the VBA Immediate Window defaults to whichever project is active. Unqualified module references (e.g., `KernelEngine.ResumeFrom`) fail with Error 424. | All documentation and user-facing VBA calls must use `Application.Run "WorkbookName!Module.Sub"` syntax. |
| **Silent pipeline outcomes** | Pipeline entry points logged results to ErrorLog but showed no visible UI feedback. Users had to manually check ErrorLog to know if anything happened. | All user-facing entry points must show a MsgBox on completion (success or failure). |

### A.3 Design Decisions

| Decision | Rationale |
|----------|-----------|
| Domain calls KernelConfig directly (no Object wrapper) | VBA standard modules cannot be passed as Object references. Direct calls work within a single VBA project. |
| LogError placed in KernelConfig | Spec calls for `config.LogError()` in domain interface. KernelConfig is the natural home. |
| In-memory seed (no .xlsx on disk) | Unified Setup creates workbook via `Workbooks.Add()` and saves directly as .xlsm. Eliminates intermediate seed file and separate script. |
| Scripting.Dictionary for ColIndex lookups | Implemented for all sizes since overhead is negligible and it future-proofs for 100+ column registries. |
| Quoted CSV parser handles embedded commas | Config CSVs use quoted fields. Naive `Split(",")` would break on DerivationRule values if they contained commas. |
| File-lock check instead of process kill | Setup.bat no longer kills all Excel processes (AP-48 relaxed). Instead, Bootstrap.ps1 checks if RDK_Model.xlsm is file-locked before COM automation. Other workbooks stay open. |
| Config sheet Hidden (not VeryHidden) | Config set to xlSheetHidden (accessible via right-click → Unhide) instead of xlSheetVeryHidden. Users need to inspect config during debugging; VeryHidden is overkill for parsed CSV data. |
| MsgBox in pipeline entry points | AP-19 originally prohibited MsgBox in kernel code. Relaxed for RunProjections/ResumeFrom completion and error paths, since these are user-invoked entry points that need visible feedback. |
| Severity color-coding in ErrorLog | LogError now color-codes the Severity cell (FATAL=red, ERROR=pink, WARN=yellow, INFO=green). Makes error scanning immediate without reading text. |

### A.4 Script Consolidation

The original 4-file script setup (Setup.ps1, Setup.bat, CreateSeedWorkbook.ps1, CreateSeedWorkbook.bat) was consolidated into 2 files:

| Before | After | Change |
|--------|-------|--------|
| CreateSeedWorkbook.ps1 | *(removed)* | Seed creation merged into Bootstrap.ps1 (in-memory) |
| CreateSeedWorkbook.bat | *(removed)* | Entry point merged into Setup.bat |
| Setup.ps1 | Bootstrap.ps1 | Renamed; seed creation added; config selection moved to Setup.bat |
| Setup.bat | Setup.bat | Now handles: zombie cleanup, config seeder menu, calls Bootstrap.ps1 |

### A.5 Amendment (v1.0.1) Changes

The Phase 1 amendment applied the Manual Bypass Protocol (MBP) and audit fixes:

| Change | Description |
|--------|-------------|
| AP renumbering | AP-45/46/47 (zombie, visible, NumberFormat) renumbered to AP-48/49/50. New AP-45/46/47 reserved for MBP protocol. |
| Execute signature fix | Removed dead `inputs()` parameter from `SampleDomainEngine.Execute`. Domain reads inputs via `KernelConfig.InputValue()`. |
| KernelConfig fallback mode | If `LoadAllConfig` fails, every getter reads directly from the Config sheet (AP-45 compliance). |
| Pipeline step markers | Each pipeline step writes status (COMPLETE/FAILED/SKIPPED/BYPASSED) + timestamp to the Config sheet PIPELINE_STATE section. |
| ResumeFrom | New `KernelEngine.ResumeFrom(stepNumber)` entry point validates prior artifacts and resumes the pipeline mid-flight. |
| Bypass messages | All `LogError` calls with SEV_ERROR or SEV_FATAL now include "MANUAL BYPASS:" instructions (AP-46 compliance). |
| ExportSchemaTemplate | New utility writes `detail_template.csv` (headers-only) and `detail_template.txt` (human-readable reference). |
| DomainEngine.bas stub | Empty 4-function contract stub for developers to copy. |
| Fixture corrections | Period 12 revenue values and total revenue corrected in DELIVERY_SUMMARY.md. |

### A.6 Validation Session Changes (v1.0.1 → v1.0.2)

Changes made during Ethan's Excel validation walkthrough:

| Change | Description |
|--------|-------------|
| Setup.bat: non-destructive Excel check | Replaced taskkill with file-lock warning. Other workbooks stay open. |
| Bootstrap.ps1: file-lock guard | Tests exclusive file access on RDK_Model.xlsm before COM automation. Exits with clear message if locked. |
| KernelBootstrap: required-tab fallback | CreateTabsFromRegistry now EnsureSheet's all 5 required kernel tabs after registry loop. |
| KernelBootstrap: blank config guard | SetupDetailHeaders early-exits when GetColumnCount()=0 instead of crashing on zero-size array. |
| KernelBootstrap: Config visibility | Changed from xlSheetVeryHidden to xlSheetHidden (right-click → Unhide accessible). |
| KernelEngine: MsgBox feedback | RunProjections shows MsgBox on validation failure, success, and unhandled errors. |
| KernelConfig: severity colors | LogError color-codes Severity cell (FATAL=red, ERROR=pink, WARN=yellow, INFO=green). |
| Walkthrough: Application.Run syntax | All Immediate Window calls use workbook-qualified Application.Run to avoid PERSONAL.XLSB context issues. |

### A.7 Knowledge Artifact Counts (Post Phase 1 + Validation)

| Artifact | Count |
|----------|-------|
| Anti-patterns (AP-01 to AP-53) | 53 |
| Patterns (PT-001 to PT-018) | 18 |
| Bugs (BUG-001 to BUG-008) | 8 |
| VBA modules | 8 (7 kernel/domain + 1 stub) |
| Config CSVs (per seeder) | 4 |
| Scripts | 2 |

---

## Appendix B: Phase 2 Build Findings

### B.1 Bugs Discovered and Fixed

| Bug | Description | Root Cause | Fix | Anti-Pattern |
|-----|-------------|------------|-----|--------------|
| BUG-009 | RunProjections macro missing after Setup (VBA import silently failed) | KernelEngine.bas had Unicode em-dash characters on lines 64 and 164 | Replaced em-dashes with ASCII double-dash (--) | AP-06 |
| BUG-010 | RunProjections not visible in Alt+F8 macro dialog | VBA Optional parameters hide macros from the Alt+F8 dialog | Added parameterless RunProjections() wrapper delegating to RunProjectionsEx(Optional) | (new pattern) |
| BUG-011 | Optional parameter default referenced cross-module Public Const | VBA does not allow cross-module constant references as Optional defaults | Replaced COMPARE_DEFAULT_THRESHOLD with literal 0.000001 in all signatures | AP-06 (adjacent) |
| BUG-012 | PRNG seed overwritten by RunProjections in AUTO mode | InitializePRNG always called AutoSeed when DeterministicMode=AUTO and DefaultSeed=0 | Added KernelRandom.IsInitialized check in AUTO branch | AP-45 |
| BUG-013 | ShowScenarios only showed 1 scenario | Nested Dir() calls reset the outer Dir() iteration | Replaced nested Dir() with FileSystemObject | (new pattern) |
| BUG-014 | RestoreSavepoint crashed with Subscript out of range | Config arrays not loaded when RestoreSavepoint called standalone | Added KernelConfig.LoadAllConfig at start of restore/load | AP-45 |
| BUG-015 | ArchiveSavepoint failed on re-archive | Archive destination had read-only files from prior archive | Pre-archive cleanup: detect, clear read-only, delete before copy | (new pattern) |
| BUG-016 | Bootstrap.ps1 COM error 80080005 with COM add-ins active | S&P Capital IQ COM add-in interfered with programmatic Excel startup | Launch Excel with /automation flag to suppress add-in loading | AP-48 |
| BUG-017 | S&P Capital IQ COM add-in permanently broken after Setup.bat | Three-link chain: (1) zombie kill corrupted DPAPI IsolatedStorage, (2) /automation + Quit wrote bad HKCU stubs overriding HKLM, (3) failed S&P auto-update compounded damage | Clear Roaming IsolatedStorage, delete HKCU stubs, harden Bootstrap.ps1 finally block, created Repair-ComAddins.ps1 | AP-54 |

### B.2 Error Patterns

| Pattern | Description | Mitigation |
|---------|-------------|------------|
| **VBA Optional parameter traps** | Optional parameters with defaults hide macros from Alt+F8 dialog. Cross-module Const references in Optional defaults cause compile errors. | Always use parameterless wrappers for user-facing entry points. Use literal values, not Const references, for Optional defaults. |
| **Nested Dir() corruption** | Calling Dir() inside a Dir() loop resets the outer iteration. Only the first matching folder is returned. | Use Scripting.FileSystemObject for any nested directory enumeration. |
| **Standalone entry point crashes** | Functions like RestoreSavepoint called outside the normal pipeline don't have config arrays loaded. | Every public entry point must call LoadAllConfig at start (AP-45 fallback). |
| **COM add-in state corruption** | Killing Excel mid-run corrupts DPAPI-encrypted IsolatedStorage. The /automation flag + Quit cycle writes incomplete HKCU registry stubs. Both permanently disable COM add-ins. | Bootstrap.ps1 guards: zombie kill only targets truly headless processes (MainWindowHandle=0), finally block deletes HKCU stubs. Repair-ComAddins.ps1 for recovery. |
| **Read-only file overwrite** | FileCopy cannot overwrite files with the read-only attribute set. Re-archiving a savepoint fails because prior archive set files read-only. | Always clear read-only attributes before overwriting. Check for and remove existing target files before copy operations. |

### B.3 Design Decisions

| Decision | Rationale |
|----------|-----------|
| Mersenne Twister (MT19937) in pure VBA | Industry-standard PRNG. Unsigned 32-bit arithmetic via Double intermediates because VBA Long is signed 32-bit. Same seed = bit-identical output across runs. |
| KernelScenarios + KernelState consolidated into KernelSnapshot | Original spec had separate modules. Consolidated to reduce cross-module coupling and keep scenario/savepoint storage patterns unified. |
| Manifest/metadata written LAST with STATUS=COMPLETE | Atomic save pattern: if any prior file write fails, the manifest never gets written, so the savepoint/scenario is recognized as incomplete on restore. |
| SHA256 via PowerShell Get-FileHash | VBA has no native SHA256. Shelling to PowerShell is standard on Windows 10+ and avoids bringing in external libraries. |
| File-based locking with 60-second stale detection | Prevents concurrent writes to savepoint directories. Stale lock detection avoids permanent lockout from crashed processes. |
| KernelFormHelpers + KernelFormSetup for UI | Separated form/button setup from kernel logic. FormHelpers provides interactive prompts (InputBox, list selectors). FormSetup wires buttons to ScenarioMgr tab. |
| /automation flag for Excel COM launch | Prevents COM add-in interference (BUG-016) without requiring add-in uninstallation. Session-only suppression. |
| HKCU stub deletion (not value-setting) for add-in repair | Setting LoadBehavior=3 on HKCU stubs doesn't work because the stubs lack Manifest/FriendlyName. Deleting the stubs lets the complete HKLM entries take effect. |

### B.4 Script Changes

| Script | Change |
|--------|--------|
| Bootstrap.ps1 | Added /automation Excel launch (BUG-016). Added COM add-in protection: snapshots HKCU state before launch, deletes new stubs in finally block (AP-54). Hardened zombie kill to require MainWindowHandle=0. |
| Toggle-ComAddins.ps1 | New script. Enables (-Enable), disables (-Disable), or repairs (-Repair) S&P COM add-ins via HKCU registry manipulation. |
| Repair-ComAddins.ps1 | New standalone script. Diagnoses (-Diagnose) or repairs S&P add-in issues: clears corrupted Roaming IsolatedStorage, Local IsolatedStorage, SNL Office cache, and HKCU stubs. Optional MSI repair with -Full. Works from any directory. |

### B.5 Knowledge Artifact Counts (Post Phase 2 + Validation)

| Artifact | Count |
|----------|-------|
| Anti-patterns (AP-01 to AP-54) | 54 |
| Patterns (PT-001 to PT-018) | 18 |
| Bugs (BUG-001 to BUG-017) | 17 |
| VBA modules | 13 (11 kernel + 1 domain sample + 1 domain stub) |
| Config CSVs (per seeder) | 6 |
| Scripts | 4 (Setup.bat, Bootstrap.ps1, Toggle-ComAddins.ps1, Repair-ComAddins.ps1) |

---

## Appendix C: Phase 3 Build Findings

### C.1 Bugs Discovered and Fixed

| Bug | Description | Root Cause | Fix | Anti-Pattern |
|-----|-------------|------------|-----|--------------|
| BUG-018 | Domain tests DOM-003/DOM-004 fail: entities B and C return entity A Revenue | SetupTestInputs wrote InputSchema defaults (single value per parameter) to ALL entities, overwriting per-entity fixture data | Removed default-writing loop from SetupTestInputs; added explicit per-entity fixture restore | AP-55 |

### C.2 Error Patterns

| Pattern | Description | Mitigation |
|---------|-------------|------------|
| **Schema default flattening** | InputSchema stores one default value per parameter. Any helper that writes defaults across all entity columns destroys per-entity test fixtures set by bootstrap. | Test helpers must only apply explicit overrides, never broadcast schema defaults across entities. |

### C.3 Design Decisions

| Decision | Rationale |
|----------|-----------|
| 5-tier test framework (Unit, Edge, Integration, Regression, Exhibit) | Provides natural progression from isolated checks through full system verification. Each tier depends on prior tiers passing. |
| Golden CSV baseline comparison | Deterministic fixture produces identical output. SaveGolden captures baseline; RunTests compares cell-by-cell within tolerance (0.000001). Eliminates manual inspection. |
| Prove-It auditor checks (Identity, Accumulate, Reconcile) | Financial model verification patterns. Identity: A - B - C = 0. Accumulate: sum of parts = total. Reconcile: two independent calculations match. |
| KernelTestHarness as orchestrator | Separates test infrastructure from test cases. Harness handles tier ordering, result writing, golden comparison. Domain tests implement the cases. |
| Explicit fixture restore in domain tests (PT-022) | Tests that modify inputs must restore all per-entity values at cleanup. Avoids cross-test contamination without requiring re-bootstrap. |

### C.4 Knowledge Artifact Counts (Post Phase 3)

| Artifact | Count |
|----------|-------|
| Anti-patterns (AP-01 to AP-55) | 55 |
| Patterns (PT-001 to PT-022) | 22 |
| Bugs (BUG-001 to BUG-018) | 18 |
| VBA modules | 16 (14 kernel + 1 domain sample + 1 domain stub) |
| Config CSVs (per seeder) | 7 |

---

## Appendix D: Phase 4 Build Findings

### D.1 Bugs Discovered and Fixed

| Bug | Description | Root Cause | Fix | Anti-Pattern |
|-----|-------------|------------|-----|--------------|
| BUG-019 | RunLint fails with "Engine directory not found" | KernelLint used ThisWorkbook.Path & \engine but workbook lives in workbook/ subdirectory | Changed to ThisWorkbook.Path & \..\engine in KernelLint, KernelDiagnostic, KernelHealth | AP-22 |
| BUG-020 | LINT-03 only caught magic numbers in outputs() pattern | CheckMagicNumbers missed direct col* = <number> assignments | Added MatchesColAssignmentMagic() detector | AP-08 |
| BUG-021 | LINT-01/LINT-12 unreliable non-ASCII and CRLF detection | StrConv codepage conversion remapped byte values | Rewrote to scan raw byte array directly | AP-06 |
| BUG-022 | MODULE_SIZE_WARN too high to catch 51KB modules | Threshold at 51200 bytes; KernelConfig was 51130 bytes (70 bytes under) | Lowered MODULE_SIZE_WARN from 51200 to 50000 bytes | (threshold calibration) |
| BUG-023 | MatchesMagicNumberPattern infinite loop | Fallback InStr used hardcoded position 1 instead of advancing pos | Changed InStr(1,...) to InStr(pos,...) | (new pattern) |
| BUG-024 | KernelDiagnostic ByRef type mismatch in BuildConfigSummary | Passed Long cidx where String parameter expected | Used intermediate colName = GetColName(cidx) | AP-56 |
| BUG-025 | Bootstrap.ps1 closes Excel without save prompt for other workbooks | DisplayAlerts=$false never restored before Quit() | Added DisplayAlerts=$true and Visible=$true in finally block | AP-58 |

### D.2 Error Patterns

| Pattern | Description | Mitigation |
|---------|-------------|------------|
| **Relative path assumption** | Workbook in workbook/ subdir means ThisWorkbook.Path is not the project root. All engine/ and config/ references must go up one level. | Use ThisWorkbook.Path & "\.." for project root. |
| **Lint self-detection** | Lint patterns that scan for keywords (REDIM PRESERVE, magic numbers) match string literals in the lint module itself. | Use Left() for statement-start matching; add IsInsideString() guard. |
| **Byte-level scanning** | VBA StrConv between byte arrays and strings is lossy for non-ASCII detection. Raw byte scanning is the only reliable method. | Always use raw byte arrays for file encoding checks. |
| **InStr loop stalls** | InStr with a fixed start position inside a While loop creates infinite loops when the pattern exists but doesn't meet secondary criteria. | Always advance the search position past each match. |

### D.3 Design Decisions

| Decision | Rationale |
|----------|-----------|
| Dual-source lint (disk + VBA project, PT-023) | Disk scan catches file-level issues (CRLF, size). VBA project scan catches in-editor changes not yet exported. Results prefixed [VBA] to distinguish. |
| Health check two modes (LIGHTWEIGHT, FULL) | LIGHTWEIGHT runs at bootstrap (fast, non-blocking). FULL runs on demand for deeper diagnostics. |
| Diagnostic dump writes to text file | Avoids recursion risk of writing diagnostics to the same ErrorLog being diagnosed. Text file is portable and attachable to bug reports. |
| Raw byte scanning for LINT-01/LINT-12 | After BUG-021, abandoned StrConv approach entirely. Direct byte comparison is both simpler and more reliable. |

### D.4 Knowledge Artifact Counts (Post Phase 4)

| Artifact | Count |
|----------|-------|
| Anti-patterns (AP-01 to AP-58) | 58 |
| Patterns (PT-001 to PT-023) | 23 |
| Bugs (BUG-001 to BUG-025) | 25 |
| VBA modules | 18 (16 kernel + 1 domain sample + 1 domain stub) |
| Config CSVs (per seeder) | 7 |

---

## Appendix E: Phase 5A Build Findings

### E.1 Bugs Discovered and Fixed

| Bug | Description | Root Cause | Fix | Anti-Pattern |
|-----|-------------|------------|-----|--------------|
| BUG-026 | Exhibits used hardcoded .Value writes, not formulas | GenerateExhibits wrote static values that didn't update when Detail changed | Rewrote to SUMIFS formulas (Incremental) and cell-reference formulas (Derived) | AP-59 |
| BUG-027 | ProveIt tab orphan red cell past last data row | Blank-row gap before AND summary left empty cell matching FALSE conditional format | Removed blank-row gap before AND summary | (off-by-one) |
| BUG-028 | Chart generation silently skipped charts with invalid metrics | If metricCol < 1 Then GoTo NextChart with no logging | Added per-chart error logging E-521 with manual bypass and skip counter | AP-46 |
| BUG-029 | LINT-03 false positive on colCount variable in KernelTests | colCount matched col* = <number> lint pattern | Renamed colCount to numCols (5 locations) | AP-08 (lint false positive) |
| BUG-030 | Exhibit tab blanket format overrode per-metric formats | Single Format column in exhibit_config applied to all metrics including GPMargin | Built per-metric format array from summary_config -> column_registry -> exhibit fallback | (format override) |
| BUG-031 | chart_registry.csv left with FakeMetric from gate testing | User edited MetricName for validation; change persisted in config_sample/ | Restored Revenue in source file | AP-60 |
| BUG-032 | No run-level error summary visible to user | Errors logged individually to ErrorLog but user got plain "success" MsgBox | Added run-level FATAL/ERROR counters with consolidated vbExclamation alert | PT-025 |
| BUG-033 | No display mode indicator on Dashboard | Users couldn't tell if workbook was in Incremental or Cumulative mode | Added row 2 status cell with color coding + dynamic toggle button caption | PT-026 |

### E.2 Error Patterns

| Pattern | Description | Mitigation |
|---------|-------------|------------|
| **Static exhibit snapshots** | Writing computed values instead of formulas creates exhibits that go stale when Detail changes. Same issue applies to any presentation tab. | All presentation data cells must use formulas referencing the source (Detail or CumulativeView). |
| **Blanket formatting** | Applying a single format to mixed metric types (currency, percentages, counts) produces incorrect display. | Build per-metric format lookup from config hierarchy. |
| **Test artifact leakage** | Config modifications for validation testing can persist in source files (config_sample/) and propagate to all future setups. | Always verify config_sample/ matches production values before packaging. |

### E.3 Design Decisions

| Decision | Rationale |
|----------|-----------|
| CumulativeView as hidden formula sheet | Running SUM formulas on a separate sheet preserves Detail as single source of truth (AP-47). Toggle switches Summary/Charts/Exhibits source reference. |
| Generic 2D section loader (LoadSection2D) | Summary, chart, exhibit configs share a common 2D String array loader with per-table column constants. Saved ~6KB vs separate loaders, keeping KernelConfig under 64KB. |
| Charts use array data read (PT-001) | Detail data read into Variant array, then chart series built from array values. More reliable than range-based chart data across tab switches. |
| Formula-driven exhibits (PT-024) | SUMIFS for Incremental, cell-reference for Derived. Matches Summary pattern. Enables Excel formula auditing (Ctrl+`). |
| Run-level error counters (PT-025) | Individual MsgBox per error is disruptive. Consolidated alert at run end shows count + pointer to ErrorLog. |

### E.4 Knowledge Artifact Counts (Post Phase 5A)

| Artifact | Count |
|----------|-------|
| Anti-patterns (AP-01 to AP-60) | 60 |
| Patterns (PT-001 to PT-026) | 26 |
| Bugs (BUG-001 to BUG-033) | 33 |
| VBA modules | 19 (17 kernel + 1 domain sample + 1 domain stub) |
| Config CSVs (per seeder) | 11 |

---

## Appendix F: Phase 5B Build Findings

### F.1 Bugs Discovered and Fixed

| Bug | Description | Root Cause | Fix | Anti-Pattern |
|-----|-------------|------------|-----|--------------|
| BUG-034 | Transform E-700: "Cannot run the macro 'SampleDomainEngine.ApplyLoading'" | Application.Run cannot pass VBA arrays as parameters. RunTransforms passed outputs() directly to Application.Run, which silently fails. | Added Public TransformOutputs handoff variable. RunTransforms copies outputs into it before calling transforms (no parameter), copies back after. Transform functions access KernelTransform.TransformOutputs directly. | (VBA Application.Run limitation) |
| BUG-035 | VBA project contained duplicate "SampleDomainEngine1" module after Setup | SampleDomainTests.bas had incorrect Attribute VB_Name = "SampleDomainEngine" (same as SampleDomainEngine.bas). Both matched the SampleDomain* import filter. VBA auto-renamed the second import. | Deleted SampleDomainTests.bas (was a stale corrupt copy with syntax error on line 131). | AP-06 (adjacent: file integrity) |
| BUG-036 | Personal.xlsb disabled after running Setup.bat | The /automation flag suppresses XLSTART loading. When the /automation Excel instance quits, it may write entries to Resiliency\DisabledItems that prevent XLSTART items from loading on next normal launch. | Added XLSTART protection to Bootstrap.ps1: snapshot DisabledItems registry key before /automation launch, remove new entries in finally block (same pattern as COM add-in protection). | AP-54 (extended) |

### F.2 Error Patterns

| Pattern | Description | Mitigation |
|---------|-------------|------------|
| **Application.Run array limitation** | VBA's Application.Run cannot pass arrays (Variant() or typed arrays) as arguments. The call either fails with error 1004 or silently drops the argument. This is a documented VBA limitation. | Use a Public module-level Variant as a handoff. Caller copies array into it before Application.Run (no args), callee reads/writes it directly, caller copies back after. |
| **VB_Name collision in .bas files** | If two .bas files have the same Attribute VB_Name, VBA import auto-renames the second with a numeric suffix (Module1 -> Module11). No warning is shown. The renamed module has wrong internal references. | Verify all .bas files have unique VB_Name attributes before packaging. Add a pre-delivery check: grep for duplicate VB_Name values. |
| **/automation XLSTART suppression** | Excel's /automation flag suppresses not just COM add-ins but also XLSTART workbooks (Personal.xlsb). The Quit() after /automation can persist disabled state to the Resiliency\DisabledItems registry key. | Snapshot the Resiliency\DisabledItems key before /automation launch. After quit, remove any entries that didn't exist before the session. |

### F.3 Design Decisions

| Decision | Rationale |
|----------|-----------|
| KernelConfig split into KernelConfig + KernelConfigLoader | KernelConfig.bas was at 59,861 bytes (5.7KB from 64KB VBA import limit). Split moves all loader functions and fallback helpers to KernelConfigLoader.bas. KernelConfig retains getters and array declarations. |
| Public arrays for config split (AP-13 exception) | KernelConfigLoader writes directly to KernelConfig.m_xxx Public arrays. Simpler than setter functions or ByRef parameter threading. Only KernelConfigLoader writes; all other modules use getters. |
| TransformOutputs handoff pattern | Application.Run cannot pass arrays. Module-level Public Variant holds the outputs array. Transforms access it directly. Caller copies in before and out after. Same pattern will apply to Phase 6 extensions. |
| Power Pivot AUTO detection | Check config setting -> check ThisWorkbook.Model availability -> proceed or skip with INFO log. No error, no MsgBox on skip. Clean degradation for Excel versions without Power Pivot. |
| Pivot tables from Detail range (not data model) | PivotCache created from Detail data range for maximum compatibility. Works regardless of Power Pivot availability. |
| Resiliency\DisabledItems snapshot/restore | Follows same defensive pattern as COM add-in HKCU stub cleanup (BUG-017 fix). Snapshot before, remove new entries after. Protects Personal.xlsb and any other XLSTART items. |
| SampleDomainTests.bas deleted (not fixed) | File was a stale corrupt copy of SampleDomainEngine.bas, not actual test code. CLAUDE.md referenced "DomainTests" module but it never contained test functions. Deleted rather than fixed. |

### F.4 Pre-Delivery Cleanup Process

**WARNING: This section has caused file loss (BUG-037). Read the PROTECTED FILES list carefully.**

#### F.4.1 Protected Files — NEVER Delete, Move, or Archive

These files MUST remain at the repo root after every phase. Claude Code must not move, archive, rename, or delete them under any circumstances:

| File | Why It's Protected |
|------|-------------------|
| `CLAUDE.md` | Project bible. Claude Code reads FIRST every session. |
| `SESSION_NOTES.md` | Rolling context log. Claude Code reads SECOND every session. Claude Online appends. |
| `config_sample/` (entire directory) | Seeder config for sample model. Required by Setup.bat. |
| `config_blank/` (entire directory) | Seeder config for blank model. Required by Setup.bat. |
| `data/anti_patterns.csv` | Institutional knowledge. Append-only. |
| `data/patterns.csv` | Institutional knowledge. Append-only. |
| `data/bug_log.csv` | Institutional knowledge. Append-only. |
| `engine/*.bas` | Source .bas files. Modified in-place, never moved. |
| `scripts/` (entire directory) | Bootstrap and utility scripts. |

**The rule is simple: if you didn't create it this session, don't delete it.**

#### F.4.2 Files That ARE Safe to Clean Up

| What | When | Notes |
|------|------|-------|
| `output/*.pdf` | Before packaging | Regenerated by ExportPDF |
| `workbook/*.xlsm` | Before packaging | Regenerated by Setup.bat |
| `scenarios/` (runtime snapshots) | Before packaging | Created by RunProjections |
| `wal/wal.log` | Before packaging | Created by KernelSnapshot |
| Phase build prompt (e.g., `PHASE5B_BUILD_PROMPT.md`) | After audit declares CLEAN/MINOR | Move to `docs/archive/` — it's consumed, no longer needed at root |
| `.zip` files in repo root | Before packaging | Old delivery artifacts |

#### F.4.3 Common Mistakes (Documented Incidents)

| Mistake | What Happened | Rule |
|---------|--------------|------|
| Moved `SESSION_NOTES.md` to `docs/archive/` | Phase 5B cleanup treated it as a "build input" to archive. Broke the context transfer protocol — Claude Code reads it second on every session start. | SESSION_NOTES.md is a **living document**, not a build artifact. It stays at root permanently. |
| Deleted `config/` directory | Phase 5B cleanup treated runtime config CSVs as regenerable artifacts. If `config/` is the active runtime directory that Bootstrap.ps1 populates, deleting it breaks the next Setup.bat run. | `config/` is populated by Setup.bat from `config_sample/` or `config_blank/`. If it exists, it was created by Setup.bat and is safe to delete before packaging (it will be regenerated). But `config_sample/` and `config_blank/` are NEVER deleted. |
| Modified `docs/archive/` files | Phase 5B appended build findings to `RDK_Developer_Flywheel_v1.0.md` in the archive. Archive files should be immutable snapshots. | Archive documents are frozen at the time they're archived. New findings go in the CURRENT version of the document, not the archived copy. |

#### F.4.4 The Pre-Delivery Checklist (Corrected)

1. **Runtime artifacts:** Delete `output/`, `workbook/`, `scenarios/`, `wal/` contents (all regenerated)
2. **Build prompt:** Move completed phase build prompt to `docs/archive/`
3. **Old ZIPs:** Delete any `.zip` files in the repo root
4. **Stale files:** Check for corrupt/duplicate `.bas` files (verify VB_Name uniqueness)
5. **DO NOT touch:** SESSION_NOTES.md, CLAUDE.md, config_sample/, config_blank/, data/, engine/, scripts/
6. **DO NOT modify:** Any file in `docs/archive/` — these are frozen historical snapshots

#### F.4.5 Claude Online Handoff Slim-Down Process

**Purpose:** Before zipping the repo for Claude Online audit, strip all regenerable artifacts and historical archives to minimize context size. This is different from Pre-Delivery Cleanup (F.4.4) — that prepares a build for Ethan. This prepares a repo for Claude Online to read, audit, and produce the next build prompt.

**Guiding principle:** Remove anything Claude Online cannot read (binaries) or does not need (regenerable artifacts, historical archives). Keep everything needed to understand the project, write specs, and produce a buildable repo.

**Step 1: Remove generated output (regenerated by pipeline)**
- `output/` — all contents (granular CSVs, PDFs). These are 88MB+ each.
- `scenarios/` — all contents (runtime scenario CSVs)
- `snapshots/` — all contents (golden snapshot CSVs)

**Step 2: Remove runtime artifacts (regenerated by Setup.bat / pipeline)**
- `config/` — all 18 CSVs (runtime copy; regenerated from config_sample/, config_insurance/, or config_blank/ by Setup.bat)
- `workbook/*.xlsm` — regenerated by Setup.bat + Bootstrap.ps1
- `workbook/~$*.xlsm` — Excel lock/temp files (should never be packaged)
- `wal/wal.log` — write-ahead log (created by KernelSnapshot at runtime)

**Step 3: Remove debugging artifacts**
- `diagnostic_dump_*.txt` — diagnostic dumps from prior debugging sessions
- Any `.zip` files in repo root — old delivery artifacts

**Step 4: Remove archived docs**
- `docs/archive/` — entire directory. These are frozen historical snapshots (old build prompts, old walkthroughs, old roadmaps, old flywheel versions). All have been superseded by current versions in `docs/` or at root. Claude Online does not need historical reference to produce the next phase.

**Step 5: Verify protected files survive**
- `CLAUDE.md` — present at root
- `SESSION_NOTES.md` — present at root
- `config_sample/` — all 18 CSVs present
- `config_insurance/` — all 18 CSVs present (if applicable)
- `config_blank/` — all 18 CSVs present
- `data/` — anti_patterns.csv, bug_log.csv, patterns.csv present
- `docs/` — main spec docs present (NOT archive)
- `engine/` — all .bas files present
- `scripts/` — all 4 scripts present

**Step 6: Create ZIP**
- Flat ZIP (no wrapper folder): `RDK_PhaseN_YYYYMMDD_HHMMSS.zip`
- Verify ZIP contents match expected file list

**What stays in the slim ZIP (typical ~2-3MB):**

| Directory | Contents | Why |
|-----------|----------|-----|
| root | CLAUDE.md, SESSION_NOTES.md, DELIVERY_SUMMARY.md | Project context |
| config_sample/ | 18 CSVs | Sample model seeder |
| config_insurance/ | 18 CSVs | Insurance model seeder |
| config_blank/ | 18 CSVs | Blank model seeder |
| data/ | 3 CSVs | Institutional knowledge |
| docs/ | ~9 spec/design docs | Living specifications |
| engine/ | ~30 .bas files | All VBA source code |
| scripts/ | 4 files | Setup and bootstrap scripts |

**What gets removed (typical ~268MB savings):**

| Directory | Why Removed |
|-----------|-------------|
| output/ | Regenerated by pipeline (88MB per run) |
| scenarios/ | Regenerated by RunProjections |
| snapshots/ | Regenerated by TakeSnapshot |
| config/ | Regenerated by Setup.bat from seeders |
| workbook/ | Regenerated by Setup.bat + Bootstrap.ps1 |
| wal/ | Regenerated by KernelSnapshot |
| docs/archive/ | Historical only; superseded by current docs |
| diagnostic_dump_*.txt | Debugging artifacts |

### F.5 Knowledge Artifact Counts (Post Phase 5B)

| Artifact | Count |
|----------|-------|
| Anti-patterns (AP-01 to AP-60) | 60 |
| Patterns (PT-001 to PT-026) | 26 |
| Bugs (BUG-001 to BUG-036) | 36 |
| VBA modules | 22 (20 kernel + 1 domain sample + 1 domain stub) |
| Config CSVs (per seeder) | 14 |
| Scripts | 4 (Setup.bat, Bootstrap.ps1, Toggle-ComAddins.ps1, Repair-ComAddins.ps1) |

---

## Appendix G: Phase 5C Build Findings

### G.1 Bugs Discovered and Fixed

| Bug | Description | Root Cause | Fix | Anti-Pattern |
|-----|-------------|------------|-----|--------------|
| BUG-037 | Protected files moved to docs/archive/ during cleanup | Claude Code treated SESSION_NOTES.md and other root files as build artifacts | Restored files; added Protected Files doctrine (F.4) | (file preservation) |
| BUG-038 | Named ranges created with wrong cell addresses | CreateNamedRanges used RowID as cell address without resolving to actual row | Resolved RowID to row number via formula_tab_config lookup, then built address | (address resolution) |
| BUG-039 | WriteQuarterlyHeaders hardcoded row 2 wiped by merge | FS_BASIS config at row 2 with ColSpan=6 merged A2:F2, overwriting quarterly headers | Changed to dynamically find first formula row and write headers one row above | (layout timing) |
| BUG-040 | Annual total column used SUM for ratio metrics | =SUM(Q1:Q4) for GP Margin produces 1.6, not 0.4 | Changed annual total to use ResolveFormulaPlaceholders with annual column | (formula semantics) |
| BUG-041 | AutoFit reset column Hidden state on formula tabs | ws.Columns.AutoFit after ws.Columns(col).Hidden = True unhid the column | Moved .Hidden = True to AFTER .Columns.AutoFit | AP-62 |
| BUG-042 | Missing number format for Pct and Currency DataTypes | COGSPct displayed as 0.6 instead of 60.0% | Added format application: Pct -> 0.0%, Currency -> #,##0.00 | (format application) |
| BUG-043 | CreateFormulaTabs MsgBox during RunProjections | Unconditional MsgBox calls paused pipeline execution | Added m_silent flag; pipeline sets True; UI wrapper sets False | AP-53 |
| BUG-044 | Toggle display mode button caption did not update | Dynamic button caption showed stale mode after toggle | Changed to static "Toggle Display Mode" caption | (usability) |

### G.2 Design Decisions

| Decision | Rationale |
|----------|-----------|
| Formula tabs from config (formula_tab_config.csv) | Declarative layout: each cell defined by RowID, Row, Col, CellType, Content, Format. No hardcoded layout code. |
| Named ranges via registry (named_range_registry.csv) | Declarative named range creation. RangeName, TabName, RowID resolve to cell addresses at runtime. |
| Quarterly aggregation with annual total | Q1-Q4 columns use quarter-scoped SUMIFS. Annual column uses ResolveFormulaPlaceholders for correct ratio recomputation. |

---

## Appendix H: Phase 6A Build Findings

### H.1 Bugs Discovered and Fixed

| Bug | Description | Root Cause | Fix | Anti-Pattern |
|-----|-------------|------------|-----|--------------|
| BUG-045 | Personal.xlsb not loading after Setup.bat | /automation Quit() writes shutdown state suppressing XLSTART; PID not captured so process kill impossible | Capture PID, kill by PID instead of Quit(), comprehensive post-cleanup | AP-54 |
| BUG-046 | PDF report cover page blank page 2 | GenerateCoverPage had no PageSetup (no PrintArea, FitToPages, or Orientation) | Added PageSetup: PrintArea=A1:B{lastRow}, FitToPagesWide=1, FitToPagesTall=1, Portrait | (print setup) |
| BUG-047 | Book counter increments and stale XLSTART lock file | Kill by PID leaves orphaned DocumentRecovery entries and ~$PERSONAL.XLSB lock file | Clear DocumentRecovery entries + remove stale lock files in Bootstrap.ps1 finally block | AP-54 |
| BUG-048 | PDF report tabs exported with default PageSetup | Ext_ReportGen.ExportReportPDF never called ConfigurePrintSettings; all tabs exported Portrait with no fit-to-page | Added ConfigurePrintSettings(silent:=True) call before export; added silent param to ConfigurePrintSettings | (integration gap) |

### H.2 Error Patterns

| Pattern | Description | Mitigation |
|---------|-------------|------------|
| **Lanczos approximation sign error** | LogGamma function used `(x - 0.5)` instead of `(x + 0.5)`. Single character corrupted all GammaCDF output. | Mathematical functions must be verified against Excel built-in equivalents (e.g., GAMMA.DIST) with exact expected values in the spec. |
| **Debug.Print array limitation** | `? FunctionReturningArray()` gives Error 13 Type Mismatch. Must index: `p = Func(): ? p(1)`. | Document in walkthrough; use multi-statement syntax for array-returning functions. |
| **Missing cross-module integration** | ReportGen assembled tabs for PDF but never called KernelPrint.ConfigurePrintSettings. Each module worked independently but the integration path was never wired. | When Module A uses Module B's output, the spec must explicitly state the call chain. Cross-module integration must be tested end-to-end, not just per-module. |
| **PageSetup.Zoom precedence** | Setting FitToPagesWide/Tall without first setting Zoom=False causes Excel to ignore fit-to-page settings entirely. | Always set `.Zoom = False` before any `.FitToPagesWide` or `.FitToPagesTall` assignment. |
| **Process kill side effects** | Killing Excel by PID avoids quit-cycle state writes but creates orphaned DocumentRecovery entries and stale lock files. Both prevent clean restart. | Post-kill cleanup must clear DocumentRecovery registry entries AND remove stale ~$* lock files from XLSTART. |
| **Cross-module constant references in tests** | Smoke tests referenced config marker constants (e.g., CFG_MARKER_GRANULARITY) that may not be Public or may not exist in KernelConstants. | All constants used in test code must be verified to exist as Public Const in KernelConstants.bas. Add a lint check for undefined cross-module constant references. |

### H.3 Design Decisions

| Decision | Rationale |
|----------|-----------|
| Extension infrastructure via KernelExtension.bas | Centralized lifecycle manager. Extensions register via extension_registry.csv. Pipeline calls RunExtensions("HookType") at each stage. Non-fatal: extension failure logged, pipeline continues. |
| CurveLib as Standalone extension | Math library with no pipeline side effects (MutatesOutputs=FALSE). Available to domain code via direct calls. Not triggered by pipeline hooks. |
| ReportGen as PostOutput extension | Generates PDF after all output tabs are written. Uses print_config.csv for tab selection and PageSetup, then adds cover page with TOC and Prove-It summary. |
| ConfigurePrintSettings silent parameter | Added to prevent MsgBox popups when called from pipeline/report paths. ExportPDF, PrintPreview, and ReportGen all call with silent=True and show their own result messages. |
| DomainModule dynamic dispatch | Config setting `DomainModule=SampleDomainEngine` drives Application.Run calls. No hardcoded domain references in kernel. Developers swap domain logic by changing one config value. |

### H.4 Process Improvement: Automated Validation (PT-028)

**Problem:** Phase 6A initial walkthrough had 11 gates, ~50 manual checkboxes, estimated 40-50 minutes. Many checks were deterministic value comparisons that could be automated but required human execution in the Immediate Window.

**Solution:** Tier 5 "Smoke" regression tests added to KernelTests.bas. These tests run automatically during `RunTests` and cover:
- Module existence (5 checks)
- Config section presence (3 checks)
- DomainModule setting (1 check)
- Extension registry state (3 checks)
- Detail fixture values: row count, specific values, total revenue, GPMargin consistency (5 checks)
- CurveLib math: CDF functions, dispatcher, interpolation, normalization (9 checks)
- Config lookup (2 checks)
- Required tabs (4 checks)

**Result:** Walkthrough reduced to 8 gates, ~20 manual checkboxes, estimated 20-25 minutes. Automated checks provide regression safety -- if a future phase breaks CurveLib math or fixture values, `RunTests` catches it immediately.

**Going forward:** Claude Online should tag each validation gate check as `[AUTO]` or `[MANUAL]` in the spec. Claude Code implements `[AUTO]` checks as Tier 5 smoke tests. The walkthrough document only includes `[MANUAL]` steps plus one step to run `RunTests`. This is now pattern PT-028.

### H.5 Phase 6A Validation Session Findings

Bugs discovered and fixed during Ethan's Excel validation of the Phase 6A build:

| Bug | Description | Root Cause | Fix | Anti-Pattern |
|-----|-------------|------------|-----|--------------|
| BUG-049 | Compile error: CFG_MARKER_GRANULARITY and TAB_ERRORLOG undefined | Smoke tests used incorrect constant names (CFG_MARKER_GRANULARITY instead of CFG_MARKER_GRANULARITY_CONFIG, TAB_ERRORLOG instead of TAB_ERROR_LOG) | Fixed to use correct constant names from KernelConstants.bas | (cross-module constant reference) |
| BUG-050 | Setup.bat fails "Cannot run macro" when Excel already running | GetActiveObject returns first Excel in ROT, not the /automation instance. Finally block killed wrong PID, orphaning the COM-connected instance | Pre-check exits if any Excel running; finally block kills ALL Excel processes | AP-54 |
| BUG-051 | SMK-030/031/032 fail -- extension registry not loaded | RunSmokeTests called extension query functions but LoadExtensionRegistry was never called (not included in LoadAllConfig) | Added LoadExtensionRegistry call at start of RunSmokeTests | (initialization gap) |
| BUG-052 | RunLint overwrites TestResults headers | WriteLintResults inserted rows at top of sheet, disrupting merged title cell | Changed to insert at row 3 (below header), pushing test results down (newest-on-top ordering) | (layout disruption) |
| BUG-053 | Walkthrough Step 6 tells user to run Setup option [1] after config edit | Options 1/2 delete config/ and re-copy from config_sample/, silently discarding user edits | Changed walkthrough to use option [3] (rebuild workbook only, keep config) | (walkthrough error) |
| BUG-054 | MsgBox "Created 1 formula-driven tab" fires during Setup | KernelBootstrap calls CreateFormulaTabs directly; m_silent defaults to False; only pipeline path set it True | Added Optional silent parameter; bootstrap passes True; audited all bootstrap sub-calls | AP-53 |

### H.6 Validation Session Error Patterns

| Pattern | Description | Mitigation |
|---------|-------------|------------|
| **Initialization gaps in test code** | Test suites call query functions (IsExtensionActive, GetActiveExtensionCount) without first loading the backing data (LoadExtensionRegistry). LoadAllConfig does not cover all registries. | Every test suite must explicitly load all registries it queries. Document which Load* calls are separate from LoadAllConfig. |
| **Walkthrough/setup option mismatch** | Walkthrough told user to use option [1] (fresh start) for config change testing, which resets config. Users naturally edit config/ (runtime) not config_sample/ (source). | Post-initial-setup walkthrough steps should always use option [3] (keep config). Only Step 1 should use option [1]. |
| **Silent MsgBox in non-interactive paths** | Functions with m_silent flags work when called through the pipeline wrapper (which sets silent=True), but fail when called from other non-interactive paths (bootstrap) that don't set the flag. | Use Optional silent parameter on the function itself, not just module-level state. All callers explicitly declare intent. |
| **TestResults layout fragility** | Any code that writes to TestResults must coordinate with existing content. Inserting at row 1 disrupts merged headers; appending at bottom buries newest results. | Newest sections insert at row 3 (below fixed header rows 1-2), pushing older content down. Never insert at row 1. |
| **PageSetup performance** | Each PageSetup property triggers a printer driver round-trip. Configuring 7 tabs x 10+ properties = 70+ round-trips, dominating report generation time. | Wrap all PageSetup changes in Application.PrintCommunication = False/True (PT-029). Disable ScreenUpdating for multi-sheet operations (PT-030). |

### H.7 Validation Session Design Decisions

| Decision | Rationale |
|----------|-----------|
| Detail tab Portrait orientation + CenterHorizontally | Detail Output is a columnar data table that reads better in Portrait. Added CenterHorizontally as column 14 in print_config.csv for all tabs. |
| Health check silent on workbook open | RunHealthCheckOnOpen now passes silent=True. Health check still runs and logs to ErrorLog, but no MsgBox interrupts workbook opening. User can run full health check from Dashboard. |
| TestResults newest-on-top ordering | WriteLintResults inserts at row 3 (below title/run-info), pushing test results down. Most recent section is always visible at top without scrolling. |
| PrintCommunication batching (PT-029) | Application.PrintCommunication = False before PageSetup loop, True after. Uses On Error Resume Next for Excel 2010+ compatibility. Restores True in error handler. |
| DomainOutputs direct assignment | Replaced O(rows x cols) element-by-element copy-back loop with single `outputs = DomainOutputs` assignment. VBA Variant assignment copies the array. |

### H.8 New Patterns Established

| Pattern | Name | Description |
|---------|------|-------------|
| PT-029 | PrintCommunication batching | Wrap PageSetup property changes in PrintCommunication = False/True to batch printer driver communication. |
| PT-030 | ScreenUpdating guard for multi-sheet ops | Any operation touching multiple sheets must disable ScreenUpdating at entry and restore at exit (including error handler). |

### H.9 Knowledge Artifact Counts (Post Phase 6A + Validation)

| Artifact | Count |
|----------|-------|
| Anti-patterns (AP-01 to AP-62) | 62 |
| Patterns (PT-001 to PT-030) | 30 |
| Bugs (BUG-001 to BUG-054) | 54 |
| VBA modules | 27 (23 kernel + 2 extension + 1 domain sample + 1 domain stub) |
| Config CSVs (per seeder) | 18 |
| Scripts | 4 (Setup.bat, Bootstrap.ps1, Toggle-ComAddins.ps1, Repair-ComAddins.ps1) |

---

## Appendix I: Post-Phase 11B Hotfixes (BUG-082)

### I.1 Problem

After Phase 11B delivery (v1.2.0), users opening RDK_Model.xlsm via double-click encounter Excel Safe Mode prompt. Root cause: Phase 11B added 3 new modules (+104KB VBA), bringing the total from 26 to 30 modules (830KB). Combined with PERSONAL.XLSB (a second VBA project loaded from XLSTART), the total VBA load at startup triggers an access violation on Office builds below 16.0.19822.20114. Setup.bat printed instructions to use OpenRDK.cmd, but users naturally double-click the .xlsm file. Additionally, OpenRDK.cmd lacked interrupt protection, leaving PERSONAL.XLSB stuck renamed if the script was interrupted.

### I.2 Bugs Found and Fixed

| Bug | Description | Root Cause | Fix | Anti-Pattern |
|-----|-------------|------------|-----|-------------|
| BUG-082 | Safe mode prompt on double-click after Setup; stuck PERSONAL.XLSB rename | (1) UX gap: Setup.bat tells users to use OpenRDK.cmd but users double-click. (2) OpenRDK.cmd lacked interrupt recovery for PERSONAL.XLSB rename. (3) All 30 modules imported unconditionally regardless of config. | (1) Setup.bat prompts "Open workbook now?" and calls OpenRDK.cmd directly. (2) OpenRDK.cmd restores stuck .rdk_bypass on startup, verifies rename success. (3) Bootstrap.ps1 reads DomainModule from granularity_config.csv and skips unused domain modules (PT-031). | AP-54 |
| BUG-083 | Excel crash persists after conditional import -- corrupt COM stubs and startup overload | Conditional import reduced to 27 modules but crash persisted. 4 corrupt HKCU COM add-in stubs (no Manifest/FriendlyName) overrode HKLM registrations. PDFMaker stub escalated LoadBehavior from 2 to 3. Combined with S&P Cap IQ (2 COM + 3 XLA) + PERSONAL.XLSB + RDK = 8 items competing for startup. | (1) Moved PERSONAL.XLSB to PERSONAL_BACKUP. (2) CleanComStubs.ps1 removes incomplete HKCU stubs. (3) OpenRDK.cmd calls CleanComStubs.ps1 before open. | AP-54 |

### I.3 Error Patterns

| Pattern | Observation | Mitigation |
|---------|-------------|------------|
| **Module count threshold** | 30-module VBA project + PERSONAL.XLSB crosses stability threshold on some Office builds. 26 modules worked; 30 did not. | Conditional import reduces to 27 (sample) or 29 (insurance). Future: add-in split (Phase 12i on roadmap). |
| **Invisible workaround instructions** | Setup.bat printed "Use OpenRDK.cmd" in scrollable terminal output. Users missed it and followed natural instinct (double-click). | Auto-open prompt at end of Setup.bat. Workaround should be the default path, not an instruction. |
| **Fragile file rename without recovery** | OpenRDK.cmd renamed PERSONAL.XLSB but had no recovery if interrupted. File stayed renamed across sessions. | Restore-on-startup check at beginning of OpenRDK.cmd. Verify rename succeeded before proceeding. |
| **Corrupt HKCU COM stubs compound startup load** | Bootstrap.ps1 /automation creates incomplete HKCU stubs (LoadBehavior only). These override complete HKLM registrations. PDFMaker stub escalated from demand-load (2) to startup-load (3), adding ~50MB of Adobe COM to every Excel launch. | CleanComStubs.ps1 runs before workbook open. Only removes stubs with no Manifest AND no FriendlyName (preserves intentional user overrides). |
| **Multiple add-in types stack** | S&P Capital IQ installs 2 COM add-ins (HKLM LoadBehavior=3) + 3 Excel add-ins (OPEN registry keys) + UDF functions. Combined with PERSONAL.XLSB + PDFMaker + RDK = 8+ simultaneous loads. | Move PERSONAL.XLSB out of XLSTART. Clean corrupt stubs. Conditional import. Future: /automation flag with selective add-in restore. |

### I.4 Design Decisions

| Decision | Rationale |
|----------|-----------|
| Config-driven module filtering (PT-031) | Bootstrap.ps1 reads DomainModule from granularity_config.csv (already authoritative for domain dispatch in KernelEngine) and skips modules not needed for the active domain. Safe fallback: unknown DomainModule imports all. |
| Setup.bat auto-open prompt | Default Y opens via OpenRDK.cmd. User never needs to know about PERSONAL.XLSB conflict. Non-interactive users can type N. |
| Kernel add-in split deferred to Phase 12i | Moving 25 modules to .xlam is the permanent fix but requires architecture review (risk: 3 VBA projects may be worse than 2). Added to roadmap for proper spec/review cycle. |

### I.5 New Patterns Established

| Pattern | Name | Description |
|---------|------|-------------|
| PT-031 | Conditional module import from config | Bootstrap reads DomainModule from granularity_config.csv and imports only needed VBA modules. Skips unused domain/companion modules to reduce project size. Unknown DomainModule falls back to all. |

### I.6 Knowledge Artifact Counts (Post Phase 11B Hotfixes)

| Artifact | Count |
|----------|-------|
| Anti-patterns (AP-01 to AP-63) | 63 |
| Patterns (PT-001 to PT-031) | 31 |
| Bugs (BUG-001 to BUG-134) | 134 |
| Kernel VBA modules | 27 |
| Domain + companion modules | 8 (3 domain, 5 companion) |
| Extension modules | 2 |
| Config CSVs (FM) | 31 |
| Scripts | 7 |

---

## F.5 Pre-Delivery Cleanup Checklist

**CRITICAL: Run this checklist before every ZIP delivery.** Failure to clean causes multi-GB ZIPs with runtime artifacts.

### Directories to Clear (contents only, keep the empty folder)
| Directory | What it contains | Why clear |
|-----------|-----------------|-----------|
| `output/` | Granular CSVs, PDFs, reports | Runtime artifacts. 1.6GB+ after many runs. |
| `scenarios/` | CSV scenario exports | Runtime artifacts from RunModel. |
| `wal/` | Write-ahead log files | Runtime debug log. |
| `config/` | Runtime copy of config CSVs | Seeded from config_insurance/ by Setup.bat. Not version-controlled. |
| `snapshots/` | Legacy snapshots (deprecated) | Replaced by workspaces/. Delete all except GOLDEN_* if present. |

### Files to Remove
| File | Location | Why |
|------|----------|-----|
| `Archive_*.xlsm` | workbook/ | Old workbook backups. Only keep the current RDK_Model.xlsm. |
| `diagnostic_dump_*.txt` | root | Generated on demand, not version-controlled. |
| `*.xlsx` reference files | root | Design references (e.g., RDK_Dashboard.xlsx). Move to docs/archive/ or delete. |

### Files to Keep
| File | Location | Why |
|------|----------|-----|
| `CLAUDE.md` | root | Project bible. Read first every session. |
| `SESSION_NOTES.md` | root | Context transfer log. Append only. |
| `workspaces/GOLDEN_*/` | workspaces/ | Golden regression baselines. Keep for testing. |
| `config_insurance/` | root | FM model configuration. Source of truth. |
| `config_sample/` | root | Sample model configuration. |
| `config_blank/` | root | Blank bootable skeleton. |
| `engine/*.bas` | engine/ | All VBA modules. Source of truth. |
| `data/*.csv` | data/ | Institutional knowledge. Append only. |
| `docs/*.md` | docs/ | Living documents + archive. |
| `scripts/*` | scripts/ | Bootstrap and utility scripts. |

### Build Prompts
Move completed build prompts to `docs/archive/` after the build is done. They are historical reference, not active documents.

### Verification
After cleanup, the repo should be under 15MB (excluding workbook). Run:
```
du -sh */
```
Expected: engine ~1.2MB, config_insurance ~400KB, docs ~350KB, everything else minimal.
