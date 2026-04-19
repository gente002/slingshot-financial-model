# RDK Codebase Adversarial Review — Technical Debt Assessment

**Instructions:** You are an independent code reviewer performing a technical debt audit of the RDK (Rapid Development Kit) — a zero-dependency Excel/VBA/CSV/PowerShell insurance carrier financial model framework. Your job is to find problems, not to praise. Be direct, specific, and cite file names + line numbers where possible.

---

## REVIEW FORMAT

### 1. Executive Summary (FIRST PARAGRAPH — MANDATORY)

Start with a single paragraph that answers: **"Is this codebase getting better or worse, and how much work is needed to stabilize it?"** Include the overall weighted score (X.X / 10.0) and a one-sentence verdict. If this is a re-review, open with: "Compared to the prior review (vX.X, score Y.Y), this version [improved/regressed] by Z.Z points, primarily due to [reason]."

### 2. Scorecard Table (IMMEDIATELY AFTER EXECUTIVE SUMMARY)

```
| # | Category                        | Score | Prior | Δ   | Key Finding |
|---|----------------------------------|-------|-------|-----|-------------|
| 1 | DRY / Code Duplication           | X/10  | -     | -   | [one line]  |
| 2 | Single Responsibility            | X/10  | -     | -   | [one line]  |
| 3 | Naming & Readability             | X/10  | -     | -   | [one line]  |
| 4 | Error Handling & Resilience      | X/10  | -     | -   | [one line]  |
| 5 | Configuration vs Hardcoding      | X/10  | -     | -   | [one line]  |
| 6 | Module Architecture & Coupling   | X/10  | -     | -   | [one line]  |
| 7 | Data Integrity & Validation      | X/10  | -     | -   | [one line]  |
| 8 | Performance & Efficiency         | X/10  | -     | -   | [one line]  |
| 9 | Testability & Test Coverage      | X/10  | -     | -   | [one line]  |
|10 | Documentation & Maintainability  | X/10  | -     | -   | [one line]  |
|   | **WEIGHTED OVERALL**             |**X.X**| -     | -   |             |
```

Weights: Categories 1, 2, 4, 6 = 1.5x. Categories 3, 5, 7 = 1.0x. Categories 8, 9, 10 = 0.75x.

For re-reviews: fill in the "Prior" and "Δ" columns from the previous assessment. Flag any category that dropped ≥1 point in bold red.

### 3. Detailed Findings (PER CATEGORY)

For each of the 10 categories, provide:

**Score: X/10**

**What went well:** 1-2 specific strengths with file/line references.

**What went wrong:** All violations found, ordered by severity (P0 critical → P3 cosmetic). Include:
- File name and line number (or function name)
- What the violation is
- What the correct implementation looks like
- Estimated effort to fix (trivial / small / medium / large)

**Grading rubric:**
- **9-10:** Exemplary. No violations found. Would pass a senior engineer's review.
- **7-8:** Good. Minor violations only. Patterns are sound, a few exceptions.
- **5-6:** Acceptable. Several violations. Core patterns work but inconsistently applied.
- **3-4:** Concerning. Systemic violations. Patterns are stated but not enforced.
- **1-2:** Critical. Fundamental problems that undermine reliability.

### 4. Top 5 Technical Debt Items (PRIORITIZED)

Rank the 5 highest-impact technical debt items across all categories. For each:
- What it is (specific, not vague)
- Why it matters (what breaks or degrades if not fixed)
- Estimated effort (hours/days)
- Recommended priority (P0-P3)

### 5. DRY Deep Dive (SPECIAL SECTION)

DRY violations are the most common debt in VBA codebases. Specifically look for:
- Functions with >80% similar logic across modules
- Copy-paste patterns (same formula structure repeated with minor variations)
- Magic numbers / magic strings that should be constants or config
- Repeated error-handling boilerplate that should be a helper
- Column/row references that are hardcoded instead of computed

List every DRY violation found with the pattern name, affected files, and refactoring approach.

### 6. Architecture Risks

Identify any structural risks:
- Modules approaching size limits (64KB VBA limit)
- Circular dependencies between modules
- God modules (doing too many unrelated things)
- Missing abstraction layers
- Config tables that have outgrown their schema

---

## CODEBASE CONTEXT

**Technology:** Excel VBA (32 modules, 920KB total), CSV config (18 tables, 2,766 rows), PowerShell bootstrap (905 lines), Markdown docs (14 files).

**Architecture:**
- Kernel layer (23 modules): domain-agnostic framework — config loading, formula engine, snapshot, output, testing
- Extension layer (2 modules): CurveLib (actuarial curves), ReportGen (PDF generation)
- Domain layer (3 modules): DomainEngine stub, SampleDomainEngine, InsuranceDomainEngine
- Companion layer (4 modules): Ins_GranularCSV, Ins_QuarterlyAgg, Ins_Triangles, Ins_Tests

**Key patterns documented in the codebase:**
- 31 patterns (data/patterns.csv) — established best practices
- 63 anti-patterns (data/anti_patterns.csv) — known pitfalls
- 117 bugs logged (data/bug_log.csv) — full history

**Known technical debt (addressed in this version — verify fixes):**
- TD-01: KernelFormula.bas was 62.6KB — should now be split into KernelFormula + KernelFormulaWriter
- TD-02: Hardcoded sheet name strings — should now use TAB_* constants
- TD-03: devTabs array duplicated in 2 modules — should now be a shared function
- TD-04: Silent error exits — should now have LogError before early exits
- TD-05: BOM in formula_tab_config.csv header — should be removed
- TD-06: Config-missing fallback behavior — should be documented and consistent

**Known technical debt (NOT yet addressed — verify still present):**
- TD-07: InsuranceDomainEngine.bas still large (~54KB) — actuarial math extraction deferred
- TD-08: Ins_Tests not integrated with KernelTests harness — deferred
- 4 modules in WARN zone (>50KB): verify which ones remain after TD-01 split

**Version:** Post-TD (technical debt remediation). Phases 1-12B complete + TD fixes.

**Prior review scores (v1.3.0, ChatGPT + Claude arbiter):**

```
| # | Category                        | v1.3.0 Score | Arbiter Adjusted |
|---|----------------------------------|--------------|------------------|
| 1 | DRY / Code Duplication           | 6/10         | 6/10             |
| 2 | Single Responsibility            | 7/10         | 6/10             |
| 3 | Naming & Readability             | 8/10         | 8/10             |
| 4 | Error Handling & Resilience      | 7/10         | 6/10             |
| 5 | Configuration vs Hardcoding      | 8/10         | 7/10             |
| 6 | Module Architecture & Coupling   | 7/10         | 6/10             |
| 7 | Data Integrity & Validation      | 7/10         | 7/10             |
| 8 | Performance & Efficiency         | 8/10         | 8/10             |
| 9 | Testability & Test Coverage      | 6/10         | 6/10             |
|10 | Documentation & Maintainability  | 8/10         | 8/10             |
|   | WEIGHTED OVERALL                 | 7.9          | 7.2              |
```

Use the "Arbiter Adjusted" column as the "Prior" scores in your scorecard.

**Corrections from prior review (known false findings to avoid repeating):**
1. CSV column naming IS consistent (PascalCase throughout) — do not flag as inconsistent.
2. KernelTests IS an automated regression harness (RunSmokeTests, RunIntegrationTests, RunRegressionTests, golden baseline comparison) — do not claim tests are manual.
3. Get*/Calc*/Run*/Create* verb prefixes are intentionally differentiated (accessor vs computation vs side-effect vs construction) — do not flag as naming inconsistency.
4. Domain engine actuarial calculations (CDF evaluation, loss development) are NOT duplicated formula logic — they are domain-specific computations distinct from config-driven formula tab writing.

---

## FILES TO REVIEW

Review ALL files. Priority order:

1. `engine/KernelFormula.bas` (62.6KB — largest, most complex, near size limit)
2. `engine/InsuranceDomainEngine.bas` (54.8KB — domain logic, recently had EP/WP basis rewrite)
3. `engine/KernelBootstrap.bas` (49.1KB — orchestration, many hardcoded patterns)
4. `engine/KernelSnapshot.bas` (56.6KB — persistence, error handling intensive)
5. `engine/KernelTabs.bas` (53.3KB — tab management, formatting)
6. `engine/KernelConfig.bas` (39.0KB — config loading, accessor functions)
7. `engine/Ext_CurveLib.bas` (35.2KB — math library, UDFs)
8. `config_insurance/formula_tab_config.csv` (2,487 rows — formula definitions)
9. `config_insurance/tab_registry.csv` (31 rows — tab definitions)
10. All remaining engine/*.bas files
11. `scripts/Bootstrap.ps1` and `scripts/Setup.bat`
12. `data/anti_patterns.csv` and `data/patterns.csv` (check for violations of stated patterns)

---

## ADVERSARIAL INSTRUCTIONS

1. **Do NOT grade on a curve.** A 5/10 means "acceptable but flawed." A 7/10 means "good." Reserve 9-10 for genuinely excellent work.
2. **Find at least 3 violations per category.** If you can't find 3, look harder. Every codebase has debt.
3. **Check stated patterns against actual code.** The codebase documents its own anti-patterns — verify they're actually avoided.
4. **Look for the bugs they haven't found.** 117 bugs were caught and fixed. What's still lurking?
5. **Challenge architecture decisions.** Is the kernel/domain separation clean? Are the right things configurable?
6. **Test the config-driven claims.** How much is truly config-driven vs how much has hardcoded fallbacks that undermine the config?
7. **Pressure-test error handling.** What happens when config is missing? When a tab doesn't exist? When a formula references a nonexistent RowID?
8. **Cross-reference anti-patterns vs code.** Pick 10 random anti-patterns from the CSV and verify the code doesn't violate them.
9. **Grade documentation honestly.** Inline comments, module headers, SESSION_NOTES — is someone new actually able to understand this codebase?
10. **Don't anchor on prior scores.** If this is a re-review, grade independently first, then compare.

---

## DELIVERABLE

One document with sections 1-6 as described above. No preamble, no disclaimers about "I'm an AI." Start with the executive summary paragraph.
