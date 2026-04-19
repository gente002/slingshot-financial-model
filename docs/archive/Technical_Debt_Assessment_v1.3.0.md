# Technical Debt Review — Arbiter Assessment & Action Spec

**Date:** 2026-04-02
**Reviewer:** ChatGPT (adversarial)
**Arbiter:** Claude Online (this document)
**Codebase:** RDK v1.3.0

---

## Agreement / Disagreement Analysis

### AGREE — Valid findings, action needed

| # | Finding | My Assessment |
|---|---|---|
| 1 | **KernelFormula.bas at 62.6KB** (P0) | Fully agree. This is a ticking time bomb. One more formula feature = hard crash. Must split before any further formula work. |
| 2 | **InsuranceDomainEngine.bas too large** (P0) | Agree. 54.8KB and mixing actuarial math, tab orchestration, and aggregation. The BUG-117 EP-basis fix shows how fragile this is. |
| 3 | **devTabs array duplicated** in KernelBootstrap + KernelFormHelpers | Confirmed — exact same Array() literal in two places. Textbook DRY violation. Trivial fix: move to KernelConstants or a shared function. |
| 4 | **Hardcoded sheet name strings** despite TAB_* constants existing | Confirmed — 10+ instances of literal "Detail", "Dashboard", "Summary" instead of using TAB_DETAIL, TAB_DASHBOARD, TAB_SUMMARY constants. |
| 5 | **Silent error exits** in KernelFormula | Confirmed — 21 Exit Function/Sub vs 11 LogError calls. Some error paths exit without logging. |
| 6 | **Config missing → fallback vs fail inconsistency** | Confirmed — KernelSnapshot uses Dir() checks with fallback paths, while formula engine fails fast on missing RowIDs. Inconsistent philosophy. |
| 7 | **Public globals in InsuranceDomainEngine** | Confirmed — 15+ Public arrays (m_wpMon, m_epMon, m_ultMon, etc.) accessible to any module. Creates hidden coupling. |
| 8 | **BOM in formula_tab_config.csv header** | ChatGPT missed this but I found it — UTF-8 BOM embedded inside the first CSV field. Latent bug that could break header parsing. |
| 9 | **Ins_Tests coverage is thin** | Confirmed — 6 test subs, no assertions framework. KernelTests has a proper harness (SmokeCheck*, WriteTestRow) but Ins_Tests is ad-hoc. |

### DISAGREE — Overstated or inaccurate findings

| # | Finding | My Assessment |
|---|---|---|
| 1 | **"Formula duplication across modules"** (P0) | Overstated. The DomainEngine does actuarial calculations (CDF evaluation, loss development) which are domain-specific, not duplicated formula logic. KernelFormula handles config-driven formula tabs. These are different responsibilities, not DRY violations. |
| 2 | **"Tab write loop duplication"** across KernelTabs/Snapshot/ReportGen | Partially overstated. These modules write to tabs for different purposes (formatting vs snapshot export vs PDF). The loop structure is similar but the content and context differ. A shared WriteTable() helper would actually increase coupling. Score this as P3, not P1. |
| 3 | **"CSV column naming inconsistent (RowID vs Row_Id vs row_id)"** | Factually incorrect. I checked every CSV header — they all use consistent PascalCase: "RowID", "TabName", "CellType". ChatGPT hallucinated this finding. |
| 4 | **"Mixed verb naming (Get vs Calc vs Run)"** | Overstated. Get* = accessor, Calc* = pure computation, Run* = side-effecting operation, Create* = construction. This is a reasonable convention, not inconsistency. |
| 5 | **"No automated regression harness"** | Incorrect. KernelTests.RunTests() IS the automated harness — it has RunSmokeTests, RunIntegrationTests, RunRegressionTests, golden baseline comparison. It runs via a Dashboard button. Not CI/CD, but automated within Excel. |
| 6 | **Overall score 7.9** | Slightly generous given the P0 module size risk. I'd score it 7.4 — the module size issue alone should drag the Architecture category to 6, not 7. |

### NEEDS ARBITER DECISION

| # | Item | Options |
|---|---|---|
| 1 | **KernelFormula split approach** | (A) Split into KernelFormula + KernelFormulaWriter (extract the tab-writing logic). (B) Full 12i add-in split (move all kernel modules to .xlam). (C) Both — do A now as a quick fix, B later as architecture improvement. **Recommendation: C.** |
| 2 | **InsuranceDomainEngine split approach** | (A) Extract actuarial math into Ins_Actuarial.bas. (B) Extract tab orchestration into Ins_TabWriter.bas. (C) Both. **Recommendation: A — the actuarial math (CDF evaluation, loss development, reserve computation) is the cleanest extraction boundary.** |
| 3 | **Error handling philosophy** | (A) Standardize on fail-fast everywhere (config missing = error). (B) Keep fallback patterns where they exist. (C) Fail-fast for domain config, fallback for infrastructure (snapshots, optional features). **Recommendation: C.** |
| 4 | **Public globals in DomainEngine** | (A) Encapsulate in a UDT (User Defined Type) and pass as argument. (B) Keep as-is — VBA doesn't support classes well enough to justify refactoring. (C) Move to a dedicated state module (Ins_State.bas). **Recommendation: B for now — the globals are scoped to the domain engine's lifecycle and aren't accessed cross-module in practice. Refactor only if it causes a real bug.** |

---

## Action Spec — What to Fix

### Priority 1 (Do Now — Before Any More Feature Work)

**TD-01: Split KernelFormula.bas**
- Extract `CreateFormulaTabs`, `RefreshFormulaTabs`, `WriteQuarterlyHeaders`, and all tab-writing logic into `KernelFormulaWriter.bas`
- KernelFormula.bas retains: `ResolveFormulaPlaceholders`, `ResolveRowID`, `ClearRowIDCache`, `CreateNamedRanges`
- Target: KernelFormula < 35KB, KernelFormulaWriter < 35KB
- Effort: 4 hours

**TD-02: Fix hardcoded sheet name strings**
- Replace all literal "Detail", "Dashboard", "Summary" etc. with TAB_* constants from KernelConstants.bas
- 10+ instances across KernelBootstrap, KernelFormHelpers, Ins_Tests, KernelProveIt
- Effort: 1 hour

**TD-03: Deduplicate devTabs array**
- Move to a Public Function in KernelConstants: `GetDevModeTabs() As Variant`
- Replace both Array() literals in KernelBootstrap and KernelFormHelpers
- Effort: 30 minutes

### Priority 2 (Do Soon — Next Sprint)

**TD-04: Silent error exit audit**
- Audit all Exit Function/Sub in KernelFormula.bas
- Add LogError before every early exit that indicates a failure (not a normal guard clause)
- Effort: 2 hours

**TD-05: Fix BOM in formula_tab_config.csv**
- Remove UTF-8 BOM from the CSV header
- Add BOM-stripping logic to KernelConfigLoader as a safety net
- Effort: 30 minutes

**TD-06: Standardize config-missing behavior**
- Document the rule: domain config missing = fail fast, infrastructure config missing = fallback with warning
- Add LogEvent for all fallback paths so they're visible in ErrorLog
- Effort: 2 hours

### Priority 3 (Do When Convenient)

**TD-07: InsuranceDomainEngine split**
- Extract actuarial math (CDF evaluation, loss development arrays, reserve computation) into Ins_Actuarial.bas
- DomainEngine retains: Initialize, WriteOutputs, tab orchestration
- Target: IDE < 40KB, Ins_Actuarial < 25KB
- Effort: 1 day

**TD-08: Ins_Tests integration with KernelTests harness**
- Refactor Ins_Tests to use KernelTests.WriteTestRow and SmokeCheck* helpers
- Register domain tests in the main test runner
- Effort: 4 hours

---

## Updated Scores (My Assessment)

| # | Category | ChatGPT | My Score | Delta | Rationale |
|---|---|---|---|---|---|
| 1 | DRY | 6 | 6 | 0 | Agree — devTabs dup, magic strings confirmed |
| 2 | Single Responsibility | 7 | 6 | -1 | KernelFormula at 62.6KB is a P0, not a 7 |
| 3 | Naming & Readability | 8 | 8 | 0 | Agree — CSV naming claim was wrong but overall good |
| 4 | Error Handling | 7 | 6 | -1 | Silent exits + inconsistent fallback worse than stated |
| 5 | Config vs Hardcoding | 8 | 7 | -1 | Magic strings undermine the config-driven architecture |
| 6 | Architecture & Coupling | 7 | 6 | -1 | Module sizes are a structural risk, not just cosmetic |
| 7 | Data Integrity | 7 | 7 | 0 | Agree |
| 8 | Performance | 8 | 8 | 0 | Agree |
| 9 | Testability | 6 | 6 | 0 | Agree |
| 10 | Documentation | 8 | 8 | 0 | Agree |
| | **WEIGHTED OVERALL** | **7.9** | **7.2** | **-0.7** | Module size risk is more severe than ChatGPT scored |
