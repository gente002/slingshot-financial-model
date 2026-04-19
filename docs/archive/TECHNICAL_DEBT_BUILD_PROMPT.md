# RDK Technical Debt Remediation — Claude Code Build Prompt

**Date:** April 2026
**Prior Phase:** Phase 12B COMPLETE at v1.3.0
**Goal:** Address 6 technical debt items (TD-01 through TD-06) identified in adversarial code review.

## BEFORE WRITING ANY CODE

1. Read `CLAUDE.md` — project bible.
2. Read `SESSION_NOTES.md` — **APPEND ONLY — do not truncate.**
3. Read `docs/Technical_Debt_Assessment_v1.3.0.md` — the full assessment with action items.

---

## TD-01: Split KernelFormula.bas (P0 — CRITICAL)

KernelFormula.bas is at 62.6KB — 1.4KB from the 64KB VBA hard limit (AD-09). Split it.

**Extract into new module `KernelFormulaWriter.bas`:**
- `CreateFormulaTabs` (and all Private helpers it calls for tab writing)
- `RefreshFormulaTabs` and `RefreshFormulaTabsUI`
- `WriteQuarterlyHeaders`
- `WriteFormula` helper
- `ColLetter` helper (if not shared elsewhere)
- All Private subs/functions used exclusively by the above

**KernelFormula.bas retains:**
- `ResolveFormulaPlaceholders`
- `ResolveRowID` and `ClearRowIDCache`
- `CreateNamedRanges`
- Any shared helpers used by both modules (keep in KernelFormula, have Writer call them)

**Target sizes:** KernelFormula.bas < 35KB, KernelFormulaWriter.bas < 35KB.

**Update callers:** Any module that calls `KernelFormula.CreateFormulaTabs` must now call `KernelFormulaWriter.CreateFormulaTabs`. Check:
- KernelEngine.bas (main pipeline)
- KernelBootstrap.bas
- Any module referencing KernelFormula public functions that moved

**Add KernelFormulaWriter to CLAUDE.md** module count (Kernel modules: 23 → 24).

---

## TD-02: Fix hardcoded sheet name strings

Replace ALL literal tab name strings with TAB_* constants from KernelConstants.bas.

Known instances (verify each, there may be more):
- `Ins_Tests.bas:461` — `"Detail"` → `TAB_DETAIL`
- `KernelBootstrap.bas:94` — `"Dashboard"` → `TAB_DASHBOARD`
- `KernelBootstrap.bas:1040` — `"Detail"` → `TAB_DETAIL`
- `KernelBootstrap.bas:1238-1239` — devTabs Array with literal strings (fixed by TD-03)
- `KernelFormHelpers.bas:738-739` — devTabs Array with literal strings (fixed by TD-03)
- `KernelProveIt.bas:88` — `"Detail"` → `TAB_DETAIL`

**Search pattern:** `grep -n '"Detail"\|"Summary"\|"Dashboard"\|"ErrorLog"\|"Config"' engine/*.bas`

Exclude: Const declarations, string content for labels/descriptions/log messages, column header values.

---

## TD-03: Deduplicate devTabs array

The dev mode tab list is duplicated as an identical Array() literal in:
- `KernelBootstrap.bas` (~line 1238)
- `KernelFormHelpers.bas` (~line 738)

**Fix:** Add a Public Function to KernelConstants.bas:

```vba
Public Function GetDevModeTabs() As Variant
    GetDevModeTabs = Array(TAB_DETAIL, TAB_QUARTERLY_SUMMARY, TAB_ERROR_LOG, _
                          TAB_TEST_RESULTS, TAB_CUMULATIVE_VIEW, TAB_ANALYSIS, _
                          TAB_SUMMARY, TAB_PROVE_IT, TAB_EXHIBITS, TAB_CHARTS)
End Function
```

Replace both Array() literals with `GetDevModeTabs()`.

---

## TD-04: Silent error exit audit

Audit every `Exit Function` and `Exit Sub` in KernelFormula.bas (and KernelFormulaWriter.bas after the split).

For each early exit:
- If it's a **normal guard clause** (e.g., `If count = 0 Then Exit Sub`) — leave as-is
- If it's a **failure path** (e.g., config not found, tab doesn't exist, formula parse error) — add `LogError` or `LogEvent` before the exit

Do NOT add logging to performance-sensitive inner loops.

---

## TD-05: Fix BOM in formula_tab_config.csv

The file `config_insurance/formula_tab_config.csv` has a UTF-8 BOM (EF BB BF) embedded inside the first quoted field. The header reads as `"﻿""TabName"""` instead of `"TabName"`.

**Fix:**
1. Remove the BOM from the CSV file (re-save as UTF-8 without BOM, or strip the first 3 bytes if they are EF BB BF)
2. Add a BOM-stripping safety net in KernelConfigLoader: when reading the first cell of any CSV, strip leading BOM characters (Chr(65279) in VBA, or check for byte sequence EF BB BF)
3. Check all other config_insurance/*.csv files for the same BOM issue

---

## TD-06: Standardize config-missing behavior

Document and enforce this rule:

**Domain config missing = fail fast with LogError (E-level error code)**
- formula_tab_config.csv rows referencing nonexistent RowIDs
- named_range_registry referencing nonexistent tabs
- curve_library_config missing required LOB/CurveType combinations

**Infrastructure config missing = fallback with LogEvent (W-level warning)**
- Snapshot files not found → skip restore, log warning
- Optional config tables empty → use defaults, log warning
- Print/chart/exhibit config empty → skip feature, log warning

Add a comment block at the top of KernelConfigLoader documenting this convention.

---

## VALIDATION

1. KernelFormula.bas < 35KB after split
2. KernelFormulaWriter.bas < 35KB and contains all tab-writing logic
3. No literal sheet name strings outside of KernelConstants.bas Const declarations
4. devTabs defined in exactly one place (GetDevModeTabs)
5. No silent Exit Function/Sub on failure paths in KernelFormula*.bas
6. No BOM in any config_insurance/*.csv header
7. Config-missing convention documented in KernelConfigLoader
8. All existing tests still pass (Run Model + RunTests)
9. BS_CHECK = 0, CFS_CHECK = 0 after all changes
10. SESSION_NOTES.md APPENDED only

## LOGGING

Append to SESSION_NOTES.md. Update CLAUDE.md (kernel module count 23→24, add KernelFormulaWriter to the list). Log any bugs discovered during refactoring. Sync config/ → config_insurance/ before ZIP delivery.

Do not bump version — this is a maintenance cycle, not a feature release.
