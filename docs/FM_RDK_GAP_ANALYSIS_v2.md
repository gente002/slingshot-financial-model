# FM-RDK Gap Analysis — Updated Post Phase 6A

**Date:** 2026-03-26
**Kernel Version:** 1.0.8 (Phase 6A COMPLETE)
**Kernel Modules:** 23 kernel + 2 extensions
**Config Tables:** 18

---

## Gap Status Summary

| # | Gap | Severity | Status | Resolved In |
|---|-----|----------|--------|-------------|
| G-01 | Formula tab generation | ~~P0~~ | ✅ RESOLVED | Phase 5C — KernelFormula.CreateFormulaTabs + formula_tab_config.csv |
| G-02 | Named range creation | ~~P0~~ | ✅ RESOLVED | Phase 5C — KernelFormula.CreateNamedRanges + named_range_registry.csv |
| G-03 | Quarterly aggregation | ~~P0~~ | ✅ RESOLVED | Phase 5C — AggregateToQuarterly PostCompute transform → QuarterlySummary tab |
| G-04 | Hybrid tab support | P1 | ⚠️ OPEN | Design during FM Q&A. Investments tab and Capital Activity are hybrid (inputs + formulas). formula_tab_config supports Input CellType — may be sufficient. |
| G-05 | Circularity resolution | P1 | ⚠️ OPEN | Investment income ↔ cash balance. Enable iterative calc via KernelBootstrap. ~5 lines of VBA. |
| G-06 | Repeating block layout | P1 | ⚠️ OPEN | UW Program Detail needs 10 program blocks. Deferred — UW Program Detail is not in v1 10-tab scope. |
| G-07 | Waterfall layout | P1 | ⚠️ OPEN | UW Exec Summary needs Gross→Ceded→Net. DomainEngine + QuarterlySummary handle this — the waterfall is a data arrangement, not a kernel gap. formula_tab_config can define the row ordering. |

**Net assessment:** All P0 gaps resolved. Of the 4 remaining P1 gaps, G-06 is deferred (not in v1 scope), G-07 is solvable with existing formula_tab_config, G-05 is ~5 lines of VBA, and G-04 needs Q&A validation but likely works with existing CellType=Input. **No kernel gaps block the FM build.**

---

## Remaining P1 Gaps — Resolution Path

### G-04: Hybrid Tabs
The Investments tab and Capital Activity tab mix user inputs (blue cells) with computed formulas. The formula_tab_config.csv already has `CellType=Input` which writes default values with blue font. The question is whether this is sufficient for a full hybrid tab experience (user edits inputs, formulas auto-recalculate).

**Likely resolution:** Yes — CellType=Input cells are editable. CellType=Formula cells reference them. Excel handles the recalculation. Test during the Investments tab Q&A design session.

### G-05: Circularity
Investment income = f(invested assets) = f(cash balance) = f(investment income). Excel's iterative calculation resolves this. Add to KernelBootstrap:
```vba
Application.Iteration = True
Application.MaxIterations = 50
Application.MaxChange = 0.01
```
This is FM-D30 from the FM Component Spec. Three lines of VBA in BootstrapWorkbook.

### G-06: Repeating Blocks (DEFERRED)
UW Program Detail (10 program blocks) is not in the v1 10-tab scope. When built, formula_tab_config can define each block with program-prefixed RowIDs (UWPD_VISR_GWP, UWPD_VIEV_GWP, etc.). The kernel doesn't need a "repeating block generator" — the config just has more rows.

### G-07: Waterfall Layout
The Gross→Ceded→Net waterfall is a row ordering decision, not a kernel capability. The QuarterlySummary tab (from Phase 5C) can be structured with Gross metrics first, then Ceded, then Net. formula_tab_config defines the row layout on UW Exec Summary referencing these rows. No kernel change needed.

---

## What the Kernel Now Provides for the FM

| FM Need | Kernel Capability | Module |
|---|---|---|
| UW actuarial projections | DomainEngine.Execute via dynamic dispatch | KernelEngine (DomainModule setting) |
| Monthly → quarterly aggregation | AggregateToQuarterly PostCompute transform | SampleDomainEngine + KernelTransform |
| Formula-driven financial statements | formula_tab_config.csv + CreateFormulaTabs | KernelFormula |
| Named range bridge contract | named_range_registry.csv + CreateNamedRanges | KernelFormula |
| Cross-tab formula references | Placeholder resolution ({REF:Tab!RowID}, {ROWID:xxx}) | KernelFormula |
| Development curves | CDF functions + curve_library_config.csv | Ext_CurveLib |
| PDF report generation | Cover page + tab export + Prove-It summary | Ext_ReportGen |
| Print configuration | print_config.csv + ConfigurePrintSettings | KernelPrint |
| Balance checks | prove_it_config.csv + native Excel formulas | KernelProveIt |
| Snapshot save/load | KernelSnapshot + savepoint manifests | KernelSnapshot |
| Scenario comparison | KernelCompare + comparison tabs | KernelCompare |
| Extension activation | extension_registry.csv + KernelExtension | KernelExtension |

---

## FM Build Sequence (10 tabs)

| Phase | Tabs | Dependencies |
|---|---|---|
| 0 | DomainEngine (insurance) + UW Inputs (input_schema) | Replaces SampleDomainEngine. Produces monthly Detail with insurance metrics. |
| 1 | Assumptions + UW Executive Summary | Assumptions = formula tab with global params. UW Exec = formula tab referencing QuarterlySummary. |
| 2 | Investments + Capital Activity | Investments = hybrid (inputs + weighted yield). Capital = hybrid (inputs + running balances). Enable iterative calc (G-05). |
| 3 | Revenue Summary + Expense Summary | Formula tabs aggregating from UW Exec + Investments. FM-D34 net presentation. |
| 4 | Income Statement + Balance Sheet + Cash Flow Statement | Tier 1 formula tabs. IS references Rev/Exp Summary. BS references IS + UW Exec + Investments + Capital. CFS derived from ΔBS. |

Each phase: Q&A design session → formula_tab_config rows → named_range_registry entries → build → validate.
