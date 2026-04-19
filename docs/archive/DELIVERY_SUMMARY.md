# Phase 12B+ Delivery Summary -- Curve Calibration, Triangles, Model Integrity

**Version:** 1.3.0
**Date:** 2026-04-02
**Prior Phase:** 12B COMPLETE at v1.3.0

---

## What Was Built (Post-12B Session)

### Development Curve Recalibration (DE-09)
- 8 curves (Property/Casualty x Paid/CI/ReportedCount/ClosedCount)
- 5-point anchor interpolation (TL=1,25,50,75,100) with per-TL P2
- Calibrated from Industry Schedule P Reserve Analysis 2024
- MaxAge=240 universal, ordering constraints verified

### Mid-Month Average Written Date (DE-08)
- CDF evaluated at age - 0.5 for all curve types
- Applied consistently across DevelopLosses, GranularCSV, Triangles

### EP-Based Ultimates (DE-10, BUG-117)
- m_ultMon = m_epMon * ELR (single source of truth)
- All development = m_ultMon * CDF (no band-aids)
- CSV, Detail, QuarterlySummary, Triangles all derive from same computation

### New Tabs
- **Curve Reference** -- 8 blocks of development curves at 10 TL increments
- **Loss Triangles** -- Accident quarter Paid + Case Incurred with All Programs total
- **Count Triangles** -- Closed Count + Reported Count with All Programs total

### Validation Suite
- 14 Prove-It checks (run every model execution)
- 22-point Ins_Tests VBA suite (reserve identities, BS balance, curve ordering)

### Bugs Fixed
BUG-100 through BUG-117 (18 bugs). Key fixes:
- BUG-100: m_horizon=0 at Initialize
- BUG-101: Tail column loses PD-05 negation
- BUG-102: Staffing Expense multiplicative formula annual totals
- BUG-104: Triangle SUMIFS cannot isolate exposure cohorts
- BUG-105-117: Calendar vs exposure period architecture (culminating in BUG-117 root cause fix)

---

## Flywheel Cleanup

- Archived: PHASE12B_BUILD_PROMPT.md, Phase12B_Excel_Validation_Walkthrough.md, CURVE_REFERENCE_BUILD_PROMPT.md
- Moved: Sample_Output.xlsx, Triangle_Calculations.xlsx to docs/
- Cleaned: 47 granular CSV outputs, scenarios, WAL, stale workbook copy
- All 32 .bas files: AP-06 clean, Sub/Function balanced, CRLF, under 64KB
