# Outstanding Questions

**Status:** Living document. Add questions as they arise. Resolve with decisions and dates.

---

## Open

### OQ-001: Tax Payment Timing
**Tab:** Balance Sheet (BS_TAX_PAY), Income Statement (IS_TAX)
**Question:** Are income taxes paid quarterly or annually?
**Context:** Currently BS_TAX_PAY = IS_TAX (quarterly tax expense treated as quarterly payment). If taxes are paid annually, the balance sheet should accumulate quarterly tax expense and show a lump payment in Q4. This changes BS_TAX_PAY from `={REF:Income Statement!IS_TAX}` to a cumulative formula with annual reset.
**Impact:** Balance Sheet tax payable line, Cash Flow Statement change-in-tax-payable, and potentially a separate "Tax Paid" line in CFS.
**Decision:** TBD. Ask: does the entity make estimated quarterly tax payments or a single annual payment?
**Date Added:** 2026-03-30

### OQ-003: BS_AR (Premium Receivable) Modeling

### OQ-003: BS_AR (Premium Receivable) Modeling
**Tab:** Balance Sheet
**Question:** Is `GWP * 0.5` the right proxy for premium receivable?
**Context:** Currently BS_AR = UWEX_GWP * 0.5 (half of quarterly GWP assumed outstanding). This is a rough simplification. At annual columns, it would incorrectly multiply annual total GWP by 0.5 instead of the quarter-end value. A more precise model would track receivable aging or use earned/written spread.
**Impact:** Balance Sheet accuracy, Cash Flow Statement AR change calculation.
**Decision:** TBD.
**Date Added:** 2026-03-30

---

## Resolved

### OQ-002: Balance Sheet Annual Total Columns
**Tab:** Balance Sheet
**Question:** What should the annual total column show for a point-in-time statement?
**Decision:** Option A -- leave annual total columns blank for BS rows. Column still exists for layout consistency with IS/CFS, but BS formula rows write no value. Grey shading preserved. Grand total column also skipped.
**Resolved:** 2026-03-30
**Implementation:** KernelFormula.bas checks `tabName = "Balance Sheet"` and skips annual/grand total formula writes for formula rows.
