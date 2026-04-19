# Phase 12A: Detail Tabs — Design Specification

**Date:** 2026-03-31
**Status:** LOCKED — all decisions from Q&A sessions applied.
**Depends on:** Phase 11B (v1.2.0) COMPLETE
**Scope:** 3 new tabs + Revenue Summary wiring change

---

## Locked Decisions

| ID | Decision |
|---|---|
| PD-01 | UW Program Detail: all 10 program blocks, full Gross→Ceded→Net waterfall per program |
| SI-01 | Software Income Detail: 5 flexible user-defined revenue types, seeded with SaaS Subscriptions, Implementation/Setup, Usage-Based/API, Licensing, Other |
| SI-02 | Software Income wires into Revenue Summary as a new "Software Revenue" line (distinct from Fee Income) |

---

## Tab 1: UW Program Detail (Formula tab — QuarterlyColumns=TRUE)

### Purpose
Per-program UW waterfall showing Gross→Ceded→Net for each of the 10 program slots. References QuarterlySummary per-program RowIDs (QS_{METRIC}_{entityIndex}).

### Layout

Each program block has ~15 rows:

```
=== PROGRAM {N}: {EntityName} ===                    (Section header)

  Premium
    Gross Written Premium          = {REF:QuarterlySummary!QS_G_WP_{N}}
    Gross Earned Premium           = {REF:QuarterlySummary!QS_G_EP_{N}}
    Ceded Written Premium          = {REF:QuarterlySummary!QS_CQ_WP_{N}}
    Ceded Earned Premium           = {REF:QuarterlySummary!QS_CQ_EP_{N}}
    Net Written Premium            = Gross - Ceded
    Net Earned Premium             = Gross - Ceded

  Losses & Reserves
    Gross Ultimate Loss            = {REF:QuarterlySummary!QS_G_ULT_{N}}
    Ceded Ultimate Loss            = {REF:QuarterlySummary!QS_C_ULT_{N}}
    Net Ultimate Loss              = Gross - Ceded
    Gross Loss Reserve (EOQ)       = {REF:QuarterlySummary!QS_G_UNPAID_{N}}
    Net Loss Reserve (EOQ)         = {REF:QuarterlySummary!QS_N_UNPAID_{N}}

  Ratios
    Gross Loss Ratio               = G_Ult / G_EP
    Net Loss Ratio                 = N_Ult / N_EP
    (spacer)
```

10 blocks × ~17 rows = ~170 rows in formula_tab_config.csv.

### RowID Convention

All RowIDs prefixed with `PD_` followed by metric and program number:
- PD_GWP_1, PD_GWP_2, ... PD_GWP_10
- PD_GEP_1, PD_NEP_1, PD_GULT_1, PD_NULT_1, PD_GLR_1, PD_NLR_1, etc.

### Entity Names

Program headers should reference the entity name from UW Inputs. Since entity names are written to the Assumptions tab row 3 starting at column C (per InsuranceDomainEngine.Initialize), the section header can use a static label like "Program 1" — the entity name will show alongside via the QuarterlySummary entity label column.

Alternatively, use a formula that reads the entity name: `=IF(Assumptions!C3="","(empty)",Assumptions!C3)` for Program 1, `=IF(Assumptions!D3="","(empty)",Assumptions!D3)` for Program 2, etc.

### No Named Ranges Needed

This tab is purely informational — no downstream formulas reference it. No named ranges required.

### GrandTotal Column

FALSE — per-program detail doesn't need a grand total across years.

---

## Tab 2: Other Revenue Detail (Formula tab — QuarterlyColumns=TRUE)

### Purpose
Breaks out non-UW, non-software, non-investment revenue. Feeds the REV_FEE placeholder on Revenue Summary.

### Layout

```
Row 1:  [Other Revenue Detail]                    (Section header)
Row 2:  Management Projections                     (basis label)

=== FEE INCOME ===
Row 5:  Section: "Fee Income"
Row 6:  ID: ORD_FEE_MGA      MGA Program Fees         [Input, blue, quarterly]
Row 7:  ID: ORD_FEE_ADMIN    Administrative Fees      [Input, blue, quarterly]
Row 8:  ID: ORD_FEE_OTHER    Other Fees               [Input, blue, quarterly]
Row 9:  ID: ORD_FEE_TOTAL    Total Fee Income          [Formula: sum of above]

=== CONSULTING ===
Row 11: Section: "Consulting Revenue"
Row 12: ID: ORD_CON_ACT      Actuarial Consulting     [Input, blue, quarterly]
Row 13: ID: ORD_CON_RISK     Risk Management          [Input, blue, quarterly]
Row 14: ID: ORD_CON_OTHER    Other Consulting         [Input, blue, quarterly]
Row 15: ID: ORD_CON_TOTAL    Total Consulting          [Formula: sum of above]

=== TOTAL ===
Row 17: ID: ORD_TOTAL        Total Other Revenue       [Formula: Fee + Consulting]
```

### Named Ranges

```
ORD_Q_FeeTot     → ORD_FEE_TOTAL   (Quarterly)
ORD_Q_ConTot     → ORD_CON_TOTAL   (Quarterly)
ORD_Q_Total      → ORD_TOTAL       (Quarterly)
```

### Revenue Summary Wiring

Change REV_FEE formula from `=0` to `={REF:Other Revenue Detail!ORD_FEE_TOTAL}`
Change REV_CONSULT formula from `=0` to `={REF:Other Revenue Detail!ORD_CON_TOTAL}`

---

## Tab 3: Software Income Detail (Hybrid tab — QuarterlyColumns=TRUE)

### Purpose
Software revenue by user-defined type. Feeds a NEW "Software Revenue" line on Revenue Summary (SI-02).

### Layout

```
Row 1:  [Software Income Detail]                  (Section header)
Row 2:  Management Projections                     (basis label)

=== REVENUE BY TYPE ===
Row 5:  Section: "Software Revenue by Type"
Row 6:  ID: SWI_TYPE1_NAME   [Input: "SaaS Subscriptions"]     Col B (label)
Row 6:  ID: SWI_TYPE1        [Input, blue, quarterly]           Col C+ (amounts)
Row 7:  ID: SWI_TYPE2_NAME   [Input: "Implementation/Setup"]
Row 7:  ID: SWI_TYPE2        [Input, blue, quarterly]
Row 8:  ID: SWI_TYPE3_NAME   [Input: "Usage-Based/API"]
Row 8:  ID: SWI_TYPE3        [Input, blue, quarterly]
Row 9:  ID: SWI_TYPE4_NAME   [Input: "Licensing"]
Row 9:  ID: SWI_TYPE4        [Input, blue, quarterly]
Row 10: ID: SWI_TYPE5_NAME   [Input: "Other Software"]
Row 10: ID: SWI_TYPE5        [Input, blue, quarterly]

=== TOTAL ===
Row 12: ID: SWI_TOTAL        Total Software Revenue    [Formula: sum of 5 types]
```

### Named Ranges

```
SWI_Q_Total      → SWI_TOTAL      (Quarterly)
```

### Revenue Summary Wiring

Add a NEW row on Revenue Summary between REV_FEE and REV_CONSULT:
- RowID: REV_SOFTWARE
- Label: "Software Revenue"
- Formula: `={REF:Software Income Detail!SWI_TOTAL}`

Update REV_OTHREV formula to include REV_SOFTWARE:
`={ROWID:REV_INVEST}+{ROWID:REV_FEE}+{ROWID:REV_SOFTWARE}+{ROWID:REV_CONSULT}`

Renumber subsequent Revenue Summary rows as needed.

---

## Data Flow

```
QuarterlySummary (per-program) ──→ UW Program Detail (read-only)

Other Revenue Detail (inputs) ──→ Revenue Summary (REV_FEE, REV_CONSULT)
Software Income Detail (inputs) ──→ Revenue Summary (REV_SOFTWARE, new line)
                                         │
                                         ▼
                                    Income Statement (via IS_OTHREV)
```

---

## Tab Registry Additions

| Tab | Type | Category | SortOrder | QuarterlyColumns | QuarterlyHorizon | GrandTotal |
|---|---|---|---|---|---|---|
| UW Program Detail | Domain | Output | 13 | TRUE | Writing | FALSE |
| Other Revenue Detail | Domain | Input | 14 | TRUE | Writing | TRUE |
| Software Income Detail | Domain | Input | 15 | TRUE | Writing | TRUE |
