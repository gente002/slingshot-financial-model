# Phase 12B: Expense Detail + Staffing + Sales Funnel — Design Specification

**Date:** 2026-04-01
**Status:** LOCKED
**Depends on:** Phase 12A COMPLETE
**Scope:** 3 new tabs + Expense Summary rewiring + PD-05 sign convention + version bump

---

## Locked Decisions

| ID | Decision |
|---|---|
| OE-01 | Other Expense Detail: Grouped — Personnel (Salaries, Benefits, Contractors) + Non-Personnel (Rent/Facilities, Travel, Technology, Other) |
| OE-02 | Staffing + Other Expense Detail replace inline inputs on Expense Summary. Detail tabs are source of truth. |
| OE-03 | Annual inputs with even quarterly spread (annual ÷ 4) |
| ST-01 | Staffing Expense: 6 departments (UW, Claims, Finance, Tech, Executive, Other) |
| SF-01 | Sales Funnel: Inform only, copy/paste-ready output matching UW Inputs column layout |
| SF-02 | Universe → cohort % split → per-cohort conversion funnel. One UW Inputs row per cohort. |
| SF-03 | Single bind quarter per cohort |
| SF-04 | Fresh universe each year |
| SF-05 | Visible, after FM tabs |
| SF-06 | Per-cohort avg premium |
| SF-07 | Max 10 cohorts |
| SF-08 | Pipeline only — no known programs section |
| PD-05 | Ceded amounts show negative with parentheses format across UW Program Detail and UW Exec Summary |

---

## Tab 1: Staffing Expense (Hybrid — QuarterlyColumns=TRUE)

### Purpose
Headcount and loaded cost by department. Feeds EXP_STAFF on Expense Summary.

### Layout

```
Row 1:  [Staffing Expense]                           Section
Row 2:  Management Projections                        Basis

=== HEADCOUNT BY DEPARTMENT ===
Row 4:  Section "Headcount"                           D9E1F2
Row 5:  Headers: Department | Y1 | Y2 | Y3 | Y4 | Y5    (static cols C-G)
Row 6:  STF_HC_UW        Underwriting        [Input, blue — integer headcount per year]
Row 7:  STF_HC_CLAIMS    Claims              [Input, blue]
Row 8:  STF_HC_FINANCE   Finance             [Input, blue]
Row 9:  STF_HC_TECH      Technology          [Input, blue]
Row 10: STF_HC_EXEC      Executive           [Input, blue]
Row 11: STF_HC_OTHER     Other               [Input, blue]
Row 12: STF_HC_TOTAL     Total Headcount      [Formula: sum]              Bold, Thin border

=== AVERAGE LOADED COST ===
Row 14: Section "Average Loaded Cost (Annual)"        D9E1F2
Row 15: STF_COST_UW      Underwriting        [Input, blue — annual $ per head]
Row 16: STF_COST_CLAIMS  Claims              [Input, blue]
Row 17: STF_COST_FINANCE Finance             [Input, blue]
Row 18: STF_COST_TECH    Technology          [Input, blue]
Row 19: STF_COST_EXEC    Executive           [Input, blue]
Row 20: STF_COST_OTHER   Other               [Input, blue]
Row 21: STF_COST_AVG     Weighted Avg Cost    [Formula: total cost / total HC]  Italic, grey

=== ANNUAL STAFFING EXPENSE ===
Row 23: Section "Annual Staffing Expense"             D9E1F2
Row 24: STF_ANN_UW       Underwriting        [Formula: HC × Cost]
Row 25: STF_ANN_CLAIMS   Claims              [Formula: HC × Cost]
Row 26: STF_ANN_FINANCE  Finance             [Formula: HC × Cost]
Row 27: STF_ANN_TECH     Technology          [Formula: HC × Cost]
Row 28: STF_ANN_EXEC     Executive           [Formula: HC × Cost]
Row 29: STF_ANN_OTHER    Other               [Formula: HC × Cost]
Row 30: STF_ANN_TOTAL    Total Annual         [Formula: sum]              Bold, Thin border

=== QUARTERLY STAFFING EXPENSE ===   (QuarterlyColumns section)
Row 32: Section "Quarterly Staffing Expense"          D9E1F2
Row 33: STF_Q_UW         Underwriting        [Formula: Annual UW ÷ 4]
Row 34: STF_Q_CLAIMS     Claims              [Formula: Annual Claims ÷ 4]
Row 35: STF_Q_FINANCE    Finance             [Formula: Annual Finance ÷ 4]
Row 36: STF_Q_TECH       Technology          [Formula: Annual Tech ÷ 4]
Row 37: STF_Q_EXEC       Executive           [Formula: Annual Exec ÷ 4]
Row 38: STF_Q_OTHER      Other               [Formula: Annual Other ÷ 4]
Row 39: STF_Q_TOTAL      Total Staffing       [Formula: sum]              Bold, Double border
```

### Headcount/Cost Input Approach

Headcount and cost inputs are **annual by year** (5 columns: Y1-Y5) in the static section (rows 6-11, 15-20). These are NOT quarterly columns — they're fixed in columns C-G.

The quarterly section (rows 33-39) computes: for each quarter, look up the year based on column position, multiply HC × Cost for that year, divide by 4.

**Formula pattern for quarterly rows:**
`=INDEX($C$6:$G$6,1,INT((COLUMN()-{QS_DATA_START})/5)+1)*INDEX($C$15:$G$15,1,INT((COLUMN()-{QS_DATA_START})/5)+1)/4`

Simpler: use the annual total from the same year:
`=INDEX($C$24:$G$24,1,INT((COLUMN()-{QS_DATA_START})/5)+1)/4`

Where `{QS_DATA_START}` is the first quarterly data column (typically col 3 = C).

**Actually — simpler still:** Since the Annual section already computes HC × Cost per year, the quarterly formula just divides the corresponding annual total by 4. Use `INT((COLUMN()-3)/5)+1` to get year index (1-5), then INDEX into the annual row.

### Named Ranges

```
"STF_Q_Total","Staffing Expense","STF_Q_TOTAL","","Quarterly","Total quarterly staffing expense"
```

---

## Tab 2: Other Expense Detail (Hybrid — QuarterlyColumns=TRUE)

### Purpose
Non-staffing operating expenses by category. Feeds EXP_OTHER on Expense Summary.

### Layout

```
Row 1:  [Other Expense Detail]                        Section
Row 2:  Management Projections                         Basis

=== PERSONNEL EXPENSES (non-salary) ===
Row 4:  Section "Personnel Expenses"                   D9E1F2
Row 5:  OED_PER_BENEFITS   Benefits & Insurance       [Input, blue — annual Y1-Y5 in cols C-G]
Row 6:  OED_PER_CONTRACT   Contractors                [Input, blue]
Row 7:  OED_PER_RECRUIT    Recruiting                 [Input, blue]
Row 8:  OED_PER_TOTAL      Total Personnel             [Formula: sum]     Bold, Thin border

=== NON-PERSONNEL EXPENSES ===
Row 10: Section "Non-Personnel Expenses"               D9E1F2
Row 11: OED_NP_RENT        Rent / Facilities          [Input, blue — annual Y1-Y5]
Row 12: OED_NP_TRAVEL      Travel & Entertainment     [Input, blue]
Row 13: OED_NP_TECH        Technology & Software      [Input, blue]
Row 14: OED_NP_PROF        Professional Services      [Input, blue]
Row 15: OED_NP_INSUR       Insurance (D&O, E&O)       [Input, blue]
Row 16: OED_NP_OTHER       Other Expenses             [Input, blue]
Row 17: OED_NP_TOTAL       Total Non-Personnel         [Formula: sum]     Bold, Thin border

=== TOTAL ANNUAL ===
Row 19: OED_ANN_TOTAL      Total Annual Other Expense  [Formula: Pers + NonPers]  Bold

=== QUARTERLY OTHER EXPENSE ===   (QuarterlyColumns section)
Row 21: Section "Quarterly Other Operating Expense"    D9E1F2
Row 22: OED_Q_PERSONNEL    Personnel Expenses          [Formula: Annual Personnel ÷ 4]
Row 23: OED_Q_NONPERS      Non-Personnel Expenses      [Formula: Annual NonPers ÷ 4]
Row 24: OED_Q_TOTAL        Total Other Operating        [Formula: sum]    Bold, Double border
```

### Input Approach

Same pattern as Staffing: annual inputs in columns C-G (Y1-Y5), quarterly section divides by 4 using year-index lookup.

**Formula for quarterly rows:**
`=INDEX($C$8:$G$8,1,INT((COLUMN()-3)/5)+1)/4` (for Personnel)
`=INDEX($C$17:$G$17,1,INT((COLUMN()-3)/5)+1)/4` (for Non-Personnel)

### Named Ranges

```
"OED_Q_Total","Other Expense Detail","OED_Q_TOTAL","","Quarterly","Total quarterly other operating expense"
```

---

## Tab 3: Sales Funnel (Hybrid — QuarterlyColumns=TRUE)

### Purpose
Pipeline planning tool. Universe → cohort split → conversion funnel → expected GWP output. Inform only (SF-01) — user manually enters programs on UW Inputs. Copy/paste-ready output section.

### Layout

```
Row 1:  [Sales Funnel]                                Section
Row 2:  Pipeline Planning Tool                         Basis

=== UNIVERSE DEFINITION ===
Row 4:  Section "Addressable Universe"                 D9E1F2
Row 5:  SF_UNIVERSE     Total Addressable MGAs         [Input, blue — single value, col C]
Row 6:  SF_ALLOC_CHECK  Allocation Check               [Formula: sum of cohort %s — must = 100%]

=== COHORT ALLOCATION ===
Row 8:  Section "Cohort Allocation"                    D9E1F2
Row 9:  Headers: Cohort Name | % of Universe | MGA Count | Avg Premium | Product Type

Row 10: SF_COH1_NAME    [Input: "Property Direct"]     col B
        SF_COH1_PCT     [Input: blue]                   col C  (% of universe)
        SF_COH1_CNT     [Formula: Universe × %]         col D
        SF_COH1_AVGPREM [Input: blue]                   col E  (avg GWP per program)
        SF_COH1_PRODUCT [Input: blue]                   col F  (Property/Casualty/Specialty)
Row 11: SF_COH2 ... (same pattern)
...
Row 19: SF_COH10 ...
Row 20: SF_COH_TOTAL    Total                           [Formula: sums]   Bold, Thin border

=== CONVERSION FUNNEL (per cohort) ===
Row 22: Section "Conversion Funnel"                    D9E1F2
Row 23: Headers: Metric | Cohort 1 | Cohort 2 | ... | Cohort 10

Row 24: SF_CONTACT_RATE   Contact Rate (%)              [Input row, blue — one value per cohort across cols C-L]
Row 25: SF_QUALIFY_RATE   Qualification Rate (%)         [Input row, blue]
Row 26: SF_QUOTE_RATE     Quote Rate (%)                 [Input row, blue]
Row 27: SF_BIND_RATE      Bind Rate (%)                  [Input row, blue]
Row 28: SF_BIND_QTR       Bind Quarter (1-20)            [Input row, blue — which quarter programs start]
Row 29: SF_RENEWAL_RATE   Renewal Rate (%)               [Input row, blue]
Row 30: SF_RENEWAL_GROWTH Renewal Growth (%)             [Input row, blue]

=== FUNNEL RESULTS ===
Row 32: Section "Funnel Results"                       D9E1F2
Row 33: SF_CONTACTED      MGAs Contacted                [Formula: Count × Contact%]
Row 34: SF_QUALIFIED       MGAs Qualified                [Formula: Contacted × Qualify%]
Row 35: SF_QUOTED          MGAs Quoted                   [Formula: Qualified × Quote%]
Row 36: SF_BOUND           Programs Bound                [Formula: Quoted × Bind%]
Row 37: SF_NEW_GWP         New Business GWP              [Formula: Bound × AvgPrem]
Row 38: SF_RENEW_GWP       Renewal GWP (Y2+)             [Formula: prior year bound × RenewalRate × (1+Growth)]

=== QUARTERLY OUTPUT (copy/paste-ready) ===
Row 40: Section "Expected Premium by Quarter"          D9E1F2
        (QuarterlyColumns section)

Row 41-50: SF_OUT_1 through SF_OUT_10
        One row per cohort showing quarterly GWP.
        Q before bind quarter = 0.
        Bind quarter = New GWP (annualized ÷ remaining quarters in Y1).
        Subsequent quarters = continuing GWP.
        Y2+ = prior year × (1 + Renewal Growth) × Renewal Rate.

Row 51: SF_OUT_TOTAL     Total Pipeline GWP             [Formula: sum]    Bold, Double border

=== VARIANCE ===
Row 53: SF_UW_TOTAL      UW Inputs Total GWP            [Formula: ={REF:UW Exec Summary!UWEX_GWP}]
Row 54: SF_VARIANCE      Variance (Pipeline - UW Inputs) [Formula: SF_OUT_TOTAL - SF_UW_TOTAL]
```

### Funnel Mechanics

The conversion funnel is a **static calculation** (not quarterly). It computes:
- Contacted = Universe × Cohort% × ContactRate
- Qualified = Contacted × QualifyRate
- Quoted = Qualified × QuoteRate
- Bound = Quoted × BindRate
- New GWP = Bound × AvgPrem

The **quarterly output** section then spreads this GWP across quarters based on BindQuarter.

### Quarterly Output Formula

For each cohort row in the quarterly section, the formula needs to:
1. Check if this quarter is >= bind quarter → if not, show 0
2. For the bind year: GWP for remaining quarters (annualized ÷ 4, prorated if bind is mid-year)
3. For renewal years: prior year GWP × renewal rate × (1 + growth)

**Simplification for v1:** New programs write for a full year starting at bind quarter. If bind = Q3Y1, then Q3Y1 and Q4Y1 get GWP/4 each. Y2 gets full year (renewal rate × (1+growth) × prior year). Fresh universe each year (SF-04) means new programs bind each year independently.

**Actually — even simpler for v1:** Each cohort produces one number: annual New Business GWP. The bind quarter determines when it starts. The quarterly output is:
- Quarter < bind quarter: 0
- Quarter >= bind quarter, same year: NewGWP ÷ 4
- Year 2+: PriorYearGWP × RenewalRate × (1 + RenewalGrowth) ÷ 4

This is a column-based formula using COLUMN() position to determine quarter/year relative to bind quarter.

### Named Ranges

```
"SF_Q_PipelineTotal","Sales Funnel","SF_OUT_TOTAL","","Quarterly","Total pipeline GWP"
```

---

## Expense Summary Rewiring (OE-02)

Remove the 3 inline input rows (EXP_STAFF_ANN, EXP_OTHER_ANN, EXP_GROWTH) from Expense Summary.

Replace EXP_STAFF formula: from `=$C$12*(1+$C$14)^(...)` to `={REF:Staffing Expense!STF_Q_TOTAL}`
Replace EXP_OTHER formula: from `=$C$13*(1+$C$14)^(...)` to `={REF:Other Expense Detail!OED_Q_TOTAL}`

Remove the Spacer between inputs and computed rows. Renumber.

---

## PD-05: Negative Sign Convention

Apply to both UW Program Detail and UW Exec Summary:
- All ceded/deduction values: negate formula, format `#,##0;(#,##0)`
- Net rows: change from subtraction to addition
- Affected rows on UW Exec Summary: UWEX_CWP, UWEX_CEP, UWEX_CLLAE, UWEX_CEDCOMM
- Affected rows on UW Program Detail: PD_CEP_{N}, PD_CULT_{N}, PD_CCOMM_{N}, PD_GFFE_{N} (and Total block equivalents)

---

## Tab Registry

| Tab | SortOrder | QuarterlyColumns | GrandTotal |
|---|---|---|---|
| Staffing Expense | 15 | TRUE | TRUE |
| Other Expense Detail | 16 | TRUE | TRUE |
| Sales Funnel | 17 | TRUE | FALSE |

---

## Data Flow

```
Staffing Expense (dept HC × cost) ──→ Expense Summary (EXP_STAFF)
Other Expense Detail (Pers + NonPers) ──→ Expense Summary (EXP_OTHER)
                                              │
                                              ▼
                                         Income Statement → BS → CFS

Sales Funnel (pipeline) ──→ informational (no formula link)
                            user copy/pastes to UW Inputs
```
