# RDK Phase 12B: Expense Detail + Staffing + Sales Funnel — Claude Code Build Prompt

**Date:** April 2026
**Prior Phase:** Phase 12A COMPLETE
**Phase 12B Goal:** Build 3 new tabs (Staffing Expense, Other Expense Detail, Sales Funnel), rewire Expense Summary to reference the detail tabs, apply PD-05 negative sign convention to ceded amounts, bump version to 1.3.0.

## BEFORE WRITING ANY CODE

1. Read `CLAUDE.md` — project bible.
2. Read `SESSION_NOTES.md` — **APPEND ONLY — do not truncate.**
3. Read `data/anti_patterns.csv` — all entries.
4. Read `data/patterns.csv` — all entries.
5. Read `docs/Phase12B_Expense_Staffing_Funnel_Design.md` — the full design specification.
6. Examine `config_insurance/formula_tab_config.csv` — understand Expense Summary inline inputs (rows 12-14) that will be replaced.
7. Examine the UW Program Detail and UW Exec Summary formula rows — you will apply PD-05 sign changes to ceded amounts.

## DECISIONS (All Locked)

| ID | Decision |
|---|---|
| OE-01 | Other Expense: Personnel (Benefits, Contractors, Recruiting) + Non-Personnel (Rent, Travel, Tech, Professional, Insurance, Other) |
| OE-02 | Staffing + Other Expense Detail replace Expense Summary inline inputs |
| OE-03 | Annual inputs ÷ 4 for quarterly |
| ST-01 | Staffing: 6 departments (UW, Claims, Finance, Tech, Executive, Other) |
| SF-01 | Funnel: inform only, copy/paste-ready output |
| SF-02 | Universe → cohort % → per-cohort conversion |
| SF-03 | Single bind quarter per cohort |
| SF-04 | Fresh universe each year |
| SF-05 | Visible after FM tabs |
| SF-06 | Per-cohort avg premium |
| SF-07 | Max 10 cohorts |
| SF-08 | Pipeline only — no known programs |
| PD-05 | Ceded amounts: negate formula, format #,##0;(#,##0), net rows use addition |

---

## STEP 1: TAB REGISTRY

Add 3 tabs to `config_insurance/tab_registry.csv`:

```
"Staffing Expense","Domain","Input","N","Visible","15","Headcount and loaded cost by department","TRUE","Writing","TRUE"
"Other Expense Detail","Domain","Input","N","Visible","16","Non-staffing operating expenses by category","TRUE","Writing","TRUE"
"Sales Funnel","Domain","Input","N","Visible","17","Pipeline planning: universe, cohort conversion, expected GWP","TRUE","Writing","FALSE"
```

---

## STEP 2: STAFFING EXPENSE TAB

Hybrid tab. The top section has **annual inputs** (5 columns for Y1-Y5 in cols C-G, NOT quarterly columns). The bottom section has **quarterly formulas** that divide the annual amounts by 4.

### Static Section: Headcount + Cost (rows 4-21)

These rows use cols C-G for years 1-5. NOT quarterly columns — they are fixed-position annual inputs.

6 departments: UW, Claims, Finance, Tech, Executive, Other.

```
Row 4:  Section "Headcount by Department"
Row 5:  Headers: Department | Y1 | Y2 | Y3 | Y4 | Y5  (cols B-G)
Row 6-11:  STF_HC_UW through STF_HC_OTHER  [Input, blue, integer, per year]
Row 12: STF_HC_TOTAL  [Formula: SUM of 6 depts per column]  Bold, Thin border

Row 14: Section "Average Loaded Cost (Annual per Head)"
Row 15-20: STF_COST_UW through STF_COST_OTHER  [Input, blue, #,##0, per year]
Row 21: STF_COST_AVG  [Formula: total cost / total HC]  Italic, grey

Row 23: Section "Annual Staffing Expense by Department"
Row 24-29: STF_ANN_UW through STF_ANN_OTHER  [Formula: HC × Cost for each year]
Row 30: STF_ANN_TOTAL  [Formula: SUM]  Bold, Thin border
```

### Quarterly Section (rows 32+)

These rows ARE quarterly — they use the standard quarterly column layout.

```
Row 32: Section "Quarterly Staffing Expense"
Row 33-38: STF_Q_UW through STF_Q_OTHER
  Formula: =INDEX($C$24:$G$24,1,INT((COLUMN()-3)/5)+1)/4
  (Looks up the annual expense for the corresponding year, divides by 4)
  Replace row reference ($24) with the appropriate STF_ANN row for each dept.
Row 39: STF_Q_TOTAL  [Formula: SUM of 6 dept quarterly rows]  Bold, Double border
```

**IMPORTANT:** The `INT((COLUMN()-3)/5)+1` pattern gives year index 1-5 based on column position. This matches the 5-column-per-year quarterly layout (Q1,Q2,Q3,Q4,AnnualTotal). For the annual total column itself, the formula returns the full annual amount (INDEX picks the right year, ÷4 still applies — but the annual total column should show the annual SUM, which KernelFormula handles via the GT/annual-total logic).

### Named Ranges

```
"STF_Q_Total","Staffing Expense","STF_Q_TOTAL","","Quarterly","Total quarterly staffing expense"
```

---

## STEP 3: OTHER EXPENSE DETAIL TAB

Same pattern as Staffing — annual inputs in cols C-G, quarterly section divides by 4.

### Static Section (rows 4-19)

```
Row 4:  Section "Personnel Expenses (Non-Salary)"
Row 5:  OED_PER_BENEFITS    Benefits & Insurance     [Input, blue, annual Y1-Y5]
Row 6:  OED_PER_CONTRACT    Contractors              [Input, blue]
Row 7:  OED_PER_RECRUIT     Recruiting               [Input, blue]
Row 8:  OED_PER_TOTAL       Total Personnel           [Formula: SUM]   Bold, Thin border

Row 10: Section "Non-Personnel Expenses"
Row 11: OED_NP_RENT         Rent / Facilities        [Input, blue, annual Y1-Y5]
Row 12: OED_NP_TRAVEL       Travel & Entertainment   [Input, blue]
Row 13: OED_NP_TECH         Technology & Software    [Input, blue]
Row 14: OED_NP_PROF         Professional Services    [Input, blue]
Row 15: OED_NP_INSUR        Insurance (D&O, E&O)     [Input, blue]
Row 16: OED_NP_OTHER        Other Expenses           [Input, blue]
Row 17: OED_NP_TOTAL        Total Non-Personnel       [Formula: SUM]   Bold, Thin border

Row 19: OED_ANN_TOTAL       Total Annual Other Expense [Formula: Pers + NonPers]  Bold
```

### Quarterly Section (rows 21+)

```
Row 21: Section "Quarterly Other Operating Expense"
Row 22: OED_Q_PERSONNEL     Personnel Expenses        [Formula: Annual Personnel ÷ 4]
Row 23: OED_Q_NONPERS       Non-Personnel Expenses    [Formula: Annual NonPers ÷ 4]
Row 24: OED_Q_TOTAL         Total Other Operating      [Formula: SUM]   Bold, Double border
```

Quarterly formula: `=INDEX($C$8:$G$8,1,INT((COLUMN()-3)/5)+1)/4` (for Personnel)

### Named Ranges

```
"OED_Q_Total","Other Expense Detail","OED_Q_TOTAL","","Quarterly","Total quarterly other operating expense"
```

---

## STEP 4: SALES FUNNEL TAB

Hybrid tab. Static cohort inputs + quarterly output section.

### Universe & Cohort Allocation (rows 4-20)

```
Row 4:  Section "Addressable Universe"
Row 5:  SF_UNIVERSE       Total Addressable MGAs      [Input, blue, single value col C]

Row 7:  Section "Cohort Allocation & Characteristics"
Row 8:  Headers: Cohort Name | % of Universe | MGA Count | Avg Premium | Product Type
Row 9-18:  SF_COH1 through SF_COH10 (10 rows)
  Col B: Input (cohort name) — seed: "Cohort 1" through "Cohort 10"
  Col C: Input (% of universe, blue, 0.0%)
  Col D: Formula (= Universe × %)
  Col E: Input (avg GWP per program, blue, #,##0)
  Col F: Input (product type label, blue — seed: "Property", "Casualty", etc.)
Row 19: SF_COH_TOTAL  [Formula: sums for cols C-D]  Bold, Thin border
Row 20: SF_ALLOC_CHECK  [Formula: IF(ABS(C19-1)<0.001,"OK","ERROR")]
```

### Conversion Funnel (rows 22-30)

Transposed layout — metrics in rows, cohorts in columns (C-L = cohorts 1-10).

```
Row 22: Section "Conversion Funnel"
Row 23: Headers: Metric | Cohort 1 | Cohort 2 | ... | Cohort 10
Row 24: SF_CONTACT     Contact Rate          [Input, blue, 0.0% — one per cohort]
Row 25: SF_QUALIFY     Qualification Rate     [Input, blue, 0.0%]
Row 26: SF_QUOTE       Quote Rate             [Input, blue, 0.0%]
Row 27: SF_BIND        Bind Rate              [Input, blue, 0.0%]
Row 28: SF_BIND_QTR    Bind Quarter (1-20)    [Input, blue, integer]
Row 29: SF_RENEWAL     Renewal Rate           [Input, blue, 0.0%]
Row 30: SF_GROWTH      Renewal Growth         [Input, blue, 0.0%]
```

### Funnel Results (rows 32-38)

Same transposed layout — computed from inputs above.

```
Row 32: Section "Funnel Results"
Row 33: SF_RES_CONTACT   Contacted             = MGA Count × Contact%
Row 34: SF_RES_QUALIFY   Qualified              = Contacted × Qualify%
Row 35: SF_RES_QUOTE     Quoted                 = Qualified × Quote%
Row 36: SF_RES_BOUND     Bound                  = Quoted × Bind%
Row 37: SF_RES_NEWGWP    New Business GWP       = Bound × AvgPrem            Bold
Row 38: SF_RES_TOTALGWP  Total Pipeline GWP     = sum of all cohorts          Bold, Double border
```

### Quarterly Output — Copy/Paste Ready (rows 40+)

This section uses quarterly columns and shows expected GWP by cohort by quarter. Formatted to match UW Inputs column layout for easy copy/paste.

```
Row 40: Section "Expected Premium by Quarter (Copy/Paste to UW Inputs)"
Row 41-50: SF_OUT_1 through SF_OUT_10 (one row per cohort)
  Label in col B = cohort name (reference from col B of allocation section)
  Formula: for each quarterly column, compute expected GWP based on bind quarter:
    - Quarter < bind quarter: 0
    - Quarter >= bind quarter, year 1: NewGWP ÷ 4
    - Year 2+: PriorYearGWP × RenewalRate × (1+RenewalGrowth) ÷ 4

Row 51: Spacer
Row 52: SF_OUT_TOTAL     Total Pipeline GWP      [Formula: SUM]   Bold, Double border

Row 54: SF_UW_TOTAL      UW Inputs Total GWP     [Formula: ={REF:UW Exec Summary!UWEX_GWP}]
Row 55: SF_VARIANCE      Variance                 [Formula: SF_OUT_TOTAL - SF_UW_TOTAL]   Italic, grey
```

### Quarterly Output Formula Logic

For cohort row N, bind quarter = SF_BIND_QTR (from the funnel section, stored in a fixed cell reference). The formula needs:

```
=IF(quarter_number < bind_quarter, 0,
  IF(year_of_quarter = year_of_bind,
    new_gwp / 4,
    prior_year_gwp * renewal_rate * (1 + growth) / 4))
```

**Implementation approach:** Since {PREV_Q:} gives prior quarter, use it for year 2+ renewal calculation. For year 1, use the static NewGWP from funnel results. The bind quarter determines when to start.

This is complex formula logic — CC should implement the simplest version that works. An acceptable v1 simplification: just spread NewGWP evenly across all quarters from bind quarter onward, with no renewal decay/growth. Mark renewal as a Phase 12C enhancement if needed.

### Named Ranges

```
"SF_Q_PipelineTotal","Sales Funnel","SF_OUT_TOTAL","","Quarterly","Total pipeline GWP"
```

---

## STEP 5: EXPENSE SUMMARY REWIRING (OE-02)

### 5A: Remove inline inputs

Delete these rows from Expense Summary in formula_tab_config.csv:
- EXP_STAFF_ANN (row 12, both Label and Input entries)
- EXP_OTHER_ANN (row 13, both Label and Input entries)
- EXP_GROWTH (row 14, both Label and Input entries)
- EXP_SPACER3 (row 15)

### 5B: Rewire formulas

Change EXP_STAFF formula from:
`=$C$12*(1+$C$14)^(INT((COLUMN()-3)/5))/IF(MOD(COLUMN()-3,5)=4,1,4)`
To: `={REF:Staffing Expense!STF_Q_TOTAL}`

Change EXP_OTHER formula from:
`=$C$13*(1+$C$14)^(INT((COLUMN()-3)/5))/IF(MOD(COLUMN()-3,5)=4,1,4)`
To: `={REF:Other Expense Detail!OED_Q_TOTAL}`

### 5C: Renumber

After removing 4 rows (12-15), renumber EXP_STAFF, EXP_OTHER, EXP_OPEXP, EXP_SPACER4, EXP_TOTAL to fill the gap.

---

## STEP 6: PD-05 — NEGATIVE SIGN CONVENTION

Apply to both UW Program Detail AND UW Exec Summary.

### Format
All ceded/deduction rows: `#,##0;(#,##0)` — negatives show in parentheses.

### UW Exec Summary changes

| RowID | Current Formula | New Formula |
|---|---|---|
| UWEX_CWP | `={REF:QS!QS_CQ_WP_TOTAL}` | `=-{REF:QS!QS_CQ_WP_TOTAL}` |
| UWEX_CEP | `={REF:QS!QS_CQ_EP_TOTAL}` | `=-{REF:QS!QS_CQ_EP_TOTAL}` |
| UWEX_CLLAE | `={REF:QS!QS_C_ULT_TOTAL}` | `=-{REF:QS!QS_C_ULT_TOTAL}` |
| UWEX_CEDCOMM | `={REF:QS!QS_C_ECOMM_TOTAL}` | `=-{REF:QS!QS_C_ECOMM_TOTAL}` |
| UWEX_NWP | `=GWP-CWP` | `=GWP+CWP` (CWP now negative) |
| UWEX_NEP | `=GEP-CEP` | `=GEP+CEP` |
| UWEX_NLLAE | `=GLLAE-CLLAE` | `=GLLAE+CLLAE` |
| UWEX_NCOMM | `=GCOMM-CEDCOMM` | `=GCOMM+CEDCOMM` |

**Also negate:** UWEX_GFFE (fronting fee is a deduction from gross premium — should show negative).

**Update Net formulas downstream:** UWEX_NACQ, UWEX_GUWRES, UWEX_NUWRES, ratio formulas — all need to account for the sign change. The key principle: everything that was `A - B` where B is now negative becomes `A + B`.

**IMPORTANT:** The existing UW Exec Summary named ranges (UWEX_Q_NEP, UWEX_Q_NLLAE, etc.) are referenced by Revenue Summary, Expense Summary, Balance Sheet, and other tabs. When you negate the ceded values, make sure the NET values (which downstream tabs reference) remain POSITIVE. The net rows should produce the same absolute values as before — only the ceded intermediate rows change sign.

### UW Program Detail changes

Same pattern — negate PD_CEP_{N}, PD_CULT_{N}, PD_CCOMM_{N}, PD_GFFE_{N} formulas. Change Net row formulas from subtraction to addition. Apply to all 10 program blocks AND the Total block.

**CAREFUL:** The UW Program Detail also has Supporting Calculations (EOQ echo rows). Those should NOT be negated — they're balance sheet references used for {PREV_Q:} deltas.

---

## STEP 7: VERSION BUMP

Change `KERNEL_VERSION` in `engine/KernelConstants.bas` from `"1.2.0"` to `"1.3.0"`.

---

## STEP 8: SESSION_NOTES + CLAUDE.md

APPEND to SESSION_NOTES.md. Update CLAUDE.md counters. Sync config/ to config_insurance/ before delivery.

---

## VALIDATION GATES

1. tab_registry: 3 new tabs (SortOrder 15-17), no collisions
2. Staffing Expense: 6 dept rows × (HC + Cost + Annual + Quarterly), totals sum correctly
3. Other Expense Detail: Personnel + Non-Personnel grouping, annual ÷ 4 quarterly
4. Sales Funnel: 10 cohort rows, allocation check, funnel results, quarterly output
5. Expense Summary: inline inputs REMOVED, EXP_STAFF references Staffing, EXP_OTHER references OED
6. EXP_OPEXP and EXP_TOTAL still compute correctly after rewiring
7. PD-05: UW Exec Summary ceded values negative with parentheses format
8. PD-05: UW Program Detail ceded values negative with parentheses format
9. PD-05: Net values (NEP, NLLAE, etc.) unchanged in absolute terms — downstream tabs unaffected
10. BS still balances (BS_CHECK = 0) after all changes
11. CFS still reconciles (CFS_CHECK = 0)
12. IS flows correctly through the rewired Expense Summary
13. Named ranges: STF_Q_Total, OED_Q_Total, SF_Q_PipelineTotal
14. No module > 64KB HARD (watch KernelFormula at 58.7KB)
15. SESSION_NOTES appended, docs not deleted
16. Version = 1.3.0
17. config_insurance/ synced from config/ before ZIP

## PROTECTED FILES

- SESSION_NOTES.md (APPEND ONLY)
- CLAUDE.md
- All docs/*.md
- data/bug_log.csv, anti_patterns.csv, patterns.csv
- **DO NOT modify** UW Program Detail's Supporting Calculations EOQ echo rows (used for {PREV_Q:} deltas)

## DELIVERY

All modified files + DELIVERY_SUMMARY.md. Version = 1.3.0. Sync config/ → config_insurance/ before ZIP. Include full directory structure.
