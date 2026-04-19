# Phase 12B — Excel Validation Walkthrough

After CC builds Phase 12B, open the workbook and validate.

## Pre-Flight

1. Open workbook. Enter at least 3 programs on UW Inputs.
2. Click "Run Model" on Dashboard.

## Gate 1: Tab Structure

- [ ] 3 new tabs: Staffing Expense (15), Other Expense Detail (16), Sales Funnel (17)
- [ ] Version = 1.3.0 (check Dashboard or Assumptions)

## Gate 2: Staffing Expense

- [ ] 6 department rows: UW, Claims, Finance, Tech, Executive, Other
- [ ] Headcount inputs: annual Y1-Y5 in cols C-G (blue font)
- [ ] Loaded cost inputs: annual Y1-Y5 (blue font)
- [ ] Annual expense = HC × Cost per dept per year
- [ ] Quarterly section: each quarter = annual ÷ 4
- [ ] Enter HC=5 and Cost=$150K for UW in Y1 → Annual=$750K → Quarterly=$187.5K
- [ ] Total row sums all 6 departments
- [ ] Annual total column shows correct annual sum

## Gate 3: Other Expense Detail

- [ ] Personnel section: Benefits, Contractors, Recruiting (3 rows)
- [ ] Non-Personnel section: Rent, Travel, Tech, Professional, Insurance, Other (6 rows)
- [ ] All inputs annual Y1-Y5 (blue font)
- [ ] Quarterly = annual ÷ 4
- [ ] Enter $200K rent Y1 → quarterly = $50K
- [ ] Total = Personnel + Non-Personnel

## Gate 4: Expense Summary Rewiring

- [ ] Inline inputs GONE (no StaffExpY1, OtherExpY1, ExpGrowth rows)
- [ ] EXP_STAFF = Staffing Expense total (not hardcoded formula)
- [ ] EXP_OTHER = Other Expense Detail total
- [ ] Enter staffing and other expense on detail tabs → Expense Summary reflects them
- [ ] Total Expenses = UW Expenses + Operating Expenses (unchanged logic)

## Gate 5: End-to-End Expense Check

- [ ] Enter staffing + other expenses on detail tabs
- [ ] Run Model
- [ ] Expense Summary shows correct staffing + other values
- [ ] Income Statement OpExp matches Expense Summary
- [ ] BS still balances (BS_CHECK = 0 all quarters)
- [ ] CFS still reconciles (CFS_CHECK = 0 all quarters)

## Gate 6: Sales Funnel — Universe & Cohorts

- [ ] Universe input: single value (e.g., enter 500)
- [ ] 10 cohort rows with editable names
- [ ] Enter 20% for Cohort 1 → MGA Count = 100
- [ ] Avg Premium input per cohort (e.g., $3M)
- [ ] Product Type input (text label)
- [ ] Allocation check: if cohort %s don't sum to 100%, shows "ERROR"

## Gate 7: Sales Funnel — Conversion Funnel

- [ ] Contact, Qualify, Quote, Bind rates all input (blue, per cohort)
- [ ] Bind Quarter input (integer 1-20)
- [ ] Renewal Rate and Growth inputs
- [ ] Funnel Results: Contacted → Qualified → Quoted → Bound (multiplicative cascade)
- [ ] New Business GWP = Bound × Avg Premium
- [ ] Enter: 100 MGAs, 50% contact, 40% qualify, 60% quote, 25% bind = 3 programs
- [ ] 3 programs × $3M avg = $9M New GWP

## Gate 8: Sales Funnel — Quarterly Output

- [ ] 10 output rows (one per cohort) with quarterly columns
- [ ] Quarters before bind quarter show 0
- [ ] Bind quarter onward shows GWP/4
- [ ] Total Pipeline GWP row sums all cohorts
- [ ] Variance row shows Pipeline - UW Inputs difference
- [ ] Output format matches UW Inputs (ready for copy/paste)

## Gate 9: PD-05 — Negative Sign Convention

- [ ] UW Exec Summary: ceded values show in parentheses (225,000) not 225,000
- [ ] UW Exec Summary: Net values UNCHANGED in absolute terms (e.g., NEP still positive $12,500)
- [ ] UW Program Detail: same parentheses convention on ceded rows
- [ ] Revenue Summary, Expense Summary, IS, BS, CFS all still show correct values
- [ ] No #REF! errors anywhere

## Gate 10: Counters & Integrity

- [ ] Version = 1.3.0
- [ ] SESSION_NOTES.md has 12B section (not truncated)
- [ ] config_insurance/ synced (matches runtime config/)
- [ ] No #REF! errors in any tab
- [ ] BS_CHECK = 0 all quarters
- [ ] CFS_CHECK = 0 all quarters
