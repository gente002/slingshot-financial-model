# Phase 11B — Excel Validation Walkthrough

After CC builds Phase 11B, open the workbook and validate.

## Pre-Flight

1. Open workbook. Run Setup if needed.
2. Click "Run Model" on Dashboard.

## Gate 1: Tab Structure

- [ ] 7 new visible tabs: Investments, Capital Activity, Revenue Summary, Expense Summary, Income Statement, Balance Sheet, Cash Flow Statement
- [ ] Tab order: Dashboard, Assumptions, UW Inputs, UW Exec Summary, Detail, Investments, Capital Activity, Revenue Summary, Expense Summary, Income Statement, Balance Sheet, Cash Flow Statement
- [ ] FinancialSummary tab is gone
- [ ] Each new tab has navy header, "Management Projections" subtitle
- [ ] Quarterly columns start at col C

## Gate 2: UW Exec Summary — Reserves & UPR

- [ ] Earned Ratio row is removed (NEP is followed directly by Losses section)
- [ ] "Retained Margin" label (not "Retention Margin")
- [ ] Reserves section: Gross, Ceded, Net Loss Reserve rows present
- [ ] "Gross Unearned Premium Reserve" row exists (UWEX_GUPR = GWP - GEP)
- [ ] "Net Unearned Premium Reserve" row exists (UWEX_NUPR = GUPR - Ceded UPR)
- [ ] No #REF! errors in any row below
- [ ] Named ranges UWEX_Q_GUPR, UWEX_Q_NUPR, UWEX_Q_CRSV exist (Formulas > Name Manager)

## Gate 3: Investments Tab

- [ ] 6 asset rows with blue input cells, allocations sum to 100%
- [ ] Allocation Check shows "OK"
- [ ] Quarterly section: Pool = Net Reserves + Net UPR + CumulEquity (NRSV + NUPR)
- [ ] Liquid Cash = MIN(Pool, MAX(Floor, Pool × LiquidAlloc%)) — floor always respected
- [ ] Invested Assets = MAX(0, Pool - Liquid Cash) — remainder after liquid
- [ ] Set large liquidity floor (> Pool) → Liquid = Pool, Invested = 0 (edge case)
- [ ] Q1Y1 investment income = 0 (no prior-quarter balance to earn on)
- [ ] Q2+ income = prior quarter balance × yield ÷ 4 (BOQ approach, no circularity)

## Gate 4: Capital Activity Tab

- [ ] Enter $10M equity Q1Y1 → Cumulative = $10M
- [ ] Enter $5M surplus note draw Q1Y1, $1M repay Q4 → Balance tracks
- [ ] Interest = prior quarter balance × rate ÷ 4, Q1 interest = 0 (BOQ approach)
- [ ] Net Financing = raises + draws - repays

## Gate 5: Revenue Summary

- [ ] NEP matches UW Exec Summary
- [ ] Fronting Fees match UW Exec Summary
- [ ] Investment Income matches Investments tab
- [ ] NO ceding commission line in revenue (FM-D34)

## Gate 6: Expense Summary

- [ ] 3 blue input cells: Staffing Y1, Other Y1, Growth Rate
- [ ] Enter $1M staffing → Q1 = $250K
- [ ] Y2 quarters show growth ($1M × 1.03 / 4 = $257.5K)
- [ ] Net Acquisition = Gross Comm - Ceding Comm

## Gate 7: Income Statement

- [ ] Tax = MAX(0, PreTax × 21%). Negative PreTax → Tax = 0
- [ ] Other Income = blue input (editable)
- [ ] Key Ratios — Revenue Composition includes Fronting Fee : Total Rev row
- [ ] Key Ratios: no #DIV/0! (IFERROR wrapping)
- [ ] Growth: Q1 = 0, Q2+ shows QoQ

## Gate 8: Balance Sheet

- [ ] Premium Receivable = GWP × 0.5
- [ ] UPR = UWEX_GUPR
- [ ] RE = prior RE + Net Income
- [ ] **BALANCE CHECK = 0 for ALL quarters** ← critical

## Gate 9: Cash Flow Statement

- [ ] Echo rows show correct EOQ values (italic, grey)
- [ ] Δ rows = current - prior (Q1 Δ = Q1 value since prior = 0)
- [ ] **CASH RECONCILIATION = 0 for ALL quarters** ← critical
- [ ] Signs: +ΔReserves, +ΔUPR, -ΔAR, -ΔRI

## Gate 10: Smoke Test

1. Enter 3 programs on UW Inputs ($5M, $10M, $3M GWP Y1)
2. Enter $15M equity Q1Y1 on Capital Activity
3. Enter $5M surplus note draw Q1Y1 at 6%
4. Enter $500K staffing, $200K other opex on Expense Summary
5. Run Model
6. Verify: BS balances, CFS reconciles, no #REF!, investment income grows over time

## Gate 11: Counters

- [ ] Version = 1.2.0
- [ ] SESSION_NOTES.md has 11B section (not truncated)
- [ ] docs/*.md all present
- [ ] Stale files cleaned up
