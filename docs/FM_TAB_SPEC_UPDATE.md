# FM Tab Specification — Updated per Ethan Feedback (2026-03-26)

**Status:** Locked decisions updated. 10 tabs for v1.
**Changes from prior version:** Items 1-7 from Ethan's feedback incorporated. FM-D34 revised. Investments tab added. IS restructured.

---

## Decision Updates

| ID | Prior | Updated | Reason |
|---|---|---|---|
| FM-D34 | Gross-up: ceding comm in revenue, gross comm in expense | **REVISED:** Net presentation — Net Acquisition Expense = Direct Comm - Ceding Comm | Ethan item 2: ceding commission is an expense offset, not revenue |
| FM-D37 | (new) | Investments tab is a Tier 3 input tab with allocation mix driving weighted-average yield | Ethan item 1 |
| FM-D38 | (new) | Income Statement shows Tier 1 categories only — details on Tier 2 exhibits | Ethan item 3 |
| FM-D39 | (new) | IS has symmetric Revenue/Expense structure with subtotals | Ethan item 4 |
| FM-D40 | (new) | IS includes full walk from Operating Income → Net Income (interest/debt expense, other income/expense, tax) | Ethan item 5 |
| FM-D41 | (new) | IS Key Ratios section uses revenue/expense composition ratios and growth metrics | Ethan items 6-7 |

---

## Updated Tab Manifest (10 tabs)

| # | Tab | Type | Tier | Feeds |
|---|---|---|---|---|
| 1 | Assumptions | Control | — | All tabs |
| 2 | UW Inputs | Input | 3 | UW Executive Summary (via RunProjections) |
| 3 | UW Executive Summary | Summary | 2 | Revenue Summary, Expense Summary, Balance Sheet |
| 4 | Investments | Input | 3 | Revenue Summary (investment income), Balance Sheet (asset allocation) |
| 5 | Capital Activity | Input | 3 | Balance Sheet, Cash Flow Statement, IS (interest expense) |
| 6 | Revenue Summary | Summary | 2 | Income Statement |
| 7 | Expense Summary | Summary | 2 | Income Statement |
| 8 | Income Statement | Summary | 1 | Balance Sheet (net income → retained earnings) |
| 9 | Balance Sheet | Summary | 1 | Cash Flow Statement |
| 10 | Cash Flow Statement | Summary | 1 | — (terminal output) |

---

## Tab 4: Investments (NEW — Tier 3 Input)

### Purpose
User defines investment allocation mix and per-class yield assumptions. Computes weighted-average portfolio yield that feeds the Investment Allocation Model (§10) and Balance Sheet asset lines.

### Layout

```
Row 1:  [Investment Portfolio]                    (navy header)
Row 2:  Management Projections — ...              (basis label)
Row 3:  (blank)

Section 1: Asset Allocation (rows 5-14)
  Col B: Asset Class
  Col C: Allocation %  [input, blue]
  Col D: Annual Yield  [input, blue]
  Col E: Duration (yrs) [input, blue]
  Col F: Credit Quality [input, blue — e.g., "AAA", "A", "BBB"]

  Row 5:   Government/Agency Bonds     40%    3.5%    4.0    AAA
  Row 6:   Corporate Bonds             25%    4.5%    3.5    A
  Row 7:   Municipal Bonds              5%    3.0%    5.0    AA
  Row 8:   Money Market / Cash         20%    2.0%    0.1    AAA
  Row 9:   Equities                     5%    6.0%    —      —
  Row 10:  Alternatives                 5%    7.0%    —      —
  Row 11:  (blank)
  Row 12:  Total Allocation           100%    [weighted avg]  [weighted avg]
  Row 13:  Allocation Check            =SUM(C5:C10)  [must = 100%, red if not]

Section 2: Portfolio Summary (rows 16-22)
  ID: INV_WTD_YIELD      Weighted-Average Yield = SUMPRODUCT(Alloc%, Yield%)
  ID: INV_WTD_DURATION   Weighted-Average Duration
  ID: INV_LIQUID_PCT     Liquid Allocation = Money Market %  [= C8]
  ID: INV_INVESTED_PCT   Invested Allocation = 1 - Liquid %
  ID: INV_FLOOR          Liquidity Floor  [input, blue — replaces CTRL_LiquidityFloor]
  ID: INV_LIQUID_YIELD   Liquid Cash Yield [= D8, Money Market yield]
  ID: INV_INVEST_YIELD   Invested Asset Yield [= weighted avg of non-cash classes]

Section 3: Quarterly Investment Income (rows 25+)
  [Formula-driven — references §10 Investment Allocation Model]
  ID: INV_Q_POOL         Investable Pool (from UW Exec: reserves + UPR float)
  ID: INV_Q_INVESTED     Invested Assets = Pool × INV_INVESTED_PCT (subject to floor)
  ID: INV_Q_LIQUID       Liquid Cash = Pool - Invested
  ID: INV_Q_INC_INVEST   Income on Invested = Avg(beg,end) × INV_INVEST_YIELD / 4
  ID: INV_Q_INC_LIQUID   Income on Liquid = Avg(beg,end) × INV_LIQUID_YIELD / 4
  ID: INV_Q_INC_TOTAL    Total Investment Income = Invest + Liquid
```

### Named Ranges

| Range | Type | Description |
|---|---|---|
| INV_WtdYield | Single | Weighted-average portfolio yield (replaces CTRL_InvestYield) |
| INV_LiquidYield | Single | Liquid cash yield (replaces CTRL_LiquidYield) |
| INV_LiquidityFloor | Single | Minimum liquid cash (replaces CTRL_LiquidityFloor) |
| INV_Q_IncTotal | Quarterly | Total investment income per quarter |
| INV_Q_Invested | Quarterly | Invested assets balance → BS |
| INV_Q_Liquid | Quarterly | Liquid cash balance → BS |

### Feeds

- **Revenue Summary:** INV_Q_IncTotal → Other Revenue → Investment Income line
- **Balance Sheet:** INV_Q_Invested → BS_INVEST, INV_Q_Liquid → BS_CASH
- **Cash Flow Statement:** Change in invested assets → CFI

### Assumptions Tab Changes

Remove from Assumptions: Investment Yield (C27), Liquidity Reserve % (C51), Target Liquidity Floor (C52), Liquid Cash Yield (C54), Investment Allocation % (C50). These all move to the Investments tab. Keep Risk-Free Rate (C26) on Assumptions as it's a global parameter not specific to the investment portfolio.

---

## Tab 6: Revenue Summary (UPDATED — FM-D34 revised)

### Layout

```
Section: Underwriting Revenue
  ID: REV_NEP          Net Earned Premium           = UWEX_Q_NEP
  ID: REV_FFE          Fronting Fees (earned)        = UWEX_Q_GFFE
  ID: REV_UWREV        Total UW Revenue             = NEP + FFE
                        (NOTE: Ceding commission NO LONGER appears here)

Section: Other Revenue
  ID: REV_INVEST       Investment Income             = INV_Q_IncTotal
  ID: REV_FEE          Fee Income                    = 0 (deferred — Other Rev Detail not built)
  ID: REV_CONSULT      Consulting Revenue            = 0 (deferred)
  ID: REV_OTHREV       Total Other Revenue           = Invest + Fee + Consult

Section: Total
  ID: REV_TOTAL        Total Revenue                = UW Revenue + Other Revenue
```

**What changed:** Ceding commission removed from UW Revenue. Under the net presentation (FM-D34 revised), ceding commission reduces acquisition expense — it doesn't appear as revenue.

---

## Tab 7: Expense Summary (UPDATED — FM-D34 revised)

### Layout

```
Section: Underwriting Expenses
  ID: EXP_NLLAE        Net Losses & LAE              = UWEX_Q_NLLAE
  ID: EXP_NACQ         Net Acquisition Expense        = UWEX_Q_GCOMM - UWEX_Q_CEDCOMM
                        (= Direct Commission - Ceding Commission)
  ID: EXP_OTHUW        Other UW Expense              = UWEX_Q_OTHUWEXP
  ID: EXP_UWEXP        Total UW Expenses             = NLLAE + Net Acq + Other UW

Section: Operating Expenses
  ID: EXP_STAFF        Staffing Expense              = 0 (deferred — Staffing tab not built)
  ID: EXP_OTHER        Other Operating Expense       = 0 (deferred — Other Exp Detail not built)
  ID: EXP_OPEXP        Total Operating Expenses      = Staff + Other

Section: Total
  ID: EXP_TOTAL        Total Expenses                = UW Expenses + Operating Expenses
```

**What changed:** Net Acquisition Expense replaces separate Gross Commission / Net Commission lines. Ceding commission is subtracted from gross commission here, not added to revenue.

---

## Tab 8: Income Statement (RESTRUCTURED — Items 3-7)

### Design Principles (FM-D38 through FM-D41)

- **Tier 1 only:** Major categories and subtotals. No Tier 2 detail. Details on Revenue Summary and Expense Summary.
- **Symmetric structure:** Revenue section and Expense section follow the same pattern (sub-categories → subtotal).
- **Full walk:** Revenue → Expenses → Operating Income → Other Items → Pre-Tax → Tax → Net Income.
- **Key Ratios:** Composition ratios (each category as % of total), not UW-specific ratios.
- **Growth:** QoQ and YoY growth for major totals.

### Layout

```
Row 1:  [Phronex — Income Statement]                    (navy header)
Row 2:  Management Projections — ...                      (basis label)
Row 3:  (blank)

===== REVENUE =====
  ID: IS_UWREV         Underwriting Revenue              = REV_UWREV
  ID: IS_OTHREV        Other Revenue                     = REV_OTHREV
  ID: IS_TOTALREV      Total Revenue                     = UW Rev + Other Rev
                        (bold, single top border)

===== EXPENSES =====
  ID: IS_UWEXP         Underwriting Expenses             = EXP_UWEXP
  ID: IS_OPEXP         Operating Expenses                = EXP_OPEXP
  ID: IS_TOTALEXP      Total Expenses                    = UW Exp + Operating Exp
                        (bold, single top border)

===== OPERATING INCOME =====
  ID: IS_OPINC         Operating Income                  = Total Rev - Total Exp
                        (bold, double top border)

===== BELOW THE LINE =====
  ID: IS_INTEXP        Interest / Debt Expense           = CAP_Q_IntExp
  ID: IS_OTHINC        Other Income / (Expense)          = 0 (manual input, blue)

===== PRE-TAX INCOME =====
  ID: IS_PRETAX        Pre-Tax Income                    = Operating Inc - Int Exp + Other Inc
                        (bold, single top border)

===== TAX =====
  ID: IS_TAX           Income Tax Expense                = MAX(0, Pre-Tax × CTRL_TaxRate)

===== NET INCOME =====
  ID: IS_NETINC        Net Income                        = Pre-Tax - Tax
                        (bold, double top border)

===== (blank separator) =====

===== KEY RATIOS — REVENUE COMPOSITION =====
  ID: IS_KR_NEP_REV    Net Earned Premium : Total Revenue    = UWEX_Q_NEP / IS_TOTALREV
  ID: IS_KR_INV_REV    Investment Income : Total Revenue     = INV_Q_IncTotal / IS_TOTALREV
  ID: IS_KR_OTH_REV    Other Revenue : Total Revenue         = (IS_OTHREV - INV_Q_IncTotal) / IS_TOTALREV

===== KEY RATIOS — EXPENSE COMPOSITION =====
  ID: IS_KR_LLAE_EXP   Net Losses & LAE : Total Expenses    = EXP_NLLAE / IS_TOTALEXP
  ID: IS_KR_ACQ_EXP    Net Acquisition Exp : Total Expenses = EXP_NACQ / IS_TOTALEXP
  ID: IS_KR_STAFF_EXP  Staffing : Total Expenses            = EXP_STAFF / IS_TOTALEXP
  ID: IS_KR_OTH_EXP    Other Expense : Total Expenses       = EXP_OTHER / IS_TOTALEXP

===== KEY RATIOS — PROFITABILITY =====
  ID: IS_KR_DEBT_OP    Interest Expense : Operating Income   = IS_INTEXP / IS_OPINC
  ID: IS_KR_TAX_NI     Tax Expense : Pre-Tax Income          = IS_TAX / IS_PRETAX

===== GROWTH =====
  ID: IS_GR_REV        Total Revenue Growth                 = (Rev_t - Rev_{t-1}) / Rev_{t-1}
  ID: IS_GR_EXP        Total Expense Growth                 = (Exp_t - Exp_{t-1}) / Exp_{t-1}
  ID: IS_GR_NI         Net Income Growth                    = (NI_t - NI_{t-1}) / NI_{t-1}
```

### What Changed from Prior IS Spec

1. **Tier 1 only (FM-D38):** Removed NEP, Fronting Fees, Ceding Commission, Net LLAE, Net Commission as separate lines. These are now sub-detail on Revenue Summary and Expense Summary (Tier 2).
2. **Symmetric structure (FM-D39):** Revenue has UW Revenue + Other Revenue → Total Revenue. Expenses has UW Expenses + Operating Expenses → Total Expenses. Parallel pattern.
3. **Total Expense row (FM-D39):** Added Total UW Expenses, Total Operating Expenses, Total Expenses subtotals.
4. **Below-the-line walk (FM-D40):** Operating Income → Interest/Debt Expense → Other Income/(Expense) → Pre-Tax Income → Tax → Net Income. Interest expense comes from Capital Activity (CAP_Q_IntExp for surplus notes and other debt).
5. **Key Ratios revised (FM-D41):** Removed UW-specific ratios (loss ratio, combined ratio — those belong on UW Exec Summary). Added revenue/expense composition ratios and profitability ratios.
6. **Growth section (FM-D41):** Total Revenue, Total Expense, and Net Income growth rates (quarter-over-quarter). First quarter shows "—" (no prior period).

---

## Impact on Other Tabs

### UW Executive Summary — No changes
UW-specific ratios (loss ratio, combined ratio, UW margin) stay here. This is the actuarial view. The IS is the management/investor view.

### Balance Sheet — Minor update
Investment assets now come from Investments tab (INV_Q_Invested, INV_Q_Liquid) instead of the old §10 computation embedded in Other Revenue Detail.

### Cash Flow Statement — Minor update
Interest expense (IS_INTEXP) flows into CFO. Investment asset changes flow into CFI from the Investments tab.

### Assumptions Tab — Slimmed
Investment-related parameters (yield, allocation %, liquidity floor, liquid yield) move to the Investments tab. Assumptions keeps: tax rates, inflation/growth rates, program registry, collection lag, risk-free rate.

---

## Updated Data Flow

```
Assumptions (global params)
    │
    ├──→ UW Inputs ──→ RunProjections ──→ UW Exec Summary ──┐
    │                                                         │
    ├──→ Investments (allocation, yields) ────────────────────┤
    │                                                         │
    ├──→ Capital Activity (raises, debt, interest) ──────────┤
    │                                                         │
    │         ┌───────────────────────────────────────────────┘
    │         │
    │         ├──→ Revenue Summary (UW Rev + Other Rev)
    │         │         │
    │         ├──→ Expense Summary (UW Exp + Operating Exp)
    │         │         │
    │         │         ▼
    │         ├──→ Income Statement (Tier 1: Rev - Exp - Interest - Tax = NI)
    │         │         │
    │         │         ▼
    │         ├──→ Balance Sheet (assets = liabilities + equity)
    │         │         │
    │         │         ▼
    │         └──→ Cash Flow Statement (derived from ΔBS)
```

---

## Updated Named Range Contracts

### Revenue Summary (revised)

| Range | ID | Description |
|---|---|---|
| REV_Q_UWRevenue | REV_UWREV | = NEP + Fronting Fees (earned). NO ceding commission. |
| REV_Q_OtherRevenue | REV_OTHREV | = Investment Income + Fee + Consulting |
| REV_Q_TotalRevenue | REV_TOTAL | = UW Revenue + Other Revenue |

### Expense Summary (revised)

| Range | ID | Description |
|---|---|---|
| EXP_Q_NLLAE | EXP_NLLAE | Net Losses & LAE |
| EXP_Q_NetAcq | EXP_NACQ | Net Acquisition Expense = Direct Comm - Ceding Comm |
| EXP_Q_UWExpense | EXP_UWEXP | = NLLAE + Net Acq + Other UW |
| EXP_Q_OpExpense | EXP_OPEXP | = Staffing + Other Operating |
| EXP_Q_TotalExpense | EXP_TOTAL | = UW Exp + Operating Exp |

### Income Statement (revised)

| Range | ID | Description |
|---|---|---|
| IS_Q_TotalRev | IS_TOTALREV | Total Revenue |
| IS_Q_TotalExp | IS_TOTALEXP | Total Expenses |
| IS_Q_OpIncome | IS_OPINC | Operating Income |
| IS_Q_IntExp | IS_INTEXP | Interest / Debt Expense |
| IS_Q_PreTax | IS_PRETAX | Pre-Tax Income |
| IS_Q_Tax | IS_TAX | Tax Expense |
| IS_Q_NetIncome | IS_NETINC | Net Income → BS Retained Earnings |

---

## Deferred Tabs (unchanged — build after v1 FM ships)

- UW Program Detail (per-program UW waterfall)
- UW KPI Summary (actuarial ratios, capital adequacy v2)
- Other Revenue Detail (fee income, consulting detail)
- Other Expense Detail (G&A categories with growth rates)
- Staffing Expense (by department, loaded cost)
- Sales Funnel (program pipeline → ProgramInputs)
- Valuation, Investor Returns, Cap Table (future)
