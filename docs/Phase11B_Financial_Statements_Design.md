# Phase 11B: Financial Statement Tabs Design

**Date:** 2026-03-26
**Status:** LOCKED — Q12-Q17 answered, all 7 tabs defined.
**Depends on:** Phase 11A (InsuranceDomainEngine, UW Exec Summary, named ranges)

---

## Locked Decisions

| ID | Decision | Answer |
|---|---|---|
| FS-01 | Investment income method | BOQ invested assets × yield ÷ 4. No circularity, no iterative calc. One-quarter reinvestment lag (immaterial). |
| FS-02 | Circularity resolution | Avoided by design (G-05 resolved). BOQ balance drives income; income flows into next quarter's BOQ. |
| FS-03 | Capital Activity scope | Full: equity raises, surplus notes, other debt — each with timing, draws, repayments, interest |
| FS-04 | Premium collection | Lag from Assumptions (PremCollectionLag days → quarterly fraction = lag/90). AR balance = uncollected WP. |
| FS-05 | Loss payment timing | Paid losses from Detail are already cash outflows. No additional lag. |
| FS-06 | Equity section | RE = Prior RE + Net Income. No dividends for v1. Total Equity = Paid-in Capital + RE. |

---

## Tab 4: Investments (Hybrid — formula_tab_config with CellType=Input)

### Layout

```
Row 1:  [Investment Portfolio]                    (navy header)
Row 2:  Management Projections                     (basis label)

=== ASSET ALLOCATION (static inputs) ===
Row 5:  Headers: Asset Class | Allocation% | Yield | Duration | Quality
Row 6:  ID: INV_GOV    Government/Agency Bonds    [Input: 40%, 3.5%, 4.0, AAA]
Row 7:  ID: INV_CORP   Corporate Bonds             [Input: 25%, 4.5%, 3.5, A]
Row 8:  ID: INV_MUNI   Municipal Bonds             [Input: 5%, 3.0%, 5.0, AA]
Row 9:  ID: INV_CASH   Money Market / Cash          [Input: 20%, 2.0%, 0.1, AAA]
Row 10: ID: INV_EQ     Equities                    [Input: 5%, 6.0%, —, —]
Row 11: ID: INV_ALT    Alternatives                [Input: 5%, 7.0%, —, —]
Row 12: (blank)
Row 13: ID: INV_TOTAL  Total / Weighted Average     [Formula: sums/SUMPRODUCT]
Row 14: ID: INV_CHECK  Allocation Check             [Formula: =SUM must = 100%]

=== PORTFOLIO SUMMARY (derived) ===
Row 17: ID: INV_WTD_YIELD    Weighted-Average Yield      [Formula: SUMPRODUCT]
Row 18: ID: INV_WTD_DUR      Weighted-Average Duration   [Formula: SUMPRODUCT]
Row 19: ID: INV_LIQUID_PCT   Liquid Allocation           [Formula: = INV_CASH alloc%]
Row 20: ID: INV_INVEST_PCT   Invested Allocation         [Formula: = 1 - liquid%]
Row 21: ID: INV_FLOOR         Liquidity Floor             [Input: $500,000]
Row 22: ID: INV_LIQUID_YIELD  Liquid Cash Yield           [Formula: = INV_CASH yield]
Row 23: ID: INV_INVEST_YIELD  Invested Asset Yield        [Formula: wtd avg of non-cash]

=== QUARTERLY INVESTMENT INCOME (formula-driven, QuarterlyColumns=TRUE) ===
Row 26: ID: INV_Q_POOL       Investable Pool              = UWEX_Q_GRSV + UWEX_Q_GUPR + CAP_EQUITY_CUMUL
Row 27: ID: INV_Q_INVESTED   Invested Assets              = MAX(Pool × InvestedPct, Pool - Floor)
Row 28: ID: INV_Q_LIQUID     Liquid Cash                  = Pool - Invested
Row 29: ID: INV_Q_INC_INV    Income on Invested           = BOQ_Invested × InvestYield ÷ 4
Row 30: ID: INV_Q_INC_LIQ    Income on Liquid             = BOQ_Liquid × LiquidYield ÷ 4
Row 31: ID: INV_Q_INC_TOTAL  Total Investment Income      = Invested + Liquid income
```

**BOQ convention:** For Q1Y1, BOQ = initial capital allocation (from Capital Activity initial equity). For subsequent quarters, BOQ = prior quarter's EOQ value. No circularity.

### Named Ranges
- INV_WtdYield (Single), INV_LiquidYield (Single), INV_InvestYield (Single)
- INV_Q_IncTotal (Quarterly), INV_Q_Invested (Quarterly), INV_Q_Liquid (Quarterly)
- INV_Q_Pool (Quarterly)

---

## Tab 5: Capital Activity (Hybrid — formula_tab_config with CellType=Input)

### Layout

Matches FM spec §7 exactly. Quarterly columns.

```
Row 1:  [Capital Activity]                        (navy header)
Row 2:  Management Projections                     (basis label)

=== SECTION 1: EQUITY RAISES ===
Row 5:  ID: CAP_EQ_DESC      Description           [Input: text]
Row 6:  ID: CAP_EQ_AMT       Amount Raised          [Input: quarterly, blue — $0 default]
Row 7:  ID: CAP_EQ_CUMUL     Cumulative Equity      [Formula: initial + running SUM]

=== SECTION 2: SURPLUS NOTES ===
Row 10: ID: CAP_SN_DESC      Description           [Input: text]
Row 11: ID: CAP_SN_DRAW      Draw Amount            [Input: quarterly, blue]
Row 12: ID: CAP_SN_REPAY     Repayment              [Input: quarterly, blue]
Row 13: ID: CAP_SN_BAL       Outstanding Balance    [Formula: prior + draw - repay]
Row 14: ID: CAP_SN_RATE      Interest Rate          [Input: annual rate, blue]
Row 15: ID: CAP_SN_INT       Interest Expense       [Formula: BOQ_balance × rate ÷ 4]

=== SECTION 3: OTHER DEBT ===
Row 18: ID: CAP_OD_DESC      Description           [Input: text]
Row 19: ID: CAP_OD_DRAW      Draw Amount            [Input: quarterly, blue]
Row 20: ID: CAP_OD_REPAY     Repayment              [Input: quarterly, blue]
Row 21: ID: CAP_OD_BAL       Outstanding Balance    [Formula: prior + draw - repay]
Row 22: ID: CAP_OD_RATE      Interest Rate          [Input: annual rate, blue]
Row 23: ID: CAP_OD_INT       Interest Expense       [Formula: BOQ_balance × rate ÷ 4]

=== SECTION 4: SUMMARY ===
Row 26: ID: CAP_TOTAL_RAISED  Total Capital Raised   [Formula: cumul equity + cumul draws]
Row 27: ID: CAP_TOTAL_DEBT    Total Debt Outstanding [Formula: SN_BAL + OD_BAL]
Row 28: ID: CAP_TOTAL_INT     Total Interest Expense [Formula: SN_INT + OD_INT]
Row 29: ID: CAP_NET_FIN       Net Financing CF        [Formula: raises + draws - repays]
```

**Interest uses BOQ balance** — same convention as investments. No circularity.

### Named Ranges
- CAP_Q_EquityRaised, CAP_Q_EquityCumul (Quarterly)
- CAP_Q_SurplusBal, CAP_Q_DebtBal (Quarterly — Balance type)
- CAP_Q_IntExp (Quarterly)
- CAP_Q_NetFinancing (Quarterly)

---

## Tab 6: Revenue Summary (Formula tab — QuarterlyColumns=TRUE)

Per FM_TAB_SPEC_UPDATE.md with FM-D34 net presentation.

```
=== UNDERWRITING REVENUE ===
  ID: REV_NEP       Net Earned Premium            = UWEX_Q_NEP
  ID: REV_FFE       Fronting Fees (earned)         = UWEX_Q_GFFE
  ID: REV_UWREV     Total UW Revenue              = NEP + FFE

=== OTHER REVENUE ===
  ID: REV_INVEST    Investment Income              = INV_Q_IncTotal
  ID: REV_FEE       Fee Income                     = 0 (placeholder)
  ID: REV_CONSULT   Consulting Revenue             = 0 (placeholder)
  ID: REV_OTHREV    Total Other Revenue            = Invest + Fee + Consult

=== TOTAL ===
  ID: REV_TOTAL     Total Revenue                 = UW Revenue + Other Revenue
```

**No ceding commission in revenue** (FM-D34 revised). Ceding commission nets against commissions in Expense Summary.

### Named Ranges
- REV_Q_UWRevenue, REV_Q_OtherRevenue, REV_Q_TotalRevenue (Quarterly)

---

## Tab 7: Expense Summary (Formula tab — QuarterlyColumns=TRUE)

Per FM_TAB_SPEC_UPDATE.md with FM-D34 net presentation.

```
=== UNDERWRITING EXPENSES ===
  ID: EXP_NLLAE     Net Losses & LAE              = UWEX_Q_NLLAE
  ID: EXP_NACQ      Net Acquisition Expense        = UWEX_Q_GCOMM - UWEX_Q_CEDCOMM
  ID: EXP_OTHUW     Other UW Expense              = 0 (placeholder)
  ID: EXP_UWEXP     Total UW Expenses             = NLLAE + Net Acq + Other UW

=== OPERATING EXPENSES ===
  ID: EXP_STAFF     Staffing Expense              = CTRL_StaffExpY1 × (1 + CTRL_ExpGrowth)^(year-1) ÷ 4
  ID: EXP_OTHER     Other Operating Expense       = CTRL_OtherExpY1 × (1 + CTRL_ExpGrowth)^(year-1) ÷ 4
  ID: EXP_OPEXP     Total Operating Expenses      = Staff + Other

=== TOTAL ===
  ID: EXP_TOTAL     Total Expenses                = UW Expenses + Operating Expenses
```

### Named Ranges
- EXP_Q_NLLAE, EXP_Q_NetAcq, EXP_Q_UWExpense, EXP_Q_OpExpense, EXP_Q_TotalExpense (Quarterly)

---

## Tab 8: Income Statement (Formula tab — QuarterlyColumns=TRUE)

Per FM_TAB_SPEC_UPDATE.md — Tier 1 only, symmetric structure, full walk.

```
=== REVENUE ===
  ID: IS_UWREV      Underwriting Revenue          = REV_Q_UWRevenue
  ID: IS_OTHREV     Other Revenue                 = REV_Q_OtherRevenue
  ID: IS_TOTALREV   Total Revenue                 = UW Rev + Other Rev         [bold, top border]

=== EXPENSES ===
  ID: IS_UWEXP      Underwriting Expenses         = EXP_Q_UWExpense
  ID: IS_OPEXP      Operating Expenses            = EXP_Q_OpExpense
  ID: IS_TOTALEXP   Total Expenses                = UW Exp + Op Exp            [bold, top border]

=== OPERATING INCOME ===
  ID: IS_OPINC      Operating Income              = Total Rev - Total Exp       [bold, double border]

=== BELOW THE LINE ===
  ID: IS_INTEXP     Interest / Debt Expense       = CAP_Q_IntExp
  ID: IS_OTHINC     Other Income / (Expense)      = 0 [Input, blue — manual override]

=== PRE-TAX ===
  ID: IS_PRETAX     Pre-Tax Income                = OpInc - IntExp + OthInc     [bold, top border]

=== TAX ===
  ID: IS_TAX        Income Tax Expense            = MAX(0, PreTax × CTRL_TaxRate)

=== NET INCOME ===
  ID: IS_NETINC     Net Income                    = PreTax - Tax                [bold, double border]

=== KEY RATIOS — REVENUE COMPOSITION ===
  ID: IS_KR_NEP     NEP : Total Revenue           = UWEX_Q_NEP / IS_TOTALREV
  ID: IS_KR_INV     Investment Inc : Total Rev     = INV_Q_IncTotal / IS_TOTALREV
  ID: IS_KR_OREV    Other Rev : Total Rev          = (IS_OTHREV - INV_Q_IncTotal) / IS_TOTALREV

=== KEY RATIOS — EXPENSE COMPOSITION ===
  ID: IS_KR_LLAE    Net L&LAE : Total Exp          = EXP_Q_NLLAE / IS_TOTALEXP
  ID: IS_KR_ACQ     Net Acq : Total Exp            = EXP_Q_NetAcq / IS_TOTALEXP
  ID: IS_KR_STAFF   Staffing : Total Exp           = EXP_Q_Staff / IS_TOTALEXP
  ID: IS_KR_OEXP    Other Exp : Total Exp          = EXP_Q_Other / IS_TOTALEXP

=== KEY RATIOS — PROFITABILITY ===
  ID: IS_KR_DEBT    Interest : Operating Inc       = IS_INTEXP / IS_OPINC
  ID: IS_KR_TAX     Tax : Pre-Tax Inc              = IS_TAX / IS_PRETAX

=== GROWTH ===
  ID: IS_GR_REV     Total Revenue Growth           = (Rev_t - Rev_{t-1}) / Rev_{t-1}
  ID: IS_GR_EXP     Total Expense Growth           = (Exp_t - Exp_{t-1}) / Exp_{t-1}
  ID: IS_GR_NI      Net Income Growth              = (NI_t - NI_{t-1}) / NI_{t-1}
```

### Named Ranges
- IS_Q_TotalRev, IS_Q_TotalExp, IS_Q_OpIncome, IS_Q_IntExp, IS_Q_PreTax, IS_Q_Tax, IS_Q_NetIncome (Quarterly)

---

## Tab 9: Balance Sheet (Formula tab — QuarterlyColumns=TRUE)

All values are EOP balances.

```
=== ASSETS ===
  ID: BS_CASH       Cash & Liquid Assets           = INV_Q_Liquid
  ID: BS_INVEST     Invested Assets                = INV_Q_Invested
  ID: BS_AR         Premium Receivable             = GWP × (PremCollLag / 90)
                    (= current quarter GWP × fraction uncollected at quarter-end)
  ID: BS_RI_RECV    Reinsurance Recoverable        = UWEX_Q_CRSV  (ceded loss reserve)
  ID: BS_OTHER_A    Other Assets                   = 0 (placeholder)
  ID: BS_TOTAL_A    Total Assets                   [bold, double border]

=== LIABILITIES ===
  ID: BS_LOSS_RSV   Gross Loss Reserve             = UWEX_Q_GRSV
  ID: BS_UPR        Unearned Premium Reserve       = UWEX_Q_GUPR
  ID: BS_AP         Accounts Payable               = 0 (placeholder — commission/expense payables)
  ID: BS_DEBT       Total Debt                     = CAP_Q_SurplusBal + CAP_Q_DebtBal
  ID: BS_TAX_PAY    Tax Payable                    = IS_Q_Tax (current quarter — simplified)
  ID: BS_OTHER_L    Other Liabilities              = 0 (placeholder)
  ID: BS_TOTAL_L    Total Liabilities              [bold, top border]

=== EQUITY ===
  ID: BS_PAIDIN     Paid-in Capital                = CAP_Q_EquityCumul
  ID: BS_RE         Retained Earnings              = Prior RE + IS_Q_NetIncome
  ID: BS_EQUITY     Total Equity                   = Paid-in + RE              [bold, top border]

=== BALANCE CHECK ===
  ID: BS_TOTAL_LE   Total Liabilities + Equity     [bold, double border]
  ID: BS_CHECK      Balance Check (A - L&E)        = Total Assets - Total L&E  [must = 0]
```

**Premium Receivable:** BS_AR = current quarter Gross WP × (PremCollectionLag / 90). At 45-day lag, ~50% of the quarter's WP is uncollected at quarter-end. This is a simplified calculation — a full AR rollforward would track each month's WP and collection timing, but for v1 this fraction approach is sufficient.

**Reinsurance Recoverable:** = Ceded loss reserves. The reinsurer owes us this amount for losses we've reserved but they haven't paid yet.

**Retained Earnings accumulation:** Q1Y1 RE = Net Income Q1. Q2Y1 RE = Q1 RE + Q2 NI. Running sum. No dividends (FS-06).

### Named Ranges
- BS_Q_TotalAssets, BS_Q_TotalLiab, BS_Q_Equity, BS_Q_Cash, BS_Q_Invest (Quarterly)
- BS_Q_LossRsv, BS_Q_UPR, BS_Q_Debt, BS_Q_RE (Quarterly)

---

## Tab 10: Cash Flow Statement (Formula tab — QuarterlyColumns=TRUE)

Derived from changes in BS balances plus IS flows.

```
=== CASH FROM OPERATIONS (CFO) ===
  ID: CFS_NI        Net Income                     = IS_Q_NetIncome
  ID: CFS_D_RSV     Change in Loss Reserves        = ΔBS_LOSS_RSV (increase = cash outflow negative)
  ID: CFS_D_UPR     Change in UPR                  = ΔUWEX_Q_GUPR (increase = cash inflow)
  ID: CFS_D_AR      Change in Premium Receivable   = ΔBS_AR (increase = cash outflow)
  ID: CFS_D_RI      Change in RI Recoverable       = ΔBS_RI_RECV (increase = cash outflow)
  ID: CFS_D_AP      Change in Payables             = ΔBS_AP
  ID: CFS_D_TAX     Change in Tax Payable          = ΔBS_TAX_PAY
  ID: CFS_CFO       Cash from Operations           = NI + working capital changes  [bold, top border]

=== CASH FROM INVESTING (CFI) ===
  ID: CFS_D_INV     Change in Invested Assets      = ΔBS_INVEST (increase = cash outflow)
  ID: CFS_CFI       Cash from Investing            = -ΔInvested                     [bold, top border]

=== CASH FROM FINANCING (CFF) ===
  ID: CFS_EQ_RAISE  Equity Raised                  = CAP_Q_EquityRaised
  ID: CFS_DEBT_NET  Net Debt Activity              = (SN_Draw + OD_Draw) - (SN_Repay + OD_Repay)
  ID: CFS_INT_PAID  Interest Paid                  = -CAP_Q_IntExp
  ID: CFS_CFF       Cash from Financing            = Equity + Net Debt - Interest   [bold, top border]

=== NET CHANGE ===
  ID: CFS_NET       Net Change in Cash             = CFO + CFI + CFF               [bold, double border]
  ID: CFS_BOQ_CASH  Beginning Cash                 = prior quarter BS_CASH + BS_INVEST
  ID: CFS_EOQ_CASH  Ending Cash                    = Beginning + Net Change
  ID: CFS_CHECK     Cash Reconciliation            = EOQ_CASH - BS_TOTAL_A cash portion [must = 0]
```

**Δ convention:** Change = EOQ - BOQ. Positive Δ in an asset = cash used. Positive Δ in a liability = cash received. The signs in CFO are: +NI, +ΔUPR (liability increase = cash in), -ΔReserves (we net paid vs reserve change), -ΔAR, -ΔRI.

**Note on reserves in CFO:** The Change in Loss Reserves line is the net reserve movement. Paid losses are already deducted in NI (through NLLAE on the IS). The reserve change adjusts for the non-cash portion of loss expense (IBNR increase, case reserve development). This is the standard indirect method.

### Named Ranges
- CFS_Q_CFO, CFS_Q_CFI, CFS_Q_CFF, CFS_Q_NetChange (Quarterly)

---

## Data Flow Summary

```
UW Exec Summary ──→ Revenue Summary ──→ Income Statement
       │          ──→ Expense Summary ──→      │
       │                                       │
       ├──→ Balance Sheet ←── Capital Activity  │
       │         │        ←── Investments       │
       │         │        ←── Income Statement ─┘
       │         │
       │         └──→ Cash Flow Statement ←── Capital Activity
       │
Assumptions ──→ (tax rate, lags, expense growth) → all formula tabs
```
