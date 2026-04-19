# Phase 11: Column Registry Design

**Date:** 2026-03-26
**Status:** LOCKED — Q1-Q4 answered, registry defined.

---

## Locked Decisions

| ID | Decision | Answer |
|---|---|---|
| CR-01 | Metrics per block | 16: WP, EP, WComm, EComm, WFrontFee, EFrontFee, Paid, CaseRsv, CaseInc, IBNR, Unpaid, Ult, ClsCt, OpenCt, RptCt, UltCt |
| CR-02 | Detail grain | Program × Month (1,200 rows for 10 programs × 120 months) |
| CR-03 | Ceded counts | Derived (= Gross counts per AR-01 QS proportional) |
| CR-04 | Reserve fields | Balance type (EOP value, not summed for quarterly aggregation) |

---

## Column Registry (52 columns)

### Dimensions (4)

| Name | DetailCol | CsvCol | Block | FieldClass | DefaultView | Format | BalGrp | DerivationRule | BalanceType |
|---|---|---|---|---|---|---|---|---|---|
| EntityName | 1 | 0 | | Dimension | | | | | Flow |
| Period | 2 | 1 | | Dimension | | | | | Flow |
| Quarter | 3 | 2 | | Dimension | | | | | Flow |
| Year | 4 | 3 | | Dimension | | | | | Flow |

### Gross Block — Incremental (16)

DomainEngine computes these directly.

| Name | DetailCol | CsvCol | Block | FieldClass | DefaultView | Format | BalGrp | DerivationRule | BalanceType |
|---|---|---|---|---|---|---|---|---|---|
| G_WP | 5 | 4 | GROSS | Incremental | Incremental | #,##0 | Income | | Flow |
| G_EP | 6 | 5 | GROSS | Incremental | Incremental | #,##0 | Income | | Flow |
| G_WComm | 7 | 6 | GROSS | Incremental | Incremental | #,##0 | Income | | Flow |
| G_EComm | 8 | 7 | GROSS | Incremental | Incremental | #,##0 | Income | | Flow |
| G_WFrontFee | 9 | 8 | GROSS | Incremental | Incremental | #,##0 | Income | | Flow |
| G_EFrontFee | 10 | 9 | GROSS | Incremental | Incremental | #,##0 | Income | | Flow |
| G_Paid | 11 | 10 | GROSS | Incremental | Incremental | #,##0 | Loss | | Flow |
| G_CaseRsv | 12 | 11 | GROSS | Incremental | Incremental | #,##0 | Reserve | | Balance |
| G_CaseInc | 13 | 12 | GROSS | Incremental | Incremental | #,##0 | Loss | | Flow |
| G_IBNR | 14 | 13 | GROSS | Incremental | Incremental | #,##0 | Reserve | | Balance |
| G_Unpaid | 15 | 14 | GROSS | Incremental | Incremental | #,##0 | Reserve | | Balance |
| G_Ult | 16 | 15 | GROSS | Incremental | Incremental | #,##0 | Loss | | Flow |
| G_ClsCt | 17 | 16 | GROSS | Incremental | Incremental | #,##0 | Count | | Flow |
| G_OpenCt | 18 | 17 | GROSS | Incremental | Incremental | #,##0 | Count | | Balance |
| G_RptCt | 19 | 18 | GROSS | Incremental | Incremental | #,##0 | Count | | Flow |
| G_UltCt | 20 | 19 | GROSS | Incremental | Incremental | #,##0 | Count | | Flow |

### Ceded Block — Incremental (12 computed + 4 derived)

DomainEngine computes dollar fields (via QS cede percentage). Count fields are Derived from Gross.

| Name | DetailCol | CsvCol | Block | FieldClass | DefaultView | Format | BalGrp | DerivationRule | BalanceType |
|---|---|---|---|---|---|---|---|---|---|
| C_WP | 21 | 20 | CEDED | Incremental | Incremental | #,##0 | Income | | Flow |
| C_EP | 22 | 21 | CEDED | Incremental | Incremental | #,##0 | Income | | Flow |
| C_WComm | 23 | 22 | CEDED | Incremental | Incremental | #,##0 | Income | | Flow |
| C_EComm | 24 | 23 | CEDED | Incremental | Incremental | #,##0 | Income | | Flow |
| C_WFrontFee | 25 | 24 | CEDED | Incremental | Incremental | #,##0 | Income | | Flow |
| C_EFrontFee | 26 | 25 | CEDED | Incremental | Incremental | #,##0 | Income | | Flow |
| C_Paid | 27 | 26 | CEDED | Incremental | Incremental | #,##0 | Loss | | Flow |
| C_CaseRsv | 28 | 27 | CEDED | Incremental | Incremental | #,##0 | Reserve | | Balance |
| C_CaseInc | 29 | 28 | CEDED | Incremental | Incremental | #,##0 | Loss | | Flow |
| C_IBNR | 30 | 29 | CEDED | Incremental | Incremental | #,##0 | Reserve | | Balance |
| C_Unpaid | 31 | 30 | CEDED | Incremental | Incremental | #,##0 | Reserve | | Balance |
| C_Ult | 32 | 31 | CEDED | Incremental | Incremental | #,##0 | Loss | | Flow |
| C_ClsCt | 33 | 32 | CEDED | Derived | Incremental | #,##0 | Count | G_ClsCt | Flow |
| C_OpenCt | 34 | 33 | CEDED | Derived | Incremental | #,##0 | Count | G_OpenCt | Balance |
| C_RptCt | 35 | 34 | CEDED | Derived | Incremental | #,##0 | Count | G_RptCt | Flow |
| C_UltCt | 36 | 35 | CEDED | Derived | Incremental | #,##0 | Count | G_UltCt | Flow |

### Net Block — Derived (16)

Kernel computes all Net fields as Gross - Ceded.

| Name | DetailCol | CsvCol | Block | FieldClass | DefaultView | Format | BalGrp | DerivationRule | BalanceType |
|---|---|---|---|---|---|---|---|---|---|
| N_WP | 37 | 36 | NET | Derived | Incremental | #,##0 | Income | G_WP - C_WP | Flow |
| N_EP | 38 | 37 | NET | Derived | Incremental | #,##0 | Income | G_EP - C_EP | Flow |
| N_WComm | 39 | 38 | NET | Derived | Incremental | #,##0 | Income | G_WComm - C_WComm | Flow |
| N_EComm | 40 | 39 | NET | Derived | Incremental | #,##0 | Income | G_EComm - C_EComm | Flow |
| N_WFrontFee | 41 | 40 | NET | Derived | Incremental | #,##0 | Income | G_WFrontFee - C_WFrontFee | Flow |
| N_EFrontFee | 42 | 41 | NET | Derived | Incremental | #,##0 | Income | G_EFrontFee - C_EFrontFee | Flow |
| N_Paid | 43 | 42 | NET | Derived | Incremental | #,##0 | Loss | G_Paid - C_Paid | Flow |
| N_CaseRsv | 44 | 43 | NET | Derived | Incremental | #,##0 | Reserve | G_CaseRsv - C_CaseRsv | Balance |
| N_CaseInc | 45 | 44 | NET | Derived | Incremental | #,##0 | Loss | G_CaseInc - C_CaseInc | Flow |
| N_IBNR | 46 | 45 | NET | Derived | Incremental | #,##0 | Reserve | G_IBNR - C_IBNR | Balance |
| N_Unpaid | 47 | 46 | NET | Derived | Incremental | #,##0 | Reserve | G_Unpaid - C_Unpaid | Balance |
| N_Ult | 48 | 47 | NET | Derived | Incremental | #,##0 | Loss | G_Ult - C_Ult | Flow |
| N_ClsCt | 49 | 48 | NET | Derived | Incremental | #,##0 | Count | G_ClsCt - C_ClsCt | Flow |
| N_OpenCt | 50 | 49 | NET | Derived | Incremental | #,##0 | Count | G_OpenCt - C_OpenCt | Balance |
| N_RptCt | 51 | 50 | NET | Derived | Incremental | #,##0 | Count | G_RptCt - C_RptCt | Flow |
| N_UltCt | 52 | 51 | NET | Derived | Incremental | #,##0 | Count | G_UltCt - C_UltCt | Flow |

---

## Summary

| Category | Count | FieldClass | Computed By |
|---|---|---|---|
| Dimensions | 4 | Dimension | Kernel |
| Gross metrics | 16 | Incremental | DomainEngine |
| Ceded dollar metrics | 12 | Incremental | DomainEngine (QS cede %) |
| Ceded count metrics | 4 | Derived | Kernel (= Gross per AR-01) |
| Net metrics | 16 | Derived | Kernel (= Gross - Ceded) |
| **Total** | **52** | | |

### Balance Type Fields (EOP, not summed for quarterly aggregation)

These fields take last-month-of-quarter value during AggregateToQuarterly:
- G_CaseRsv, G_IBNR, G_Unpaid, G_OpenCt
- C_CaseRsv, C_IBNR, C_Unpaid, C_OpenCt
- N_CaseRsv, N_IBNR, N_Unpaid, N_OpenCt

All other fields are Flow type (summed for quarterly aggregation).

### What's NOT in this registry (handled elsewhere)

- **Severity metrics** (PaidSev, CaseSev, etc.): Derived on UW Exec Summary formula tab as Paid/ClsCt, CaseInc/RptCt, etc. Not stored on Detail.
- **Loss ratios, combined ratios**: Derived on UW Exec Summary formula tab. Not stored on Detail.
- **Earning ratio** (EP/WP): Derived on UW Exec Summary formula tab.
- **UPR** (Unearned Premium Reserve = WP - EP): Derived on formula tabs. Not stored on Detail.
- **Net Acquisition Expense** (G_EComm - C_EComm): Derived on Expense Summary formula tab per FM-D34.
- **Investment income, tax, interest expense**: Computed on Investments/Capital/IS formula tabs, not on Detail.

---

## Comparison to UWM v2.9.4

| | UWM v2.9.4 | RDK FM |
|---|---|---|
| Total Detail columns | 152 | 52 |
| Blocks | 6 (MTD×3 + ITD×3) | 3 (Gross + Ceded + Net) |
| Metrics per block | 24 | 16 |
| ITD handling | Explicit columns | Kernel display mode toggle |
| Net handling | Explicit columns | Kernel Derived (Gross - Ceded) |
| Severity handling | Explicit columns | Formula tab derivation |
| Row grain | Program × LossType × Month (3,600) | Program × Month (1,200) |
| Reduction | — | 66% fewer columns, 67% fewer rows |
