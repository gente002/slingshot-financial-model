# Phase 11: Assumptions Tab Design

**Date:** 2026-03-26
**Status:** LOCKED — Q9-Q11 answered.
**Tab Type:** input_schema (scalar parameters, kernel-generated)

---

## Locked Decisions

| ID | Decision | Answer |
|---|---|---|
| CR-09 | Assumptions tab mechanism | input_schema (scalar params, kernel-generated) |
| CR-10 | Sections | All 6: Model Identity, Tax, Economic, Collection/Payment, Operating Expense, Program Registry |
| CR-11 | Loss inflation | Removed — ELR inputs already embed the actuary's trend view. No separate loss inflation parameter. |

---

## input_schema.csv Entries

```csv
"Section","ParamName","Row","DataType","Default","Validation","Required","Tooltip"
"Model Identity","ModelName","4","Text","Phronex FM v1","NotBlank","Y","Model name displayed on reports and cover pages"
"Model Identity","ProjectionStart","5","Date","2026-01-01","IsDate","Y","First month of projection"
"Model Identity","ProjectionYears","6","Integer","5","Range(1,10)","Y","Number of projection years (TimeHorizon = Years x 12)"
"Model Identity","ScenarioName","7","Text","Base","NotBlank","Y","Active scenario label"
"Tax & Regulatory","FederalTaxRate","10","Pct","0.21","Range(0,0.5)","Y","Federal corporate income tax rate"
"Tax & Regulatory","StateTaxRate","11","Pct","0.00","Range(0,0.15)","N","State income tax rate (placeholder)"
"Tax & Regulatory","EffectiveTaxRate","12","Pct","0.21","Range(0,0.5)","Y","Combined effective tax rate (default = Federal + State)"
"Economic","RiskFreeRate","15","Pct","0.04","Range(0,0.2)","Y","Risk-free discount rate (Treasury)"
"Economic","GeneralInflation","16","Pct","0.025","Range(-0.05,0.2)","N","General inflation rate for expense growth"
"Collection & Payment","PremCollectionLag","19","Integer","45","Range(0,180)","Y","Average days to collect premium receivable"
"Collection & Payment","LossPaymentLag","20","Integer","30","Range(0,180)","Y","Average days from reserve recognition to loss payment"
"Collection & Payment","CommPaymentLag","21","Integer","30","Range(0,180)","Y","Average days to pay commission"
"Operating Expense","StaffingExpenseY1","24","Currency","0","Range(0,100000000)","N","Annual staffing expense Year 1 (placeholder until Staffing tab)"
"Operating Expense","OtherOpExpenseY1","25","Currency","0","Range(0,100000000)","N","Annual other operating expense Year 1 (placeholder)"
"Operating Expense","ExpenseGrowthRate","26","Pct","0.03","Range(-0.1,0.5)","N","Annual growth rate for operating expenses"
"Program Registry","MaxPrograms","29","Integer","10","Range(1,10)","Y","Maximum programs supported in this model"
```

---

## Tab Layout (rendered by KernelBootstrap.GenerateInputsTab)

```
Row 1:  Assumptions                          (tab header)
Row 2:  (blank)
Row 3:  === Model Identity ===               (section header)
Row 4:  ModelName          Text    Phronex FM v1
Row 5:  ProjectionStart    Date    2026-01-01
Row 6:  ProjectionYears    Integer 5
Row 7:  ScenarioName       Text    Base
Row 8:  (blank)
Row 9:  === Tax & Regulatory ===
Row 10: FederalTaxRate     Pct     21.0%
Row 11: StateTaxRate       Pct     0.0%
Row 12: EffectiveTaxRate   Pct     21.0%
Row 13: (blank)
Row 14: === Economic ===
Row 15: RiskFreeRate       Pct     4.0%
Row 16: GeneralInflation   Pct     2.5%
Row 17: (blank)
Row 18: === Collection & Payment ===
Row 19: PremCollectionLag  Integer 45
Row 20: LossPaymentLag     Integer 30
Row 21: CommPaymentLag     Integer 30
Row 22: (blank)
Row 23: === Operating Expense ===
Row 24: StaffingExpenseY1  Currency $0
Row 25: OtherOpExpenseY1   Currency $0
Row 26: ExpenseGrowthRate  Pct     3.0%
Row 27: (blank)
Row 28: === Program Registry ===
Row 29: MaxPrograms        Integer 10
```

**Note:** Unlike UW Inputs, the Assumptions tab does NOT have per-entity columns. These are global parameters — one value each. The kernel's GenerateInputsTab renders them in a single-column layout (Col A = param name, Col B = data type, Col C = value).

---

## How Other Tabs Read Assumptions

| Consumer | Parameters Used | Mechanism |
|---|---|---|
| DomainEngine | ProjectionStart, ProjectionYears, MaxPrograms | KernelConfig.InputValue() |
| Income Statement | EffectiveTaxRate | Named range CTRL_TaxRate → formula |
| Balance Sheet | PremCollectionLag, LossPaymentLag, CommPaymentLag | Named ranges → AR/AP formulas |
| Expense Summary | StaffingExpenseY1, OtherOpExpenseY1, ExpenseGrowthRate | Named ranges → quarterly expense formulas |
| Investments | RiskFreeRate | Named range (floor rate reference) |
| KernelEngine | ProjectionYears → TimeHorizon | granularity_config override |

---

## Named Ranges Created

All Assumptions parameters get Single-type named ranges in named_range_registry.csv:

| RangeName | Cell | Description |
|---|---|---|
| CTRL_ModelName | C4 | Model name |
| CTRL_ProjStart | C5 | Projection start date |
| CTRL_ProjYears | C6 | Projection years |
| CTRL_Scenario | C7 | Scenario name |
| CTRL_FedTaxRate | C10 | Federal tax rate |
| CTRL_StateTaxRate | C11 | State tax rate |
| CTRL_TaxRate | C12 | Effective tax rate |
| CTRL_RiskFreeRate | C15 | Risk-free rate |
| CTRL_Inflation | C16 | General inflation |
| CTRL_PremCollLag | C19 | Premium collection lag (days) |
| CTRL_LossPayLag | C20 | Loss payment lag (days) |
| CTRL_CommPayLag | C21 | Commission payment lag (days) |
| CTRL_StaffExpY1 | C24 | Staffing expense Y1 |
| CTRL_OtherExpY1 | C25 | Other operating expense Y1 |
| CTRL_ExpGrowth | C26 | Expense growth rate |
| CTRL_MaxPrograms | C29 | Maximum programs |
