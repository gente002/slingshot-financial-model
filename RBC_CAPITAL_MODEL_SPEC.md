# RBC Capital Model — Design Specification & Build Prompt

**Date:** April 2026
**Status:** HISTORICAL — DO NOT USE AS CURRENT TRUTH
**Depends on:** CLEANUP sprint complete
**Scope:** One new formula tab (RBC Capital Model) via formula_tab_config.csv. No new VBA modules.

> ### ⚠ HISTORICAL DOC — SUPERSEDED
>
> **As of 2026-04-18**, this document describes the INITIAL design which has
> been significantly revised for NAIC compliance. Specifically, it shows R2 as
> "Insurance Risk" (old 5-component covariance) and R4 as a 3% "Business Risk"
> factor on GWP growth — both of those descriptions are OUTDATED.
>
> **Current authoritative sources for the RBC tab:**
> 1. `config/formula_tab_config.csv` rows tagged `"RBC Capital Model"` —
>    the actual formulas and layout.
> 2. `SESSION_NOTES.md` entries for 2026-04-18 under
>    *"RBC Capital Model: NAIC compliance corrections"* and
>    *"RBC Capital Model: Excessive Growth Charge + NWP annualization"*.
>
> **Current RBC calculation** uses the full NAIC 6-component covariance
> `R0 + sqrt(R1² + R2² + R3² + R4² + R5² + Rcat²)` with per-LOB NAIC factors,
> Loss/Premium Concentration Factor, and Excessive Growth Charge at NAIC
> multipliers (0.45 for reserves, 0.225 for premium). Validation against
> reference model matches within 0.0% at 2026 Q2, 2028 Q2, and 2030 Q4
> on the first-10-programs scope.
>
> **Deviations from full NAIC compliance (documented in SESSION_NOTES):**
> - R1 bond granularity: 6 buckets vs NAIC's 12+ classes
> - R3 reinsurance: legacy 10% rule vs 2018+ transaction-level model
> - Company-experience blending: not wired (defaults to industry factors)
> - Loss-sensitive contract discount: not yet implemented
> - LOB factors: currently from CAS 2018 illustrative data — verify against
>   current NAIC 2024+ forecasting instructions before regulatory filing

---


---

## Locked Decisions

| ID | Decision |
|---|---|
| RBC-01 | RBC only — defer AM Best and Internal Capital models |
| RBC-02 | Formula tab via formula_tab_config.csv — no new VBA |
| RBC-03 | Quarterly computation (QuarterlyColumns=TRUE, GrandTotal=FALSE) |
| RBC-04 | Roadmap: after CLEANUP, before Pipeline Full |

---

## Background

Risk-Based Capital (RBC) is the NAIC regulatory formula for P&C insurers. It computes minimum required capital based on risk profile. The RBC ratio (Total Adjusted Capital / Authorized Control Level) determines regulatory action thresholds.

**P&C RBC Components:**
- **R0** — Asset risk (affiliates) — typically 0 for a startup carrier
- **R1** — Fixed income risk (default risk by asset quality grade)
- **R2** — Insurance risk (reserve risk + premium/growth risk)
- **R3** — Interest rate risk (duration × rate factor)
- **R4** — Business risk (premium growth)
- **R5** — Credit risk (reinsurance recoverables × reinsurer quality factor)

**Covariance formula:** `ACL = R0 + √(R1² + R2² + R3² + R4² + R5²)`

**Regulatory thresholds:**
- ≥200% Company Action Level (CAL) — no action required
- 150-200% — company must file action plan
- 100-150% — regulator can take action
- 70-100% — regulator authorized to seize
- <70% — mandatory control

---

## Tab Layout

### Tab Registry Entry
```csv
"RBC Capital Model","Domain","Output","N","Visible","15.5","P&C Risk-Based Capital model with ratio trajectory","TRUE","Writing","FALSE","BF8F00","TRUE","FALSE"
```

SortOrder 15.5 places it between Balance Sheet (15) and Cash Flow Statement (16). QuarterlyColumns=TRUE. SkipAnnual=TRUE (RBC is point-in-time like Balance Sheet — annual total is meaningless for a ratio).

### formula_tab_config.csv Rows

```
Row 1:  [RBC Capital Model]                           Section header, 1F3864, white
Row 2:  "P&C Risk-Based Capital — Quarterly Trajectory"   Basis label

=== RISK CHARGES (Input section — blue cells, static column C) ===
Row 4:  Section "Risk Charge Factors"                  D9E1F2
Row 5:  RBC_R1_GOV      Govt Bond Risk Charge          Input, 0.003    (0.3% NAIC 1)
Row 6:  RBC_R1_CORP     Corp Bond Risk Charge          Input, 0.010    (1.0% NAIC 2)
Row 7:  RBC_R1_MUNI     Muni Bond Risk Charge          Input, 0.003    (0.3% NAIC 1)
Row 8:  RBC_R1_CASH     Cash Risk Charge               Input, 0.000    (0% — no default risk)
Row 9:  RBC_R1_EQ       Equity Risk Charge             Input, 0.150    (15% common stock)
Row 10: RBC_R1_ALT      Alternative Risk Charge        Input, 0.200    (20% other invested)
Row 11: RBC_R2_RSV      Reserve Risk Factor            Input, 0.150    (15% of net reserves)
Row 12: RBC_R2_PREM     Premium Risk Factor            Input, 0.100    (10% of NWP)
Row 13: RBC_R3_RATE     Interest Rate Risk Factor      Input, 0.020    (2% of invested × duration)
Row 14: RBC_R4_BIZ      Business Risk Factor           Input, 0.030    (3% of GWP growth)
Row 15: RBC_R5_CRED     RI Credit Risk Factor          Input, 0.100    (10% of ceded reserves)
Row 16: RBC_TARGET      Target RBC Ratio               Input, 3.000    (300% — management target)
Row 17: Spacer

=== QUARTERLY RBC COMPUTATION (formula section) ===
Row 18: Section "Risk Components"                       D9E1F2

Row 19: RBC_INVESTED     Total Invested Assets          = {REF:Investments!INV_Q_INVESTED}
Row 20: RBC_LIQUID       Liquid Cash                    = {REF:Investments!INV_Q_LIQUID}
Row 21: Spacer

Row 22: RBC_R0           R0: Affiliate Risk             = 0                                    (no affiliates for startup)
Row 23: RBC_R1           R1: Fixed Income Risk           = {REF:Investments!INV_Q_INVESTED}*$C$5*{REF:Investments!INV_GOV_ALLOC} + ... (see formula below)
Row 24: RBC_R2           R2: Insurance Risk              = {REF:UW Exec Summary!UWEX_NRSV}*$C$11 + {REF:UW Exec Summary!UWEX_NWP}*$C$12
Row 25: RBC_R3           R3: Interest Rate Risk          = {REF:Investments!INV_Q_INVESTED}*$C$13
Row 26: RBC_R4           R4: Business Risk               = MAX(0,{REF:UW Exec Summary!UWEX_GWP}-{PREV_Q:RBC_GWP_PRIOR})*$C$14
Row 27: RBC_R5           R5: Credit Risk                 = ABS({REF:UW Exec Summary!UWEX_CRSV})*$C$15
Row 28: Spacer

Row 29: Section "Capital Adequacy"                      D9E1F2
Row 30: RBC_ACL          Authorized Control Level        = {ROWID:RBC_R0}+SQRT({ROWID:RBC_R1}^2+{ROWID:RBC_R2}^2+{ROWID:RBC_R3}^2+{ROWID:RBC_R4}^2+{ROWID:RBC_R5}^2)
Row 31: RBC_TAC          Total Adjusted Capital          = {REF:Balance Sheet!BS_EQUITY}
Row 32: RBC_RATIO        RBC Ratio                       = IFERROR({ROWID:RBC_TAC}/{ROWID:RBC_ACL},0)     Format: 0.0%
Row 33: RBC_TARGET_LINE  Target Ratio                    = $C$16                                           Format: 0.0%, Italic, grey
Row 34: RBC_SURPLUS      Capital Surplus / (Deficit)     = {ROWID:RBC_TAC}-{ROWID:RBC_ACL}*{ROWID:RBC_TARGET_LINE}
Row 35: Spacer

Row 36: Section "Regulatory Thresholds"                 D9E1F2
Row 37: RBC_CAL          Company Action Level (200%)     = {ROWID:RBC_ACL}*2                    Format: #,##0
Row 38: RBC_RAL          Regulatory Action Level (150%)  = {ROWID:RBC_ACL}*1.5                  Format: #,##0
Row 39: RBC_AUTH         Authorized Control (100%)       = {ROWID:RBC_ACL}                      Format: #,##0
Row 40: RBC_MAND         Mandatory Control (70%)         = {ROWID:RBC_ACL}*0.7                  Format: #,##0
Row 41: Spacer

Row 42: Section "Key Ratios"                            Italic, grey
Row 43: RBC_RATIO_DISP   RBC Ratio                      = {ROWID:RBC_RATIO}                    0.0%
Row 44: RBC_R1_PCT       R1 as % of ACL                 = IFERROR({ROWID:RBC_R1}/{ROWID:RBC_ACL},0)    0.0%
Row 45: RBC_R2_PCT       R2 as % of ACL                 = IFERROR({ROWID:RBC_R2}/{ROWID:RBC_ACL},0)    0.0%
Row 46: RBC_R5_PCT       R5 as % of ACL                 = IFERROR({ROWID:RBC_R5}/{ROWID:RBC_ACL},0)    0.0%

=== SUPPORTING CALCULATIONS (for {PREV_Q:} deltas) ===
Row 48: Section "Supporting Calculations"               Italic, grey
Row 49: RBC_GWP_PRIOR    Prior Quarter GWP (for R4)     = {REF:UW Exec Summary!UWEX_GWP}      (echo row for {PREV_Q:} delta)
```

### R1 Formula Detail

R1 = sum of (invested assets × asset allocation % × risk charge) for each asset class:

```
={REF:Investments!INV_Q_POOL}*($C$5*$C$5_alloc + $C$6*$C$6_alloc + ... )
```

Actually, since allocation % and risk charges are both static inputs in column C, and INV_Q_POOL is the total investable pool:

```
=({REF:Investments!INV_Q_POOL})*SUMPRODUCT($C$5:$C$10, [allocation percentages from Investments tab])
```

This is tricky because the Investments allocations are on a different tab at fixed rows. The cleanest approach:

```
R1 = {REF:Investments!INV_Q_POOL} * (
    {REF:Investments!INV_GOV_ALLOC} * $C$5 +
    {REF:Investments!INV_CORP_ALLOC} * $C$6 +
    {REF:Investments!INV_MUNI_ALLOC} * $C$7 +
    {REF:Investments!INV_CASH_ALLOC} * $C$8 +
    {REF:Investments!INV_EQ_ALLOC} * $C$9 +
    {REF:Investments!INV_ALT_ALLOC} * $C$10
)
```

**Note:** INV_*_ALLOC are static inputs (column C only, not quarterly). The {REF:} placeholder resolves to the static cell. This formula computes the weighted-average risk charge and applies it to the quarterly pool.

### R2 Formula Detail

R2 has two sub-components:
- Reserve risk = |Net Loss Reserves| × reserve factor
- Premium risk = |Net Written Premium| × premium factor

```
R2 = ABS({REF:UW Exec Summary!UWEX_NRSV}) * $C$11 + ABS({REF:UW Exec Summary!UWEX_NWP}) * $C$12
```

Note: UWEX_NWP and UWEX_NRSV may be negative after PD-05 sign convention — use ABS().

Actually wait — UWEX_NWP is Net Written = Gross + Ceded (where Ceded is negative). So NWP is positive (smaller than GWP). UWEX_NRSV = Gross - Ceded = positive. These should already be positive. But use ABS() as a safety net.

### R4 Formula Detail

Business risk = premium GROWTH × factor. Growth = current GWP - prior quarter GWP. Only positive growth contributes risk.

```
R4 = MAX(0, {REF:UW Exec Summary!UWEX_GWP} - {PREV_Q:RBC_GWP_PRIOR}) * $C$14
```

The echo row RBC_GWP_PRIOR mirrors UWEX_GWP so {PREV_Q:} can compute the delta.

### R5 Formula Detail

Credit risk = ceded reserves × credit quality factor. UWEX_CRSV is the ceded loss reserve (negative after PD-05). Use ABS().

```
R5 = ABS({REF:UW Exec Summary!UWEX_CRSV}) * $C$15
```

### Capital Surplus / Deficit

Shows how much capital exceeds (or falls short of) the TARGET ratio × ACL:

```
Surplus = TAC - (ACL × Target Ratio)
```

Positive = surplus capital above target. Negative = deficit, need to raise capital.

---

## Conditional Formatting

Apply via ApplyHealthFormatting (or add to health_config if it exists):

**RBC_RATIO row:**
- ≥ 2.0 (200%) → Green fill (C6EFCE), green font
- 1.5 to 2.0 → Yellow fill (FFEB9C), dark yellow font
- 1.0 to 1.5 → Orange fill (FFC000), dark font
- < 1.0 → Red fill (FFC7CE), red font

**RBC_SURPLUS row:**
- ≥ 0 → Green font
- < 0 → Red font, bold

---

## Named Ranges

```csv
"RBC_Q_Ratio","RBC Capital Model","RBC_RATIO","","Quarterly","RBC ratio"
"RBC_Q_ACL","RBC Capital Model","RBC_ACL","","Quarterly","Authorized Control Level"
"RBC_Q_TAC","RBC Capital Model","RBC_TAC","","Quarterly","Total Adjusted Capital"
"RBC_Q_Surplus","RBC Capital Model","RBC_SURPLUS","","Quarterly","Capital surplus/deficit vs target"
```

---

## Investor Story

The RBC trajectory tells this story:
- Q1Y1: Equity raised → high RBC ratio (e.g., 500%)
- As programs ramp → reserves build → R2 increases → ACL increases → ratio declines
- Steady state: ratio stabilizes at 250-350% depending on premium volume and retention
- Capital Surplus shows whether additional equity raises are needed and when
- If surplus goes negative → model shows exactly which quarter the carrier needs more capital

---

## CC Build Prompt

Read `CLAUDE.md`, then `SESSION_NOTES.md`.

Add the RBC Capital Model tab to the insurance financial model:

1. Add tab to `config_insurance/tab_registry.csv` — SortOrder between Balance Sheet and Cash Flow Statement. QuarterlyColumns=TRUE, SkipAnnual=TRUE, GrandTotal=FALSE.

2. Add all rows to `config_insurance/formula_tab_config.csv` per the layout above. Use the exact RowIDs specified (RBC_R0 through RBC_GWP_PRIOR).

3. Add named ranges to `config_insurance/named_range_registry.csv` per the spec.

4. Add RBC_RATIO to health_config (or wire into ApplyHealthFormatting) with the 4-tier color thresholds.

5. Add RBC_SURPLUS conditional formatting (green positive, red negative).

6. Seed the risk charge inputs with the defaults specified (these are approximate NAIC factors — the user can adjust).

7. Add to regression_config.csv so workspace saves capture the RBC tab.

8. Append to SESSION_NOTES.md. Update CLAUDE.md tab count.

### Validation
- Run Model with 3 programs, $15M equity
- RBC tab shows quarterly ratios declining as premium ramps
- RBC Ratio > 0% (not div-by-zero)
- ACL > 0 when reserves and premium exist
- Capital Surplus = TAC - (ACL × Target)
- Regulatory threshold rows show correct multiples of ACL
- BS still balances, CFS still reconciles
- No #REF! errors
