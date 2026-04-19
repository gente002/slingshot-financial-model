# Phase 11: DomainEngine Computation Design

**Date:** 2026-03-26
**Status:** LOCKED — computation pipeline defined.
**Ported from:** UWMEngine v2.9.4 (ComputeEarning, ComputeUltimates, ComputeMonthly, UWMCsvIO.WriteCSV)

---

## Computation Pipeline

The insurance DomainEngine replaces SampleDomainEngine. It reads UW Inputs directly (CR-08), computes monthly actuarial projections, and writes Program×Month rows to the kernel outputs array. The kernel handles Net derivation and quarterly aggregation.

```
Initialize()  → Read UW Inputs tab, load curve params, register transforms
Validate()    → Check programs defined, ELR check, reinsurance validation
Reset()       → Clear internal arrays
Execute()     → Run 6-step computation, write to outputs array
```

### Execute() — 6-Step Pipeline

```
Step 1: ReadUWInputs          Read all 4 sections from UW Inputs tab
Step 2: SpreadPremium         Quarterly GWP → monthly WP, apply growth Y6-Y10
Step 3: EarnPremium           Monthly WP → monthly EP via term-based earning
Step 4: ComputeUltimates      EP × ELR → ultimate losses per exposure month
Step 5: DevelopLosses         CurveLib CDF → monthly emergence (Paid, Case, IBNR, counts)
Step 6: WriteOutputs          Assemble Gross + Ceded blocks, write to outputs array
```

---

## Step 1: ReadUWInputs

Reads the UW Inputs tab directly (not via InputValue). Populates module-level arrays:

```vba
' Program identity (Section 1)
Private m_progName(1 To 10) As String
Private m_progBU(1 To 10) As String
Private m_progTerm(1 To 10) As Long        ' policy term in months

' Premium schedule (Section 1) — 20 quarterly values
Private m_gwpQtr(1 To 10, 1 To 20) As Double
Private m_gwpGrowth(1 To 10) As Double     ' annual growth for Y6-Y10

' Commission rates (Section 2) — 5 annual rates
Private m_commRate(1 To 10, 1 To 5) As Double

' Loss assumptions (Section 3) — 3 loss types per program
Private m_lyrLOB(1 To 10, 1 To 3) As String       ' "Property" or "Casualty"
Private m_lyrLossTL(1 To 10, 1 To 3) As Long      ' curve trend level (1-100)
Private m_lyrCntTL(1 To 10, 1 To 3) As Long       ' count curve trend level
Private m_lyrELR(1 To 10, 1 To 3, 1 To 4) As Double  ' quarterly ELRs
Private m_lyrSev(1 To 10, 1 To 3) As Double       ' average severity
Private m_lyrFreq(1 To 10, 1 To 3, 1 To 4) As Double ' quarterly frequencies
Private m_lyrActive(1 To 10, 1 To 3) As Boolean   ' TRUE if ELR > 0

' Reinsurance terms (Section 4) — 5-year blocks
Private m_reinsCedePct(1 To 10, 1 To 5) As Double
Private m_reinsCedeComm(1 To 10, 1 To 5) As Double
Private m_reinsFrontFee(1 To 10, 1 To 5) As Double

' Program count (how many rows are populated)
Private m_numProgs As Long
```

**Attritional ELR rule:** For loss type 1 (Attritional), Q2-Q4 ELR = Q1 ELR. The user only enters Q1. DomainEngine copies Q1 to Q2-Q4 during ReadUWInputs.

---

## Step 2: SpreadPremium

Converts quarterly GWP to monthly WP. Applies growth for Y6-Y10.

```vba
Private m_wpMon(1 To 10, 1 To 120) As Double   ' monthly written premium

' For each program:
'   Quarters 1-20: m_wpMon(p, m) = m_gwpQtr(p, qIdx) / 3
'     where qIdx = Int((m-1)/3) + 1
'   Months 61-120 (Y6-Y10): apply compound growth from Y5 values
'     Y6 quarterly = Y5 quarterly × (1 + growth)
'     Y7 quarterly = Y6 quarterly × (1 + growth)
'     etc.
```

**Growth application:** Growth compounds annually from the Y5 premium base. Each year Y6-Y10, the quarterly premium = prior year same quarter × (1 + annual growth rate). Monthly = quarterly / 3.

---

## Step 3: EarnPremium

Earns each month's written premium across the policy term using mid-month assumption.

```vba
Private m_epMon(1 To 10, 1 To 120) As Double   ' monthly earned premium

' For each program p, for each write month m:
'   If m_wpMon(p, m) = 0, skip
'   term = m_progTerm(p)
'   Spread WP across earning window [m, m+term]:
'     First month (ew=m):  frac = 1 / (2 × term)     ← half-month assumption
'     Last month (ew=m+t): frac = 1 / (2 × term)     ← half-month assumption
'     Middle months:        frac = 1 / term            ← full month
'   m_epMon(p, ew) += m_wpMon(p, m) × frac
```

This is the UWM's mid-month earning assumption. Total earning fraction sums to exactly 1.0 over the term (the half-month at start + T-1 full months + half-month at end = T/T = 1.0).

**Commission earning:** Written commission = WP × commission rate. Earned commission follows the same earning pattern as premium — it earns as the premium earns.

**Fronting fee earning:** Written fronting fee = WP × fronting fee rate. Same earning pattern.

```vba
' Computed during Step 6 (WriteOutputs) from WP/EP and rates:
'   G_WComm  = WP(m) × commRate(year)
'   G_EComm  = EP(m) × commRate(year)
'   G_WFrontFee = WP(m) × frontFeeRate(year)
'   G_EFrontFee = EP(m) × frontFeeRate(year)
```

---

## Step 4: ComputeUltimates

For each exposure month with earned premium, compute ultimate losses and claim counts.

```vba
Private m_ultMon(1 To 10, 1 To 3, 1 To 120) As Double  ' ultimate loss per layer per exposure month
Private m_cntUlt(1 To 10, 1 To 3, 1 To 120) As Double  ' ultimate count per layer per exposure month

' For each program p, layer l (1=Attr, 2=Seas, 3=CAT):
'   For each exposure month ep with EP > 0:
'     qWithinYr = ((ep-1) Mod 12) \ 3 + 1     ← which quarter (1-4)
'     If Attritional: elr = m_lyrELR(p, 1, 1)  ← always Q1 (uniform)
'     Else:           elr = m_lyrELR(p, l, qWithinYr)
'
'     m_ultMon(p, l, ep) = m_epMon(p, ep) × elr
'     m_cntUlt(p, l, ep) = m_epMon(p, ep) × m_lyrFreq(p, l, qWithinYr) / 1,000,000
```

**Frequency convention:** Frequency is per $1M earned premium. UltCount = EP × Freq / 1,000,000.

---

## Step 5: DevelopLosses

Apply CurveLib CDF functions to develop losses from ultimate to monthly emerged amounts.

```vba
' Working arrays (cumulative by calendar month)
Private m_cumPaid(1 To 10, 1 To 120) As Double
Private m_cumCI(1 To 10, 1 To 120) As Double     ' cumulative case incurred
Private m_cumUlt(1 To 10, 1 To 120) As Double    ' cumulative ultimate
Private m_cumRpt(1 To 10, 1 To 120) As Double    ' cumulative reported count
Private m_cumCls(1 To 10, 1 To 120) As Double    ' cumulative closed count
Private m_cumCntUlt(1 To 10, 1 To 120) As Double ' cumulative ultimate count
Private m_cumEP(1 To 10, 1 To 120) As Double     ' cumulative earned premium

' For each program p, calendar month cm (1 to devEnd):
'   For each layer l:
'     For each exposure month ep (1 to cm):
'       age = cm - ep + 1           ← 1-based development age (B96)
'
'       ' Look up curve parameters from GetDefaultParams (via LOB + TrendLevel)
'       ' 4 curves: Paid, CaseIncurred, ReportedCount, ClosedCount
'       paidPct = EvaluateCurve(distPd, p1Pd, p2Pd, age, maxAgePd)
'       ciPct   = EvaluateCurve(distCI, p1CI, p2CI, age, maxAgeCI)
'       rptPct  = EvaluateCurve(distRC, p1RC, p2RC, age, maxAgeRC)
'       clsPct  = EvaluateCurve(distCC, p1CC, p2CC, age, maxAgeCC)
'
'       ' Accumulate into layer totals for this calendar month
'       layerPaid += m_ultMon(p,l,ep) × paidPct
'       layerCI   += m_ultMon(p,l,ep) × ciPct
'       layerRpt  += m_cntUlt(p,l,ep) × rptPct
'       layerCls  += m_cntUlt(p,l,ep) × clsPct
'
'   ' Sum across layers into program-level cumulative arrays
'   m_cumPaid(p, cm) = sum of layerPaid across all active layers
'   m_cumCI(p, cm)   = sum of layerCI
'   m_cumUlt(p, cm)  = sum of all ultMon across all layers and all ep
'   m_cumCntUlt(p, cm) = sum of all cntUlt
'   m_cumRpt(p, cm)  = sum of layerRpt
'   m_cumCls(p, cm)  = sum of layerCls
'   m_cumEP(p, cm)   = running sum of m_epMon(p, 1..cm)
```

**Curve parameter lookup:** Uses Ext_CurveLib.CalcCDF for the math, but the parameters (distribution type, p1, p2, maxAge) come from GetDefaultParams — which is insurance-specific and lives in the DomainEngine (not CurveLib). GetDefaultParams uses LOB × CurveType × TrendLevel to interpolate parameters from the UWM's hardcoded anchor tables. For v1, we port GetDefaultParams into the DomainEngine. For v2, these could move to curve_library_config.csv.

**Development endpoint:** Per-program — the first calendar month where ALL curves ≥ 99.9999% emerged. Beyond this point, no further development. Reduces computation for programs that mature early.

---

## Step 6: WriteOutputs

Converts cumulative arrays to MTD (incremental) values and writes to the kernel outputs array. Applies reinsurance.

```vba
' For each program p, calendar month cm (1 to projection horizon):
'   row = (p-1) × horizon + cm
'
'   ' --- Dimensions ---
'   outputs(row, ColIndex("EntityName")) = m_progName(p)
'   outputs(row, ColIndex("Period"))     = cm
'   outputs(row, ColIndex("Quarter"))    = ((cm-1) Mod 12) \ 3 + 1
'   outputs(row, ColIndex("Year"))       = ((cm-1) \ 12) + 1
'
'   ' --- Determine rate year for this calendar month ---
'   rateYr = Min(((cm-1) \ 12) + 1, 5)    ← caps at Y5 (constant after Y5)
'   commPct = m_commRate(p, rateYr)
'   cedePct = m_reinsCedePct(p, rateYr)
'   cedeCommPct = m_reinsCedeComm(p, rateYr)
'   frontFeePct = m_reinsFrontFee(p, rateYr)
'
'   ' --- Incremental (MTD) from cumulative ---
'   If cm = 1 Then
'     mtdPaid = m_cumPaid(p, 1)
'     mtdCI   = m_cumCI(p, 1)
'     ' ... etc for all cumulative fields
'   Else
'     mtdPaid = m_cumPaid(p, cm) - m_cumPaid(p, cm-1)
'     mtdCI   = m_cumCI(p, cm)  - m_cumCI(p, cm-1)
'     ' ...
'   End If
'
'   ' --- Gross Block ---
'   G_WP        = m_wpMon(p, cm)
'   G_EP        = m_epMon(p, cm)     ← NOT cumulative; this is monthly earning increment
'   G_WComm     = G_WP × commPct
'   G_EComm     = G_EP × commPct
'   G_WFrontFee = G_WP × frontFeePct
'   G_EFrontFee = G_EP × frontFeePct
'   G_Paid      = mtdPaid
'   G_CaseRsv   = m_cumCI(p,cm) - m_cumPaid(p,cm)       ← EOP balance (Balance type)
'   G_CaseInc   = mtdCI
'   G_IBNR      = m_cumUlt(p,cm) - m_cumCI(p,cm)        ← EOP balance (Balance type)
'   G_Unpaid    = m_cumUlt(p,cm) - m_cumPaid(p,cm)       ← EOP balance (Balance type)
'   G_Ult       = mtdUlt      ← incremental ultimate (= EP × blended ELR for this month)
'   G_ClsCt     = mtdCls
'   G_OpenCt    = m_cumRpt(p,cm) - m_cumCls(p,cm)        ← EOP balance (Balance type)
'   G_RptCt     = mtdRpt
'   G_UltCt     = mtdCntUlt
'
'   ' --- Ceded Block (QS proportional) ---
'   C_WP        = G_WP × cedePct
'   C_EP        = G_EP × cedePct
'   C_WComm     = C_WP × cedeCommPct    ← ceding commission on ceded WP
'   C_EComm     = C_EP × cedeCommPct
'   C_WFrontFee = 0                      ← DR-018: no ceded fronting fee
'   C_EFrontFee = 0
'   C_Paid      = G_Paid × cedePct
'   C_CaseRsv   = G_CaseRsv × cedePct   ← Balance type, proportional
'   C_CaseInc   = G_CaseInc × cedePct
'   C_IBNR      = G_IBNR × cedePct
'   C_Unpaid    = G_Unpaid × cedePct
'   C_Ult       = G_Ult × cedePct
'   ' C_ClsCt, C_OpenCt, C_RptCt, C_UltCt → Derived by kernel (= Gross per CR-03/AR-01)
'
'   ' --- Net Block → Derived by kernel (= Gross - Ceded per column_registry) ---
'
'   Write all Gross and Ceded Incremental fields to outputs(row, ColIndex(...))
```

**Key detail — Balance fields:** G_CaseRsv, G_IBNR, G_Unpaid, G_OpenCt are computed as EOP balances from cumulative arrays (not differenced like flow fields). They are stored as the Balance value for that month. The kernel's quarterly aggregation takes the last-month-of-quarter value (not sum) because BalanceType = "Balance" (CR-04).

**Key detail — EP is already incremental:** m_epMon(p, cm) is the earning increment for month cm (how much premium earned THIS month from all prior written months). It is NOT cumulative. So G_EP = m_epMon(p, cm) directly, no differencing needed. Same for G_WP = m_wpMon(p, cm).

---

## Arbiter Decisions

| ID | Decision | Answer |
|---|---|---|
| DE-01 | Earning assumption | Mid-month: first and last month of term earn ½ fraction, middle months earn full fraction. Matches UWM. |
| DE-02 | Development age indexing | 1-based: age = cm - ep + 1. At cm=ep (same month as exposure), age=1 and CDF(1)>0. Matches UWM B96 fix. |
| DE-03 | Curve parameter source | DomainEngine.GetDefaultParams (ported from UWMCurves) for v1. Hardcoded anchor tables by LOB × CurveType × TrendLevel. Future: move to curve_library_config.csv. |
| DE-04 | Growth application | Annual compound from Y5 base. Y6 = Y5 × (1+g), Y7 = Y6 × (1+g), etc. Applied per-quarter then split to monthly ÷3. |
| DE-05 | Fronting fee on ceded | Zero (DR-018). Fronting fee is charged by the fronting carrier on gross premium. Ceded premium to the reinsurer does not carry a fronting fee. |
| DE-06 | Commission computation | Written commission = WP × commRate. Earned commission = EP × commRate. Both Gross. Ceding commission = Ceded WP × cedeCommRate (per QS terms). |
| DE-07 | Development endpoint | Per-program: first cm where all 4 curves ≥ 99.9999%. Caps computation. If no curves configured (no severity), use max(maxAge) across all configured curves. |

---

## Module Structure

```
InsuranceDomainEngine.bas  (~800-1000 lines)
├── Public Sub Initialize()        — read inputs, load curves, register transforms
├── Public Function Validate() As Boolean  — program/ELR/reins validation
├── Public Sub Reset()             — clear all arrays
├── Public Sub Execute()           — 6-step pipeline, write to DomainOutputs
├── Private Sub ReadUWInputs()     — read 4 sections from UW Inputs tab
├── Private Sub SpreadPremium()    — quarterly → monthly WP, growth
├── Private Sub EarnPremium()      — WP → EP via term-based earning
├── Private Sub ComputeUltimates() — EP × ELR → ultimate losses
├── Private Sub DevelopLosses()    — CurveLib CDF → emerged amounts
├── Private Sub GetDefaultParams() — LOB/TL → distribution/params/maxAge
├── Private Sub WriteOutputs()     — assemble rows, write to outputs
└── Public Sub AggregateToQuarterly() — PostCompute transform (moved from SampleDomainEngine)
```

**Size estimate:** ~800-1000 lines. Well under 64KB. GetDefaultParams is the largest single function (~80 lines of Select Case).

**AggregateToQuarterly moves** from SampleDomainEngine to InsuranceDomainEngine. It's domain-specific (it knows which fields are Balance vs Flow, which ratios to recompute). The sample model keeps its own version in SampleDomainEngine.

---

## CurveLib Integration

The DomainEngine calls Ext_CurveLib functions directly:

```vba
' In DevelopLosses:
paidPct = Ext_CurveLib.EvaluateCurve(distPd, p1Pd, p2Pd, age, maxAgePd)

' In GetDefaultParams:
p1 = Ext_CurveLib.LogInterp3(trendLevel, anchor1, anchor50, anchor100)
maxAge = CLng(Ext_CurveLib.LinInterp3(trendLevel, age1, age50, age100))
```

CurveLib provides the math. The DomainEngine provides the parameters (which distribution, what anchor values, what maxAge). This is the kernel/domain split from the CurveLib design session.

---

## Validation Rules (Validate())

| Check | Severity | Description |
|---|---|---|
| No programs defined | BLOCK | At least 1 program must have a name on UW Inputs |
| Missing BU | WARN | BU blank — defaults to "Unassigned" |
| Term ≤ 0 | BLOCK | Policy term must be positive |
| ELR out of range | WARN | ELR > 2.0 or < 0 — likely input error |
| Freq×Sev check | WARN | Freq×Sev/1M should ≈ ELR within 0.5% (per UWM) |
| CedePct out of range | BLOCK | Must be 0-1 |
| CedeComm out of range | WARN | Must be 0-1 |
| FrontFee out of range | WARN | Must be 0-1 |
| Commission out of range | WARN | Must be 0-1 |
| No premium entered | WARN | All 20 quarters = 0 — program produces no output |
