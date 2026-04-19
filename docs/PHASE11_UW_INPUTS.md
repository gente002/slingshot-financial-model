# Phase 11: UW Inputs Tab Design

**Date:** 2026-03-26
**Status:** LOCKED — Q5-Q8 answered, layout defined.
**Tab Type:** formula_tab_config (hybrid — CellType=Input for user-editable cells, CellType=Formula for check formulas)
**NOT generated from input_schema.** DomainEngine reads this tab directly.

---

## Locked Decisions

| ID | Decision | Answer |
|---|---|---|
| CR-05 | Loss type handling | Keep 3 loss types at input. DomainEngine blends internally before writing Program×Month Detail. |
| CR-06 | Premium schedule | 20 quarterly GWP values (Q1Y1–Q4Y5) + annual growth rate for Y6–Y10 |
| CR-07 | Reinsurance terms | Annual rates for 5 years, constant after Y5 |
| CR-08 | Input tab mechanism | Hybrid tab via formula_tab_config (not input_schema). DomainEngine reads directly. |

---

## Tab Layout

The UW Inputs tab has 4 sections stacked vertically, each with per-program rows. Programs are rows; time/parameters are columns. Maximum 10 programs.

### Section 1: Program Identity & Written Premium (rows 1–16)

```
Row 1:  [Underwriting Inputs]                              (navy header)
Row 2:  Enter program details and premium schedule          (italic, grey)
Row 3:  (blank)
Row 4:  [Section 1: Program Definitions & Written Premium]  (section header)
Row 5:  Headers:  A:# | B:BU | C:Program | D:Term(mo) | E:Q1Y1 | F:Q2Y1 | ... | X:Q4Y5 | Y:Growth%
Row 6:  Program 1  [input cells, blue]
Row 7:  Program 2
...
Row 15: Program 10
Row 16: (blank separator)
```

**Columns:**
| Col | Header | DataType | Description |
|---|---|---|---|
| A | # | Label | Program number (1-10, auto) |
| B | BU | Input/Text | Business unit code (e.g., "Property") |
| C | Program | Input/Text | Program name (e.g., "VISR") — becomes EntityName in Detail |
| D | Term | Input/Integer | Policy term in months (12 or 36 typical) |
| E-X | Q1Y1–Q4Y5 | Input/Currency | Gross Written Premium per quarter ($) |
| Y | Growth% | Input/Pct | Annual growth rate applied to Y6-Y10 (compounds quarterly) |

**Named Ranges:**
- UWIN_ProgName_{N}: Cell C{row} for each program
- UWIN_BU_{N}: Cell B{row}
- UWIN_Term_{N}: Cell D{row}
- UWIN_GWP_{N}: Range E{row}:X{row} (20 quarterly values)
- UWIN_Growth_{N}: Cell Y{row}

### Section 2: Direct Commission Rates (rows 17–30)

```
Row 17: [Section 2: Direct Commission Rates]               (section header)
Row 18: Headers:  A:# | B:(blank) | C:Program | D:(blank) | E:Y1 | F:Y2 | G:Y3 | H:Y4 | I:Y5
Row 19: Program 1  [input cells, blue — commission rate as %]
...
Row 28: Program 10
Row 29: (blank separator)
```

**Columns:**
| Col | Header | DataType | Description |
|---|---|---|---|
| A | # | Label | Program number |
| C | Program | Formula | =Section 1 program name (linked, not re-entered) |
| E-I | Y1–Y5 | Input/Pct | Direct commission rate per year (e.g., 25.0%) |

**Named Ranges:**
- UWIN_CommRate_{N}: Range E{row}:I{row} (5 annual values)

### Section 3: Loss Assumptions (rows 30–68)

This is the most complex section. Each program has 3 sub-rows (Attritional, Seasonal, Catastrophe). Total: 10 programs × 3 loss types = 30 data rows.

```
Row 30: [Section 3: Loss Assumptions]                       (section header)
Row 31: Headers:  A:# | B:Type | C:Program | D:LOB | E:LossTL | F:CntTL | G:Q1 ELR | H:Q2 ELR | I:Q3 ELR | J:Q4 ELR | K:Severity | L:Q1 Freq | M:Q2 Freq | N:Q3 Freq | O:Q4 Freq | P:Check
Row 32: Program 1 — Attritional  [input cells, blue]
Row 33: Program 1 — Seasonal
Row 34: Program 1 — Catastrophe
Row 35: Program 2 — Attritional
Row 36: Program 2 — Seasonal
Row 37: Program 2 — Catastrophe
...
Row 61: Program 10 — Catastrophe
Row 62: (blank separator)
```

**Columns:**
| Col | Header | DataType | Description |
|---|---|---|---|
| A | # | Label | Program number |
| B | Type | Label | "Attr", "Seas", "CAT" (pre-filled, not editable) |
| C | Program | Formula | =Section 1 program name |
| D | LOB | Input/Text | Line of business for curve selection ("Property"/"Casualty") |
| E | LossTL | Input/Integer | Loss development curve selection (1-100 trend level) |
| F | CntTL | Input/Integer | Count development curve selection (1-100 trend level) |
| G-J | Q1-Q4 ELR | Input/Pct | Expected Loss Ratio per quarter |
| K | Severity | Input/Currency | Average claim severity ($) |
| L-O | Q1-Q4 Freq | Input/Number | Claim frequency per quarter (per $1M earned premium) |
| P | Check | Formula | =Freq×Sev/1000000 vs blended ELR (display-only validation) |

**Loss type behavior (per UWM):**
- **Attritional:** Uniform ELR. Q2-Q4 auto-set = Q1 by DomainEngine. User only enters Q1.
- **Seasonal:** 4 independent quarterly ELRs. Annual cycle repeats.
- **Catastrophe:** 4 quarterly ELRs; typically one non-zero quarter.

**Named Ranges:**
- UWIN_LossBlock_{N}: Range for all 3 loss type rows of program N (3 rows × 15 columns)

### Section 4: QS Reinsurance Terms (rows 63–78)

```
Row 63: [Section 4: Quota Share Reinsurance]                (section header)
Row 64: Headers: A:# | B:(blank) | C:Program | D:(blank) | E:SubjY1 | F:Cede%Y1 | G:CdComm%Y1 | H:FrFee%Y1 | I:SubjY2 | ...
Row 65: Program 1  [input cells, blue]
...
Row 74: Program 10
Row 75: (blank separator)
```

**Columns (repeating 4-column block per year, 5 years):**
| Col | Header | DataType | Description |
|---|---|---|---|
| E+(y×4) | SubjectY{y} | Input/Text | Subject business description |
| F+(y×4) | Cede%Y{y} | Input/Pct | QS cede percentage (e.g., 70%) |
| G+(y×4) | CdComm%Y{y} | Input/Pct | Ceding commission rate |
| H+(y×4) | FrFee%Y{y} | Input/Pct | Fronting fee rate |

**Named Ranges:**
- UWIN_Reins_{N}: Range for all 5 years of program N (1 row × 20 columns)

---

## How DomainEngine Reads This Tab

The DomainEngine does NOT use `KernelConfig.InputValue()`. Instead it reads the UW Inputs tab directly:

```vba
Public Sub ReadUWInputs()
    Dim wsUW As Worksheet
    Set wsUW = ThisWorkbook.Sheets("UW Inputs")
    
    ' Section 1: Program identity + premium schedule
    For p = 1 To MAX_PROGRAMS
        Dim dataRow As Long
        dataRow = UWIN_S1_DATA_ROW + p - 1
        If Len(Trim(CStr(wsUW.Cells(dataRow, UWIN_COL_NAME).Value))) = 0 Then Exit For
        m_progName(p) = Trim(CStr(wsUW.Cells(dataRow, UWIN_COL_NAME).Value))
        m_progBU(p) = Trim(CStr(wsUW.Cells(dataRow, UWIN_COL_BU).Value))
        m_progTerm(p) = CLng(wsUW.Cells(dataRow, UWIN_COL_TERM).Value)
        For q = 1 To 20
            m_gwpSchedule(p, q) = CDbl(wsUW.Cells(dataRow, UWIN_GWP_START_COL + q - 1).Value)
        Next q
        m_gwpGrowth(p) = CDbl(wsUW.Cells(dataRow, UWIN_GROWTH_COL).Value)
    Next p
    
    ' Section 2: Commission rates
    ' Section 3: Loss assumptions (3 rows per program)
    ' Section 4: Reinsurance terms (5-year blocks)
End Sub
```

Constants for row/column positions are defined in the DomainEngine module (insurance-specific, not kernel constants).

---

## Relationship to Other Tabs

| Tab | What it reads from UW Inputs |
|---|---|
| DomainEngine.Execute | Everything — drives the monthly Detail computation |
| Assumptions | Nothing — Assumptions has its own global params |
| UW Executive Summary | Nothing — reads from QuarterlySummary (which reads from Detail) |
| Investments | Nothing — independent input tab |

---

## Comparison to UWM ProgramInputs

| | UWM ProgramInputs | RDK UW Inputs |
|---|---|---|
| Section 1 | Premium schedule (20 quarters + growth) | Same — 20 quarterly GWP + growth rate |
| Section 2 | Commission rates (5 years) | Same — 5 annual rates |
| Section 3 | 3 loss types × (LOB, LossTL, CntTL, Q1-Q4 ELR, Sev, Q1-Q4 Freq, Check) | Same — 30 rows for 10 programs × 3 types |
| Section 4 | QS reins (5 years × Subject/Cede%/CdComm%/FrFee%) | Same — 5-year blocks, 4 terms per year |
| Mechanism | VBA reads cells directly | Same — DomainEngine reads cells directly |
| Layout | Hardcoded row/col positions | formula_tab_config defines layout, DomainEngine uses constants |
| Entity limit | 10 programs | 10 programs (configurable via MAX_PROGRAMS) |

The UW Inputs tab is functionally identical to UWM ProgramInputs. The difference is implementation: RDK generates it from formula_tab_config.csv (config-driven layout) instead of hardcoded VBA formatting.
