# Assumptions Register — CC Build Prompt Addendum

Read `CLAUDE.md`, then `SESSION_NOTES.md`. This is an addendum to the current session.

## Overview

Build a full Assumptions Register: config CSV with 21 seed entries, kernel module (KernelAssumptions.bas), rendered visible tab with hyperlinks and conditional formatting, and a manager panel for CRUD operations. This is kernel-level — any domain model can use it.

---

## 1. NEW CONFIG: assumptions_config.csv

Create `config_insurance/assumptions_config.csv` with this schema:

```csv
"AssumptionID","Category","TabName","RowID","Description","Rationale","Source","Confidence","Sensitivity","SensitivityDetail","Owner","LastReviewed","History"
```

**Columns:** AssumptionID (unique A-001+), Category (Staffing/Revenue/Capital/Underwriting/Expense/Investment/Regulatory/Technology), TabName (from tab_registry), RowID (enables hyperlink, blank for general), Description (what the assumption IS), Rationale (WHY), Source (Management Estimate/Benchmark/LOI/Contract/Market Data/Regulatory), Confidence (High/Medium/Low), Sensitivity (High/Medium/Low), SensitivityDetail (brief impact explanation), Owner, LastReviewed (YYYY-MM-DD), History (semicolon-delimited change log).

### Seed 21 entries:

```csv
"A-001","Staffing","Staffing Expense","STF_HC_UW","2 UW staff in Y1, growing to 4 by Y3","Expected program volume of 3-5 programs requiring dedicated UW oversight","Management estimate","Medium","Medium","±1 head = ±$180K loaded cost","Ethan","2026-04-03","2026-04-03: Initial estimate"
"A-002","Staffing","Staffing Expense","STF_HC_CLAIMS","1 claims adjuster in Y1","Outsourced TPA handles most claims; in-house for oversight","Management estimate","High","Low","±1 head = ±$140K loaded cost","Ethan","2026-04-03","2026-04-03: Initial estimate"
"A-003","Staffing","Staffing Expense","STF_HC_ACTUARY","2 actuaries in Y1","Pricing + reserving require dedicated actuarial staff from day one","Management estimate","High","Medium","±1 head = ±$225K loaded cost","Ethan","2026-04-03","2026-04-03: Initial estimate"
"A-004","Staffing","Staffing Expense","STF_SAL_EXEC","$200K avg executive salary","Benchmarked to Des Moines/Midwest market. Below coastal markets.","Glassdoor, Payscale","Medium","Medium","±$30K per exec × 3 = ±$90K","Ethan","2026-04-03","2026-04-03: Initial estimate"
"A-005","Staffing","Staffing Expense","STF_BENEFITS","30% benefits loading factor","Covers health, dental, 401k match, FICA, disability, life","Industry benchmark","High","Low","±5% = ±$50K total","Ethan","2026-04-03","2026-04-03: Initial estimate"
"A-006","Staffing","Staffing Expense","STF_BONUS_EXEC","30% executive bonus target","Performance-based. Aligns with startup incentive structure.","Board compensation committee","Medium","Medium","±10% = ±$60K for 3 execs","Ethan","2026-04-03","2026-04-03: Initial estimate"
"A-007","Revenue","Other Revenue Detail","ORD_FEE_CARRIER_RATE","2% carrier access fee on GWP","Standard fronting/carrier fee range is 1-5%. Conservative at 2%.","Market survey of fronting carriers","Medium","High","±1% on $20M GWP = ±$200K revenue","Ethan","2026-04-03","2026-04-03: Initial estimate"
"A-008","Revenue","Other Revenue Detail","ORD_SW_PLATFORM_RATE","$25K/program/year platform subscription","Based on comparable InsurTech platform pricing. Mid-market.","Competitor analysis","Low","Medium","±$10K × 5 programs = ±$50K","Ethan","2026-04-03","2026-04-03: Initial estimate"
"A-009","Revenue","Other Revenue Detail","ORD_SW_API_RATE","$15/policy/year API fee","0.03% of avg premium ($50K). Market range $10-$25.","Competitor pricing","Low","Low","Small revenue driver at startup scale","Ethan","2026-04-03","2026-04-03: Initial estimate"
"A-010","Revenue","Other Revenue Detail","ORD_AVG_PREMIUM","$50K average premium per policy","Blended across GL, Property, Specialty. E&S tends higher.","Program pipeline analysis","Medium","Medium","Affects estimated policy count for per-policy fees","Ethan","2026-04-03","2026-04-03: Initial estimate"
"A-011","Capital","Capital Activity","CAP_EQ_AMT","$15M initial equity raise","Rally Ventures $10M + strategic reinsurer $5M committed","Term sheets","High","High","Determines initial RBC ratio and investable pool","Ethan","2026-04-03","2026-04-03: Based on committed capital"
"A-012","Investment","Investments","INV_GOV_ALLOC","40% govt bond allocation","Conservative for startup building track record. Shift higher yield after Y2.","Investment policy statement draft","Medium","Medium","10% shift changes portfolio yield by ~50bps","Ethan","2026-04-03","2026-04-03: Initial IPS draft"
"A-013","Investment","Investments","INV_MGMT_FEE_BPS","25 bps investment management fee","Mid-range for outsourced investment management","RFP responses","Medium","Low","±10bps on $20M = ±$20K","Ethan","2026-04-03","2026-04-03: Initial estimate"
"A-014","Underwriting","UW Inputs","","3 programs in Y1, growing to 7 by Y5","Pipeline-driven. GL confirmed, Property and Specialty in negotiation.","MGA pipeline tracker","Medium","High","Each program adds $3-5M GWP","Ethan","2026-04-03","2026-04-03: Based on pipeline"
"A-015","Underwriting","UW Inputs","","60% average expected loss ratio","Blended ELR. GL ~55%, Property ~65%, Specialty varies.","Actuarial pricing analysis","Medium","High","±5% ELR on $20M NEP = ±$1M loss","Ethan","2026-04-03","2026-04-03: Initial pricing review"
"A-016","Underwriting","UW Inputs","","70% average QS cession rate","High cession in early years to limit net retention while building surplus","Reinsurance strategy","High","High","±10% cession shifts $2M gross/net","Ethan","2026-04-03","2026-04-03: RI strategy confirmed"
"A-017","Expense","Other Expense Detail","OED_NP_AUDIT","$150K annual statutory audit","Mid-range for small carrier. Regional firm targeted.","Audit firm proposals","High","Low","Fixed cost, well-benchmarked","Ethan","2026-04-03","2026-04-03: Based on proposals"
"A-018","Expense","Other Expense Detail","OED_NP_LEGAL","$200K annual legal","Heavy Y1 for shell acquisition, regulatory, contracts. Declining Y2+.","Outside counsel estimates","Medium","Medium","Y1 could be $250-300K","Ethan","2026-04-03","2026-04-03: Initial estimate"
"A-019","Expense","Other Expense Detail","OED_NP_STARTUP","$100K one-time startup costs","Regulatory filings, initial technology buildout, branding","Management estimate","Medium","Low","One-time, Y1 only","Ethan","2026-04-03","2026-04-03: Initial estimate"
"A-020","Capital","RBC Capital Model","RBC_TARGET","300% target RBC ratio","Above 200% CAL with buffer for growth and volatility.","Board risk appetite","High","High","Determines capital call timing","Ethan","2026-04-03","2026-04-03: Board-approved target"
"A-021","Capital","RBC Capital Model","RBC_R2_RSV","15% reserve risk factor","NAIC standard P&C reserve risk charge","NAIC RBC instructions","High","Medium","Factor-based, well-established","Ethan","2026-04-03","2026-04-03: Per NAIC guidance"
```

---

## 2. NEW MODULE: KernelAssumptions.bas

### Public API

```vba
Public Sub GenerateAssumptionsRegister()
' Read assumptions_config from Config sheet
' Create or clear "Assumptions Register" tab
' Write summary: total count, by confidence, by category
' Write column headers (bold, frozen, auto-filter)
' Group by Category (section headers with D9E1F2 fill)
' For each assumption:
'   - Write all fields
'   - If TabName + RowID non-blank: create HYPERLINK("'" & TabName & "'!" & cell, RowID)
'   - Conditional format Confidence: High=C6EFCE, Medium=FFEB9C, Low=FFC7CE
'   - Conditional format Sensitivity: High=FFC7CE, Medium=FFEB9C, Low=C6EFCE (inverse)
' Column widths: ID=8, Category=12, Tab=20, Input=15, Description=40, Rationale=40,
'   Source=20, Confidence=12, Sensitivity=12, Impact=30, Owner=10, Reviewed=12, History=50
' Wrap text on Description, Rationale, History
' No gridlines

Public Sub ShowAssumptionManager()
' Main menu via MsgBox:
' "Assumption Manager" [View Register] [Add New] [Edit] [Archive] [Review Stale] [Cancel]

Public Sub AddAssumption(...)
' Prompt for all fields via InputBox sequence
' Auto-suggest next AssumptionID
' Default Owner = "Ethan", LastReviewed = today
' Initialize History = "{today}: Created"
' Write to Config sheet, regenerate register

Public Sub EditAssumption()
' Prompt for AssumptionID
' Show current values
' Prompt for field to update and new value
' Append History: "{today}: {field} changed from '{old}' to '{new}'"
' Update LastReviewed to today
' Write to Config sheet, regenerate register

Public Sub ArchiveAssumption()
' Prompt for AssumptionID, confirm
' Prefix Category with "ARCHIVED-"
' Append History: "{today}: Archived"
' Regenerate — archived assumptions at bottom in grey

Public Function GetStaleAssumptions(Optional daysThreshold As Long = 90) As String
' Find assumptions where LastReviewed > daysThreshold days ago
' Return comma-delimited list of IDs
```

### View Register
Activate the Assumptions Register tab. Regenerate if it doesn't exist.

### Add New Flow
1. AssumptionID: suggest next (e.g., "A-022")
2. Category: InputBox with hint "Staffing/Revenue/Capital/Underwriting/Expense/Investment/Regulatory/Technology"
3. TabName: InputBox with hint "Tab name from registry, or blank"
4. RowID: InputBox with hint "Input row ID, or blank for general"
5. Description: InputBox
6. Rationale: InputBox
7. Source: InputBox with hint "Management Estimate/Benchmark/LOI/Contract/Market Data/Regulatory"
8. Confidence: InputBox "High/Medium/Low"
9. Sensitivity: InputBox "High/Medium/Low"
10. SensitivityDetail: InputBox
11. Owner: default "Ethan" (InputBox with default)
12. Write + regenerate

### Edit Flow
1. InputBox: "Enter Assumption ID to edit (e.g., A-007)"
2. Find on Config sheet. If not found, error.
3. Display current Description + Rationale in MsgBox
4. InputBox: "Which field? Description/Rationale/Source/Confidence/Sensitivity/SensitivityDetail/Owner"
5. InputBox: "New value for {field}:"
6. Update + append History + regenerate

### Archive Flow
1. InputBox: "Enter Assumption ID to archive"
2. Confirm via MsgBox
3. Prefix Category, append History, regenerate

### Review Stale Flow
1. Scan for LastReviewed > 90 days
2. MsgBox: "Stale assumptions (>90 days since review): A-003, A-008, ..."
3. Offer: "Mark all as reviewed today?" [Yes] [No]

---

## 3. TAB REGISTRY

```csv
"Assumptions Register","Kernel","Output","N","Visible","17","Documented model assumptions with rationale and sensitivity","FALSE","","FALSE","548235","FALSE","FALSE"
```

---

## 4. DASHBOARD BUTTON

```csv
"Dashboard","ASSUMPTIONS","Manage Assumptions","KernelFormHelpers.ShowAssumptionManager","FALSE","17","TRUE","8","2"
```

Add to KernelFormHelpers.bas:
```vba
Public Sub ShowAssumptionManager()
    KernelAssumptions.ShowAssumptionManager
End Sub
```

---

## 5. GENERATE ON RUN MODEL

At end of KernelEngine.RunModel, after other post-run tasks:
```vba
KernelAssumptions.GenerateAssumptionsRegister
```

Also generate during bootstrap if the tab doesn't exist.

---

## 6. UPDATE CLAUDE.md

- Kernel modules: +1 (KernelAssumptions)
- Config tables: +1 (assumptions_config.csv)
- Note Assumptions Register tab

---

## VALIDATION GATES (Assumptions)

30. assumptions_config.csv exists with 21 seed entries
31. KernelAssumptions.bas exists with all public subs
32. Assumptions Register tab renders with all entries grouped by category
33. Hyperlinks navigate to correct input cells
34. Confidence: green/yellow/red conditional formatting
35. Sensitivity: red/yellow/green conditional formatting (inverse)
36. Add New creates entry and regenerates
37. Edit updates field and appends to History
38. Archive marks assumption inactive
39. Review Stale identifies >90 day assumptions
40. Button appears on Dashboard (user-visible)
41. Register regenerates after Run Model
42. SESSION_NOTES.md appended only
