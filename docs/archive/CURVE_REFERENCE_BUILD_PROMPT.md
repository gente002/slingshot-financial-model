# RDK Curve Reference Tab — Claude Code Build Prompt

**Date:** April 2026
**Depends on:** Phase 12B COMPLETE
**Scope:** 1 new tab (Curve Reference) + CurveRefPct() UDF wrapper in Ext_CurveLib.bas

## BEFORE WRITING ANY CODE

Read `CLAUDE.md`, then `SESSION_NOTES.md`, then `engine/Ext_CurveLib.bas`.

---

## Goal

Build a "Curve Reference" tab that shows cumulative % of ultimate for all 8 development curves at trend level increments of 10 (TL=10, 20, 30, 40, 50, 60, 70, 80, 90, 100). No charts — table only.

---

## Step 1: Add CurveRefPct() wrapper to Ext_CurveLib.bas

Add this public function (callable as an Excel UDF):

```vba
Public Function CurveRefPct(lob As String, curveType As String, _
                            trendLevel As Long, age As Long) As Double
    ' Thin wrapper: resolves curve params by TL, evaluates CDF, returns % of ultimate.
    ' Callable from Excel cells: =CurveRefPct("Property","Paid",50,12)
    If trendLevel < 1 Or trendLevel > 100 Or age < 1 Then
        CurveRefPct = 0
        Exit Function
    End If
    
    Dim params As Variant
    params = GetCurveParamsByTL(lob, curveType, trendLevel)
    ' params: (1)=distName, (2)=p1, (3)=p2, (4)=p3, (5)=maxAge
    
    Dim distName As String: distName = CStr(params(1))
    Dim p1 As Double: p1 = CDbl(params(2))
    Dim p2 As Double: p2 = CDbl(params(3))
    Dim p3 As Double: p3 = CDbl(params(4))
    Dim maxAge As Long: maxAge = CLng(params(5))
    
    If age > maxAge Then
        CurveRefPct = 1
        Exit Function
    End If
    
    CurveRefPct = EvaluateCurve(distName, p1, p2, age, maxAge, p3)
End Function
```

Also add a companion function to return the MaxAge for a given LOB/CurveType/TL:

```vba
Public Function CurveRefMaxAge(lob As String, curveType As String, _
                               trendLevel As Long) As Long
    If trendLevel < 1 Or trendLevel > 100 Then
        CurveRefMaxAge = 0
        Exit Function
    End If
    Dim params As Variant
    params = GetCurveParamsByTL(lob, curveType, trendLevel)
    CurveRefMaxAge = CLng(params(5))
End Function
```

---

## Step 2: Add Curve Reference tab to tab_registry

Add to `config_insurance/tab_registry.csv`:

```
"Curve Reference","Domain","Input","N","Visible","18","Development curve shapes by trend level","FALSE","","FALSE"
```

QuarterlyColumns=FALSE — this is a static reference tab, not a quarterly formula tab. SortOrder=18 (after Sales Funnel at 17).

---

## Step 3: Build the tab via VBA (NOT formula_tab_config)

Since this tab uses UDF formulas (`=CurveRefPct(...)`) rather than kernel formula placeholders, it should be built by a VBA subroutine that writes the formulas directly to the worksheet. Add a `BuildCurveReferenceTab` subroutine to Ext_CurveLib.bas (or a new companion module if CurveLib is near the size limit). Call it from KernelBootstrap after tab creation.

### Tab Layout

**Row 1:** Section header "Curve Reference" (Bold, navy background 1F3864, white font)
**Row 2:** "Development Patterns — Cumulative % of Ultimate by Trend Level"

**Row 4:** Section: "PROPERTY — Paid Loss Development"

**Row 5:** Column headers: Dev Age (months) | TL=10 | TL=20 | TL=30 | TL=40 | TL=50 | TL=60 | TL=70 | TL=80 | TL=90 | TL=100

**Rows 6-26:** Development ages: 1, 3, 6, 9, 12, 18, 24, 30, 36, 42, 48, 54, 60, 72, 84, 96, 108, 120, 144, 180, 240

Each cell formula: `=CurveRefPct("Property","Paid",{TL},{age})`
Format: 0.0%

**Row 27:** MaxAge row — `=CurveRefMaxAge("Property","Paid",{TL})` with label "Max Age (months)". Format: integer. Italic, grey font.

**Row 28:** Spacer

**Then repeat the same block for:**
- Property — Case Incurred Development
- Property — Reported Count Development
- Property — Closed Count Development
- Casualty — Paid Loss Development
- Casualty — Case Incurred Development
- Casualty — Reported Count Development
- Casualty — Closed Count Development

**Total: 8 blocks × ~25 rows (21 ages + header + MaxAge + spacer + section) = ~200 rows.**

### Development Ages

Use these 21 ages for every block:

```
1, 3, 6, 9, 12, 18, 24, 30, 36, 42, 48, 54, 60, 72, 84, 96, 108, 120, 144, 180, 240
```

240 months = 20 years — clean actuarial horizon. Casualty curves at high TLs will still be below 100% at 240, which is the point.

### Formatting

- Age column (col B): integer, right-aligned
- % cells: 0.0% format
- MaxAge row: integer, Italic, grey font (808080)
- Column headers (TL=10, TL=20, ...): Bold, center-aligned
- Section headers: Bold, D9E1F2 fill, 000000 font
- Alternate row shading (light grey F2F2F2 every other data row) for readability
- Column widths: col B = 16, TL columns = 10

---

## Step 4: Wire into KernelBootstrap

After the tab is created by the tab_registry loop, call `BuildCurveReferenceTab` to populate it. Gate this behind a check: only run if the "Curve Reference" tab exists (i.e., insurance config only).

```vba
' In KernelBootstrap, after tab creation loop:
If SheetExists("Curve Reference") Then
    Ext_CurveLib.BuildCurveReferenceTab
End If
```

---

## Step 5: Logging

Append to SESSION_NOTES.md. Log any bugs found. Update CLAUDE.md tab count. Do not modify any other tabs or formulas.

No named ranges needed — this tab is informational only.

---

## Validation

1. Open workbook, Run Model
2. Curve Reference tab shows 8 blocks of development percentages
3. Property Paid at TL=50, Age=12: should show a meaningful % (not 0%, not 100%)
4. All TL=10 columns develop faster (higher % at early ages) than TL=100 columns
5. At MaxAge, the % should be 100% (or very close)
6. At Age=240, Casualty curves at high TLs should be below 100%
7. Casualty curves develop slower than Property curves at the same TL
8. MaxAge row shows increasing values as TL increases
9. No #VALUE! or #REF! errors
10. Ext_CurveLib.bas remains under 64KB after adding the 2 UDFs + BuildCurveReferenceTab
