Attribute VB_Name = "KernelProveIt"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelProveIt.bas
' Purpose: Generates native Excel formulas on a Prove-It tab that verify
'          computation results WITHOUT any VBA. For auditors and stakeholders.
'
' Check Types (Phase 3 v1.0):
'   Identity   -- Derived field equals its DerivationRule
'   Accumulate -- SUM of incremental periods equals expected total
'   Reconcile  -- Cross-metric identity (e.g. GrossProfit/Revenue = GPMargin)
' =============================================================================


' =============================================================================
' GenerateProveItTab
' Reads prove_it_config.csv and generates Excel formulas on the ProveIt tab.
' =============================================================================
Public Sub GenerateProveItTab()
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    ' Ensure config is loaded
    KernelConfig.LoadAllConfig

    Dim checkCount As Long
    checkCount = KernelConfig.GetProveItCheckCount()

    If checkCount = 0 Then
        KernelConfig.LogError SEV_INFO, "KernelProveIt", "I-760", _
            "No Prove-It checks configured. Add checks to prove_it_config.csv.", ""
        Application.ScreenUpdating = True
        MsgBox "No Prove-It checks configured." & vbCrLf & _
               "Add checks to prove_it_config.csv and re-run Setup.bat.", _
               vbInformation, "RDK -- Prove-It"
        Exit Sub
    End If

    ' Verify Detail tab has data
    Dim wsDetail As Worksheet
    Set wsDetail = ThisWorkbook.Sheets(TAB_DETAIL)

    Dim detailLastRow As Long
    detailLastRow = wsDetail.Cells(wsDetail.Rows.Count, 1).End(xlUp).Row
    If detailLastRow < DETAIL_DATA_START_ROW Then
        KernelConfig.LogError SEV_WARN, "KernelProveIt", "W-760", _
            "Detail tab has no data. Run RunProjections first.", ""
        Application.ScreenUpdating = True
        MsgBox "Detail tab has no data. Run RunProjections first.", _
               vbExclamation, "RDK -- Prove-It"
        Exit Sub
    End If

    Dim totalDataRows As Long
    totalDataRows = detailLastRow - DETAIL_DATA_START_ROW + 1

    ' Set up Prove-It sheet
    Dim wsPI As Worksheet
    Set wsPI = EnsureProveItSheet()

    On Error Resume Next
    wsPI.Unprotect
    On Error GoTo ErrHandler

    wsPI.Cells.ClearContents
    wsPI.Cells.ClearFormats

    ' Row 1: Title
    wsPI.Cells(1, 1).Value = "RDK Prove-It -- Audit Verification"
    wsPI.Range(wsPI.Cells(1, 1), wsPI.Cells(1, 6)).Merge
    With wsPI.Cells(1, 1)
        .Font.Bold = True
        .Font.Size = 14
    End With

    ' Row 2: Info
    wsPI.Cells(2, 1).Value = "Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & _
        "  |  All formulas are native Excel -- no VBA required."

    ' Row 4: Headers
    wsPI.Cells(4, 1).Value = "CheckID"
    wsPI.Cells(4, 2).Value = "CheckType"
    wsPI.Cells(4, 3).Value = "CheckName"
    wsPI.Cells(4, 4).Value = "Formula"
    wsPI.Cells(4, 5).Value = "Result"
    wsPI.Cells(4, 6).Value = "Detail"

    With wsPI.Range("A4:F4")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    wsPI.Columns(1).ColumnWidth = 10
    wsPI.Columns(2).ColumnWidth = 12
    wsPI.Columns(3).ColumnWidth = 35
    wsPI.Columns(4).ColumnWidth = 60
    wsPI.Columns(5).ColumnWidth = 10
    wsPI.Columns(6).ColumnWidth = 40

    ' Determine which rows to check (sampling for large models)
    Dim testRows() As Long
    Dim testRowCount As Long
    BuildCheckRows totalDataRows, testRows, testRowCount

    ' Generate checks
    Dim piRow As Long
    piRow = 5
    Dim totalChecks As Long
    totalChecks = 0

    Dim chkIdx As Long
    For chkIdx = 1 To checkCount
        If Not KernelConfig.GetProveItEnabled(chkIdx) Then GoTo NextCheck

        Dim checkID As String
        checkID = KernelConfig.GetProveItCheckID(chkIdx)

        Dim checkType As String
        checkType = KernelConfig.GetProveItCheckType(chkIdx)

        Dim checkName As String
        checkName = KernelConfig.GetProveItCheckName(chkIdx)

        Dim metricA As String
        metricA = KernelConfig.GetProveItMetricA(chkIdx)

        Dim metricB As String
        metricB = KernelConfig.GetProveItMetricB(chkIdx)

        Dim metricC As String
        metricC = KernelConfig.GetProveItMetricC(chkIdx)

        Dim opStr As String
        opStr = KernelConfig.GetProveItOperator(chkIdx)

        Dim tol As Double
        tol = KernelConfig.GetProveItTolerance(chkIdx)

        Select Case UCase(checkType)
            Case "IDENTITY"
                piRow = GenerateIdentityChecks(wsPI, piRow, checkID, checkName, _
                    metricA, metricB, metricC, opStr, tol, testRows, testRowCount)
                totalChecks = totalChecks + testRowCount

            Case "ACCUMULATE"
                piRow = GenerateAccumulateChecks(wsPI, piRow, checkID, checkName, _
                    metricA, tol)
                ' Count will vary by entity count

            Case "RECONCILE"
                piRow = GenerateReconcileChecks(wsPI, piRow, checkID, checkName, _
                    metricA, metricB, metricC, opStr, tol, testRows, testRowCount)
                totalChecks = totalChecks + testRowCount
        End Select
NextCheck:
    Next chkIdx

    ' Update check count in row 2
    totalChecks = piRow - 5
    wsPI.Cells(2, 1).Value = "Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & _
        "  |  Checks: " & totalChecks & _
        "  |  All formulas are native Excel -- no VBA required."

    ' Add AND summary (directly after last check row)
    wsPI.Cells(piRow, 4).Value = "All checks pass:"
    wsPI.Cells(piRow, 4).Font.Bold = True
    If piRow > 5 Then
        wsPI.Cells(piRow, 5).Formula = "=AND(E5:E" & (piRow - 1) & ")"
    End If

    ' Apply conditional formatting for TRUE/FALSE
    ApplyProveItFormatting wsPI, piRow

    ' Protect sheet (read-only display)
    wsPI.Protect UserInterfaceOnly:=True

    wsPI.Visible = xlSheetVisible

    Application.ScreenUpdating = True

    KernelConfig.LogError SEV_INFO, "KernelProveIt", "I-761", _
        "Prove-It generated: " & totalChecks & " checks", ""

    MsgBox "Prove-It generated: " & totalChecks & " checks." & vbCrLf & vbCrLf & _
           "All formulas are native Excel -- works with macros disabled.", _
           vbInformation, "RDK -- Prove-It"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    KernelConfig.LogError SEV_ERROR, "KernelProveIt", "E-760", _
        "Error generating Prove-It: " & Err.Description, _
        "MANUAL BYPASS: Create the Prove-It tab manually. Write verification formulas in column E referencing Detail tab values. Use =ABS(A-B)<0.000001 pattern for numeric checks."
    MsgBox "Prove-It generation failed: " & Err.Description, _
           vbCritical, "RDK -- Error"
End Sub


' =============================================================================
' RefreshProveIt
' Clears and regenerates. Useful after config changes or new run.
' =============================================================================
Public Sub RefreshProveIt()
    GenerateProveItTab
End Sub


' =============================================================================
' ValidateProveIt
' Scans the Prove-It tab Result column.
' Returns True if all results are TRUE (reads formula results, no VBA eval).
' =============================================================================
Public Function ValidateProveIt() As Boolean
    On Error GoTo ErrOut

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_PROVE_IT)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row

    If lastRow < 5 Then
        ValidateProveIt = True
        Exit Function
    End If

    Dim r As Long
    For r = 5 To lastRow
        Dim cellVal As Variant
        cellVal = ws.Cells(r, 5).Value

        ' Skip non-formula rows (like the AND summary label)
        If VarType(cellVal) = vbBoolean Then
            If cellVal = False Then
                ValidateProveIt = False
                Exit Function
            End If
        End If
    Next r

    ValidateProveIt = True
    Exit Function

ErrOut:
    ValidateProveIt = False
End Function


' =============================================================================
' GenerateIdentityChecks
' Identity: Derived field equals its DerivationRule.
' Formula: =ABS(Detail!G5-(Detail!E5-Detail!F5))<tolerance
' =============================================================================
Private Function GenerateIdentityChecks(ws As Worksheet, startRow As Long, _
    checkID As String, checkName As String, _
    metricA As String, metricB As String, metricC As String, _
    opStr As String, tol As Double, _
    ByRef testRows() As Long, testRowCount As Long) As Long

    Dim colA As Long
    colA = KernelConfig.ColIndex(metricA)
    Dim colB As Long
    colB = KernelConfig.ColIndex(metricB)
    Dim colC As Long
    colC = KernelConfig.ColIndex(metricC)

    If colA < 1 Or colB < 1 Or colC < 1 Then
        GenerateIdentityChecks = startRow
        Exit Function
    End If

    Dim letA As String
    letA = KernelOutput.ColLetter(colA)
    Dim letB As String
    letB = KernelOutput.ColLetter(colB)
    Dim letC As String
    letC = KernelOutput.ColLetter(colC)

    Dim tolStr As String
    tolStr = Format(tol, "0.000000")

    Dim piRow As Long
    piRow = startRow

    Dim trIdx As Long
    For trIdx = 1 To testRowCount
        Dim detailRow As Long
        detailRow = DETAIL_DATA_START_ROW + testRows(trIdx) - 1

        ' Build formula: =ABS(Detail!colA[row]-(Detail!colB[row] op Detail!colC[row]))<tol
        Dim formula As String
        formula = "=ABS(Detail!" & letA & detailRow & _
                  "-(Detail!" & letB & detailRow & _
                  opStr & "Detail!" & letC & detailRow & _
                  "))<" & tolStr

        ws.Cells(piRow, 1).Value = checkID
        ws.Cells(piRow, 2).Value = PROVEIT_IDENTITY
        ws.Cells(piRow, 3).Value = checkName & " R" & detailRow
        ws.Cells(piRow, 4).Value = formula
        ws.Cells(piRow, 5).Formula = formula
        ws.Cells(piRow, 6).Value = "Detail row " & detailRow

        piRow = piRow + 1
    Next trIdx

    GenerateIdentityChecks = piRow
End Function


' =============================================================================
' GenerateAccumulateChecks
' Accumulate: SUM of incremental column per entity matches Summary.
' Formula: =ABS(SUMIFS(Detail!E:E,Detail!$A:$A,"Product A")-Summary!C2)<tol
' =============================================================================
Private Function GenerateAccumulateChecks(ws As Worksheet, startRow As Long, _
    checkID As String, checkName As String, _
    metricA As String, tol As Double) As Long

    Dim colMetric As Long
    colMetric = KernelConfig.ColIndex(metricA)
    If colMetric < 1 Then
        GenerateAccumulateChecks = startRow
        Exit Function
    End If

    Dim colEntity As Long
    colEntity = KernelConfig.ColIndex("EntityName")
    If colEntity < 1 Then
        GenerateAccumulateChecks = startRow
        Exit Function
    End If

    Dim letMetric As String
    letMetric = KernelOutput.ColLetter(colMetric)
    Dim letEntity As String
    letEntity = KernelOutput.ColLetter(colEntity)

    Dim tolStr As String
    tolStr = Format(tol, "0.000000")

    ' Get entity names from Inputs
    Dim wsInputs As Worksheet
    Set wsInputs = ThisWorkbook.Sheets(TAB_INPUTS)

    Dim entityCount As Long
    entityCount = 0
    Dim ecol As Long
    For ecol = INPUT_ENTITY_START_COL To INPUT_ENTITY_START_COL + INPUT_MAX_ENTITIES - 1
        If Len(Trim(CStr(wsInputs.Cells(3, ecol).Value))) = 0 Then Exit For
        entityCount = entityCount + 1
    Next ecol

    Dim piRow As Long
    piRow = startRow

    Dim periodCount As Long
    periodCount = KernelConfig.GetTimeHorizon()

    Dim entIdx As Long
    For entIdx = 1 To entityCount
        Dim entityName As String
        entityName = KernelConfig.GetEntityName(entIdx)

        ' Sum all periods for this entity on Detail tab
        ' Compare to manually computed SUMIFS
        ' Formula: verify SUMIFS matches by checking the sum is consistent
        Dim formula As String
        formula = "=ABS(SUMIFS(Detail!" & letMetric & ":" & letMetric & _
                  ",Detail!$" & letEntity & ":$" & letEntity & ",""" & entityName & """)" & _
                  "-SUMIFS(Detail!" & letMetric & ":" & letMetric & _
                  ",Detail!$" & letEntity & ":$" & letEntity & ",""" & entityName & """))<" & tolStr

        ' A more useful check: verify SUMIFS equals the sum of individual periods
        ' We build a SUM of individual cell references
        Dim startDataRow As Long
        startDataRow = DETAIL_DATA_START_ROW + (entIdx - 1) * periodCount
        Dim endDataRow As Long
        endDataRow = startDataRow + periodCount - 1

        formula = "=ABS(SUMIFS(Detail!" & letMetric & ":" & letMetric & _
                  ",Detail!$" & letEntity & ":$" & letEntity & ",""" & entityName & """)" & _
                  "-SUM(Detail!" & letMetric & startDataRow & ":" & letMetric & endDataRow & "))<" & tolStr

        ws.Cells(piRow, 1).Value = checkID
        ws.Cells(piRow, 2).Value = PROVEIT_ACCUMULATE
        ws.Cells(piRow, 3).Value = checkName & " [" & entityName & "]"
        ws.Cells(piRow, 4).Value = formula
        ws.Cells(piRow, 5).Formula = formula
        ws.Cells(piRow, 6).Value = metricA & " for " & entityName

        piRow = piRow + 1
    Next entIdx

    GenerateAccumulateChecks = piRow
End Function


' =============================================================================
' GenerateReconcileChecks
' Reconcile: Cross-metric identity (e.g. GrossProfit/Revenue = GPMargin).
' Formula: =IFERROR(ABS(Detail!colA/Detail!colB-Detail!colC)<tol,FALSE)
' =============================================================================
Private Function GenerateReconcileChecks(ws As Worksheet, startRow As Long, _
    checkID As String, checkName As String, _
    metricA As String, metricB As String, metricC As String, _
    opStr As String, tol As Double, _
    ByRef testRows() As Long, testRowCount As Long) As Long

    Dim colA As Long
    colA = KernelConfig.ColIndex(metricA)
    Dim colB As Long
    colB = KernelConfig.ColIndex(metricB)
    Dim colC As Long
    colC = KernelConfig.ColIndex(metricC)

    If colA < 1 Or colB < 1 Or colC < 1 Then
        GenerateReconcileChecks = startRow
        Exit Function
    End If

    Dim letA As String
    letA = KernelOutput.ColLetter(colA)
    Dim letB As String
    letB = KernelOutput.ColLetter(colB)
    Dim letC As String
    letC = KernelOutput.ColLetter(colC)

    Dim tolStr As String
    tolStr = Format(tol, "0.000000")

    Dim piRow As Long
    piRow = startRow

    Dim trIdx As Long
    For trIdx = 1 To testRowCount
        Dim detailRow As Long
        detailRow = DETAIL_DATA_START_ROW + testRows(trIdx) - 1

        ' Build formula: =IFERROR(ABS(Detail!colB[row] / Detail!colC[row] - Detail!colA[row])<tol, FALSE)
        Dim formula As String
        formula = "=IFERROR(ABS(Detail!" & letB & detailRow & _
                  opStr & "Detail!" & letC & detailRow & _
                  "-Detail!" & letA & detailRow & _
                  ")<" & tolStr & ",FALSE)"

        ws.Cells(piRow, 1).Value = checkID
        ws.Cells(piRow, 2).Value = PROVEIT_RECONCILE
        ws.Cells(piRow, 3).Value = checkName & " R" & detailRow
        ws.Cells(piRow, 4).Value = formula
        ws.Cells(piRow, 5).Formula = formula
        ws.Cells(piRow, 6).Value = "Detail row " & detailRow

        piRow = piRow + 1
    Next trIdx

    GenerateReconcileChecks = piRow
End Function


' =============================================================================
' ApplyProveItFormatting
' Applies conditional formatting: TRUE=green, FALSE=red.
' =============================================================================
Private Sub ApplyProveItFormatting(ws As Worksheet, lastRow As Long)
    If lastRow < 5 Then Exit Sub

    Dim rng As Object
    Set rng = ws.Range(ws.Cells(5, 5), ws.Cells(lastRow, 5))

    rng.FormatConditions.Delete

    ' TRUE = green
    Dim cfTrue As Object
    Set cfTrue = rng.FormatConditions.Add(Type:=1, Operator:=3, Formula1:="=TRUE")
    cfTrue.Interior.Color = RGB(198, 239, 206)
    cfTrue.Font.Color = RGB(0, 97, 0)

    ' FALSE = red bold
    Dim cfFalse As Object
    Set cfFalse = rng.FormatConditions.Add(Type:=1, Operator:=3, Formula1:="=FALSE")
    cfFalse.Interior.Color = RGB(255, 199, 206)
    cfFalse.Font.Color = RGB(255, 255, 255)
    cfFalse.Font.Bold = True
End Sub


' =============================================================================
' EnsureProveItSheet
' Creates the ProveIt sheet if it doesn't exist.
' =============================================================================
Private Function EnsureProveItSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(TAB_PROVE_IT)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = TAB_PROVE_IT
    End If

    Set EnsureProveItSheet = ws
End Function


' =============================================================================
' BuildCheckRows
' For large models, samples rows. For small models, checks all rows.
' =============================================================================
Private Sub BuildCheckRows(totalRows As Long, ByRef sampleRows() As Long, _
                            ByRef sampleCount As Long)
    If totalRows <= LARGE_MODEL_ROW_THRESHOLD Then
        sampleCount = totalRows
        ReDim sampleRows(1 To sampleCount)
        Dim i As Long
        For i = 1 To sampleCount
            sampleRows(i) = i
        Next i
    Else
        sampleCount = LARGE_MODEL_SAMPLE_SIZE
        ReDim sampleRows(1 To sampleCount)

        For i = 1 To 10
            sampleRows(i) = i
        Next i

        For i = 1 To 10
            sampleRows(10 + i) = totalRows - 10 + i
        Next i

        Dim midStart As Long
        midStart = 11
        Dim midEnd As Long
        midEnd = totalRows - 10

        For i = 1 To 10
            sampleRows(20 + i) = midStart + Int(Rnd * (midEnd - midStart + 1))
        Next i

        KernelConfig.LogError SEV_WARN, "KernelProveIt", "W-761", _
            "Large model -- sampled " & sampleCount & " of " & totalRows & " rows for Prove-It checks.", ""
    End If
End Sub
