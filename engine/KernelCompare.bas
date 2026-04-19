Attribute VB_Name = "KernelCompare"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' ---
' KernelCompare.bas
' Purpose: Structured comparison of snapshots and current state.
'          Produces delta tabs with conditional formatting, comparison CSVs,
'          input comparison, and summary comparison. (AD-40: separate module)
' ---

' ---
' CompareSnapshots
' ---
Public Sub CompareSnapshots(baseName As String, variantName As String, _
                            Optional includeInputs As Boolean = True, _
                            Optional includeSummary As Boolean = True, _
                            Optional threshold As Double = 0.000001)
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Dim pr As String: pr = GetProjectRoot()
    Dim bDir As String: bDir = pr & "\" & DIR_SNAPSHOTS & "\" & baseName
    Dim vDir As String: vDir = pr & "\" & DIR_SNAPSHOTS & "\" & variantName
    If Not ValidateSnapshotExists(bDir, baseName) Then GoTo Cleanup
    If Not ValidateSnapshotExists(vDir, variantName) Then GoTo Cleanup
    Dim bd As Variant: bd = LoadCsvToArray(bDir & "\detail.csv")
    If IsEmpty(bd) Then GoTo Cleanup
    Dim vd As Variant: vd = LoadCsvToArray(vDir & "\detail.csv")
    If IsEmpty(vd) Then GoTo Cleanup
    Dim inpSrc As String: inpSrc = "snapshot"
    RunCoreComparison bd, vd, baseName, variantName, threshold, includeInputs, includeSummary, inpSrc
Cleanup:
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    LogCompareError "snapshots", Err.Description
    Resume Cleanup
End Sub

' ---
' CompareCurrentToSnapshot
' ---
Public Sub CompareCurrentToSnapshot(snapshotName As String, _
                                    Optional includeInputs As Boolean = True, _
                                    Optional threshold As Double = 0.000001)
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Dim pr As String: pr = GetProjectRoot()
    Dim bDir As String: bDir = pr & "\" & DIR_SNAPSHOTS & "\" & snapshotName
    If Not ValidateSnapshotExists(bDir, snapshotName) Then GoTo Cleanup
    Dim bd As Variant: bd = LoadCsvToArray(bDir & "\detail.csv")
    If IsEmpty(bd) Then GoTo Cleanup
    Dim vd As Variant: vd = LoadCurrentDetailToArray()
    If IsEmpty(vd) Then
        KernelConfig.LogError SEV_ERROR, "KernelCompare", "E-720", _
            "Detail tab has no data. Run RunProjections first.", _
            "MANUAL BYPASS: Run RunProjections to populate the Detail tab, then retry."
        MsgBox "Detail tab has no data. Run RunProjections first.", vbExclamation, "RDK -- Compare"
        GoTo Cleanup
    End If
    Dim inpSrc As String: inpSrc = "current_vs_snapshot"
    RunCoreComparison bd, vd, snapshotName, "Current", threshold, includeInputs, False, inpSrc, snapshotName
Cleanup:
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    LogCompareError "current vs snapshot", Err.Description
    Resume Cleanup
End Sub

' ---
' CompareInputsOnly
' ---
Public Sub CompareInputsOnly(baseName As String, variantName As String)
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    WriteInputComparisonTab baseName, variantName
    KernelConfig.LogError SEV_INFO, "KernelCompare", "I-740", _
        "Input comparison complete: " & baseName & " vs " & variantName, "snapshot"
    MsgBox "Input comparison complete. See the InputComp tab.", vbInformation, "RDK -- Compare Inputs"
Cleanup:
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    LogCompareError "inputs", Err.Description
    Resume Cleanup
End Sub

' ---
' RemoveComparisonTabs
' ---
Public Sub RemoveComparisonTabs()
    On Error GoTo ErrHandler
    Application.DisplayAlerts = False
    Dim removed As Long: removed = 0
    Dim i As Long
    For i = ThisWorkbook.Sheets.Count To 1 Step -1
        Dim sn As String: sn = ThisWorkbook.Sheets(i).Name
        If Left(sn, Len(COMPARE_TAB_PREFIX)) = COMPARE_TAB_PREFIX Or _
           Left(sn, Len(COMPARE_INPUT_PREFIX)) = COMPARE_INPUT_PREFIX Or _
           Left(sn, Len(COMPARE_SUMMARY_PREFIX)) = COMPARE_SUMMARY_PREFIX Then
            ThisWorkbook.Sheets(i).Delete
            removed = removed + 1
        End If
    Next i
    Application.DisplayAlerts = True
    KernelConfig.LogError SEV_INFO, "KernelCompare", "I-750", _
        "Removed " & removed & " comparison tab(s)", ""
    MsgBox "Removed " & removed & " comparison tab(s).", vbInformation, "RDK -- Clear Comparisons"
    Exit Sub
ErrHandler:
    Application.DisplayAlerts = True
    KernelConfig.LogError SEV_ERROR, "KernelCompare", "E-759", _
        "Error removing comparison tabs: " & Err.Description, _
        "MANUAL BYPASS: Right-click and delete tabs starting with Compare_, InputComp_, or SummComp_."
    MsgBox "Error: " & Err.Description, vbCritical, "RDK -- Clear Comparisons"
End Sub

' ---
' CompareSnapshotPrompt
' ---
Public Sub CompareSnapshotPrompt()
    Dim names() As String
    names = KernelSnapshot.ListSnapshots()
    If UBound(names) < 2 Then
        MsgBox "Need at least 2 saved snapshots to compare.", vbInformation, "RDK -- Compare"
        Exit Sub
    End If
    Dim lst As String
    lst = BuildNameList(names)
    Dim bn As String: bn = InputBox("Select BASE snapshot:" & vbCrLf & lst & vbCrLf & "Enter number or name:", "Compare Snapshots")
    If Len(bn) = 0 Then Exit Sub
    bn = ResolvePick(bn, names)
    Dim vn As String: vn = InputBox("Select VARIANT snapshot (compare against base):" & vbCrLf & lst & vbCrLf & "Enter number or name:", "Compare Snapshots")
    If Len(vn) = 0 Then Exit Sub
    vn = ResolvePick(vn, names)
    CompareSnapshots bn, vn
End Sub

' ---
' CompareCurrentToSnapshotPrompt
' ---
Public Sub CompareCurrentToSnapshotPrompt()
    Dim names() As String
    names = KernelSnapshot.ListSnapshots()
    If UBound(names) < 1 Then
        MsgBox "No snapshots found.", vbInformation, "RDK -- Compare"
        Exit Sub
    End If
    Dim lst As String
    lst = BuildNameList(names)
    Dim sp As String: sp = InputBox("Select snapshot to compare current state against:" & vbCrLf & lst & vbCrLf & "Enter number or name:", "Compare to Snapshot")
    If Len(sp) = 0 Then Exit Sub
    sp = ResolvePick(sp, names)
    CompareCurrentToSnapshot sp
End Sub

' ---
' CompareInputsPrompt
' ---
Public Sub CompareInputsPrompt()
    Dim names() As String
    names = KernelSnapshot.ListSnapshots()
    If UBound(names) < 2 Then
        MsgBox "Need at least 2 snapshots to compare inputs.", vbInformation, "RDK -- Compare"
        Exit Sub
    End If
    Dim lst As String
    lst = BuildNameList(names)
    Dim bn As String: bn = InputBox("Select BASE:" & vbCrLf & lst & vbCrLf & "Enter number or name:", "Compare Inputs")
    If Len(bn) = 0 Then Exit Sub
    bn = ResolvePick(bn, names)
    Dim vn As String: vn = InputBox("Select VARIANT:" & vbCrLf & lst & vbCrLf & "Enter number or name:", "Compare Inputs")
    If Len(vn) = 0 Then Exit Sub
    vn = ResolvePick(vn, names)
    CompareInputsOnly bn, vn
End Sub

' ---
' BuildNameList (Private) - Formats a numbered list for InputBox display
' ---
Private Function BuildNameList(names() As String) As String
    Dim msg As String
    msg = ""
    Dim i As Long
    For i = 1 To UBound(names)
        msg = msg & vbCrLf & "  " & i & ". " & names(i)
    Next i
    BuildNameList = msg
End Function

' ---
' ResolvePick (Private) - Resolves a number or name input against a names array
' ---
Private Function ResolvePick(input1 As String, names() As String) As String
    If IsNumeric(input1) Then
        Dim idx As Long
        idx = CLng(input1)
        If idx >= 1 And idx <= UBound(names) Then
            ResolvePick = names(idx)
            Exit Function
        End If
    End If
    ResolvePick = Trim(input1)
End Function


' ---
' PRIVATE: Core comparison engine
' ---

Private Sub RunCoreComparison(baseData As Variant, variantData As Variant, _
                               baseName As String, variantName As String, _
                               threshold As Double, includeInputs As Boolean, _
                               includeSummary As Boolean, inputSourceType As String, _
                               Optional inputBaseName As String = "")
    If Not CheckStructuralCompat(baseData, variantData, baseName, variantName) Then Exit Sub

    Dim metricNames() As String
    Dim metricCount As Long
    GetMetricInfo metricNames, metricCount
    If metricCount = 0 Then
        KernelConfig.LogError SEV_ERROR, "KernelCompare", "E-700", _
            "No metric columns found in column registry", _
            "MANUAL BYPASS: Check Config sheet COLUMN_REGISTRY for Incremental or Derived columns."
        Exit Sub
    End If

    Dim grainCols() As Long
    GetGrainCols grainCols

    Dim matched As Object
    Set matched = MatchRowsByGrain(baseData, variantData, grainCols)

    Dim changedCount As Long
    Dim compData As Variant
    compData = ComputeDeltas(baseData, variantData, matched, metricNames, metricCount, grainCols, threshold, changedCount)

    WriteOutputComparisonTab compData, baseName, variantName, threshold, metricNames, metricCount, changedCount
    WriteComparisonCsv compData, baseName, variantName, metricNames, metricCount, grainCols

    If includeInputs Then
        Dim inpBase As String
        If Len(inputBaseName) > 0 Then
            inpBase = inputBaseName
        Else
            inpBase = baseName
        End If
        WriteInputComparisonTab inpBase, variantName
    End If

    If includeSummary Then
        WriteSummaryComparisonTab compData, baseName, variantName, metricNames, metricCount, grainCols
    End If

    Dim totalRows As Long: totalRows = 0
    If IsArray(compData) Then totalRows = UBound(compData, 1)

    KernelConfig.LogError SEV_INFO, "KernelCompare", "I-700", _
        "Comparison complete: " & baseName & " vs " & variantName, _
        totalRows & " rows, " & changedCount & " changed"
    MsgBox "Comparison complete: " & totalRows & " rows compared, " & changedCount & " changed." & vbCrLf & _
           "See comparison tabs.", vbInformation, "RDK -- Compare"
End Sub

Private Sub LogCompareError(context As String, errDesc As String)
    KernelConfig.LogError SEV_ERROR, "KernelCompare", "E-799", _
        "Error comparing " & context & ": " & errDesc, _
        "MANUAL BYPASS: Open both CSV files in a diff tool to compare manually."
    MsgBox "Error comparing " & context & ": " & errDesc, vbCritical, "RDK -- Compare"
End Sub


' ---
' PRIVATE: Validation helpers
' ---

Private Function GetProjectRoot() As String
    Dim wbPath As String: wbPath = ThisWorkbook.Path
    GetProjectRoot = Left(wbPath, InStrRev(wbPath, "\") - 1)
End Function

Private Function ValidateSnapshotExists(sDir As String, sName As String) As Boolean
    ValidateSnapshotExists = False
    If Dir(sDir, vbDirectory) = "" Then
        KernelConfig.LogError SEV_ERROR, "KernelCompare", "E-701", _
            "Snapshot not found: " & sName, _
            "MANUAL BYPASS: Check snapshots/ directory. Run ShowSnapshots."
        MsgBox "Snapshot not found: " & sName, vbExclamation, "RDK -- Compare"
        Exit Function
    End If
    If Dir(sDir & "\manifest.json") = "" Then
        KernelConfig.LogError SEV_ERROR, "KernelCompare", "E-702", _
            "Snapshot incomplete (no manifest.json): " & sName, _
            "MANUAL BYPASS: Snapshot may not have saved properly."
        MsgBox "Snapshot incomplete: " & sName, vbExclamation, "RDK -- Compare"
        Exit Function
    End If
    ValidateSnapshotExists = True
End Function


' ---
' PRIVATE: Data loading
' ---

Private Function LoadCsvToArray(csvPath As String) As Variant
    LoadCsvToArray = Empty
    If Dir(csvPath) = "" Then
        KernelConfig.LogError SEV_ERROR, "KernelCompare", "E-705", _
            "CSV not found: " & csvPath, _
            "MANUAL BYPASS: Open the CSV in Excel manually."
        Exit Function
    End If
    Dim fn As Integer: fn = FreeFile
    Dim fc As String
    Open csvPath For Binary Access Read As #fn
    Dim fs As Long: fs = LOF(fn)
    If fs = 0 Then
        Close #fn
        Exit Function
    End If
    fc = Space$(fs)
    Get #fn, , fc
    Close #fn
    fc = Replace(Replace(fc, vbCrLf, vbLf), vbCr, vbLf)
    If Right(fc, 1) = vbLf Then fc = Left(fc, Len(fc) - 1)
    Dim lines() As String: lines = Split(fc, vbLf)
    Dim lc As Long: lc = UBound(lines) + 1
    If lc = 0 Then Exit Function
    Dim hf() As String: hf = ParseCsvLine(lines(0))
    Dim cc As Long: cc = UBound(hf) + 1
    Dim result() As Variant
    ReDim result(1 To lc, 1 To cc)
    Dim ri As Long
    For ri = 0 To UBound(lines)
        If Len(Trim(lines(ri))) = 0 Then GoTo NextCL
        Dim flds() As String: flds = ParseCsvLine(lines(ri))
        Dim ci As Long
        For ci = 0 To UBound(flds)
            If ci < cc Then
                If IsNumeric(flds(ci)) And Len(flds(ci)) > 0 Then
                    result(ri + 1, ci + 1) = CDbl(flds(ci))
                Else
                    result(ri + 1, ci + 1) = flds(ci)
                End If
            End If
        Next ci
NextCL:
    Next ri
    LoadCsvToArray = result
End Function

Private Function LoadCurrentDetailToArray() As Variant
    LoadCurrentDetailToArray = Empty
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(TAB_DETAIL)
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim lc As Long: lc = KernelConfig.GetColumnCount()
    If lc = 0 Then lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lr < DETAIL_DATA_START_ROW Then Exit Function
    LoadCurrentDetailToArray = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)).Value
End Function

Private Function LoadInputsFromCsv(csvPath As String) As Variant
    LoadInputsFromCsv = Empty
    If Dir(csvPath) = "" Then Exit Function
    Dim fn As Integer: fn = FreeFile
    Dim fc As String
    Open csvPath For Binary Access Read As #fn
    Dim fs As Long: fs = LOF(fn)
    If fs = 0 Then
        Close #fn
        Exit Function
    End If
    fc = Space$(fs)
    Get #fn, , fc
    Close #fn
    fc = Replace(Replace(fc, vbCrLf, vbLf), vbCr, vbLf)
    If Right(fc, 1) = vbLf Then fc = Left(fc, Len(fc) - 1)
    Dim lines() As String: lines = Split(fc, vbLf)
    If UBound(lines) < 0 Then Exit Function
    Dim hf() As String: hf = ParseCsvLine(lines(0))
    Dim cc As Long: cc = UBound(hf) + 1
    Dim lc As Long: lc = UBound(lines) + 1
    Dim result() As Variant
    ReDim result(1 To lc, 1 To cc)
    Dim ri As Long
    For ri = 0 To UBound(lines)
        If Len(Trim(lines(ri))) = 0 Then GoTo NextICL
        Dim flds() As String: flds = ParseCsvLine(lines(ri))
        Dim ci As Long
        For ci = 0 To UBound(flds)
            If ci < cc Then result(ri + 1, ci + 1) = flds(ci)
        Next ci
NextICL:
    Next ri
    LoadInputsFromCsv = result
End Function

Private Function LoadCurrentInputs() As Variant
    LoadCurrentInputs = Empty
    Dim pc As Long: pc = KernelConfig.GetInputCount()
    If pc = 0 Then Exit Function
    Dim ec As Long: ec = DetectEntityCount()
    If ec = 0 Then Exit Function
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(TAB_INPUTS)
    Dim cc As Long: cc = 2 + ec
    Dim result() As Variant
    ReDim result(1 To pc + 1, 1 To cc)
    result(1, 1) = "Section"
    result(1, 2) = "ParamName"
    Dim ei As Long
    For ei = 1 To ec
        result(1, 2 + ei) = "Entity" & ei
    Next ei
    Dim pi As Long
    For pi = 1 To pc
        result(pi + 1, 1) = KernelConfig.GetInputSection(pi)
        result(pi + 1, 2) = KernelConfig.GetInputParam(pi)
        Dim pr As Long: pr = KernelConfig.GetInputRow(pi)
        For ei = 1 To ec
            result(pi + 1, 2 + ei) = ws.Cells(pr, INPUT_ENTITY_START_COL + ei - 1).Value
        Next ei
    Next pi
    LoadCurrentInputs = result
End Function

Private Function DetectEntityCount() As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(TAB_INPUTS)
    Dim pc As Long: pc = KernelConfig.GetInputCount()
    Dim er As Long: er = 0
    Dim pi As Long
    For pi = 1 To pc
        If StrComp(KernelConfig.GetInputParam(pi), "EntityName", vbTextCompare) = 0 Then
            er = KernelConfig.GetInputRow(pi)
            Exit For
        End If
    Next pi
    If er = 0 Then
        DetectEntityCount = 0
        Exit Function
    End If
    Dim cnt As Long: cnt = 0
    Dim ci As Long: ci = INPUT_ENTITY_START_COL
    Do While ci < INPUT_ENTITY_START_COL + INPUT_MAX_ENTITIES
        If Trim(CStr(ws.Cells(er, ci).Value)) = "" Then Exit Do
        cnt = cnt + 1
        ci = ci + 1
    Loop
    DetectEntityCount = cnt
End Function


' ---
' PRIVATE: Structural checks and metric info
' ---

Private Function CheckStructuralCompat(bd As Variant, vd As Variant, _
                                        bn As String, vn As String) As Boolean
    CheckStructuralCompat = False
    Dim bc As Long: bc = UBound(bd, 2)
    Dim vc As Long: vc = UBound(vd, 2)
    If bc <> vc Then
        KernelConfig.LogError SEV_ERROR, "KernelCompare", "E-706", _
            "Structural incompatibility: Base=" & bc & " cols, Variant=" & vc, _
            "MANUAL BYPASS: Export both CSVs and compare manually."
        MsgBox "Cannot compare: column count mismatch (" & bc & " vs " & vc & ").", _
               vbCritical, "RDK -- Compare"
        Exit Function
    End If
    Dim diffs As String: diffs = ""
    Dim ci As Long
    For ci = 1 To bc
        If StrComp(CStr(bd(1, ci)), CStr(vd(1, ci)), vbTextCompare) <> 0 Then
            If Len(diffs) > 0 Then diffs = diffs & ", "
            diffs = diffs & "col" & ci & "(" & CStr(bd(1, ci)) & "/" & CStr(vd(1, ci)) & ")"
        End If
    Next ci
    If Len(diffs) > 0 Then
        KernelConfig.LogError SEV_ERROR, "KernelCompare", "E-707", _
            "Column names differ: " & diffs, _
            "MANUAL BYPASS: Export both CSVs and compare manually."
        MsgBox "Cannot compare: column name mismatch.", vbCritical, "RDK -- Compare"
        Exit Function
    End If
    CheckStructuralCompat = True
End Function

Private Sub GetMetricInfo(ByRef mn() As String, ByRef mc As Long)
    Dim tc As Long: tc = KernelConfig.GetColumnCount()
    mc = 0
    Dim i As Long
    For i = 1 To tc
        Dim f As String: f = KernelConfig.GetFieldClass(KernelConfig.GetColName(i))
        If f = "Incremental" Or f = "Derived" Then mc = mc + 1
    Next i
    If mc = 0 Then Exit Sub
    ReDim mn(1 To mc)
    Dim p As Long: p = 0
    For i = 1 To tc
        f = KernelConfig.GetFieldClass(KernelConfig.GetColName(i))
        If f = "Incremental" Or f = "Derived" Then
            p = p + 1
            mn(p) = KernelConfig.GetColName(i)
        End If
    Next i
End Sub

Private Sub GetGrainCols(ByRef gc() As Long)
    ReDim gc(1 To 2)
    gc(1) = KernelConfig.ColIndex("EntityName")
    gc(2) = KernelConfig.TryColIndex("Period")
    If gc(2) < 1 Then gc(2) = KernelConfig.ColIndex("CalPeriod")
End Sub


' ---
' PRIVATE: Row matching and delta computation
' ---

Private Function MatchRowsByGrain(bd As Variant, vd As Variant, gc() As Long) As Object
    Dim bDict As Object: Set bDict = CreateObject("Scripting.Dictionary")
    bDict.CompareMode = vbTextCompare
    Dim rDict As Object: Set rDict = CreateObject("Scripting.Dictionary")
    rDict.CompareMode = vbTextCompare
    Dim br As Long
    For br = 2 To UBound(bd, 1)
        Dim bk As String: bk = GrainKey(bd, br, gc)
        If Len(bk) > 0 And Not bDict.Exists(bk) Then bDict.Add bk, br
    Next br
    Dim vr As Long
    For vr = 2 To UBound(vd, 1)
        Dim vk As String: vk = GrainKey(vd, vr, gc)
        If Len(vk) = 0 Then GoTo NextVR
        If bDict.Exists(vk) Then
            rDict.Add vk, Array(CLng(bDict(vk)), vr)
            bDict.Remove vk
        Else
            rDict.Add vk, Array(CLng(-1), vr)
        End If
NextVR:
    Next vr
    Dim bKeys As Variant: bKeys = bDict.Keys
    Dim ki As Long
    For ki = 0 To bDict.Count - 1
        rDict.Add bKeys(ki), Array(CLng(bDict(bKeys(ki))), CLng(-1))
    Next ki
    Set MatchRowsByGrain = rDict
End Function

Private Function GrainKey(d As Variant, r As Long, gc() As Long) As String
    Dim k As String: k = ""
    Dim gi As Long
    For gi = 1 To UBound(gc)
        If gi > 1 Then k = k & "|"
        k = k & CStr(d(r, gc(gi)))
    Next gi
    GrainKey = k
End Function

Private Function ComputeDeltas(bd As Variant, vd As Variant, matched As Object, _
                                mn() As String, mc As Long, gc() As Long, _
                                threshold As Double, ByRef changed As Long) As Variant
    Dim gcn As Long: gcn = UBound(gc)
    Dim oc As Long: oc = gcn + (mc * COMPARE_COLS_PER_METRIC) + 1
    Dim tr As Long: tr = matched.Count
    If tr = 0 Then
        ComputeDeltas = Empty
        changed = 0
        Exit Function
    End If
    Dim result() As Variant
    ReDim result(1 To tr, 1 To oc)
    changed = 0
    Dim keys As Variant: keys = matched.Keys
    Dim ki As Long
    For ki = 0 To matched.Count - 1
        Dim or2 As Long: or2 = ki + 1
        Dim pair As Variant: pair = matched(keys(ki))
        Dim bRow As Long: bRow = CLng(pair(0))
        Dim vRow As Long: vRow = CLng(pair(1))
        Dim gi As Long
        For gi = 1 To gcn
            If vRow > 0 Then
                result(or2, gi) = vd(vRow, gc(gi))
            Else
                result(or2, gi) = bd(bRow, gc(gi))
            End If
        Next gi
        Dim anyChg As Boolean: anyChg = False
        Dim mi As Long
        For mi = 1 To mc
            Dim mci As Long: mci = KernelConfig.ColIndex(mn(mi))
            Dim bo As Long: bo = gcn + (mi - 1) * COMPARE_COLS_PER_METRIC + 1
            If bRow < 0 Then
                result(or2, bo) = "--"
                If IsNumeric(vd(vRow, mci)) Then
                    result(or2, bo + 1) = CDbl(vd(vRow, mci))
                Else
                    result(or2, bo + 1) = vd(vRow, mci)
                End If
                result(or2, bo + 2) = "--"
                result(or2, bo + 3) = "NEW"
                anyChg = True
            ElseIf vRow < 0 Then
                If IsNumeric(bd(bRow, mci)) Then
                    result(or2, bo) = CDbl(bd(bRow, mci))
                Else
                    result(or2, bo) = bd(bRow, mci)
                End If
                result(or2, bo + 1) = "--"
                result(or2, bo + 2) = "--"
                result(or2, bo + 3) = "REMOVED"
                anyChg = True
            Else
                Dim bv As Double: bv = 0
                Dim vv As Double: vv = 0
                If IsNumeric(bd(bRow, mci)) Then bv = CDbl(bd(bRow, mci))
                If IsNumeric(vd(vRow, mci)) Then vv = CDbl(vd(vRow, mci))
                result(or2, bo) = bv
                result(or2, bo + 1) = vv
                Dim dlt As Double: dlt = vv - bv
                result(or2, bo + 2) = dlt
                If bv = 0 And vv <> 0 Then
                    result(or2, bo + 3) = "N/A"
                    anyChg = True
                ElseIf bv <> 0 And vv = 0 Then
                    result(or2, bo + 3) = -1#
                    anyChg = True
                ElseIf bv = 0 And vv = 0 Then
                    result(or2, bo + 3) = 0#
                Else
                    result(or2, bo + 3) = (vv - bv) / bv
                    If Abs(dlt) > threshold Then anyChg = True
                End If
            End If
        Next mi
        result(or2, oc) = anyChg
        If anyChg Then changed = changed + 1
    Next ki
    ComputeDeltas = result
End Function


' ---
' PRIVATE: Output writing
' ---

Private Sub WriteOutputComparisonTab(cd As Variant, bn As String, vn As String, _
                                      threshold As Double, mn() As String, _
                                      mc As Long, changed As Long)
    If IsEmpty(cd) Then Exit Sub
    Dim gcn As Long: gcn = 2
    Dim oc As Long: oc = gcn + (mc * COMPARE_COLS_PER_METRIC) + 1
    Dim sn As String
    sn = COMPARE_TAB_PREFIX & Left(bn, 10) & "_v_" & Left(vn, 10)
    If Len(sn) > 31 Then sn = Left(sn, 31)
    Dim ws As Worksheet: Set ws = EnsureCompSheet(sn)
    ws.Cells.ClearContents
    Dim tr As Long: tr = UBound(cd, 1)
    ' Title
    ws.Cells(1, 1).Value = "Comparison: " & vn & " vs " & bn
    ws.Range(ws.Cells(1, 1), ws.Cells(1, oc)).Merge
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.Bold = True
    ' Metadata
    ws.Cells(2, 1).Value = "Generated: " & FmtTS() & "  |  Threshold: " & threshold & "  |  Changed: " & changed & " of " & tr
    ' Headers row 4
    ws.Cells(4, 1).Value = "EntityName"
    ws.Cells(4, 2).Value = "Period"
    Dim mi As Long
    For mi = 1 To mc
        Dim bc As Long: bc = gcn + (mi - 1) * COMPARE_COLS_PER_METRIC + 1
        ws.Cells(4, bc).Value = mn(mi) & "_Base"
        ws.Cells(4, bc + 1).Value = mn(mi) & "_Variant"
        ws.Cells(4, bc + 2).Value = mn(mi) & "_Delta"
        ws.Cells(4, bc + 3).Value = mn(mi) & "_Pct"
    Next mi
    ws.Cells(4, oc).Value = "AnyChange"
    With ws.Range(ws.Cells(4, 1), ws.Cells(4, oc))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    ' Data (PT-001)
    If tr > 0 Then
        ws.Range(ws.Cells(5, 1), ws.Cells(4 + tr, oc)).Value = cd
    End If
    ' Number formats
    Dim er As Long: er = 4 + tr
    If er >= 5 Then
        For mi = 1 To mc
            bc = gcn + (mi - 1) * COMPARE_COLS_PER_METRIC + 1
            ws.Range(ws.Cells(5, bc), ws.Cells(er, bc + 1)).NumberFormat = "#,##0.000000"
            ws.Range(ws.Cells(5, bc + 2), ws.Cells(er, bc + 2)).NumberFormat = "#,##0.000000"
            ws.Range(ws.Cells(5, bc + 3), ws.Cells(er, bc + 3)).NumberFormat = "0.00%"
        Next mi
        ApplyDeltaFmt ws, 5, er, gcn, mc, oc
    End If
    ws.Range(ws.Cells(4, 1), ws.Cells(4, oc)).AutoFilter
    ws.Rows(5).Select
    ActiveWindow.FreezePanes = True
    ws.Cells(1, 1).Select
    ws.Columns.AutoFit
End Sub

Private Sub ApplyDeltaFmt(ws As Worksheet, sr As Long, er As Long, _
                           gcn As Long, mc As Long, oc As Long)
    Dim mi As Long
    For mi = 1 To mc
        Dim dc As Long: dc = gcn + (mi - 1) * COMPARE_COLS_PER_METRIC + 3
        Dim pc As Long: pc = dc + 1
        Dim rw As Long
        For rw = sr To er
            Dim dv As Variant: dv = ws.Cells(rw, dc).Value
            If IsNumeric(dv) Then
                If CDbl(dv) > 0 Then
                    ws.Cells(rw, dc).Font.Color = RGB(0, 128, 0)
                ElseIf CDbl(dv) < 0 Then
                    ws.Cells(rw, dc).Font.Color = RGB(192, 0, 0)
                Else
                    ws.Cells(rw, dc).Font.Color = RGB(160, 160, 160)
                End If
            End If
            Dim pv As Variant: pv = ws.Cells(rw, pc).Value
            If IsNumeric(pv) Then
                If CDbl(pv) > 0 Then
                    ws.Cells(rw, pc).Font.Color = RGB(0, 128, 0)
                ElseIf CDbl(pv) < 0 Then
                    ws.Cells(rw, pc).Font.Color = RGB(192, 0, 0)
                Else
                    ws.Cells(rw, pc).Font.Color = RGB(160, 160, 160)
                End If
            End If
        Next rw
    Next mi
    Dim rw2 As Long
    For rw2 = sr To er
        If ws.Cells(rw2, oc).Value = True Then
            ws.Cells(rw2, oc).Interior.Color = RGB(255, 255, 200)
            ws.Range(ws.Cells(rw2, 1), ws.Cells(rw2, oc - 1)).Interior.Color = RGB(255, 255, 230)
        End If
    Next rw2
End Sub

Private Sub WriteComparisonCsv(cd As Variant, bn As String, vn As String, _
                                mn() As String, mc As Long, gc() As Long)
    If IsEmpty(cd) Then Exit Sub
    Dim pr As String: pr = GetProjectRoot()
    Dim sd As String: sd = pr & "\" & DIR_SNAPSHOTS
    If Dir(sd, vbDirectory) = "" Then MkDir sd
    Dim csvPath As String
    csvPath = sd & "\compare_" & Left(bn, 15) & "_vs_" & Left(vn, 15) & "_" & FmtTSFile() & ".csv"
    Dim tmp As String: tmp = csvPath & ".tmp"
    Dim gcn As Long: gcn = UBound(gc)
    Dim oc As Long: oc = gcn + (mc * COMPARE_COLS_PER_METRIC) + 1
    Dim fn As Integer: fn = FreeFile
    Open tmp For Output As #fn
    Dim hl As String: hl = """EntityName"",""Period"""
    Dim mi As Long
    For mi = 1 To mc
        hl = hl & ",""" & mn(mi) & "_Base"",""" & mn(mi) & "_Variant"",""" & mn(mi) & "_Delta"",""" & mn(mi) & "_Pct"""
    Next mi
    hl = hl & ",""AnyChange"""
    Print #fn, hl
    Dim tr As Long: tr = UBound(cd, 1)
    Dim rw As Long
    For rw = 1 To tr
        Dim dl As String: dl = ""
        Dim ci As Long
        For ci = 1 To oc
            If ci > 1 Then dl = dl & ","
            Dim cv As Variant: cv = cd(rw, ci)
            If IsNumeric(cv) And Not IsEmpty(cv) Then
                If VarType(cv) = vbBoolean Then
                    dl = dl & CStr(cv)
                Else
                    dl = dl & FmtNum6(CDbl(cv))
                End If
            Else
                dl = dl & """" & Replace(CStr(cv), """", """""") & """"
            End If
        Next ci
        Print #fn, dl
    Next rw
    Close #fn
    If Dir(csvPath) <> "" Then Kill csvPath
    Name tmp As csvPath
End Sub

Private Sub WriteInputComparisonTab(bn As String, vn As String)
    Dim pr As String: pr = GetProjectRoot()
    Dim bi As Variant
    Dim vi As Variant
    If StrComp(vn, "Current", vbTextCompare) = 0 Then
        bi = LoadInputsFromCsv(pr & "\" & DIR_SNAPSHOTS & "\" & bn & "\inputs.csv")
        vi = LoadCurrentInputs()
    Else
        bi = LoadInputsFromCsv(pr & "\" & DIR_SNAPSHOTS & "\" & bn & "\inputs.csv")
        vi = LoadInputsFromCsv(pr & "\" & DIR_SNAPSHOTS & "\" & vn & "\inputs.csv")
    End If
    If IsEmpty(bi) Or IsEmpty(vi) Then
        KernelConfig.LogError SEV_WARN, "KernelCompare", "W-740", _
            "Could not load input files for comparison", _
            "MANUAL BYPASS: Check that inputs.csv exists in both source directories."
        Exit Sub
    End If
    Dim sn As String
    sn = COMPARE_INPUT_PREFIX & Left(bn, 8) & "_v_" & Left(vn, 8)
    If Len(sn) > 31 Then sn = Left(sn, 31)
    Dim ws As Worksheet: Set ws = EnsureCompSheet(sn)
    ws.Cells.ClearContents
    ws.Cells(1, 1).Value = "Input Comparison: " & vn & " vs " & bn
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 8)).Merge
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(3, 1).Value = "Section"
    ws.Cells(3, 2).Value = "ParamName"
    ws.Cells(3, 3).Value = "EntityName"
    ws.Cells(3, 4).Value = "Base_Value"
    ws.Cells(3, 5).Value = "Variant_Value"
    ws.Cells(3, 6).Value = "Delta"
    ws.Cells(3, 7).Value = "Pct"
    ws.Cells(3, 8).Value = "Changed"
    With ws.Range("A3:H3")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    Dim bRows As Long: bRows = UBound(bi, 1)
    Dim bCols As Long: bCols = UBound(bi, 2)
    Dim vCols As Long: vCols = UBound(vi, 2)
    Dim bEnt As Long: bEnt = bCols - 2
    Dim vEnt As Long: vEnt = vCols - 2
    Dim mxE As Long
    If bEnt > vEnt Then mxE = bEnt Else mxE = vEnt
    Dim mxR As Long: mxR = (bRows - 1) * mxE
    If mxR <= 0 Then Exit Sub
    Dim oa() As Variant
    ReDim oa(1 To mxR, 1 To 8)
    Dim outR As Long: outR = 0
    Dim rw As Long
    For rw = 2 To bRows
        Dim sec As String: sec = CStr(bi(rw, 1))
        Dim prm As String: prm = CStr(bi(rw, 2))
        Dim ei As Long
        For ei = 1 To mxE
            outR = outR + 1
            If outR > mxR Then Exit For
            oa(outR, 1) = sec
            oa(outR, 2) = prm
            oa(outR, 3) = "Entity" & ei
            Dim bVal As Variant: bVal = ""
            If ei + 2 <= bCols Then bVal = bi(rw, ei + 2)
            oa(outR, 4) = bVal
            Dim vVal As Variant: vVal = ""
            Dim vr As Long
            For vr = 2 To UBound(vi, 1)
                If StrComp(CStr(vi(vr, 1)), sec, vbTextCompare) = 0 And _
                   StrComp(CStr(vi(vr, 2)), prm, vbTextCompare) = 0 Then
                    If ei + 2 <= vCols Then vVal = vi(vr, ei + 2)
                    Exit For
                End If
            Next vr
            oa(outR, 5) = vVal
            If IsNumeric(bVal) And IsNumeric(vVal) And Len(CStr(bVal)) > 0 And Len(CStr(vVal)) > 0 Then
                Dim bd2 As Double: bd2 = CDbl(bVal)
                Dim vd2 As Double: vd2 = CDbl(vVal)
                oa(outR, 6) = vd2 - bd2
                If bd2 <> 0 Then
                    oa(outR, 7) = (vd2 - bd2) / bd2
                ElseIf vd2 <> 0 Then
                    oa(outR, 7) = "N/A"
                Else
                    oa(outR, 7) = 0
                End If
                oa(outR, 8) = (Abs(vd2 - bd2) > COMPARE_DEFAULT_THRESHOLD)
            Else
                If CStr(bVal) = CStr(vVal) Then
                    oa(outR, 6) = ""
                    oa(outR, 7) = ""
                    oa(outR, 8) = False
                Else
                    oa(outR, 6) = "Changed"
                    oa(outR, 7) = ""
                    oa(outR, 8) = True
                End If
            End If
        Next ei
    Next rw
    If outR > 0 Then
        Dim wa() As Variant
        ReDim wa(1 To outR, 1 To 8)
        Dim wr As Long
        For wr = 1 To outR
            Dim wc As Long
            For wc = 1 To 8
                wa(wr, wc) = oa(wr, wc)
            Next wc
        Next wr
        ws.Range(ws.Cells(4, 1), ws.Cells(3 + outR, 8)).Value = wa
        For wr = 4 To 3 + outR
            Dim dv2 As Variant: dv2 = ws.Cells(wr, 6).Value
            If IsNumeric(dv2) Then
                If CDbl(dv2) > 0 Then ws.Cells(wr, 6).Font.Color = RGB(0, 128, 0)
                If CDbl(dv2) < 0 Then ws.Cells(wr, 6).Font.Color = RGB(192, 0, 0)
            End If
            If ws.Cells(wr, 8).Value = True Then ws.Cells(wr, 8).Interior.Color = RGB(255, 255, 200)
        Next wr
        ws.Range(ws.Cells(4, 7), ws.Cells(3 + outR, 7)).NumberFormat = "0.00%"
    End If
    ws.Range("A3:H3").AutoFilter
    ws.Columns.AutoFit
End Sub

Private Sub WriteSummaryComparisonTab(cd As Variant, bn As String, vn As String, _
                                       mn() As String, mc As Long, gc() As Long)
    If IsEmpty(cd) Then Exit Sub
    Dim gcn As Long: gcn = UBound(gc)
    Dim tr As Long: tr = UBound(cd, 1)
    Dim sn As String
    sn = COMPARE_SUMMARY_PREFIX & Left(bn, 7) & "_v_" & Left(vn, 7)
    If Len(sn) > 31 Then sn = Left(sn, 31)
    Dim ws As Worksheet: Set ws = EnsureCompSheet(sn)
    ws.Cells.ClearContents
    ws.Cells(1, 1).Value = "Summary Comparison: " & vn & " vs " & bn
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 5)).Merge
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.Bold = True
    ' Collect entity names
    Dim eDict As Object: Set eDict = CreateObject("Scripting.Dictionary")
    eDict.CompareMode = vbTextCompare
    Dim rw As Long
    For rw = 1 To tr
        Dim en As String: en = CStr(cd(rw, 1))
        If Not eDict.Exists(en) And Len(en) > 0 Then eDict.Add en, en
    Next rw
    Dim cr As Long: cr = 3
    ' TOTAL section
    ws.Cells(cr, 1).NumberFormat = "@"
    ws.Cells(cr, 1).Value = "=== TOTAL ==="
    ws.Cells(cr, 1).Font.Bold = True
    ws.Cells(cr, 1).Interior.Color = RGB(217, 225, 242)
    cr = cr + 1
    ws.Cells(cr, 1).Value = "Metric"
    ws.Cells(cr, 2).Value = "Base_Total"
    ws.Cells(cr, 3).Value = "Variant_Total"
    ws.Cells(cr, 4).Value = "Delta"
    ws.Cells(cr, 5).Value = "Pct"
    With ws.Range(ws.Cells(cr, 1), ws.Cells(cr, 5))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    cr = cr + 1
    Dim tBases() As Double: ReDim tBases(1 To mc)
    Dim tVars() As Double: ReDim tVars(1 To mc)
    Dim mi As Long
    For mi = 1 To mc
        Dim bo As Long: bo = gcn + (mi - 1) * COMPARE_COLS_PER_METRIC + 1
        If KernelConfig.GetFieldClass(mn(mi)) = "Incremental" Then
            Dim sb As Double: sb = 0
            Dim sv As Double: sv = 0
            For rw = 1 To tr
                If IsNumeric(cd(rw, bo)) Then sb = sb + CDbl(cd(rw, bo))
                If IsNumeric(cd(rw, bo + 1)) Then sv = sv + CDbl(cd(rw, bo + 1))
            Next rw
            tBases(mi) = sb
            tVars(mi) = sv
        End If
    Next mi
    ' Derived totals
    For mi = 1 To mc
        If KernelConfig.GetFieldClass(mn(mi)) = "Derived" Then
            Dim rl As String: rl = KernelConfig.GetDerivationRule(mn(mi))
            If Len(rl) > 0 Then
                Dim oA As String: Dim oS As String: Dim oB As String
                If ParseRule(rl, oA, oS, oB) Then
                    Dim iA As Long: iA = FindMI(mn, mc, oA)
                    Dim iB As Long: iB = FindMI(mn, mc, oB)
                    If iA > 0 And iB > 0 Then
                        tBases(mi) = CalcOp(tBases(iA), tBases(iB), oS)
                        tVars(mi) = CalcOp(tVars(iA), tVars(iB), oS)
                    End If
                End If
            End If
        End If
    Next mi
    Dim tsr As Long: tsr = cr
    For mi = 1 To mc
        ws.Cells(cr, 1).Value = mn(mi)
        ws.Cells(cr, 2).Value = tBases(mi)
        ws.Cells(cr, 3).Value = tVars(mi)
        Dim td As Double: td = tVars(mi) - tBases(mi)
        ws.Cells(cr, 4).Value = td
        If tBases(mi) <> 0 Then ws.Cells(cr, 5).Value = td / tBases(mi) Else ws.Cells(cr, 5).Value = 0
        If td > 0 Then ws.Cells(cr, 4).Font.Color = RGB(0, 128, 0)
        If td < 0 Then ws.Cells(cr, 4).Font.Color = RGB(192, 0, 0)
        cr = cr + 1
    Next mi
    ws.Range(ws.Cells(tsr, 2), ws.Cells(cr - 1, 4)).NumberFormat = "#,##0.000000"
    ws.Range(ws.Cells(tsr, 5), ws.Cells(cr - 1, 5)).NumberFormat = "0.00%"
    ' Per-entity sections
    Dim eKeys As Variant: eKeys = eDict.Keys
    Dim ei As Long
    For ei = 0 To eDict.Count - 1
        cr = cr + 1
        Dim eName As String: eName = eKeys(ei)
        ws.Cells(cr, 1).NumberFormat = "@"
        ws.Cells(cr, 1).Value = "=== " & eName & " ==="
        ws.Cells(cr, 1).Font.Bold = True
        ws.Cells(cr, 1).Interior.Color = RGB(217, 225, 242)
        cr = cr + 1
        Dim eBases() As Double: ReDim eBases(1 To mc)
        Dim eVars() As Double: ReDim eVars(1 To mc)
        For mi = 1 To mc
            bo = gcn + (mi - 1) * COMPARE_COLS_PER_METRIC + 1
            If KernelConfig.GetFieldClass(mn(mi)) = "Incremental" Then
                sb = 0: sv = 0
                For rw = 1 To tr
                    If CStr(cd(rw, 1)) = eName Then
                        If IsNumeric(cd(rw, bo)) Then sb = sb + CDbl(cd(rw, bo))
                        If IsNumeric(cd(rw, bo + 1)) Then sv = sv + CDbl(cd(rw, bo + 1))
                    End If
                Next rw
                eBases(mi) = sb
                eVars(mi) = sv
            End If
        Next mi
        For mi = 1 To mc
            If KernelConfig.GetFieldClass(mn(mi)) = "Derived" Then
                rl = KernelConfig.GetDerivationRule(mn(mi))
                If Len(rl) > 0 Then
                    If ParseRule(rl, oA, oS, oB) Then
                        iA = FindMI(mn, mc, oA)
                        iB = FindMI(mn, mc, oB)
                        If iA > 0 And iB > 0 Then
                            eBases(mi) = CalcOp(eBases(iA), eBases(iB), oS)
                            eVars(mi) = CalcOp(eVars(iA), eVars(iB), oS)
                        End If
                    End If
                End If
            End If
        Next mi
        Dim esr As Long: esr = cr
        For mi = 1 To mc
            ws.Cells(cr, 1).Value = mn(mi)
            ws.Cells(cr, 2).Value = eBases(mi)
            ws.Cells(cr, 3).Value = eVars(mi)
            td = eVars(mi) - eBases(mi)
            ws.Cells(cr, 4).Value = td
            If eBases(mi) <> 0 Then ws.Cells(cr, 5).Value = td / eBases(mi) Else ws.Cells(cr, 5).Value = 0
            If td > 0 Then ws.Cells(cr, 4).Font.Color = RGB(0, 128, 0)
            If td < 0 Then ws.Cells(cr, 4).Font.Color = RGB(192, 0, 0)
            cr = cr + 1
        Next mi
        If cr > esr Then
            ws.Range(ws.Cells(esr, 2), ws.Cells(cr - 1, 4)).NumberFormat = "#,##0.000000"
            ws.Range(ws.Cells(esr, 5), ws.Cells(cr - 1, 5)).NumberFormat = "0.00%"
        End If
    Next ei
    ws.Columns.AutoFit
End Sub


' ---
' PRIVATE: Utility helpers
' ---

Private Function EnsureCompSheet(sn As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sn)
    On Error GoTo 0
    If Not ws Is Nothing Then
        KernelConfig.LogError SEV_WARN, "KernelCompare", "W-700", _
            "Replacing existing comparison sheet: " & sn, ""
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        Set ws = Nothing
    End If
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sn
    Set EnsureCompSheet = ws
End Function

Private Function ParseCsvLine(lineText As String) As String()
    Dim result() As String
    Dim fc As Long: fc = 0
    Dim mf As Long: mf = 1
    Dim sp As Long
    For sp = 1 To Len(lineText)
        If Mid(lineText, sp, 1) = "," Then mf = mf + 1
    Next sp
    ReDim result(0 To mf - 1)
    Dim pos As Long: pos = 1
    Dim ll As Long: ll = Len(lineText)
    Do While pos <= ll
        Dim fv As String: fv = ""
        If Mid(lineText, pos, 1) = """" Then
            pos = pos + 1
            Do While pos <= ll
                If Mid(lineText, pos, 1) = """" Then
                    If pos < ll And Mid(lineText, pos + 1, 1) = """" Then
                        fv = fv & """"
                        pos = pos + 2
                    Else
                        pos = pos + 1
                        Exit Do
                    End If
                Else
                    fv = fv & Mid(lineText, pos, 1)
                    pos = pos + 1
                End If
            Loop
            If pos <= ll And Mid(lineText, pos, 1) = "," Then pos = pos + 1
        Else
            Dim cp As Long: cp = InStr(pos, lineText, ",")
            If cp = 0 Then
                fv = Mid(lineText, pos)
                pos = ll + 1
            Else
                fv = Mid(lineText, pos, cp - pos)
                pos = cp + 1
            End If
        End If
        If fc <= UBound(result) Then result(fc) = fv
        fc = fc + 1
    Loop
    If fc > 0 And fc <= mf Then
        ReDim Preserve result(0 To fc - 1)
    ElseIf fc = 0 Then
        ReDim result(0 To 0)
    End If
    ParseCsvLine = result
End Function

Private Function ParseRule(rl As String, ByRef oA As String, ByRef oS As String, ByRef oB As String) As Boolean
    ParseRule = False
    Dim ops As Variant: ops = Array(" - ", " + ", " * ", " / ")
    Dim oi As Long
    For oi = LBound(ops) To UBound(ops)
        Dim p As Long: p = InStr(1, rl, ops(oi), vbTextCompare)
        If p > 0 Then
            oA = Trim(Mid(rl, 1, p - 1))
            oS = Trim(CStr(ops(oi)))
            oB = Trim(Mid(rl, p + Len(CStr(ops(oi)))))
            If Len(oA) > 0 And Len(oB) > 0 Then
                ParseRule = True
                Exit Function
            End If
        End If
    Next oi
End Function

Private Function CalcOp(a As Double, b As Double, op As String) As Double
    Select Case op
        Case "-": CalcOp = a - b
        Case "+": CalcOp = a + b
        Case "*": CalcOp = a * b
        Case "/":
            If b = 0 Then CalcOp = 0 Else CalcOp = a / b
        Case Else: CalcOp = 0
    End Select
End Function

Private Function FindMI(mn() As String, mc As Long, nm As String) As Long
    FindMI = 0
    Dim i As Long
    For i = 1 To mc
        If StrComp(mn(i), nm, vbTextCompare) = 0 Then
            FindMI = i
            Exit Function
        End If
    Next i
End Function

Private Function FmtTS() As String
    FmtTS = CStr(Year(Now)) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & " " & Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2)
End Function

Private Function FmtTSFile() As String
    FmtTSFile = CStr(Year(Now)) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
End Function

Private Function FmtNum6(val As Double) As String
    Dim ip As String: ip = CStr(Fix(val))
    Dim fv As Double: fv = Abs(val - Fix(val))
    Dim fi As Long: fi = CLng(fv * 1000000)
    FmtNum6 = ip & "." & Right("000000" & CStr(fi), 6)
End Function
