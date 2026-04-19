Attribute VB_Name = "KernelTabIO"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelTabIO.bas
' Purpose: Generic tab export/import for snapshots and workspaces.
'          Exports/imports any worksheet's cell values as CSV.
' =============================================================================


' =============================================================================
' ExportAllInputTabs
' Exports every tab with Category=Input in tab_registry to dir/input_tabs/.
' Each tab gets a CSV file named after the tab (sanitized).
' =============================================================================
Public Sub ExportAllInputTabs(baseDir As String)
    On Error Resume Next
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    If wsConfig Is Nothing Then Exit Sub

    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_TAB_REGISTRY)
    If sr = 0 Then Exit Sub

    Dim tabDir As String
    tabDir = baseDir & "\input_tabs"
    KernelSnapshot.EnsureDirectoryExists tabDir

    Dim dr As Long
    dr = sr + 2
    Do While Len(Trim(CStr(wsConfig.Cells(dr, TREG_COL_TABNAME).Value))) > 0
        Dim tabName As String
        tabName = Trim(CStr(wsConfig.Cells(dr, TREG_COL_TABNAME).Value))
        Dim category As String
        category = Trim(CStr(wsConfig.Cells(dr, 3).Value))

        Dim tabType As String
        tabType = Trim(CStr(wsConfig.Cells(dr, 2).Value))

        ' Only export Domain Input tabs (skip Kernel tabs like Dashboard,
        ' Cover Page, User Guide, Assumptions -- BUG-161/BUG-162)
        If StrComp(category, "Input", vbTextCompare) = 0 And _
           StrComp(tabType, "Domain", vbTextCompare) = 0 Then
            Dim ws As Worksheet
            Set ws = Nothing
            Set ws = ThisWorkbook.Sheets(tabName)
            If Not ws Is Nothing Then
                Dim csvPath As String
                csvPath = tabDir & "\" & SanitizeTabName(tabName) & ".csv"
                ExportTabToCSV ws, csvPath
            End If
        End If
        dr = dr + 1
    Loop
    On Error GoTo 0
End Sub


' =============================================================================
' ImportAllInputTabs
' Restores input tab data from dir/input_tabs/ CSVs.
' =============================================================================
Public Sub ImportAllInputTabs(baseDir As String)
    On Error Resume Next
    Dim tabDir As String
    tabDir = baseDir & "\input_tabs"
    If Dir(tabDir, vbDirectory) = "" Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(tabDir) Then Exit Sub

    Dim csvFile As Object
    For Each csvFile In fso.GetFolder(tabDir).Files
        If StrComp(fso.GetExtensionName(csvFile.Name), "csv", vbTextCompare) = 0 Then
            Dim baseName As String
            baseName = fso.GetBaseName(csvFile.Name)
            Dim ws As Worksheet
            Set ws = Nothing
            Set ws = ThisWorkbook.Sheets(baseName)
            If ws Is Nothing Then
                Set ws = ThisWorkbook.Sheets(Replace(baseName, "_", " "))
            End If
            If Not ws Is Nothing Then
                ImportTabFromCSV ws, csvFile.Path
            End If
        End If
    Next csvFile
    Set fso = Nothing
    On Error GoTo 0
End Sub


' =============================================================================
' ExportTabToCSV
' Exports a worksheet's used range to CSV (values only).
' =============================================================================
Public Sub ExportTabToCSV(ws As Worksheet, csvPath As String)
    On Error Resume Next
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    ' Expand to used range if wider/taller
    If ws.UsedRange.row + ws.UsedRange.Rows.Count - 1 > lastRow Then
        lastRow = ws.UsedRange.row + ws.UsedRange.Rows.Count - 1
    End If
    If ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1 > lastCol Then
        lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    End If
    If lastRow < 1 Or lastCol < 1 Then Exit Sub

    ' Batch read all values
    Dim data As Variant
    data = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value

    ' Write CSV
    Dim fileNum As Integer
    fileNum = FreeFile
    Open csvPath For Output As #fileNum

    Dim r As Long
    For r = 1 To lastRow
        Dim lineStr As String
        lineStr = ""
        Dim c As Long
        For c = 1 To lastCol
            Dim cellVal As String
            If IsEmpty(data(r, c)) Then
                cellVal = ""
            Else
                cellVal = CStr(data(r, c))
            End If
            If InStr(1, cellVal, ",") > 0 Or InStr(1, cellVal, """") > 0 Or InStr(1, cellVal, vbLf) > 0 Then
                cellVal = """" & Replace(cellVal, """", """""") & """"
            End If
            If c > 1 Then lineStr = lineStr & ","
            lineStr = lineStr & cellVal
        Next c
        Print #fileNum, lineStr
    Next r
    Close #fileNum
    On Error GoTo 0
End Sub


' =============================================================================
' ImportTabFromCSV
' Reads a CSV and writes values back to a worksheet (batch array write).
' =============================================================================
Public Sub ImportTabFromCSV(ws As Worksheet, csvPath As String)
    On Error Resume Next
    Dim fileNum As Integer
    fileNum = FreeFile
    Dim fileContent As String
    Dim fileSize As Long

    Open csvPath For Binary Access Read As #fileNum
    fileSize = LOF(fileNum)
    If fileSize = 0 Then
        Close #fileNum
        Exit Sub
    End If
    fileContent = Space$(fileSize)
    Get #fileNum, , fileContent
    Close #fileNum

    fileContent = Replace(fileContent, vbCrLf, vbLf)
    fileContent = Replace(fileContent, vbCr, vbLf)

    Dim lines() As String
    lines = Split(fileContent, vbLf)

    ' Count lines and columns
    Dim lineCount As Long
    lineCount = 0
    Dim maxCols As Long
    maxCols = 0
    Dim li As Long
    For li = 0 To UBound(lines)
        If Len(Trim(lines(li))) > 0 Then
            lineCount = lineCount + 1
            Dim cc As Long
            cc = Len(lines(li)) - Len(Replace(lines(li), ",", ""))
            If cc + 1 > maxCols Then maxCols = cc + 1
        End If
    Next li
    If lineCount = 0 Or maxCols = 0 Then Exit Sub

    ws.Unprotect

    ' Build array
    Dim dataArr() As Variant
    ReDim dataArr(1 To lineCount, 1 To maxCols)
    Dim rowIdx As Long
    rowIdx = 0
    For li = 0 To UBound(lines)
        If Len(Trim(lines(li))) > 0 Then
            rowIdx = rowIdx + 1
            Dim fields() As String
            fields = ParseCSVLine(lines(li))
            Dim colIdx As Long
            For colIdx = 0 To UBound(fields)
                If colIdx + 1 <= maxCols Then
                    Dim fVal As String
                    fVal = fields(colIdx)
                    If IsNumeric(fVal) And Len(fVal) > 0 Then
                        dataArr(rowIdx, colIdx + 1) = CDbl(fVal)
                    Else
                        dataArr(rowIdx, colIdx + 1) = fVal
                    End If
                End If
            Next colIdx
        End If
    Next li

    ' Batch write: clear and overwrite entire used area
    ' RefreshFormulaTabs must be called after import to rebuild formula cells
    ws.Cells.ClearContents
    ws.Range(ws.Cells(1, 1), ws.Cells(lineCount, maxCols)).Value = dataArr
    On Error GoTo 0
End Sub


' SanitizeTabName -- converts tab name to safe filename
Private Function SanitizeTabName(tabName As String) As String
    Dim result As String
    result = Replace(Trim(tabName), " ", "_")
    Dim cleaned As String
    cleaned = ""
    Dim i As Long
    For i = 1 To Len(result)
        Dim ch As String
        ch = Mid(result, i, 1)
        If ch Like "[A-Za-z0-9_-]" Then cleaned = cleaned & ch
    Next i
    SanitizeTabName = cleaned
End Function


' ParseCSVLine -- CSV line parser with quoted field support
Private Function ParseCSVLine(lineText As String) As String()
    Dim fields As New Collection
    Dim pos As Long
    pos = 1
    Dim inQuote As Boolean
    inQuote = False
    Dim fieldVal As String
    fieldVal = ""

    Do While pos <= Len(lineText)
        Dim ch As String
        ch = Mid(lineText, pos, 1)
        If inQuote Then
            If ch = """" Then
                If pos < Len(lineText) And Mid(lineText, pos + 1, 1) = """" Then
                    fieldVal = fieldVal & """"
                    pos = pos + 1
                Else
                    inQuote = False
                End If
            Else
                fieldVal = fieldVal & ch
            End If
        Else
            If ch = """" Then
                inQuote = True
            ElseIf ch = "," Then
                fields.Add fieldVal
                fieldVal = ""
            Else
                fieldVal = fieldVal & ch
            End If
        End If
        pos = pos + 1
    Loop
    fields.Add fieldVal

    Dim result() As String
    ReDim result(0 To fields.Count - 1)
    Dim i As Long
    For i = 1 To fields.Count
        result(i - 1) = fields(i)
    Next i
    ParseCSVLine = result
End Function


' =============================================================================
' VerifyHealthAfterLoad
' Reads health_config and checks each indicator after snapshot restore.
' Returns a summary string: "All 3 health checks passed" or "2 passed, FAIL: ..."
' =============================================================================
Public Function VerifyHealthAfterLoad() As String
    On Error Resume Next
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    If wsConfig Is Nothing Then
        VerifyHealthAfterLoad = "Config sheet not found"
        Exit Function
    End If

    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_HEALTH_CONFIG)
    If sr = 0 Then
        VerifyHealthAfterLoad = "No health checks configured"
        Exit Function
    End If

    Dim failures As String
    failures = ""
    Dim passCount As Long
    passCount = 0

    Dim dr As Long
    dr = sr + 2
    Do While Len(Trim(CStr(wsConfig.Cells(dr, HLCFG_COL_TABNAME).Value))) > 0
        Dim hTabName As String
        hTabName = Trim(CStr(wsConfig.Cells(dr, HLCFG_COL_TABNAME).Value))
        Dim hRowID As String
        hRowID = Trim(CStr(wsConfig.Cells(dr, HLCFG_COL_ROWID).Value))
        Dim hCheckType As String
        hCheckType = Trim(CStr(wsConfig.Cells(dr, HLCFG_COL_CHECKTYPE).Value))
        Dim hGoodValue As String
        hGoodValue = Trim(CStr(wsConfig.Cells(dr, HLCFG_COL_GOODVALUE).Value))
        Dim hThreshold As String
        hThreshold = Trim(CStr(wsConfig.Cells(dr, HLCFG_COL_THRESHOLD).Value))

        Dim hWs As Worksheet
        Set hWs = Nothing
        Set hWs = ThisWorkbook.Sheets(hTabName)
        If Not hWs Is Nothing Then
            Dim hRow As Long
            hRow = KernelFormula.ResolveRowID(hTabName, hRowID)
            If hRow > 0 Then
                If StrComp(hCheckType, "NumericZero", vbTextCompare) = 0 Then
                    Dim thresh As Double
                    If IsNumeric(hThreshold) And Len(hThreshold) > 0 Then thresh = CDbl(hThreshold) Else thresh = 0.01
                    Dim worstVal As Double
                    worstVal = 0
                    Dim timeHorizon As Long
                    timeHorizon = KernelConfig.GetTimeHorizon()
                    If timeHorizon <= 0 Then timeHorizon = 12
                    Dim numYears As Long
                    numYears = (timeHorizon \ 3) \ QS_QUARTERS_PER_YEAR
                    If numYears < 1 Then numYears = 1
                    Dim lastQCol As Long
                    lastQCol = QS_DATA_START_COL + numYears * QS_COLS_PER_YEAR
                    Dim qc As Long
                    For qc = QS_DATA_START_COL To lastQCol
                        Dim cv As Variant
                        cv = hWs.Cells(hRow, qc).Value
                        If IsNumeric(cv) Then
                            If Abs(CDbl(cv)) > Abs(worstVal) Then worstVal = CDbl(cv)
                        End If
                    Next qc
                    If Abs(worstVal) < thresh Then
                        passCount = passCount + 1
                    Else
                        If Len(failures) > 0 Then failures = failures & ", "
                        failures = failures & hRowID & "=" & Format(worstVal, "#,##0.00")
                    End If
                ElseIf StrComp(hCheckType, "TextMatch", vbTextCompare) = 0 Then
                    Dim hColStart As String
                    hColStart = Trim(CStr(wsConfig.Cells(dr, HLCFG_COL_COLSTART).Value))
                    Dim checkCol As Long
                    If IsNumeric(hColStart) Then checkCol = CLng(hColStart) Else checkCol = 3
                    Dim cellText As String
                    cellText = Trim(CStr(hWs.Cells(hRow, checkCol).Value))
                    If StrComp(cellText, hGoodValue, vbTextCompare) = 0 Then
                        passCount = passCount + 1
                    Else
                        If Len(failures) > 0 Then failures = failures & ", "
                        failures = failures & hRowID & "=""" & cellText & """ (expected """ & hGoodValue & """)"
                    End If
                End If
            End If
        End If
        dr = dr + 1
    Loop

    If Len(failures) = 0 Then
        VerifyHealthAfterLoad = "All " & passCount & " health checks passed"
    Else
        VerifyHealthAfterLoad = passCount & " passed, FAIL: " & failures
    End If
    On Error GoTo 0
End Function


' =============================================================================
' ExportRegressionTabs
' Exports all tabs listed in regression_config (except Detail) to baseDir/regression_tabs/.
' Called during SaveGolden to capture formula tab outputs for regression comparison.
' =============================================================================
Public Sub ExportRegressionTabs(baseDir As String)
    On Error Resume Next
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    If wsConfig Is Nothing Then Exit Sub

    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_REGRESSION_CONFIG)
    If sr = 0 Then Exit Sub

    Dim regDir As String
    regDir = baseDir & "\regression_tabs"
    KernelSnapshot.EnsureDirectoryExists regDir

    Dim dr As Long
    dr = sr + 2
    Do While Len(Trim(CStr(wsConfig.Cells(dr, REGCFG_COL_TABNAME).Value))) > 0
        Dim tabName As String
        tabName = Trim(CStr(wsConfig.Cells(dr, REGCFG_COL_TABNAME).Value))
        ' Skip Detail (already exported as detail.csv)
        If StrComp(tabName, TAB_DETAIL, vbTextCompare) <> 0 Then
            Dim ws As Worksheet
            Set ws = Nothing
            Set ws = ThisWorkbook.Sheets(tabName)
            If Not ws Is Nothing Then
                ExportTabToCSV ws, regDir & "\" & SanitizeTabName(tabName) & ".csv"
            End If
        End If
        dr = dr + 1
    Loop
    On Error GoTo 0
End Sub


' =============================================================================
' CompareRegressionTabs
' Compares current tab values against golden regression_tabs/ CSVs.
' Returns a summary string and writes test results via the callback.
' =============================================================================
Public Function CompareRegressionTabs(goldenDir As String) As String
    On Error Resume Next
    Dim regDir As String
    regDir = goldenDir & "\regression_tabs"
    If Dir(regDir, vbDirectory) = "" Then
        CompareRegressionTabs = "No regression_tabs in golden (legacy snapshot)"
        Exit Function
    End If

    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    If wsConfig Is Nothing Then
        CompareRegressionTabs = "Config sheet not found"
        Exit Function
    End If

    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_REGRESSION_CONFIG)
    If sr = 0 Then
        CompareRegressionTabs = "No regression_config"
        Exit Function
    End If

    Dim totalTabs As Long
    totalTabs = 0
    Dim passedTabs As Long
    passedTabs = 0
    Dim failedTabs As String
    failedTabs = ""

    Dim dr As Long
    dr = sr + 2
    Do While Len(Trim(CStr(wsConfig.Cells(dr, REGCFG_COL_TABNAME).Value))) > 0
        Dim tabName As String
        tabName = Trim(CStr(wsConfig.Cells(dr, REGCFG_COL_TABNAME).Value))
        ' Skip Detail (compared separately)
        If StrComp(tabName, TAB_DETAIL, vbTextCompare) <> 0 Then
            Dim tolerance As Double
            Dim tolStr As String
            tolStr = Trim(CStr(wsConfig.Cells(dr, REGCFG_COL_TOLERANCE).Value))
            If IsNumeric(tolStr) And Len(tolStr) > 0 Then tolerance = CDbl(tolStr) Else tolerance = 0.01

            Dim csvPath As String
            csvPath = regDir & "\" & SanitizeTabName(tabName) & ".csv"
            If Dir(csvPath) <> "" Then
                totalTabs = totalTabs + 1
                Dim ws As Worksheet
                Set ws = Nothing
                Set ws = ThisWorkbook.Sheets(tabName)
                If Not ws Is Nothing Then
                    Dim mismatches As Long
                    mismatches = CompareTabToCSV(ws, csvPath, tolerance)
                    If mismatches = 0 Then
                        passedTabs = passedTabs + 1
                    Else
                        If Len(failedTabs) > 0 Then failedTabs = failedTabs & ", "
                        failedTabs = failedTabs & tabName & "(" & mismatches & ")"
                    End If
                End If
            End If
        End If
        dr = dr + 1
    Loop

    If totalTabs = 0 Then
        CompareRegressionTabs = "No regression tab CSVs found"
    ElseIf Len(failedTabs) = 0 Then
        CompareRegressionTabs = "All " & passedTabs & " tabs passed"
    Else
        CompareRegressionTabs = passedTabs & "/" & totalTabs & " passed. FAIL: " & failedTabs
    End If
    On Error GoTo 0
End Function


' CompareTabToCSV -- compares a worksheet values to a CSV file, returns mismatch count
Private Function CompareTabToCSV(ws As Worksheet, csvPath As String, tolerance As Double) As Long
    CompareTabToCSV = 0
    On Error Resume Next

    ' Read golden CSV
    Dim fileNum As Integer
    fileNum = FreeFile
    Dim fc As String
    Open csvPath For Binary Access Read As #fileNum
    Dim fSize As Long: fSize = LOF(fileNum)
    If fSize = 0 Then Close #fileNum: Exit Function
    fc = Space$(fSize)
    Get #fileNum, , fc
    Close #fileNum
    fc = Replace(fc, vbCrLf, vbLf)
    fc = Replace(fc, vbCr, vbLf)
    Dim lines() As String
    lines = Split(fc, vbLf)

    ' Count golden rows/cols
    Dim gRows As Long: gRows = 0
    Dim gCols As Long: gCols = 0
    Dim li As Long
    For li = 0 To UBound(lines)
        If Len(Trim(lines(li))) > 0 Then
            gRows = gRows + 1
            Dim cc As Long
            cc = Len(lines(li)) - Len(Replace(lines(li), ",", ""))
            If cc + 1 > gCols Then gCols = cc + 1
        End If
    Next li
    If gRows = 0 Then Exit Function

    ' Read current tab values
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1 > lastRow Then
        lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    End If
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1 > lastCol Then
        lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    End If
    If lastRow < 1 Or lastCol < 1 Then Exit Function

    Dim currentData As Variant
    currentData = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value

    ' Parse golden and compare
    Dim mismatches As Long: mismatches = 0
    Dim rowIdx As Long: rowIdx = 0
    For li = 0 To UBound(lines)
        If Len(Trim(lines(li))) > 0 Then
            rowIdx = rowIdx + 1
            If rowIdx > lastRow Then
                mismatches = mismatches + 1
            Else
                Dim fields() As String
                fields = ParseCSVLine(lines(li))
                Dim colIdx As Long
                For colIdx = 0 To UBound(fields)
                    If colIdx + 1 <= lastCol Then
                        Dim gVal As String: gVal = fields(colIdx)
                        Dim cVal As Variant: cVal = currentData(rowIdx, colIdx + 1)
                        ' Flag error values as automatic mismatches
                        Dim cStr2 As String
                        If IsError(cVal) Then
                            mismatches = mismatches + 1
                            GoTo NextCompareCol
                        End If
                        cStr2 = ""
                        If Not IsEmpty(cVal) Then cStr2 = CStr(cVal)
                        If Left(gVal, 1) = "#" Or Left(cStr2, 1) = "#" Then
                            ' Error values -- always a mismatch
                            mismatches = mismatches + 1
                            GoTo NextCompareCol
                        End If
                        If IsNumeric(gVal) And Len(gVal) > 0 And IsNumeric(cVal) Then
                            If Abs(CDbl(gVal) - CDbl(cVal)) > tolerance Then
                                mismatches = mismatches + 1
                            End If
                        Else
                            ' String comparison using cStr2 (already set above)
                            Dim gStr As String: gStr = gVal
                            If StrComp(gStr, cStr2, vbBinaryCompare) <> 0 Then
                                If Len(gStr) > 0 Or Len(cStr2) > 0 Then
                                    mismatches = mismatches + 1
                                End If
                            End If
                        End If
NextCompareCol:
                    End If
                Next colIdx
            End If
        End If
    Next li
    CompareTabToCSV = mismatches
    On Error GoTo 0
End Function
