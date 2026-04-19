Attribute VB_Name = "KernelSnapshotIO"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelSnapshotIO.bas
' Purpose: Export/import functions for snapshot and workspace persistence.
'          Split from KernelSnapshot.bas (AD-09: module size limit).
'          Calls KernelSnapshot for shared helpers (GetProjectRoot,
'          EnsureDirectoryExists, GetInputsSheet, ParseCsvLine, etc.)
' =============================================================================


Public Sub ExportDetailToFile(csvPath As String)
    On Error GoTo ExpDetailErr
    Dim wsDetail As Worksheet
    Set wsDetail = ThisWorkbook.Sheets(TAB_DETAIL)
    On Error Resume Next
    wsDetail.Unprotect
    On Error GoTo ExpDetailErr
    Dim totalCols As Long
    totalCols = KernelConfig.GetColumnCount()
    If totalCols = 0 Then
        Err.Raise vbObjectError + 810, "KSnapIO", _
            "Column count is 0. Config may not be loaded. " & _
            "MANUAL BYPASS: Run Bootstrap or LoadConfig first."
    End If
    Dim lastRow As Long
    lastRow = wsDetail.Cells(wsDetail.Rows.Count, 1).End(xlUp).Row
    If lastRow < DETAIL_DATA_START_ROW Then lastRow = DETAIL_HEADER_ROW
    Dim fileNum As Integer
    fileNum = FreeFile
    Open csvPath For Output As #fileNum
    Dim headerLine As String
    headerLine = ""
    Dim colIdx As Long
    For colIdx = 1 To totalCols
        If colIdx > 1 Then headerLine = headerLine & ","
        headerLine = headerLine & """" & CStr(wsDetail.Cells(DETAIL_HEADER_ROW, colIdx).Value) & """"
    Next colIdx
    Print #fileNum, headerLine
    Dim rowIdx As Long
    For rowIdx = DETAIL_DATA_START_ROW To lastRow
        Dim dataLine As String
        dataLine = ""
        For colIdx = 1 To totalCols
            If colIdx > 1 Then dataLine = dataLine & ","
            Dim cellVal As Variant
            cellVal = wsDetail.Cells(rowIdx, colIdx).Value
            If IsNumeric(cellVal) And Not IsEmpty(cellVal) Then
                dataLine = dataLine & Format(CDbl(cellVal), "0.000000")
            Else
                dataLine = dataLine & """" & Replace(CStr(cellVal), """", """""") & """"
            End If
        Next colIdx
        Print #fileNum, dataLine
    Next rowIdx
    Close #fileNum
    On Error Resume Next
    wsDetail.Protect UserInterfaceOnly:=True
    On Error GoTo 0
    Exit Sub
ExpDetailErr:
    On Error Resume Next
    Close #fileNum
    wsDetail.Protect UserInterfaceOnly:=True
    On Error GoTo 0
    Err.Raise vbObjectError + 811, "KSnapIO", _
        "Export detail failed: " & Err.Description & ". " & _
        "MANUAL BYPASS: Copy Detail tab data to " & csvPath & " manually."
End Sub

Public Sub ExportInputsToFile(csvPath As String)
    On Error GoTo ExpInputErr
    Dim wsInputs As Worksheet
    Set wsInputs = KernelSnapshot.GetInputsSheet()
    Dim paramCount As Long
    paramCount = KernelConfig.GetInputCount()
    Dim entityCount As Long
    entityCount = DetectEntityCount()
    Dim fileNum As Integer
    fileNum = FreeFile
    Open csvPath For Output As #fileNum
    Dim headerLine As String
    headerLine = """Section"",""ParamName"""
    Dim entIdx As Long
    For entIdx = 1 To entityCount
        headerLine = headerLine & ",""Entity" & entIdx & """"
    Next entIdx
    Print #fileNum, headerLine
    Dim pidx As Long
    For pidx = 1 To paramCount
        Dim section As String
        section = KernelConfig.GetInputSection(pidx)
        Dim paramName As String
        paramName = KernelConfig.GetInputParam(pidx)
        Dim paramRow As Long
        paramRow = KernelConfig.GetInputRow(pidx)
        Dim dataLine As String
        dataLine = """" & section & """,""" & paramName & """"
        For entIdx = 1 To entityCount
            Dim cellVal As Variant
            cellVal = wsInputs.Cells(paramRow, INPUT_ENTITY_START_COL + entIdx - 1).Value
            If IsNumeric(cellVal) And Not IsEmpty(cellVal) Then
                dataLine = dataLine & "," & CStr(cellVal)
            Else
                dataLine = dataLine & ",""" & Replace(CStr(cellVal), """", """""") & """"
            End If
        Next entIdx
        Print #fileNum, dataLine
    Next pidx
    Close #fileNum
    Exit Sub
ExpInputErr:
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
    Err.Raise vbObjectError + 812, "KSnapIO", _
        "Export inputs failed: " & Err.Description & ". " & _
        "MANUAL BYPASS: Copy Inputs tab data to " & csvPath & " manually."
End Sub

Public Sub ExportSettingsToFile(csvPath As String)
    On Error GoTo ExpSettErr
    Dim fileNum As Integer
    fileNum = FreeFile
    Open csvPath For Output As #fileNum
    Print #fileNum, """Setting"",""Value"""
    Dim sk As Variant, sv As Variant
    sk = Array("TimeHorizon", "MaxEntities", "DefaultSummaryView")
    Dim si As Long
    For si = 0 To 2
        Print #fileNum, """" & sk(si) & """,""" & KernelConfig.GetSetting(CStr(sk(si))) & """"
    Next si
    sk = Array("DeterministicMode", "DefaultSeed", "FloatPrecision")
    For si = 0 To 2
        Print #fileNum, """" & sk(si) & """,""" & KernelConfig.GetReproSetting(CStr(sk(si))) & """"
    Next si
    If KernelRandom.IsInitialized() Then
        Print #fileNum, """PRNGSeed"",""" & KernelRandom.GetSeed() & """"
        Print #fileNum, """PRNGCallCount"",""" & KernelRandom.GetCallCount() & """"
    End If
    Close #fileNum
    Exit Sub
ExpSettErr:
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
    Err.Raise vbObjectError + 814, "KSnapIO", _
        "Export settings failed: " & Err.Description & ". " & _
        "MANUAL BYPASS: Manually create " & csvPath
End Sub

Public Sub ExportErrorLogToFile(csvPath As String)
    On Error GoTo ExpErrLogErr
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets(TAB_ERROR_LOG)
    Dim lastRow As Long
    lastRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
    Dim fileNum As Integer
    fileNum = FreeFile
    Open csvPath For Output As #fileNum
    Print #fileNum, """Timestamp"",""Severity"",""Source"",""Code"",""Message"",""Detail"""
    If lastRow >= 2 Then
        Dim rowIdx As Long
        For rowIdx = 2 To lastRow
            Dim dLine As String
            dLine = """" & Replace(CStr(wsLog.Cells(rowIdx, 1).Value), """", """""") & """"
            Dim c As Long
            For c = 2 To 6
                dLine = dLine & ",""" & Replace(CStr(wsLog.Cells(rowIdx, c).Value), """", """""") & """"
            Next c
            Print #fileNum, dLine
        Next rowIdx
    End If
    Close #fileNum
    Exit Sub
ExpErrLogErr:
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
    Err.Raise vbObjectError + 813, "KSnapIO", _
        "Export error log failed: " & Err.Description & ". " & _
        "MANUAL BYPASS: Copy ErrorLog tab to " & csvPath & " manually."
End Sub

Public Sub ImportDetailFromCsv(csvPath As String)
    Dim wsDetail As Worksheet
    Set wsDetail = ThisWorkbook.Sheets(TAB_DETAIL)
    On Error Resume Next
    wsDetail.Unprotect
    On Error GoTo 0
    Dim lines() As String
    lines = KernelSnapshot.ReadFileLinesFromPath(csvPath)
    If UBound(lines) < 1 Then GoTo DoneDetail
    Dim totalCols As Long
    totalCols = KernelConfig.GetColumnCount()
    ' Validate CSV column count vs current config
    Dim hdrFields() As String
    hdrFields = KernelSnapshot.ParseCsvLine(lines(0))
    Dim csvColCount As Long
    csvColCount = UBound(hdrFields) + 1
    If csvColCount <> totalCols Then
        KernelConfig.LogError SEV_WARN, "KSnapIO", "W-825", _
            "Column count mismatch: CSV has " & csvColCount & _
            " columns, config has " & totalCols & ". " & _
            "Data may be truncated or padded.", ""
    End If
    ' Use the larger of CSV vs config to avoid silent truncation
    Dim importCols As Long
    If csvColCount > totalCols Then
        importCols = csvColCount
    Else
        importCols = totalCols
    End If
    Dim lastRow As Long
    lastRow = wsDetail.Cells(wsDetail.Rows.Count, 1).End(xlUp).Row
    If lastRow >= DETAIL_DATA_START_ROW Then
        wsDetail.Range(wsDetail.Cells(DETAIL_DATA_START_ROW, 1), _
                       wsDetail.Cells(lastRow, importCols)).ClearContents
    End If
    Dim dataRowCount As Long
    dataRowCount = UBound(lines)
    If dataRowCount <= 0 Then GoTo DoneDetail
    Dim outputArr() As Variant
    ReDim outputArr(1 To dataRowCount, 1 To importCols)
    Dim rowIdx As Long
    For rowIdx = 1 To dataRowCount
        If Len(Trim(lines(rowIdx))) = 0 Then GoTo NextDLine
        Dim fields() As String
        fields = KernelSnapshot.ParseCsvLine(lines(rowIdx))
        Dim colIdx As Long
        For colIdx = 0 To UBound(fields)
            If colIdx < importCols Then
                If IsNumeric(fields(colIdx)) And Len(fields(colIdx)) > 0 Then
                    outputArr(rowIdx, colIdx + 1) = CDbl(fields(colIdx))
                Else
                    outputArr(rowIdx, colIdx + 1) = fields(colIdx)
                End If
            End If
        Next colIdx
NextDLine:
    Next rowIdx
    ' Write headers from CSV (preserves snapshot column names)
    Dim hIdx As Long
    For hIdx = 0 To UBound(hdrFields)
        If hIdx < importCols Then
            wsDetail.Cells(DETAIL_HEADER_ROW, hIdx + 1).Value = hdrFields(hIdx)
        End If
    Next hIdx
    wsDetail.Range(wsDetail.Cells(DETAIL_DATA_START_ROW, 1), _
                   wsDetail.Cells(DETAIL_DATA_START_ROW + dataRowCount - 1, importCols)).Value = outputArr
DoneDetail:
    On Error Resume Next
    wsDetail.Protect UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Public Sub ImportInputsFromCsv(csvPath As String)
    Dim wsInputs As Worksheet
    Set wsInputs = KernelSnapshot.GetInputsSheet()
    Dim lines() As String
    lines = KernelSnapshot.ReadFileLinesFromPath(csvPath)
    If UBound(lines) < 1 Then Exit Sub
    Dim lineIdx As Long
    For lineIdx = 1 To UBound(lines)
        If Len(Trim(lines(lineIdx))) = 0 Then GoTo NextIPLine
        Dim fields() As String
        fields = KernelSnapshot.ParseCsvLine(lines(lineIdx))
        If UBound(fields) < 1 Then GoTo NextIPLine
        Dim section As String
        section = fields(0)
        Dim paramName As String
        paramName = fields(1)
        Dim paramCount As Long
        paramCount = KernelConfig.GetInputCount()
        Dim pidx As Long
        For pidx = 1 To paramCount
            If StrComp(KernelConfig.GetInputSection(pidx), section, vbTextCompare) = 0 And _
               StrComp(KernelConfig.GetInputParam(pidx), paramName, vbTextCompare) = 0 Then
                Dim targetRow As Long
                targetRow = KernelConfig.GetInputRow(pidx)
                Dim entIdx As Long
                For entIdx = 2 To UBound(fields)
                    wsInputs.Cells(targetRow, INPUT_ENTITY_START_COL + entIdx - 2).Value = fields(entIdx)
                Next entIdx
                Exit For
            End If
        Next pidx
NextIPLine:
    Next lineIdx
End Sub

Public Sub ImportErrorLogFromCsv(csvPath As String)
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets(TAB_ERROR_LOG)
    Dim lines() As String
    lines = KernelSnapshot.ReadFileLinesFromPath(csvPath)
    If UBound(lines) < 1 Then Exit Sub
    Dim lastRow As Long
    lastRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then
        wsLog.Range(wsLog.Cells(2, 1), wsLog.Cells(lastRow, 6)).ClearContents
        wsLog.Range(wsLog.Cells(2, 1), wsLog.Cells(lastRow, 6)).Interior.ColorIndex = xlNone
        wsLog.Range(wsLog.Cells(2, 1), wsLog.Cells(lastRow, 6)).Font.ColorIndex = xlAutomatic
        wsLog.Range(wsLog.Cells(2, 1), wsLog.Cells(lastRow, 6)).Font.Bold = False
    End If
    Dim lineIdx As Long
    For lineIdx = 1 To UBound(lines)
        If Len(Trim(lines(lineIdx))) = 0 Then GoTo NextELLine
        Dim fields() As String
        fields = KernelSnapshot.ParseCsvLine(lines(lineIdx))
        Dim outRow As Long
        outRow = lineIdx + 1
        Dim c As Long
        For c = 0 To UBound(fields)
            If c < 6 Then
                wsLog.Cells(outRow, c + 1).Value = fields(c)
            End If
        Next c
        If UBound(fields) >= 1 Then
            Dim sevText As String
            sevText = UCase(Trim(fields(1)))
            Select Case sevText
                Case "FATAL", "0"
                    wsLog.Cells(outRow, 2).Interior.Color = RGB(192, 0, 0)
                    wsLog.Cells(outRow, 2).Font.Color = RGB(255, 255, 255)
                    wsLog.Cells(outRow, 2).Font.Bold = True
                Case "ERROR", "1"
                    wsLog.Cells(outRow, 2).Interior.Color = RGB(255, 199, 206)
                    wsLog.Cells(outRow, 2).Font.Color = RGB(156, 0, 6)
                    wsLog.Cells(outRow, 2).Font.Bold = True
                Case "WARN", "WARNING", "2"
                    wsLog.Cells(outRow, 2).Interior.Color = RGB(255, 235, 156)
                    wsLog.Cells(outRow, 2).Font.Color = RGB(156, 101, 0)
                Case "INFO", "3"
                    wsLog.Cells(outRow, 2).Interior.Color = RGB(198, 239, 206)
                    wsLog.Cells(outRow, 2).Font.Color = RGB(0, 97, 0)
            End Select
        End If
NextELLine:
    Next lineIdx
End Sub

Public Function DetectEntityCount() As Long
    Dim wsInputs As Worksheet
    Set wsInputs = KernelSnapshot.GetInputsSheet()
    Dim paramCount As Long
    paramCount = KernelConfig.GetInputCount()
    Dim entityRow As Long
    entityRow = 0
    Dim pidx As Long
    For pidx = 1 To paramCount
        If StrComp(KernelConfig.GetInputParam(pidx), "EntityName", vbTextCompare) = 0 Then
            entityRow = KernelConfig.GetInputRow(pidx)
            Exit For
        End If
    Next pidx
    If entityRow = 0 Then
        DetectEntityCount = 0
        Exit Function
    End If
    Dim cnt As Long
    cnt = 0
    Dim colIdx As Long
    colIdx = INPUT_ENTITY_START_COL
    Do While colIdx < INPUT_ENTITY_START_COL + INPUT_MAX_ENTITIES
        If Trim(CStr(wsInputs.Cells(entityRow, colIdx).Value)) = "" Then Exit Do
        cnt = cnt + 1
        colIdx = colIdx + 1
    Loop
    DetectEntityCount = cnt
End Function

' =============================================================================
' SaveConfigToSnapshot
' Copies all CSV files from config/ into snapDir\config\ so the snapshot
' captures the full model definition (tabs, formulas, branding, validation).
' =============================================================================
Public Sub SaveConfigToSnapshot(snapDir As String)
    On Error Resume Next
    Dim root As String
    root = KernelSnapshot.GetProjectRoot()
    Dim configDir As String
    configDir = root & "\config"
    If Dir(configDir, vbDirectory) = "" Then Exit Sub

    Dim destDir As String
    destDir = snapDir & "\config"
    KernelSnapshot.EnsureDirectoryExists destDir

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim csvFile As Object
    For Each csvFile In fso.GetFolder(configDir).Files
        If StrComp(fso.GetExtensionName(csvFile.Name), "csv", vbTextCompare) = 0 Then
            fso.CopyFile csvFile.Path, destDir & "\" & csvFile.Name, True
        End If
    Next csvFile
    Set fso = Nothing
    On Error GoTo 0
End Sub


' =============================================================================
' RestoreConfigFromSnapshot
' Copies CSV files from snapDir\config\ back to config/. If the snapshot
' has no config/ subfolder (pre-enhancement snapshot), skips silently.
' Returns True if config was restored, False if no config in snapshot.
' =============================================================================
Public Function RestoreConfigFromSnapshot(snapDir As String) As Boolean
    RestoreConfigFromSnapshot = False
    On Error Resume Next
    Dim srcDir As String
    srcDir = snapDir & "\config"
    If Dir(srcDir, vbDirectory) = "" Then Exit Function

    Dim root As String
    root = KernelSnapshot.GetProjectRoot()
    Dim configDir As String
    configDir = root & "\config"
    KernelSnapshot.EnsureDirectoryExists configDir

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim csvFile As Object
    For Each csvFile In fso.GetFolder(srcDir).Files
        If StrComp(fso.GetExtensionName(csvFile.Name), "csv", vbTextCompare) = 0 Then
            fso.CopyFile csvFile.Path, configDir & "\" & csvFile.Name, True
        End If
    Next csvFile
    Set fso = Nothing
    RestoreConfigFromSnapshot = True
    On Error GoTo 0
End Function
