Attribute VB_Name = "KernelFormHelpers"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' KernelFormHelpers.bas
' Shared utilities for VBA UserForms in the RDK project.
' Provides timestamp formatting, file counting, JSON field reading,
' and ListBox population helpers for Snapshot/Compare explorers.

' ---------------------------------------------------------------------------
' 1. AppendTimestamp
' ---------------------------------------------------------------------------
Public Function AppendTimestamp(baseName As String) As String
    Dim dt As Date
    dt = Now()

    Dim y As String, m As String, d As String
    Dim h As String, mi As String, s As String

    y = CStr(Year(dt))
    m = Right$("0" & CStr(Month(dt)), 2)
    d = Right$("0" & CStr(Day(dt)), 2)
    h = Right$("0" & CStr(Hour(dt)), 2)
    mi = Right$("0" & CStr(Minute(dt)), 2)
    s = Right$("0" & CStr(Second(dt)), 2)

    AppendTimestamp = baseName & "_" & y & m & d & "_" & h & mi & s
End Function

' ---------------------------------------------------------------------------
' 2. GetProjectRootPublic
' ---------------------------------------------------------------------------
Public Function GetProjectRootPublic() As String
    GetProjectRootPublic = ThisWorkbook.Path & "\.."
End Function

' ---------------------------------------------------------------------------
' 3. CountFilesInFolder
' ---------------------------------------------------------------------------
Public Function CountFilesInFolder(folderPath As String) As Long
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        CountFilesInFolder = 0
        Exit Function
    End If

    Dim fld As Object
    Set fld = fso.GetFolder(folderPath)
    CountFilesInFolder = fld.Files.Count

    Set fld = Nothing
    Set fso = Nothing
End Function

' ---------------------------------------------------------------------------
' 4. ReadJsonField
' ---------------------------------------------------------------------------
Public Function ReadJsonField(jsonPath As String, fieldName As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(jsonPath) Then
        ReadJsonField = ""
        Set fso = Nothing
        Exit Function
    End If
    Set fso = Nothing

    Dim fileNum As Integer
    Dim lineText As String
    Dim searchKey As String
    Dim posColon As Long
    Dim posQuote1 As Long
    Dim posQuote2 As Long

    searchKey = """" & fieldName & """"

    fileNum = FreeFile
    Open jsonPath For Input As #fileNum

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText

        If InStr(1, lineText, searchKey, vbTextCompare) > 0 Then
            ' Find the colon after the field name
            posColon = InStr(1, lineText, ":")
            If posColon > 0 Then
                ' Find the opening quote of the value
                posQuote1 = InStr(posColon + 1, lineText, """")
                If posQuote1 > 0 Then
                    ' Find the closing quote of the value
                    posQuote2 = InStr(posQuote1 + 1, lineText, """")
                    If posQuote2 > posQuote1 Then
                        ReadJsonField = Mid$(lineText, posQuote1 + 1, posQuote2 - posQuote1 - 1)
                        Close #fileNum
                        Exit Function
                    End If
                End If
            End If
        End If
    Loop

    Close #fileNum
    ReadJsonField = ""
End Function

' ---------------------------------------------------------------------------
' 5. GetItemCreatedDate
' ---------------------------------------------------------------------------
Public Function GetItemCreatedDate(folderPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim metaPath As String
    metaPath = folderPath & "\metadata.json"

    If fso.FileExists(metaPath) Then
        Set fso = Nothing
        Dim val1 As String
        val1 = ReadJsonField(metaPath, "created")
        If Len(val1) > 0 Then
            GetItemCreatedDate = val1
            Exit Function
        End If
    End If

    Dim manPath As String
    manPath = folderPath & "\manifest.json"

    If fso.FileExists(manPath) Then
        Set fso = Nothing
        Dim val2 As String
        val2 = ReadJsonField(manPath, "created")
        If Len(val2) > 0 Then
            GetItemCreatedDate = val2
            Exit Function
        End If
    End If

    Set fso = Nothing
    GetItemCreatedDate = "unknown"
End Function

' ---------------------------------------------------------------------------
' 6. PopulateSnapshotListBox
' ---------------------------------------------------------------------------
Public Sub PopulateSnapshotListBox(lb As Object, Optional sortByDate As Boolean = False, Optional sortAsc As Boolean = True)
    lb.Clear
    lb.ColumnCount = 7
    lb.ColumnWidths = "160;110;50;120;35;50;35"

    Dim names() As String
    names = KernelSnapshot.ListSnapshots()

    Dim cnt As Long
    cnt = 0
    Dim i As Long
    For i = LBound(names) To UBound(names)
        If Len(names(i)) > 0 Then cnt = cnt + 1
    Next i

    If cnt = 0 Then Exit Sub

    ' Build parallel arrays for sorting
    Dim arrName() As String
    Dim arrDate() As String
    Dim arrElapsed() As String
    Dim arrDesc() As String
    Dim arrStale() As String
    Dim arrStatus() As String
    Dim arrFiles() As String
    ReDim arrName(1 To cnt)
    ReDim arrDate(1 To cnt)
    ReDim arrElapsed(1 To cnt)
    ReDim arrDesc(1 To cnt)
    ReDim arrStale(1 To cnt)
    ReDim arrStatus(1 To cnt)
    ReDim arrFiles(1 To cnt)

    Dim basePath As String
    Dim itemPath As String
    Dim manPath As String
    Dim idx As Long: idx = 0

    basePath = GetProjectRootPublic() & "\snapshots\"

    For i = LBound(names) To UBound(names)
        If Len(names(i)) > 0 Then
            idx = idx + 1
            itemPath = basePath & names(i)
            manPath = itemPath & "\manifest.json"
            arrName(idx) = names(i)
            arrDate(idx) = GetItemCreatedDate(itemPath)
            arrElapsed(idx) = ReadJsonField(manPath, "elapsedSeconds")
            Dim rawDesc As String
            rawDesc = ReadJsonField(manPath, "description")
            If Len(rawDesc) > 30 Then rawDesc = Left(rawDesc, 30)
            arrDesc(idx) = rawDesc
            Dim rawStale As String
            rawStale = ReadJsonField(manPath, "resultsStale")
            If StrComp(rawStale, "true", vbTextCompare) = 0 Then
                arrStale(idx) = "Y"
            ElseIf StrComp(rawStale, "false", vbTextCompare) = 0 Then
                arrStale(idx) = "N"
            Else
                arrStale(idx) = "-"
            End If
            arrStatus(idx) = ReadJsonField(manPath, "status")
            arrFiles(idx) = CStr(CountFilesInFolder(itemPath))
        End If
    Next i

    ' Sort (bubble sort -- small N expected)
    If cnt > 1 Then
        Dim j As Long
        Dim swapped As Boolean
        Dim tmpStr As String
        Dim doSwap As Boolean
        For i = 1 To cnt - 1
            swapped = False
            For j = 1 To cnt - i
                If sortByDate Then
                    If sortAsc Then
                        doSwap = (arrDate(j) > arrDate(j + 1))
                    Else
                        doSwap = (arrDate(j) < arrDate(j + 1))
                    End If
                Else
                    If sortAsc Then
                        doSwap = (LCase(arrName(j)) > LCase(arrName(j + 1)))
                    Else
                        doSwap = (LCase(arrName(j)) < LCase(arrName(j + 1)))
                    End If
                End If
                If doSwap Then
                    tmpStr = arrName(j): arrName(j) = arrName(j + 1): arrName(j + 1) = tmpStr
                    tmpStr = arrDate(j): arrDate(j) = arrDate(j + 1): arrDate(j + 1) = tmpStr
                    tmpStr = arrElapsed(j): arrElapsed(j) = arrElapsed(j + 1): arrElapsed(j + 1) = tmpStr
                    tmpStr = arrDesc(j): arrDesc(j) = arrDesc(j + 1): arrDesc(j + 1) = tmpStr
                    tmpStr = arrStale(j): arrStale(j) = arrStale(j + 1): arrStale(j + 1) = tmpStr
                    tmpStr = arrStatus(j): arrStatus(j) = arrStatus(j + 1): arrStatus(j + 1) = tmpStr
                    tmpStr = arrFiles(j): arrFiles(j) = arrFiles(j + 1): arrFiles(j + 1) = tmpStr
                    swapped = True
                End If
            Next j
            If Not swapped Then Exit For
        Next i
    End If

    ' Populate ListBox
    For i = 1 To cnt
        lb.AddItem arrName(i)
        lb.List(lb.ListCount - 1, 1) = arrDate(i)
        lb.List(lb.ListCount - 1, 2) = arrElapsed(i)
        lb.List(lb.ListCount - 1, 3) = arrDesc(i)
        lb.List(lb.ListCount - 1, 4) = arrStale(i)
        lb.List(lb.ListCount - 1, 5) = arrStatus(i)
        lb.List(lb.ListCount - 1, 6) = arrFiles(i)
    Next i
End Sub

' ---------------------------------------------------------------------------
' 7. PopulateArchivedSnapshotListBox
' ---------------------------------------------------------------------------
Public Sub PopulateArchivedSnapshotListBox(lb As Object, Optional sortByDate As Boolean = False, Optional sortAsc As Boolean = True)
    lb.Clear
    lb.ColumnCount = 7
    lb.ColumnWidths = "160;110;50;120;35;50;35"

    Dim names() As String
    names = KernelSnapshot.ListArchivedSnapshots()

    Dim cnt As Long
    cnt = 0
    Dim i As Long
    For i = LBound(names) To UBound(names)
        If Len(names(i)) > 0 Then cnt = cnt + 1
    Next i

    If cnt = 0 Then Exit Sub

    ' Build parallel arrays for sorting
    Dim arrName() As String
    Dim arrDate() As String
    Dim arrElapsed() As String
    Dim arrDesc() As String
    Dim arrStale() As String
    Dim arrStatus() As String
    Dim arrFiles() As String
    ReDim arrName(1 To cnt)
    ReDim arrDate(1 To cnt)
    ReDim arrElapsed(1 To cnt)
    ReDim arrDesc(1 To cnt)
    ReDim arrStale(1 To cnt)
    ReDim arrStatus(1 To cnt)
    ReDim arrFiles(1 To cnt)

    Dim basePath As String
    Dim itemPath As String
    Dim manPath As String
    Dim idx As Long: idx = 0

    basePath = GetProjectRootPublic() & "\archive\snapshots\"

    For i = LBound(names) To UBound(names)
        If Len(names(i)) > 0 Then
            idx = idx + 1
            itemPath = basePath & names(i)
            manPath = itemPath & "\manifest.json"
            arrName(idx) = names(i)
            arrDate(idx) = GetItemCreatedDate(itemPath)
            arrElapsed(idx) = ReadJsonField(manPath, "elapsedSeconds")
            Dim rawDescA As String
            rawDescA = ReadJsonField(manPath, "description")
            If Len(rawDescA) > 30 Then rawDescA = Left(rawDescA, 30)
            arrDesc(idx) = rawDescA
            Dim rawStaleA As String
            rawStaleA = ReadJsonField(manPath, "resultsStale")
            If StrComp(rawStaleA, "true", vbTextCompare) = 0 Then
                arrStale(idx) = "Y"
            ElseIf StrComp(rawStaleA, "false", vbTextCompare) = 0 Then
                arrStale(idx) = "N"
            Else
                arrStale(idx) = "-"
            End If
            arrStatus(idx) = ReadJsonField(manPath, "status")
            arrFiles(idx) = CStr(CountFilesInFolder(itemPath))
        End If
    Next i

    ' Sort (bubble sort -- small N expected)
    If cnt > 1 Then
        Dim j As Long
        Dim swapped As Boolean
        Dim tmpStr As String
        Dim doSwap As Boolean
        For i = 1 To cnt - 1
            swapped = False
            For j = 1 To cnt - i
                If sortByDate Then
                    If sortAsc Then
                        doSwap = (arrDate(j) > arrDate(j + 1))
                    Else
                        doSwap = (arrDate(j) < arrDate(j + 1))
                    End If
                Else
                    If sortAsc Then
                        doSwap = (LCase(arrName(j)) > LCase(arrName(j + 1)))
                    Else
                        doSwap = (LCase(arrName(j)) < LCase(arrName(j + 1)))
                    End If
                End If
                If doSwap Then
                    tmpStr = arrName(j): arrName(j) = arrName(j + 1): arrName(j + 1) = tmpStr
                    tmpStr = arrDate(j): arrDate(j) = arrDate(j + 1): arrDate(j + 1) = tmpStr
                    tmpStr = arrElapsed(j): arrElapsed(j) = arrElapsed(j + 1): arrElapsed(j + 1) = tmpStr
                    tmpStr = arrDesc(j): arrDesc(j) = arrDesc(j + 1): arrDesc(j + 1) = tmpStr
                    tmpStr = arrStale(j): arrStale(j) = arrStale(j + 1): arrStale(j + 1) = tmpStr
                    tmpStr = arrStatus(j): arrStatus(j) = arrStatus(j + 1): arrStatus(j + 1) = tmpStr
                    tmpStr = arrFiles(j): arrFiles(j) = arrFiles(j + 1): arrFiles(j + 1) = tmpStr
                    swapped = True
                End If
            Next j
            If Not swapped Then Exit For
        Next i
    End If

    ' Populate ListBox
    For i = 1 To cnt
        lb.AddItem arrName(i)
        lb.List(lb.ListCount - 1, 1) = arrDate(i)
        lb.List(lb.ListCount - 1, 2) = arrElapsed(i)
        lb.List(lb.ListCount - 1, 3) = arrDesc(i)
        lb.List(lb.ListCount - 1, 4) = arrStale(i)
        lb.List(lb.ListCount - 1, 5) = arrStatus(i)
        lb.List(lb.ListCount - 1, 6) = arrFiles(i)
    Next i
End Sub

' ---------------------------------------------------------------------------
' 8. ShowSnapshotExplorer
' ---------------------------------------------------------------------------
Public Sub ShowSnapshotExplorer()
    SnapshotExplorer.Show vbModal
End Sub

' ---------------------------------------------------------------------------
' 9. ShowCompareExplorer
' ---------------------------------------------------------------------------
Public Sub ShowCompareExplorer()
    CompareExplorer.Show vbModal
End Sub

' ---------------------------------------------------------------------------
' ShowWorkspaceExplorer
' Single entry point for Save/Load. Reuses SnapshotExplorer form.
' ---------------------------------------------------------------------------
Public Sub ShowWorkspaceExplorer()
    WorkspaceExplorer.Show vbModal
End Sub

' ---------------------------------------------------------------------------
' ShowReportExplorer
' Opens the Report Explorer form for PDF/Print export.
' ---------------------------------------------------------------------------
Public Sub ShowReportExplorer()
    ReportExplorer.Show vbModal
End Sub

' ---------------------------------------------------------------------------
' PrintTabName
' Demo button handler: shows a MsgBox with the active sheet's tab name.
' Demonstrates that button_config can place buttons on any tab.
' ---------------------------------------------------------------------------
Public Sub PrintTabName()
    MsgBox "Current tab: " & ActiveSheet.Name, vbInformation, "RDK"
End Sub

' ---------------------------------------------------------------------------
' ListAllTabs
' Creates a new sheet listing all tab names with metadata from tab_registry.
' ---------------------------------------------------------------------------
Public Sub ListAllTabs()
    Application.ScreenUpdating = False
    On Error Resume Next

    ' Delete existing Tab List sheet if present
    Application.DisplayAlerts = False
    Dim existing As Worksheet
    Set existing = ThisWorkbook.Sheets("Tab List")
    If Not existing Is Nothing Then existing.Delete
    Application.DisplayAlerts = True

    ' Create new sheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "Tab List"
    On Error GoTo 0

    ' Headers
    ws.Cells(1, 1).Value = "Tab List"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.Color = RGB(31, 56, 100)
    ws.Cells(2, 1).Value = "Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(2, 1).Font.Color = RGB(128, 128, 128)

    ' Column headers
    ws.Cells(4, 1).Value = "#"
    ws.Cells(4, 2).Value = "Tab Name"
    ws.Cells(4, 3).Value = "Type"
    ws.Cells(4, 4).Value = "Category"
    ws.Cells(4, 5).Value = "Visible"
    ws.Cells(4, 6).Value = "Sort Order"
    ws.Cells(4, 7).Value = "Description"
    ws.Cells(4, 8).Value = "In Registry"
    ws.Range("A4:H4").Font.Bold = True
    ws.Range("A4:H4").Interior.Color = RGB(31, 56, 100)
    ws.Range("A4:H4").Font.Color = RGB(255, 255, 255)

    ' Build registry lookup dictionary from Config sheet
    Dim regDict As Object
    Set regDict = CreateObject("Scripting.Dictionary")
    regDict.CompareMode = vbTextCompare

    Dim wsConfig As Worksheet
    Set wsConfig = Nothing
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    On Error GoTo 0

    If Not wsConfig Is Nothing Then
        Dim sr As Long
        sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_TAB_REGISTRY)
        If sr > 0 Then
            Dim dr As Long
            dr = sr + 2
            Do While Len(Trim(CStr(wsConfig.Cells(dr, TREG_COL_TABNAME).Value))) > 0
                Dim regName As String
                regName = Trim(CStr(wsConfig.Cells(dr, TREG_COL_TABNAME).Value))
                If Not regDict.Exists(regName) Then
                    ' Store: Type, Category, Visible, SortOrder, Description
                    regDict.Add regName, Array( _
                        Trim(CStr(wsConfig.Cells(dr, 2).Value)), _
                        Trim(CStr(wsConfig.Cells(dr, 3).Value)), _
                        Trim(CStr(wsConfig.Cells(dr, TREG_COL_VISIBLE).Value)), _
                        Trim(CStr(wsConfig.Cells(dr, TREG_COL_SORTORDER).Value)), _
                        Trim(CStr(wsConfig.Cells(dr, 7).Value)))
                End If
                dr = dr + 1
            Loop
        End If
    End If

    ' Iterate ACTUAL worksheets in the workbook
    Dim outRow As Long
    outRow = 5
    Dim idx As Long
    idx = 0
    Dim sht As Worksheet
    For Each sht In ThisWorkbook.Worksheets
        idx = idx + 1
        ws.Cells(outRow, 1).Value = idx
        ws.Cells(outRow, 2).Value = sht.Name

        ' Look up metadata from registry
        If regDict.Exists(sht.Name) Then
            Dim meta As Variant
            meta = regDict(sht.Name)
            ws.Cells(outRow, 3).Value = meta(0)   ' Type
            ws.Cells(outRow, 4).Value = meta(1)   ' Category
            ws.Cells(outRow, 5).Value = meta(2)   ' Visible
            ws.Cells(outRow, 6).Value = meta(3)   ' SortOrder
            ws.Cells(outRow, 7).Value = meta(4)   ' Description
            ws.Cells(outRow, 8).Value = "Yes"
            ws.Cells(outRow, 8).Font.Color = RGB(0, 128, 0)
        Else
            ' Tab exists but not in registry
            If sht.Visible = xlSheetVisible Then
                ws.Cells(outRow, 5).Value = "Visible"
            Else
                ws.Cells(outRow, 5).Value = "Hidden"
            End If
            ws.Cells(outRow, 8).Value = "No"
            ws.Cells(outRow, 8).Font.Color = RGB(192, 0, 0)
        End If

        ' Alternating row color
        If idx Mod 2 = 0 Then
            ws.Range(ws.Cells(outRow, 1), ws.Cells(outRow, 8)).Interior.Color = RGB(242, 242, 242)
        End If
        outRow = outRow + 1
    Next sht

    ' Column widths
    ws.Columns(1).ColumnWidth = 4
    ws.Columns(2).ColumnWidth = 24
    ws.Columns(3).ColumnWidth = 10
    ws.Columns(4).ColumnWidth = 10
    ws.Columns(5).ColumnWidth = 10
    ws.Columns(6).ColumnWidth = 10
    ws.Columns(7).ColumnWidth = 50
    ws.Columns(8).ColumnWidth = 12

    Application.ScreenUpdating = True
    ws.Activate
    MsgBox "Tab List created: " & idx & " tabs.", vbInformation, "RDK"
End Sub

' ---------------------------------------------------------------------------
' 10. EnsureOutputDir
' Returns the full path to [project_root]\output\ and creates it if needed.
' Project root = parent of workbook directory (ThisWorkbook.Path\..\).
' ---------------------------------------------------------------------------
Public Function EnsureOutputDir() As String
    Dim wbPath As String
    wbPath = ThisWorkbook.Path
    Dim parentDir As String
    parentDir = Left(wbPath, InStrRev(wbPath, "\") - 1)
    Dim fullDir As String
    fullDir = parentDir & "\" & DIR_OUTPUT

    On Error Resume Next
    If Dir(fullDir, vbDirectory) = "" Then
        MkDir fullDir
    End If
    On Error GoTo 0

    EnsureOutputDir = fullDir
End Function

' ---------------------------------------------------------------------------
' 11. BuildInputHash
' ---------------------------------------------------------------------------
Public Function BuildInputHash() As String
    On Error GoTo HashError
    Dim wsInputs As Worksheet
    Set wsInputs = ThisWorkbook.Sheets(TAB_INPUTS)
    Dim paramCount As Long
    paramCount = KernelConfig.GetInputCount()
    If paramCount = 0 Then
        BuildInputHash = ""
        Exit Function
    End If
    Dim entityCount As Long
    entityCount = 0
    Dim ecol As Long
    For ecol = INPUT_ENTITY_START_COL To INPUT_ENTITY_START_COL + INPUT_MAX_ENTITIES - 1
        If Len(Trim(CStr(wsInputs.Cells(3, ecol).Value))) = 0 Then Exit For
        entityCount = entityCount + 1
    Next ecol
    If entityCount = 0 Then
        BuildInputHash = ""
        Exit Function
    End If
    Dim combined As String
    combined = ""
    Dim pidx As Long
    For pidx = 1 To paramCount
        Dim pRow As Long
        pRow = KernelConfig.GetInputRow(pidx)
        Dim eidx As Long
        For eidx = 0 To entityCount - 1
            combined = combined & CStr(wsInputs.Cells(pRow, INPUT_ENTITY_START_COL + eidx).Value) & "|"
        Next eidx
    Next pidx
    Dim tmpPath As String
    tmpPath = Environ("TEMP") & "\rdk_inputhash_" & Format(Timer * 1000, "0") & ".txt"
    Dim fileNum As Integer
    fileNum = FreeFile
    Open tmpPath For Output As #fileNum
    Print #fileNum, combined;
    Close #fileNum
    BuildInputHash = KernelSnapshot.ComputeSHA256(tmpPath)
    On Error Resume Next
    Kill tmpPath
    On Error GoTo 0
    Exit Function
HashError:
    BuildInputHash = ""
End Function

' ---------------------------------------------------------------------------
' 11. ReadRunStateValue
' ---------------------------------------------------------------------------
Public Function ReadRunStateValue(key As String) As String
    On Error GoTo ErrOut
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    Dim r As Long
    Dim lastRow As Long
    lastRow = wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).Row
    Dim inSection As Boolean
    inSection = False
    For r = 1 To lastRow
        Dim cv As String
        cv = CStr(wsConfig.Cells(r, 1).Value)
        If StrComp(cv, CFG_MARKER_RUN_STATE, vbTextCompare) = 0 Then
            inSection = True
        ElseIf Left(cv, 4) = "=== " And inSection Then
            Exit For
        ElseIf inSection Then
            If StrComp(cv, key, vbTextCompare) = 0 Then
                ReadRunStateValue = CStr(wsConfig.Cells(r, 2).Value)
                Exit Function
            End If
        End If
    Next r
    ReadRunStateValue = ""
    Exit Function
ErrOut:
    ReadRunStateValue = ""
End Function

' ---------------------------------------------------------------------------
' 12. WriteRunStateValue
' ---------------------------------------------------------------------------
Public Sub WriteRunStateValue(key As String, value As String)
    On Error Resume Next
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    Dim lastRow As Long
    lastRow = wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).Row
    Dim markerRow As Long
    markerRow = 0
    Dim r As Long
    For r = 1 To lastRow
        If StrComp(CStr(wsConfig.Cells(r, 1).Value), CFG_MARKER_RUN_STATE, vbTextCompare) = 0 Then
            markerRow = r
            Exit For
        End If
    Next r
    If markerRow = 0 Then
        markerRow = lastRow + 2
        wsConfig.Cells(markerRow, 1).NumberFormat = "@"
        wsConfig.Cells(markerRow, 1).Value = CFG_MARKER_RUN_STATE
    End If
    Dim inSection As Boolean
    inSection = True
    For r = markerRow + 1 To lastRow + 20
        Dim cv As String
        cv = CStr(wsConfig.Cells(r, 1).Value)
        If Left(cv, 4) = "=== " Then Exit For
        If StrComp(cv, key, vbTextCompare) = 0 Then
            wsConfig.Cells(r, 2).Value = value
            Exit Sub
        End If
        If Len(Trim(cv)) = 0 Then
            wsConfig.Cells(r, 1).Value = key
            wsConfig.Cells(r, 2).Value = value
            Exit Sub
        End If
    Next r
    wsConfig.Cells(r, 1).Value = key
    wsConfig.Cells(r, 2).Value = value
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' 13. IsResultsStale
' ---------------------------------------------------------------------------
Public Function IsResultsStale() As Boolean
    Dim sv As String
    sv = ReadRunStateValue(RS_KEY_STALE)
    If StrComp(sv, "TRUE", vbTextCompare) = 0 Then
        IsResultsStale = True
    ElseIf Len(sv) = 0 Then
        IsResultsStale = True
    Else
        IsResultsStale = False
    End If
End Function

' ---------------------------------------------------------------------------
' 14. CheckStalenessBeforeSave
' Returns: 0=Cancel, 1=Run&Save, 2=SaveAnyway
' ---------------------------------------------------------------------------
Public Function CheckStalenessBeforeSave() As Long
    Dim curHash As String
    curHash = BuildInputHash()
    Dim lastHash As String
    lastHash = ReadRunStateValue(RS_KEY_INPUT_HASH)
    If Len(curHash) > 0 And Len(lastHash) > 0 Then
        If StrComp(curHash, lastHash, vbTextCompare) = 0 Then
            Dim cfgCur As String
            cfgCur = KernelSnapshot.BuildConfigHash()
            Dim cfgLast As String
            cfgLast = ReadRunStateValue(RS_KEY_CONFIG_HASH)
            If Len(cfgCur) > 0 And Len(cfgLast) > 0 Then
                If StrComp(cfgCur, cfgLast, vbTextCompare) = 0 Then
                    CheckStalenessBeforeSave = 2
                    Exit Function
                End If
            Else
                CheckStalenessBeforeSave = 2
                Exit Function
            End If
        End If
    End If
    If Len(lastHash) = 0 And Len(curHash) = 0 Then
        CheckStalenessBeforeSave = 2
        Exit Function
    End If
    Dim elapsed As String
    elapsed = ReadRunStateValue(RS_KEY_TOTAL_ELAPSED)
    Dim threshold As String
    threshold = KernelConfig.GetScaleSetting(SCALE_LARGE_MODEL_SEC)
    Dim threshVal As Double
    If IsNumeric(threshold) And Len(threshold) > 0 Then
        threshVal = CDbl(threshold)
    Else
        threshVal = 60
    End If
    Dim warnMsg As String
    warnMsg = "Results may not reflect current inputs." & vbCrLf & vbCrLf
    If IsNumeric(elapsed) And Len(elapsed) > 0 Then
        If CDbl(elapsed) > threshVal Then
            warnMsg = warnMsg & "Last run took " & Format(CDbl(elapsed), "0.0") & "s (large model)." & vbCrLf & vbCrLf
        End If
    End If
    warnMsg = warnMsg & "Run model first, save anyway, or cancel?"
    Dim ans As Long
    ans = MsgBox(warnMsg, vbYesNoCancel Or vbQuestion Or vbDefaultButton2, "RDK -- Stale Results")
    Select Case ans
        Case vbYes: CheckStalenessBeforeSave = 1
        Case vbNo: CheckStalenessBeforeSave = 2
        Case Else: CheckStalenessBeforeSave = 0
    End Select
End Function

' ---------------------------------------------------------------------------
' 15. PatchManifestStaleAndElapsed
' ---------------------------------------------------------------------------
Public Sub PatchManifestStaleAndElapsed(manifestPath As String, isStale As Boolean, elapsed As String)
    On Error GoTo PatchErr
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(manifestPath) Then Exit Sub
    Dim fileNum As Integer
    fileNum = FreeFile
    Dim content As String
    Open manifestPath For Binary Access Read As #fileNum
    Dim fSize As Long
    fSize = LOF(fileNum)
    If fSize = 0 Then Close #fileNum: Exit Sub
    content = Space$(fSize)
    Get #fileNum, , content
    Close #fileNum
    Dim staleStr As String
    If isStale Then staleStr = "true" Else staleStr = "false"
    If Len(elapsed) = 0 Then elapsed = "0"
    Dim insertLine As String
    insertLine = "  ""resultsStale"": " & staleStr & "," & vbCrLf & _
                 "  ""elapsedSeconds"": """ & elapsed & ""","
    Dim insertPos As Long
    insertPos = InStr(1, content, """prngSeed""")
    If insertPos > 0 Then
        Dim lineEnd As Long
        lineEnd = InStr(insertPos, content, vbLf)
        If lineEnd > 0 Then
            content = Left(content, lineEnd) & insertLine & vbCrLf & Mid(content, lineEnd + 1)
        End If
    Else
        Dim lastBrace As Long
        lastBrace = InStrRev(content, "}")
        If lastBrace > 1 Then
            content = Left(content, lastBrace - 1) & "," & vbCrLf & insertLine & vbCrLf & "}"
        End If
    End If
    fileNum = FreeFile
    Open manifestPath For Output As #fileNum
    Print #fileNum, content;
    Close #fileNum
    Exit Sub
PatchErr:
    Debug.Print "PatchManifest error: " & Err.Description
End Sub

' ---------------------------------------------------------------------------
' 16. EditSnapshotDescription
' ---------------------------------------------------------------------------
Public Sub EditSnapshotDescription(snapshotName As String, newDesc As String)
    On Error GoTo EditErr
    Dim manPath As String
    manPath = GetProjectRootPublic() & "\snapshots\" & snapshotName & "\manifest.json"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(manPath) Then
        MsgBox "manifest.json not found.", vbExclamation, "RDK"
        Exit Sub
    End If
    Set fso = Nothing
    Dim fileNum As Integer
    fileNum = FreeFile
    Dim content As String
    Open manPath For Binary Access Read As #fileNum
    Dim fSize As Long
    fSize = LOF(fileNum)
    content = Space$(fSize)
    Get #fileNum, , content
    Close #fileNum
    Dim searchKey As String
    searchKey = """description"""
    Dim keyPos As Long
    keyPos = InStr(1, content, searchKey)
    If keyPos > 0 Then
        Dim colonPos As Long
        colonPos = InStr(keyPos, content, ":")
        If colonPos > 0 Then
            Dim q1 As Long
            q1 = InStr(colonPos, content, """")
            If q1 > 0 Then
                Dim q2 As Long
                q2 = InStr(q1 + 1, content, """")
                If q2 > q1 Then
                    content = Left(content, q1) & Replace(newDesc, """", "\""") & Mid(content, q2)
                End If
            End If
        End If
    End If
    fileNum = FreeFile
    Open manPath For Output As #fileNum
    Print #fileNum, content;
    Close #fileNum
    MsgBox "Description updated.", vbInformation, "RDK"
    Exit Sub
EditErr:
    KernelConfig.LogError SEV_ERROR, "KFH", "E-880", _
        "Error editing description: " & Err.Description, _
        "MANUAL BYPASS: Edit manifest.json description field directly."
    MsgBox "Error: " & Err.Description, vbCritical, "RDK"
End Sub

' ---------------------------------------------------------------------------
' 17. ShowSnapshots (moved from KernelSnapshot)
' ---------------------------------------------------------------------------
Public Sub ShowSnapshots()
    SnapshotExplorer.Show vbModal
End Sub


' ---------------------------------------------------------------------------
' 18. ToggleDevMode
' Toggles visibility of developer/system tabs and persists state.
' Called from Dashboard button. Shows MsgBox confirming new state (AP-53).
' Dev tabs: Detail, QuarterlySummary, ErrorLog, TestResults, CumulativeView,
'           Analysis, Summary, ProveIt, Exhibits, Charts.
' ---------------------------------------------------------------------------
Public Sub ToggleDevMode()
    On Error GoTo ErrHandler

    Dim devTabs As Variant
    devTabs = GetDevModeTabs()

    ' Read persisted state and toggle
    Dim curMode As String
    curMode = KernelConfig.GetDevMode()
    If Len(curMode) = 0 Then curMode = DEV_MODE_OFF

    Dim newMode As String
    If curMode = DEV_MODE_ON Then
        newMode = DEV_MODE_OFF
    Else
        newMode = DEV_MODE_ON
    End If

    Dim newState As XlSheetVisibility
    If newMode = DEV_MODE_ON Then
        newState = xlSheetVisible
    Else
        newState = xlSheetHidden
    End If

    Dim idx As Long
    For idx = LBound(devTabs) To UBound(devTabs)
        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(devTabs(idx))
        On Error GoTo ErrHandler
        If Not ws Is Nothing Then
            ws.Visible = newState
        End If
    Next idx

    ' Persist state
    KernelConfig.SetDevMode newMode

    ' Update Dashboard indicator and toggle button caption
    KernelTabs.UpdateDashboardDevMode

    ' Show/hide dev-only buttons on Dashboard based on new mode
    ToggleDevOnlyButtons newMode

    KernelConfig.LogError SEV_INFO, "KernelFormHelpers", "I-560", _
        "Dev mode toggled to " & newMode, ""

    MsgBox "Dev Mode: " & newMode, vbInformation, "RDK"
    Exit Sub

ErrHandler:
    MsgBox "Error toggling dev mode: " & Err.Description & vbCrLf & _
           "MANUAL BYPASS: Unhide tabs via right-click on tab bar.", _
           vbExclamation, "RDK"
End Sub


' ---------------------------------------------------------------------------
' 19. ToggleDevOnlyButtons
' Shows or hides dev-only Dashboard buttons based on dev mode state.
' Reads button_config from Config sheet to determine which ButtonIDs are
' DevOnly=TRUE, then sets shape visibility for shapes named "btn_<ID>".
' ---------------------------------------------------------------------------
Private Sub ToggleDevOnlyButtons(newMode As String)
    On Error Resume Next

    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets(TAB_DASHBOARD)
    If wsDash Is Nothing Then Exit Sub

    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    If wsConfig Is Nothing Then Exit Sub
    On Error GoTo 0

    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_BUTTON_CONFIG)
    If sr = 0 Then Exit Sub

    Dim showDev As Boolean
    showDev = (StrComp(newMode, DEV_MODE_ON, vbTextCompare) = 0)

    ' Walk button_config rows and toggle visibility of dev-only shapes
    Dim dr As Long
    dr = sr + 2
    Do While Len(Trim(CStr(wsConfig.Cells(dr, BTNCFG_COL_TABNAME).Value))) > 0
        Dim btnTab As String
        btnTab = Trim(CStr(wsConfig.Cells(dr, BTNCFG_COL_TABNAME).Value))
        Dim btnDevOnly As String
        btnDevOnly = Trim(CStr(wsConfig.Cells(dr, BTNCFG_COL_DEVONLY).Value))

        If StrComp(btnTab, TAB_DASHBOARD, vbTextCompare) = 0 And _
           StrComp(btnDevOnly, "TRUE", vbTextCompare) = 0 Then
            Dim btnID As String
            btnID = Trim(CStr(wsConfig.Cells(dr, BTNCFG_COL_ID).Value))
            Dim shapeName As String
            shapeName = "btn_" & btnID

            On Error Resume Next
            Dim shp As Shape
            Set shp = wsDash.Shapes(shapeName)
            If Not shp Is Nothing Then
                If showDev Then
                    shp.Visible = msoTrue
                Else
                    shp.Visible = msoFalse
                End If
            End If
            Set shp = Nothing
            On Error GoTo 0
        End If
        dr = dr + 1
    Loop

    ' Dev button shapes are toggled via Visible = msoTrue/msoFalse above.
    ' No row hiding -- rows stay visible, only button shapes toggle.
End Sub


' ---------------------------------------------------------------------------
' ShowConfigMsgBox
' Resolves a msgbox_config entry by ID, applies placeholder replacements,
' maps icon strings to VBA constants, and shows the MsgBox.
' Falls back to a generic MsgBox if the config entry is missing.
' Usage: ShowConfigMsgBox "RUN_COMPLETE", "{ENTITIES}", entityCount, "{ELAPSED}", elapsed
' ---------------------------------------------------------------------------
Public Sub ShowConfigMsgBox(msgID As String, ParamArray replacements() As Variant)
    Dim msg As Variant
    msg = KernelConfig.GetMsgBox(msgID)
    If IsEmpty(msg) Then
        ' Fallback: show generic message with the ID
        MsgBox msgID, vbInformation, "RDK"
        Exit Sub
    End If
    Dim txt As String: txt = CStr(msg(0))
    Dim ttl As String: ttl = CStr(msg(1))
    Dim icon As String: icon = CStr(msg(2))
    ' Replace \n with line breaks
    txt = Replace(txt, "\n", vbCrLf)
    ' Apply placeholder replacements (pairs: placeholder, value)
    Dim i As Long
    If UBound(replacements) >= 1 Then
        For i = LBound(replacements) To UBound(replacements) - 1 Step 2
            txt = Replace(txt, CStr(replacements(i)), CStr(replacements(i + 1)))
        Next i
    End If
    ' Map icon string to VBA constant
    Dim vbIcon As Long
    Select Case LCase(icon)
        Case "information": vbIcon = vbInformation
        Case "exclamation": vbIcon = vbExclamation
        Case "critical": vbIcon = vbCritical
        Case Else: vbIcon = vbInformation
    End Select
    MsgBox txt, vbIcon, ttl
End Sub


' ---------------------------------------------------------------------------
' ShowAssumptionManager
' Wrapper for KernelAssumptions.ShowAssumptionManager (Dashboard button target).
' ---------------------------------------------------------------------------
Public Sub ShowAssumptionManager()
    On Error Resume Next
    AssumptionExplorer.Show vbModal
    If Err.Number <> 0 Then
        Err.Clear
        ' Fallback to InputBox-based menu if form not available
        KernelAssumptions.ShowAssumptionManager
    End If
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' Model Lock -- state stored as workbook custom property "_ModelLocked"
' ---------------------------------------------------------------------------

' IsModelLocked -- checks current lock state
Public Function IsModelLocked() As Boolean
    On Error Resume Next
    Dim val As String
    val = ThisWorkbook.CustomDocumentProperties("_ModelLocked").Value
    IsModelLocked = (StrComp(val, "TRUE", vbTextCompare) = 0)
    On Error GoTo 0
End Function

' SetModelLocked -- sets lock state and applies/removes tab protection
Public Sub SetModelLocked(locked As Boolean)
    On Error Resume Next
    ' Store state as custom doc property
    Dim props As Object
    Set props = ThisWorkbook.CustomDocumentProperties
    Dim exists As Boolean: exists = False
    Dim p As Object
    For Each p In props
        If p.Name = "_ModelLocked" Then exists = True: Exit For
    Next p
    If exists Then
        props("_ModelLocked").Value = IIf(locked, "TRUE", "FALSE")
    Else
        props.Add "_ModelLocked", False, 4, IIf(locked, "TRUE", "FALSE")
    End If

    ' Protect/unprotect configured tabs
    Dim lockTabs As String
    lockTabs = KernelConfig.GetLockSetting("LockTabs")
    If Len(lockTabs) > 0 Then
        Dim tabs() As String
        tabs = Split(lockTabs, ",")
        Dim i As Long
        For i = 0 To UBound(tabs)
            Dim tabName As String: tabName = Trim$(tabs(i))
            Dim ws As Worksheet
            Set ws = Nothing
            Set ws = ThisWorkbook.Sheets(tabName)
            If Not ws Is Nothing Then
                If locked Then
                    ws.Protect UserInterfaceOnly:=False
                Else
                    ws.Unprotect
                End If
            End If
        Next i
    End If

    ' Update Dashboard visual
    UpdateLockVisual locked
    On Error GoTo 0
End Sub

' ToggleModelLock -- Dashboard button handler
Public Sub ToggleModelLock()
    Dim lockEnabled As String
    lockEnabled = KernelConfig.GetLockSetting("LockEnabled")
    If StrComp(lockEnabled, "TRUE", vbTextCompare) <> 0 Then
        MsgBox "Model lock is not enabled.", vbInformation, "RDK"
        Exit Sub
    End If

    If IsModelLocked() Then
        SetModelLocked False
        MsgBox "Model unlocked. You can now edit inputs and re-run.", _
               vbInformation, "RDK -- Unlocked"
    Else
        SetModelLocked True
        Dim msg As String
        msg = KernelConfig.GetLockSetting("LockMessage")
        If Len(msg) = 0 Then msg = "Model is now locked."
        MsgBox msg, vbInformation, "RDK -- Locked"
    End If
End Sub

' CheckLockGate -- returns True if action is blocked by lock.
' Shows MsgBox and returns True if blocked, False if allowed.
Public Function CheckLockGate(actionName As String) As Boolean
    CheckLockGate = False
    If Not IsModelLocked() Then Exit Function

    Dim lockActions As String
    lockActions = KernelConfig.GetLockSetting("LockActions")
    If Len(lockActions) = 0 Then Exit Function

    ' Check if this action is in the blocked list
    Dim actions() As String
    actions = Split(lockActions, ",")
    Dim i As Long
    For i = 0 To UBound(actions)
        If StrComp(Trim$(actions(i)), actionName, vbTextCompare) = 0 Then
            Dim msg As String
            msg = KernelConfig.GetLockSetting("LockMessage")
            If Len(msg) = 0 Then msg = "The model is locked. Please unlock to proceed."
            MsgBox msg, vbExclamation, "RDK -- Model Locked"
            CheckLockGate = True
            Exit Function
        End If
    Next i
End Function

' UpdateLockVisual -- updates Dashboard Run Model button and status text
Private Sub UpdateLockVisual(locked As Boolean)
    On Error Resume Next
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets(TAB_DASHBOARD)
    If wsDash Is Nothing Then Exit Sub

    ' Update lock status indicator (cell F1)
    If locked Then
        wsDash.Cells(11, 6).Value = "LOCKED"
        wsDash.Cells(11, 6).Font.Color = RGB(255, 255, 255)
        wsDash.Cells(11, 6).Font.Bold = True
        wsDash.Cells(11, 6).Font.Size = 10
        wsDash.Cells(11, 6).Interior.Color = RGB(192, 0, 0)
        wsDash.Cells(11, 6).HorizontalAlignment = xlCenter
    Else
        wsDash.Cells(11, 6).Value = ""
        wsDash.Cells(11, 6).Interior.ColorIndex = xlNone
    End If

    ' Update Run Model button appearance
    Dim shp As Object
    For Each shp In wsDash.Shapes
        If shp.Name = "btnRUN_MODEL" Or InStr(1, shp.OnAction, "RunModel", vbTextCompare) > 0 Then
            If locked Then
                shp.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = RGB(180, 180, 180)
            Else
                shp.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            End If
            Exit For
        End If
    Next shp
    On Error GoTo 0
End Sub

' RestoreLockVisual -- called at bootstrap/load to restore visual state
Public Sub RestoreLockVisual()
    If IsModelLocked() Then
        UpdateLockVisual True
    End If
End Sub

' ---------------------------------------------------------------------------
' ShowWorkspaceCompare (stub)
' Placeholder for workspace comparison feature.
' ---------------------------------------------------------------------------
Public Sub ShowWorkspaceCompare()
    MsgBox "Compare Workspaces: coming soon.", vbInformation, "RDK"
End Sub
