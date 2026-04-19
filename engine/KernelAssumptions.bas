Attribute VB_Name = "KernelAssumptions"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelAssumptions.bas
' Purpose: Assumptions Register -- renders assumptions_config.csv as a
'          formatted tab with hyperlinks, conditional formatting, and CRUD
'          management via InputBox-based flows.
' =============================================================================

' Column positions on the Assumptions Register tab
Private Const AR_COL_ID As Long = 1
Private Const AR_COL_CATEGORY As Long = 2
Private Const AR_COL_TAB As Long = 3
Private Const AR_COL_ROWID As Long = 4
Private Const AR_COL_DESC As Long = 5
Private Const AR_COL_RATIONALE As Long = 6
Private Const AR_COL_SOURCE As Long = 7
Private Const AR_COL_CONFIDENCE As Long = 8
Private Const AR_COL_SENSITIVITY As Long = 9
Private Const AR_COL_IMPACT As Long = 10
Private Const AR_COL_OWNER As Long = 11
Private Const AR_COL_REVIEWED As Long = 12
Private Const AR_COL_HISTORY As Long = 13

' Config sheet column positions (assumptions_config.csv)
Private Const ACFG_COL_ID As Long = 1
Private Const ACFG_COL_CATEGORY As Long = 2
Private Const ACFG_COL_TAB As Long = 3
Private Const ACFG_COL_ROWID As Long = 4
Private Const ACFG_COL_DESC As Long = 5
Private Const ACFG_COL_RATIONALE As Long = 6
Private Const ACFG_COL_SOURCE As Long = 7
Private Const ACFG_COL_CONFIDENCE As Long = 8
Private Const ACFG_COL_SENSITIVITY As Long = 9
Private Const ACFG_COL_IMPACT As Long = 10
Private Const ACFG_COL_OWNER As Long = 11
Private Const ACFG_COL_REVIEWED As Long = 12
Private Const ACFG_COL_HISTORY As Long = 13

Private Const TAB_ASSUMPTIONS_REGISTER As String = "Assumptions Register"

' Confidence/Sensitivity fill colors
Private Const CLR_GREEN As Long = &HC6EFCE   ' High confidence / Low sensitivity
Private Const CLR_YELLOW As Long = &HFFEB9C  ' Medium
Private Const CLR_RED As Long = &HFFC7CE     ' Low confidence / High sensitivity
Private Const CLR_SECTION As Long = &HD9E1F2  ' Section header fill
Private Const CLR_ARCHIVED As Long = &HD9D9D9 ' Archived row fill (grey)

' =============================================================================
' GenerateAssumptionsRegister
' Read assumptions_config from Config sheet, create/refresh the Assumptions
' Register tab with grouped rows, hyperlinks, and conditional formatting.
' =============================================================================
Public Sub GenerateAssumptionsRegister()
    On Error GoTo ErrHandler

    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)

    ' Find assumptions_config section on Config sheet
    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_ASSUMPTIONS_CONFIG)
    If sr = 0 Then
        KernelConfig.LogError SEV_WARN, "KernelAssumptions", "W-900", _
            "assumptions_config section not found on Config sheet. Skipping register generation.", _
            "MANUAL BYPASS: Ensure assumptions_config.csv is in config/ and re-run bootstrap."
        Exit Sub
    End If

    ' Read all assumption rows into arrays
    Dim dataRows() As Variant
    Dim rowCount As Long
    rowCount = 0
    Dim dr As Long
    dr = sr + 2  ' skip marker + header row

    ' Count rows first
    Do While wsConfig.Cells(dr + rowCount, 1).Value <> "" And _
             Left$(CStr(wsConfig.Cells(dr + rowCount, 1).Value), 3) <> "==="
        rowCount = rowCount + 1
    Loop

    If rowCount = 0 Then
        KernelConfig.LogError SEV_INFO, "KernelAssumptions", "I-900", _
            "No assumption entries found in assumptions_config.", ""
        Exit Sub
    End If

    ' Load into 2D array
    ReDim dataRows(1 To rowCount, 1 To 13)
    Dim i As Long, j As Long
    For i = 1 To rowCount
        For j = 1 To 13
            dataRows(i, j) = CStr(wsConfig.Cells(dr + i - 1, j).Value)
        Next j
    Next i

    ' Create or clear tab
    Dim ws As Worksheet
    Set ws = EnsureTab(TAB_ASSUMPTIONS_REGISTER)

    Application.ScreenUpdating = False

    ws.Cells.Clear
    ws.DisplayPageBreaks = False

    ' --- Summary block ---
    Dim r As Long
    r = 1

    ' Title
    ws.Cells(r, 1).Value = "Assumptions Register"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = 14
    ws.Cells(r, 1).Font.Color = RGB(31, 56, 100) ' 1F3864
    r = r + 1

    ' Counts
    Dim cntHigh As Long, cntMed As Long, cntLow As Long
    Dim cntTotal As Long
    cntTotal = 0
    For i = 1 To rowCount
        If Left$(dataRows(i, ACFG_COL_CATEGORY), 9) <> "ARCHIVED-" Then
            cntTotal = cntTotal + 1
            Select Case dataRows(i, ACFG_COL_CONFIDENCE)
                Case "High": cntHigh = cntHigh + 1
                Case "Medium": cntMed = cntMed + 1
                Case "Low": cntLow = cntLow + 1
            End Select
        End If
    Next i

    ws.Cells(r, 1).Value = "Total Active: " & cntTotal & _
        "   |   High Confidence: " & cntHigh & _
        "   |   Medium: " & cntMed & _
        "   |   Low: " & cntLow
    ws.Cells(r, 1).Font.Size = 10
    ws.Cells(r, 1).Font.Color = RGB(128, 128, 128)
    r = r + 1

    ' Category summary
    Dim cats As Object
    Set cats = CreateObject("Scripting.Dictionary")
    For i = 1 To rowCount
        Dim catName As String
        catName = dataRows(i, ACFG_COL_CATEGORY)
        If Left$(catName, 9) = "ARCHIVED-" Then catName = "ARCHIVED"
        If Not cats.Exists(catName) Then cats.Add catName, 0
        cats(catName) = cats(catName) + 1
    Next i

    Dim catSummary As String
    catSummary = "By Category: "
    Dim k As Variant
    For Each k In cats.Keys
        If CStr(k) <> "ARCHIVED" Then
            catSummary = catSummary & CStr(k) & " (" & cats(k) & ")  "
        End If
    Next k
    ws.Cells(r, 1).Value = catSummary
    ws.Cells(r, 1).Font.Size = 10
    ws.Cells(r, 1).Font.Color = RGB(128, 128, 128)
    r = r + 1

    r = r + 1 ' spacer

    ' --- Column headers ---
    Dim headers As Variant
    headers = Array("ID", "Category", "Tab", "Input", "Description", "Rationale", _
                    "Source", "Confidence", "Sensitivity", "Impact", "Owner", "Last Reviewed", "History")
    For j = 0 To UBound(headers)
        ws.Cells(r, j + 1).Value = headers(j)
        ws.Cells(r, j + 1).Font.Bold = True
        ws.Cells(r, j + 1).Interior.Color = RGB(31, 56, 100)  ' 1F3864
        ws.Cells(r, j + 1).Font.Color = RGB(255, 255, 255)
    Next j
    Dim headerRow As Long
    headerRow = r
    r = r + 1

    ' Freeze panes at header row (only if tab is visible/active)
    On Error Resume Next
    If ws.Visible = xlSheetVisible Then
        ws.Activate
        ActiveWindow.FreezePanes = False
        ws.Cells(r, 1).Select
        ActiveWindow.FreezePanes = True
    End If
    On Error GoTo ErrHandler

    ' Sort data by category (active first, then archived)
    ' Build sorted index array
    Dim sortedIdx() As Long
    ReDim sortedIdx(1 To rowCount)

    ' Collect unique categories in order (active, then ARCHIVED)
    Dim catOrder As Object
    Set catOrder = CreateObject("Scripting.Dictionary")
    For i = 1 To rowCount
        catName = dataRows(i, ACFG_COL_CATEGORY)
        If Left$(catName, 9) <> "ARCHIVED-" Then
            If Not catOrder.Exists(catName) Then catOrder.Add catName, catOrder.Count
        End If
    Next i
    ' Add ARCHIVED- categories at the end
    For i = 1 To rowCount
        catName = dataRows(i, ACFG_COL_CATEGORY)
        If Left$(catName, 9) = "ARCHIVED-" Then
            If Not catOrder.Exists(catName) Then catOrder.Add catName, catOrder.Count
        End If
    Next i

    ' Build sorted index by category order
    Dim sIdx As Long
    sIdx = 0
    Dim catKey As Variant
    For Each catKey In catOrder.Keys
        For i = 1 To rowCount
            If dataRows(i, ACFG_COL_CATEGORY) = CStr(catKey) Then
                sIdx = sIdx + 1
                sortedIdx(sIdx) = i
            End If
        Next i
    Next catKey

    ' --- Write data rows grouped by category ---
    Dim prevCat As String
    prevCat = ""
    Dim dataStartRow As Long
    dataStartRow = r

    For sIdx = 1 To rowCount
        i = sortedIdx(sIdx)
        catName = dataRows(i, ACFG_COL_CATEGORY)

        ' Section header on category change
        If catName <> prevCat Then
            Dim sectionLabel As String
            If Left$(catName, 9) = "ARCHIVED-" Then
                sectionLabel = "ARCHIVED"
            Else
                sectionLabel = catName
            End If
            ws.Cells(r, 1).Value = sectionLabel
            ws.Cells(r, 1).Font.Bold = True
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 13)).Interior.Color = CLR_SECTION
            r = r + 1
            prevCat = catName
        End If

        Dim isArchived As Boolean
        isArchived = (Left$(catName, 9) = "ARCHIVED-")

        ' Write fields
        ws.Cells(r, AR_COL_ID).Value = dataRows(i, ACFG_COL_ID)
        ws.Cells(r, AR_COL_CATEGORY).Value = dataRows(i, ACFG_COL_CATEGORY)
        ws.Cells(r, AR_COL_TAB).Value = dataRows(i, ACFG_COL_TAB)

        ' RowID with hyperlink if TabName + RowID are non-blank
        Dim tabNm As String, rowId As String
        tabNm = dataRows(i, ACFG_COL_TAB)
        rowId = dataRows(i, ACFG_COL_ROWID)
        If Len(tabNm) > 0 And Len(rowId) > 0 Then
            ' Find cell address for RowID on the target tab
            Dim linkAddr As String
            linkAddr = FindRowIDAddress(tabNm, rowId)
            If Len(linkAddr) > 0 Then
                ws.Hyperlinks.Add Anchor:=ws.Cells(r, AR_COL_ROWID), _
                    Address:="", SubAddress:="'" & tabNm & "'!" & linkAddr, _
                    TextToDisplay:=rowId
            Else
                ws.Cells(r, AR_COL_ROWID).Value = rowId
            End If
        Else
            ws.Cells(r, AR_COL_ROWID).Value = rowId
        End If

        ws.Cells(r, AR_COL_DESC).Value = dataRows(i, ACFG_COL_DESC)
        ws.Cells(r, AR_COL_RATIONALE).Value = dataRows(i, ACFG_COL_RATIONALE)
        ws.Cells(r, AR_COL_SOURCE).Value = dataRows(i, ACFG_COL_SOURCE)
        ws.Cells(r, AR_COL_CONFIDENCE).Value = dataRows(i, ACFG_COL_CONFIDENCE)
        ws.Cells(r, AR_COL_SENSITIVITY).Value = dataRows(i, ACFG_COL_SENSITIVITY)
        ws.Cells(r, AR_COL_IMPACT).Value = dataRows(i, ACFG_COL_IMPACT)
        ws.Cells(r, AR_COL_OWNER).Value = dataRows(i, ACFG_COL_OWNER)
        ws.Cells(r, AR_COL_REVIEWED).Value = dataRows(i, ACFG_COL_REVIEWED)
        ws.Cells(r, AR_COL_HISTORY).Value = dataRows(i, ACFG_COL_HISTORY)

        ' Conditional formatting: Confidence
        Select Case dataRows(i, ACFG_COL_CONFIDENCE)
            Case "High":   ws.Cells(r, AR_COL_CONFIDENCE).Interior.Color = CLR_GREEN
            Case "Medium": ws.Cells(r, AR_COL_CONFIDENCE).Interior.Color = CLR_YELLOW
            Case "Low":    ws.Cells(r, AR_COL_CONFIDENCE).Interior.Color = CLR_RED
        End Select

        ' Conditional formatting: Sensitivity (inverse)
        Select Case dataRows(i, ACFG_COL_SENSITIVITY)
            Case "High":   ws.Cells(r, AR_COL_SENSITIVITY).Interior.Color = CLR_RED
            Case "Medium": ws.Cells(r, AR_COL_SENSITIVITY).Interior.Color = CLR_YELLOW
            Case "Low":    ws.Cells(r, AR_COL_SENSITIVITY).Interior.Color = CLR_GREEN
        End Select

        ' Grey out archived rows
        If isArchived Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 13)).Interior.Color = CLR_ARCHIVED
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 13)).Font.Color = RGB(128, 128, 128)
        End If

        r = r + 1
    Next sIdx

    ' --- Column widths ---
    ws.Columns(AR_COL_ID).ColumnWidth = 8
    ws.Columns(AR_COL_CATEGORY).ColumnWidth = 12
    ws.Columns(AR_COL_TAB).ColumnWidth = 20
    ws.Columns(AR_COL_ROWID).ColumnWidth = 15
    ws.Columns(AR_COL_DESC).ColumnWidth = 40
    ws.Columns(AR_COL_RATIONALE).ColumnWidth = 40
    ws.Columns(AR_COL_SOURCE).ColumnWidth = 20
    ws.Columns(AR_COL_CONFIDENCE).ColumnWidth = 12
    ws.Columns(AR_COL_SENSITIVITY).ColumnWidth = 12
    ws.Columns(AR_COL_IMPACT).ColumnWidth = 30
    ws.Columns(AR_COL_OWNER).ColumnWidth = 10
    ws.Columns(AR_COL_REVIEWED).ColumnWidth = 12
    ws.Columns(AR_COL_HISTORY).ColumnWidth = 50

    ' Wrap text on Description, Rationale, History
    ws.Columns(AR_COL_DESC).WrapText = True
    ws.Columns(AR_COL_RATIONALE).WrapText = True
    ws.Columns(AR_COL_HISTORY).WrapText = True

    ' Auto-filter on header row
    If Not ws.AutoFilterMode Then
        ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, 13)).AutoFilter
    End If

    ' No gridlines (only if this tab is active)
    On Error Resume Next
    If ActiveSheet.Name = TAB_ASSUMPTIONS_REGISTER Then
        ActiveWindow.DisplayGridlines = False
    End If
    On Error GoTo ErrHandler

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    KernelConfig.LogError SEV_ERROR, "KernelAssumptions", "E-900", _
        "Failed to generate Assumptions Register: " & Err.Description, _
        "MANUAL BYPASS: Check assumptions_config.csv data on Config sheet. " & _
        "Verify tab exists in tab_registry. Re-run bootstrap if needed."
End Sub


' =============================================================================
' ShowAssumptionManager
' Main menu for Assumption CRUD operations via MsgBox.
' =============================================================================
Public Sub ShowAssumptionManager()
    Dim choice As VbMsgBoxResult
    Dim msg As String
    msg = "Assumption Manager" & vbCrLf & vbCrLf & _
          "1 - View Register" & vbCrLf & _
          "2 - Add New Assumption" & vbCrLf & _
          "3 - Edit Assumption" & vbCrLf & _
          "4 - Archive Assumption" & vbCrLf & _
          "5 - Review Stale Assumptions" & vbCrLf & _
          "6 - Clear All Assumptions" & vbCrLf & vbCrLf & _
          "Enter choice number in the next prompt."

    Dim inp As String
    inp = InputBox(msg, "Assumption Manager", "1")
    If Len(inp) = 0 Then Exit Sub

    Select Case inp
        Case "1"
            ViewRegister
        Case "2"
            AddAssumption
        Case "3"
            EditAssumption
        Case "4"
            ArchiveAssumption
        Case "5"
            ReviewStale
        Case "6"
            ClearAllAssumptions
        Case Else
            MsgBox "Invalid choice. Please enter 1-6.", vbExclamation, "Assumption Manager"
    End Select
End Sub


' =============================================================================
' ViewRegister
' Activate the Assumptions Register tab; regenerate if missing.
' =============================================================================
Private Sub ViewRegister()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_ASSUMPTIONS_REGISTER)
    On Error GoTo 0

    If ws Is Nothing Then
        GenerateAssumptionsRegister
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(TAB_ASSUMPTIONS_REGISTER)
        On Error GoTo 0
    End If

    If Not ws Is Nothing Then ws.Activate
End Sub


' =============================================================================
' AddAssumption
' Prompt user for all fields via InputBox sequence, write to Config sheet,
' and regenerate the register.
' =============================================================================
Public Sub AddAssumption()
    On Error GoTo ErrHandler

    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)

    ' Find next AssumptionID
    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_ASSUMPTIONS_CONFIG)
    If sr = 0 Then
        MsgBox "assumptions_config section not found on Config sheet." & vbCrLf & _
               "MANUAL BYPASS: Re-run bootstrap to load config.", _
               vbExclamation, "Add Assumption"
        Exit Sub
    End If

    Dim maxNum As Long
    maxNum = 0
    Dim dr As Long
    dr = sr + 2
    Do While wsConfig.Cells(dr, 1).Value <> "" And _
             Left$(CStr(wsConfig.Cells(dr, 1).Value), 3) <> "==="
        Dim idStr As String
        idStr = CStr(wsConfig.Cells(dr, ACFG_COL_ID).Value)
        If Left$(idStr, 2) = "A-" Then
            Dim idNum As Long
            idNum = 0
            On Error Resume Next
            idNum = CLng(Mid$(idStr, 3))
            On Error GoTo ErrHandler
            If idNum > maxNum Then maxNum = idNum
        End If
        dr = dr + 1
    Loop

    Dim nextId As String
    nextId = "A-" & Format(maxNum + 1, "000")

    ' Prompt for fields
    Dim inp As String

    inp = InputBox("Assumption ID:", "Add Assumption", nextId)
    If Len(inp) = 0 Then Exit Sub
    Dim newId As String: newId = inp

    inp = InputBox("Category:" & vbCrLf & _
        "(Staffing/Revenue/Capital/Underwriting/Expense/Investment/Regulatory/Technology)", _
        "Add Assumption", "")
    If Len(inp) = 0 Then Exit Sub
    Dim newCat As String: newCat = inp

    inp = InputBox("Tab Name (from tab registry, or blank for general):", _
        "Add Assumption", "")
    Dim newTab As String: newTab = inp

    inp = InputBox("Row ID (input row ID, or blank for general):", _
        "Add Assumption", "")
    Dim newRowId As String: newRowId = inp

    inp = InputBox("Description (what the assumption IS):", _
        "Add Assumption", "")
    If Len(inp) = 0 Then Exit Sub
    Dim newDesc As String: newDesc = inp

    inp = InputBox("Rationale (WHY this assumption):", _
        "Add Assumption", "")
    Dim newRat As String: newRat = inp

    inp = InputBox("Source:" & vbCrLf & _
        "(Management Estimate/Benchmark/LOI/Contract/Market Data/Regulatory)", _
        "Add Assumption", "Management estimate")
    If Len(inp) = 0 Then Exit Sub
    Dim newSrc As String: newSrc = inp

    inp = InputBox("Confidence (High/Medium/Low):", _
        "Add Assumption", "Medium")
    If Len(inp) = 0 Then Exit Sub
    Dim newConf As String: newConf = inp

    inp = InputBox("Sensitivity (High/Medium/Low):", _
        "Add Assumption", "Medium")
    If Len(inp) = 0 Then Exit Sub
    Dim newSens As String: newSens = inp

    inp = InputBox("Sensitivity Detail (brief impact explanation):", _
        "Add Assumption", "")
    Dim newImpact As String: newImpact = inp

    inp = InputBox("Owner:", "Add Assumption", "Ethan")
    If Len(inp) = 0 Then Exit Sub
    Dim newOwner As String: newOwner = inp

    ' Write to Config sheet at end of section
    Dim writeRow As Long
    writeRow = dr  ' dr is already one past the last data row

    Dim todayStr As String
    todayStr = Format(Date, "yyyy-mm-dd")

    wsConfig.Cells(writeRow, ACFG_COL_ID).NumberFormat = "@"
    wsConfig.Cells(writeRow, ACFG_COL_ID).Value = newId
    wsConfig.Cells(writeRow, ACFG_COL_CATEGORY).NumberFormat = "@"
    wsConfig.Cells(writeRow, ACFG_COL_CATEGORY).Value = newCat
    wsConfig.Cells(writeRow, ACFG_COL_TAB).NumberFormat = "@"
    wsConfig.Cells(writeRow, ACFG_COL_TAB).Value = newTab
    wsConfig.Cells(writeRow, ACFG_COL_ROWID).NumberFormat = "@"
    wsConfig.Cells(writeRow, ACFG_COL_ROWID).Value = newRowId
    wsConfig.Cells(writeRow, ACFG_COL_DESC).NumberFormat = "@"
    wsConfig.Cells(writeRow, ACFG_COL_DESC).Value = newDesc
    wsConfig.Cells(writeRow, ACFG_COL_RATIONALE).NumberFormat = "@"
    wsConfig.Cells(writeRow, ACFG_COL_RATIONALE).Value = newRat
    wsConfig.Cells(writeRow, ACFG_COL_SOURCE).NumberFormat = "@"
    wsConfig.Cells(writeRow, ACFG_COL_SOURCE).Value = newSrc
    wsConfig.Cells(writeRow, ACFG_COL_CONFIDENCE).NumberFormat = "@"
    wsConfig.Cells(writeRow, ACFG_COL_CONFIDENCE).Value = newConf
    wsConfig.Cells(writeRow, ACFG_COL_SENSITIVITY).NumberFormat = "@"
    wsConfig.Cells(writeRow, ACFG_COL_SENSITIVITY).Value = newSens
    wsConfig.Cells(writeRow, ACFG_COL_IMPACT).NumberFormat = "@"
    wsConfig.Cells(writeRow, ACFG_COL_IMPACT).Value = newImpact
    wsConfig.Cells(writeRow, ACFG_COL_OWNER).NumberFormat = "@"
    wsConfig.Cells(writeRow, ACFG_COL_OWNER).Value = newOwner
    wsConfig.Cells(writeRow, ACFG_COL_REVIEWED).NumberFormat = "@"
    wsConfig.Cells(writeRow, ACFG_COL_REVIEWED).Value = todayStr
    wsConfig.Cells(writeRow, ACFG_COL_HISTORY).NumberFormat = "@"
    wsConfig.Cells(writeRow, ACFG_COL_HISTORY).Value = todayStr & ": Created"

    ' Regenerate register
    GenerateAssumptionsRegister

    MsgBox "Assumption " & newId & " added successfully.", vbInformation, "Add Assumption"
    Exit Sub

ErrHandler:
    MsgBox "Failed to add assumption: " & Err.Description & vbCrLf & _
           "MANUAL BYPASS: Add the entry directly to the Config sheet under " & _
           "the assumptions_config section, then call GenerateAssumptionsRegister.", _
           vbExclamation, "Add Assumption"
End Sub


' =============================================================================
' EditAssumption
' Prompt for AssumptionID, show current values, update a field, append History.
' =============================================================================
Public Sub EditAssumption()
    On Error GoTo ErrHandler

    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)

    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_ASSUMPTIONS_CONFIG)
    If sr = 0 Then
        MsgBox "assumptions_config section not found on Config sheet." & vbCrLf & _
               "MANUAL BYPASS: Re-run bootstrap to load config.", _
               vbExclamation, "Edit Assumption"
        Exit Sub
    End If

    ' Prompt for ID
    Dim targetId As String
    targetId = InputBox("Enter Assumption ID to edit (e.g., A-007):", "Edit Assumption", "")
    If Len(targetId) = 0 Then Exit Sub

    ' Find row on Config sheet
    Dim foundRow As Long
    foundRow = FindAssumptionRow(wsConfig, sr, targetId)
    If foundRow = 0 Then
        MsgBox "Assumption " & targetId & " not found." & vbCrLf & _
               "MANUAL BYPASS: Check the Config sheet assumptions_config section.", _
               vbExclamation, "Edit Assumption"
        Exit Sub
    End If

    ' Show current values
    Dim curDesc As String, curRat As String
    curDesc = CStr(wsConfig.Cells(foundRow, ACFG_COL_DESC).Value)
    curRat = CStr(wsConfig.Cells(foundRow, ACFG_COL_RATIONALE).Value)
    MsgBox "Current values for " & targetId & ":" & vbCrLf & vbCrLf & _
           "Description: " & curDesc & vbCrLf & _
           "Rationale: " & curRat, vbInformation, "Edit Assumption"

    ' Prompt for field to update
    Dim fieldName As String
    fieldName = InputBox("Which field to update?" & vbCrLf & _
        "Description / Rationale / Source / Confidence / Sensitivity / SensitivityDetail / Owner", _
        "Edit Assumption", "")
    If Len(fieldName) = 0 Then Exit Sub

    ' Map field name to column
    Dim fieldCol As Long
    fieldCol = MapFieldToColumn(fieldName)
    If fieldCol = 0 Then
        MsgBox "Unknown field: " & fieldName, vbExclamation, "Edit Assumption"
        Exit Sub
    End If

    ' Get old value
    Dim oldVal As String
    oldVal = CStr(wsConfig.Cells(foundRow, fieldCol).Value)

    ' Prompt for new value
    Dim newVal As String
    newVal = InputBox("New value for " & fieldName & ":", "Edit Assumption", oldVal)
    If Len(newVal) = 0 Then Exit Sub
    If newVal = oldVal Then
        MsgBox "No change made.", vbInformation, "Edit Assumption"
        Exit Sub
    End If

    ' Update field
    wsConfig.Cells(foundRow, fieldCol).NumberFormat = "@"
    wsConfig.Cells(foundRow, fieldCol).Value = newVal

    ' Append to History
    Dim todayStr As String
    todayStr = Format(Date, "yyyy-mm-dd")
    Dim curHistory As String
    curHistory = CStr(wsConfig.Cells(foundRow, ACFG_COL_HISTORY).Value)
    wsConfig.Cells(foundRow, ACFG_COL_HISTORY).NumberFormat = "@"
    wsConfig.Cells(foundRow, ACFG_COL_HISTORY).Value = curHistory & "; " & _
        todayStr & ": " & fieldName & " changed from '" & oldVal & "' to '" & newVal & "'"

    ' Update LastReviewed
    wsConfig.Cells(foundRow, ACFG_COL_REVIEWED).NumberFormat = "@"
    wsConfig.Cells(foundRow, ACFG_COL_REVIEWED).Value = todayStr

    ' Regenerate
    GenerateAssumptionsRegister

    MsgBox "Assumption " & targetId & " updated.", vbInformation, "Edit Assumption"
    Exit Sub

ErrHandler:
    MsgBox "Failed to edit assumption: " & Err.Description & vbCrLf & _
           "MANUAL BYPASS: Edit directly on Config sheet under assumptions_config section.", _
           vbExclamation, "Edit Assumption"
End Sub


' =============================================================================
' ArchiveAssumption
' Prefix Category with "ARCHIVED-", append History, regenerate.
' =============================================================================
Public Sub ArchiveAssumption()
    On Error GoTo ErrHandler

    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)

    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_ASSUMPTIONS_CONFIG)
    If sr = 0 Then
        MsgBox "assumptions_config section not found." & vbCrLf & _
               "MANUAL BYPASS: Re-run bootstrap.", vbExclamation, "Archive Assumption"
        Exit Sub
    End If

    Dim targetId As String
    targetId = InputBox("Enter Assumption ID to archive (e.g., A-007):", "Archive Assumption", "")
    If Len(targetId) = 0 Then Exit Sub

    Dim foundRow As Long
    foundRow = FindAssumptionRow(wsConfig, sr, targetId)
    If foundRow = 0 Then
        MsgBox "Assumption " & targetId & " not found." & vbCrLf & _
               "MANUAL BYPASS: Check Config sheet.", vbExclamation, "Archive Assumption"
        Exit Sub
    End If

    ' Check if already archived
    Dim curCat As String
    curCat = CStr(wsConfig.Cells(foundRow, ACFG_COL_CATEGORY).Value)
    If Left$(curCat, 9) = "ARCHIVED-" Then
        MsgBox targetId & " is already archived.", vbInformation, "Archive Assumption"
        Exit Sub
    End If

    ' Confirm
    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Archive assumption " & targetId & "?" & vbCrLf & _
        "Description: " & CStr(wsConfig.Cells(foundRow, ACFG_COL_DESC).Value), _
        vbYesNo + vbQuestion, "Archive Assumption")
    If confirm <> vbYes Then Exit Sub

    ' Prefix category
    wsConfig.Cells(foundRow, ACFG_COL_CATEGORY).NumberFormat = "@"
    wsConfig.Cells(foundRow, ACFG_COL_CATEGORY).Value = "ARCHIVED-" & curCat

    ' Append History
    Dim todayStr As String
    todayStr = Format(Date, "yyyy-mm-dd")
    Dim curHistory As String
    curHistory = CStr(wsConfig.Cells(foundRow, ACFG_COL_HISTORY).Value)
    wsConfig.Cells(foundRow, ACFG_COL_HISTORY).NumberFormat = "@"
    wsConfig.Cells(foundRow, ACFG_COL_HISTORY).Value = curHistory & "; " & todayStr & ": Archived"

    ' Regenerate
    GenerateAssumptionsRegister

    MsgBox "Assumption " & targetId & " archived.", vbInformation, "Archive Assumption"
    Exit Sub

ErrHandler:
    MsgBox "Failed to archive assumption: " & Err.Description & vbCrLf & _
           "MANUAL BYPASS: Prefix Category with ARCHIVED- on Config sheet.", _
           vbExclamation, "Archive Assumption"
End Sub


' =============================================================================
' GetStaleAssumptions
' Find assumptions where LastReviewed > daysThreshold days ago.
' Returns comma-delimited list of IDs.
' =============================================================================
Public Function GetStaleAssumptions(Optional daysThreshold As Long = 90) As String
    On Error GoTo ErrHandler

    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)

    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_ASSUMPTIONS_CONFIG)
    If sr = 0 Then
        GetStaleAssumptions = ""
        Exit Function
    End If

    Dim staleList As String
    staleList = ""
    Dim dr As Long
    dr = sr + 2
    Dim today As Date
    today = Date

    Do While wsConfig.Cells(dr, 1).Value <> "" And _
             Left$(CStr(wsConfig.Cells(dr, 1).Value), 3) <> "==="
        ' Skip archived
        If Left$(CStr(wsConfig.Cells(dr, ACFG_COL_CATEGORY).Value), 9) <> "ARCHIVED-" Then
            Dim reviewed As String
            reviewed = CStr(wsConfig.Cells(dr, ACFG_COL_REVIEWED).Value)
            If Len(reviewed) > 0 Then
                Dim revDate As Date
                On Error Resume Next
                revDate = CDate(reviewed)
                If Err.Number = 0 Then
                    If DateDiff("d", revDate, today) > daysThreshold Then
                        If Len(staleList) > 0 Then staleList = staleList & ", "
                        staleList = staleList & CStr(wsConfig.Cells(dr, ACFG_COL_ID).Value)
                    End If
                End If
                Err.Clear
                On Error GoTo ErrHandler
            End If
        End If
        dr = dr + 1
    Loop

    GetStaleAssumptions = staleList
    Exit Function

ErrHandler:
    GetStaleAssumptions = ""
End Function


' =============================================================================
' ReviewStale
' Scan for stale assumptions (>90 days) and offer to mark as reviewed.
' =============================================================================
Private Sub ReviewStale()
    Dim staleIds As String
    staleIds = GetStaleAssumptions(90)

    If Len(staleIds) = 0 Then
        MsgBox "No stale assumptions found. All assumptions reviewed within 90 days.", _
               vbInformation, "Review Stale"
        Exit Sub
    End If

    Dim resp As VbMsgBoxResult
    resp = MsgBox("Stale assumptions (>90 days since review):" & vbCrLf & vbCrLf & _
                  staleIds & vbCrLf & vbCrLf & _
                  "Mark all as reviewed today?", _
                  vbYesNo + vbQuestion, "Review Stale Assumptions")

    If resp = vbYes Then
        MarkAllReviewed staleIds
        GenerateAssumptionsRegister
        MsgBox "Marked as reviewed: " & staleIds, vbInformation, "Review Stale"
    End If
End Sub


' =============================================================================
' ClearAllAssumptions
' Removes all assumption rows from the Config sheet and regenerates register.
' =============================================================================
Private Sub ClearAllAssumptions()
    On Error GoTo ErrHandler

    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)

    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_ASSUMPTIONS_CONFIG)
    If sr = 0 Then
        MsgBox "No assumptions_config section found.", vbInformation, "Clear Assumptions"
        Exit Sub
    End If

    ' Count rows
    Dim rowCount As Long
    rowCount = 0
    Dim dr As Long
    dr = sr + 2
    Do While wsConfig.Cells(dr + rowCount, 1).Value <> "" And _
             Left$(CStr(wsConfig.Cells(dr + rowCount, 1).Value), 3) <> "==="
        rowCount = rowCount + 1
    Loop

    If rowCount = 0 Then
        MsgBox "No assumptions to clear.", vbInformation, "Clear Assumptions"
        Exit Sub
    End If

    Dim resp As VbMsgBoxResult
    resp = MsgBox("This will remove all " & rowCount & " assumptions from the register." & vbCrLf & vbCrLf & _
                  "This cannot be undone. Continue?", _
                  vbYesNo + vbExclamation, "Clear All Assumptions")
    If resp <> vbYes Then Exit Sub

    ' Clear the data rows (keep marker + header)
    wsConfig.Range(wsConfig.Cells(dr, 1), wsConfig.Cells(dr + rowCount - 1, 13)).ClearContents

    ' Regenerate (will show empty register)
    GenerateAssumptionsRegister

    MsgBox rowCount & " assumptions cleared.", vbInformation, "Clear Assumptions"
    Exit Sub

ErrHandler:
    MsgBox "Failed to clear assumptions: " & Err.Description, vbExclamation, "Clear Assumptions"
End Sub


' =============================================================================
' MarkAllReviewed
' Mark a comma-separated list of assumption IDs as reviewed today.
' =============================================================================
Private Sub MarkAllReviewed(idList As String)
    On Error Resume Next

    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)

    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_ASSUMPTIONS_CONFIG)
    If sr = 0 Then Exit Sub

    Dim todayStr As String
    todayStr = Format(Date, "yyyy-mm-dd")

    Dim ids() As String
    ids = Split(idList, ", ")

    Dim idx As Long
    For idx = 0 To UBound(ids)
        Dim foundRow As Long
        foundRow = FindAssumptionRow(wsConfig, sr, Trim$(ids(idx)))
        If foundRow > 0 Then
            wsConfig.Cells(foundRow, ACFG_COL_REVIEWED).NumberFormat = "@"
            wsConfig.Cells(foundRow, ACFG_COL_REVIEWED).Value = todayStr

            Dim curHist As String
            curHist = CStr(wsConfig.Cells(foundRow, ACFG_COL_HISTORY).Value)
            wsConfig.Cells(foundRow, ACFG_COL_HISTORY).NumberFormat = "@"
            wsConfig.Cells(foundRow, ACFG_COL_HISTORY).Value = curHist & "; " & todayStr & ": Reviewed"
        End If
    Next idx

    On Error GoTo 0
End Sub


' =============================================================================
' Helper: FindAssumptionRow
' Find the Config sheet row for a given AssumptionID.
' =============================================================================
Private Function FindAssumptionRow(wsConfig As Worksheet, sectionStart As Long, _
                                   targetId As String) As Long
    Dim dr As Long
    dr = sectionStart + 2
    Do While wsConfig.Cells(dr, 1).Value <> "" And _
             Left$(CStr(wsConfig.Cells(dr, 1).Value), 3) <> "==="
        If StrComp(CStr(wsConfig.Cells(dr, ACFG_COL_ID).Value), targetId, vbTextCompare) = 0 Then
            FindAssumptionRow = dr
            Exit Function
        End If
        dr = dr + 1
    Loop
    FindAssumptionRow = 0
End Function


' =============================================================================
' Helper: MapFieldToColumn
' Map user-friendly field name to config column index.
' =============================================================================
Private Function MapFieldToColumn(fieldName As String) As Long
    Select Case LCase$(fieldName)
        Case "description":       MapFieldToColumn = ACFG_COL_DESC
        Case "rationale":         MapFieldToColumn = ACFG_COL_RATIONALE
        Case "source":            MapFieldToColumn = ACFG_COL_SOURCE
        Case "confidence":        MapFieldToColumn = ACFG_COL_CONFIDENCE
        Case "sensitivity":       MapFieldToColumn = ACFG_COL_SENSITIVITY
        Case "sensitivitydetail": MapFieldToColumn = ACFG_COL_IMPACT
        Case "owner":             MapFieldToColumn = ACFG_COL_OWNER
        Case Else:                MapFieldToColumn = 0
    End Select
End Function


' =============================================================================
' Helper: FindRowIDAddress
' Find the cell address for a RowID on a given tab by scanning column B.
' Returns cell address string (e.g., "C5") or empty string.
' =============================================================================
Private Function FindRowIDAddress(tabName As String, rowId As String) As String
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(tabName)
    If ws Is Nothing Then
        FindRowIDAddress = ""
        Exit Function
    End If
    On Error GoTo 0

    ' Scan column B for the RowID (formula_tab_config convention: RowID in col B)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow > 500 Then lastRow = 500  ' safety cap

    Dim r As Long
    For r = 1 To lastRow
        If StrComp(CStr(ws.Cells(r, 2).Value), rowId, vbTextCompare) = 0 Then
            ' Return column C of the matching row (first data column)
            FindRowIDAddress = "C" & r
            Exit Function
        End If
    Next r

    ' Also check column A (some tabs use col A for RowID labels)
    For r = 1 To lastRow
        If StrComp(CStr(ws.Cells(r, 1).Value), rowId, vbTextCompare) = 0 Then
            FindRowIDAddress = "B" & r
            Exit Function
        End If
    Next r

    FindRowIDAddress = ""
End Function


' =============================================================================
' Helper: EnsureTab
' Get or create a worksheet by name.
' =============================================================================
Private Function EnsureTab(tabName As String) As Worksheet
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(tabName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = tabName
    End If

    Set EnsureTab = ws
End Function


' =============================================================================
' SyncAssumptionsToCSV
' Exports the assumptions_config section from the Config sheet back to
' config/assumptions_config.csv. Called before workspace save to ensure
' user edits (Add/Edit/Archive) are persisted to disk.
' =============================================================================
Public Sub SyncAssumptionsToCSV()
    On Error Resume Next

    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    If wsConfig Is Nothing Then Exit Sub

    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_ASSUMPTIONS_CONFIG)
    If sr = 0 Then Exit Sub

    ' Header row is sr + 1
    Dim headerRow As Long
    headerRow = sr + 1

    ' Count data rows
    Dim rowCount As Long
    rowCount = 0
    Dim dr As Long
    dr = headerRow + 1
    Do While wsConfig.Cells(dr + rowCount, 1).Value <> "" And _
             Left$(CStr(wsConfig.Cells(dr + rowCount, 1).Value), 3) <> "==="
        rowCount = rowCount + 1
    Loop

    ' Build CSV output
    Dim root As String
    root = ThisWorkbook.Path & "\.."
    Dim csvPath As String
    csvPath = root & "\config\assumptions_config.csv"

    Dim fileNum As Integer
    fileNum = FreeFile
    Open csvPath For Output As #fileNum

    ' Write header
    Print #fileNum, """AssumptionID"",""Category"",""TabName"",""RowID"",""Description"",""Rationale"",""Source"",""Confidence"",""Sensitivity"",""SensitivityDetail"",""Owner"",""LastReviewed"",""History"""

    ' Write data rows
    Dim i As Long
    For i = 1 To rowCount
        Dim r As Long
        r = dr + i - 1
        Dim lineStr As String
        lineStr = ""
        Dim j As Long
        For j = 1 To 13
            Dim cellVal As String
            cellVal = CStr(wsConfig.Cells(r, j).Value)
            ' Escape quotes and wrap in quotes
            cellVal = Replace(cellVal, """", """""")
            If j > 1 Then lineStr = lineStr & ","
            lineStr = lineStr & """" & cellVal & """"
        Next j
        Print #fileNum, lineStr
    Next i

    Close #fileNum
    On Error GoTo 0
End Sub
