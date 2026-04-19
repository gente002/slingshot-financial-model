Attribute VB_Name = "KernelBootstrapUI"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelBootstrapUI.bas
' Purpose: Tab creation, inputs generation, dashboard setup, cover page,
'          fingerprinting, and workspace seeding. Split from KernelBootstrap.bas
'          to stay under the 64KB VBA module limit (AD-09).
' =============================================================================

' Module-level diagnostic variable for sub-step tracking (used by GenerateInputsTab)
Public m_subStep As String


' =============================================================================
' CreateTabsFromRegistry
' Creates all tabs from tab_registry, orders by SortOrder, applies tab colors.
' =============================================================================
Public Sub CreateTabsFromRegistry()
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)

    Dim startRow As Long
    startRow = FindTabRegistryStart(wsConfig)
    If startRow = 0 Then Exit Sub

    ' Skip header row
    Dim dataRow As Long
    dataRow = startRow + 2

    ' Read tab entries
    Do While Trim(CStr(wsConfig.Cells(dataRow, TREG_COL_TABNAME).Value)) <> ""
        Dim tabName As String
        tabName = Trim(CStr(wsConfig.Cells(dataRow, TREG_COL_TABNAME).Value))

        Dim isProtected As String
        isProtected = Trim(CStr(wsConfig.Cells(dataRow, TREG_COL_PROTECTED).Value))

        Dim visibility As String
        visibility = Trim(CStr(wsConfig.Cells(dataRow, TREG_COL_VISIBLE).Value))

        ' Create sheet if not exists (skip Config - already exists)
        If tabName <> TAB_CONFIG Then
            Dim ws As Worksheet
            Set ws = EnsureSheet(tabName)

            ' Set visibility
            If StrComp(visibility, "Hidden", vbTextCompare) = 0 Then
                ws.Visible = xlSheetHidden
            ElseIf StrComp(visibility, "VeryHidden", vbTextCompare) = 0 Then
                ws.Visible = xlSheetVeryHidden
            Else
                ws.Visible = xlSheetVisible
            End If
        End If

        dataRow = dataRow + 1
    Loop

    ' Minimum kernel infrastructure tabs (always required regardless of tab_registry)
    EnsureSheet TAB_CONFIG
    EnsureSheet TAB_DASHBOARD
    EnsureSheet TAB_DETAIL
    EnsureSheet TAB_ERROR_LOG
    ' All other tabs come solely from tab_registry (config is the authority)

    ' --- Order tabs by SortOrder and apply tab colors ---
    Dim tabCount As Long
    tabCount = 0
    Dim countRow As Long
    countRow = startRow + 2
    Do While Trim(CStr(wsConfig.Cells(countRow, TREG_COL_TABNAME).Value)) <> ""
        tabCount = tabCount + 1
        countRow = countRow + 1
    Loop

    If tabCount > 0 Then
        Dim sortNames() As String
        Dim sortOrders() As Long
        Dim sortColors() As String
        ReDim sortNames(1 To tabCount)
        ReDim sortOrders(1 To tabCount)
        ReDim sortColors(1 To tabCount)

        Dim idx As Long
        Dim sortVal As String
        For idx = 1 To tabCount
            sortNames(idx) = Trim(CStr(wsConfig.Cells(startRow + 1 + idx, TREG_COL_TABNAME).Value))
            sortVal = Trim(CStr(wsConfig.Cells(startRow + 1 + idx, TREG_COL_SORTORDER).Value))
            If IsNumeric(sortVal) And Len(sortVal) > 0 Then
                sortOrders(idx) = CLng(sortVal)
            Else
                sortOrders(idx) = 999
            End If
            sortColors(idx) = Trim(CStr(wsConfig.Cells(startRow + 1 + idx, TREG_COL_TABCOLOR).Value))
        Next idx

        ' Bubble sort by SortOrder
        Dim swapped As Boolean
        Dim tmpName As String
        Dim tmpOrder As Long
        Dim tmpColor As String
        Do
            swapped = False
            For idx = 1 To tabCount - 1
                If sortOrders(idx) > sortOrders(idx + 1) Then
                    tmpName = sortNames(idx)
                    tmpOrder = sortOrders(idx)
                    tmpColor = sortColors(idx)
                    sortNames(idx) = sortNames(idx + 1)
                    sortOrders(idx) = sortOrders(idx + 1)
                    sortColors(idx) = sortColors(idx + 1)
                    sortNames(idx + 1) = tmpName
                    sortOrders(idx + 1) = tmpOrder
                    sortColors(idx + 1) = tmpColor
                    swapped = True
                End If
            Next idx
        Loop While swapped

        ' Move tabs in sorted order
        Dim lastPlaced As String
        Dim wsTab As Worksheet
        lastPlaced = ""
        On Error Resume Next
        For idx = 1 To tabCount
            Set wsTab = Nothing
            Set wsTab = ThisWorkbook.Sheets(sortNames(idx))
            If Not wsTab Is Nothing Then
                If Len(lastPlaced) = 0 Then
                    wsTab.Move Before:=ThisWorkbook.Sheets(1)
                Else
                    wsTab.Move After:=ThisWorkbook.Sheets(lastPlaced)
                End If
                lastPlaced = sortNames(idx)
            End If
        Next idx
        On Error GoTo 0

        ' Apply tab colors from hex values (6-char hex without #)
        Dim hexClr As String
        Dim rVal As Long
        Dim gVal As Long
        Dim bVal As Long
        On Error Resume Next
        For idx = 1 To tabCount
            If Len(sortColors(idx)) = 6 Then
                Set wsTab = Nothing
                Set wsTab = ThisWorkbook.Sheets(sortNames(idx))
                If Not wsTab Is Nothing Then
                    hexClr = sortColors(idx)
                    rVal = CLng("&H" & Mid(hexClr, 1, 2))
                    gVal = CLng("&H" & Mid(hexClr, 3, 2))
                    bVal = CLng("&H" & Mid(hexClr, 5, 2))
                    wsTab.Tab.Color = RGB(rVal, gVal, bVal)
                End If
            End If
        Next idx
        On Error GoTo 0
    End If

    ' Delete the default "Sheet1" if it exists and we have other sheets
    On Error Resume Next
    Dim defaultSheet As Worksheet
    Set defaultSheet = ThisWorkbook.Sheets("Sheet1")
    If Not defaultSheet Is Nothing Then
        If ThisWorkbook.Sheets.Count > 1 Then
            defaultSheet.Delete
        End If
    End If
    On Error GoTo 0
End Sub


' =============================================================================
' FindTabRegistryStart
' Finds the TAB_REGISTRY section on the Config sheet.
' =============================================================================
Public Function FindTabRegistryStart(ws As Worksheet) As Long
    Dim scanRow As Long
    For scanRow = 1 To 500
        If Trim(CStr(ws.Cells(scanRow, 1).Value)) = CFG_MARKER_TAB_REGISTRY Then
            FindTabRegistryStart = scanRow
            Exit Function
        End If
    Next scanRow
    FindTabRegistryStart = 0
End Function


' =============================================================================
' EnsureSheet
' Creates a sheet if it does not exist. Returns the sheet reference.
' =============================================================================
Public Function EnsureSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    End If

    Set EnsureSheet = ws
End Function


' =============================================================================
' GenerateInputsTab
' Generates the Inputs tab layout from InputSchema.
' =============================================================================
Public Sub GenerateInputsTab()
    m_subStep = "GetSheet"

    ' Phase 11A: Support configurable tab name (Assumptions for insurance)
    Dim inputsTabName As String
    inputsTabName = CStr(KernelConfig.GetSetting("InputsTabName"))
    If Len(inputsTabName) = 0 Then inputsTabName = TAB_INPUTS

    Dim wsInputs As Worksheet
    On Error Resume Next
    Set wsInputs = ThisWorkbook.Sheets(inputsTabName)
    On Error GoTo 0
    If wsInputs Is Nothing Then
        Set wsInputs = ThisWorkbook.Sheets(TAB_INPUTS)
    End If

    m_subStep = "ClearContents"
    wsInputs.Cells.ClearContents

    m_subStep = "UnMerge"
    ' Unmerge all cells first in case sheet was reused
    wsInputs.Cells.UnMerge

    ' Row 1: Title
    m_subStep = "WriteTitle"
    wsInputs.Cells(1, 1).Value = "RDK Model -- " & wsInputs.Name

    m_subStep = "MergeTitle"
    wsInputs.Range(wsInputs.Cells(1, 1), wsInputs.Cells(1, 5)).Merge

    m_subStep = "FormatTitle"
    With wsInputs.Cells(1, 1)
        .Font.Size = 14
        .Font.Bold = True
    End With

    ' Row 3: Entity header row (only if EntityName is in input_schema)
    m_subStep = "EntityHeader"
    Dim hasEntityParam As Boolean
    hasEntityParam = False
    Dim epIdx As Long
    For epIdx = 1 To KernelConfig.GetInputCount()
        If StrComp(KernelConfig.GetInputParam(epIdx), "EntityName", vbTextCompare) = 0 Then
            hasEntityParam = True
            Exit For
        End If
    Next epIdx
    If hasEntityParam Then
        wsInputs.Cells(3, 1).Value = "Entity ->"
        wsInputs.Cells(3, 1).Font.Bold = True
    End If

    ' Read input parameters and organize by section
    m_subStep = "GetInputCount"
    Dim paramCount As Long
    paramCount = KernelConfig.GetInputCount()

    Dim lastSection As String
    lastSection = ""

    Dim pidx As Long
    For pidx = 1 To paramCount
        m_subStep = "ReadParam_" & pidx

        Dim section As String
        section = KernelConfig.GetInputSection(pidx)

        Dim paramName As String
        paramName = KernelConfig.GetInputParam(pidx)

        Dim paramRow As Long
        paramRow = KernelConfig.GetInputRow(pidx)

        m_subStep = "Param_" & pidx & "_" & paramName & "_row" & paramRow

        ' Check if we need a section header
        If StrComp(section, lastSection, vbTextCompare) <> 0 Then
            ' Find the section header row (1 row before first param in section)
            Dim sectionHeaderRow As Long
            sectionHeaderRow = paramRow - 1

            m_subStep = "SectionHdr_" & section & "_row" & sectionHeaderRow

            ' Write section header (force text format to prevent "===" being parsed as formula)
            wsInputs.Cells(sectionHeaderRow, 1).NumberFormat = "@"
            wsInputs.Cells(sectionHeaderRow, 1).Value = "=== " & UCase(section) & " ==="
            With wsInputs.Cells(sectionHeaderRow, 1)
                .Font.Bold = True
                .Interior.Color = RGB(217, 225, 242)
            End With

            lastSection = section
        End If

        ' Write parameter name in column A
        m_subStep = "WriteParam_" & paramName
        wsInputs.Cells(paramRow, 1).Value = paramName

        ' Write tooltip/type info in column B
        m_subStep = "WriteType_" & paramName
        Dim paramType As String
        paramType = KernelConfig.GetInputType(pidx)
        wsInputs.Cells(paramRow, 2).Value = paramType
        wsInputs.Cells(paramRow, 2).Font.Color = RGB(128, 128, 128)
        wsInputs.Cells(paramRow, 2).Font.Italic = True

        ' Apply number format based on DataType
        Dim fmtEndCol As Long
        If hasEntityParam Then
            fmtEndCol = INPUT_ENTITY_START_COL + INPUT_MAX_ENTITIES - 1
        Else
            fmtEndCol = INPUT_ENTITY_START_COL
        End If
        If StrComp(paramType, "Pct", vbTextCompare) = 0 Then
            wsInputs.Range(wsInputs.Cells(paramRow, INPUT_ENTITY_START_COL), _
                wsInputs.Cells(paramRow, fmtEndCol)).NumberFormat = "0.0%"
        ElseIf StrComp(paramType, "Currency", vbTextCompare) = 0 Then
            wsInputs.Range(wsInputs.Cells(paramRow, INPUT_ENTITY_START_COL), _
                wsInputs.Cells(paramRow, fmtEndCol)).NumberFormat = "#,##0.00"
        End If
    Next pidx

    ' Populate sample fixture data
    m_subStep = "PopulateSampleData"
    PopulateSampleData wsInputs, hasEntityParam

    ' Auto-fit columns
    m_subStep = "AutoFit"
    wsInputs.Columns.AutoFit
End Sub


' =============================================================================
' PopulateSampleData
' Populates 3 entities with the deterministic fixture values.
' =============================================================================
Public Sub PopulateSampleData(ws As Worksheet, Optional hasEntities As Boolean = True)
    ' Entity names in row 3 (only for entity-based tabs)
    If hasEntities Then
        ws.Cells(3, INPUT_ENTITY_START_COL).Value = "Product A"
        ws.Cells(3, INPUT_ENTITY_START_COL + 1).Value = "Product B"
        ws.Cells(3, INPUT_ENTITY_START_COL + 2).Value = "Product C"
    End If

    ' Populate from InputSchema defaults and fixture overrides
    Dim paramCount As Long
    paramCount = KernelConfig.GetInputCount()

    Dim pidx As Long
    For pidx = 1 To paramCount
        Dim paramName As String
        paramName = KernelConfig.GetInputParam(pidx)

        Dim paramRow As Long
        paramRow = KernelConfig.GetInputRow(pidx)

        Select Case paramName
            Case "EntityName"
                ws.Cells(paramRow, INPUT_ENTITY_START_COL).Value = "Product A"
                ws.Cells(paramRow, INPUT_ENTITY_START_COL + 1).Value = "Product B"
                ws.Cells(paramRow, INPUT_ENTITY_START_COL + 2).Value = "Product C"

            Case "StartDate"
                ws.Cells(paramRow, INPUT_ENTITY_START_COL).Value = "2026-01-01"
                ws.Cells(paramRow, INPUT_ENTITY_START_COL + 1).Value = "2026-01-01"
                ws.Cells(paramRow, INPUT_ENTITY_START_COL + 2).Value = "2026-01-01"

            Case "TermMonths"
                ws.Cells(paramRow, INPUT_ENTITY_START_COL).Value = 12
                ws.Cells(paramRow, INPUT_ENTITY_START_COL + 1).Value = 12
                ws.Cells(paramRow, INPUT_ENTITY_START_COL + 2).Value = 12

            Case "Units"
                ws.Cells(paramRow, INPUT_ENTITY_START_COL).Value = 100
                ws.Cells(paramRow, INPUT_ENTITY_START_COL + 1).Value = 200
                ws.Cells(paramRow, INPUT_ENTITY_START_COL + 2).Value = 50

            Case "UnitPrice"
                ws.Cells(paramRow, INPUT_ENTITY_START_COL).Value = 250#
                ws.Cells(paramRow, INPUT_ENTITY_START_COL + 1).Value = 125#
                ws.Cells(paramRow, INPUT_ENTITY_START_COL + 2).Value = 500#

            Case "MonthlyGrowth"
                ws.Cells(paramRow, INPUT_ENTITY_START_COL).Value = 0.009
                ws.Cells(paramRow, INPUT_ENTITY_START_COL + 1).Value = 0.005
                ws.Cells(paramRow, INPUT_ENTITY_START_COL + 2).Value = 0.012

            Case "COGSPct"
                ws.Cells(paramRow, INPUT_ENTITY_START_COL).Value = 0.6
                ws.Cells(paramRow, INPUT_ENTITY_START_COL + 1).Value = 0.6
                ws.Cells(paramRow, INPUT_ENTITY_START_COL + 2).Value = 0.6

            Case Else
                ' Write defaults: column C only for global params, C-E for entity params
                Dim defVal As String
                defVal = KernelConfig.GetInputDefault(pidx)
                If Len(defVal) > 0 Then
                    ws.Cells(paramRow, INPUT_ENTITY_START_COL).Value = defVal
                    ' Only spread to entity columns if this is an entity-based tab
                    If hasEntities Then
                        ws.Cells(paramRow, INPUT_ENTITY_START_COL + 1).Value = defVal
                        ws.Cells(paramRow, INPUT_ENTITY_START_COL + 2).Value = defVal
                    End If
                End If
        End Select
    Next pidx
End Sub


' =============================================================================
' SetupDetailHeaders
' =============================================================================
Public Sub SetupDetailHeaders()
    Dim wsDetail As Worksheet
    Set wsDetail = ThisWorkbook.Sheets(TAB_DETAIL)
    wsDetail.Cells.ClearContents

    Dim totalCols As Long
    totalCols = KernelConfig.GetColumnCount()
    If totalCols = 0 Then Exit Sub

    ' Write headers using array batch write (PT-001)
    Dim headerArr() As Variant
    ReDim headerArr(1 To 1, 1 To totalCols)

    Dim colIdx As Long
    For colIdx = 1 To totalCols
        headerArr(1, KernelConfig.GetDetailCol(colIdx)) = KernelConfig.GetColName(colIdx)
    Next colIdx

    wsDetail.Range(wsDetail.Cells(DETAIL_HEADER_ROW, 1), _
                   wsDetail.Cells(DETAIL_HEADER_ROW, totalCols)).Value = headerArr

    With wsDetail.Range(wsDetail.Cells(DETAIL_HEADER_ROW, 1), _
                        wsDetail.Cells(DETAIL_HEADER_ROW, totalCols))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
End Sub


' =============================================================================
' SetupSummaryStructure
' =============================================================================
Public Sub SetupSummaryStructure()
    Dim wsSummary As Worksheet
    Set wsSummary = Nothing
    On Error Resume Next
    Set wsSummary = ThisWorkbook.Sheets(TAB_SUMMARY)
    On Error GoTo 0
    If wsSummary Is Nothing Then Exit Sub

    wsSummary.Cells.ClearContents
    wsSummary.Cells(1, 1).Value = "Entity"
    wsSummary.Cells(1, 2).Value = "Metric"
    wsSummary.Cells(1, 1).Font.Bold = True
    wsSummary.Cells(1, 2).Font.Bold = True
End Sub


' =============================================================================
' SetupErrorLog
' =============================================================================
Public Sub SetupErrorLog()
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets(TAB_ERROR_LOG)
    wsLog.Cells.ClearContents

    wsLog.Cells(1, 1).Value = "Timestamp"
    wsLog.Cells(1, 2).Value = "Severity"
    wsLog.Cells(1, 3).Value = "Source"
    wsLog.Cells(1, 4).Value = "Code"
    wsLog.Cells(1, 5).Value = "Message"
    wsLog.Cells(1, 6).Value = "Detail"

    With wsLog.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    wsLog.Columns(1).ColumnWidth = 20
    wsLog.Columns(2).ColumnWidth = 10
    wsLog.Columns(3).ColumnWidth = 20
    wsLog.Columns(4).ColumnWidth = 10
    wsLog.Columns(5).ColumnWidth = 50
    wsLog.Columns(6).ColumnWidth = 50
End Sub


' =============================================================================
' SetupDashboardTab
' =============================================================================
Public Sub SetupDashboardTab()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_DASHBOARD)
    ws.Cells.ClearContents

    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp

    ws.Cells.Interior.Color = RGB(255, 255, 255)

    ws.Columns(1).ColumnWidth = 2
    ws.Columns(2).ColumnWidth = 28
    ws.Columns(3).ColumnWidth = 6
    ws.Columns(4).ColumnWidth = 14
    ws.Columns(5).ColumnWidth = 10
    ws.Columns(6).ColumnWidth = 8
    ws.Columns(7).ColumnWidth = 2

    Dim rh As Long
    For rh = 1 To 30
        ws.Rows(rh).RowHeight = 20
    Next rh
    ws.Rows(1).RowHeight = 6
    ws.Rows(2).RowHeight = 30
    ws.Rows(3).RowHeight = 16
    ws.Rows(4).RowHeight = 8
    ws.Rows(5).RowHeight = 6

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0

    Dim modelTitle As String
    modelTitle = KernelConfig.GetBrandingSetting("CompanyName")
    Dim modelSubtitle As String
    modelSubtitle = KernelConfig.GetBrandingSetting("ModelTitle")
    If Len(modelTitle) > 0 And Len(modelSubtitle) > 0 Then
        modelTitle = modelTitle & " " & modelSubtitle
    ElseIf Len(modelTitle) = 0 Then
        modelTitle = "RDK Dashboard"
    End If
    ws.Cells(2, 2).Value = modelTitle
    With ws.Cells(2, 2)
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(31, 56, 100)
    End With

    Dim cfgVer As String
    cfgVer = KernelConfig.GetConfigVersion("ConfigVersion")
    Dim verText As String
    verText = "Kernel v" & KERNEL_VERSION
    If Len(cfgVer) > 0 Then verText = verText & "  |  Config v" & cfgVer
    ws.Cells(3, 2).Value = verText
    ws.Cells(3, 2).Font.Size = 10
    ws.Cells(3, 2).Font.Italic = True
    ws.Cells(3, 2).Font.Color = RGB(128, 128, 128)

    ws.Range(ws.Cells(4, 2), ws.Cells(4, 7)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range(ws.Cells(4, 2), ws.Cells(4, 7)).Borders(xlEdgeBottom).Color = RGB(31, 56, 100)
    ws.Range(ws.Cells(4, 2), ws.Cells(4, 7)).Borders(xlEdgeBottom).Weight = xlMedium

    ws.Range(ws.Cells(6, 4), ws.Cells(6, 6)).Merge
    ws.Cells(6, 4).Value = "MODEL OUTPUT"
    ws.Cells(6, 4).Font.Color = RGB(255, 255, 255)
    ws.Cells(6, 4).Font.Bold = True
    ws.Cells(6, 4).Font.Size = 10
    ws.Cells(6, 4).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(6, 4), ws.Cells(6, 6)).Interior.Color = RGB(31, 56, 100)

    Dim moLabels As Variant
    moLabels = Array("Last Run", "Scenario", "Programs", "Total GWP")
    Dim moRow As Long
    For moRow = 0 To 3
        ws.Cells(7 + moRow, 4).Value = moLabels(moRow)
        ws.Cells(7 + moRow, 4).Font.Bold = True
        ws.Cells(7 + moRow, 4).Font.Size = 10
        ws.Range(ws.Cells(7 + moRow, 5), ws.Cells(7 + moRow, 6)).Merge
        ws.Cells(7 + moRow, 5).HorizontalAlignment = xlRight
        ws.Cells(7 + moRow, 5).Font.Size = 10
        If moRow Mod 2 = 0 Then
            ws.Range(ws.Cells(7 + moRow, 4), ws.Cells(7 + moRow, 6)).Interior.Color = RGB(214, 228, 240)
        End If
        ws.Range(ws.Cells(7 + moRow, 4), ws.Cells(7 + moRow, 6)).Borders.LineStyle = xlContinuous
        ws.Range(ws.Cells(7 + moRow, 4), ws.Cells(7 + moRow, 6)).Borders.Color = RGB(191, 191, 191)
        ws.Range(ws.Cells(7 + moRow, 4), ws.Cells(7 + moRow, 6)).Borders.Weight = xlThin
    Next moRow

    ws.Cells(14, 2).Value = "DEVELOPER TOOLS"
    ws.Cells(14, 2).Font.Bold = True
    ws.Cells(14, 2).Font.Italic = True
    ws.Cells(14, 2).Font.Size = 9
    ws.Cells(14, 2).Font.Color = RGB(100, 100, 100)
    ws.Range(ws.Cells(14, 2), ws.Cells(14, 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range(ws.Cells(14, 2), ws.Cells(14, 2)).Borders(xlEdgeBottom).Color = RGB(191, 191, 191)
    ws.Range(ws.Cells(14, 2), ws.Cells(14, 2)).Borders(xlEdgeBottom).Weight = xlThin

    Dim initDevMode As String
    initDevMode = KernelConfig.GetDevMode()
    If Len(initDevMode) = 0 Then initDevMode = DEV_MODE_OFF
    KernelButtons.CreateButtonsFromConfig ws, TAB_DASHBOARD, initDevMode
End Sub


' =============================================================================
' ProtectOutputTabs
' =============================================================================
Public Sub ProtectOutputTabs()
    On Error Resume Next
    ThisWorkbook.Sheets(TAB_DETAIL).Protect UserInterfaceOnly:=True
    Dim wsSumProt As Worksheet
    Set wsSumProt = ThisWorkbook.Sheets(TAB_SUMMARY)
    If Not wsSumProt Is Nothing Then wsSumProt.Protect UserInterfaceOnly:=True
    On Error GoTo 0
End Sub


' =============================================================================
' ApplyDefaultDevMode
' =============================================================================
Public Sub ApplyDefaultDevMode()
    On Error Resume Next
    Dim devMode As String
    devMode = KernelConfig.GetDevMode()
    If Len(devMode) = 0 Then devMode = DEV_MODE_OFF
    KernelConfig.SetDevMode devMode

    If devMode = DEV_MODE_OFF Then
        Dim devTabs As Variant
        devTabs = GetDevModeTabs()
        Dim idx As Long
        For idx = LBound(devTabs) To UBound(devTabs)
            Dim ws As Worksheet
            Set ws = Nothing
            Set ws = ThisWorkbook.Sheets(devTabs(idx))
            If Not ws Is Nothing Then
                ws.Visible = xlSheetHidden
            End If
        Next idx
    End If
    On Error GoTo 0
End Sub


' =============================================================================
' PopulateCoverPage (12c)
' =============================================================================
Public Sub PopulateCoverPage()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_COVER_PAGE)
    If ws Is Nothing Then Exit Sub
    On Error GoTo 0

    ws.Cells.ClearContents

    Dim cpName As String
    cpName = KernelConfig.GetBrandingSetting("CompanyName")
    If Len(cpName) = 0 Then cpName = "RDK Model"
    ws.Cells(2, 2).Value = cpName
    ws.Cells(2, 2).Font.Bold = True
    ws.Cells(2, 2).Font.Size = 20
    ws.Cells(2, 2).Font.Color = RGB(31, 56, 100)

    Dim cpTitle As String
    cpTitle = KernelConfig.GetBrandingSetting("ModelTitle")
    ws.Cells(3, 2).Value = cpTitle
    ws.Cells(3, 2).Font.Bold = True
    ws.Cells(3, 2).Font.Size = 14

    Dim cpTag As String
    cpTag = KernelConfig.GetBrandingSetting("Tagline")
    If Len(cpTag) > 0 Then
        ws.Cells(4, 2).Value = cpTag
        ws.Cells(4, 2).Font.Italic = True
        ws.Cells(4, 2).Font.Size = 11
        ws.Cells(4, 2).Font.Color = RGB(128, 128, 128)
    End If

    ws.Cells(6, 2).Value = "Kernel Version"
    ws.Cells(6, 2).Font.Bold = True
    ws.Cells(6, 3).Value = KERNEL_VERSION

    Dim cfgVer As String
    cfgVer = KernelConfig.GetConfigVersion("ConfigVersion")
    If Len(cfgVer) > 0 Then
        ws.Cells(7, 2).Value = "Config Version"
        ws.Cells(7, 2).Font.Bold = True
        ws.Cells(7, 3).Value = cfgVer
    End If

    ws.Cells(8, 2).Value = "Date"
    ws.Cells(8, 2).Font.Bold = True
    ws.Cells(8, 3).formula = "=TODAY()"
    ws.Cells(8, 3).NumberFormat = "mm/dd/yyyy"
    ws.Cells(8, 3).HorizontalAlignment = xlLeft

    ws.Cells(9, 2).Value = "Scenario"
    ws.Cells(9, 2).Font.Bold = True
    On Error Resume Next
    ws.Cells(9, 3).Value = CStr(KernelConfig.InputValue("Global Assumptions", "ScenarioName", 1))
    On Error GoTo 0

    ws.Cells(11, 2).Value = "Table of Contents"
    ws.Cells(10, 2).Font.Bold = True
    ws.Cells(10, 2).Font.Size = 12

    Dim wsConfig As Worksheet
    Set wsConfig = Nothing
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    On Error GoTo 0
    If wsConfig Is Nothing Then Exit Sub

    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_TAB_REGISTRY)
    If sr = 0 Then Exit Sub

    Dim dr As Long
    dr = sr + 2

    Dim tocCount As Long
    tocCount = 0
    Dim scanRow As Long
    scanRow = dr
    Do While Len(Trim(CStr(wsConfig.Cells(scanRow, TREG_COL_TABNAME).Value))) > 0
        If StrComp(Trim(CStr(wsConfig.Cells(scanRow, TREG_COL_VISIBLE).Value)), _
                   "Visible", vbTextCompare) = 0 Then
            tocCount = tocCount + 1
        End If
        scanRow = scanRow + 1
    Loop
    If tocCount = 0 Then Exit Sub

    Dim tocNames() As String
    Dim tocDescs() As String
    Dim tocSorts() As Long
    ReDim tocNames(1 To tocCount)
    ReDim tocDescs(1 To tocCount)
    ReDim tocSorts(1 To tocCount)

    Dim ti As Long
    ti = 0
    scanRow = dr
    Do While Len(Trim(CStr(wsConfig.Cells(scanRow, TREG_COL_TABNAME).Value))) > 0
        If StrComp(Trim(CStr(wsConfig.Cells(scanRow, TREG_COL_VISIBLE).Value)), _
                   "Visible", vbTextCompare) = 0 Then
            ti = ti + 1
            tocNames(ti) = Trim(CStr(wsConfig.Cells(scanRow, TREG_COL_TABNAME).Value))
            tocDescs(ti) = Trim(CStr(wsConfig.Cells(scanRow, 7).Value))
            Dim soVal As String
            soVal = Trim(CStr(wsConfig.Cells(scanRow, TREG_COL_SORTORDER).Value))
            If IsNumeric(soVal) Then tocSorts(ti) = CLng(soVal) Else tocSorts(ti) = 999
        End If
        scanRow = scanRow + 1
    Loop

    Dim i As Long
    Dim j As Long
    For i = 1 To tocCount - 1
        For j = i + 1 To tocCount
            If tocSorts(j) < tocSorts(i) Then
                Dim tmpN As String
                tmpN = tocNames(i): tocNames(i) = tocNames(j): tocNames(j) = tmpN
                Dim tmpD As String
                tmpD = tocDescs(i): tocDescs(i) = tocDescs(j): tocDescs(j) = tmpD
                Dim tmpS As Long
                tmpS = tocSorts(i): tocSorts(i) = tocSorts(j): tocSorts(j) = tmpS
            End If
        Next j
    Next i

    Dim tocRow As Long
    tocRow = 12
    For i = 1 To tocCount
        If StrComp(tocNames(i), TAB_COVER_PAGE, vbTextCompare) <> 0 Then
            ws.Cells(tocRow, 2).formula = "=HYPERLINK(""#'" & _
                tocNames(i) & "'!A1"",""" & tocNames(i) & """)"
            ws.Cells(tocRow, 2).Font.Color = RGB(5, 99, 193)
            ws.Cells(tocRow, 2).Font.Underline = xlUnderlineStyleSingle
            ws.Cells(tocRow, 3).Value = tocDescs(i)
            ws.Cells(tocRow, 3).Font.Color = RGB(128, 128, 128)
            tocRow = tocRow + 1
        End If
    Next i

    ws.Columns(1).ColumnWidth = 5
    ws.Columns(2).ColumnWidth = 30
    ws.Columns(3).ColumnWidth = 50
    ws.Cells.Interior.Color = RGB(255, 255, 255)

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
End Sub


' =============================================================================
' StampFingerprint
' =============================================================================
Public Sub StampFingerprint()
    On Error Resume Next
    Dim wsFP As Worksheet
    Set wsFP = Nothing
    Set wsFP = ThisWorkbook.Sheets("_fp")
    If wsFP Is Nothing Then
        Set wsFP = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsFP.Name = "_fp"
    End If
    wsFP.Visible = xlSheetVeryHidden
    wsFP.Cells.ClearContents
    wsFP.Cells(1, 1).Value = "Author"
    wsFP.Cells(1, 2).Value = KERNEL_AUTHOR
    wsFP.Cells(2, 1).Value = "BuildID"
    wsFP.Cells(2, 2).Value = KERNEL_BUILD_ID
    wsFP.Cells(3, 1).Value = "BuildDate"
    wsFP.Cells(3, 2).Value = KERNEL_BUILD_DATE
    wsFP.Cells(4, 1).Value = "KernelVersion"
    wsFP.Cells(4, 2).Value = KERNEL_VERSION
    wsFP.Cells(5, 1).Value = "StampedAt"
    wsFP.Cells(5, 2).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    wsFP.Cells(6, 1).Value = "Machine"
    wsFP.Cells(6, 2).Value = Environ("COMPUTERNAME")
    wsFP.Cells(7, 1).Value = "License"
    wsFP.Cells(7, 2).Value = "Proprietary. All rights reserved."
    wsFP.Cells(1, 2).AddComment "39ee874dd4a5df1a6547dbaa06ad94ce"
    wsFP.Cells(1, 2).Comment.Visible = False
    wsFP.Protect Password:="fp" & KERNEL_BUILD_ID
    Dim props As Object
    Set props = ThisWorkbook.CustomDocumentProperties
    Dim pExists As Boolean: pExists = False
    Dim p As Object
    For Each p In props
        If p.Name = "_dp" Then pExists = True: Exit For
    Next p
    If pExists Then
        props("_dp").Value = "39ee874dd4a5df1a6547dbaa06ad94ce"
    Else
        props.Add "_dp", False, 4, "39ee874dd4a5df1a6547dbaa06ad94ce"
    End If
    Dim nm As Name
    For Each nm In ThisWorkbook.Names
        If nm.Name = "_xlfn.RES" Then nm.Delete: Exit For
    Next nm
    ThisWorkbook.Names.Add "_xlfn.RES", "=""39ee874dd4a5df1a6547dbaa06ad94ce""", False
    On Error GoTo 0
End Sub


' =============================================================================
' SeedStarterWorkspaces
' Creates Base_Case and Sensitivity Testing workspaces at bootstrap time.
' =============================================================================
Public Sub SeedStarterWorkspaces()
    On Error Resume Next

    Dim wsEnabled As String
    wsEnabled = KernelConfig.GetWorkspaceSetting("WorkspacesEnabled")
    If StrComp(wsEnabled, "FALSE", vbTextCompare) = 0 Then Exit Sub

    Dim root As String
    root = ThisWorkbook.Path & "\.."
    Dim wsDir As String
    wsDir = root & "\workspaces"

    ' Always import seed data (populates UW Inputs, Capital, etc.)
    SeedBaseModelInputs

    ' Check if workspaces already exist -- skip workspace creation if so
    Dim skipWS As Boolean: skipWS = False
    If Dir(wsDir, vbDirectory) <> "" Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(wsDir) Then
            If fso.GetFolder(wsDir).SubFolders.Count > 0 Then
                skipWS = True
            End If
        End If
        Set fso = Nothing
    End If

    ' Run the model (includes RefreshFormulaTabs internally,
    ' which rebuilds formula cells destroyed by seed import)
    On Error Resume Next
    KernelEngine.SilentMode = True
    KernelEngine.RunProjectionsEx
    KernelEngine.SilentMode = False

    ' AutoFit all tabs after silent run (SilentMode skips AutoFit)
    Dim afWs As Worksheet
    For Each afWs In ThisWorkbook.Sheets
        If afWs.Visible = xlSheetVisible Then afWs.Columns.AutoFit
    Next afWs
    On Error GoTo 0

    If Not skipWS Then
        ' Save as Base Model (with results)
        KernelWorkspace.SaveWorkspace "Base Model", True

        ' Copy Base Model to Sensitivity Testing (avoids second SaveWorkspace)
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim bmDir As String: bmDir = root & "\workspaces\Base Model"
        Dim stDir As String: stDir = root & "\workspaces\Sensitivity Testing"
        If fso.FolderExists(bmDir) And Not fso.FolderExists(stDir) Then
            fso.CopyFolder bmDir, stDir
            Dim stJson As String: stJson = stDir & "\workspace.json"
            If fso.FileExists(stJson) Then
                Dim jContent As String
                jContent = KernelSnapshot.ReadEntireFile(stJson)
                jContent = Replace(jContent, """Base Model""", """Sensitivity Testing""")
                Dim fn As Integer: fn = FreeFile
                Open stJson For Output As #fn
                Print #fn, jContent;
                Close #fn
            End If
        End If
        Set fso = Nothing

        KernelConfig.LogError SEV_INFO, "KernelBootstrap", "I-510", _
            "Seeded starter workspaces: Base Model, Sensitivity Testing", ""
    End If

    On Error GoTo 0
End Sub


' =============================================================================
' SeedBaseModelInputs
' Loads inputs from the Base Model workspace (single source of truth).
' The workspace directory ships with the repo and is preserved across setups.
' To update the seed: save the Base Model workspace from Excel.
' =============================================================================
Private Sub SeedBaseModelInputs()
    On Error Resume Next

    Dim root As String
    root = ThisWorkbook.Path & "\.."

    ' Find latest version in Base Model workspace
    Dim wsDir As String
    wsDir = root & "\workspaces\Base Model"
    If Dir(wsDir, vbDirectory) = "" Then Exit Sub

    Dim fsoSeed As Object
    Set fsoSeed = CreateObject("Scripting.FileSystemObject")
    If Not fsoSeed.FolderExists(wsDir) Then
        Set fsoSeed = Nothing
        Exit Sub
    End If

    Dim latestVer As String: latestVer = ""
    Dim subF As Object
    For Each subF In fsoSeed.GetFolder(wsDir).SubFolders
        If Left(subF.Name, 1) = "v" Then
            If subF.Name > latestVer Then latestVer = subF.Name
        End If
    Next subF
    Set fsoSeed = Nothing

    If Len(latestVer) = 0 Then Exit Sub

    ' Import input tabs from the workspace
    Dim verDir As String: verDir = wsDir & "\" & latestVer
    KernelTabIO.ImportAllInputTabs verDir

    ' Clear assumptions (0 assumptions for Base Model)
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    Dim acSr As Long
    acSr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_ASSUMPTIONS_CONFIG)
    If acSr > 0 Then
        Dim acDr As Long: acDr = acSr + 2
        Dim acCnt As Long: acCnt = 0
        Do While wsConfig.Cells(acDr + acCnt, 1).Value <> "" And _
                 Left$(CStr(wsConfig.Cells(acDr + acCnt, 1).Value), 3) <> "==="
            acCnt = acCnt + 1
        Loop
        If acCnt > 0 Then
            wsConfig.Range(wsConfig.Cells(acDr, 1), wsConfig.Cells(acDr + acCnt - 1, 13)).ClearContents
        End If
    End If

    On Error GoTo 0
End Sub
