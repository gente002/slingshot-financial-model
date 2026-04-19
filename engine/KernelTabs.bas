Attribute VB_Name = "KernelTabs"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelTabs.bas
' Purpose: Generates presentation artifacts from config: Summary tab layout,
'          Dashboard charts, Exhibit tables, display mode toggle, and
'          CumulativeView formula sheet.
' =============================================================================


' =============================================================================
' RefreshAllPresentations
' Convenience: GenerateSummary + GenerateCharts + GenerateExhibits.
' =============================================================================
Public Sub RefreshAllPresentations()
    On Error GoTo ErrHandler
    GenerateSummary
    GenerateCharts
    GenerateExhibits
    UpdateDashboardMode
    Exit Sub
ErrHandler:
    KernelConfig.LogError SEV_WARN, "KernelTabs", "W-500", _
        "Error in RefreshAllPresentations: " & Err.Description, _
        "MANUAL BYPASS: Run GenerateSummary, GenerateCharts, GenerateExhibits individually."
End Sub


' =============================================================================
' UpdateDashboardMode
' Updates the Dashboard tab display mode indicator and toggle button caption.
' =============================================================================
Public Sub UpdateDashboardMode()
    On Error Resume Next
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets(TAB_DASHBOARD)
    If wsDash Is Nothing Then Exit Sub

    Dim curMode As String
    curMode = KernelConfig.GetCurrentDisplayMode()
    If Len(curMode) = 0 Then curMode = DISPLAY_INCREMENTAL

    ' Update status cell (row 2)
    wsDash.Cells(2, 1).Value = "Display Mode: " & curMode
    wsDash.Cells(2, 1).Font.Bold = True
    wsDash.Cells(2, 1).Font.Size = 11
    If curMode = DISPLAY_CUMULATIVE Then
        wsDash.Cells(2, 1).Font.Color = RGB(0, 128, 0)
    Else
        wsDash.Cells(2, 1).Font.Color = RGB(0, 70, 173)
    End If

    ' Update toggle button caption
    Dim shp As Shape
    For Each shp In wsDash.Shapes
        If shp.OnAction = "KernelTabs.ToggleDisplayMode" Then
            If curMode = DISPLAY_CUMULATIVE Then
                shp.TextFrame.Characters.Text = "Switch to Incremental"
            Else
                shp.TextFrame.Characters.Text = "Switch to Cumulative"
            End If
            Exit For
        End If
    Next shp
    On Error GoTo 0
End Sub


' =============================================================================
' UpdateDashboardDevMode
' Updates the Dashboard tab dev mode indicator and toggle button caption.
' =============================================================================
Public Sub UpdateDashboardDevMode()
    On Error Resume Next
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets(TAB_DASHBOARD)
    If wsDash Is Nothing Then Exit Sub

    Dim curMode As String
    curMode = KernelConfig.GetDevMode()
    If Len(curMode) = 0 Then curMode = DEV_MODE_OFF

    ' Update toggle button caption (find by shape name or OnAction fallback)
    Dim shp As Shape
    Set shp = Nothing
    On Error Resume Next
    Set shp = wsDash.Shapes("btn_DEV_MODE")
    On Error GoTo 0
    If shp Is Nothing Then
        ' Fallback: find by OnAction for legacy dashboards
        Dim s As Shape
        For Each s In wsDash.Shapes
            If s.OnAction = "KernelFormHelpers.ToggleDevMode" Then
                Set shp = s
                Exit For
            End If
        Next s
    End If
    If Not shp Is Nothing Then
        On Error Resume Next
        If curMode = DEV_MODE_ON Then
            shp.TextFrame.Characters.Text = "Dev Mode: ON (click to hide)"
        Else
            shp.TextFrame.Characters.Text = "Dev Mode: OFF (click to show)"
        End If
        On Error GoTo 0
    End If
End Sub


' =============================================================================
' GenerateSummary
' Reads summary_config.csv from Config sheet. Rebuilds Summary tab with
' metrics grouped by SectionName, ordered by SortOrder.
' =============================================================================
Public Sub GenerateSummary()
    On Error GoTo ErrHandler

    Dim savedScreenUpdating As Boolean
    savedScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim sumCount As Long
    sumCount = KernelConfig.GetSummaryConfigCount()
    If sumCount = 0 Then
        KernelConfig.LogError SEV_WARN, "KernelTabs", "W-510", _
            "No summary_config data found. Summary not generated.", _
            "MANUAL BYPASS: Build Summary tab manually with SUMIFS formulas referencing Detail."
        Application.ScreenUpdating = savedScreenUpdating
        Exit Sub
    End If

    Dim wsSummary As Worksheet
    Set wsSummary = Nothing
    On Error Resume Next
    Set wsSummary = ThisWorkbook.Sheets(TAB_SUMMARY)
    On Error GoTo ErrHandler
    If wsSummary Is Nothing Then
        Application.ScreenUpdating = savedScreenUpdating
        Exit Sub
    End If

    On Error Resume Next
    wsSummary.Unprotect
    On Error GoTo ErrHandler

    wsSummary.Cells.ClearContents

    ' Determine current display mode and source tab
    Dim curMode As String
    curMode = KernelConfig.GetCurrentDisplayMode()
    Dim srcTab As String
    If curMode = DISPLAY_CUMULATIVE Then
        srcTab = TAB_CUMULATIVE_VIEW
    Else
        srcTab = TAB_DETAIL
    End If

    ' Detect entity count and period count
    Dim entityCount As Long
    entityCount = DetectEntityCountFromDetail()
    Dim periodCount As Long
    periodCount = KernelConfig.GetTimeHorizon()
    If entityCount = 0 Or periodCount = 0 Then
        KernelConfig.LogError SEV_WARN, "KernelTabs", "W-511", _
            "No data on Detail tab. Summary not generated.", _
            "MANUAL BYPASS: Run RunProjections first, then re-run GenerateSummary."
        Application.ScreenUpdating = savedScreenUpdating
        Exit Sub
    End If

    ' Get entity names
    Dim entityNames() As String
    ReDim entityNames(1 To entityCount)
    Dim eIdx As Long
    For eIdx = 1 To entityCount
        entityNames(eIdx) = KernelConfig.GetEntityName(eIdx)
    Next eIdx

    ' Build ordered metric list from summary_config
    Dim metricNames() As String
    Dim metricSections() As String
    Dim metricFormats() As String
    ReDim metricNames(1 To sumCount)
    ReDim metricSections(1 To sumCount)
    ReDim metricFormats(1 To sumCount)

    Dim visCount As Long
    visCount = 0
    Dim sIdx As Long
    For sIdx = 1 To sumCount
        If StrComp(KernelConfig.GetSummaryConfigField(sIdx, SUMCFG_COL_SHOW), "TRUE", vbTextCompare) = 0 Then
            visCount = visCount + 1
            metricNames(visCount) = KernelConfig.GetSummaryConfigField(sIdx, SUMCFG_COL_METRIC)
            metricSections(visCount) = KernelConfig.GetSummaryConfigField(sIdx, SUMCFG_COL_SECTION)
            metricFormats(visCount) = KernelConfig.GetSummaryConfigField(sIdx, SUMCFG_COL_FORMAT)
            If Len(metricFormats(visCount)) = 0 Then
                metricFormats(visCount) = KernelConfig.GetFormat(metricNames(visCount))
            End If
        End If
    Next sIdx

    If visCount = 0 Then
        Application.ScreenUpdating = savedScreenUpdating
        Exit Sub
    End If

    ' Column references for SUMIFS
    Dim entityColLetter As String
    entityColLetter = KernelOutput.ColLetter(KernelConfig.ColIndex("EntityName"))
    Dim periodCI As Long
    periodCI = KernelConfig.TryColIndex("Period")
    If periodCI < 1 Then periodCI = KernelConfig.ColIndex("CalPeriod")
    Dim periodColLetter As String
    periodColLetter = KernelOutput.ColLetter(periodCI)

    ' Row 1: Title
    wsSummary.Cells(1, 1).Value = "Model Summary"
    wsSummary.Range(wsSummary.Cells(1, 1), wsSummary.Cells(1, 3)).Merge
    With wsSummary.Cells(1, 1)
        .Font.Size = 14
        .Font.Bold = True
    End With

    ' Row 2: Metadata
    wsSummary.Cells(2, 1).Value = "Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & _
        "  |  Mode: " & curMode

    ' Row 3: blank

    ' Row 4: Column headers
    Dim headerRow As Long
    headerRow = 4
    wsSummary.Cells(headerRow, 1).Value = "Metric"
    Dim prd As Long
    For prd = 1 To periodCount
        wsSummary.Cells(headerRow, 1 + prd).Value = "Period " & prd
    Next prd
    wsSummary.Cells(headerRow, periodCount + 2).Value = "Total"

    With wsSummary.Range(wsSummary.Cells(headerRow, 1), wsSummary.Cells(headerRow, periodCount + 2))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    Dim curRow As Long
    curRow = headerRow + 1

    ' Track metric rows for derived field references
    Dim metricRows() As Long
    ReDim metricRows(1 To visCount)

    Dim lastSection As String
    lastSection = ""

    ' Write TOTAL section first
    wsSummary.Cells(curRow, 1).NumberFormat = "@"
    wsSummary.Cells(curRow, 1).Value = "=== TOTAL ==="
    wsSummary.Cells(curRow, 1).Font.Bold = True
    wsSummary.Cells(curRow, 1).Interior.Color = RGB(217, 225, 242)
    curRow = curRow + 1

    Dim mIdx As Long
    For mIdx = 1 To visCount
        wsSummary.Cells(curRow, 1).Value = metricNames(mIdx)
        metricRows(mIdx) = curRow

        Dim fc As String
        fc = KernelConfig.GetFieldClass(metricNames(mIdx))

        If fc = "Incremental" Then
            Dim metColLetter As String
            metColLetter = KernelOutput.ColLetter(KernelConfig.ColIndex(metricNames(mIdx)))
            For prd = 1 To periodCount
                wsSummary.Cells(curRow, 1 + prd).Formula = _
                    "=SUMIFS(" & srcTab & "!$" & metColLetter & ":$" & metColLetter & _
                    "," & srcTab & "!$" & periodColLetter & ":$" & periodColLetter & "," & prd & ")"
            Next prd
            ' Total column = SUM of period columns
            Dim firstDataCol As String
            firstDataCol = KernelOutput.ColLetter(2)
            Dim lastDataCol As String
            lastDataCol = KernelOutput.ColLetter(periodCount + 1)
            wsSummary.Cells(curRow, periodCount + 2).Formula = _
                "=SUM(" & firstDataCol & curRow & ":" & lastDataCol & curRow & ")"
        End If

        ' Apply format
        If Len(metricFormats(mIdx)) > 0 Then
            wsSummary.Range(wsSummary.Cells(curRow, 2), _
                wsSummary.Cells(curRow, periodCount + 2)).NumberFormat = metricFormats(mIdx)
        End If

        curRow = curRow + 1
    Next mIdx

    ' Fill derived formulas in TOTAL using cell references
    FillDerivedFormulas wsSummary, metricRows, metricNames, visCount, periodCount

    ' Blank row
    curRow = curRow + 1

    ' Entity sections
    For eIdx = 1 To entityCount
        Dim entMetricRows() As Long
        ReDim entMetricRows(1 To visCount)

        ' Section header
        wsSummary.Cells(curRow, 1).NumberFormat = "@"
        wsSummary.Cells(curRow, 1).Value = "=== " & entityNames(eIdx) & " ==="
        wsSummary.Cells(curRow, 1).Font.Bold = True
        wsSummary.Cells(curRow, 1).Interior.Color = RGB(217, 225, 242)
        curRow = curRow + 1

        For mIdx = 1 To visCount
            wsSummary.Cells(curRow, 1).Value = metricNames(mIdx)
            entMetricRows(mIdx) = curRow

            fc = KernelConfig.GetFieldClass(metricNames(mIdx))

            If fc = "Incremental" Then
                metColLetter = KernelOutput.ColLetter(KernelConfig.ColIndex(metricNames(mIdx)))
                For prd = 1 To periodCount
                    wsSummary.Cells(curRow, 1 + prd).Formula = _
                        "=SUMIFS(" & srcTab & "!$" & metColLetter & ":$" & metColLetter & _
                        "," & srcTab & "!$" & entityColLetter & ":$" & entityColLetter & ",""" & entityNames(eIdx) & """" & _
                        "," & srcTab & "!$" & periodColLetter & ":$" & periodColLetter & "," & prd & ")"
                Next prd
                firstDataCol = KernelOutput.ColLetter(2)
                lastDataCol = KernelOutput.ColLetter(periodCount + 1)
                wsSummary.Cells(curRow, periodCount + 2).Formula = _
                    "=SUM(" & firstDataCol & curRow & ":" & lastDataCol & curRow & ")"
            End If

            If Len(metricFormats(mIdx)) > 0 Then
                wsSummary.Range(wsSummary.Cells(curRow, 2), _
                    wsSummary.Cells(curRow, periodCount + 2)).NumberFormat = metricFormats(mIdx)
            End If

            curRow = curRow + 1
        Next mIdx

        ' Fill derived formulas for this entity
        FillDerivedFormulas wsSummary, entMetricRows, metricNames, visCount, periodCount

        ' Blank row between entities
        curRow = curRow + 1
    Next eIdx

    ' Freeze panes below header
    wsSummary.Activate
    wsSummary.Cells(headerRow + 1, 1).Select
    ActiveWindow.FreezePanes = True

    ' AutoFit deferred to caller (AutoFitAllOutputTabs for pipeline, manual for standalone)
    wsSummary.Protect UserInterfaceOnly:=True

    KernelConfig.LogError SEV_INFO, "KernelTabs", "I-510", _
        "Summary tab generated", visCount & " metrics, " & entityCount & " entities, mode=" & curMode

    MsgBox "Summary tab generated." & vbCrLf & _
        visCount & " metrics, " & entityCount & " entities." & vbCrLf & _
        "Display mode: " & curMode, vbInformation, "RDK -- Summary"

    Application.ScreenUpdating = savedScreenUpdating
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = savedScreenUpdating
    KernelConfig.LogError SEV_ERROR, "KernelTabs", "E-510", _
        "Error generating Summary: " & Err.Description, _
        "MANUAL BYPASS: Build Summary tab manually with SUMIFS formulas referencing Detail."
    MsgBox "Summary generation failed: " & Err.Description, vbExclamation, "RDK -- Error"
End Sub


' =============================================================================
' FillDerivedFormulas
' For derived metrics, writes cell-reference formulas instead of SUMIFS.
' =============================================================================
Private Sub FillDerivedFormulas(ws As Worksheet, ByRef rowArr() As Long, _
    ByRef nameArr() As String, cnt As Long, periodCount As Long)

    Dim mIdx As Long
    For mIdx = 1 To cnt
        Dim fc As String
        fc = KernelConfig.GetFieldClass(nameArr(mIdx))
        If fc <> "Derived" Then GoTo NextDeriv

        Dim rule As String
        rule = KernelConfig.GetDerivationRule(nameArr(mIdx))
        If Len(rule) = 0 Then GoTo NextDeriv

        Dim opA As String
        Dim opStr As String
        Dim opB As String
        If Not ParseRule(rule, opA, opStr, opB) Then GoTo NextDeriv

        Dim rowA As Long
        rowA = FindMetricInRows(rowArr, nameArr, cnt, opA)
        Dim rowB As Long
        rowB = FindMetricInRows(rowArr, nameArr, cnt, opB)
        If rowA = 0 Or rowB = 0 Then GoTo NextDeriv

        Dim prd As Long
        For prd = 1 To periodCount + 1
            Dim dataCol As Long
            dataCol = 1 + prd
            Dim cA As String
            cA = KernelOutput.ColLetter(dataCol) & rowA
            Dim cB As String
            cB = KernelOutput.ColLetter(dataCol) & rowB
            Select Case opStr
                Case "-"
                    ws.Cells(rowArr(mIdx), dataCol).Formula = "=" & cA & "-" & cB
                Case "+"
                    ws.Cells(rowArr(mIdx), dataCol).Formula = "=" & cA & "+" & cB
                Case "*"
                    ws.Cells(rowArr(mIdx), dataCol).Formula = "=" & cA & "*" & cB
                Case "/"
                    ws.Cells(rowArr(mIdx), dataCol).Formula = "=IFERROR(" & cA & "/" & cB & ",0)"
            End Select
        Next prd
NextDeriv:
    Next mIdx
End Sub


' =============================================================================
' GenerateCharts
' Reads chart_registry.csv. Creates charts on Charts tab.
' =============================================================================
Public Sub GenerateCharts()
    On Error GoTo ErrHandler

    Dim savedScreenUpdating As Boolean
    savedScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim chtCount As Long
    chtCount = KernelConfig.GetChartRegistryCount()
    If chtCount = 0 Then
        KernelConfig.LogError SEV_WARN, "KernelTabs", "W-520", _
            "No chart_registry data found. Charts not generated.", _
            "MANUAL BYPASS: Insert a chart manually on the Charts tab. Select data from Detail tab."
        Application.ScreenUpdating = savedScreenUpdating
        Exit Sub
    End If

    Dim wsCharts As Worksheet
    Set wsCharts = Nothing
    On Error Resume Next
    Set wsCharts = ThisWorkbook.Sheets(TAB_CHARTS)
    On Error GoTo ErrHandler
    ' DEPRECATED: TAB_CHARTS removed from tab_registry; no-op if tab absent
    If wsCharts Is Nothing Then
        Application.ScreenUpdating = savedScreenUpdating
        Exit Sub
    End If

    ' Clear existing charts
    ClearChartsOnSheet wsCharts

    ' Add title
    wsCharts.Cells(1, 1).Value = "Charts"
    With wsCharts.Cells(1, 1)
        .Font.Size = 14
        .Font.Bold = True
    End With

    ' Read Detail data into array for chart data source
    Dim wsDetail As Worksheet
    Dim curMode As String
    curMode = KernelConfig.GetCurrentDisplayMode()
    If curMode = DISPLAY_CUMULATIVE Then
        On Error Resume Next
        Set wsDetail = ThisWorkbook.Sheets(TAB_CUMULATIVE_VIEW)
        On Error GoTo ErrHandler
        If wsDetail Is Nothing Then Set wsDetail = ThisWorkbook.Sheets(TAB_DETAIL)
    Else
        Set wsDetail = ThisWorkbook.Sheets(TAB_DETAIL)
    End If

    Dim entityCount As Long
    entityCount = DetectEntityCountFromDetail()
    Dim periodCount As Long
    periodCount = KernelConfig.GetTimeHorizon()
    If entityCount = 0 Or periodCount = 0 Then
        KernelConfig.LogError SEV_WARN, "KernelTabs", "W-521", _
            "No data on Detail tab. Charts not generated.", _
            "MANUAL BYPASS: Run RunProjections first, then re-run GenerateCharts."
        Application.ScreenUpdating = savedScreenUpdating
        Exit Sub
    End If

    ' Get entity names
    Dim entityNames() As String
    ReDim entityNames(1 To entityCount)
    Dim eIdx As Long
    For eIdx = 1 To entityCount
        entityNames(eIdx) = KernelConfig.GetEntityName(eIdx)
    Next eIdx

    ' Read all data from Detail into array (PT-001)
    Dim totalRows As Long
    totalRows = entityCount * periodCount
    Dim totalCols As Long
    totalCols = KernelConfig.GetColumnCount()
    Dim detailData() As Variant
    detailData = wsDetail.Range(wsDetail.Cells(DETAIL_DATA_START_ROW, 1), _
        wsDetail.Cells(DETAIL_DATA_START_ROW + totalRows - 1, totalCols)).Value

    ' Chart positioning: 2 per row, flowing down
    ' Place charts below existing buttons (start at row 20)
    Dim chartLeft As Double
    Dim chartTop As Double
    Dim chartIdx As Long
    Dim chartsPlaced As Long
    chartsPlaced = 0
    Dim chartsSkipped As Long
    chartsSkipped = 0

    Dim baseTop As Double
    baseTop = wsCharts.Cells(3, 1).Top

    For chartIdx = 1 To chtCount
        If StrComp(KernelConfig.GetChartRegistryField(chartIdx, CHTCFG_COL_ENABLED), "TRUE", vbTextCompare) <> 0 Then
            GoTo NextChart
        End If

        Dim chtName As String
        chtName = KernelConfig.GetChartRegistryField(chartIdx, CHTCFG_COL_NAME)
        Dim chtType As String
        chtType = KernelConfig.GetChartRegistryField(chartIdx, CHTCFG_COL_TYPE)
        Dim chtMetric As String
        chtMetric = KernelConfig.GetChartRegistryField(chartIdx, CHTCFG_COL_METRIC)
        Dim chtGroupBy As String
        chtGroupBy = KernelConfig.GetChartRegistryField(chartIdx, CHTCFG_COL_GROUPBY)

        Dim chtWidth As Long
        Dim widthStr As String
        widthStr = KernelConfig.GetChartRegistryField(chartIdx, CHTCFG_COL_WIDTH)
        If IsNumeric(widthStr) And Len(widthStr) > 0 Then
            chtWidth = CLng(widthStr)
        Else
            chtWidth = 500
        End If

        Dim chtHeight As Long
        Dim heightStr As String
        heightStr = KernelConfig.GetChartRegistryField(chartIdx, CHTCFG_COL_HEIGHT)
        If IsNumeric(heightStr) And Len(heightStr) > 0 Then
            chtHeight = CLng(heightStr)
        Else
            chtHeight = 300
        End If

        ' Calculate position (2 per row)
        Dim colPos As Long
        colPos = chartsPlaced Mod 2
        Dim rowPos As Long
        rowPos = chartsPlaced \ 2

        chartLeft = 10 + colPos * (chtWidth + 20)
        chartTop = baseTop + rowPos * (chtHeight + 30)

        ' Get metric column index
        Dim metricCol As Long
        metricCol = KernelConfig.ColIndex(chtMetric)
        If metricCol < 1 Then
            chartsSkipped = chartsSkipped + 1
            KernelConfig.LogError SEV_ERROR, "KernelTabs", "E-521", _
                "Chart '" & chtName & "' skipped: metric '" & chtMetric & "' not found in column_registry.", _
                "MANUAL BYPASS: Fix MetricName in chart_registry.csv (Config sheet '=== CHART_REGISTRY ===' section), " & _
                "or insert chart manually on the Charts tab using data from Detail."
            GoTo NextChart
        End If

        Dim periodCol As Long
        periodCol = KernelConfig.TryColIndex("Period")
        If periodCol < 1 Then periodCol = KernelConfig.ColIndex("CalPeriod")
        Dim entityCol As Long
        entityCol = KernelConfig.ColIndex("EntityName")

        ' Create chart
        Dim chtObj As ChartObject
        Set chtObj = wsCharts.ChartObjects.Add(chartLeft, chartTop, chtWidth, chtHeight)

        With chtObj.Chart
            ' Set chart type
            Select Case chtType
                Case CHART_LINE: .ChartType = xlLine
                Case CHART_BAR: .ChartType = xlColumnClustered
                Case CHART_STACKED: .ChartType = xlColumnStacked
                Case CHART_PIE: .ChartType = xlPie
                Case CHART_AREA: .ChartType = xlArea
                Case Else: .ChartType = xlLine
            End Select

            .HasTitle = True
            .ChartTitle.Text = chtName

            ' Build series from data array
            If StrComp(chtGroupBy, "Entity", vbTextCompare) = 0 Then
                ' One series per entity, X axis = Period
                For eIdx = 1 To entityCount
                    Dim serValues() As Double
                    ReDim serValues(1 To periodCount)
                    Dim serCategories() As Long
                    ReDim serCategories(1 To periodCount)

                    Dim rIdx As Long
                    For rIdx = 1 To totalRows
                        If CStr(detailData(rIdx, entityCol)) = entityNames(eIdx) Then
                            Dim p As Long
                            If IsNumeric(detailData(rIdx, periodCol)) Then
                                p = CLng(detailData(rIdx, periodCol))
                                If p >= 1 And p <= periodCount Then
                                    serCategories(p) = p
                                    If IsNumeric(detailData(rIdx, metricCol)) Then
                                        serValues(p) = CDbl(detailData(rIdx, metricCol))
                                    End If
                                End If
                            End If
                        End If
                    Next rIdx

                    Dim newSeries As Series
                    Set newSeries = .SeriesCollection.NewSeries
                    newSeries.Name = entityNames(eIdx)
                    newSeries.Values = serValues
                    newSeries.XValues = serCategories
                Next eIdx
            Else
                ' One series per period, X axis = Entity (less common)
                Dim pIdx As Long
                For pIdx = 1 To periodCount
                    Dim pSerValues() As Double
                    ReDim pSerValues(1 To entityCount)
                    Dim pSerCats() As String
                    ReDim pSerCats(1 To entityCount)

                    For rIdx = 1 To totalRows
                        If IsNumeric(detailData(rIdx, periodCol)) Then
                            If CLng(detailData(rIdx, periodCol)) = pIdx Then
                                For eIdx = 1 To entityCount
                                    If CStr(detailData(rIdx, entityCol)) = entityNames(eIdx) Then
                                        pSerCats(eIdx) = entityNames(eIdx)
                                        If IsNumeric(detailData(rIdx, metricCol)) Then
                                            pSerValues(eIdx) = CDbl(detailData(rIdx, metricCol))
                                        End If
                                    End If
                                Next eIdx
                            End If
                        End If
                    Next rIdx

                    Set newSeries = .SeriesCollection.NewSeries
                    newSeries.Name = "Period " & pIdx
                    newSeries.Values = pSerValues
                    newSeries.XValues = pSerCats
                Next pIdx
            End If

            .HasLegend = True
        End With

        chartsPlaced = chartsPlaced + 1
NextChart:
    Next chartIdx

    KernelConfig.LogError SEV_INFO, "KernelTabs", "I-520", _
        "Charts generated on Charts tab", chartsPlaced & " placed, " & chartsSkipped & " skipped"

    Dim chartMsg As String
    chartMsg = chartsPlaced & " chart(s) generated on Charts tab."
    If chartsSkipped > 0 Then
        chartMsg = chartMsg & vbCrLf & chartsSkipped & " chart(s) skipped (invalid metric). Check ErrorLog for details." & _
            vbCrLf & vbCrLf & "MANUAL BYPASS: Fix MetricName in chart_registry.csv, or insert charts manually on Charts tab."
    End If
    MsgBox chartMsg, IIf(chartsSkipped > 0, vbExclamation, vbInformation), "RDK -- Charts"

    Application.ScreenUpdating = savedScreenUpdating
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = savedScreenUpdating
    KernelConfig.LogError SEV_ERROR, "KernelTabs", "E-520", _
        "Error generating charts: " & Err.Description, _
        "MANUAL BYPASS: Insert a chart manually on the Charts tab. Select data from Detail tab. Set chart type and formatting as desired."
    MsgBox "Chart generation failed: " & Err.Description, vbExclamation, "RDK -- Error"
End Sub


' =============================================================================
' GenerateExhibits
' Reads exhibit_config.csv. Creates exhibit tables on Exhibits tab.
' =============================================================================
Public Sub GenerateExhibits()
    On Error GoTo ErrHandler

    Dim savedScreenUpdating As Boolean
    savedScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim exhCount As Long
    exhCount = KernelConfig.GetExhibitConfigCount()
    If exhCount = 0 Then
        KernelConfig.LogError SEV_WARN, "KernelTabs", "W-530", _
            "No exhibit_config data found. Exhibits not generated.", _
            "MANUAL BYPASS: Create Exhibits tab manually with SUMIFS formulas referencing Detail."
        Application.ScreenUpdating = savedScreenUpdating
        Exit Sub
    End If

    Dim wsExh As Worksheet
    Set wsExh = Nothing
    On Error Resume Next
    Set wsExh = ThisWorkbook.Sheets(TAB_EXHIBITS)
    On Error GoTo ErrHandler
    ' DEPRECATED: TAB_EXHIBITS removed from tab_registry; no-op if tab absent
    If wsExh Is Nothing Then
        Application.ScreenUpdating = savedScreenUpdating
        Exit Sub
    End If

    On Error Resume Next
    wsExh.Unprotect
    On Error GoTo ErrHandler

    wsExh.Cells.ClearContents
    wsExh.Cells.ClearFormats

    ' Determine source tab (display mode aware)
    Dim curMode As String
    curMode = KernelConfig.GetCurrentDisplayMode()
    Dim srcTab As String
    If curMode = DISPLAY_CUMULATIVE Then
        Dim wsCumTest As Worksheet
        On Error Resume Next
        Set wsCumTest = ThisWorkbook.Sheets(TAB_CUMULATIVE_VIEW)
        On Error GoTo ErrHandler
        If wsCumTest Is Nothing Then
            srcTab = TAB_DETAIL
        Else
            srcTab = TAB_CUMULATIVE_VIEW
        End If
    Else
        srcTab = TAB_DETAIL
    End If

    Dim entityCount As Long
    entityCount = DetectEntityCountFromDetail()
    Dim periodCount As Long
    periodCount = KernelConfig.GetTimeHorizon()
    If entityCount = 0 Or periodCount = 0 Then
        KernelConfig.LogError SEV_WARN, "KernelTabs", "W-531", _
            "No data on Detail tab. Exhibits not generated.", _
            "MANUAL BYPASS: Run RunProjections first, then re-run GenerateExhibits."
        Application.ScreenUpdating = savedScreenUpdating
        Exit Sub
    End If

    ' Get entity names
    Dim entityNames() As String
    ReDim entityNames(1 To entityCount)
    Dim eIdx As Long
    For eIdx = 1 To entityCount
        entityNames(eIdx) = KernelConfig.GetEntityName(eIdx)
    Next eIdx

    ' Source tab column letters for SUMIFS formulas
    Dim entColLtr As String
    entColLtr = KernelOutput.ColLetter(KernelConfig.ColIndex("EntityName"))
    Dim prdCI As Long
    prdCI = KernelConfig.TryColIndex("Period")
    If prdCI < 1 Then prdCI = KernelConfig.ColIndex("CalPeriod")
    Dim prdColLtr As String
    prdColLtr = KernelOutput.ColLetter(prdCI)

    Dim curRow As Long
    curRow = 1
    Dim exhibitsPlaced As Long
    exhibitsPlaced = 0

    Dim exIdx As Long
    For exIdx = 1 To exhCount
        If StrComp(KernelConfig.GetExhibitConfigField(exIdx, EXHCFG_COL_ENABLED), "TRUE", vbTextCompare) <> 0 Then
            GoTo NextExhibit
        End If

        Dim exName As String
        exName = KernelConfig.GetExhibitConfigField(exIdx, EXHCFG_COL_NAME)
        Dim exMetricList As String
        exMetricList = KernelConfig.GetExhibitConfigField(exIdx, EXHCFG_COL_METRICS)
        Dim exGroupBy As String
        exGroupBy = KernelConfig.GetExhibitConfigField(exIdx, EXHCFG_COL_GROUPBY)
        Dim exIncTotal As Boolean
        exIncTotal = (StrComp(KernelConfig.GetExhibitConfigField(exIdx, EXHCFG_COL_TOTAL), "TRUE", vbTextCompare) = 0)
        Dim exFormat As String
        exFormat = KernelConfig.GetExhibitConfigField(exIdx, EXHCFG_COL_FORMAT)
        If Len(exFormat) = 0 Then exFormat = "#,##0"

        ' Parse metric list
        Dim metrics() As String
        metrics = Split(exMetricList, ",")
        Dim metCount As Long
        metCount = UBound(metrics) - LBound(metrics) + 1

        Dim mIdx As Long
        For mIdx = LBound(metrics) To UBound(metrics)
            metrics(mIdx) = Trim(metrics(mIdx))
        Next mIdx

        ' Build per-metric format array (summary_config -> column_registry -> exFormat)
        Dim metFmts() As String
        ReDim metFmts(0 To metCount - 1)
        For mIdx = 0 To metCount - 1
            Dim mfName As String
            mfName = metrics(LBound(metrics) + mIdx)
            Dim mfFmt As String
            mfFmt = ""
            Dim scIdx As Long
            For scIdx = 1 To KernelConfig.GetSummaryConfigCount()
                If StrComp(KernelConfig.GetSummaryConfigField(scIdx, SUMCFG_COL_METRIC), mfName, vbTextCompare) = 0 Then
                    mfFmt = KernelConfig.GetSummaryConfigField(scIdx, SUMCFG_COL_FORMAT)
                    Exit For
                End If
            Next scIdx
            If Len(mfFmt) = 0 Then mfFmt = KernelConfig.GetFormat(mfName)
            If Len(mfFmt) = 0 Then mfFmt = exFormat
            metFmts(mIdx) = mfFmt
        Next mIdx

        ' Exhibit header
        wsExh.Cells(curRow, 1).NumberFormat = "@"
        wsExh.Cells(curRow, 1).Value = "=== " & exName & " ==="
        wsExh.Cells(curRow, 1).Font.Bold = True
        wsExh.Cells(curRow, 1).Font.Size = 12
        wsExh.Cells(curRow, 1).Interior.Color = RGB(217, 225, 242)
        curRow = curRow + 1

        If StrComp(exGroupBy, "Period", vbTextCompare) = 0 Then
            ' GroupBy Period: for each entity, rows=periods, cols=metrics (SUMIFS formulas)
            For eIdx = 1 To entityCount
                wsExh.Cells(curRow, 1).Value = entityNames(eIdx)
                wsExh.Cells(curRow, 1).Font.Bold = True
                wsExh.Cells(curRow, 1).Font.Italic = True
                curRow = curRow + 1

                ' Column headers
                wsExh.Cells(curRow, 1).Value = "Period"
                wsExh.Cells(curRow, 1).Font.Bold = True
                For mIdx = 0 To metCount - 1
                    wsExh.Cells(curRow, 2 + mIdx).Value = metrics(LBound(metrics) + mIdx)
                    wsExh.Cells(curRow, 2 + mIdx).Font.Bold = True
                Next mIdx
                With wsExh.Range(wsExh.Cells(curRow, 1), wsExh.Cells(curRow, 1 + metCount))
                    .Interior.Color = RGB(68, 114, 196)
                    .Font.Color = RGB(255, 255, 255)
                End With
                curRow = curRow + 1

                Dim firstDataRow As Long
                firstDataRow = curRow

                Dim prd As Long
                For prd = 1 To periodCount
                    wsExh.Cells(curRow, 1).Value = prd
                    For mIdx = 0 To metCount - 1
                        Dim metName As String
                        metName = metrics(LBound(metrics) + mIdx)
                        Dim metFc As String
                        metFc = KernelConfig.GetFieldClass(metName)
                        If metFc = "Incremental" Then
                            Dim metLtr As String
                            metLtr = KernelOutput.ColLetter(KernelConfig.ColIndex(metName))
                            wsExh.Cells(curRow, 2 + mIdx).Formula = _
                                "=SUMIFS(" & srcTab & "!$" & metLtr & ":$" & metLtr & _
                                "," & srcTab & "!$" & entColLtr & ":$" & entColLtr & ",""" & entityNames(eIdx) & """" & _
                                "," & srcTab & "!$" & prdColLtr & ":$" & prdColLtr & "," & prd & ")"
                        End If
                        wsExh.Cells(curRow, 2 + mIdx).NumberFormat = metFmts(mIdx)
                    Next mIdx
                    ' Fill derived formulas (cell references within exhibit)
                    FillExhDerivedRow wsExh, curRow, metrics, metCount
                    curRow = curRow + 1
                Next prd

                ' Total row
                If exIncTotal Then
                    wsExh.Cells(curRow, 1).Value = "Total"
                    wsExh.Cells(curRow, 1).Font.Bold = True
                    For mIdx = 0 To metCount - 1
                        metFc = KernelConfig.GetFieldClass(metrics(LBound(metrics) + mIdx))
                        If metFc = "Incremental" Then
                            Dim fCol As String
                            fCol = KernelOutput.ColLetter(2 + mIdx)
                            wsExh.Cells(curRow, 2 + mIdx).Formula = _
                                "=SUM(" & fCol & firstDataRow & ":" & fCol & (curRow - 1) & ")"
                        End If
                        wsExh.Cells(curRow, 2 + mIdx).NumberFormat = metFmts(mIdx)
                    Next mIdx
                    FillExhDerivedRow wsExh, curRow, metrics, metCount
                    wsExh.Range(wsExh.Cells(curRow, 1), _
                        wsExh.Cells(curRow, 1 + metCount)).Borders(xlEdgeTop).LineStyle = xlContinuous
                    curRow = curRow + 1
                End If

                curRow = curRow + 1  ' blank row between entities
            Next eIdx

        Else
            ' GroupBy Entity: rows=entities, cols=metrics (SUMIFS formulas)
            wsExh.Cells(curRow, 1).Value = "Entity"
            wsExh.Cells(curRow, 1).Font.Bold = True
            For mIdx = 0 To metCount - 1
                wsExh.Cells(curRow, 2 + mIdx).Value = metrics(LBound(metrics) + mIdx)
                wsExh.Cells(curRow, 2 + mIdx).Font.Bold = True
            Next mIdx
            With wsExh.Range(wsExh.Cells(curRow, 1), wsExh.Cells(curRow, 1 + metCount))
                .Interior.Color = RGB(68, 114, 196)
                .Font.Color = RGB(255, 255, 255)
            End With
            curRow = curRow + 1

            Dim firstEntRow As Long
            firstEntRow = curRow

            For eIdx = 1 To entityCount
                wsExh.Cells(curRow, 1).Value = entityNames(eIdx)
                For mIdx = 0 To metCount - 1
                    metName = metrics(LBound(metrics) + mIdx)
                    metFc = KernelConfig.GetFieldClass(metName)
                    If metFc = "Incremental" Then
                        metLtr = KernelOutput.ColLetter(KernelConfig.ColIndex(metName))
                        wsExh.Cells(curRow, 2 + mIdx).Formula = _
                            "=SUMIFS(" & srcTab & "!$" & metLtr & ":$" & metLtr & _
                            "," & srcTab & "!$" & entColLtr & ":$" & entColLtr & ",""" & entityNames(eIdx) & """)"
                    End If
                    wsExh.Cells(curRow, 2 + mIdx).NumberFormat = metFmts(mIdx)
                Next mIdx
                FillExhDerivedRow wsExh, curRow, metrics, metCount
                curRow = curRow + 1
            Next eIdx

            ' Total row
            If exIncTotal Then
                wsExh.Cells(curRow, 1).Value = "Total"
                wsExh.Cells(curRow, 1).Font.Bold = True
                For mIdx = 0 To metCount - 1
                    metFc = KernelConfig.GetFieldClass(metrics(LBound(metrics) + mIdx))
                    If metFc = "Incremental" Then
                        fCol = KernelOutput.ColLetter(2 + mIdx)
                        wsExh.Cells(curRow, 2 + mIdx).Formula = _
                            "=SUM(" & fCol & firstEntRow & ":" & fCol & (curRow - 1) & ")"
                    End If
                    wsExh.Cells(curRow, 2 + mIdx).NumberFormat = metFmts(mIdx)
                Next mIdx
                FillExhDerivedRow wsExh, curRow, metrics, metCount
                wsExh.Range(wsExh.Cells(curRow, 1), _
                    wsExh.Cells(curRow, 1 + metCount)).Borders(xlEdgeTop).LineStyle = xlContinuous
                curRow = curRow + 1
            End If
        End If

        curRow = curRow + 1  ' blank row between exhibits
        exhibitsPlaced = exhibitsPlaced + 1
NextExhibit:
    Next exIdx

    wsExh.Columns.AutoFit
    wsExh.Protect UserInterfaceOnly:=True

    KernelConfig.LogError SEV_INFO, "KernelTabs", "I-530", _
        "Exhibits generated (formula-driven)", exhibitsPlaced & " exhibits"

    MsgBox exhibitsPlaced & " exhibit(s) generated on Exhibits tab." & vbCrLf & _
        "All values are native Excel formulas.", _
        vbInformation, "RDK -- Exhibits"

    Application.ScreenUpdating = savedScreenUpdating
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = savedScreenUpdating
    KernelConfig.LogError SEV_ERROR, "KernelTabs", "E-530", _
        "Error generating exhibits: " & Err.Description, _
        "MANUAL BYPASS: Create Exhibits tab manually with SUMIFS formulas referencing Detail."
    MsgBox "Exhibit generation failed: " & Err.Description, vbExclamation, "RDK -- Error"
End Sub


' =============================================================================
' ToggleDisplayMode
' Switches between Incremental and Cumulative display.
' =============================================================================
Public Sub ToggleDisplayMode()
    On Error GoTo ErrHandler

    Dim curMode As String
    curMode = KernelConfig.GetCurrentDisplayMode()

    Dim newMode As String
    If curMode = DISPLAY_INCREMENTAL Then
        newMode = DISPLAY_CUMULATIVE
    Else
        newMode = DISPLAY_INCREMENTAL
    End If

    ' If switching to Cumulative, ensure CumulativeView exists
    If newMode = DISPLAY_CUMULATIVE Then
        GenerateCumulativeView
    End If

    ' Store new mode
    KernelConfig.SetCurrentDisplayMode newMode

    ' Refresh all presentation artifacts
    GenerateSummary
    GenerateCharts
    GenerateExhibits
    UpdateDashboardMode

    ' Refresh Prove-It if exists
    On Error Resume Next
    Dim wsPI As Worksheet
    Set wsPI = ThisWorkbook.Sheets(TAB_PROVE_IT)
    If Not wsPI Is Nothing Then
        If Len(Trim(CStr(wsPI.Cells(5, 1).Value))) > 0 Then
            KernelProveIt.RefreshProveIt
        End If
    End If
    On Error GoTo ErrHandler

    KernelConfig.LogError SEV_INFO, "KernelTabs", "I-540", _
        "Display mode toggled to " & newMode, ""

    MsgBox "Display mode: " & newMode, vbInformation, "RDK -- Display Mode"

    Exit Sub

ErrHandler:
    KernelConfig.LogError SEV_ERROR, "KernelTabs", "E-540", _
        "Error toggling display mode: " & Err.Description, _
        "MANUAL BYPASS: Switch display mode manually by changing the 'DefaultMode' value on the Config sheet display_mode_config section. Re-run GenerateSummary."
    MsgBox "Display mode toggle failed: " & Err.Description, vbExclamation, "RDK -- Error"
End Sub


' =============================================================================
' GenerateCumulativeView
' Creates (or refreshes) the CumulativeView sheet with formulas referencing
' Detail. For Incremental columns: running SUM. For Derived: recomputed.
' Detail tab is NEVER modified (AP-47).
' =============================================================================
Public Sub GenerateCumulativeView()
    On Error GoTo ErrHandler

    Dim savedScreenUpdating As Boolean
    savedScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim wsCum As Worksheet
    Set wsCum = Nothing
    On Error Resume Next
    Set wsCum = ThisWorkbook.Sheets(TAB_CUMULATIVE_VIEW)
    On Error GoTo ErrHandler
    ' DEPRECATED: TAB_CUMULATIVE_VIEW removed from tab_registry; no-op if tab absent
    If wsCum Is Nothing Then
        Application.ScreenUpdating = savedScreenUpdating
        Exit Sub
    End If

    wsCum.Cells.ClearContents

    Dim entityCount As Long
    entityCount = DetectEntityCountFromDetail()
    Dim periodCount As Long
    periodCount = KernelConfig.GetTimeHorizon()
    If entityCount = 0 Or periodCount = 0 Then
        Application.ScreenUpdating = savedScreenUpdating
        Exit Sub
    End If

    Dim totalRows As Long
    totalRows = entityCount * periodCount
    Dim totalCols As Long
    totalCols = KernelConfig.GetColumnCount()

    ' Copy headers from Detail
    Dim wsDetail As Worksheet
    Set wsDetail = ThisWorkbook.Sheets(TAB_DETAIL)
    wsDetail.Range(wsDetail.Cells(DETAIL_HEADER_ROW, 1), _
        wsDetail.Cells(DETAIL_HEADER_ROW, totalCols)).Copy _
        Destination:=wsCum.Cells(DETAIL_HEADER_ROW, 1)

    ' For each data row, write formulas
    Dim rIdx As Long
    For rIdx = 1 To totalRows
        Dim detRow As Long
        detRow = DETAIL_DATA_START_ROW + rIdx - 1

        ' Determine entity block start row (first row for this entity)
        Dim entityBlockStart As Long
        entityBlockStart = DETAIL_DATA_START_ROW + ((rIdx - 1) \ periodCount) * periodCount

        Dim cIdx As Long
        For cIdx = 1 To totalCols
            Dim colName As String
            colName = KernelConfig.GetColName(cIdx)
            Dim detCol As Long
            detCol = KernelConfig.GetDetailCol(cIdx)
            Dim fc As String
            fc = KernelConfig.GetFieldClass(colName)
            Dim colLtr As String
            colLtr = KernelOutput.ColLetter(detCol)

            Select Case fc
                Case "Dimension"
                    ' Pass through from Detail
                    wsCum.Cells(detRow, detCol).Formula = "=Detail!" & colLtr & detRow

                Case "Incremental"
                    ' Running SUM from entity block start to current row
                    wsCum.Cells(detRow, detCol).Formula = _
                        "=SUM(Detail!" & colLtr & entityBlockStart & ":Detail!" & colLtr & detRow & ")"

                Case "Derived"
                    ' Recompute from cumulative incremental values
                    Dim rule As String
                    rule = KernelConfig.GetDerivationRule(colName)
                    If Len(rule) > 0 Then
                        Dim opA As String
                        Dim opStr As String
                        Dim opB As String
                        If ParseRule(rule, opA, opStr, opB) Then
                            Dim colA As Long
                            colA = KernelConfig.ColIndex(opA)
                            Dim colB As Long
                            colB = KernelConfig.ColIndex(opB)
                            If colA > 0 And colB > 0 Then
                                Dim ltrA As String
                                ltrA = KernelOutput.ColLetter(colA)
                                Dim ltrB As String
                                ltrB = KernelOutput.ColLetter(colB)
                                Dim refA As String
                                refA = ltrA & detRow
                                Dim refB As String
                                refB = ltrB & detRow
                                Select Case opStr
                                    Case "-"
                                        wsCum.Cells(detRow, detCol).Formula = "=" & refA & "-" & refB
                                    Case "+"
                                        wsCum.Cells(detRow, detCol).Formula = "=" & refA & "+" & refB
                                    Case "*"
                                        wsCum.Cells(detRow, detCol).Formula = "=" & refA & "*" & refB
                                    Case "/"
                                        wsCum.Cells(detRow, detCol).Formula = _
                                            "=IFERROR(" & refA & "/" & refB & ",0)"
                                End Select
                            End If
                        End If
                    End If
            End Select
        Next cIdx
    Next rIdx

    ' Apply number formats from column registry
    For cIdx = 1 To totalCols
        colName = KernelConfig.GetColName(cIdx)
        Dim fmtStr As String
        fmtStr = KernelConfig.GetFormat(colName)
        If Len(fmtStr) > 0 Then
            detCol = KernelConfig.GetDetailCol(cIdx)
            wsCum.Range(wsCum.Cells(DETAIL_DATA_START_ROW, detCol), _
                wsCum.Cells(DETAIL_DATA_START_ROW + totalRows - 1, detCol)).NumberFormat = fmtStr
        End If
    Next cIdx

    KernelConfig.LogError SEV_INFO, "KernelTabs", "I-550", _
        "CumulativeView sheet generated", totalRows & " rows"

    Application.ScreenUpdating = savedScreenUpdating
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = savedScreenUpdating
    KernelConfig.LogError SEV_ERROR, "KernelTabs", "E-550", _
        "Error generating CumulativeView: " & Err.Description, _
        "MANUAL BYPASS: Create CumulativeView sheet manually with running SUM formulas referencing Detail."
End Sub


' =============================================================================
' ClearCharts
' Removes all chart objects from the Charts tab.
' =============================================================================
Public Sub ClearCharts()
    On Error Resume Next
    Dim wsCharts As Worksheet
    Set wsCharts = ThisWorkbook.Sheets(TAB_CHARTS)
    If Not wsCharts Is Nothing Then
        ClearChartsOnSheet wsCharts
    End If
    On Error GoTo 0
End Sub


' =============================================================================
' ClearExhibits
' Clears the Exhibits tab content.
' =============================================================================
Public Sub ClearExhibits()
    On Error Resume Next
    Dim wsExh As Worksheet
    Set wsExh = ThisWorkbook.Sheets(TAB_EXHIBITS)
    If Not wsExh Is Nothing Then
        wsExh.Unprotect
        wsExh.Cells.ClearContents
    End If
    On Error GoTo 0
End Sub


' =============================================================================
' Private Helpers
' =============================================================================

Private Sub ClearChartsOnSheet(ws As Worksheet)
    Dim chtObj As ChartObject
    For Each chtObj In ws.ChartObjects
        chtObj.Delete
    Next chtObj
End Sub


Private Function DetectEntityCountFromDetail() As Long
    On Error Resume Next
    Dim wsDetail As Worksheet
    Set wsDetail = ThisWorkbook.Sheets(TAB_DETAIL)
    If wsDetail Is Nothing Then
        DetectEntityCountFromDetail = 0
        Exit Function
    End If
    On Error GoTo 0

    Dim entityCol As Long
    entityCol = KernelConfig.ColIndex("EntityName")
    If entityCol < 1 Then
        DetectEntityCountFromDetail = 0
        Exit Function
    End If

    ' Count unique entities in Detail
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim r As Long
    r = DETAIL_DATA_START_ROW
    Do While Len(Trim(CStr(wsDetail.Cells(r, entityCol).Value))) > 0
        Dim eName As String
        eName = Trim(CStr(wsDetail.Cells(r, entityCol).Value))
        If Not dict.Exists(eName) Then dict.Add eName, 1
        r = r + 1
    Loop
    DetectEntityCountFromDetail = dict.Count
End Function


Private Function ParseRule(rule As String, ByRef opA As String, _
    ByRef opStr As String, ByRef opB As String) As Boolean
    ParseRule = False
    Dim ops As Variant
    ops = Array(" - ", " + ", " * ", " / ")
    Dim i As Long
    For i = LBound(ops) To UBound(ops)
        Dim pos As Long
        pos = InStr(1, rule, ops(i), vbTextCompare)
        If pos > 0 Then
            opA = Trim(Mid(rule, 1, pos - 1))
            opStr = Trim(CStr(ops(i)))
            opB = Trim(Mid(rule, pos + Len(CStr(ops(i)))))
            If Len(opA) > 0 And Len(opB) > 0 Then
                ParseRule = True
                Exit Function
            End If
        End If
    Next i
End Function


Private Function FindMetricInRows(ByRef rowArr() As Long, ByRef nameArr() As String, _
    cnt As Long, metricName As String) As Long
    FindMetricInRows = 0
    Dim i As Long
    For i = 1 To cnt
        If StrComp(nameArr(i), metricName, vbTextCompare) = 0 Then
            FindMetricInRows = rowArr(i)
            Exit Function
        End If
    Next i
End Function


' FillExhDerivedRow - Writes cell-reference formulas for derived metrics in an exhibit row.
' Uses operand columns within the same exhibit row (no hardcoded values).
Private Sub FillExhDerivedRow(ws As Worksheet, row As Long, _
    ByRef metrics() As String, metCount As Long)
    Dim mIdx As Long
    For mIdx = 0 To metCount - 1
        Dim metName As String
        metName = metrics(LBound(metrics) + mIdx)
        If KernelConfig.GetFieldClass(metName) <> "Derived" Then GoTo NextExhD

        Dim rule As String
        rule = KernelConfig.GetDerivationRule(metName)
        If Len(rule) = 0 Then GoTo NextExhD

        Dim opA As String
        Dim opStr As String
        Dim opB As String
        If Not ParseRule(rule, opA, opStr, opB) Then GoTo NextExhD

        Dim idxA As Long
        idxA = FindExhMetIdx(opA, metrics, metCount)
        Dim idxB As Long
        idxB = FindExhMetIdx(opB, metrics, metCount)
        If idxA < 0 Or idxB < 0 Then GoTo NextExhD

        Dim cA As String
        cA = KernelOutput.ColLetter(2 + idxA) & row
        Dim cB As String
        cB = KernelOutput.ColLetter(2 + idxB) & row

        Select Case opStr
            Case "-": ws.Cells(row, 2 + mIdx).Formula = "=" & cA & "-" & cB
            Case "+": ws.Cells(row, 2 + mIdx).Formula = "=" & cA & "+" & cB
            Case "*": ws.Cells(row, 2 + mIdx).Formula = "=" & cA & "*" & cB
            Case "/": ws.Cells(row, 2 + mIdx).Formula = "=IFERROR(" & cA & "/" & cB & ",0)"
        End Select
NextExhD:
    Next mIdx
End Sub


Private Function FindExhMetIdx(metName As String, ByRef metrics() As String, _
    metCount As Long) As Long
    FindExhMetIdx = -1
    Dim i As Long
    For i = 0 To metCount - 1
        If StrComp(metrics(LBound(metrics) + i), metName, vbTextCompare) = 0 Then
            FindExhMetIdx = i
            Exit Function
        End If
    Next i
End Function
