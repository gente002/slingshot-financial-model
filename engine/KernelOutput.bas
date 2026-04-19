Attribute VB_Name = "KernelOutput"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelOutput.bas
' Purpose: Writes computation results to the Detail tab and generates
'          SUMIFS formulas on the Summary tab.
' =============================================================================


' =============================================================================
' WriteDetailTab
' Writes dimension and fact columns to the Detail sheet using array batch write.
' =============================================================================
Public Sub WriteDetailTab(ByRef outputs() As Variant, _
                          totalRows As Long)

    Dim wsDetail As Worksheet
    Set wsDetail = ThisWorkbook.Sheets(TAB_DETAIL)

    ' Unprotect if protected
    On Error Resume Next
    wsDetail.Unprotect
    On Error GoTo 0

    ' Clear existing data (preserve nothing - full rewrite)
    wsDetail.Cells.ClearContents

    Dim totalCols As Long
    totalCols = KernelConfig.GetColumnCount()

    ' Write headers in row 1
    Dim headerArr() As Variant
    ReDim headerArr(1 To 1, 1 To totalCols)

    Dim colIdx As Long
    For colIdx = 1 To totalCols
        headerArr(1, KernelConfig.GetDetailCol(colIdx)) = KernelConfig.GetColName(colIdx)
    Next colIdx

    wsDetail.Range(wsDetail.Cells(DETAIL_HEADER_ROW, 1), _
                   wsDetail.Cells(DETAIL_HEADER_ROW, totalCols)).Value = headerArr

    ' Format header row
    With wsDetail.Range(wsDetail.Cells(DETAIL_HEADER_ROW, 1), _
                        wsDetail.Cells(DETAIL_HEADER_ROW, totalCols))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Write data using array batch write (PT-001)
    If totalRows > 0 Then
        wsDetail.Range(wsDetail.Cells(DETAIL_DATA_START_ROW, 1), _
                       wsDetail.Cells(DETAIL_DATA_START_ROW + totalRows - 1, totalCols)).Value = outputs
    End If

    ' Apply number formats from ColumnRegistry
    For colIdx = 1 To totalCols
        Dim fmtStr As String
        fmtStr = KernelConfig.GetFormat(KernelConfig.GetColName(colIdx))
        If Len(fmtStr) > 0 Then
            Dim detCol As Long
            detCol = KernelConfig.GetDetailCol(colIdx)
            wsDetail.Range(wsDetail.Cells(DETAIL_DATA_START_ROW, detCol), _
                           wsDetail.Cells(DETAIL_DATA_START_ROW + totalRows - 1, detCol)).NumberFormat = fmtStr
        End If
    Next colIdx

    ' AutoFit deferred to AutoFitAllOutputTabs in KernelEngine Cleanup

    ' Re-protect output tab (AD-04)
    wsDetail.Protect UserInterfaceOnly:=True

    KernelConfig.LogError SEV_INFO, "KernelOutput", "I-300", _
                          "Detail tab written", totalRows & " rows, " & totalCols & " columns"
End Sub


' =============================================================================
' WriteSummaryFormulas
' Generates the Summary tab with SUMIFS formulas referencing Detail.
' =============================================================================
Public Sub WriteSummaryFormulas(entityCount As Long, maxPeriod As Long)
    Dim wsSummary As Worksheet
    Set wsSummary = Nothing
    On Error Resume Next
    Set wsSummary = ThisWorkbook.Sheets(TAB_SUMMARY)
    On Error GoTo 0
    If wsSummary Is Nothing Then Exit Sub

    ' Unprotect if protected
    On Error Resume Next
    wsSummary.Unprotect
    On Error GoTo 0

    ' Clear existing content
    wsSummary.Cells.ClearContents

    ' Get metric columns (Incremental + Derived)
    Dim incCols As Variant
    incCols = KernelConfig.GetIncrementalColumns()

    Dim derCols As Variant
    derCols = KernelConfig.GetDerivedColumns()

    ' Build combined metric list maintaining registry order
    Dim metricNames() As String
    Dim metricClasses() As String
    Dim metricCount As Long
    metricCount = 0

    ' Count metrics
    Dim totalCols As Long
    totalCols = KernelConfig.GetColumnCount()
    Dim regIdx As Long
    For regIdx = 1 To totalCols
        Dim fc As String
        fc = KernelConfig.GetFieldClass(KernelConfig.GetColName(regIdx))
        If fc = "Incremental" Or fc = "Derived" Then
            metricCount = metricCount + 1
        End If
    Next regIdx

    If metricCount = 0 Then Exit Sub

    ReDim metricNames(1 To metricCount)
    ReDim metricClasses(1 To metricCount)
    Dim mIdx As Long
    mIdx = 0
    For regIdx = 1 To totalCols
        fc = KernelConfig.GetFieldClass(KernelConfig.GetColName(regIdx))
        If fc = "Incremental" Or fc = "Derived" Then
            mIdx = mIdx + 1
            metricNames(mIdx) = KernelConfig.GetColName(regIdx)
            metricClasses(mIdx) = fc
        End If
    Next regIdx

    ' Get entity names from Inputs
    Dim entityNames() As String
    ReDim entityNames(1 To entityCount)
    Dim entIdx As Long
    For entIdx = 1 To entityCount
        entityNames(entIdx) = KernelConfig.GetEntityName(entIdx)
    Next entIdx

    ' Column references for SUMIFS
    Dim entityColLetter As String
    entityColLetter = ColLetter(KernelConfig.ColIndex("EntityName"))

    Dim periodColIdx As Long
    periodColIdx = KernelConfig.TryColIndex("Period")
    If periodColIdx < 1 Then periodColIdx = KernelConfig.ColIndex("CalPeriod")
    Dim periodColLetter As String
    periodColLetter = ColLetter(periodColIdx)

    ' ---- Row 1: Headers ----
    wsSummary.Cells(1, SUMMARY_COL_ENTITY).Value = "Entity"
    wsSummary.Cells(1, SUMMARY_COL_METRIC).Value = "Metric"
    Dim prd As Long
    For prd = 1 To maxPeriod
        wsSummary.Cells(1, SUMMARY_DATA_START_COL + prd - 1).Value = "Period " & prd
    Next prd

    ' Format header row
    With wsSummary.Range(wsSummary.Cells(1, SUMMARY_COL_ENTITY), wsSummary.Cells(1, SUMMARY_DATA_START_COL + maxPeriod - 1))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' ---- Row 2: TOTAL section ----
    Dim curRow As Long
    curRow = 2

    ' For TOTAL, we need one row per metric
    ' Track row positions for derived field references
    Dim totalMetricRows() As Long
    ReDim totalMetricRows(1 To metricCount)

    For mIdx = 1 To metricCount
        wsSummary.Cells(curRow, SUMMARY_COL_ENTITY).Value = "TOTAL"
        wsSummary.Cells(curRow, SUMMARY_COL_METRIC).Value = metricNames(mIdx)
        totalMetricRows(mIdx) = curRow

        If metricClasses(mIdx) = "Incremental" Then
            ' SUMIFS formula for each period
            Dim metColLetter As String
            metColLetter = ColLetter(KernelConfig.ColIndex(metricNames(mIdx)))

            For prd = 1 To maxPeriod
                Dim sumCol As Long
                sumCol = SUMMARY_DATA_START_COL + prd - 1
                wsSummary.Cells(curRow, sumCol).Formula = _
                    "=SUMIFS(Detail!$" & metColLetter & ":$" & metColLetter & _
                    ",Detail!$" & periodColLetter & ":$" & periodColLetter & "," & prd & ")"
            Next prd
        End If
        ' Derived fields in TOTAL will be handled after all incremental rows are placed

        curRow = curRow + 1
    Next mIdx

    ' Now fill in TOTAL derived formulas using cell references
    Dim tmIdx As Long
    For tmIdx = 1 To metricCount
        If metricClasses(tmIdx) = "Derived" Then
            Dim derivRule As String
            derivRule = KernelConfig.GetDerivationRule(metricNames(tmIdx))
            If Len(derivRule) > 0 Then
                Dim opA As String
                Dim opStr As String
                Dim opB As String
                If ParseRuleForSummary(derivRule, opA, opStr, opB) Then
                    Dim rowA As Long
                    rowA = FindMetricRow(totalMetricRows, metricNames, metricCount, opA)
                    Dim rowB As Long
                    rowB = FindMetricRow(totalMetricRows, metricNames, metricCount, opB)

                    If rowA > 0 And rowB > 0 Then
                        For prd = 1 To maxPeriod
                            sumCol = SUMMARY_DATA_START_COL + prd - 1
                            Dim cellA As String
                            cellA = ColLetter(sumCol) & rowA
                            Dim cellB As String
                            cellB = ColLetter(sumCol) & rowB

                            Select Case opStr
                                Case "-"
                                    wsSummary.Cells(totalMetricRows(tmIdx), sumCol).Formula = _
                                        "=" & cellA & "-" & cellB
                                Case "+"
                                    wsSummary.Cells(totalMetricRows(tmIdx), sumCol).Formula = _
                                        "=" & cellA & "+" & cellB
                                Case "*"
                                    wsSummary.Cells(totalMetricRows(tmIdx), sumCol).Formula = _
                                        "=" & cellA & "*" & cellB
                                Case "/"
                                    wsSummary.Cells(totalMetricRows(tmIdx), sumCol).Formula = _
                                        "=IFERROR(" & cellA & "/" & cellB & ",0)"
                            End Select
                        Next prd
                    End If
                End If
            End If
        End If
    Next tmIdx

    ' ---- Row after TOTAL: blank separator ----
    curRow = curRow + 1

    ' ---- Entity sections ----
    For entIdx = 1 To entityCount
        Dim entityMetricRows() As Long
        ReDim entityMetricRows(1 To metricCount)

        For mIdx = 1 To metricCount
            wsSummary.Cells(curRow, SUMMARY_COL_ENTITY).Value = entityNames(entIdx)
            wsSummary.Cells(curRow, SUMMARY_COL_METRIC).Value = metricNames(mIdx)
            entityMetricRows(mIdx) = curRow

            If metricClasses(mIdx) = "Incremental" Then
                metColLetter = ColLetter(KernelConfig.ColIndex(metricNames(mIdx)))

                For prd = 1 To maxPeriod
                    sumCol = SUMMARY_DATA_START_COL + prd - 1
                    wsSummary.Cells(curRow, sumCol).Formula = _
                        "=SUMIFS(Detail!$" & metColLetter & ":$" & metColLetter & _
                        ",Detail!$" & entityColLetter & ":$" & entityColLetter & ",""" & entityNames(entIdx) & """" & _
                        ",Detail!$" & periodColLetter & ":$" & periodColLetter & "," & prd & ")"
                Next prd
            End If

            curRow = curRow + 1
        Next mIdx

        ' Fill in derived formulas for this entity section
        For tmIdx = 1 To metricCount
            If metricClasses(tmIdx) = "Derived" Then
                derivRule = KernelConfig.GetDerivationRule(metricNames(tmIdx))
                If Len(derivRule) > 0 Then
                    If ParseRuleForSummary(derivRule, opA, opStr, opB) Then
                        rowA = FindMetricRow(entityMetricRows, metricNames, metricCount, opA)
                        rowB = FindMetricRow(entityMetricRows, metricNames, metricCount, opB)

                        If rowA > 0 And rowB > 0 Then
                            For prd = 1 To maxPeriod
                                sumCol = SUMMARY_DATA_START_COL + prd - 1
                                cellA = ColLetter(sumCol) & rowA
                                cellB = ColLetter(sumCol) & rowB

                                Select Case opStr
                                    Case "-"
                                        wsSummary.Cells(entityMetricRows(tmIdx), sumCol).Formula = _
                                            "=" & cellA & "-" & cellB
                                    Case "+"
                                        wsSummary.Cells(entityMetricRows(tmIdx), sumCol).Formula = _
                                            "=" & cellA & "+" & cellB
                                    Case "*"
                                        wsSummary.Cells(entityMetricRows(tmIdx), sumCol).Formula = _
                                            "=" & cellA & "*" & cellB
                                    Case "/"
                                        wsSummary.Cells(entityMetricRows(tmIdx), sumCol).Formula = _
                                            "=IFERROR(" & cellA & "/" & cellB & ",0)"
                                End Select
                            Next prd
                        End If
                    End If
                End If
            End If
        Next tmIdx

        ' Blank row between entities
        curRow = curRow + 1
    Next entIdx

    ' Apply number formats
    For mIdx = 1 To metricCount
        Dim nf As String
        nf = KernelConfig.GetFormat(metricNames(mIdx))
        If Len(nf) > 0 Then
            ' Apply to all data columns for all rows with this metric
            ' (format entire columns 3+ through maxPeriod)
        End If
    Next mIdx

    ' AutoFit deferred to AutoFitAllOutputTabs in KernelEngine Cleanup

    ' Protect output tab (AD-04)
    wsSummary.Protect UserInterfaceOnly:=True

    KernelConfig.LogError SEV_INFO, "KernelOutput", "I-310", _
                          "Summary tab written", entityCount & " entities, " & metricCount & " metrics"
End Sub


' =============================================================================
' ExportSchemaTemplate
' Writes detail_template.csv (empty CSV with headers) and detail_template.txt
' (human-readable column documentation) to the config/ directory.
' =============================================================================
Public Sub ExportSchemaTemplate()
    On Error GoTo ErrHandler

    ' Ensure config is loaded
    Dim totalCols As Long
    totalCols = KernelConfig.GetColumnCount()

    If totalCols = 0 Then
        KernelConfig.LogError SEV_ERROR, "KernelOutput", "E-320", _
                              "Cannot export schema: no columns loaded", _
                              "MANUAL BYPASS: Run LoadAllConfig first, or verify the Config sheet has a valid COLUMN_REGISTRY section."
        Exit Sub
    End If

    Dim configDir As String
    Dim wbPath As String
    wbPath = ThisWorkbook.Path
    Dim parentDir As String
    parentDir = Left(wbPath, InStrRev(wbPath, "\") - 1)
    configDir = parentDir & "\config"

    If Dir(configDir, vbDirectory) = "" Then
        MkDir configDir
    End If

    ' --- detail_template.csv ---
    Dim csvPath As String
    csvPath = configDir & "\detail_template.csv"

    Dim fileNum As Integer
    fileNum = FreeFile

    Open csvPath For Output As #fileNum

    ' Write header row ordered by CsvIndex
    Dim headers() As String
    ReDim headers(0 To totalCols - 1)

    Dim regIdx As Long
    For regIdx = 1 To totalCols
        Dim colName As String
        colName = KernelConfig.GetColName(regIdx)
        Dim csvCol As Long
        csvCol = KernelConfig.CsvIndex(colName)
        If csvCol >= 0 And csvCol < totalCols Then
            headers(csvCol) = colName
        End If
    Next regIdx

    Dim headerLine As String
    headerLine = ""
    Dim cidx As Long
    For cidx = 0 To totalCols - 1
        If cidx > 0 Then headerLine = headerLine & ","
        headerLine = headerLine & headers(cidx)
    Next cidx

    Print #fileNum, headerLine
    Close #fileNum

    ' --- detail_template.txt ---
    Dim txtPath As String
    txtPath = configDir & "\detail_template.txt"

    fileNum = FreeFile
    Open txtPath For Output As #fileNum

    Print #fileNum, "RDK Detail Tab Schema Reference"
    Print #fileNum, "================================"
    Print #fileNum, "Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Print #fileNum, "Kernel Version: " & KERNEL_VERSION
    Print #fileNum, ""
    Print #fileNum, "Columns (" & totalCols & " total):"
    Print #fileNum, String(60, "-")

    For regIdx = 1 To totalCols
        colName = KernelConfig.GetColName(regIdx)
        Dim fc As String
        fc = KernelConfig.GetFieldClass(colName)
        Dim fmt As String
        fmt = KernelConfig.GetFormat(colName)
        Dim deriv As String
        deriv = KernelConfig.GetDerivationRule(colName)
        Dim detCol As Long
        detCol = KernelConfig.GetDetailCol(regIdx)

        Dim line As String
        line = "  " & regIdx & ". " & colName
        line = line & "  [" & fc & "]"
        line = line & "  DetailCol=" & detCol
        csvCol = KernelConfig.CsvIndex(colName)
        line = line & "  CsvCol=" & csvCol

        If Len(fmt) > 0 Then line = line & "  Format=" & fmt
        If Len(deriv) > 0 Then line = line & "  Rule=" & deriv

        Print #fileNum, line
    Next regIdx

    Print #fileNum, ""
    Print #fileNum, "Field Classes:"
    Print #fileNum, "  Dimension   - Entity/time identifiers (written by domain)"
    Print #fileNum, "  Incremental - Period-level metrics (written by domain)"
    Print #fileNum, "  Derived     - Auto-computed by kernel from DerivationRule"

    Close #fileNum

    KernelConfig.LogError SEV_INFO, "KernelOutput", "I-320", _
                          "Schema templates exported", csvPath & " and " & txtPath

    Exit Sub

ErrHandler:
    KernelConfig.LogError SEV_ERROR, "KernelOutput", "E-329", _
                          "Error exporting schema template: " & Err.Description, _
                          "MANUAL BYPASS: Manually create detail_template.csv with headers from the Config sheet COLUMN_REGISTRY."
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
End Sub


' =============================================================================
' ColLetter
' Converts a 1-based column number to an Excel column letter (1=A, 26=Z, 27=AA).
' =============================================================================
Public Function ColLetter(colNum As Long) As String
    Dim result As String
    Dim num As Long
    num = colNum

    result = ""
    Do While num > 0
        Dim remainder As Long
        remainder = (num - 1) Mod 26
        result = Chr(65 + remainder) & result
        num = (num - 1) \ 26
    Loop

    ColLetter = result
End Function


' =============================================================================
' ParseRuleForSummary
' Parses a DerivationRule for use in Summary formula generation.
' =============================================================================
Private Function ParseRuleForSummary(rule As String, ByRef opA As String, _
                                     ByRef opStr As String, ByRef opB As String) As Boolean
    ParseRuleForSummary = False

    Dim ops As Variant
    ops = Array(" - ", " + ", " * ", " / ")

    Dim opIdx As Long
    For opIdx = LBound(ops) To UBound(ops)
        Dim pos As Long
        pos = InStr(1, rule, ops(opIdx), vbTextCompare)
        If pos > 0 Then
            opA = Trim(Mid(rule, 1, pos - 1))
            opStr = Trim(CStr(ops(opIdx)))
            opB = Trim(Mid(rule, pos + Len(CStr(ops(opIdx)))))

            If Len(opA) > 0 And Len(opB) > 0 Then
                ParseRuleForSummary = True
                Exit Function
            End If
        End If
    Next opIdx
End Function


' =============================================================================
' FindMetricRow
' Finds the row number for a metric name within a row-tracking array.
' =============================================================================
Private Function FindMetricRow(ByRef rowArr() As Long, ByRef nameArr() As String, _
                               cnt As Long, metricName As String) As Long
    FindMetricRow = 0
    Dim idx As Long
    For idx = 1 To cnt
        If StrComp(nameArr(idx), metricName, vbTextCompare) = 0 Then
            FindMetricRow = rowArr(idx)
            Exit Function
        End If
    Next idx
End Function
