Attribute VB_Name = "Ins_QuarterlyAgg"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' Ins_QuarterlyAgg.bas
' Purpose: Quarterly aggregation transform for insurance domain.
'          Groups monthly Detail rows by Entity + Quarter, writes SUMIFS
'          formulas on QuarterlySummary tab. Flow fields summed per quarter.
'          Balance fields: Detail stores true incremental (change-in-EOP),
'          so SUMIFS(<=lastMon) reconstructs EOP at quarter-end.
'          Derived fields recomputed from quarterly operands.
' Split from InsuranceDomainEngine.bas for modularity (AD-09).
' Phase 11A. ASCII only (AP-06). All column refs via ColIndex (AP-08).


' AggregateToQuarterly (PostCompute Transform)
' Groups monthly Detail rows by Entity + Quarter, sums Flow fields,
' reconstructs EOP for Balance fields (cumulative sum of incremental),
' recomputes Derived fields from quarterly sums.
' Writes QuarterlySummary tab with FM quarterly column layout.
Public Sub AggregateToQuarterly()
    On Error GoTo ErrHandler

    Dim outputs As Variant
    outputs = KernelTransform.TransformOutputs

    Dim totalRows As Long
    totalRows = UBound(outputs, 1)

    ' Column indices (AP-08)
    Dim colEntityName As Long: colEntityName = KernelConfig.ColIndex("EntityName")
    Dim colPeriod As Long: colPeriod = KernelConfig.ColIndex("CalPeriod")
    Dim colQuarter As Long: colQuarter = KernelConfig.ColIndex("CalQuarter")
    Dim colYear As Long: colYear = KernelConfig.ColIndex("CalYear")

    If colEntityName < 1 Or colPeriod < 1 Or colQuarter < 1 Then
        KernelConfig.LogError SEV_ERROR, "Ins_QuarterlyAgg", "E-360", _
            "Missing required dimension columns for quarterly aggregation", _
            "MANUAL BYPASS: Verify EntityName, Period, Quarter columns exist in column_registry."
        Exit Sub
    End If

    ' Detect entity count and names
    Dim entityNames() As String
    Dim entityCount As Long
    entityCount = 0
    ReDim entityNames(1 To totalRows)

    Dim r As Long
    For r = 1 To totalRows
        Dim eName As String
        eName = CStr(outputs(r, colEntityName))
        If Len(eName) > 0 Then
            Dim found As Boolean
            found = False
            Dim e As Long
            For e = 1 To entityCount
                If StrComp(entityNames(e), eName, vbTextCompare) = 0 Then
                    found = True
                    Exit For
                End If
            Next e
            If Not found Then
                entityCount = entityCount + 1
                entityNames(entityCount) = eName
            End If
        End If
    Next r

    If entityCount = 0 Then
        KernelConfig.LogError SEV_WARN, "Ins_QuarterlyAgg", "W-360", _
            "No entities found for quarterly aggregation.", ""
        Exit Sub
    End If

    ' Determine number of years from actual data (supports run-off beyond TimeHorizon)
    Dim maxYearVal As Long
    maxYearVal = 0
    If colYear > 0 Then
        For r = 1 To totalRows
            Dim yrVal As Long
            If IsNumeric(outputs(r, colYear)) Then
                yrVal = CLng(outputs(r, colYear))
                If yrVal > maxYearVal Then maxYearVal = yrVal
            End If
        Next r
    End If
    If maxYearVal < 1 Then
        ' Fallback to TimeHorizon if Year column not populated
        Dim timeHorizon As Long
        timeHorizon = KernelConfig.GetTimeHorizon()
        If timeHorizon <= 0 Then timeHorizon = 60
        maxYearVal = (timeHorizon - 1) \ 12 + 1
    End If
    Dim numQuarters As Long
    numQuarters = maxYearVal * QS_QUARTERS_PER_YEAR
    Dim numYears As Long
    numYears = maxYearVal

    ' Tail column: aggregates all development beyond the writing horizon
    Dim writingMonths As Long
    writingMonths = KernelConfig.GetTimeHorizon()
    If writingMonths <= 0 Then writingMonths = 60
    Dim writingYears As Long
    writingYears = (writingMonths - 1) \ 12 + 1
    Dim hasTail As Boolean
    hasTail = (numYears > writingYears)
    Dim tailCol As Long
    tailCol = QS_DATA_START_COL + numYears * QS_COLS_PER_YEAR
    Dim tailCriteria As String
    tailCriteria = """>" & writingMonths & """"

    ' Get all column names and their BalanceType
    Dim colCount As Long
    colCount = KernelConfig.GetColumnCount()

    ' Build list of Incremental + Derived columns with their BalanceType
    Dim incrCols As Variant
    incrCols = KernelConfig.GetIncrementalColumns()
    Dim derivCols As Variant
    derivCols = KernelConfig.GetDerivedColumns()

    Dim incrCount As Long
    incrCount = 0
    If IsArray(incrCols) Then
        On Error Resume Next
        incrCount = UBound(incrCols) - LBound(incrCols) + 1
        If Err.Number <> 0 Then incrCount = 0
        On Error GoTo ErrHandler
    End If

    Dim derivCount As Long
    derivCount = 0
    If IsArray(derivCols) Then
        On Error Resume Next
        derivCount = UBound(derivCols) - LBound(derivCols) + 1
        If Err.Number <> 0 Then derivCount = 0
        On Error GoTo ErrHandler
    End If

    ' Track which metrics are Balance type
    Dim isBalance() As Boolean
    If incrCount > 0 Then
        ReDim isBalance(1 To incrCount)
        Dim ic As Long
        For ic = LBound(incrCols) To UBound(incrCols)
            Dim micIdx As Long
            micIdx = ic - LBound(incrCols) + 1
            Dim balType As String
            balType = KernelConfig.GetBalanceType(CStr(incrCols(ic)))
            isBalance(micIdx) = (StrComp(balType, BALANCE_TYPE_BALANCE, vbTextCompare) = 0)
        Next ic
    End If

    ' Detail tab column letters for SUMIFS formula references
    Dim dEntL As String: dEntL = ColNumToLetter(KernelConfig.ColIndex("EntityName"))
    Dim dPerL As String: dPerL = ColNumToLetter(KernelConfig.ColIndex("CalPeriod"))
    Dim dQtrL As String: dQtrL = ColNumToLetter(KernelConfig.ColIndex("CalQuarter"))
    Dim dYrL As String: dYrL = ColNumToLetter(KernelConfig.ColIndex("CalYear"))

    ' Write QuarterlySummary tab
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(TAB_QUARTERLY_SUMMARY)
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = TAB_QUARTERLY_SUMMARY
    End If

    On Error Resume Next
    ws.Unprotect
    On Error GoTo ErrHandler
    ws.Cells.ClearContents

    ' Header row
    ws.Cells(1, 1).Value = "RowID"
    ws.Cells(1, 2).Value = "Metric"

    ' Write quarterly column headers
    Dim yr As Long
    For yr = 1 To numYears
        Dim q As Long
        For q = 1 To QS_QUARTERS_PER_YEAR
            Dim hCol As Long
            hCol = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + (q - 1)
            ws.Cells(1, hCol).Value = "Q" & q & " Y" & yr
            ws.Cells(1, hCol).Font.Bold = True
            ws.Cells(1, hCol).HorizontalAlignment = xlCenter
        Next q
        Dim annHCol As Long
        annHCol = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + QS_QUARTERS_PER_YEAR
        ws.Cells(1, annHCol).Value = "Y" & yr & " Total"
        ws.Cells(1, annHCol).Font.Bold = True
        ws.Cells(1, annHCol).HorizontalAlignment = xlCenter
        ws.Cells(1, annHCol).Interior.Color = RGB(217, 217, 217)
    Next yr

    ' Tail column header
    If hasTail Then
        ws.Cells(1, tailCol).Value = "Tail"
        ws.Cells(1, tailCol).Font.Bold = True
        ws.Cells(1, tailCol).HorizontalAlignment = xlCenter
        ws.Cells(1, tailCol).Interior.Color = RGB(198, 224, 180)
    End If

    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 2).Font.Bold = True

    Dim curRow As Long
    curRow = 2

    ' Write each Incremental metric section
    If incrCount > 0 Then
        Dim mIdx As Long
        For mIdx = LBound(incrCols) To UBound(incrCols)
            Dim metricName As String
            metricName = CStr(incrCols(mIdx))
            Dim mPos As Long
            mPos = mIdx - LBound(incrCols) + 1

            ' Use full metric name for RowID (not truncated to 4 chars)
            Dim metricUpper As String
            metricUpper = UCase(metricName)

            ' Section label
            curRow = curRow + 1
            ws.Cells(curRow, 1).Value = "QS_SEC_" & metricUpper
            ws.Cells(curRow, 2).Value = KernelConfig.GetDisplayAlias(metricName)
            ws.Cells(curRow, 2).Font.Bold = True
            ws.Cells(curRow, 2).Interior.Color = RGB(217, 225, 242)

            ' Detail column letter for this metric
            Dim mL As String
            mL = ColNumToLetter(KernelConfig.ColIndex(metricName))

            ' Per-entity rows (SUMIFS formulas referencing Detail tab)
            For e = 1 To entityCount
                curRow = curRow + 1
                ws.Cells(curRow, 1).Value = "QS_" & metricUpper & "_" & e
                ws.Cells(curRow, 2).Value = entityNames(e)
                ws.Cells(curRow, 2).IndentLevel = 1

                For yr = 1 To numYears
                    For q = 1 To QS_QUARTERS_PER_YEAR
                        Dim globalQ As Long
                        globalQ = (yr - 1) * QS_QUARTERS_PER_YEAR + q
                        If globalQ <= numQuarters Then
                            Dim dataCol As Long
                            dataCol = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + (q - 1)
                            If isBalance(mPos) Then
                                ' Balance: sum incremental values up through last month of quarter
                                ' to reconstruct EOP balance at quarter-end
                                Dim lastMon As Long
                                lastMon = (yr - 1) * 12 + q * 3
                                ws.Cells(curRow, dataCol).formula = _
                                    "=SUMIFS(Detail!" & mL & ":" & mL & _
                                    ",Detail!$" & dEntL & ":$" & dEntL & ",$B" & curRow & _
                                    ",Detail!$" & dPerL & ":$" & dPerL & ",""<=""&" & lastMon & ")"
                            Else
                                ' Flow: sum of 3 months in quarter
                                ws.Cells(curRow, dataCol).formula = _
                                    "=SUMIFS(Detail!" & mL & ":" & mL & _
                                    ",Detail!$" & dEntL & ":$" & dEntL & ",$B" & curRow & _
                                    ",Detail!$" & dYrL & ":$" & dYrL & "," & yr & _
                                    ",Detail!$" & dQtrL & ":$" & dQtrL & "," & q & ")"
                            End If
                        End If
                    Next q
                    ' Annual total
                    Dim annDataCol As Long
                    annDataCol = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + QS_QUARTERS_PER_YEAR
                    If isBalance(mPos) Then
                        Dim q4col As Long
                        q4col = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + 3
                        ws.Cells(curRow, annDataCol).formula = "=" & ColNumToLetter(q4col) & curRow
                    Else
                        Dim q1c As Long
                        q1c = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR
                        Dim q4c As Long
                        q4c = q1c + QS_QUARTERS_PER_YEAR - 1
                        ws.Cells(curRow, annDataCol).formula = "=SUM(" & _
                            ColNumToLetter(q1c) & curRow & ":" & ColNumToLetter(q4c) & curRow & ")"
                    End If
                    ws.Cells(curRow, annDataCol).Interior.Color = RGB(217, 217, 217)
                Next yr

                ' Tail column: SUMIFS for periods beyond writing horizon
                If hasTail Then
                    ws.Cells(curRow, tailCol).formula = _
                        "=SUMIFS(Detail!" & mL & ":" & mL & _
                        ",Detail!$" & dEntL & ":$" & dEntL & ",$B" & curRow & _
                        ",Detail!$" & dPerL & ":$" & dPerL & "," & tailCriteria & ")"
                    ws.Cells(curRow, tailCol).Interior.Color = RGB(198, 224, 180)
                End If

                ' Apply number format
                Dim mFmt As String
                mFmt = KernelConfig.GetFormat(metricName)
                If Len(mFmt) > 0 Then
                    Dim lastDCol As Long
                    lastDCol = QS_DATA_START_COL + numYears * QS_COLS_PER_YEAR - 1
                    ws.Range(ws.Cells(curRow, QS_DATA_START_COL), ws.Cells(curRow, lastDCol)).NumberFormat = mFmt
                    If hasTail Then ws.Cells(curRow, tailCol).NumberFormat = mFmt
                End If
            Next e

            ' Total row
            curRow = curRow + 1
            ws.Cells(curRow, 1).Value = "QS_" & metricUpper & "_TOTAL"
            ws.Cells(curRow, 2).Value = "Total " & KernelConfig.GetDisplayAlias(metricName)
            ws.Cells(curRow, 2).Font.Bold = True

            Dim firstEntRow As Long
            firstEntRow = curRow - entityCount
            Dim lastEntRow As Long
            lastEntRow = curRow - 1

            For yr = 1 To numYears
                For q = 1 To QS_QUARTERS_PER_YEAR
                    dataCol = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + (q - 1)
                    ws.Cells(curRow, dataCol).formula = "=SUM(" & _
                        ColNumToLetter(dataCol) & firstEntRow & ":" & ColNumToLetter(dataCol) & lastEntRow & ")"
                Next q
                annDataCol = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + QS_QUARTERS_PER_YEAR
                If isBalance(mPos) Then
                    q4c = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + 3
                    ws.Cells(curRow, annDataCol).formula = "=" & ColNumToLetter(q4c) & curRow
                Else
                    q1c = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR
                    q4c = q1c + QS_QUARTERS_PER_YEAR - 1
                    ws.Cells(curRow, annDataCol).formula = "=SUM(" & _
                        ColNumToLetter(q1c) & curRow & ":" & ColNumToLetter(q4c) & curRow & ")"
                End If
                ws.Cells(curRow, annDataCol).Interior.Color = RGB(217, 217, 217)
            Next yr

            ' Tail column for total row: SUM of entity tail cells
            If hasTail Then
                ws.Cells(curRow, tailCol).formula = "=SUM(" & _
                    ColNumToLetter(tailCol) & firstEntRow & ":" & ColNumToLetter(tailCol) & lastEntRow & ")"
                ws.Cells(curRow, tailCol).Interior.Color = RGB(198, 224, 180)
            End If

            ' Apply number format to total row
            If Len(mFmt) > 0 Then
                lastDCol = QS_DATA_START_COL + numYears * QS_COLS_PER_YEAR - 1
                ws.Range(ws.Cells(curRow, QS_DATA_START_COL), ws.Cells(curRow, lastDCol)).NumberFormat = mFmt
                If hasTail Then ws.Cells(curRow, tailCol).NumberFormat = mFmt
            End If

            ' Bottom border
            Dim borderEndCol As Long
            borderEndCol = QS_DATA_START_COL + numYears * QS_COLS_PER_YEAR - 1
            If hasTail Then borderEndCol = tailCol
            ws.Range(ws.Cells(curRow, 2), ws.Cells(curRow, borderEndCol)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            ws.Range(ws.Cells(curRow, 2), ws.Cells(curRow, borderEndCol)).Borders(xlEdgeBottom).Weight = xlThin
        Next mIdx
    End If

    ' Write Derived metric sections (recomputed from quarterly operands)
    If derivCount > 0 Then
        Dim dIdx As Long
        For dIdx = LBound(derivCols) To UBound(derivCols)
            Dim dMetricName As String
            dMetricName = CStr(derivCols(dIdx))
            Dim dRule As String
            dRule = KernelConfig.GetDerivationRule(dMetricName)
            If Len(dRule) = 0 Then GoTo NextDerivedMetric

            Dim opA As String
            Dim opB As String
            Dim opStr As String
            Dim isCopyRule As Boolean
            isCopyRule = False

            If Not ParseSimpleRule(dRule, opA, opStr, opB) Then
                ' Single-operand copy rule: rule is just a column name
                ' (e.g., C_ClsCt = "G_ClsCt" means copy from G_ClsCt)
                If KernelConfig.ColIndex(Trim(dRule)) > 0 Then
                    opA = Trim(dRule)
                    isCopyRule = True
                Else
                    GoTo NextDerivedMetric
                End If
            End If

            Dim dMetricUpper As String
            dMetricUpper = UCase(dMetricName)

            ' Section label
            curRow = curRow + 2
            ws.Cells(curRow, 1).Value = "QS_SEC_" & dMetricUpper
            ws.Cells(curRow, 2).Value = KernelConfig.GetDisplayAlias(dMetricName)
            ws.Cells(curRow, 2).Font.Bold = True
            ws.Cells(curRow, 2).Interior.Color = RGB(217, 225, 242)

            ' Find rows of operand totals
            Dim opARowID As String
            opARowID = "QS_" & UCase(opA) & "_TOTAL"
            Dim opBRowID As String
            If Not isCopyRule Then
                opBRowID = "QS_" & UCase(opB) & "_TOTAL"
            End If

            ' Per-entity rows
            For e = 1 To entityCount
                curRow = curRow + 1
                ws.Cells(curRow, 1).Value = "QS_" & dMetricUpper & "_" & e
                ws.Cells(curRow, 2).Value = entityNames(e)
                ws.Cells(curRow, 2).IndentLevel = 1

                Dim opAEntRowID As String
                opAEntRowID = "QS_" & UCase(opA) & "_" & e

                Dim opAEntRow As Long
                opAEntRow = FindRowIDInSheet(ws, opAEntRowID)

                Dim opBEntRow As Long
                Dim opBEntRowID As String
                If isCopyRule Then
                    opBEntRow = -1
                Else
                    opBEntRowID = "QS_" & UCase(opB) & "_" & e
                    opBEntRow = FindRowIDInSheet(ws, opBEntRowID)
                End If

                Dim canWriteFormula As Boolean
                If isCopyRule Then
                    canWriteFormula = (opAEntRow > 0)
                Else
                    canWriteFormula = (opAEntRow > 0 And opBEntRow > 0)
                End If

                If canWriteFormula Then
                    For yr = 1 To numYears
                        For q = 1 To QS_QUARTERS_PER_YEAR
                            dataCol = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + (q - 1)
                            Dim dFormula As String
                            Dim cLtr As String
                            cLtr = ColNumToLetter(dataCol)
                            If isCopyRule Then
                                dFormula = "=" & cLtr & opAEntRow
                            Else
                                Select Case opStr
                                    Case "-"
                                        dFormula = "=" & cLtr & opAEntRow & "-" & cLtr & opBEntRow
                                    Case "+"
                                        dFormula = "=" & cLtr & opAEntRow & "+" & cLtr & opBEntRow
                                    Case "*"
                                        dFormula = "=" & cLtr & opAEntRow & "*" & cLtr & opBEntRow
                                    Case "/"
                                        dFormula = "=IFERROR(" & cLtr & opAEntRow & "/" & cLtr & opBEntRow & ",0)"
                                End Select
                            End If
                            ws.Cells(curRow, dataCol).formula = dFormula
                        Next q
                        ' Annual total
                        annDataCol = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + QS_QUARTERS_PER_YEAR
                        If isCopyRule Then
                            ws.Cells(curRow, annDataCol).formula = "=" & _
                                ColNumToLetter(annDataCol) & opAEntRow
                        Else
                            Select Case opStr
                                Case "/"
                                    Dim annALetter As String
                                    annALetter = ColNumToLetter(annDataCol)
                                    ws.Cells(curRow, annDataCol).formula = "=IFERROR(" & _
                                        annALetter & opAEntRow & "/" & annALetter & opBEntRow & ",0)"
                                Case Else
                                    q1c = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR
                                    q4c = q1c + QS_QUARTERS_PER_YEAR - 1
                                    ws.Cells(curRow, annDataCol).formula = "=SUM(" & _
                                        ColNumToLetter(q1c) & curRow & ":" & ColNumToLetter(q4c) & curRow & ")"
                            End Select
                        End If
                        ws.Cells(curRow, annDataCol).Interior.Color = RGB(217, 217, 217)
                    Next yr

                    ' Tail column for derived entity row
                    If hasTail Then
                        Dim tLtr As String
                        tLtr = ColNumToLetter(tailCol)
                        If isCopyRule Then
                            ws.Cells(curRow, tailCol).formula = "=" & tLtr & opAEntRow
                        Else
                            Select Case opStr
                                Case "-"
                                    ws.Cells(curRow, tailCol).formula = "=" & tLtr & opAEntRow & "-" & tLtr & opBEntRow
                                Case "+"
                                    ws.Cells(curRow, tailCol).formula = "=" & tLtr & opAEntRow & "+" & tLtr & opBEntRow
                                Case "*"
                                    ws.Cells(curRow, tailCol).formula = "=" & tLtr & opAEntRow & "*" & tLtr & opBEntRow
                                Case "/"
                                    ws.Cells(curRow, tailCol).formula = "=IFERROR(" & tLtr & opAEntRow & "/" & tLtr & opBEntRow & ",0)"
                            End Select
                        End If
                        ws.Cells(curRow, tailCol).Interior.Color = RGB(198, 224, 180)
                    End If
                End If

                ' Apply number format
                Dim dFmt As String
                dFmt = KernelConfig.GetFormat(dMetricName)
                If Len(dFmt) > 0 Then
                    lastDCol = QS_DATA_START_COL + numYears * QS_COLS_PER_YEAR - 1
                    ws.Range(ws.Cells(curRow, QS_DATA_START_COL), ws.Cells(curRow, lastDCol)).NumberFormat = dFmt
                    If hasTail Then ws.Cells(curRow, tailCol).NumberFormat = dFmt
                End If
            Next e

            ' Total row for derived metric
            curRow = curRow + 1
            ws.Cells(curRow, 1).Value = "QS_" & dMetricUpper & "_TOTAL"
            ws.Cells(curRow, 2).Value = "Total " & KernelConfig.GetDisplayAlias(dMetricName)
            ws.Cells(curRow, 2).Font.Bold = True

            Dim opATotalRow As Long
            opATotalRow = FindRowIDInSheet(ws, opARowID)

            Dim opBTotalRow As Long
            Dim canWriteTotal As Boolean
            If isCopyRule Then
                canWriteTotal = (opATotalRow > 0)
            Else
                opBTotalRow = FindRowIDInSheet(ws, opBRowID)
                canWriteTotal = (opATotalRow > 0 And opBTotalRow > 0)
            End If

            If canWriteTotal Then
                For yr = 1 To numYears
                    For q = 1 To QS_QUARTERS_PER_YEAR
                        dataCol = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + (q - 1)
                        cLtr = ColNumToLetter(dataCol)
                        If isCopyRule Then
                            dFormula = "=" & cLtr & opATotalRow
                        Else
                            Select Case opStr
                                Case "-"
                                    dFormula = "=" & cLtr & opATotalRow & "-" & cLtr & opBTotalRow
                                Case "+"
                                    dFormula = "=" & cLtr & opATotalRow & "+" & cLtr & opBTotalRow
                                Case "*"
                                    dFormula = "=" & cLtr & opATotalRow & "*" & cLtr & opBTotalRow
                                Case "/"
                                    dFormula = "=IFERROR(" & cLtr & opATotalRow & "/" & cLtr & opBTotalRow & ",0)"
                            End Select
                        End If
                        ws.Cells(curRow, dataCol).formula = dFormula
                    Next q
                    annDataCol = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + QS_QUARTERS_PER_YEAR
                    If isCopyRule Then
                        ws.Cells(curRow, annDataCol).formula = "=" & _
                            ColNumToLetter(annDataCol) & opATotalRow
                    Else
                        Select Case opStr
                            Case "/"
                                annALetter = ColNumToLetter(annDataCol)
                                ws.Cells(curRow, annDataCol).formula = "=IFERROR(" & _
                                    annALetter & opATotalRow & "/" & annALetter & opBTotalRow & ",0)"
                            Case Else
                                q1c = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR
                                q4c = q1c + QS_QUARTERS_PER_YEAR - 1
                                ws.Cells(curRow, annDataCol).formula = "=SUM(" & _
                                    ColNumToLetter(q1c) & curRow & ":" & ColNumToLetter(q4c) & curRow & ")"
                        End Select
                    End If
                    ws.Cells(curRow, annDataCol).Interior.Color = RGB(217, 217, 217)
                Next yr

                ' Tail column for derived total row
                If hasTail Then
                    tLtr = ColNumToLetter(tailCol)
                    If isCopyRule Then
                        ws.Cells(curRow, tailCol).formula = "=" & tLtr & opATotalRow
                    Else
                        Select Case opStr
                            Case "-"
                                ws.Cells(curRow, tailCol).formula = "=" & tLtr & opATotalRow & "-" & tLtr & opBTotalRow
                            Case "+"
                                ws.Cells(curRow, tailCol).formula = "=" & tLtr & opATotalRow & "+" & tLtr & opBTotalRow
                            Case "*"
                                ws.Cells(curRow, tailCol).formula = "=" & tLtr & opATotalRow & "*" & tLtr & opBTotalRow
                            Case "/"
                                ws.Cells(curRow, tailCol).formula = "=IFERROR(" & tLtr & opATotalRow & "/" & tLtr & opBTotalRow & ",0)"
                        End Select
                    End If
                    ws.Cells(curRow, tailCol).Interior.Color = RGB(198, 224, 180)
                End If
            End If

            ' Apply number format to total row
            If Len(dFmt) > 0 Then
                lastDCol = QS_DATA_START_COL + numYears * QS_COLS_PER_YEAR - 1
                ws.Range(ws.Cells(curRow, QS_DATA_START_COL), ws.Cells(curRow, lastDCol)).NumberFormat = dFmt
                If hasTail Then ws.Cells(curRow, tailCol).NumberFormat = dFmt
            End If

            ' Bottom border
            borderEndCol = QS_DATA_START_COL + numYears * QS_COLS_PER_YEAR - 1
            If hasTail Then borderEndCol = tailCol
            ws.Range(ws.Cells(curRow, 2), ws.Cells(curRow, borderEndCol)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            ws.Range(ws.Cells(curRow, 2), ws.Cells(curRow, borderEndCol)).Borders(xlEdgeBottom).Weight = xlThin
NextDerivedMetric:
        Next dIdx
    End If

    ' AutoFit then hide RowID column (AP-62: Hidden AFTER AutoFit)
    On Error Resume Next
    ws.Columns.AutoFit
    On Error GoTo ErrHandler
    ws.Columns(1).Hidden = True

    KernelConfig.LogError SEV_INFO, "Ins_QuarterlyAgg", "I-360", _
        "AggregateToQuarterly completed: " & entityCount & " entities, " & _
        numQuarters & " quarters on QuarterlySummary tab.", ""

    Exit Sub

ErrHandler:
    KernelConfig.LogError SEV_ERROR, "Ins_QuarterlyAgg", "E-361", _
        "AggregateToQuarterly error: " & Err.Description, _
        "MANUAL BYPASS: Create QuarterlySummary tab manually with quarterly " & _
        "sums from the Detail tab. Use RowIDs in Column A."
End Sub


' ColNumToLetter -- column number to letter helper
Private Function ColNumToLetter(colNum As Long) As String
    Dim n As Long
    n = colNum
    ColNumToLetter = ""
    Do While n > 0
        Dim remainder As Long
        remainder = (n - 1) Mod 26
        ColNumToLetter = Chr(65 + remainder) & ColNumToLetter
        n = (n - 1) \ 26
    Loop
End Function


' FindRowIDInSheet -- finds a RowID in column A, returns row or 0
' Uses array read for performance instead of cell-by-cell scan
Private Function FindRowIDInSheet(ws As Worksheet, rowID As String) As Long
    FindRowIDInSheet = 0
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastRow > 2000 Then lastRow = 2000
    If lastRow < 1 Then Exit Function

    ' Read column A into array for fast scan (PT-001)
    Dim colData As Variant
    colData = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 1)).Value

    Dim r As Long
    For r = 1 To lastRow
        If StrComp(Trim(CStr(colData(r, 1))), rowID, vbTextCompare) = 0 Then
            FindRowIDInSheet = r
            Exit Function
        End If
    Next r
End Function


' ParseSimpleRule -- parses "A - B" into operandA, operator, operandB
Private Function ParseSimpleRule(rule As String, ByRef opA As String, _
    ByRef opStr As String, ByRef opB As String) As Boolean
    ParseSimpleRule = False
    Dim ops As Variant
    ops = Array(" - ", " + ", " * ", " / ")
    Dim oi As Long
    For oi = LBound(ops) To UBound(ops)
        Dim p As Long
        p = InStr(1, rule, ops(oi), vbTextCompare)
        If p > 0 Then
            opA = Trim(Left(rule, p - 1))
            opStr = Trim(CStr(ops(oi)))
            opB = Trim(Mid(rule, p + Len(CStr(ops(oi)))))
            If Len(opA) > 0 And Len(opB) > 0 Then
                ParseSimpleRule = True
                Exit Function
            End If
        End If
    Next oi
End Function
