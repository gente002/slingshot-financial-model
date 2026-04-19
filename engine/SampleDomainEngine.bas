Attribute VB_Name = "SampleDomainEngine"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' SampleDomainEngine.bas
' Purpose: Implements the toy model computation. This is the ONLY domain code.
'          Implements the 4-function domain contract (AP-43).
' =============================================================================

' Module-level state
Private m_initialized As Boolean


' =============================================================================
' Initialize
' Called once at bootstrap. For Phase 1, domain code calls KernelConfig
' public functions directly since all modules are in the same VBA project.
' =============================================================================
Public Sub Initialize()
    m_initialized = True
    ' Phase 5C: Register quarterly aggregation transform
    KernelTransform.RegisterTransform "QuarterlyAgg", _
        "SampleDomainEngine", "AggregateToQuarterly", 100
End Sub


' =============================================================================
' Validate
' Check that required inputs are present and valid.
' Returns True if all valid, False if any fail.
' =============================================================================
Public Function Validate() As Boolean
    Validate = True

    ' Determine entity count
    Dim entityCount As Long
    entityCount = DetectEntityCountForValidation()

    If entityCount = 0 Then
        KernelConfig.LogError SEV_ERROR, "SampleDomainEngine", "E-300", _
                              "No entities found for validation", _
                              "MANUAL BYPASS: Add entity names to the Inputs tab row 3, columns C onward."
        Validate = False
        Exit Function
    End If

    Dim entIdx As Long
    For entIdx = 1 To entityCount
        ' Validate Units > 0
        Dim units As Variant
        units = KernelConfig.InputValue("Assumptions", "Units", entIdx)
        If Not IsNumeric(units) Then
            KernelConfig.LogError SEV_ERROR, "SampleDomainEngine", "E-301", _
                                  "Units is not numeric for entity " & entIdx, _
                                  "MANUAL BYPASS: Enter a numeric value > 0 in the Units row for entity " & entIdx & "."
            Validate = False
        ElseIf CDbl(units) <= 0 Then
            KernelConfig.LogError SEV_ERROR, "SampleDomainEngine", "E-302", _
                                  "Units must be > 0 for entity " & entIdx, _
                                  "MANUAL BYPASS: Enter a value > 0 in the Units row for entity " & entIdx & "."
            Validate = False
        End If

        ' Validate UnitPrice > 0
        Dim unitPrice As Variant
        unitPrice = KernelConfig.InputValue("Assumptions", "UnitPrice", entIdx)
        If Not IsNumeric(unitPrice) Then
            KernelConfig.LogError SEV_ERROR, "SampleDomainEngine", "E-303", _
                                  "UnitPrice is not numeric for entity " & entIdx, _
                                  "MANUAL BYPASS: Enter a numeric value > 0 in the UnitPrice row for entity " & entIdx & "."
            Validate = False
        ElseIf CDbl(unitPrice) <= 0 Then
            KernelConfig.LogError SEV_ERROR, "SampleDomainEngine", "E-304", _
                                  "UnitPrice must be > 0 for entity " & entIdx, _
                                  "MANUAL BYPASS: Enter a value > 0 in the UnitPrice row for entity " & entIdx & "."
            Validate = False
        End If

        ' Validate COGSPct between 0 and 1
        Dim cogsPct As Variant
        cogsPct = KernelConfig.InputValue("Assumptions", "COGSPct", entIdx)
        If Not IsNumeric(cogsPct) Then
            KernelConfig.LogError SEV_ERROR, "SampleDomainEngine", "E-305", _
                                  "COGSPct is not numeric for entity " & entIdx, _
                                  "MANUAL BYPASS: Enter a decimal between 0 and 1 in the COGSPct row for entity " & entIdx & "."
            Validate = False
        ElseIf CDbl(cogsPct) < 0 Or CDbl(cogsPct) > 1 Then
            KernelConfig.LogError SEV_ERROR, "SampleDomainEngine", "E-306", _
                                  "COGSPct must be between 0 and 1 for entity " & entIdx, _
                                  "MANUAL BYPASS: Enter a decimal between 0 and 1 in the COGSPct row for entity " & entIdx & "."
            Validate = False
        End If
    Next entIdx
End Function


' =============================================================================
' Reset
' Nothing to reset for this simple model.
' =============================================================================
Public Sub Reset()
    ' More complex domain modules would Erase module-level arrays here.
    ' This simple model has no computation state to reset.
End Sub


' =============================================================================
' Execute
' Computes Revenue and COGS for each entity and period.
' Writes Dimension values and Incremental values ONLY.
' Does NOT compute GrossProfit or GPMargin (kernel derives these via AP-42).
' Domain reads inputs via KernelConfig.InputValue() -- no inputs array needed.
' =============================================================================
Public Sub Execute()
    Dim outputs As Variant
    outputs = KernelEngine.DomainOutputs

    Dim entityCount As Long
    entityCount = DetectEntityCountForValidation()

    Dim periodCount As Long
    periodCount = KernelConfig.GetTimeHorizon()

    ' Column indices via ColIndex (AP-08: no magic numbers)
    Dim colEntityName As Long
    colEntityName = KernelConfig.ColIndex("EntityName")

    Dim colPeriod As Long
    colPeriod = KernelConfig.ColIndex("Period")

    Dim colQuarter As Long
    colQuarter = KernelConfig.ColIndex("Quarter")

    Dim colYear As Long
    colYear = KernelConfig.ColIndex("Year")

    Dim colRevenue As Long
    colRevenue = KernelConfig.ColIndex("Revenue")

    Dim colCOGS As Long
    colCOGS = KernelConfig.ColIndex("COGS")

    ' Process each entity
    Dim entIdx As Long
    For entIdx = 1 To entityCount
        ' Read inputs via KernelConfig.InputValue
        Dim units As Double
        units = CDbl(KernelConfig.InputValue("Assumptions", "Units", entIdx))

        Dim price As Double
        price = CDbl(KernelConfig.InputValue("Assumptions", "UnitPrice", entIdx))

        Dim growth As Double
        growth = CDbl(KernelConfig.InputValue("Assumptions", "MonthlyGrowth", entIdx))

        Dim cogsPct As Double
        cogsPct = CDbl(KernelConfig.InputValue("Assumptions", "COGSPct", entIdx))

        Dim entityName As String
        entityName = CStr(KernelConfig.InputValue("Attributes", "EntityName", entIdx))

        ' Process each period
        Dim prd As Long
        For prd = 1 To periodCount
            Dim row As Long
            row = (entIdx - 1) * periodCount + prd

            ' Compute Incremental values only (AP-42: never compute cumulative)
            Dim revenue As Double
            revenue = units * price * (1 + growth) ^ (prd - 1)

            Dim cogs As Double
            cogs = revenue * cogsPct

            ' Write to output array using ColIndex (AP-08: no magic numbers)
            outputs(row, colRevenue) = revenue
            outputs(row, colCOGS) = cogs

            ' Write dimension values
            outputs(row, colEntityName) = entityName
            outputs(row, colPeriod) = prd
            outputs(row, colQuarter) = Int((prd - 1) / 3) + 1
            outputs(row, colYear) = 1

            ' DO NOT compute GrossProfit or GPMargin - kernel derives these
        Next prd
    Next entIdx

    KernelEngine.DomainOutputs = outputs
End Sub


' =============================================================================
' DetectEntityCountForValidation
' Counts non-empty entity names on the Inputs sheet.
' =============================================================================
' =============================================================================
' ApplyLoading (Sample Transform)
' Demonstrates the transform contract. Transforms access the outputs array
' via KernelTransform.TransformOutputs (Application.Run cannot pass arrays).
' Register with: KernelTransform.RegisterTransform "TestLoad", _
'                "SampleDomainEngine", "ApplyLoading", 1
' =============================================================================
Public Sub ApplyLoading()
    KernelConfig.LogError SEV_INFO, "SampleDomainEngine", "I-350", _
        "ApplyLoading transform executed (sample no-op)", _
        "Rows: " & UBound(KernelTransform.TransformOutputs, 1)
End Sub


' =============================================================================
' AggregateToQuarterly (Phase 5C Transform)
' Groups monthly Detail rows by Entity + Quarter, sums Incremental fields,
' takes EOP for Balance fields, recomputes Derived fields from quarterly sums.
' Writes QuarterlySummary tab with FM quarterly column layout.
' Registered at sortOrder 100 in Initialize().
' =============================================================================
Public Sub AggregateToQuarterly()
    On Error GoTo ErrHandler

    Dim outputs As Variant
    outputs = KernelTransform.TransformOutputs

    Dim totalRows As Long
    totalRows = UBound(outputs, 1)

    ' Column indices (AP-08)
    Dim colEntityName As Long
    colEntityName = KernelConfig.ColIndex("EntityName")
    Dim colPeriod As Long
    colPeriod = KernelConfig.ColIndex("Period")
    Dim colQuarter As Long
    colQuarter = KernelConfig.ColIndex("Quarter")
    Dim colYear As Long
    colYear = KernelConfig.ColIndex("Year")

    If colEntityName < 1 Or colPeriod < 1 Or colQuarter < 1 Then
        KernelConfig.LogError SEV_ERROR, "SampleDomainEngine", "E-360", _
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
        KernelConfig.LogError SEV_WARN, "SampleDomainEngine", "W-360", _
            "No entities found for quarterly aggregation.", ""
        Exit Sub
    End If

    ' Determine number of quarters and years
    Dim timeHorizon As Long
    timeHorizon = KernelConfig.GetTimeHorizon()
    If timeHorizon <= 0 Then timeHorizon = 12
    Dim numQuarters As Long
    numQuarters = timeHorizon \ 3
    If numQuarters < 1 Then numQuarters = 1
    Dim numYears As Long
    numYears = numQuarters \ QS_QUARTERS_PER_YEAR
    If numYears < 1 Then numYears = 1

    ' Get Incremental and Derived column info
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

    ' Build quarterly aggregation arrays
    ' qData(entity, quarter, metric) = summed value
    Dim qData() As Double
    ReDim qData(1 To entityCount, 1 To numQuarters, 1 To incrCount)

    ' Aggregate Incremental fields by entity + quarter
    For r = 1 To totalRows
        Dim entName As String
        entName = CStr(outputs(r, colEntityName))
        Dim qtrNum As Long
        If IsNumeric(outputs(r, colQuarter)) Then
            qtrNum = CLng(outputs(r, colQuarter))
        Else
            qtrNum = 0
        End If
        If qtrNum < 1 Or qtrNum > numQuarters Then GoTo NextAggRow

        ' Find entity index
        Dim entIdx As Long
        entIdx = 0
        For e = 1 To entityCount
            If StrComp(entityNames(e), entName, vbTextCompare) = 0 Then
                entIdx = e
                Exit For
            End If
        Next e
        If entIdx = 0 Then GoTo NextAggRow

        ' Sum Incremental columns
        If incrCount > 0 Then
            Dim ic As Long
            For ic = LBound(incrCols) To UBound(incrCols)
                Dim icIdx As Long
                icIdx = KernelConfig.ColIndex(CStr(incrCols(ic)))
                If icIdx > 0 Then
                    Dim micIdx As Long
                    micIdx = ic - LBound(incrCols) + 1
                    If IsNumeric(outputs(r, icIdx)) Then
                        qData(entIdx, qtrNum, micIdx) = qData(entIdx, qtrNum, micIdx) + CDbl(outputs(r, icIdx))
                    End If
                End If
            Next ic
        End If
NextAggRow:
    Next r

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

    ' Format header row
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

            ' Section label
            curRow = curRow + 1
            ws.Cells(curRow, 1).Value = "QS_SEC_" & UCase(metricName)
            ws.Cells(curRow, 2).Value = metricName
            ws.Cells(curRow, 2).Font.Bold = True
            ws.Cells(curRow, 2).Interior.Color = RGB(217, 225, 242)

            ' Per-entity rows
            For e = 1 To entityCount
                curRow = curRow + 1
                ws.Cells(curRow, 1).Value = "QS_" & UCase(Left(metricName, 4)) & "_" & e
                ws.Cells(curRow, 2).Value = entityNames(e)
                ws.Cells(curRow, 2).IndentLevel = 1

                For yr = 1 To numYears
                    For q = 1 To QS_QUARTERS_PER_YEAR
                        Dim globalQ As Long
                        globalQ = (yr - 1) * QS_QUARTERS_PER_YEAR + q
                        If globalQ <= numQuarters Then
                            Dim dataCol As Long
                            dataCol = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + (q - 1)
                            ws.Cells(curRow, dataCol).Value = qData(e, globalQ, mPos)
                        End If
                    Next q
                    ' Annual total formula
                    Dim annDataCol As Long
                    annDataCol = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + QS_QUARTERS_PER_YEAR
                    Dim q1c As Long
                    q1c = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR
                    Dim q4c As Long
                    q4c = q1c + QS_QUARTERS_PER_YEAR - 1
                    ws.Cells(curRow, annDataCol).formula = "=SUM(" & _
                        QSColLetter(q1c) & curRow & ":" & QSColLetter(q4c) & curRow & ")"
                    ws.Cells(curRow, annDataCol).Interior.Color = RGB(217, 217, 217)
                Next yr

                ' Apply number format
                Dim mFmt As String
                mFmt = KernelConfig.GetFormat(metricName)
                If Len(mFmt) > 0 Then
                    Dim lastDCol As Long
                    lastDCol = QS_DATA_START_COL + numYears * QS_COLS_PER_YEAR - 1
                    ws.Range(ws.Cells(curRow, QS_DATA_START_COL), ws.Cells(curRow, lastDCol)).NumberFormat = mFmt
                End If
            Next e

            ' Total row
            curRow = curRow + 1
            ws.Cells(curRow, 1).Value = "QS_" & UCase(Left(metricName, 4)) & "_TOTAL"
            ws.Cells(curRow, 2).Value = "Total " & metricName
            ws.Cells(curRow, 2).Font.Bold = True

            Dim firstEntRow As Long
            firstEntRow = curRow - entityCount
            Dim lastEntRow As Long
            lastEntRow = curRow - 1

            For yr = 1 To numYears
                For q = 1 To QS_QUARTERS_PER_YEAR
                    dataCol = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + (q - 1)
                    ws.Cells(curRow, dataCol).formula = "=SUM(" & _
                        QSColLetter(dataCol) & firstEntRow & ":" & QSColLetter(dataCol) & lastEntRow & ")"
                Next q
                annDataCol = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + QS_QUARTERS_PER_YEAR
                q1c = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR
                q4c = q1c + QS_QUARTERS_PER_YEAR - 1
                ws.Cells(curRow, annDataCol).formula = "=SUM(" & _
                    QSColLetter(q1c) & curRow & ":" & QSColLetter(q4c) & curRow & ")"
                ws.Cells(curRow, annDataCol).Interior.Color = RGB(217, 217, 217)
            Next yr

            ' Apply number format to total row
            If Len(mFmt) > 0 Then
                lastDCol = QS_DATA_START_COL + numYears * QS_COLS_PER_YEAR - 1
                ws.Range(ws.Cells(curRow, QS_DATA_START_COL), ws.Cells(curRow, lastDCol)).NumberFormat = mFmt
            End If

            ' Bottom border on total row
            ws.Range(ws.Cells(curRow, 2), ws.Cells(curRow, lastDCol)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            ws.Range(ws.Cells(curRow, 2), ws.Cells(curRow, lastDCol)).Borders(xlEdgeBottom).Weight = xlThin
        Next mIdx
    End If

    ' Write Derived metric sections (recomputed from quarterly sums)
    If derivCount > 0 Then
        Dim dIdx As Long
        For dIdx = LBound(derivCols) To UBound(derivCols)
            Dim dMetricName As String
            dMetricName = CStr(derivCols(dIdx))
            Dim dRule As String
            dRule = KernelConfig.GetDerivationRule(dMetricName)
            If Len(dRule) = 0 Then GoTo NextDerivedMetric

            ' Parse derivation rule to get operands
            Dim opA As String
            Dim opB As String
            Dim opStr As String
            If Not ParseSimpleRule(dRule, opA, opStr, opB) Then GoTo NextDerivedMetric

            ' Section label
            curRow = curRow + 2
            ws.Cells(curRow, 1).Value = "QS_SEC_" & UCase(Left(dMetricName, 4))
            ws.Cells(curRow, 2).Value = dMetricName
            ws.Cells(curRow, 2).Font.Bold = True
            ws.Cells(curRow, 2).Interior.Color = RGB(217, 225, 242)

            ' Find rows of operand totals
            Dim opARowID As String
            opARowID = "QS_" & UCase(Left(opA, 4)) & "_TOTAL"
            Dim opBRowID As String
            opBRowID = "QS_" & UCase(Left(opB, 4)) & "_TOTAL"

            ' Per-entity rows
            For e = 1 To entityCount
                curRow = curRow + 1
                ws.Cells(curRow, 1).Value = "QS_" & UCase(Left(dMetricName, 4)) & "_" & e
                ws.Cells(curRow, 2).Value = entityNames(e)
                ws.Cells(curRow, 2).IndentLevel = 1

                ' Find operand entity rows
                Dim opAEntRowID As String
                opAEntRowID = "QS_" & UCase(Left(opA, 4)) & "_" & e
                Dim opBEntRowID As String
                opBEntRowID = "QS_" & UCase(Left(opB, 4)) & "_" & e

                Dim opAEntRow As Long
                opAEntRow = FindRowIDInSheet(ws, opAEntRowID)
                Dim opBEntRow As Long
                opBEntRow = FindRowIDInSheet(ws, opBEntRowID)

                If opAEntRow > 0 And opBEntRow > 0 Then
                    For yr = 1 To numYears
                        For q = 1 To QS_QUARTERS_PER_YEAR
                            dataCol = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + (q - 1)
                            Dim dFormula As String
                            Dim cLtr As String
                            cLtr = QSColLetter(dataCol)
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
                            ws.Cells(curRow, dataCol).formula = dFormula
                        Next q
                        ' Annual total
                        annDataCol = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + QS_QUARTERS_PER_YEAR
                        Select Case opStr
                            Case "/"
                                ' Ratio: recompute from annual totals
                                q1c = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR
                                q4c = q1c + QS_QUARTERS_PER_YEAR - 1
                                Dim annALetter As String
                                annALetter = QSColLetter(annDataCol)
                                ws.Cells(curRow, annDataCol).formula = "=IFERROR(" & _
                                    annALetter & opAEntRow & "/" & annALetter & opBEntRow & ",0)"
                            Case Else
                                q1c = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR
                                q4c = q1c + QS_QUARTERS_PER_YEAR - 1
                                ws.Cells(curRow, annDataCol).formula = "=SUM(" & _
                                    QSColLetter(q1c) & curRow & ":" & QSColLetter(q4c) & curRow & ")"
                        End Select
                        ws.Cells(curRow, annDataCol).Interior.Color = RGB(217, 217, 217)
                    Next yr
                End If

                ' Apply number format
                Dim dFmt As String
                dFmt = KernelConfig.GetFormat(dMetricName)
                If Len(dFmt) > 0 Then
                    lastDCol = QS_DATA_START_COL + numYears * QS_COLS_PER_YEAR - 1
                    ws.Range(ws.Cells(curRow, QS_DATA_START_COL), ws.Cells(curRow, lastDCol)).NumberFormat = dFmt
                End If
            Next e

            ' Total row for derived metric
            curRow = curRow + 1
            ws.Cells(curRow, 1).Value = "QS_" & UCase(Left(dMetricName, 4)) & "_TOTAL"
            ws.Cells(curRow, 2).Value = "Total " & dMetricName
            ws.Cells(curRow, 2).Font.Bold = True

            Dim opATotalRow As Long
            opATotalRow = FindRowIDInSheet(ws, opARowID)
            Dim opBTotalRow As Long
            opBTotalRow = FindRowIDInSheet(ws, opBRowID)

            If opATotalRow > 0 And opBTotalRow > 0 Then
                For yr = 1 To numYears
                    For q = 1 To QS_QUARTERS_PER_YEAR
                        dataCol = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + (q - 1)
                        cLtr = QSColLetter(dataCol)
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
                        ws.Cells(curRow, dataCol).formula = dFormula
                    Next q
                    annDataCol = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + QS_QUARTERS_PER_YEAR
                    Select Case opStr
                        Case "/"
                            annALetter = QSColLetter(annDataCol)
                            ws.Cells(curRow, annDataCol).formula = "=IFERROR(" & _
                                annALetter & opATotalRow & "/" & annALetter & opBTotalRow & ",0)"
                        Case Else
                            q1c = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR
                            q4c = q1c + QS_QUARTERS_PER_YEAR - 1
                            ws.Cells(curRow, annDataCol).formula = "=SUM(" & _
                                QSColLetter(q1c) & curRow & ":" & QSColLetter(q4c) & curRow & ")"
                    End Select
                    ws.Cells(curRow, annDataCol).Interior.Color = RGB(217, 217, 217)
                Next yr
            End If

            ' Apply number format to total row
            If Len(dFmt) > 0 Then
                lastDCol = QS_DATA_START_COL + numYears * QS_COLS_PER_YEAR - 1
                ws.Range(ws.Cells(curRow, QS_DATA_START_COL), ws.Cells(curRow, lastDCol)).NumberFormat = dFmt
            End If

            ' Bottom border on total row
            lastDCol = QS_DATA_START_COL + numYears * QS_COLS_PER_YEAR - 1
            ws.Range(ws.Cells(curRow, 2), ws.Cells(curRow, lastDCol)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            ws.Range(ws.Cells(curRow, 2), ws.Cells(curRow, lastDCol)).Borders(xlEdgeBottom).Weight = xlThin
NextDerivedMetric:
        Next dIdx
    End If

    ' Hide RowID column
    ws.Columns(1).Hidden = True

    ' AutoFit
    On Error Resume Next
    ws.Columns.AutoFit
    On Error GoTo ErrHandler

    KernelConfig.LogError SEV_INFO, "SampleDomainEngine", "I-360", _
        "AggregateToQuarterly completed: " & entityCount & " entities, " & _
        numQuarters & " quarters on QuarterlySummary tab.", ""

    Exit Sub

ErrHandler:
    KernelConfig.LogError SEV_ERROR, "SampleDomainEngine", "E-361", _
        "AggregateToQuarterly error: " & Err.Description, _
        "MANUAL BYPASS: Create QuarterlySummary tab manually with quarterly " & _
        "sums from the Detail tab. Use RowIDs in Column A."
End Sub


' =============================================================================
' QSColLetter -- column number to letter helper (local to avoid cross-module)
' =============================================================================
Private Function QSColLetter(colNum As Long) As String
    Dim n As Long
    n = colNum
    QSColLetter = ""
    Do While n > 0
        Dim remainder As Long
        remainder = (n - 1) Mod 26
        QSColLetter = Chr(65 + remainder) & QSColLetter
        n = (n - 1) \ 26
    Loop
End Function


' =============================================================================
' FindRowIDInSheet -- finds a RowID in column A, returns row or 0
' Uses array read for performance instead of cell-by-cell scan
' =============================================================================
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


' =============================================================================
' ParseSimpleRule -- parses "A - B" into operandA, operator, operandB
' =============================================================================
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


Private Function DetectEntityCountForValidation() As Long
    Dim wsInputs As Worksheet
    Set wsInputs = ThisWorkbook.Sheets(TAB_INPUTS)

    ' Find EntityName row
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
        DetectEntityCountForValidation = 0
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

    DetectEntityCountForValidation = cnt
End Function
