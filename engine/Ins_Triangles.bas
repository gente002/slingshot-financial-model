Attribute VB_Name = "Ins_Triangles"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' Ins_Triangles.bas
' Purpose: Builds development triangles on Loss Triangles and Count Triangles
'          tabs. Computes values in VBA from per-exposure-month ultimate and
'          curve arrays, isolating each exposure quarter cohort.
'          Left section: cumulative amounts. Right section: Ult + % of ultimate.
'          BUG-104: Cannot use SUMIFS on Detail tab because Detail aggregates
'          across all exposure months per calendar month.
'          Registered as PostCompute transform (runs after AggregateToQuarterly).
' ASCII only (AP-06).

Private Const TRI_DEV_QTRS As Long = 20
Private Const TRI_DATA_COL As Long = 3
Private Const TRI_NUM_LAYERS As Long = 3

' Metric type constants for curve selection
Private Const MET_PAID As Long = 1
Private Const MET_CASE_INCURRED As Long = 2
Private Const MET_REPORTED_COUNT As Long = 3
Private Const MET_CLOSED_COUNT As Long = 4


Public Sub BuildTriangleTab()
    On Error GoTo ErrHandler
    BuildTriangle "Loss Triangles", "Loss Development Triangles", _
        "Cumulative Dollar Amounts (left) | Ult + % of Ultimate (right)", _
        MET_PAID, "Gross Paid", MET_CASE_INCURRED, "Gross Case Incurred", _
        "#,##0", False
    Exit Sub
ErrHandler:
    KernelConfig.LogError SEV_ERROR, "Ins_Triangles", "E-370", _
        "BuildTriangleTab error: " & Err.Description, _
        "MANUAL BYPASS: Loss Triangles tab is informational only."
End Sub


Public Sub BuildCountTriangleTab()
    On Error GoTo ErrHandler
    BuildTriangle "Count Triangles", "Count Development Triangles", _
        "Cumulative Counts (left) | Ult + % of Ultimate (right)", _
        MET_CLOSED_COUNT, "Closed Count", MET_REPORTED_COUNT, "Reported Count", _
        "#,##0", True
    Exit Sub
ErrHandler:
    KernelConfig.LogError SEV_ERROR, "Ins_Triangles", "E-371", _
        "BuildCountTriangleTab error: " & Err.Description, _
        "MANUAL BYPASS: Count Triangles tab is informational only."
End Sub


' Core triangle builder -- shared by both entry points
Private Sub BuildTriangle(tabName As String, title As String, subtitle As String, _
    met1 As Long, met1Name As String, met2 As Long, met2Name As String, _
    numFmt As String, isCount As Boolean)

    Dim nProg As Long
    nProg = InsuranceDomainEngine.m_numProgs
    If nProg = 0 Then Exit Sub

    Dim horizon As Long
    horizon = InsuranceDomainEngine.m_horizon
    If horizon <= 0 Then horizon = 60
    Dim numExpQtrs As Long
    numExpQtrs = (horizon \ 3)
    If numExpQtrs > TRI_DEV_QTRS Then numExpQtrs = TRI_DEV_QTRS

    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(tabName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = tabName
    End If

    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0
    ws.Cells.ClearContents

    ' Layout
    Dim spacerCol As Long: spacerCol = TRI_DATA_COL + TRI_DEV_QTRS
    Dim ultCol As Long: ultCol = spacerCol + 1
    Dim pctStartCol As Long: pctStartCol = ultCol + 1
    Dim lastPctCol As Long: lastPctCol = pctStartCol + TRI_DEV_QTRS - 1

    ' Header -- left
    ws.Cells(1, 2).Value = title
    ws.Range(ws.Cells(1, 2), ws.Cells(1, TRI_DATA_COL + TRI_DEV_QTRS - 1)).Font.Bold = True
    ws.Range(ws.Cells(1, 2), ws.Cells(1, TRI_DATA_COL + TRI_DEV_QTRS - 1)).Interior.Color = RGB(31, 56, 100)
    ws.Range(ws.Cells(1, 2), ws.Cells(1, TRI_DATA_COL + TRI_DEV_QTRS - 1)).Font.Color = RGB(255, 255, 255)
    ws.Cells(2, 2).Value = subtitle
    ws.Cells(2, 2).Font.Italic = True
    ws.Cells(2, 2).Font.Color = RGB(128, 128, 128)

    ' Header -- right
    ws.Range(ws.Cells(1, ultCol), ws.Cells(1, lastPctCol)).Font.Bold = True
    ws.Range(ws.Cells(1, ultCol), ws.Cells(1, lastPctCol)).Interior.Color = RGB(31, 56, 100)
    ws.Range(ws.Cells(1, ultCol), ws.Cells(1, lastPctCol)).Font.Color = RGB(255, 255, 255)
    ws.Cells(1, ultCol).Value = "% of Ultimate"

    Dim curRow As Long: curRow = 4
    Dim p As Long
    Dim eq As Long
    Dim dq As Long
    Dim ep As Long
    Dim lyr As Long
    Dim age As Long
    Dim ageAdj As Double
    Dim pct As Double
    Dim ultVal As Double

    Dim metricIDs(1 To 2) As Long
    Dim metricLabels(1 To 2) As String
    metricIDs(1) = met1: metricLabels(1) = met1Name
    metricIDs(2) = met2: metricLabels(2) = met2Name

    ' === Total (All Programs) block first ===
    Dim mi As Long
    For mi = 1 To 2
        Dim metID As Long: metID = metricIDs(mi)

        ws.Cells(curRow, 2).Value = "All Programs -- " & metricLabels(mi) & IIf(isCount, "", " ($)")
        ws.Cells(curRow, 2).Font.Bold = True
        ws.Range(ws.Cells(curRow, 2), ws.Cells(curRow, TRI_DATA_COL + TRI_DEV_QTRS - 1)).Interior.Color = RGB(217, 225, 242)
        ws.Cells(curRow, ultCol).Value = "All Programs -- " & metricLabels(mi) & " (%)"
        ws.Cells(curRow, ultCol).Font.Bold = True
        ws.Range(ws.Cells(curRow, ultCol), ws.Cells(curRow, lastPctCol)).Interior.Color = RGB(217, 225, 242)

        curRow = curRow + 1
        ws.Cells(curRow, 2).Value = "Exp Qtr"
        ws.Cells(curRow, 2).Font.Bold = True
        ws.Cells(curRow, ultCol).Value = "Ult"
        ws.Cells(curRow, ultCol).Font.Bold = True
        ws.Cells(curRow, ultCol).HorizontalAlignment = xlCenter
        For dq = 1 To TRI_DEV_QTRS
            ws.Cells(curRow, TRI_DATA_COL + dq - 1).Value = "DQ" & dq
            ws.Cells(curRow, TRI_DATA_COL + dq - 1).Font.Bold = True
            ws.Cells(curRow, TRI_DATA_COL + dq - 1).HorizontalAlignment = xlCenter
            ws.Cells(curRow, pctStartCol + dq - 1).Value = "DQ" & dq
            ws.Cells(curRow, pctStartCol + dq - 1).Font.Bold = True
            ws.Cells(curRow, pctStartCol + dq - 1).HorizontalAlignment = xlCenter
        Next dq

        For eq = 1 To numExpQtrs
            curRow = curRow + 1
            Dim eqYrT As Long: eqYrT = ((eq - 1) \ 4) + 1
            Dim eqQT As Long: eqQT = ((eq - 1) Mod 4) + 1
            ws.Cells(curRow, 2).Value = "Q" & eqQT & "Y" & eqYrT

            Dim epStartT As Long: epStartT = (eq - 1) * 3 + 1
            Dim epEndT As Long: epEndT = eq * 3
            If epEndT > horizon Then epEndT = horizon

            ' Sum across ALL programs
            Dim totalUltAll As Double: totalUltAll = 0
            For p = 1 To nProg
                For ep = epStartT To epEndT
                    For lyr = 1 To TRI_NUM_LAYERS
                        If InsuranceDomainEngine.m_lyrActive(p, lyr) Then
                            If isCount Then
                                totalUltAll = totalUltAll + InsuranceDomainEngine.m_cntUlt(p, lyr, ep)
                            Else
                                totalUltAll = totalUltAll + InsuranceDomainEngine.m_ultMon(p, lyr, ep)
                            End If
                        End If
                    Next lyr
                Next ep
            Next p
            ws.Cells(curRow, ultCol).Value = totalUltAll
            ws.Cells(curRow, ultCol).NumberFormat = numFmt

            For dq = 1 To TRI_DEV_QTRS
                Dim endCmT As Long
                endCmT = epStartT - 1 + dq * 3
                Dim cumAll As Double: cumAll = 0
                For p = 1 To nProg
                    For ep = epStartT To epEndT
                        age = endCmT - ep + 1
                        If age >= 1 Then
                            ageAdj = CDbl(age) - 0.5
                            For lyr = 1 To TRI_NUM_LAYERS
                                If InsuranceDomainEngine.m_lyrActive(p, lyr) Then
                                    If isCount Then
                                        ultVal = InsuranceDomainEngine.m_cntUlt(p, lyr, ep)
                                    Else
                                        ultVal = InsuranceDomainEngine.m_ultMon(p, lyr, ep)
                                    End If
                                    If ultVal <> 0 Then
                                        If age >= InsuranceDomainEngine.m_devEnd(p) Then
                                            pct = 1
                                        Else
                                            pct = EvalMetricCurve(p, lyr, metID, ageAdj)
                                        End If
                                        cumAll = cumAll + ultVal * pct
                                    End If
                                End If
                            Next lyr
                        End If
                    Next ep
                Next p
                ws.Cells(curRow, TRI_DATA_COL + dq - 1).Value = cumAll
                ws.Cells(curRow, TRI_DATA_COL + dq - 1).NumberFormat = numFmt
                If totalUltAll > 0 Then
                    ws.Cells(curRow, pctStartCol + dq - 1).Value = cumAll / totalUltAll
                Else
                    ws.Cells(curRow, pctStartCol + dq - 1).Value = 0
                End If
                ws.Cells(curRow, pctStartCol + dq - 1).NumberFormat = "0.0%"
            Next dq

            If eq Mod 2 = 0 Then
                ws.Range(ws.Cells(curRow, 2), ws.Cells(curRow, TRI_DATA_COL + TRI_DEV_QTRS - 1)).Interior.Color = RGB(242, 242, 242)
                ws.Range(ws.Cells(curRow, ultCol), ws.Cells(curRow, lastPctCol)).Interior.Color = RGB(242, 242, 242)
            End If
        Next eq
        curRow = curRow + 2
    Next mi

    ' === Per-program blocks ===
    For p = 1 To nProg
        Dim pName As String
        pName = InsuranceDomainEngine.m_progName(p)
        Dim pDevEnd As Long
        pDevEnd = InsuranceDomainEngine.m_devEnd(p)

        For mi = 1 To 2
            metID = metricIDs(mi)

            ' --- Section headers ---
            ws.Cells(curRow, 2).Value = pName & " -- " & metricLabels(mi) & IIf(isCount, "", " ($)")
            ws.Cells(curRow, 2).Font.Bold = True
            ws.Range(ws.Cells(curRow, 2), ws.Cells(curRow, TRI_DATA_COL + TRI_DEV_QTRS - 1)).Interior.Color = RGB(217, 225, 242)
            ws.Cells(curRow, ultCol).Value = pName & " -- " & metricLabels(mi) & " (%)"
            ws.Cells(curRow, ultCol).Font.Bold = True
            ws.Range(ws.Cells(curRow, ultCol), ws.Cells(curRow, lastPctCol)).Interior.Color = RGB(217, 225, 242)

            ' DQ headers
            curRow = curRow + 1
            ws.Cells(curRow, 2).Value = "Exp Qtr"
            ws.Cells(curRow, 2).Font.Bold = True
            ws.Cells(curRow, ultCol).Value = "Ult"
            ws.Cells(curRow, ultCol).Font.Bold = True
            ws.Cells(curRow, ultCol).HorizontalAlignment = xlCenter
            For dq = 1 To TRI_DEV_QTRS
                ws.Cells(curRow, TRI_DATA_COL + dq - 1).Value = "DQ" & dq
                ws.Cells(curRow, TRI_DATA_COL + dq - 1).Font.Bold = True
                ws.Cells(curRow, TRI_DATA_COL + dq - 1).HorizontalAlignment = xlCenter
                ws.Cells(curRow, pctStartCol + dq - 1).Value = "DQ" & dq
                ws.Cells(curRow, pctStartCol + dq - 1).Font.Bold = True
                ws.Cells(curRow, pctStartCol + dq - 1).HorizontalAlignment = xlCenter
            Next dq

            ' Data rows -- Accident Quarter triangle
            ' BUG-117: m_ultMon is EP-based. No epScale. Direct computation.
            ' Same m_ultMon * CDF as DevelopLosses = single source of truth.
            For eq = 1 To numExpQtrs
                curRow = curRow + 1
                Dim eqYr As Long: eqYr = ((eq - 1) \ 4) + 1
                Dim eqQ As Long: eqQ = ((eq - 1) Mod 4) + 1
                ws.Cells(curRow, 2).Value = "Q" & eqQ & "Y" & eqYr

                Dim epStart As Long: epStart = (eq - 1) * 3 + 1
                Dim epEnd As Long: epEnd = eq * 3
                If epEnd > horizon Then epEnd = horizon

                ' EP-based ultimate for this accident quarter
                Dim totalEpUlt As Double: totalEpUlt = 0
                For ep = epStart To epEnd
                    For lyr = 1 To TRI_NUM_LAYERS
                        If InsuranceDomainEngine.m_lyrActive(p, lyr) Then
                            If isCount Then
                                totalEpUlt = totalEpUlt + InsuranceDomainEngine.m_cntUlt(p, lyr, ep)
                            Else
                                totalEpUlt = totalEpUlt + InsuranceDomainEngine.m_ultMon(p, lyr, ep)
                            End If
                        End If
                    Next lyr
                Next ep
                ws.Cells(curRow, ultCol).Value = totalEpUlt
                ws.Cells(curRow, ultCol).NumberFormat = numFmt

                ' Dev quarter columns
                For dq = 1 To TRI_DEV_QTRS
                    Dim endCm As Long
                    endCm = epStart - 1 + dq * 3

                    Dim cumMetric As Double: cumMetric = 0

                    For ep = epStart To epEnd
                        age = endCm - ep + 1
                        If age < 1 Then GoTo NextTriEP

                        ageAdj = CDbl(age) - 0.5

                        For lyr = 1 To TRI_NUM_LAYERS
                            If Not InsuranceDomainEngine.m_lyrActive(p, lyr) Then GoTo NextTriLyr

                            If isCount Then
                                ultVal = InsuranceDomainEngine.m_cntUlt(p, lyr, ep)
                            Else
                                ultVal = InsuranceDomainEngine.m_ultMon(p, lyr, ep)
                            End If
                            If ultVal = 0 Then GoTo NextTriLyr

                            If age >= pDevEnd Then
                                pct = 1
                            Else
                                pct = EvalMetricCurve(p, lyr, metID, ageAdj)
                            End If

                            cumMetric = cumMetric + ultVal * pct
NextTriLyr:
                        Next lyr
NextTriEP:
                    Next ep

                    Dim dolCol As Long: dolCol = TRI_DATA_COL + dq - 1
                    ws.Cells(curRow, dolCol).Value = cumMetric
                    ws.Cells(curRow, dolCol).NumberFormat = numFmt

                    Dim pctCol As Long: pctCol = pctStartCol + dq - 1
                    If totalEpUlt > 0 Then
                        ws.Cells(curRow, pctCol).Value = cumMetric / totalEpUlt
                    Else
                        ws.Cells(curRow, pctCol).Value = 0
                    End If
                    ws.Cells(curRow, pctCol).NumberFormat = "0.0%"
                Next dq

                If eq Mod 2 = 0 Then
                    ws.Range(ws.Cells(curRow, 2), _
                        ws.Cells(curRow, TRI_DATA_COL + TRI_DEV_QTRS - 1)).Interior.Color = RGB(242, 242, 242)
                    ws.Range(ws.Cells(curRow, ultCol), _
                        ws.Cells(curRow, lastPctCol)).Interior.Color = RGB(242, 242, 242)
                End If
            Next eq

            curRow = curRow + 2
        Next mi
    Next p

    ' Column widths
    ws.Columns(2).ColumnWidth = 12
    Dim cw As Long
    For cw = TRI_DATA_COL To TRI_DATA_COL + TRI_DEV_QTRS - 1
        ws.Columns(cw).ColumnWidth = 11
    Next cw
    ws.Columns(spacerCol).ColumnWidth = 3
    ws.Columns(ultCol).ColumnWidth = 11
    For cw = pctStartCol To lastPctCol
        ws.Columns(cw).ColumnWidth = 8
    Next cw
    ws.Columns(1).Hidden = True

    KernelConfig.LogError SEV_INFO, "Ins_Triangles", "I-370", _
        tabName & " completed: " & nProg & " programs, " & _
        numExpQtrs & " exposure quarters, " & TRI_DEV_QTRS & " dev quarters.", ""
End Sub


' EvalMetricCurve -- evaluates the correct curve for a given metric type
Private Function EvalMetricCurve(p As Long, lyr As Long, metID As Long, _
    ageAdj As Double) As Double
    Select Case metID
        Case MET_PAID
            EvalMetricCurve = Ext_CurveLib.EvaluateCurve( _
                InsuranceDomainEngine.m_curves(p, lyr).distPd, _
                InsuranceDomainEngine.m_curves(p, lyr).p1Pd, _
                InsuranceDomainEngine.m_curves(p, lyr).p2Pd, _
                ageAdj, InsuranceDomainEngine.m_curves(p, lyr).maxAgePd)
        Case MET_CASE_INCURRED
            EvalMetricCurve = Ext_CurveLib.EvaluateCurve( _
                InsuranceDomainEngine.m_curves(p, lyr).distCI, _
                InsuranceDomainEngine.m_curves(p, lyr).p1CI, _
                InsuranceDomainEngine.m_curves(p, lyr).p2CI, _
                ageAdj, InsuranceDomainEngine.m_curves(p, lyr).maxAgeCI)
        Case MET_REPORTED_COUNT
            EvalMetricCurve = Ext_CurveLib.EvaluateCurve( _
                InsuranceDomainEngine.m_curves(p, lyr).distRC, _
                InsuranceDomainEngine.m_curves(p, lyr).p1RC, _
                InsuranceDomainEngine.m_curves(p, lyr).p2RC, _
                ageAdj, InsuranceDomainEngine.m_curves(p, lyr).maxAgeRC)
        Case MET_CLOSED_COUNT
            EvalMetricCurve = Ext_CurveLib.EvaluateCurve( _
                InsuranceDomainEngine.m_curves(p, lyr).distCC, _
                InsuranceDomainEngine.m_curves(p, lyr).p1CC, _
                InsuranceDomainEngine.m_curves(p, lyr).p2CC, _
                ageAdj, InsuranceDomainEngine.m_curves(p, lyr).maxAgeCC)
        Case Else
            EvalMetricCurve = 0
    End Select
End Function
