Attribute VB_Name = "Ins_GranularCSV"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' Ins_GranularCSV.bas
' Purpose: Writes granular triangle CSV at Program x LossType x ExposureMonth x
'          CalendarMonth grain. 157 columns matching UWM v2.9.4 schema.
'          Called from InsuranceDomainEngine.Execute() after DevelopLosses.
' Phase 11A Addendum. ASCII only (AP-06).

Private Const LAYER_NAMES As String = "Attritional,Seasonal,CAT"

' SafeDiv helper
Private Function SD(num As Double, den As Double) As Double
    If den = 0 Then SD = 0 Else SD = num / den
End Function

Public Sub WriteGranularCSV()
    ' Access InsuranceDomainEngine public state
    Dim nProg As Long
    Dim horizon As Long
    Dim scenName As String
    Dim safeName As String
    Dim ch As Long
    Dim c As String
    Dim outDir As String
    Dim ts As String
    Dim fPath As String
    Dim fNum As Integer
    Dim hdr As String
    Dim blocks As Variant
    Dim metrics As Variant
    Dim bi As Long
    Dim mi As Long
    Dim lyrNames() As String
    Dim p As Long
    Dim lyr As Long
    Dim ep As Long
    Dim cm As Long

    ' Variables used inside the loops
    Dim itdGPaid As Double
    Dim itdGCI As Double
    Dim itdGUlt As Double
    Dim itdGRpt As Double
    Dim itdGCls As Double
    Dim itdGCntUlt As Double
    Dim itdGWP As Double
    Dim itdGEP As Double
    Dim age As Long
    Dim ageAdj As Double
    Dim prevAgeAdj As Double
    Dim paidPct As Double
    Dim ciPct As Double
    Dim rptPct As Double
    Dim clsPct As Double
    Dim prevPaid As Double
    Dim prevCI As Double
    Dim prevRpt As Double
    Dim prevCls As Double
    Dim ultLoss As Double
    Dim ultCnt As Double

    ' MTD Gross loss metrics
    Dim mgPaid As Double
    Dim mgCI As Double
    Dim mgUlt As Double
    Dim mgRpt As Double
    Dim mgCls As Double
    Dim mgCntUlt As Double

    ' ITD Gross loss metrics
    Dim igPaid As Double
    Dim igCaseRsv As Double
    Dim igCI As Double
    Dim igIBNR As Double
    Dim igUnpaid As Double
    Dim igUlt As Double
    Dim igCls As Double
    Dim igOpen As Double
    Dim igRpt As Double
    Dim igIBNRCt As Double
    Dim igUnclCt As Double
    Dim igUltCt As Double

    ' MTD Gross balance metrics
    Dim mgCaseRsv As Double
    Dim mgIBNR As Double
    Dim mgUnpaid As Double
    Dim mgOpen As Double
    Dim mgIBNRCt As Double
    Dim mgUnclCt As Double

    ' MTD Gross premium
    Dim mgWP As Double
    Dim mgEP As Double
    Dim mgWComm As Double
    Dim mgEComm As Double
    Dim mgWFF As Double
    Dim mgEFF As Double
    Dim termMo As Long
    Dim eFrac As Double
    Dim rateYr As Long

    ' Ceded variables
    Dim ryC As Long
    Dim cdPct As Double

    ' MTD Ceded
    Dim mcPaid As Double
    Dim mcCI As Double
    Dim mcUlt As Double
    Dim mcCaseRsv As Double
    Dim mcIBNR As Double
    Dim mcUnpaid As Double
    Dim mcRpt As Double
    Dim mcCls As Double
    Dim mcCntUlt As Double
    Dim mcOpen As Double
    Dim mcIBNRCt As Double
    Dim mcUnclCt As Double
    Dim mcWP As Double
    Dim mcEP As Double
    Dim mcWComm As Double
    Dim mcEComm As Double
    Dim mcWFF As Double
    Dim mcEFF As Double

    ' ITD Ceded
    Dim icPaid As Double
    Dim icCI As Double
    Dim icUlt As Double
    Dim icCaseRsv As Double
    Dim icIBNR As Double
    Dim icUnpaid As Double
    Dim icCls As Double
    Dim icOpen As Double
    Dim icRpt As Double
    Dim icIBNRCt As Double
    Dim icUnclCt As Double
    Dim icUltCt As Double

    ' Dimension variables
    Dim calYr As Long
    Dim calMo As Long
    Dim calQ As Long
    Dim epYr As Long
    Dim epMo As Long
    Dim epQ As Long
    Dim devQ As Long
    Dim devY As Long

    ' CSV line builder
    Dim ln As String

    ' Severity variables - MTD Gross
    Dim mgClosedSev As Double
    Dim mgOpenSev As Double
    Dim mgRptSev As Double
    Dim mgIBNRSev As Double
    Dim mgUnclSev As Double
    Dim mgUltSev As Double

    ' Severity variables - MTD Ceded
    Dim mcClosedSev As Double
    Dim mcOpenSev As Double
    Dim mcRptSev As Double
    Dim mcIBNRSev As Double
    Dim mcUnclSev As Double
    Dim mcUltSev As Double

    ' MTD Net
    Dim mnPaid As Double
    Dim mnCI As Double
    Dim mnUlt As Double
    Dim mnCaseRsv As Double
    Dim mnIBNR As Double
    Dim mnUnpaid As Double
    Dim mnRpt As Double
    Dim mnCls As Double
    Dim mnCntUlt As Double
    Dim mnOpen As Double
    Dim mnIBNRCt As Double
    Dim mnUnclCt As Double
    Dim mnWP As Double
    Dim mnEP As Double
    Dim mnWComm As Double
    Dim mnEComm As Double

    ' ITD Net for severities
    Dim inPaid As Double
    Dim inCI As Double
    Dim inCaseRsv As Double
    Dim inIBNR As Double
    Dim inUnpaid As Double
    Dim inCls As Double
    Dim inOpen As Double
    Dim inRpt As Double
    Dim inIBNRCt As Double
    Dim inUnclCt As Double
    Dim inUltCt As Double
    Dim inUlt As Double

    ' Net severity
    Dim mnClosedSev As Double
    Dim mnOpenSev As Double
    Dim mnRptSev As Double
    Dim mnIBNRSev As Double
    Dim mnUnclSev As Double
    Dim mnUltSev As Double

    ' ITD Gross premium
    Dim igWComm As Double
    Dim igEComm As Double
    Dim igWFF As Double
    Dim igEFF As Double

    ' ITD Ceded premium
    Dim icWP As Double
    Dim icEP As Double
    Dim icWComm As Double
    Dim icEComm As Double

    ' ITD Net premium
    Dim inWP As Double
    Dim inEP As Double
    Dim inWComm As Double
    Dim inEComm As Double

    ' ========== Begin executable code ==========

    nProg = InsuranceDomainEngine.m_numProgs
    horizon = InsuranceDomainEngine.m_horizon

    ' Get scenario name from Assumptions tab
    scenName = CStr(KernelConfig.InputValue("Global Assumptions", "ScenarioName", 1))
    If Len(scenName) = 0 Then scenName = "Base"

    ' Sanitize scenario name for filename
    safeName = scenName
    For ch = 1 To Len(safeName)
        c = Mid(safeName, ch, 1)
        If InStr(1, " /\:*?""<>|", c) > 0 Then Mid(safeName, ch, 1) = "_"
    Next ch

    ' Build output path using shared utility
    outDir = KernelFormHelpers.EnsureOutputDir()

    ts = Format(Now, "yyyymmdd_hhnnss")
    fPath = outDir & "\granular_detail_" & safeName & "_" & ts & ".csv"

    ' Store path for KernelSnapshot access
    InsuranceDomainEngine.GranularCSVPath = fPath

    fNum = FreeFile
    Open fPath For Output As #fNum

    ' Write header (157 columns: 13 dims + 144 metrics)
    hdr = "ScenarioID,BusinessUnit,Program,LossType," & _
          "CalMonth,CalQuarter,CalYear," & _
          "ExposureMonth,ExposureQuarter,ExposureYear," & _
          "DevAgeMo,DevAgeQtr,DevAgeYr"

    ' 6 blocks x 24 metrics
    blocks = Array("MTD_Gross", "MTD_Ceded", "MTD_Net", _
                   "ITD_Gross", "ITD_Ceded", "ITD_Net")
    metrics = Array("WP", "EP", "WComm", "EComm", "WFrontFee", "EFrontFee", _
                    "Paid", "CaseRsv", "CaseInc", "IBNR", "Unpaid", "Ult", _
                    "ClosedCt", "OpenCt", "RptCt", "IBNRCt", "UnclCt", "UltCt", _
                    "ClosedSev", "OpenSev", "RptSev", "IBNRSev", "UnclSev", "UltSev")

    For bi = LBound(blocks) To UBound(blocks)
        For mi = LBound(metrics) To UBound(metrics)
            hdr = hdr & "," & blocks(bi) & "_" & metrics(mi)
        Next mi
    Next bi
    Print #fNum, hdr

    ' Layer name lookup
    lyrNames = Split(LAYER_NAMES, ",")

    ' Iterate: p -> lyr -> ep -> cm (optimal for ITD accumulation)
    For p = 1 To nProg
        For lyr = 1 To 3
            If Not InsuranceDomainEngine.m_lyrActive(p, lyr) Then GoTo NextGLayer

            For ep = 1 To horizon
                ' Skip exposure months with no ultimate
                If InsuranceDomainEngine.m_ultMon(p, lyr, ep) = 0 And _
                   InsuranceDomainEngine.m_cntUlt(p, lyr, ep) = 0 Then GoTo NextGEP

                ' ITD accumulators (reset per exposure month)
                itdGPaid = 0
                itdGCI = 0
                itdGUlt = 0
                itdGRpt = 0
                itdGCls = 0
                itdGCntUlt = 0
                itdGWP = 0
                itdGEP = 0

                For cm = ep To horizon
                    age = cm - ep + 1

                    ' Skip if beyond dev endpoint
                    If age > InsuranceDomainEngine.m_devEnd(p) Then Exit For

                    ' DE-08: mid-month average written date offset
                    ageAdj = CDbl(age) - 0.5
                    prevAgeAdj = ageAdj - 1#

                    ' Evaluate curves
                    paidPct = Ext_CurveLib.EvaluateCurve( _
                        InsuranceDomainEngine.m_curves(p, lyr).distPd, _
                        InsuranceDomainEngine.m_curves(p, lyr).p1Pd, _
                        InsuranceDomainEngine.m_curves(p, lyr).p2Pd, _
                        ageAdj, InsuranceDomainEngine.m_curves(p, lyr).maxAgePd)

                    ciPct = Ext_CurveLib.EvaluateCurve( _
                        InsuranceDomainEngine.m_curves(p, lyr).distCI, _
                        InsuranceDomainEngine.m_curves(p, lyr).p1CI, _
                        InsuranceDomainEngine.m_curves(p, lyr).p2CI, _
                        ageAdj, InsuranceDomainEngine.m_curves(p, lyr).maxAgeCI)

                    rptPct = Ext_CurveLib.EvaluateCurve( _
                        InsuranceDomainEngine.m_curves(p, lyr).distRC, _
                        InsuranceDomainEngine.m_curves(p, lyr).p1RC, _
                        InsuranceDomainEngine.m_curves(p, lyr).p2RC, _
                        ageAdj, InsuranceDomainEngine.m_curves(p, lyr).maxAgeRC)

                    clsPct = Ext_CurveLib.EvaluateCurve( _
                        InsuranceDomainEngine.m_curves(p, lyr).distCC, _
                        InsuranceDomainEngine.m_curves(p, lyr).p1CC, _
                        InsuranceDomainEngine.m_curves(p, lyr).p2CC, _
                        ageAdj, InsuranceDomainEngine.m_curves(p, lyr).maxAgeCC)

                    ' Previous-age CDFs for MTD calculation
                    prevPaid = 0
                    prevCI = 0
                    prevRpt = 0
                    prevCls = 0
                    If age > 1 Then
                        prevPaid = Ext_CurveLib.EvaluateCurve( _
                            InsuranceDomainEngine.m_curves(p, lyr).distPd, _
                            InsuranceDomainEngine.m_curves(p, lyr).p1Pd, _
                            InsuranceDomainEngine.m_curves(p, lyr).p2Pd, _
                            prevAgeAdj, InsuranceDomainEngine.m_curves(p, lyr).maxAgePd)
                        prevCI = Ext_CurveLib.EvaluateCurve( _
                            InsuranceDomainEngine.m_curves(p, lyr).distCI, _
                            InsuranceDomainEngine.m_curves(p, lyr).p1CI, _
                            InsuranceDomainEngine.m_curves(p, lyr).p2CI, _
                            prevAgeAdj, InsuranceDomainEngine.m_curves(p, lyr).maxAgeCI)
                        prevRpt = Ext_CurveLib.EvaluateCurve( _
                            InsuranceDomainEngine.m_curves(p, lyr).distRC, _
                            InsuranceDomainEngine.m_curves(p, lyr).p1RC, _
                            InsuranceDomainEngine.m_curves(p, lyr).p2RC, _
                            prevAgeAdj, InsuranceDomainEngine.m_curves(p, lyr).maxAgeRC)
                        prevCls = Ext_CurveLib.EvaluateCurve( _
                            InsuranceDomainEngine.m_curves(p, lyr).distCC, _
                            InsuranceDomainEngine.m_curves(p, lyr).p1CC, _
                            InsuranceDomainEngine.m_curves(p, lyr).p2CC, _
                            prevAgeAdj, InsuranceDomainEngine.m_curves(p, lyr).maxAgeCC)
                    End If

                    ' Ultimate amounts for this (p, lyr, ep)
                    ultLoss = InsuranceDomainEngine.m_ultMon(p, lyr, ep)
                    ultCnt = InsuranceDomainEngine.m_cntUlt(p, lyr, ep)

                    ' --- MTD Gross loss metrics ---
                    mgPaid = ultLoss * (paidPct - prevPaid)
                    mgCI = ultLoss * (ciPct - prevCI)
                    If age = 1 Then mgUlt = ultLoss Else mgUlt = 0
                    mgRpt = ultCnt * (rptPct - prevRpt)
                    mgCls = ultCnt * (clsPct - prevCls)
                    If age = 1 Then mgCntUlt = ultCnt Else mgCntUlt = 0

                    ' ITD Gross loss metrics (cumulative CDF x ultimate)
                    igPaid = ultLoss * paidPct
                    igCaseRsv = ultLoss * ciPct - ultLoss * paidPct
                    igCI = ultLoss * ciPct
                    igIBNR = ultLoss - igCI
                    igUnpaid = ultLoss - igPaid
                    igUlt = ultLoss
                    igCls = ultCnt * clsPct
                    igOpen = ultCnt * rptPct - igCls
                    igRpt = ultCnt * rptPct
                    igIBNRCt = ultCnt - igRpt
                    igUnclCt = ultCnt - igCls
                    igUltCt = ultCnt

                    ' MTD Gross balance metrics (derived from ITD)
                    mgCaseRsv = igCaseRsv  ' Balance = EOP
                    mgIBNR = igIBNR
                    mgUnpaid = igUnpaid
                    mgOpen = igOpen
                    mgIBNRCt = igIBNRCt
                    mgUnclCt = igUnclCt

                    ' MTD Gross premium (layer 1 only, at cm=ep)
                    mgWP = 0
                    mgEP = 0
                    mgWComm = 0
                    mgEComm = 0
                    mgWFF = 0
                    mgEFF = 0

                    If lyr = 1 Then
                        If cm = ep Then mgWP = InsuranceDomainEngine.m_wpMon(p, ep)
                        ' Earning fraction for this cm from ep
                        termMo = InsuranceDomainEngine.m_progTerm(p)
                        If termMo <= 0 Then termMo = 12
                        If cm >= ep And cm <= ep + termMo Then
                            If cm = ep Then
                                eFrac = 1# / (2# * CDbl(termMo))
                            ElseIf cm = ep + termMo Then
                                eFrac = 1# / (2# * CDbl(termMo))
                            Else
                                eFrac = 1# / CDbl(termMo)
                            End If
                            mgEP = InsuranceDomainEngine.m_wpMon(p, ep) * eFrac
                        End If

                        ' Rate year for commission/reins
                        rateYr = ((cm - 1) \ 12) + 1
                        If rateYr > 5 Then rateYr = 5
                        mgWComm = mgWP * InsuranceDomainEngine.m_commRate(p, rateYr)
                        mgEComm = mgEP * InsuranceDomainEngine.m_commRate(p, rateYr)
                        mgWFF = mgWP * InsuranceDomainEngine.m_reinsFrontFee(p, rateYr)
                        mgEFF = mgEP * InsuranceDomainEngine.m_reinsFrontFee(p, rateYr)
                    End If

                    ' ITD premium accumulators
                    itdGWP = itdGWP + mgWP
                    itdGEP = itdGEP + mgEP

                    ' Cede pct for this rate year
                    ryC = ((cm - 1) \ 12) + 1
                    If ryC > 5 Then ryC = 5
                    If lyr <= 2 Then
                        cdPct = InsuranceDomainEngine.m_reinsCedePct(p, ryC)
                    Else
                        cdPct = 0  ' CAT not ceded via QS
                    End If

                    ' --- MTD Ceded (QS on Attr+Seas) ---
                    mcPaid = mgPaid * cdPct
                    mcCI = mgCI * cdPct
                    mcUlt = mgUlt * cdPct
                    mcCaseRsv = mgCaseRsv * cdPct
                    mcIBNR = mgIBNR * cdPct
                    mcUnpaid = mgUnpaid * cdPct
                    mcRpt = mgRpt * cdPct
                    mcCls = mgCls * cdPct
                    mcCntUlt = mgCntUlt  ' Counts not ceded
                    mcOpen = mgOpen * cdPct
                    mcIBNRCt = mgIBNRCt  ' Counts = gross
                    mcUnclCt = mgUnclCt
                    mcWP = mgWP * cdPct
                    mcEP = mgEP * cdPct
                    mcWComm = mgWComm * cdPct
                    mcEComm = mgEComm * cdPct
                    mcWFF = 0  ' No ceded fronting fee
                    mcEFF = 0

                    ' --- ITD Ceded ---
                    icPaid = igPaid * cdPct
                    icCI = igCI * cdPct
                    icUlt = igUlt * cdPct
                    icCaseRsv = igCaseRsv * cdPct
                    icIBNR = igIBNR * cdPct
                    icUnpaid = igUnpaid * cdPct
                    icCls = igCls * cdPct
                    icOpen = igOpen * cdPct
                    icRpt = igRpt * cdPct
                    icIBNRCt = igIBNRCt  ' Counts = gross
                    icUnclCt = igUnclCt
                    icUltCt = igUltCt

                    ' --- Dimensions ---
                    calYr = ((cm - 1) \ 12) + 1
                    calMo = ((cm - 1) Mod 12) + 1
                    calQ = ((calMo - 1) \ 3) + 1
                    epYr = ((ep - 1) \ 12) + 1
                    epMo = ((ep - 1) Mod 12) + 1
                    epQ = ((epMo - 1) \ 3) + 1
                    devQ = ((age - 1) \ 3) + 1
                    devY = ((age - 1) \ 12) + 1

                    ' Build CSV line
                    ' 13 dimensions
                    ln = scenName & "," & _
                         InsuranceDomainEngine.m_progBU(p) & "," & _
                         InsuranceDomainEngine.m_progName(p) & "," & _
                         lyrNames(lyr - 1) & "," & _
                         "Y" & Format(calYr, "0000") & "_M" & Format(calMo, "00") & "," & _
                         "Y" & Format(calYr, "0000") & "_Q" & calQ & "," & _
                         "Y" & Format(calYr, "0000") & "," & _
                         "E" & Format(epYr, "0000") & "_M" & Format(epMo, "00") & "," & _
                         "E" & Format(epYr, "0000") & "_Q" & epQ & "," & _
                         "E" & Format(epYr, "0000") & "," & _
                         age & "," & devQ & "," & devY

                    ' MTD_Gross (24 metrics)
                    ' Severities
                    mgClosedSev = SD(igPaid, igCls)
                    mgOpenSev = SD(igCaseRsv, igOpen)
                    mgRptSev = SD(igCI, igRpt)
                    mgIBNRSev = SD(igIBNR, igIBNRCt)
                    mgUnclSev = SD(igUnpaid, igUnclCt)
                    mgUltSev = SD(igUlt, igUltCt)

                    ln = ln & "," & mgWP & "," & mgEP & "," & mgWComm & "," & mgEComm & _
                         "," & mgWFF & "," & mgEFF & _
                         "," & mgPaid & "," & mgCaseRsv & "," & mgCI & _
                         "," & mgIBNR & "," & mgUnpaid & "," & mgUlt & _
                         "," & mgCls & "," & mgOpen & "," & mgRpt & _
                         "," & mgIBNRCt & "," & mgUnclCt & "," & mgCntUlt & _
                         "," & mgClosedSev & "," & mgOpenSev & "," & mgRptSev & _
                         "," & mgIBNRSev & "," & mgUnclSev & "," & mgUltSev

                    ' MTD_Ceded (24 metrics)
                    mcClosedSev = SD(icPaid, icCls)
                    mcOpenSev = SD(icCaseRsv, icOpen)
                    mcRptSev = SD(icCI, icRpt)
                    mcIBNRSev = SD(icIBNR, icIBNRCt)
                    mcUnclSev = SD(icUnpaid, icUnclCt)
                    mcUltSev = SD(icUlt, icUltCt)

                    ln = ln & "," & mcWP & "," & mcEP & "," & mcWComm & "," & mcEComm & _
                         "," & mcWFF & "," & mcEFF & _
                         "," & mcPaid & "," & mcCaseRsv & "," & mcCI & _
                         "," & mcIBNR & "," & mcUnpaid & "," & mcUlt & _
                         "," & mcCls & "," & mcOpen & "," & mcRpt & _
                         "," & mcIBNRCt & "," & mcUnclCt & "," & mcCntUlt & _
                         "," & mcClosedSev & "," & mcOpenSev & "," & mcRptSev & _
                         "," & mcIBNRSev & "," & mcUnclSev & "," & mcUltSev

                    ' MTD_Net = Gross - Ceded (24 metrics)
                    mnPaid = mgPaid - mcPaid
                    mnCI = mgCI - mcCI
                    mnUlt = mgUlt - mcUlt
                    mnCaseRsv = mgCaseRsv - mcCaseRsv
                    mnIBNR = mgIBNR - mcIBNR
                    mnUnpaid = mgUnpaid - mcUnpaid
                    mnRpt = mgRpt - mcRpt
                    mnCls = mgCls - mcCls
                    mnCntUlt = mgCntUlt - mcCntUlt
                    mnOpen = mgOpen - mcOpen
                    mnIBNRCt = mgIBNRCt - mcIBNRCt
                    mnUnclCt = mgUnclCt - mcUnclCt
                    mnWP = mgWP - mcWP
                    mnEP = mgEP - mcEP
                    mnWComm = mgWComm - mcWComm
                    mnEComm = mgEComm - mcEComm

                    ' Net ITD for severities
                    inPaid = igPaid - icPaid
                    inCI = igCI - icCI
                    inCaseRsv = igCaseRsv - icCaseRsv
                    inIBNR = igIBNR - icIBNR
                    inUnpaid = igUnpaid - icUnpaid
                    inCls = igCls - icCls
                    inOpen = igOpen - icOpen
                    inRpt = igRpt - icRpt
                    inIBNRCt = igIBNRCt - icIBNRCt
                    inUnclCt = igUnclCt - icUnclCt
                    inUltCt = igUltCt - icUltCt
                    inUlt = igUlt - icUlt

                    mnClosedSev = SD(inPaid, inCls)
                    mnOpenSev = SD(inCaseRsv, inOpen)
                    mnRptSev = SD(inCI, inRpt)
                    mnIBNRSev = SD(inIBNR, inIBNRCt)
                    mnUnclSev = SD(inUnpaid, inUnclCt)
                    mnUltSev = SD(inUlt, inUltCt)

                    ln = ln & "," & mnWP & "," & mnEP & "," & mnWComm & "," & mnEComm & _
                         "," & 0 & "," & 0 & _
                         "," & mnPaid & "," & mnCaseRsv & "," & mnCI & _
                         "," & mnIBNR & "," & mnUnpaid & "," & mnUlt & _
                         "," & mnCls & "," & mnOpen & "," & mnRpt & _
                         "," & mnIBNRCt & "," & mnUnclCt & "," & mnCntUlt & _
                         "," & mnClosedSev & "," & mnOpenSev & "," & mnRptSev & _
                         "," & mnIBNRSev & "," & mnUnclSev & "," & mnUltSev

                    ' ITD_Gross (24 metrics)
                    igWComm = itdGWP * InsuranceDomainEngine.m_commRate(p, ryC)
                    igEComm = itdGEP * InsuranceDomainEngine.m_commRate(p, ryC)
                    igWFF = itdGWP * InsuranceDomainEngine.m_reinsFrontFee(p, ryC)
                    igEFF = itdGEP * InsuranceDomainEngine.m_reinsFrontFee(p, ryC)

                    ln = ln & "," & itdGWP & "," & itdGEP & "," & igWComm & "," & igEComm & _
                         "," & igWFF & "," & igEFF & _
                         "," & igPaid & "," & igCaseRsv & "," & igCI & _
                         "," & igIBNR & "," & igUnpaid & "," & igUlt & _
                         "," & igCls & "," & igOpen & "," & igRpt & _
                         "," & igIBNRCt & "," & igUnclCt & "," & igUltCt & _
                         "," & mgClosedSev & "," & mgOpenSev & "," & mgRptSev & _
                         "," & mgIBNRSev & "," & mgUnclSev & "," & mgUltSev

                    ' ITD_Ceded (24 metrics)
                    icWP = itdGWP * cdPct
                    icEP = itdGEP * cdPct
                    icWComm = igWComm * cdPct
                    icEComm = igEComm * cdPct

                    ln = ln & "," & icWP & "," & icEP & "," & icWComm & "," & icEComm & _
                         "," & 0 & "," & 0 & _
                         "," & icPaid & "," & icCaseRsv & "," & icCI & _
                         "," & icIBNR & "," & icUnpaid & "," & icUlt & _
                         "," & icCls & "," & icOpen & "," & icRpt & _
                         "," & icIBNRCt & "," & icUnclCt & "," & icUltCt & _
                         "," & mcClosedSev & "," & mcOpenSev & "," & mcRptSev & _
                         "," & mcIBNRSev & "," & mcUnclSev & "," & mcUltSev

                    ' ITD_Net (24 metrics)
                    inWP = itdGWP - icWP
                    inEP = itdGEP - icEP
                    inWComm = igWComm - icWComm
                    inEComm = igEComm - icEComm

                    ln = ln & "," & inWP & "," & inEP & "," & inWComm & "," & inEComm & _
                         "," & 0 & "," & 0 & _
                         "," & inPaid & "," & inCaseRsv & "," & inCI & _
                         "," & inIBNR & "," & inUnpaid & "," & inUlt & _
                         "," & inCls & "," & inOpen & "," & inRpt & _
                         "," & inIBNRCt & "," & inUnclCt & "," & inUltCt & _
                         "," & mnClosedSev & "," & mnOpenSev & "," & mnRptSev & _
                         "," & mnIBNRSev & "," & mnUnclSev & "," & mnUltSev

                    Print #fNum, ln
                Next cm
NextGEP:
            Next ep
NextGLayer:
        Next lyr
    Next p

    Close #fNum

    KernelConfig.LogError SEV_INFO, "Ins_GranularCSV", "I-370", _
        "Granular CSV written: " & fPath, ""
End Sub
