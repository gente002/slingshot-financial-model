Attribute VB_Name = "Ext_CurveLib"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' Ext_CurveLib.bas
' Purpose: Domain-agnostic curve math library. Provides CDF functions,
'          interpolation, normalization, and batch evaluation utilities.
'          HookType: Standalone -- not called during pipeline. Domain code
'          and other extensions call its functions directly.
' Phase 6A extension. All errors include manual bypass instructions (AP-46).
' =============================================================================


' =============================================================================
' CDF Functions
' =============================================================================

' -----------------------------------------------------------------------------
' WeibullCDF
' Cumulative distribution function for the Weibull distribution.
' F(t) = 1 - exp(-(t/theta)^k)
' Parameters: theta = scale, k = shape
' Returns 0 for t <= 0.
' -----------------------------------------------------------------------------
Public Function WeibullCDF(t As Double, theta As Double, k As Double) As Double
    If t <= 0 Or theta <= 0 Or k <= 0 Then
        WeibullCDF = 0
        Exit Function
    End If

    WeibullCDF = 1 - Exp(-((t / theta) ^ k))
End Function


' -----------------------------------------------------------------------------
' LognormalCDF
' Cumulative distribution function for the Lognormal distribution.
' Uses the approximation of the standard normal CDF.
' F(t) = Phi((ln(t) - mu) / sigma)
' Parameters: mu = log-mean, sigma = log-std-dev
' Returns 0 for t <= 0.
' -----------------------------------------------------------------------------
Public Function LognormalCDF(t As Double, mu As Double, sigma As Double) As Double
    If t <= 0 Or sigma <= 0 Then
        LognormalCDF = 0
        Exit Function
    End If

    Dim z As Double
    z = (Log(t) - mu) / sigma

    LognormalCDF = NormSDist(z)
End Function


' -----------------------------------------------------------------------------
' LogLogisticCDF
' Cumulative distribution function for the Log-Logistic distribution.
' F(t) = 1 / (1 + (t/alpha)^(-beta))
' Equivalently: F(t) = (t/alpha)^beta / (1 + (t/alpha)^beta)
' Parameters: alpha = scale, beta = shape
' Returns 0 for t <= 0.
' -----------------------------------------------------------------------------
Public Function LogLogisticCDF(t As Double, alpha As Double, beta As Double) As Double
    If t <= 0 Or alpha <= 0 Or beta <= 0 Then
        LogLogisticCDF = 0
        Exit Function
    End If

    Dim ratio As Double
    ratio = (t / alpha) ^ beta

    LogLogisticCDF = ratio / (1 + ratio)
End Function


' -----------------------------------------------------------------------------
' GammaCDF
' Cumulative distribution function for the Gamma distribution.
' Uses the regularized incomplete gamma function via series expansion.
' F(t) = P(alpha, t/theta) where P is the regularized lower incomplete gamma.
' Parameters: alpha = shape, theta = scale
' Returns 0 for t <= 0.
' -----------------------------------------------------------------------------
Public Function GammaCDF(t As Double, alpha As Double, theta As Double) As Double
    If t <= 0 Or alpha <= 0 Or theta <= 0 Then
        GammaCDF = 0
        Exit Function
    End If

    Dim x As Double
    x = t / theta

    GammaCDF = RegularizedGammaP(alpha, x)
End Function


' =============================================================================
' Dispatcher
' =============================================================================

' -----------------------------------------------------------------------------
' CalcCDF
' Routes to the appropriate CDF function by distribution name.
' Supports: "Weibull", "Lognormal", "LogLogistic", "Gamma"
' p1/p2/p3 map to the distribution-specific parameters.
' -----------------------------------------------------------------------------
Public Function CalcCDF(t As Double, distName As String, _
                        p1 As Double, p2 As Double, _
                        Optional p3 As Double = 0) As Double
    Select Case LCase(distName)
        Case "weibull"
            CalcCDF = WeibullCDF(t, p1, p2)
        Case "lognormal"
            CalcCDF = LognormalCDF(t, p1, p2)
        Case "loglogistic"
            CalcCDF = LogLogisticCDF(t, p1, p2)
        Case "gamma"
            CalcCDF = GammaCDF(t, p1, p2)
        Case Else
            KernelConfig.LogError SEV_WARN, "Ext_CurveLib", "W-850", _
                "Unknown distribution: " & distName & ". Returning 0.", _
                "MANUAL BYPASS: Call the specific CDF function directly " & _
                "(WeibullCDF, LognormalCDF, LogLogisticCDF, GammaCDF)."
            CalcCDF = 0
    End Select
End Function


' =============================================================================
' Interpolation
' =============================================================================

' -----------------------------------------------------------------------------
' LinInterp3
' Linear interpolation with 3 anchor points at TL=1, TL=50, TL=100.
' For TL 1-50: linear between v1 and v50.
' For TL 50-100: linear between v50 and v100.
' Clamps to v1 or v100 outside range.
' -----------------------------------------------------------------------------
Public Function LinInterp3(tl As Long, v1 As Double, v50 As Double, _
                           v100 As Double) As Double
    If tl <= 1 Then
        LinInterp3 = v1
    ElseIf tl <= 50 Then
        LinInterp3 = v1 + (v50 - v1) * (tl - 1) / 49
    ElseIf tl <= 100 Then
        LinInterp3 = v50 + (v100 - v50) * (tl - 50) / 50
    Else
        LinInterp3 = v100
    End If
End Function


' -----------------------------------------------------------------------------
' LinInterp5
' 5-point piecewise linear interpolation at TL=1,25,50,75,100.
' -----------------------------------------------------------------------------
Public Function LinInterp5(tl As Long, v1 As Double, v25 As Double, _
    v50 As Double, v75 As Double, v100 As Double) As Double
    If tl <= 1 Then
        LinInterp5 = v1
    ElseIf tl <= 25 Then
        LinInterp5 = v1 + (v25 - v1) * (tl - 1) / 24
    ElseIf tl <= 50 Then
        LinInterp5 = v25 + (v50 - v25) * (tl - 25) / 25
    ElseIf tl <= 75 Then
        LinInterp5 = v50 + (v75 - v50) * (tl - 50) / 25
    ElseIf tl <= 100 Then
        LinInterp5 = v75 + (v100 - v75) * (tl - 75) / 25
    Else
        LinInterp5 = v100
    End If
End Function


' -----------------------------------------------------------------------------
' LogInterp3
' Log-linear (geometric) interpolation for scale parameters.
' Uses logarithms of anchor values for interpolation.
' For TL 1-50: geometric between v1 and v50.
' For TL 50-100: geometric between v50 and v100.
' All values must be > 0.
' -----------------------------------------------------------------------------
Public Function LogInterp3(tl As Long, v1 As Double, v50 As Double, _
                           v100 As Double) As Double
    If v1 <= 0 Or v50 <= 0 Or v100 <= 0 Then
        LogInterp3 = LinInterp3(tl, v1, v50, v100)
        Exit Function
    End If

    Dim logV1 As Double
    logV1 = Log(v1)
    Dim logV50 As Double
    logV50 = Log(v50)
    Dim logV100 As Double
    logV100 = Log(v100)

    Dim logResult As Double
    If tl <= 1 Then
        logResult = logV1
    ElseIf tl <= 50 Then
        logResult = logV1 + (logV50 - logV1) * (tl - 1) / 49
    ElseIf tl <= 100 Then
        logResult = logV50 + (logV100 - logV50) * (tl - 50) / 50
    Else
        logResult = logV100
    End If

    LogInterp3 = Exp(logResult)
End Function


' -----------------------------------------------------------------------------
' LogInterp5
' 5-point log-linear (geometric) interpolation at TL=1,25,50,75,100.
' All values must be > 0.
' -----------------------------------------------------------------------------
Public Function LogInterp5(tl As Long, v1 As Double, v25 As Double, _
    v50 As Double, v75 As Double, v100 As Double) As Double
    If v1 <= 0 Or v25 <= 0 Or v50 <= 0 Or v75 <= 0 Or v100 <= 0 Then
        LogInterp5 = LinInterp5(tl, v1, v25, v50, v75, v100)
        Exit Function
    End If
    Dim lv1 As Double: lv1 = Log(v1)
    Dim lv25 As Double: lv25 = Log(v25)
    Dim lv50 As Double: lv50 = Log(v50)
    Dim lv75 As Double: lv75 = Log(v75)
    Dim lv100 As Double: lv100 = Log(v100)
    Dim lr As Double
    If tl <= 1 Then
        lr = lv1
    ElseIf tl <= 25 Then
        lr = lv1 + (lv25 - lv1) * (tl - 1) / 24
    ElseIf tl <= 50 Then
        lr = lv25 + (lv50 - lv25) * (tl - 25) / 25
    ElseIf tl <= 75 Then
        lr = lv50 + (lv75 - lv50) * (tl - 50) / 25
    ElseIf tl <= 100 Then
        lr = lv75 + (lv100 - lv75) * (tl - 75) / 25
    Else
        lr = lv100
    End If
    LogInterp5 = Exp(lr)
End Function


' =============================================================================
' Normalization
' =============================================================================

' -----------------------------------------------------------------------------
' NormalizeCDF
' Returns cdfAtAge / cdfAtMaxAge, capped at 1.0.
' If cdfAtMaxAge <= 0, returns 0.
' -----------------------------------------------------------------------------
Public Function NormalizeCDF(cdfAtAge As Double, cdfAtMaxAge As Double) As Double
    If cdfAtMaxAge <= 0 Then
        NormalizeCDF = 0
        Exit Function
    End If

    Dim result As Double
    result = cdfAtAge / cdfAtMaxAge

    If result > 1 Then result = 1
    If result < 0 Then result = 0

    NormalizeCDF = result
End Function


' =============================================================================
' Evaluation
' =============================================================================

' -----------------------------------------------------------------------------
' EvaluateCurve
' Full evaluation: CalcCDF(age) / CalcCDF(maxAge), capped at 1.0.
' Convenience wrapper combining CDF + normalization.
' -----------------------------------------------------------------------------
Public Function EvaluateCurve(distName As String, p1 As Double, p2 As Double, _
                              age As Double, maxAge As Long, _
                              Optional p3 As Double = 0) As Double
    ' DE-08: age is Double to support mid-period offsets (e.g. age-0.5).
    If maxAge <= 0 Then
        EvaluateCurve = 0
        Exit Function
    End If

    Dim cdfAge As Double
    cdfAge = CalcCDF(age, distName, p1, p2, p3)

    Dim cdfMax As Double
    cdfMax = CalcCDF(CDbl(maxAge), distName, p1, p2, p3)

    EvaluateCurve = NormalizeCDF(cdfAge, cdfMax)
End Function


' =============================================================================
' Batch Evaluation
' =============================================================================

' -----------------------------------------------------------------------------
' EvaluateCurveBatch
' Returns an array of normalized CDF values for each age.
' Performance optimization for generating full emergence tables.
' ages() is a 1-based Long array. Returns a matching Double array.
' -----------------------------------------------------------------------------
Public Function EvaluateCurveBatch(distName As String, p1 As Double, p2 As Double, _
                                   ages() As Long, maxAge As Long, _
                                   Optional p3 As Double = 0) As Double()
    Dim result() As Double
    Dim lb As Long
    Dim ub As Long
    lb = LBound(ages)
    ub = UBound(ages)
    ReDim result(lb To ub)

    If maxAge <= 0 Then
        EvaluateCurveBatch = result
        Exit Function
    End If

    ' Pre-compute CDF at maxAge once
    Dim cdfMax As Double
    cdfMax = CalcCDF(CDbl(maxAge), distName, p1, p2, p3)

    If cdfMax <= 0 Then
        EvaluateCurveBatch = result
        Exit Function
    End If

    Dim i As Long
    For i = lb To ub
        Dim cdfAge As Double
        cdfAge = CalcCDF(CDbl(ages(i)), distName, p1, p2, p3)
        result(i) = NormalizeCDF(cdfAge, cdfMax)
    Next i

    EvaluateCurveBatch = result
End Function


' =============================================================================
' Config-Driven Lookup
' =============================================================================

' -----------------------------------------------------------------------------
' LookupCurveByID
' Reads curve_library_config from Config sheet.
' Returns array: (distName, p1, p2, p3, maxAge)
' If curveID not found, logs WARN and returns default (Weibull, 12, 1.6, 0, 120).
' -----------------------------------------------------------------------------
Public Function LookupCurveByID(curveID As String) As Variant
    Dim params() As Variant
    Dim p1str As String
    Dim p2str As String
    Dim p3str As String
    Dim maStr As String
    Dim cnt As Long
    cnt = KernelConfig.GetCurveLibraryCount()

    Dim idx As Long
    For idx = 1 To cnt
        Dim cid As String
        cid = KernelConfig.GetCurveLibraryField(idx, CLCFG_COL_ID)
        If StrComp(cid, curveID, vbTextCompare) = 0 Then
            ReDim params(1 To 5)
            params(1) = KernelConfig.GetCurveLibraryField(idx, CLCFG_COL_DIST)
            p1str = KernelConfig.GetCurveLibraryField(idx, CLCFG_COL_P1)
            If IsNumeric(p1str) Then params(2) = CDbl(p1str) Else params(2) = 0
            p2str = KernelConfig.GetCurveLibraryField(idx, CLCFG_COL_P2)
            If IsNumeric(p2str) Then params(3) = CDbl(p2str) Else params(3) = 0
            p3str = KernelConfig.GetCurveLibraryField(idx, CLCFG_COL_P3)
            If IsNumeric(p3str) Then params(4) = CDbl(p3str) Else params(4) = 0
            maStr = KernelConfig.GetCurveLibraryField(idx, CLCFG_COL_MAXAGE)
            If IsNumeric(maStr) Then params(5) = CLng(maStr) Else params(5) = 120
            LookupCurveByID = params
            Exit Function
        End If
    Next idx

    ' Not found -- return defaults
    KernelConfig.LogError SEV_WARN, "Ext_CurveLib", "W-851", _
        "CurveID not found: " & curveID & ". Using default (Weibull, 12, 1.6, 0, 120).", _
        "MANUAL BYPASS: Call EvaluateCurve directly with explicit parameters " & _
        "instead of looking up from config."

    ReDim params(1 To 5)
    params(1) = "Weibull"
    params(2) = 12#
    params(3) = 1.6
    params(4) = 0#
    params(5) = 120
    LookupCurveByID = params
End Function


' -----------------------------------------------------------------------------
' Rnd12
' Round to nearest multiple of 12, minimum 12.
' -----------------------------------------------------------------------------
Public Function Rnd12(val As Double) As Long
    Dim result As Long
    result = CLng(Int(val / 12 + 0.5)) * 12
    If result < 12 Then result = 12
    Rnd12 = result
End Function


' -----------------------------------------------------------------------------
' PropMaxAge
' 3-anchor linear MaxAge interpolation for Property curves (Prop3 method).
' Anchors at TL=1, TL=50, TL=100. Result rounded to nearest 12.
' -----------------------------------------------------------------------------
Public Function PropMaxAge(tl As Long, v1 As Double, v50 As Double, v100 As Double) As Long
    Dim raw As Double
    If tl <= 1 Then
        raw = v1
    ElseIf tl <= 50 Then
        raw = v1 + (CDbl(tl) - 1#) / 49# * (v50 - v1)
    ElseIf tl >= 100 Then
        raw = v100
    Else
        raw = v50 + (CDbl(tl) - 50#) / 50# * (v100 - v50)
    End If
    PropMaxAge = Rnd12(raw)
End Function


' -----------------------------------------------------------------------------
' CasMaxAge
' 4-anchor piecewise linear MaxAge interpolation for Casualty curves (Cas4).
' Anchors at TL=1, TL=80, TL=90, TL=100. Result rounded to nearest 12.
' -----------------------------------------------------------------------------
Public Function CasMaxAge(tl As Long, v1 As Double, v80 As Double, _
    v90 As Double, v100 As Double) As Long
    Dim raw As Double
    If tl <= 1 Then
        raw = v1
    ElseIf tl <= 80 Then
        raw = v1 + (CDbl(tl) - 1#) / 79# * (v80 - v1)
    ElseIf tl <= 90 Then
        raw = v80 + (CDbl(tl) - 80#) / 10# * (v90 - v80)
    ElseIf tl >= 100 Then
        raw = v100
    Else
        raw = v90 + (CDbl(tl) - 90#) / 10# * (v100 - v90)
    End If
    CasMaxAge = Rnd12(raw)
End Function


' -----------------------------------------------------------------------------
' GetCurveParamsByTL
' 5-point anchor lookup: LOB x CurveType x TrendLevel -> (dist, p1, p2, p3, maxAge).
' Reads curve_library_config from Config sheet. Interpolates P1, P2, MaxAge
' using Log or Linear method across 5 anchors (TL=1,25,50,75,100).
' Returns 5-element Variant array: (distName, p1, p2, p3, maxAge).
' Falls back to Property|Paid defaults if LOB/CurveType not found.
' -----------------------------------------------------------------------------
Public Function GetCurveParamsByTL(lob As String, curveType As String, _
    trendLevel As Long) As Variant
    Dim params() As Variant
    ReDim params(1 To 5)
    Dim cnt As Long
    cnt = KernelConfig.GetCurveLibraryCount()

    Dim idx As Long
    For idx = 1 To cnt
        Dim cfgLOB As String
        cfgLOB = KernelConfig.GetCurveLibraryField(idx, CLA_COL_LOB)
        Dim cfgType As String
        cfgType = KernelConfig.GetCurveLibraryField(idx, CLA_COL_TYPE)

        If StrComp(cfgLOB, lob, vbTextCompare) = 0 And _
           StrComp(cfgType, curveType, vbTextCompare) = 0 Then
            ' Distribution name
            params(1) = KernelConfig.GetCurveLibraryField(idx, CLA_COL_DIST)

            ' P1: 5-point interpolation (Log or Lin)
            Dim p1Method As String
            p1Method = KernelConfig.GetCurveLibraryField(idx, CLA_COL_P1METHOD)
            Dim p1s(1 To 5) As String
            Dim p1v(1 To 5) As Double
            p1s(1) = KernelConfig.GetCurveLibraryField(idx, CLA_COL_P1_TL1)
            p1s(2) = KernelConfig.GetCurveLibraryField(idx, CLA_COL_P1_TL25)
            p1s(3) = KernelConfig.GetCurveLibraryField(idx, CLA_COL_P1_TL50)
            p1s(4) = KernelConfig.GetCurveLibraryField(idx, CLA_COL_P1_TL75)
            p1s(5) = KernelConfig.GetCurveLibraryField(idx, CLA_COL_P1_TL100)
            Dim pi As Long
            For pi = 1 To 5
                If IsNumeric(p1s(pi)) Then p1v(pi) = CDbl(p1s(pi)) Else p1v(pi) = 12
            Next pi
            If StrComp(p1Method, "Log", vbTextCompare) = 0 Then
                params(2) = LogInterp5(trendLevel, p1v(1), p1v(2), p1v(3), p1v(4), p1v(5))
            Else
                params(2) = LinInterp5(trendLevel, p1v(1), p1v(2), p1v(3), p1v(4), p1v(5))
            End If

            ' P2: 5-point linear interpolation (per-TL shape parameter)
            Dim p2s(1 To 5) As String
            Dim p2v(1 To 5) As Double
            p2s(1) = KernelConfig.GetCurveLibraryField(idx, CLA_COL_P2_TL1)
            p2s(2) = KernelConfig.GetCurveLibraryField(idx, CLA_COL_P2_TL25)
            p2s(3) = KernelConfig.GetCurveLibraryField(idx, CLA_COL_P2_TL50)
            p2s(4) = KernelConfig.GetCurveLibraryField(idx, CLA_COL_P2_TL75)
            p2s(5) = KernelConfig.GetCurveLibraryField(idx, CLA_COL_P2_TL100)
            For pi = 1 To 5
                If IsNumeric(p2s(pi)) Then p2v(pi) = CDbl(p2s(pi)) Else p2v(pi) = 1
            Next pi
            params(3) = LinInterp5(trendLevel, p2v(1), p2v(2), p2v(3), p2v(4), p2v(5))

            ' P3 (fixed, not interpolated)
            Dim sp3 As String: sp3 = KernelConfig.GetCurveLibraryField(idx, CLA_COL_P3)
            If IsNumeric(sp3) Then params(4) = CDbl(sp3) Else params(4) = 0

            ' MaxAge: 5-point linear interpolation, rounded to nearest 12
            Dim mas(1 To 5) As String
            Dim mav(1 To 5) As Double
            mas(1) = KernelConfig.GetCurveLibraryField(idx, CLA_COL_MA_TL1)
            mas(2) = KernelConfig.GetCurveLibraryField(idx, CLA_COL_MA_TL25)
            mas(3) = KernelConfig.GetCurveLibraryField(idx, CLA_COL_MA_TL50)
            mas(4) = KernelConfig.GetCurveLibraryField(idx, CLA_COL_MA_TL75)
            mas(5) = KernelConfig.GetCurveLibraryField(idx, CLA_COL_MA_TL100)
            For pi = 1 To 5
                If IsNumeric(mas(pi)) Then mav(pi) = CDbl(mas(pi)) Else mav(pi) = 120
            Next pi
            params(5) = Rnd12(LinInterp5(trendLevel, mav(1), mav(2), mav(3), mav(4), mav(5)))

            GetCurveParamsByTL = params
            Exit Function
        End If
    Next idx

    ' Not found -- return Property|Paid defaults at TL=50
    KernelConfig.LogError SEV_WARN, "Ext_CurveLib", "W-852", _
        "Anchor config not found: " & lob & "|" & curveType & _
        ". Using default (Weibull, 5.2, 0.8, 0, 60).", _
        "MANUAL BYPASS: Call EvaluateCurve directly with explicit parameters."
    params(1) = "Weibull"
    params(2) = 5.2
    params(3) = 0.8
    params(4) = 0#
    params(5) = 60
    GetCurveParamsByTL = params
End Function


' =============================================================================
' UDF Wrappers (callable from Excel cells)
' =============================================================================

' -----------------------------------------------------------------------------
' CurveRefPct
' Thin wrapper: resolves curve params by TL, evaluates CDF, returns % of ultimate.
' Callable from Excel cells: =CurveRefPct("Property","Paid",50,12)
' -----------------------------------------------------------------------------
Public Function CurveRefPct(lob As String, curveType As String, _
                            trendLevel As Long, age As Double) As Double
    ' DE-08: age is Double to support mid-period offsets (e.g. 0.5, 1.5).
    If trendLevel < 1 Or trendLevel > 100 Or age < 0.01 Then
        CurveRefPct = 0
        Exit Function
    End If

    Dim params As Variant
    params = GetCurveParamsByTL(lob, curveType, trendLevel)

    Dim distName As String: distName = CStr(params(1))
    Dim p1 As Double: p1 = CDbl(params(2))
    Dim p2 As Double: p2 = CDbl(params(3))
    Dim p3 As Double: p3 = CDbl(params(4))
    Dim maxAge As Long: maxAge = CLng(params(5))

    If age > maxAge Then
        CurveRefPct = 1
        Exit Function
    End If

    CurveRefPct = EvaluateCurve(distName, p1, p2, age, maxAge, p3)
End Function


' -----------------------------------------------------------------------------
' CurveRefMaxAge
' Returns the MaxAge for a given LOB/CurveType/TL combination.
' Callable from Excel cells: =CurveRefMaxAge("Property","Paid",50)
' -----------------------------------------------------------------------------
Public Function CurveRefMaxAge(lob As String, curveType As String, _
                               trendLevel As Long) As Long
    If trendLevel < 1 Or trendLevel > 100 Then
        CurveRefMaxAge = 0
        Exit Function
    End If
    Dim params As Variant
    params = GetCurveParamsByTL(lob, curveType, trendLevel)
    CurveRefMaxAge = CLng(params(5))
End Function


' =============================================================================
' Curve Reference Tab Builder
' =============================================================================

' BuildCurveReferenceTab
' Populates the "Curve Reference" tab with 8 blocks of development curve data.
' Each block shows cumulative % of ultimate at 10 trend level increments.
' Called from KernelBootstrap after tab creation (insurance config only).
Public Sub BuildCurveReferenceTab()
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Curve Reference")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0
    ws.Cells.ClearContents

    ' --- Constants ---
    Dim TLs(1 To 10) As Long
    TLs(1) = 10: TLs(2) = 20: TLs(3) = 30: TLs(4) = 40: TLs(5) = 50
    TLs(6) = 60: TLs(7) = 70: TLs(8) = 80: TLs(9) = 90: TLs(10) = 100

    Dim ages(1 To 21) As Long
    ages(1) = 1: ages(2) = 3: ages(3) = 6: ages(4) = 9: ages(5) = 12
    ages(6) = 18: ages(7) = 24: ages(8) = 30: ages(9) = 36: ages(10) = 42
    ages(11) = 48: ages(12) = 54: ages(13) = 60: ages(14) = 72: ages(15) = 84
    ages(16) = 96: ages(17) = 108: ages(18) = 120: ages(19) = 144
    ages(20) = 180: ages(21) = 240

    Dim NUM_AGES As Long: NUM_AGES = 21
    Dim NUM_TLS As Long: NUM_TLS = 10
    Dim DATA_COL As Long: DATA_COL = 3

    ' LOB/CurveType pairs (8 blocks)
    Dim lobArr(1 To 8) As String
    Dim typeArr(1 To 8) As String
    Dim labelArr(1 To 8) As String
    lobArr(1) = "Property": typeArr(1) = "Paid": labelArr(1) = "PROPERTY -- Paid Loss Development"
    lobArr(2) = "Property": typeArr(2) = "Case Incurred": labelArr(2) = "PROPERTY -- Case Incurred Development"
    lobArr(3) = "Property": typeArr(3) = "Reported Count": labelArr(3) = "PROPERTY -- Reported Count Development"
    lobArr(4) = "Property": typeArr(4) = "Closed Count": labelArr(4) = "PROPERTY -- Closed Count Development"
    lobArr(5) = "Casualty": typeArr(5) = "Paid": labelArr(5) = "CASUALTY -- Paid Loss Development"
    lobArr(6) = "Casualty": typeArr(6) = "Case Incurred": labelArr(6) = "CASUALTY -- Case Incurred Development"
    lobArr(7) = "Casualty": typeArr(7) = "Reported Count": labelArr(7) = "CASUALTY -- Reported Count Development"
    lobArr(8) = "Casualty": typeArr(8) = "Closed Count": labelArr(8) = "CASUALTY -- Closed Count Development"

    ' --- Tab header ---
    ws.Cells(1, 2).Value = "Curve Reference"
    ws.Range(ws.Cells(1, 2), ws.Cells(1, DATA_COL + NUM_TLS - 1)).Font.Bold = True
    ws.Range(ws.Cells(1, 2), ws.Cells(1, DATA_COL + NUM_TLS - 1)).Interior.Color = RGB(31, 56, 100)
    ws.Range(ws.Cells(1, 2), ws.Cells(1, DATA_COL + NUM_TLS - 1)).Font.Color = RGB(255, 255, 255)
    ws.Cells(2, 2).Value = "Development Patterns -- Cumulative % of Ultimate by Trend Level"
    ws.Cells(2, 2).Font.Italic = True
    ws.Cells(2, 2).Font.Color = RGB(128, 128, 128)

    Dim curRow As Long: curRow = 4
    Dim blk As Long
    Dim t As Long
    Dim a As Long
    Dim formulaStr As String
    Dim lobQ As String
    Dim typeQ As String

    For blk = 1 To 8
        lobQ = lobArr(blk)
        typeQ = typeArr(blk)

        ' Section header
        ws.Cells(curRow, 2).Value = labelArr(blk)
        ws.Cells(curRow, 2).Font.Bold = True
        ws.Range(ws.Cells(curRow, 2), ws.Cells(curRow, DATA_COL + NUM_TLS - 1)).Interior.Color = RGB(217, 225, 242)

        ' Column headers
        curRow = curRow + 1
        ws.Cells(curRow, 2).Value = "Dev Age (months)"
        ws.Cells(curRow, 2).Font.Bold = True
        For t = 1 To NUM_TLS
            ws.Cells(curRow, DATA_COL + t - 1).Value = "TL=" & TLs(t)
            ws.Cells(curRow, DATA_COL + t - 1).Font.Bold = True
            ws.Cells(curRow, DATA_COL + t - 1).HorizontalAlignment = xlCenter
        Next t

        ' Data rows
        For a = 1 To NUM_AGES
            curRow = curRow + 1
            ws.Cells(curRow, 2).Value = ages(a)
            ws.Cells(curRow, 2).HorizontalAlignment = xlRight

            For t = 1 To NUM_TLS
                formulaStr = "=CurveRefPct(""" & lobQ & """,""" & typeQ & """," & _
                    TLs(t) & "," & ages(a) & ")"
                ws.Cells(curRow, DATA_COL + t - 1).formula = formulaStr
                ws.Cells(curRow, DATA_COL + t - 1).NumberFormat = "0.0%"
            Next t

            ' Alternate row shading
            If a Mod 2 = 0 Then
                ws.Range(ws.Cells(curRow, 2), ws.Cells(curRow, DATA_COL + NUM_TLS - 1)).Interior.Color = RGB(242, 242, 242)
            End If
        Next a

        ' MaxAge row
        curRow = curRow + 1
        ws.Cells(curRow, 2).Value = "Max Age (months)"
        ws.Cells(curRow, 2).Font.Italic = True
        ws.Cells(curRow, 2).Font.Color = RGB(128, 128, 128)
        For t = 1 To NUM_TLS
            formulaStr = "=CurveRefMaxAge(""" & lobQ & """,""" & typeQ & """," & TLs(t) & ")"
            ws.Cells(curRow, DATA_COL + t - 1).formula = formulaStr
            ws.Cells(curRow, DATA_COL + t - 1).NumberFormat = "#,##0"
            ws.Cells(curRow, DATA_COL + t - 1).Font.Italic = True
            ws.Cells(curRow, DATA_COL + t - 1).Font.Color = RGB(128, 128, 128)
        Next t

        ' Spacer + advance to next block
        curRow = curRow + 2
    Next blk

    ' Column widths
    ws.Columns(2).ColumnWidth = 16
    For t = 1 To NUM_TLS
        ws.Columns(DATA_COL + t - 1).ColumnWidth = 10
    Next t

    ' Hide column A (consistent with other tabs)
    ws.Columns(1).Hidden = True
End Sub


' =============================================================================
' Extension Entry Point (required by contract, but no-op for Standalone)
' =============================================================================
Public Sub CurveLib_Execute()
    KernelConfig.LogError SEV_INFO, "Ext_CurveLib", "I-850", _
        "CurveLib is a Standalone extension. Call its functions directly.", ""
End Sub


' =============================================================================
' Helper Functions (Private)
' =============================================================================

' -----------------------------------------------------------------------------
' NormSDist
' Standard normal CDF approximation (Abramowitz and Stegun 26.2.17).
' Accurate to ~7.5e-8.
' -----------------------------------------------------------------------------
Private Function NormSDist(z As Double) As Double
    Dim absZ As Double
    absZ = Abs(z)

    If absZ > 37 Then
        If z > 0 Then
            NormSDist = 1
        Else
            NormSDist = 0
        End If
        Exit Function
    End If

    Dim b1 As Double
    b1 = 0.319381530#
    Dim b2 As Double
    b2 = -0.356563782#
    Dim b3 As Double
    b3 = 1.781477937#
    Dim b4 As Double
    b4 = -1.821255978#
    Dim b5 As Double
    b5 = 1.330274429#
    Dim pp As Double
    pp = 0.2316419#

    Dim tVal As Double
    tVal = 1 / (1 + pp * absZ)

    Dim poly As Double
    poly = ((((b5 * tVal + b4) * tVal + b3) * tVal + b2) * tVal + b1) * tVal

    Dim pdf As Double
    pdf = Exp(-0.5 * absZ * absZ) / 2.506628274631#

    Dim cdf As Double
    cdf = 1 - pdf * poly

    If z < 0 Then
        NormSDist = 1 - cdf
    Else
        NormSDist = cdf
    End If
End Function


' -----------------------------------------------------------------------------
' RegularizedGammaP
' Regularized lower incomplete gamma function P(a, x) via series expansion.
' P(a, x) = sum_{n=0}^{inf} (-1)^n * x^n / (n! * (a + n))  *  x^a * exp(-x) / Gamma(a)
' Uses the series representation for x < a+1, continued fraction otherwise.
' For RDK purposes, the series expansion is sufficient.
' -----------------------------------------------------------------------------
Private Function RegularizedGammaP(a As Double, x As Double) As Double
    If x <= 0 Then
        RegularizedGammaP = 0
        Exit Function
    End If

    If x < a + 1 Then
        ' Series expansion
        RegularizedGammaP = GammaPSeries(a, x)
    Else
        ' Continued fraction (complement)
        RegularizedGammaP = 1 - GammaQCF(a, x)
    End If
End Function


' -----------------------------------------------------------------------------
' GammaPSeries
' Series expansion for regularized lower incomplete gamma P(a, x).
' P(a,x) = exp(-x) * x^a * sum_{n=0}^{inf} x^n / (a*(a+1)*...*(a+n))  / Gamma(a)
' Simplified: uses log-gamma for numerical stability.
' -----------------------------------------------------------------------------
Private Function GammaPSeries(a As Double, x As Double) As Double
    Dim maxIter As Long
    maxIter = 200
    Dim eps As Double
    eps = 0.0000000001#

    Dim ap As Double
    ap = a
    Dim delVal As Double
    delVal = 1 / a
    Dim sumVal As Double
    sumVal = delVal

    Dim n As Long
    For n = 1 To maxIter
        ap = ap + 1
        delVal = delVal * x / ap
        sumVal = sumVal + delVal
        If Abs(delVal) < Abs(sumVal) * eps Then Exit For
    Next n

    Dim logGammaA As Double
    logGammaA = LogGamma(a)

    GammaPSeries = sumVal * Exp(-x + a * Log(x) - logGammaA)
End Function


' -----------------------------------------------------------------------------
' GammaQCF
' Continued fraction for upper incomplete gamma Q(a, x) = 1 - P(a, x).
' Uses modified Lentz's method.
' -----------------------------------------------------------------------------
Private Function GammaQCF(a As Double, x As Double) As Double
    Dim maxIter As Long
    maxIter = 200
    Dim eps As Double
    eps = 0.0000000001#
    Dim tiny As Double
    tiny = 0.00000000000000000000000000001#

    Dim b As Double
    b = x + 1 - a
    Dim cVal As Double
    cVal = 1 / tiny
    Dim d As Double
    d = 1 / b
    Dim h As Double
    h = d

    Dim i As Long
    For i = 1 To maxIter
        Dim an As Double
        an = -i * (i - a)
        b = b + 2
        d = an * d + b
        If Abs(d) < tiny Then d = tiny
        cVal = b + an / cVal
        If Abs(cVal) < tiny Then cVal = tiny
        d = 1 / d
        Dim delV As Double
        delV = d * cVal
        h = h * delV
        If Abs(delV - 1) < eps Then Exit For
    Next i

    Dim logGammaA As Double
    logGammaA = LogGamma(a)

    GammaQCF = Exp(-x + a * Log(x) - logGammaA) * h
End Function


' -----------------------------------------------------------------------------
' LogGamma
' Stirling's approximation for ln(Gamma(x)).
' Uses Lanczos approximation for better accuracy.
' -----------------------------------------------------------------------------
Private Function LogGamma(x As Double) As Double
    Dim coef(0 To 5) As Double
    coef(0) = 76.18009172947146#
    coef(1) = -86.50532032941677#
    coef(2) = 24.01409824083091#
    coef(3) = -1.231739572450155#
    coef(4) = 0.001208650973866179#
    coef(5) = -0.000005395239384953#

    Dim y As Double
    y = x
    Dim tmp As Double
    tmp = x + 5.5
    tmp = tmp - (x + 0.5) * Log(tmp)

    Dim ser As Double
    ser = 1.000000000190015#

    Dim j As Long
    For j = 0 To 5
        y = y + 1
        ser = ser + coef(j) / y
    Next j

    LogGamma = -tmp + Log(2.506628274631# * ser / x)
End Function
