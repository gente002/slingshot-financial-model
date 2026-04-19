Attribute VB_Name = "InsuranceDomainEngine"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' InsuranceDomainEngine.bas
' Purpose: Insurance actuarial computation engine. Implements the 4-function
'          domain contract (AP-43). Reads UW Inputs directly (CR-08).
'          Computes monthly Program x Period rows with 52 columns.
'          Ported from UWMEngine v2.9.4 computation pipeline.
' Phase 11A. ASCII only (AP-06). All column refs via ColIndex (AP-08).

' --- UW Inputs layout constants (must match formula_tab_config) ---
Private Const MAX_PROGRAMS As Long = 10
Private Const MAX_HORIZON As Long = 120
Private Const MAX_DEV_ENDPOINT As Long = 480
Private Const NUM_LAYERS As Long = 3
Private Const UWIN_S1_DATA_ROW As Long = 6
Private Const UWIN_COL_BU As Long = 2
Private Const UWIN_COL_NAME As Long = 3
Private Const UWIN_COL_TERM As Long = 4
Private Const UWIN_GWP_START_COL As Long = 5
Private Const UWIN_GWP_QUARTERS As Long = 20
Private Const UWIN_GROWTH_COL As Long = 25
Private Const UWIN_S2_DATA_ROW As Long = 19
Private Const UWIN_COMM_START_COL As Long = 5
Private Const UWIN_S3_DATA_ROW As Long = 32
Private Const UWIN_LOSS_COL_PATTERN As Long = 4
Private Const UWIN_LOSS_COL_LOSSDEV As Long = 5
Private Const UWIN_LOSS_COL_CNTDEV As Long = 6
Private Const UWIN_LOSS_COL_ELR1 As Long = 7
Private Const UWIN_LOSS_COL_SEV As Long = 11
Private Const UWIN_LOSS_COL_FREQ1 As Long = 12
Private Const UWIN_S4_DATA_ROW As Long = 65
Private Const UWIN_REINS_START_COL As Long = 5
Private Const UWIN_REINS_COLS_PER_YEAR As Long = 3
Private Const UWIN_S5_DATA_ROW As Long = 78
Private Const UWIN_XOL_CAT_START_COL As Long = 5
Private Const UWIN_XOL_OTHER_START_COL As Long = 10

' --- Curve development endpoint threshold ---
Private Const DEV_ENDPOINT_PCT As Double = 0.999999

' --- Module-level state ---
Private m_initialized As Boolean
Public m_numProgs As Long
Public m_horizon As Long
Public GranularCSVPath As String

' Program identity (Section 1)
Public m_progName(1 To 10) As String
Public m_progBU(1 To 10) As String
Public m_progTerm(1 To 10) As Long

' Premium schedule (Section 1) - 20 quarterly values
Private m_gwpQtr(1 To 10, 1 To 20) As Double
Private m_gwpGrowth(1 To 10) As Double

' Commission rates (Section 2) - 5 annual rates
Public m_commRate(1 To 10, 1 To 5) As Double

' Loss assumptions (Section 3) - 3 loss types per program
Private m_lyrLOB(1 To 10, 1 To 3) As String
Private m_lyrLossTL(1 To 10, 1 To 3) As Long
Private m_lyrCntTL(1 To 10, 1 To 3) As Long
Private m_lyrELR(1 To 10, 1 To 3, 1 To 4) As Double
Private m_lyrSev(1 To 10, 1 To 3) As Double
Private m_lyrFreq(1 To 10, 1 To 3, 1 To 4) As Double
Public m_lyrActive(1 To 10, 1 To 3) As Boolean

' Reinsurance terms (Section 4) - QS for Attr + Seas only, 5-year blocks
Public m_reinsCedePct(1 To 10, 1 To 5) As Double
Private m_reinsCedeComm(1 To 10, 1 To 5) As Double
Public m_reinsFrontFee(1 To 10, 1 To 5) As Double

' XOL reinsurance spend (Section 5) - expense only, no ceded losses
Private m_xolCat(1 To 10, 1 To 5) As Double
Private m_xolOther(1 To 10, 1 To 5) As Double

' Computation arrays -- sized to MAX_DEV_ENDPOINT for run-off beyond TimeHorizon
Public m_wpMon(1 To 10, 1 To 480) As Double
Public m_epMon(1 To 10, 1 To 480) As Double
Public m_ultMon(1 To 10, 1 To 3, 1 To 120) As Double
Public m_cntUlt(1 To 10, 1 To 3, 1 To 120) As Double
' BUG-117: m_ultMon is now EP-based. No separate EP arrays needed.
Private m_cumPaid(1 To 10, 1 To 480) As Double
Private m_cumCI(1 To 10, 1 To 480) As Double
Private m_cumUlt(1 To 10, 1 To 480) As Double
Private m_cumRpt(1 To 10, 1 To 480) As Double
Private m_cumCls(1 To 10, 1 To 480) As Double
Private m_cumCntUlt(1 To 10, 1 To 480) As Double
Private m_cumEP(1 To 10, 1 To 480) As Double

' QS-subject cumulatives (Attr+Seas only, CAT excluded)
Private m_qsPaid(1 To 10, 1 To 480) As Double
Private m_qsCI(1 To 10, 1 To 480) As Double
Private m_qsUlt(1 To 10, 1 To 480) As Double
Private m_qsRpt(1 To 10, 1 To 480) As Double
Private m_qsCls(1 To 10, 1 To 480) As Double
Private m_qsCntUlt(1 To 10, 1 To 480) As Double

' Curve parameter cache per program per layer (4 curves each)
Public Type CurveParams
    distPd As String
    p1Pd As Double
    p2Pd As Double
    maxAgePd As Long
    distCI As String
    p1CI As Double
    p2CI As Double
    maxAgeCI As Long
    distRC As String
    p1RC As Double
    p2RC As Double
    maxAgeRC As Long
    distCC As String
    p1CC As Double
    p2CC As Double
    maxAgeCC As Long
End Type

Public m_curves(1 To 10, 1 To 3) As CurveParams
Public m_devEnd(1 To 10) As Long


' Initialize
' Called once at bootstrap. Reads UW Inputs, loads curves, registers transforms.
Public Sub Initialize()
    ' BUG-100: Set m_horizon BEFORE LoadCurveParams so devEnd floor
    ' check (BUG-099 line 653) works. Previously m_horizon was 0 here,
    ' causing Property programs with short curves to lack Y5 Detail rows.
    m_horizon = KernelConfig.GetTimeHorizon()
    If m_horizon <= 0 Then m_horizon = 60
    If m_horizon > MAX_HORIZON Then m_horizon = MAX_HORIZON
    ReadUWInputs
    LoadCurveParams
    ApplyUWInputsFormatting
    ' Entity names stay on UW Inputs where ReadUWInputs reads them.
    ' Kernel's GetEntityName() reads directly from the input tab.
    ' (BUG-062 workaround removed -- no longer needed after Assumptions cleanup)
    ' Register AggregateToQuarterly as PostCompute transform
    KernelTransform.RegisterTransform "QuarterlyAgg", _
        "Ins_QuarterlyAgg", "AggregateToQuarterly", 100
    ' Register triangle builders (run after quarterly agg)
    KernelTransform.RegisterTransform "LossTriangles", _
        "Ins_Triangles", "BuildTriangleTab", 110
    KernelTransform.RegisterTransform "CountTriangles", _
        "Ins_Triangles", "BuildCountTriangleTab", 111
    m_initialized = True
End Sub



' GetRowCount
' Returns total output rows needed. Sum of devEnd(p)+1 across all programs.
' The +1 per program is the tail closure row that forces ITD reserves to 0.
' Called by kernel after Initialize to size the outputs array.
Public Function GetRowCount() As Long
    Dim total As Long
    Dim p As Long
    For p = 1 To m_numProgs
        total = total + m_devEnd(p) + 1
    Next p
    GetRowCount = total
End Function


' GetMaxPeriod
' Returns the maximum calendar period across all programs (= max devEnd + 1).
' The +1 accounts for the tail closure row.
' Used by kernel for Summary tab column count.
Public Function GetMaxPeriod() As Long
    Dim mx As Long
    Dim p As Long
    For p = 1 To m_numProgs
        If m_devEnd(p) + 1 > mx Then mx = m_devEnd(p) + 1
    Next p
    GetMaxPeriod = mx
End Function


' Validate
' Pre-run validation of UW Inputs. Returns False to halt pipeline.
Public Function Validate() As Boolean
    Validate = True

    If m_numProgs = 0 Then
        KernelConfig.LogError SEV_ERROR, "InsuranceDomainEngine", "E-300", _
            "No programs defined on UW Inputs tab", _
            "MANUAL BYPASS: Enter at least one program name in column C of the UW Inputs tab, row 6."
        Validate = False
        Exit Function
    End If

    Dim p As Long
    For p = 1 To m_numProgs
        ' Term must be positive
        If m_progTerm(p) <= 0 Then
            KernelConfig.LogError SEV_ERROR, "InsuranceDomainEngine", "E-301", _
                "Policy term <= 0 for program " & p & " (" & m_progName(p) & ")", _
                "MANUAL BYPASS: Set Term(mo) > 0 on UW Inputs Section 1, row " & (UWIN_S1_DATA_ROW + p - 1) & "."
            Validate = False
        End If

        ' Missing BU warning
        If Len(m_progBU(p)) = 0 Then
            KernelConfig.LogError SEV_WARN, "InsuranceDomainEngine", "W-300", _
                "BU blank for program " & p & " (" & m_progName(p) & "). Defaulting to Unassigned.", ""
            m_progBU(p) = "Unassigned"
        End If

        ' No premium warning
        Dim hasPrem As Boolean
        hasPrem = False
        Dim qi As Long
        For qi = 1 To UWIN_GWP_QUARTERS
            If m_gwpQtr(p, qi) <> 0 Then
                hasPrem = True
                Exit For
            End If
        Next qi
        If Not hasPrem Then
            KernelConfig.LogError SEV_WARN, "InsuranceDomainEngine", "W-301", _
                "No premium entered for program " & p & " (" & m_progName(p) & "). Program produces no output.", ""
        End If

        ' Commission rate validation
        Dim yr As Long
        For yr = 1 To 5
            If m_commRate(p, yr) < 0 Or m_commRate(p, yr) > 1 Then
                KernelConfig.LogError SEV_WARN, "InsuranceDomainEngine", "W-302", _
                    "Commission rate out of range [0,1] for program " & p & " year " & yr, ""
            End If
        Next yr

        ' Reinsurance validation
        For yr = 1 To 5
            If m_reinsCedePct(p, yr) < 0 Or m_reinsCedePct(p, yr) > 1 Then
                KernelConfig.LogError SEV_ERROR, "InsuranceDomainEngine", "E-302", _
                    "Cede % out of range [0,1] for program " & p & " year " & yr, _
                    "MANUAL BYPASS: Set Cede% between 0 and 1 on UW Inputs Section 4."
                Validate = False
            End If
            If m_reinsCedeComm(p, yr) < 0 Or m_reinsCedeComm(p, yr) > 1 Then
                KernelConfig.LogError SEV_WARN, "InsuranceDomainEngine", "W-303", _
                    "Ceding commission out of range [0,1] for program " & p & " year " & yr, ""
            End If
            If m_reinsFrontFee(p, yr) < 0 Or m_reinsFrontFee(p, yr) > 1 Then
                KernelConfig.LogError SEV_WARN, "InsuranceDomainEngine", "W-304", _
                    "Fronting fee out of range [0,1] for program " & p & " year " & yr, ""
            End If
        Next yr

        ' Loss assumption validation
        Dim lyr As Long
        For lyr = 1 To NUM_LAYERS
            If Not m_lyrActive(p, lyr) Then GoTo NextLayer

            Dim elrQ As Long
            For elrQ = 1 To 4
                If m_lyrELR(p, lyr, elrQ) < 0 Or m_lyrELR(p, lyr, elrQ) > 2 Then
                    KernelConfig.LogError SEV_WARN, "InsuranceDomainEngine", "W-305", _
                        "ELR out of range [0,2] for program " & p & " layer " & lyr & " Q" & elrQ & _
                        ". Value=" & Format(m_lyrELR(p, lyr, elrQ), "0.000"), ""
                End If
            Next elrQ
NextLayer:
        Next lyr
    Next p
End Function


' Reset
' Clears all module-level computation arrays.
Public Sub Reset()
    Dim p As Long
    Dim m As Long
    Dim lyr As Long

    For p = 1 To MAX_PROGRAMS
        ' Clear premium/cumulative arrays up to full dev endpoint
        For m = 1 To MAX_DEV_ENDPOINT
            m_wpMon(p, m) = 0
            m_epMon(p, m) = 0
            m_cumPaid(p, m) = 0
            m_cumCI(p, m) = 0
            m_cumUlt(p, m) = 0
            m_cumRpt(p, m) = 0
            m_cumCls(p, m) = 0
            m_cumCntUlt(p, m) = 0
            m_cumEP(p, m) = 0
            m_cumUlt(p, m) = 0
            m_cumCntUlt(p, m) = 0
            m_qsUlt(p, m) = 0
            m_qsCntUlt(p, m) = 0
            m_qsPaid(p, m) = 0
            m_qsCI(p, m) = 0
            m_qsUlt(p, m) = 0
            m_qsRpt(p, m) = 0
            m_qsCls(p, m) = 0
            m_qsCntUlt(p, m) = 0
        Next m
        ' Clear per-layer ultimates (only within exposure horizon)
        For m = 1 To MAX_HORIZON
            For lyr = 1 To NUM_LAYERS
                m_ultMon(p, lyr, m) = 0
                m_cntUlt(p, lyr, m) = 0
            Next lyr
        Next m
        ' m_devEnd NOT cleared: set by LoadCurveParams in Initialize (BUG-059)
    Next p
End Sub


' Execute
' 6-step computation pipeline. Writes to DomainOutputs (AP-43 contract).
Public Sub Execute()
    Dim outputs As Variant
    outputs = KernelEngine.DomainOutputs

    m_horizon = KernelConfig.GetTimeHorizon()
    If m_horizon <= 0 Then m_horizon = 60
    If m_horizon > MAX_HORIZON Then m_horizon = MAX_HORIZON

    ' 6-step pipeline
    SpreadPremium
    EarnPremium
    ComputeUltimates
    DevelopLosses
    Ins_GranularCSV.WriteGranularCSV
    WriteOutputs outputs

    KernelEngine.DomainOutputs = outputs
End Sub


' ReadUWInputs
' Reads all 4 sections from UW Inputs tab. Populates module-level arrays.
Private Sub ReadUWInputs()
    Dim wsUW As Worksheet
    On Error Resume Next
    Set wsUW = ThisWorkbook.Sheets("UW Inputs")
    On Error GoTo 0

    If wsUW Is Nothing Then
        KernelConfig.LogError SEV_WARN, "InsuranceDomainEngine", "W-310", _
            "UW Inputs tab not found. No programs loaded.", _
            "MANUAL BYPASS: Create UW Inputs tab with program data or run Bootstrap first."
        m_numProgs = 0
        Exit Sub
    End If

    m_numProgs = 0

    ' --- Section 1: Program identity + premium schedule ---
    Dim p As Long
    For p = 1 To MAX_PROGRAMS
        Dim dataRow As Long
        dataRow = UWIN_S1_DATA_ROW + p - 1
        Dim nameVal As String
        nameVal = Trim(CStr(wsUW.Cells(dataRow, UWIN_COL_NAME).Value))
        If Len(nameVal) = 0 Then Exit For

        m_numProgs = m_numProgs + 1
        m_progName(p) = nameVal
        m_progBU(p) = Trim(CStr(wsUW.Cells(dataRow, UWIN_COL_BU).Value))

        Dim termVal As Variant
        termVal = wsUW.Cells(dataRow, UWIN_COL_TERM).Value
        If IsNumeric(termVal) Then
            m_progTerm(p) = CLng(termVal)
        Else
            m_progTerm(p) = 12
        End If

        Dim q As Long
        For q = 1 To UWIN_GWP_QUARTERS
            Dim gwpVal As Variant
            gwpVal = wsUW.Cells(dataRow, UWIN_GWP_START_COL + q - 1).Value
            If IsNumeric(gwpVal) Then
                m_gwpQtr(p, q) = CDbl(gwpVal)
            Else
                m_gwpQtr(p, q) = 0
            End If
        Next q

        Dim growthVal As Variant
        growthVal = wsUW.Cells(dataRow, UWIN_GROWTH_COL).Value
        If IsNumeric(growthVal) Then
            m_gwpGrowth(p) = CDbl(growthVal)
        Else
            m_gwpGrowth(p) = 0
        End If
    Next p

    ' --- Section 2: Commission rates ---
    For p = 1 To m_numProgs
        Dim commRow As Long
        commRow = UWIN_S2_DATA_ROW + p - 1
        Dim yr As Long
        For yr = 1 To 5
            Dim commVal As Variant
            commVal = wsUW.Cells(commRow, UWIN_COMM_START_COL + yr - 1).Value
            If IsNumeric(commVal) Then
                m_commRate(p, yr) = CDbl(commVal)
            Else
                m_commRate(p, yr) = 0
            End If
        Next yr
    Next p

    ' --- Section 3: Loss assumptions (3 rows per program) ---
    For p = 1 To m_numProgs
        Dim lyr As Long
        For lyr = 1 To NUM_LAYERS
            Dim lossRow As Long
            lossRow = UWIN_S3_DATA_ROW + (p - 1) * NUM_LAYERS + (lyr - 1)

            m_lyrLOB(p, lyr) = Trim(CStr(wsUW.Cells(lossRow, UWIN_LOSS_COL_PATTERN).Value))

            Dim tlVal As Variant
            tlVal = wsUW.Cells(lossRow, UWIN_LOSS_COL_LOSSDEV).Value
            If IsNumeric(tlVal) Then
                m_lyrLossTL(p, lyr) = CLng(tlVal)
            Else
                m_lyrLossTL(p, lyr) = 50
            End If

            Dim ctlVal As Variant
            ctlVal = wsUW.Cells(lossRow, UWIN_LOSS_COL_CNTDEV).Value
            If IsNumeric(ctlVal) Then
                m_lyrCntTL(p, lyr) = CLng(ctlVal)
            Else
                m_lyrCntTL(p, lyr) = 50
            End If

            ' Read Q1-Q4 ELR
            Dim elrQ As Long
            For elrQ = 1 To 4
                Dim elrVal As Variant
                elrVal = wsUW.Cells(lossRow, UWIN_LOSS_COL_ELR1 + elrQ - 1).Value
                If IsNumeric(elrVal) Then
                    m_lyrELR(p, lyr, elrQ) = CDbl(elrVal)
                Else
                    m_lyrELR(p, lyr, elrQ) = 0
                End If
            Next elrQ

            ' Attritional: Q2-Q4 = Q1 (uniform ELR)
            If lyr = 1 Then
                m_lyrELR(p, 1, 2) = m_lyrELR(p, 1, 1)
                m_lyrELR(p, 1, 3) = m_lyrELR(p, 1, 1)
                m_lyrELR(p, 1, 4) = m_lyrELR(p, 1, 1)
            End If

            ' Read severity
            Dim sevVal As Variant
            sevVal = wsUW.Cells(lossRow, UWIN_LOSS_COL_SEV).Value
            If IsNumeric(sevVal) Then
                m_lyrSev(p, lyr) = CDbl(sevVal)
            Else
                m_lyrSev(p, lyr) = 0
            End If

            ' Read Q1-Q4 frequency
            Dim fqQ As Long
            For fqQ = 1 To 4
                Dim fqVal As Variant
                fqVal = wsUW.Cells(lossRow, UWIN_LOSS_COL_FREQ1 + fqQ - 1).Value
                If IsNumeric(fqVal) Then
                    m_lyrFreq(p, lyr, fqQ) = CDbl(fqVal)
                Else
                    m_lyrFreq(p, lyr, fqQ) = 0
                End If
            Next fqQ

            ' Determine if layer is active (any ELR > 0)
            m_lyrActive(p, lyr) = False
            For elrQ = 1 To 4
                If m_lyrELR(p, lyr, elrQ) > 0 Then
                    m_lyrActive(p, lyr) = True
                    Exit For
                End If
            Next elrQ
        Next lyr
    Next p

    ' --- Section 4: QS Reinsurance terms (Attr + Seas only, CAT excluded) ---
    For p = 1 To m_numProgs
        Dim reinsRow As Long
        reinsRow = UWIN_S4_DATA_ROW + p - 1
        For yr = 1 To 5
            Dim baseCol As Long
            baseCol = UWIN_REINS_START_COL + (yr - 1) * UWIN_REINS_COLS_PER_YEAR
            Dim cedeVal As Variant
            cedeVal = wsUW.Cells(reinsRow, baseCol).Value
            If IsNumeric(cedeVal) Then
                m_reinsCedePct(p, yr) = CDbl(cedeVal)
            Else
                m_reinsCedePct(p, yr) = 0
            End If

            Dim cedeCommVal As Variant
            cedeCommVal = wsUW.Cells(reinsRow, baseCol + 1).Value
            If IsNumeric(cedeCommVal) Then
                m_reinsCedeComm(p, yr) = CDbl(cedeCommVal)
            Else
                m_reinsCedeComm(p, yr) = 0
            End If

            Dim frontVal As Variant
            frontVal = wsUW.Cells(reinsRow, baseCol + 2).Value
            If IsNumeric(frontVal) Then
                m_reinsFrontFee(p, yr) = CDbl(frontVal)
            Else
                m_reinsFrontFee(p, yr) = 0
            End If
        Next yr
    Next p

    ' --- Section 5: XOL Reinsurance spend ---
    For p = 1 To m_numProgs
        Dim xolRow As Long
        xolRow = UWIN_S5_DATA_ROW + p - 1
        For yr = 1 To 5
            Dim catVal As Variant
            catVal = wsUW.Cells(xolRow, UWIN_XOL_CAT_START_COL + yr - 1).Value
            If IsNumeric(catVal) Then
                m_xolCat(p, yr) = CDbl(catVal)
            Else
                m_xolCat(p, yr) = 0
            End If
            Dim othVal As Variant
            othVal = wsUW.Cells(xolRow, UWIN_XOL_OTHER_START_COL + yr - 1).Value
            If IsNumeric(othVal) Then
                m_xolOther(p, yr) = CDbl(othVal)
            Else
                m_xolOther(p, yr) = 0
            End If
        Next yr
    Next p

    KernelConfig.LogError SEV_INFO, "InsuranceDomainEngine", "I-310", _
        "ReadUWInputs: " & m_numProgs & " programs loaded.", ""
End Sub


' ApplyUWInputsFormatting
' Applies number formats and data validation to UW Inputs input cells.
Private Sub ApplyUWInputsFormatting()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("UW Inputs")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim s1E As Long: s1E = UWIN_S1_DATA_ROW + MAX_PROGRAMS - 1
    Dim s2E As Long: s2E = UWIN_S2_DATA_ROW + MAX_PROGRAMS - 1
    Dim s3E As Long: s3E = UWIN_S3_DATA_ROW + MAX_PROGRAMS * NUM_LAYERS - 1
    Dim s4E As Long: s4E = UWIN_S4_DATA_ROW + MAX_PROGRAMS - 1
    Dim s5E As Long: s5E = UWIN_S5_DATA_ROW + MAX_PROGRAMS - 1

    ' S1: GWP cols = #,##0; Growth col = 0.0%
    ws.Range(ws.Cells(UWIN_S1_DATA_ROW, UWIN_GWP_START_COL), _
        ws.Cells(s1E, UWIN_GWP_START_COL + UWIN_GWP_QUARTERS - 1)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(UWIN_S1_DATA_ROW, UWIN_GROWTH_COL), _
        ws.Cells(s1E, UWIN_GROWTH_COL)).NumberFormat = "0.0%"

    ' S2: Commission rates = 0.0%
    ws.Range(ws.Cells(UWIN_S2_DATA_ROW, UWIN_COMM_START_COL), _
        ws.Cells(s2E, UWIN_COMM_START_COL + 4)).NumberFormat = "0.0%"

    ' S3: Dev tails = #,##0; ELR = 0.0%; Severity = #,##0
    ws.Range(ws.Cells(UWIN_S3_DATA_ROW, UWIN_LOSS_COL_LOSSDEV), _
        ws.Cells(s3E, UWIN_LOSS_COL_CNTDEV)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(UWIN_S3_DATA_ROW, UWIN_LOSS_COL_ELR1), _
        ws.Cells(s3E, UWIN_LOSS_COL_ELR1 + 3)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(UWIN_S3_DATA_ROW, UWIN_LOSS_COL_SEV), _
        ws.Cells(s3E, UWIN_LOSS_COL_SEV)).NumberFormat = "#,##0"

    ' S3: Pattern data validation dropdown (Property or Casualty)
    Dim rngPat As Range
    Set rngPat = ws.Range(ws.Cells(UWIN_S3_DATA_ROW, UWIN_LOSS_COL_PATTERN), _
        ws.Cells(s3E, UWIN_LOSS_COL_PATTERN))
    rngPat.Validation.Delete
    rngPat.Validation.Add Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, Formula1:="Property,Casualty"

    ' S4: All reinsurance pct cols = 0.0%
    ws.Range(ws.Cells(UWIN_S4_DATA_ROW, UWIN_REINS_START_COL), _
        ws.Cells(s4E, UWIN_REINS_START_COL + 5 * UWIN_REINS_COLS_PER_YEAR - 1)).NumberFormat = "0.0%"

    ' S5: XOL spend = #,##0
    ws.Range(ws.Cells(UWIN_S5_DATA_ROW, UWIN_XOL_CAT_START_COL), _
        ws.Cells(s5E, UWIN_XOL_OTHER_START_COL + 4)).NumberFormat = "#,##0"

    ' S3: Check col conditional formatting (Pass=green, FAIL=red)
    Dim rC As Range
    Set rC = ws.Range(ws.Cells(UWIN_S3_DATA_ROW, UWIN_LOSS_COL_FREQ1 + 4), _
        ws.Cells(s3E, UWIN_LOSS_COL_FREQ1 + 4))
    rC.FormatConditions.Delete
    With rC.FormatConditions.Add(xlCellValue, xlEqual, "=""Pass""")
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
    End With
    With rC.FormatConditions.Add(xlCellValue, xlEqual, "=""FAIL""")
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
    End With

    ' Hide RowID column (Column A) -- consistent with all other tabs
    ws.Columns(1).Hidden = True
End Sub


' LoadCurveParams
' Loads curve parameters for each program x layer using GetDefaultParams.
Private Sub LoadCurveParams()
    Dim p As Long
    Dim lyr As Long

    For p = 1 To m_numProgs
        Dim maxDevEnd As Long
        maxDevEnd = 0

        For lyr = 1 To NUM_LAYERS
            If Not m_lyrActive(p, lyr) Then GoTo NextCurveLayer

            Dim lob As String
            lob = m_lyrLOB(p, lyr)
            If Len(lob) = 0 Then lob = "Casualty"

            Dim lossTL As Long
            lossTL = m_lyrLossTL(p, lyr)
            Dim cntTL As Long
            cntTL = m_lyrCntTL(p, lyr)

            ' Get Paid curve params
            GetDefaultParams lob, "Paid", lossTL, _
                m_curves(p, lyr).distPd, m_curves(p, lyr).p1Pd, _
                m_curves(p, lyr).p2Pd, m_curves(p, lyr).maxAgePd

            ' Get Case Incurred curve params
            GetDefaultParams lob, "Case Incurred", lossTL, _
                m_curves(p, lyr).distCI, m_curves(p, lyr).p1CI, _
                m_curves(p, lyr).p2CI, m_curves(p, lyr).maxAgeCI

            ' Get Reported Count curve params
            GetDefaultParams lob, "Reported Count", cntTL, _
                m_curves(p, lyr).distRC, m_curves(p, lyr).p1RC, _
                m_curves(p, lyr).p2RC, m_curves(p, lyr).maxAgeRC

            ' Get Closed Count curve params
            GetDefaultParams lob, "Closed Count", cntTL, _
                m_curves(p, lyr).distCC, m_curves(p, lyr).p1CC, _
                m_curves(p, lyr).p2CC, m_curves(p, lyr).maxAgeCC

            ' Track max dev endpoint
            If m_curves(p, lyr).maxAgePd > maxDevEnd Then maxDevEnd = m_curves(p, lyr).maxAgePd
            If m_curves(p, lyr).maxAgeCI > maxDevEnd Then maxDevEnd = m_curves(p, lyr).maxAgeCI
            If m_curves(p, lyr).maxAgeRC > maxDevEnd Then maxDevEnd = m_curves(p, lyr).maxAgeRC
            If m_curves(p, lyr).maxAgeCC > maxDevEnd Then maxDevEnd = m_curves(p, lyr).maxAgeCC
NextCurveLayer:
        Next lyr

        If maxDevEnd = 0 Then maxDevEnd = 120
        ' BUG-099: Ensure devEnd >= horizon so all premium months are output
        ' BUG-114: devEnd must cover EP spillover (m_horizon + term)
        ' so DevelopLosses and WriteOutputs include all EP months
        Dim minDevEnd As Long
        minDevEnd = m_horizon + m_progTerm(p)
        If minDevEnd > MAX_DEV_ENDPOINT Then minDevEnd = MAX_DEV_ENDPOINT
        If maxDevEnd < minDevEnd Then maxDevEnd = minDevEnd
        m_devEnd(p) = maxDevEnd
    Next p
End Sub


' GetDefaultParams
' Config-driven curve lookup via anchor-based curve_library_config.
' Delegates to Ext_CurveLib.GetCurveParamsByTL for interpolation.
Private Sub GetDefaultParams(lob As String, curveType As String, tl As Long, _
    ByRef distName As String, ByRef p1 As Double, ByRef p2 As Double, ByRef maxAge As Long)
    Dim result As Variant
    result = Ext_CurveLib.GetCurveParamsByTL(lob, curveType, tl)
    distName = CStr(result(1))
    p1 = CDbl(result(2))
    p2 = CDbl(result(3))
    maxAge = CLng(result(5))
    If maxAge < 12 Then maxAge = 12
End Sub


' SpreadPremium
' Converts quarterly GWP to monthly WP. Applies growth for Y6-Y10 (DE-04).
Private Sub SpreadPremium()
    Dim p As Long
    For p = 1 To m_numProgs
        ' Quarters 1-20: direct from input
        Dim q As Long
        For q = 1 To UWIN_GWP_QUARTERS
            Dim qGWP As Double
            qGWP = m_gwpQtr(p, q)
            If qGWP = 0 Then GoTo NextQtr

            Dim m1 As Long
            m1 = (q - 1) * 3 + 1
            Dim mi As Long
            For mi = m1 To m1 + 2
                If mi <= m_horizon Then
                    m_wpMon(p, mi) = qGWP / 3
                End If
            Next mi
NextQtr:
        Next q

        ' Y6-Y10 growth: compound annually from Y5 base (DE-04)
        If m_gwpGrowth(p) <> 0 And m_horizon > 60 Then
            Dim growthYr As Long
            For growthYr = 6 To 10
                If (growthYr - 1) * 12 >= m_horizon Then Exit For
                Dim growthFactor As Double
                growthFactor = (1 + m_gwpGrowth(p)) ^ (growthYr - 5)

                ' Apply growth to same-quarter Y5 values
                Dim qInYr As Long
                For qInYr = 1 To 4
                    Dim y5qIdx As Long
                    y5qIdx = 16 + qInYr  ' Q1Y5=17, Q4Y5=20
                    Dim baseQGWP As Double
                    baseQGWP = m_gwpQtr(p, y5qIdx)
                    Dim grownQGWP As Double
                    grownQGWP = baseQGWP * growthFactor

                    m1 = ((growthYr - 1) * 4 + (qInYr - 1)) * 3 + 1
                    For mi = m1 To m1 + 2
                        If mi <= m_horizon Then
                            m_wpMon(p, mi) = grownQGWP / 3
                        End If
                    Next mi
                Next qInYr
            Next growthYr
        End If
    Next p
End Sub


' EarnPremium
' Earns monthly WP across policy term using mid-month assumption (DE-01).
Private Sub EarnPremium()
    Dim p As Long
    For p = 1 To m_numProgs
        Dim termMo As Long
        termMo = m_progTerm(p)
        If termMo <= 0 Then termMo = 12

        Dim m As Long
        For m = 1 To m_horizon
            If m_wpMon(p, m) = 0 Then GoTo NextEarnMonth

            Dim wp As Double
            wp = m_wpMon(p, m)

            ' Spread WP across earning window [m, m+term] (BUG-065)
            ' First month: half-month, last month: half-month, middle: full
            ' term+1 months: 1/(2T) + (T-1)*(1/T) + 1/(2T) = 1.0
            Dim ew As Long
            For ew = m To m + termMo
                If ew > MAX_HORIZON Then Exit For

                Dim frac As Double
                If ew = m Then
                    ' First month: half-month assumption
                    frac = 1# / (2# * CDbl(termMo))
                ElseIf ew = m + termMo Then
                    ' Spill-over month: remaining half-month
                    frac = 1# / (2# * CDbl(termMo))
                Else
                    ' Middle months: full month
                    frac = 1# / CDbl(termMo)
                End If

                m_epMon(p, ew) = m_epMon(p, ew) + wp * frac
            Next ew
NextEarnMonth:
        Next m
    Next p
End Sub


' ComputeUltimates
' EP x ELR -> ultimate losses per exposure month (Step 4).
' Frequency per $1M EP -> ultimate claim count.
Private Sub ComputeUltimates()
    Dim p As Long
    For p = 1 To m_numProgs
        Dim lyr As Long
        For lyr = 1 To NUM_LAYERS
            If Not m_lyrActive(p, lyr) Then GoTo NextUltLayer

            Dim ep As Long
            For ep = 1 To m_horizon
                If m_epMon(p, ep) = 0 Then GoTo NextUltMonth

                ' BUG-117: Ultimate = EP x ELR per exposure month.
                ' EP is the single source of truth for all development.
                ' Paid = m_ultMon * PaidCDF. CI = m_ultMon * CI_CDF.
                ' No epScale band-aids. CSV, Detail, Triangles all derive
                ' from m_ultMon which is EP-based.

                ' Determine quarter within year for this exposure month
                Dim qWithinYr As Long
                qWithinYr = ((ep - 1) Mod 12) \ 3 + 1

                ' Get ELR for this quarter
                Dim elr As Double
                If lyr = 1 Then
                    ' Attritional: always Q1 (uniform)
                    elr = m_lyrELR(p, 1, 1)
                Else
                    elr = m_lyrELR(p, lyr, qWithinYr)
                End If

                m_ultMon(p, lyr, ep) = m_epMon(p, ep) * elr

                ' Frequency: per $1M EP
                Dim freq As Double
                freq = m_lyrFreq(p, lyr, qWithinYr)
                m_cntUlt(p, lyr, ep) = m_epMon(p, ep) * freq / 1000000#
NextUltMonth:
            Next ep
NextUltLayer:
        Next lyr

        ' BUG-117: m_ultMon is now EP-based (= m_epMon * ELR).
        ' No separate blended ELR computation needed.
    Next p
End Sub


' DevelopLosses
' CurveLib CDF -> monthly emergence. Builds cumulative arrays (Step 5).
Private Sub DevelopLosses()
    Dim p As Long
    For p = 1 To m_numProgs
        Dim cm As Long
        ' Loop to full development endpoint (run-off beyond TimeHorizon)
        Dim devHorizon As Long
        devHorizon = m_devEnd(p)
        If devHorizon > MAX_DEV_ENDPOINT Then devHorizon = MAX_DEV_ENDPOINT
        For cm = 1 To devHorizon
            Dim layerPaid As Double
            Dim layerCI As Double
            Dim layerUlt As Double
            Dim layerRpt As Double
            Dim layerCls As Double
            Dim layerCntUlt As Double
            Dim qsPaid As Double
            Dim qsCI As Double
            Dim qsUlt As Double
            Dim qsRpt As Double
            Dim qsCls As Double
            Dim qsCntUlt As Double
            layerPaid = 0: layerCI = 0: layerUlt = 0
            layerRpt = 0: layerCls = 0: layerCntUlt = 0
            qsPaid = 0: qsCI = 0: qsUlt = 0
            qsRpt = 0: qsCls = 0: qsCntUlt = 0

            Dim lyr As Long
            For lyr = 1 To NUM_LAYERS
                If Not m_lyrActive(p, lyr) Then GoTo NextDevLayer

                ' Exposure months capped at TimeHorizon (premium only written 1..m_horizon)
                Dim epMax As Long
                epMax = cm
                If epMax > m_horizon Then epMax = m_horizon
                Dim ep As Long
                For ep = 1 To epMax
                    If m_ultMon(p, lyr, ep) = 0 And m_cntUlt(p, lyr, ep) = 0 Then GoTo NextDevEP

                    ' 1-based development age (DE-02)
                    Dim age As Long
                    age = cm - ep + 1

                    ' Skip if beyond dev endpoint
                    If age > m_devEnd(p) Then GoTo NextDevEP

                    ' DE-08: Mid-month average written date offset.
                    ' Policies written uniformly within exposure month,
                    ' so average inception = mid-month. CDF lookup uses
                    ' age - 0.5 to reflect actual elapsed development time.
                    Dim ageAdj As Double
                    ageAdj = CDbl(age) - 0.5

                    ' Evaluate curves
                    Dim paidPct As Double
                    paidPct = Ext_CurveLib.EvaluateCurve( _
                        m_curves(p, lyr).distPd, m_curves(p, lyr).p1Pd, _
                        m_curves(p, lyr).p2Pd, ageAdj, m_curves(p, lyr).maxAgePd)

                    Dim ciPct As Double
                    ciPct = Ext_CurveLib.EvaluateCurve( _
                        m_curves(p, lyr).distCI, m_curves(p, lyr).p1CI, _
                        m_curves(p, lyr).p2CI, ageAdj, m_curves(p, lyr).maxAgeCI)

                    Dim rptPct As Double
                    rptPct = Ext_CurveLib.EvaluateCurve( _
                        m_curves(p, lyr).distRC, m_curves(p, lyr).p1RC, _
                        m_curves(p, lyr).p2RC, ageAdj, m_curves(p, lyr).maxAgeRC)

                    Dim clsPct As Double
                    clsPct = Ext_CurveLib.EvaluateCurve( _
                        m_curves(p, lyr).distCC, m_curves(p, lyr).p1CC, _
                        m_curves(p, lyr).p2CC, ageAdj, m_curves(p, lyr).maxAgeCC)

                    ' BUG-117: m_ultMon is now EP-based. No epScale needed.
                    ' All development = m_ultMon * CDF directly.
                    layerPaid = layerPaid + m_ultMon(p, lyr, ep) * paidPct
                    layerCI = layerCI + m_ultMon(p, lyr, ep) * ciPct
                    layerUlt = layerUlt + m_ultMon(p, lyr, ep)
                    layerRpt = layerRpt + m_cntUlt(p, lyr, ep) * rptPct
                    layerCls = layerCls + m_cntUlt(p, lyr, ep) * clsPct
                    layerCntUlt = layerCntUlt + m_cntUlt(p, lyr, ep)

                    ' QS-subject: Attr(1) + Seas(2) only, CAT(3) excluded
                    If lyr <= 2 Then
                        qsPaid = qsPaid + m_ultMon(p, lyr, ep) * paidPct
                        qsCI = qsCI + m_ultMon(p, lyr, ep) * ciPct
                        qsUlt = qsUlt + m_ultMon(p, lyr, ep)
                        qsRpt = qsRpt + m_cntUlt(p, lyr, ep) * rptPct
                        qsCls = qsCls + m_cntUlt(p, lyr, ep) * clsPct
                        qsCntUlt = qsCntUlt + m_cntUlt(p, lyr, ep)
                    End If
NextDevEP:
                Next ep
NextDevLayer:
            Next lyr

            ' Store cumulative program-level values
            m_cumPaid(p, cm) = layerPaid
            m_cumCI(p, cm) = layerCI
            m_cumUlt(p, cm) = layerUlt
            m_cumRpt(p, cm) = layerRpt
            m_cumCls(p, cm) = layerCls
            m_cumCntUlt(p, cm) = layerCntUlt
            ' QS-subject cumulatives (Attr + Seas only)
            m_qsPaid(p, cm) = qsPaid
            m_qsCI(p, cm) = qsCI
            m_qsUlt(p, cm) = qsUlt
            m_qsRpt(p, cm) = qsRpt
            m_qsCls(p, cm) = qsCls
            m_qsCntUlt(p, cm) = qsCntUlt

            ' BUG-117: m_cumUlt is now EP-based directly (no QS ratio needed).
            ' m_qsUlt already accumulated from EP-based qsUlt in the layer loop.
            ' Cumulative EP for reference
            If cm = 1 Then
                m_cumEP(p, cm) = m_epMon(p, cm)
            Else
                m_cumEP(p, cm) = m_cumEP(p, cm - 1) + m_epMon(p, cm)
            End If
        Next cm
    Next p
End Sub


' Inc -- extracts monthly increment from a cumulative array.
' Returns cum(cm) when cm=1, else cum(cm)-cum(cm-1).
' Eliminates the repeated If cm=1 Then ... Else ... pattern (DRY).
Private Function Inc(ByRef cum() As Double, p As Long, cm As Long) As Double
    If cm = 1 Then
        Inc = cum(p, 1)
    Else
        Inc = cum(p, cm) - cum(p, cm - 1)
    End If
End Function


' WriteOutputs
' Assembles Gross + Ceded blocks and writes to outputs array (Step 6).
' All fields written as true incremental (MTD change).
' Balance fields: change in EOP = EOP(cm) - EOP(cm-1).
' Flow fields: monthly increment from cumulative.
' Ceded count fields: NOT written (Derived by kernel per CR-03).
' Net fields: NOT written (Derived by kernel).
Private Sub WriteOutputs(ByRef outputs As Variant)
    ' Column indices via ColIndex (AP-08)
    Dim cEntity As Long: cEntity = KernelConfig.ColIndex("EntityName")
    Dim cPeriod As Long: cPeriod = KernelConfig.ColIndex("CalPeriod")
    Dim cQuarter As Long: cQuarter = KernelConfig.ColIndex("CalQuarter")
    Dim cYear As Long: cYear = KernelConfig.ColIndex("CalYear")
    Dim cGWP As Long: cGWP = KernelConfig.ColIndex("G_WP")
    Dim cGEP As Long: cGEP = KernelConfig.ColIndex("G_EP")
    Dim cGWComm As Long: cGWComm = KernelConfig.ColIndex("G_WComm")
    Dim cGEComm As Long: cGEComm = KernelConfig.ColIndex("G_EComm")
    Dim cGWFF As Long: cGWFF = KernelConfig.ColIndex("G_WFrontFee")
    Dim cGEFF As Long: cGEFF = KernelConfig.ColIndex("G_EFrontFee")
    Dim cGPaid As Long: cGPaid = KernelConfig.ColIndex("G_Paid")
    Dim cGCaseRsv As Long: cGCaseRsv = KernelConfig.ColIndex("G_CaseRsv")
    Dim cGCaseInc As Long: cGCaseInc = KernelConfig.ColIndex("G_CaseInc")
    Dim cGIBNR As Long: cGIBNR = KernelConfig.ColIndex("G_IBNR")
    Dim cGUnpaid As Long: cGUnpaid = KernelConfig.ColIndex("G_Unpaid")
    Dim cGUlt As Long: cGUlt = KernelConfig.ColIndex("G_Ult")
    Dim cGClsCt As Long: cGClsCt = KernelConfig.ColIndex("G_ClsCt")
    Dim cGOpenCt As Long: cGOpenCt = KernelConfig.ColIndex("G_OpenCt")
    Dim cGRptCt As Long: cGRptCt = KernelConfig.ColIndex("G_RptCt")
    Dim cGUltCt As Long: cGUltCt = KernelConfig.ColIndex("G_UltCt")
    Dim cGIBNRCt As Long: cGIBNRCt = KernelConfig.TryColIndex("G_IBNRCt")
    Dim cGUnclCt As Long: cGUnclCt = KernelConfig.TryColIndex("G_UnclosedCt")
    Dim cCWP As Long: cCWP = KernelConfig.ColIndex("C_WP")
    Dim cCEP As Long: cCEP = KernelConfig.ColIndex("C_EP")
    Dim cCWComm As Long: cCWComm = KernelConfig.ColIndex("C_WComm")
    Dim cCEComm As Long: cCEComm = KernelConfig.ColIndex("C_EComm")
    Dim cCWFF As Long: cCWFF = KernelConfig.ColIndex("C_WFrontFee")
    Dim cCEFF As Long: cCEFF = KernelConfig.ColIndex("C_EFrontFee")
    Dim cCPaid As Long: cCPaid = KernelConfig.ColIndex("C_Paid")
    Dim cCCaseRsv As Long: cCCaseRsv = KernelConfig.ColIndex("C_CaseRsv")
    Dim cCCaseInc As Long: cCCaseInc = KernelConfig.ColIndex("C_CaseInc")
    Dim cCIBNR As Long: cCIBNR = KernelConfig.ColIndex("C_IBNR")
    Dim cCUnpaid As Long: cCUnpaid = KernelConfig.ColIndex("C_Unpaid")
    Dim cCUlt As Long: cCUlt = KernelConfig.ColIndex("C_Ult")
    ' Ceded count columns (QS-subject: Attr+Seas only, CAT excluded)
    Dim cCClsCt As Long: cCClsCt = KernelConfig.ColIndex("C_ClsCt")
    Dim cCOpenCt As Long: cCOpenCt = KernelConfig.ColIndex("C_OpenCt")
    Dim cCRptCt As Long: cCRptCt = KernelConfig.ColIndex("C_RptCt")
    Dim cCUltCt As Long: cCUltCt = KernelConfig.ColIndex("C_UltCt")
    Dim cCIBNRCt As Long: cCIBNRCt = KernelConfig.TryColIndex("C_IBNRCt")
    Dim cCUnclCt As Long: cCUnclCt = KernelConfig.TryColIndex("C_UnclosedCt")
    ' XOL breakout columns
    Dim cCQWP As Long: cCQWP = KernelConfig.ColIndex("CQ_WP")
    Dim cCQEP As Long: cCQEP = KernelConfig.ColIndex("CQ_EP")
    Dim cXCWP As Long: cXCWP = KernelConfig.ColIndex("XC_WP")
    Dim cXCEP As Long: cXCEP = KernelConfig.ColIndex("XC_EP")
    Dim cXOWP As Long: cXOWP = KernelConfig.ColIndex("XO_WP")
    Dim cXOEP As Long: cXOEP = KernelConfig.ColIndex("XO_EP")

    ' Variable row offsets per program (each has different devEnd)
    Dim rowOffset As Long
    rowOffset = 0
    Dim p As Long
    For p = 1 To m_numProgs
        Dim pDevEnd As Long
        pDevEnd = m_devEnd(p)
        If pDevEnd > MAX_DEV_ENDPOINT Then pDevEnd = MAX_DEV_ENDPOINT
        Dim cm As Long
        For cm = 1 To pDevEnd
            Dim row As Long
            row = rowOffset + cm

            ' --- Dimensions ---
            outputs(row, cEntity) = m_progName(p)
            outputs(row, cPeriod) = cm
            outputs(row, cQuarter) = ((cm - 1) Mod 12) \ 3 + 1
            outputs(row, cYear) = ((cm - 1) \ 12) + 1

            ' --- Rate year for this month (caps at Y5 per CR-07) ---
            Dim rateYr As Long
            rateYr = ((cm - 1) \ 12) + 1
            If rateYr > 5 Then rateYr = 5

            Dim commPct As Double: commPct = m_commRate(p, rateYr)
            Dim cedePct As Double: cedePct = m_reinsCedePct(p, rateYr)
            Dim cedeCommPct As Double: cedeCommPct = m_reinsCedeComm(p, rateYr)
            Dim frontFeePct As Double: frontFeePct = m_reinsFrontFee(p, rateYr)

            ' --- Incremental (MTD) from cumulative via Inc() helper ---
            Dim mtdPaid As Double: mtdPaid = Inc(m_cumPaid, p, cm)
            Dim mtdCI As Double: mtdCI = Inc(m_cumCI, p, cm)
            ' BUG-112: G_Ult uses EP-based ultimate for calendar quarter
            ' financial statements (loss expense = EP x ELR).
            ' WP-based m_cumUlt drives IBNR/Unpaid reserves on BS.
            ' BS imbalance resolved via IBNR adjustment (see below).
            Dim mtdUlt As Double: mtdUlt = Inc(m_cumUlt, p, cm)
            Dim mtdRpt As Double: mtdRpt = Inc(m_cumRpt, p, cm)
            Dim mtdCls As Double: mtdCls = Inc(m_cumCls, p, cm)
            Dim mtdCntUlt As Double: mtdCntUlt = Inc(m_cumCntUlt, p, cm)

            ' --- Gross Block ---
            Dim gWP As Double: gWP = m_wpMon(p, cm)
            Dim gEP As Double: gEP = m_epMon(p, cm)

            outputs(row, cGWP) = gWP
            outputs(row, cGEP) = gEP
            outputs(row, cGWComm) = gWP * commPct
            outputs(row, cGEComm) = gEP * commPct
            outputs(row, cGWFF) = gWP * frontFeePct
            outputs(row, cGEFF) = gEP * frontFeePct
            outputs(row, cGPaid) = mtdPaid
            outputs(row, cGCaseInc) = mtdCI
            outputs(row, cGUlt) = mtdUlt
            outputs(row, cGClsCt) = mtdCls
            outputs(row, cGRptCt) = mtdRpt
            outputs(row, cGUltCt) = mtdCntUlt

            ' Balance columns: true incremental = EOP(cm) - EOP(cm-1)
            ' BUG-112: IBNR and Unpaid use EP-based cumulative ultimate
            ' m_cumUlt is now EP-based (BUG-117).
            ' Case Reserve = CI - Paid (both WP-based development, correct).
            Dim eopCaseRsv As Double: eopCaseRsv = m_cumCI(p, cm) - m_cumPaid(p, cm)
            Dim eopIBNR As Double: eopIBNR = m_cumUlt(p, cm) - m_cumCI(p, cm)
            Dim eopUnpaid As Double: eopUnpaid = m_cumUlt(p, cm) - m_cumPaid(p, cm)
            Dim eopOpenCt As Double: eopOpenCt = m_cumRpt(p, cm) - m_cumCls(p, cm)
            Dim eopIBNRCt As Double: eopIBNRCt = m_cumCntUlt(p, cm) - m_cumRpt(p, cm)
            Dim eopUnclCt As Double: eopUnclCt = m_cumCntUlt(p, cm) - m_cumCls(p, cm)
            If cm = 1 Then
                outputs(row, cGCaseRsv) = eopCaseRsv
                outputs(row, cGIBNR) = eopIBNR
                outputs(row, cGUnpaid) = eopUnpaid
                outputs(row, cGOpenCt) = eopOpenCt
                If cGIBNRCt > 0 Then outputs(row, cGIBNRCt) = eopIBNRCt
                If cGUnclCt > 0 Then outputs(row, cGUnclCt) = eopUnclCt
            Else
                Dim prCaseRsv As Double: prCaseRsv = m_cumCI(p, cm - 1) - m_cumPaid(p, cm - 1)
                Dim prIBNR As Double: prIBNR = m_cumUlt(p, cm - 1) - m_cumCI(p, cm - 1)
                Dim prUnpaid As Double: prUnpaid = m_cumUlt(p, cm - 1) - m_cumPaid(p, cm - 1)
                Dim prOpenCt As Double: prOpenCt = m_cumRpt(p, cm - 1) - m_cumCls(p, cm - 1)
                Dim prIBNRCt As Double: prIBNRCt = m_cumCntUlt(p, cm - 1) - m_cumRpt(p, cm - 1)
                Dim prUnclCt As Double: prUnclCt = m_cumCntUlt(p, cm - 1) - m_cumCls(p, cm - 1)
                outputs(row, cGCaseRsv) = eopCaseRsv - prCaseRsv
                outputs(row, cGIBNR) = eopIBNR - prIBNR
                outputs(row, cGUnpaid) = eopUnpaid - prUnpaid
                outputs(row, cGOpenCt) = eopOpenCt - prOpenCt
                If cGIBNRCt > 0 Then outputs(row, cGIBNRCt) = eopIBNRCt - prIBNRCt
                If cGUnclCt > 0 Then outputs(row, cGUnclCt) = eopUnclCt - prUnclCt
            End If

            ' --- Ceded Block ---
            ' QS breakout: proportional cession
            Dim qsWP As Double: qsWP = gWP * cedePct
            Dim qsEP As Double: qsEP = gEP * cedePct
            outputs(row, cCQWP) = qsWP
            outputs(row, cCQEP) = qsEP
            ' XOL breakout: annual premium written at 1/1 (Q1), earned evenly over 12 months
            ' XOL only applies during writing period (no coverage after treaty expires)
            Dim xolCatEP As Double
            Dim xolOtherEP As Double
            Dim xolCatWP As Double
            Dim xolOtherWP As Double
            If cm <= m_horizon Then
                xolCatEP = m_xolCat(p, rateYr) / 12
                xolOtherEP = m_xolOther(p, rateYr) / 12
                ' Written premium: full annual amount in month 1 of each year
                Dim monthInYear As Long
                monthInYear = ((cm - 1) Mod 12) + 1
                If monthInYear = 1 Then
                    xolCatWP = m_xolCat(p, rateYr)
                    xolOtherWP = m_xolOther(p, rateYr)
                Else
                    xolCatWP = 0
                    xolOtherWP = 0
                End If
            Else
                xolCatEP = 0
                xolOtherEP = 0
                xolCatWP = 0
                xolOtherWP = 0
            End If
            outputs(row, cXCWP) = xolCatWP
            outputs(row, cXCEP) = xolCatEP
            outputs(row, cXOWP) = xolOtherWP
            outputs(row, cXOEP) = xolOtherEP
            ' Total ceded = QS + CAT XOL + Other XOL
            outputs(row, cCWP) = qsWP + xolCatWP + xolOtherWP
            outputs(row, cCEP) = qsEP + xolCatEP + xolOtherEP
            outputs(row, cCWComm) = gWP * cedePct * cedeCommPct
            outputs(row, cCEComm) = gEP * cedePct * cedeCommPct
            outputs(row, cCWFF) = 0   ' DE-05: no ceded fronting fee
            outputs(row, cCEFF) = 0   ' DE-05: no ceded fronting fee
            ' QS losses: Attr+Seas only (CAT excluded)
            Dim qsMtdPaid As Double: qsMtdPaid = Inc(m_qsPaid, p, cm)
            Dim qsMtdCI As Double: qsMtdCI = Inc(m_qsCI, p, cm)
            ' C_Ult uses EP-based ultimate (BUG-112)
            Dim qsMtdUlt As Double: qsMtdUlt = Inc(m_qsUlt, p, cm)
            outputs(row, cCPaid) = qsMtdPaid * cedePct
            outputs(row, cCCaseInc) = qsMtdCI * cedePct
            outputs(row, cCUlt) = qsMtdUlt * cedePct

            ' Ceded Balance columns: true incremental = EOP(cm) - EOP(cm-1)
            ' BUG-117: Ceded IBNR/Unpaid use EP-based m_qsUlt
            Dim cEopCR As Double: cEopCR = (m_qsCI(p, cm) - m_qsPaid(p, cm)) * cedePct
            Dim cEopIB As Double: cEopIB = (m_qsUlt(p, cm) - m_qsCI(p, cm)) * cedePct
            Dim cEopUP As Double: cEopUP = (m_qsUlt(p, cm) - m_qsPaid(p, cm)) * cedePct
            If cm = 1 Then
                outputs(row, cCCaseRsv) = cEopCR
                outputs(row, cCIBNR) = cEopIB
                outputs(row, cCUnpaid) = cEopUP
            Else
                Dim cPrCR As Double: cPrCR = (m_qsCI(p, cm - 1) - m_qsPaid(p, cm - 1)) * cedePct
                Dim cPrIB As Double: cPrIB = (m_qsUlt(p, cm - 1) - m_qsCI(p, cm - 1)) * cedePct
                Dim cPrUP As Double: cPrUP = (m_qsUlt(p, cm - 1) - m_qsPaid(p, cm - 1)) * cedePct
                outputs(row, cCCaseRsv) = cEopCR - cPrCR
                outputs(row, cCIBNR) = cEopIB - cPrIB
                outputs(row, cCUnpaid) = cEopUP - cPrUP
            End If
            ' Ceded counts: QS-subject (Attr+Seas only, CAT excluded)
            Dim qsMtdCls As Double: qsMtdCls = Inc(m_qsCls, p, cm)
            Dim qsMtdRpt As Double: qsMtdRpt = Inc(m_qsRpt, p, cm)
            Dim qsMtdCntUlt As Double: qsMtdCntUlt = Inc(m_qsCntUlt, p, cm)
            outputs(row, cCClsCt) = qsMtdCls
            outputs(row, cCRptCt) = qsMtdRpt
            outputs(row, cCUltCt) = qsMtdCntUlt
            ' Ceded balance counts: true incremental = EOP(cm) - EOP(cm-1)
            Dim cEopOpen As Double: cEopOpen = m_qsRpt(p, cm) - m_qsCls(p, cm)
            If cm = 1 Then
                outputs(row, cCOpenCt) = cEopOpen
            Else
                Dim cPrOpen As Double: cPrOpen = m_qsRpt(p, cm - 1) - m_qsCls(p, cm - 1)
                outputs(row, cCOpenCt) = cEopOpen - cPrOpen
            End If
            If cCIBNRCt > 0 Then
                Dim cEopIBCt As Double: cEopIBCt = m_qsCntUlt(p, cm) - m_qsRpt(p, cm)
                If cm = 1 Then
                    outputs(row, cCIBNRCt) = cEopIBCt
                Else
                    Dim cPrIBCt As Double: cPrIBCt = m_qsCntUlt(p, cm - 1) - m_qsRpt(p, cm - 1)
                    outputs(row, cCIBNRCt) = cEopIBCt - cPrIBCt
                End If
            End If
            If cCUnclCt > 0 Then
                Dim cEopUCt As Double: cEopUCt = m_qsCntUlt(p, cm) - m_qsCls(p, cm)
                If cm = 1 Then
                    outputs(row, cCUnclCt) = cEopUCt
                Else
                    Dim cPrUCt As Double: cPrUCt = m_qsCntUlt(p, cm - 1) - m_qsCls(p, cm - 1)
                    outputs(row, cCUnclCt) = cEopUCt - cPrUCt
                End If
            End If
            ' Net block: NOT written -- Derived by kernel
        Next cm

        ' --- Tail Closure Row ---
        ' Forces ITD: paid=ultimate, reserves=0, counts closed.
        ' Written as incremental deltas from final EOP to target.
        Dim tRow As Long: tRow = rowOffset + pDevEnd + 1
        Dim tPer As Long: tPer = pDevEnd + 1

        ' Dimensions
        outputs(tRow, cEntity) = m_progName(p)
        outputs(tRow, cPeriod) = tPer
        outputs(tRow, cQuarter) = ((tPer - 1) Mod 12) \ 3 + 1
        outputs(tRow, cYear) = ((tPer - 1) \ 12) + 1

        ' Rate year for tail (capped at Y5 per CR-07)
        Dim tRateYr As Long: tRateYr = ((tPer - 1) \ 12) + 1
        If tRateYr > 5 Then tRateYr = 5
        Dim tCedePct As Double: tCedePct = m_reinsCedePct(p, tRateYr)

        ' Gross tail: delta to close
        ' BUG-112: Close to EP-based ultimate (consistent with IBNR/Unpaid)
        ' Paid delta = EP Ult - ITD Paid (so Sum(Paid) = EP Ult)
        outputs(tRow, cGPaid) = m_cumUlt(p, pDevEnd) - m_cumPaid(p, pDevEnd)
        ' CaseInc delta = EP Ult - ITD CI
        outputs(tRow, cGCaseInc) = m_cumUlt(p, pDevEnd) - m_cumCI(p, pDevEnd)
        ' Ult delta = 0 (no new ultimate at tail)
        outputs(tRow, cGUlt) = 0
        ' Reserve deltas: negate final EOP to bring Sum to 0
        Dim fCaseRsv As Double: fCaseRsv = m_cumCI(p, pDevEnd) - m_cumPaid(p, pDevEnd)
        Dim fIBNR As Double: fIBNR = m_cumUlt(p, pDevEnd) - m_cumCI(p, pDevEnd)
        Dim fUnpaid As Double: fUnpaid = m_cumUlt(p, pDevEnd) - m_cumPaid(p, pDevEnd)
        outputs(tRow, cGCaseRsv) = -fCaseRsv
        outputs(tRow, cGIBNR) = -fIBNR
        outputs(tRow, cGUnpaid) = -fUnpaid
        ' Count deltas: close all open counts
        outputs(tRow, cGClsCt) = m_cumRpt(p, pDevEnd) - m_cumCls(p, pDevEnd)
        outputs(tRow, cGRptCt) = m_cumCntUlt(p, pDevEnd) - m_cumRpt(p, pDevEnd)
        outputs(tRow, cGUltCt) = 0
        Dim fOpenCt As Double: fOpenCt = m_cumRpt(p, pDevEnd) - m_cumCls(p, pDevEnd)
        outputs(tRow, cGOpenCt) = -fOpenCt
        If cGIBNRCt > 0 Then
            Dim fIBNRCt As Double: fIBNRCt = m_cumCntUlt(p, pDevEnd) - m_cumRpt(p, pDevEnd)
            outputs(tRow, cGIBNRCt) = -fIBNRCt
        End If
        If cGUnclCt > 0 Then
            Dim fUnclCt As Double: fUnclCt = m_cumCntUlt(p, pDevEnd) - m_cumCls(p, pDevEnd)
            outputs(tRow, cGUnclCt) = -fUnclCt
        End If
        ' Premium/commission: 0 at tail (no new premium)
        outputs(tRow, cGWP) = 0
        outputs(tRow, cGEP) = 0
        outputs(tRow, cGWComm) = 0
        outputs(tRow, cGEComm) = 0
        outputs(tRow, cGWFF) = 0
        outputs(tRow, cGEFF) = 0

        ' Ceded tail: QS portion of gross tail deltas
        ' BUG-112: Ceded tail uses EP-based QS ultimate
        Dim qfCR As Double: qfCR = (m_qsCI(p, pDevEnd) - m_qsPaid(p, pDevEnd)) * tCedePct
        Dim qfIB As Double: qfIB = (m_qsUlt(p, pDevEnd) - m_qsCI(p, pDevEnd)) * tCedePct
        Dim qfUP As Double: qfUP = (m_qsUlt(p, pDevEnd) - m_qsPaid(p, pDevEnd)) * tCedePct
        outputs(tRow, cCPaid) = (m_qsUlt(p, pDevEnd) - m_qsPaid(p, pDevEnd)) * tCedePct
        outputs(tRow, cCCaseInc) = (m_qsUlt(p, pDevEnd) - m_qsCI(p, pDevEnd)) * tCedePct
        outputs(tRow, cCUlt) = 0
        outputs(tRow, cCCaseRsv) = -qfCR
        outputs(tRow, cCIBNR) = -qfIB
        outputs(tRow, cCUnpaid) = -qfUP
        outputs(tRow, cCWP) = 0
        outputs(tRow, cCEP) = 0
        outputs(tRow, cCWComm) = 0
        outputs(tRow, cCEComm) = 0
        outputs(tRow, cCWFF) = 0
        outputs(tRow, cCEFF) = 0
        ' XOL breakout: 0 at tail (no premium)
        outputs(tRow, cCQWP) = 0
        outputs(tRow, cCQEP) = 0
        outputs(tRow, cXCWP) = 0
        outputs(tRow, cXCEP) = 0
        outputs(tRow, cXOWP) = 0
        outputs(tRow, cXOEP) = 0
        ' Ceded tail counts: close QS-subject open counts
        outputs(tRow, cCClsCt) = m_qsRpt(p, pDevEnd) - m_qsCls(p, pDevEnd)
        outputs(tRow, cCRptCt) = m_qsCntUlt(p, pDevEnd) - m_qsRpt(p, pDevEnd)
        outputs(tRow, cCUltCt) = 0
        Dim qfOpen As Double: qfOpen = m_qsRpt(p, pDevEnd) - m_qsCls(p, pDevEnd)
        outputs(tRow, cCOpenCt) = -qfOpen
        If cCIBNRCt > 0 Then
            Dim qfIBCt As Double: qfIBCt = m_qsCntUlt(p, pDevEnd) - m_qsRpt(p, pDevEnd)
            outputs(tRow, cCIBNRCt) = -qfIBCt
        End If
        If cCUnclCt > 0 Then
            Dim qfUCt As Double: qfUCt = m_qsCntUlt(p, pDevEnd) - m_qsCls(p, pDevEnd)
            outputs(tRow, cCUnclCt) = -qfUCt
        End If

        rowOffset = rowOffset + pDevEnd + 1
    Next p
End Sub

