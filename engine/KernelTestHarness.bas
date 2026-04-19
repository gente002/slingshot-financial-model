Attribute VB_Name = "KernelTestHarness"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelTestHarness.bas
' Purpose: Helper functions for domain test development. Makes it easy to write
'          tests without manual array setup.
' =============================================================================

' Sequential test ID counter for domain tests
Private m_domTestSeq As Long


' =============================================================================
' SetupTestInputs
' Populates the Inputs tab with default values from InputSchema,
' then applies overrides from the dictionary.
'
' overrides is a Scripting.Dictionary:
'   Key = "Section|ParamName|EntityIndex" (e.g., "Assumptions|Units|1")
'   Value = the override value
' =============================================================================
Public Function SetupTestInputs(overrides As Object) As Boolean
    On Error GoTo ErrHandler

    ' Ensure config is loaded
    KernelConfig.LoadAllConfig

    Dim wsInputs As Worksheet
    Set wsInputs = ThisWorkbook.Sheets(TAB_INPUTS)

    Dim paramCount As Long
    paramCount = KernelConfig.GetInputCount()

    ' Count entities currently on the Inputs tab
    Dim entityCount As Long
    entityCount = 0
    Dim ecol As Long
    For ecol = INPUT_ENTITY_START_COL To INPUT_ENTITY_START_COL + INPUT_MAX_ENTITIES - 1
        If Len(Trim(CStr(wsInputs.Cells(3, ecol).Value))) = 0 Then Exit For
        entityCount = entityCount + 1
    Next ecol

    If entityCount = 0 Then entityCount = 1

    ' Apply overrides to existing Inputs tab values (preserves per-entity fixture data)
    If Not overrides Is Nothing Then
        Dim key As Variant
        For Each key In overrides.Keys
            Dim parts() As String
            parts = Split(CStr(key), "|")

            If UBound(parts) >= 2 Then
                Dim oSection As String
                oSection = parts(0)

                Dim oParam As String
                oParam = parts(1)

                Dim oEntity As Long
                oEntity = CLng(parts(2))

                ' Find the row for this parameter
                Dim oRow As Long
                oRow = 0
                Dim opi As Long
                For opi = 1 To paramCount
                    If StrComp(KernelConfig.GetInputSection(opi), oSection, vbTextCompare) = 0 And _
                       StrComp(KernelConfig.GetInputParam(opi), oParam, vbTextCompare) = 0 Then
                        oRow = KernelConfig.GetInputRow(opi)
                        Exit For
                    End If
                Next opi

                If oRow > 0 And oEntity >= 1 Then
                    wsInputs.Cells(oRow, INPUT_ENTITY_START_COL + oEntity - 1).Value = overrides(key)
                End If
            End If
        Next key
    End If

    SetupTestInputs = True
    Exit Function

ErrHandler:
    KernelConfig.LogError SEV_ERROR, "KernelTestHarness", "E-750", _
        "SetupTestInputs failed: " & Err.Description, _
        "MANUAL BYPASS: Set input values on the Inputs tab manually."
    SetupTestInputs = False
End Function


' =============================================================================
' RunComputeWithOverrides
' Convenience: SetupTestInputs(overrides) then RunProjections.
' Returns True if RunProjections completes without error.
' =============================================================================
Public Function RunComputeWithOverrides(overrides As Object) As Boolean
    On Error GoTo ErrHandler

    If Not SetupTestInputs(overrides) Then
        RunComputeWithOverrides = False
        Exit Function
    End If

    ' Run projections silently (suppress MsgBox by calling internal pipeline)
    Application.ScreenUpdating = False
    KernelEngine.RunProjectionsEx
    Application.ScreenUpdating = True

    RunComputeWithOverrides = True
    Exit Function

ErrHandler:
    Application.ScreenUpdating = True
    KernelConfig.LogError SEV_ERROR, "KernelTestHarness", "E-751", _
        "RunComputeWithOverrides failed: " & Err.Description, ""
    RunComputeWithOverrides = False
End Function


' =============================================================================
' AssertOutputMetric
' Reads the actual value from the Detail tab and compares to expected.
' Writes one row to TestResults.
' =============================================================================
Public Sub AssertOutputMetric(testName As String, metricName As String, _
                              entityIdx As Long, periodIdx As Long, _
                              expected As Double, _
                              Optional tolerance As Double = 0.000001)
    On Error GoTo ErrHandler

    Dim periodCount As Long
    periodCount = KernelConfig.GetTimeHorizon()

    Dim detailRow As Long
    detailRow = (entityIdx - 1) * periodCount + periodIdx + DETAIL_HEADER_ROW

    Dim detailCol As Long
    detailCol = KernelConfig.ColIndex(metricName)

    If detailCol < 1 Then
        KernelTests.WriteTestRow TEST_TIER_UNIT, _
            "DOM-" & NextDomSeq(), testName, _
            Format(expected, "0.000000"), "N/A", _
            TEST_FAIL, metricName & " not found in ColIndex"
        Exit Sub
    End If

    Dim wsDetail As Worksheet
    Set wsDetail = ThisWorkbook.Sheets(TAB_DETAIL)

    Dim actual As Double
    Dim rawVal As Variant
    rawVal = wsDetail.Cells(detailRow, detailCol).Value
    If IsNumeric(rawVal) Then
        actual = CDbl(rawVal)
    Else
        KernelTests.WriteTestRow TEST_TIER_UNIT, _
            "DOM-" & NextDomSeq(), testName, _
            Format(expected, "0.000000"), CStr(rawVal), _
            TEST_FAIL, metricName & " Entity " & entityIdx & " Period " & periodIdx & " -- non-numeric"
        Exit Sub
    End If

    Dim delta As Double
    delta = Abs(expected - actual)

    Dim result As String
    If delta < tolerance Then
        result = TEST_PASS
    Else
        result = TEST_FAIL
    End If

    KernelTests.WriteTestRow TEST_TIER_UNIT, _
        "DOM-" & NextDomSeq(), testName, _
        Format(expected, "0.000000"), Format(actual, "0.000000"), _
        result, _
        metricName & " Entity " & entityIdx & " Period " & periodIdx
    Exit Sub

ErrHandler:
    KernelTests.WriteTestRow TEST_TIER_UNIT, _
        "DOM-" & NextDomSeq(), testName, _
        Format(expected, "0.000000"), "ERROR", _
        TEST_FAIL, "Error: " & Err.Description
End Sub


' =============================================================================
' AssertOutputMetricCumulative
' Reads Detail values for the given entity from period 1 through periodIdx.
' Sums them. Compares sum to expected.
' =============================================================================
Public Sub AssertOutputMetricCumulative(testName As String, metricName As String, _
                                        entityIdx As Long, periodIdx As Long, _
                                        expected As Double, _
                                        Optional tolerance As Double = 0.000001)
    On Error GoTo ErrHandler

    Dim periodCount As Long
    periodCount = KernelConfig.GetTimeHorizon()

    Dim detailCol As Long
    detailCol = KernelConfig.ColIndex(metricName)

    If detailCol < 1 Then
        KernelTests.WriteTestRow TEST_TIER_UNIT, _
            "DOM-" & NextDomSeq(), testName, _
            Format(expected, "0.000000"), "N/A", _
            TEST_FAIL, metricName & " not found in ColIndex"
        Exit Sub
    End If

    Dim wsDetail As Worksheet
    Set wsDetail = ThisWorkbook.Sheets(TAB_DETAIL)

    Dim cumSum As Double
    cumSum = 0

    Dim prd As Long
    For prd = 1 To periodIdx
        Dim detailRow As Long
        detailRow = (entityIdx - 1) * periodCount + prd + DETAIL_HEADER_ROW

        Dim rawVal As Variant
        rawVal = wsDetail.Cells(detailRow, detailCol).Value
        If IsNumeric(rawVal) Then
            cumSum = cumSum + CDbl(rawVal)
        End If
    Next prd

    Dim delta As Double
    delta = Abs(expected - cumSum)

    Dim result As String
    If delta < tolerance Then
        result = TEST_PASS
    Else
        result = TEST_FAIL
    End If

    KernelTests.WriteTestRow TEST_TIER_UNIT, _
        "DOM-" & NextDomSeq(), testName, _
        Format(expected, "0.000000"), Format(cumSum, "0.000000"), _
        result, _
        metricName & " Entity " & entityIdx & " Cumulative(1.." & periodIdx & ")"
    Exit Sub

ErrHandler:
    KernelTests.WriteTestRow TEST_TIER_UNIT, _
        "DOM-" & NextDomSeq(), testName, _
        Format(expected, "0.000000"), "ERROR", _
        TEST_FAIL, "Error: " & Err.Description
End Sub


' =============================================================================
' RunWithSeed
' Sets PRNG seed, runs computation with overrides, returns success.
' =============================================================================
Public Function RunWithSeed(seed As Long, overrides As Object) As Boolean
    KernelRandom.InitSeed seed
    RunWithSeed = RunComputeWithOverrides(overrides)
End Function


' =============================================================================
' AssertEqual
' Generic equality assertion. Writes to TestResults.
' =============================================================================
Public Sub AssertEqual(testName As String, expected As Variant, actual As Variant, _
                       Optional detail As String = "")
    Dim result As String
    Dim expStr As String
    Dim actStr As String

    If IsNumeric(expected) And IsNumeric(actual) Then
        expStr = Format(CDbl(expected), "0.000000")
        actStr = Format(CDbl(actual), "0.000000")
        If Abs(CDbl(expected) - CDbl(actual)) < TEST_DEFAULT_TOLERANCE Then
            result = TEST_PASS
        Else
            result = TEST_FAIL
        End If
    Else
        expStr = CStr(expected)
        actStr = CStr(actual)
        If StrComp(expStr, actStr, vbBinaryCompare) = 0 Then
            result = TEST_PASS
        Else
            result = TEST_FAIL
        End If
    End If

    KernelTests.WriteTestRow TEST_TIER_UNIT, _
        "DOM-" & NextDomSeq(), testName, _
        expStr, actStr, result, detail
End Sub


' =============================================================================
' AssertTrue
' Boolean assertion. PASS if condition=True, FAIL if False.
' =============================================================================
Public Sub AssertTrue(testName As String, condition As Boolean, _
                      Optional detail As String = "")
    Dim result As String
    If condition Then
        result = TEST_PASS
    Else
        result = TEST_FAIL
    End If

    KernelTests.WriteTestRow TEST_TIER_UNIT, _
        "DOM-" & NextDomSeq(), testName, _
        "TRUE", CStr(condition), result, detail
End Sub


' =============================================================================
' GetTestResultCount
' Returns how many test rows are on the TestResults sheet.
' =============================================================================
Public Function GetTestResultCount() As Long
    On Error GoTo ErrOut
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_TEST_RESULTS)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, TR_COL_RESULT).End(xlUp).Row
    If lastRow < TR_DATA_START_ROW Then
        GetTestResultCount = 0
    Else
        GetTestResultCount = lastRow - TR_DATA_START_ROW + 1
    End If
    Exit Function
ErrOut:
    GetTestResultCount = 0
End Function


' =============================================================================
' GetTestFailCount
' Returns how many FAIL rows are on the TestResults sheet.
' =============================================================================
Public Function GetTestFailCount() As Long
    On Error GoTo ErrOut
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_TEST_RESULTS)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, TR_COL_RESULT).End(xlUp).Row
    If lastRow < TR_DATA_START_ROW Then
        GetTestFailCount = 0
        Exit Function
    End If

    Dim cnt As Long
    cnt = 0
    Dim r As Long
    For r = TR_DATA_START_ROW To lastRow
        If CStr(ws.Cells(r, TR_COL_RESULT).Value) = TEST_FAIL Then
            cnt = cnt + 1
        End If
    Next r
    GetTestFailCount = cnt
    Exit Function
ErrOut:
    GetTestFailCount = 0
End Function


' =============================================================================
' NextDomSeq
' Returns the next sequential domain test ID number.
' =============================================================================
Private Function NextDomSeq() As String
    m_domTestSeq = m_domTestSeq + 1
    NextDomSeq = Format(m_domTestSeq, "000")
End Function
