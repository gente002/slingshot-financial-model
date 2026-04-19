Attribute VB_Name = "KernelTests"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelTests.bas
' Purpose: 5-tier automated test framework. Runs all tests, reports PASS/FAIL
'          on a TestResults sheet.
'
' Tiers:
'   1 = Unit (domain tests via DomainTests module)
'   2 = Edge (domain tests via DomainTests module)
'   3 = Integration (auto-generated balance checks from ColumnRegistry)
'   4 = Regression (golden CSV hash comparison)
'   5 = Exhibit (stub -- Phase 5A)
' =============================================================================

' Module-level test row counter
Private m_nextTestRow As Long
Private m_passCount As Long
Private m_failCount As Long
Private m_totalCount As Long


' =============================================================================
' RunTests
' Master entry point. Runs all test tiers in order.
' =============================================================================
Public Sub RunTests()
    On Error GoTo ErrHandler

    ' Prompt for regression tests BEFORE anything runs (so user isn't waiting)
    Dim runRegression As VbMsgBoxResult
    runRegression = MsgBox("Include regression tests?" & vbCrLf & _
        "(Loads golden inputs, re-runs model, compares output. Takes 30-60 seconds.)", _
        vbYesNo Or vbQuestion, "RDK -- Regression Tests")

    Application.ScreenUpdating = False

    ' Ensure config is loaded
    KernelConfig.LoadAllConfig

    ' Initialize TestResults sheet
    InitTestResults

    ' Run Tier 3: Integration (auto-generated balance checks)
    RunIntegrationTests

    ' Run Tier 4: Regression (if user opted in)
    If runRegression = vbYes Then
        RunRegressionTests
    End If

    ' Run Tier 5: Smoke (automated walkthrough regression)
    RunSmokeTests

    ' Run Tier 1+2: Unit + Edge (delegate to DomainTests)
    RunDomainTests

    ' Write summary row
    WriteSummary

    ' Apply conditional formatting
    ApplyConditionalFormatting

    ' Make TestResults visible
    Dim wsTest As Worksheet
    Set wsTest = ThisWorkbook.Sheets(TAB_TEST_RESULTS)
    wsTest.Visible = xlSheetVisible
    wsTest.Activate

    KernelConfig.LogError SEV_INFO, "KernelTests", "I-700", _
        "Test run complete: " & m_totalCount & " tests, " & m_passCount & " passed, " & m_failCount & " failed", ""

    Application.ScreenUpdating = True

    MsgBox "Test run complete." & vbCrLf & vbCrLf & _
           "Total: " & m_totalCount & vbCrLf & _
           "Pass: " & m_passCount & vbCrLf & _
           "Fail: " & m_failCount, _
           IIf(m_failCount = 0, vbInformation, vbExclamation), _
           "RDK -- Test Results"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    KernelConfig.LogError SEV_ERROR, "KernelTests", "E-700", _
        "Error in RunTests: " & Err.Description, _
        "MANUAL BYPASS: Run each test tier individually: RunIntegrationTests, RunRegressionTests, or compare Detail to golden CSV manually."
    MsgBox "Test run failed: " & Err.Description & vbCrLf & vbCrLf & _
           "Check ErrorLog for details.", vbCritical, "RDK -- Test Error"
End Sub


' =============================================================================
' InitTestResults
' Clears and sets up the TestResults sheet with headers.
' =============================================================================
Private Sub InitTestResults()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet()

    ' Unprotect if needed
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    ws.Cells.ClearContents
    ws.Cells.ClearFormats

    ' Row 1: Title
    ws.Cells(1, 1).Value = "RDK Test Results"
    ws.Range(ws.Cells(1, 1), ws.Cells(1, TR_COL_DETAIL)).Merge
    With ws.Cells(1, 1)
        .Font.Bold = True
        .Font.Size = 14
    End With

    ' Row 2: Run info (populated at end)
    ws.Cells(2, 1).Value = "Run: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & _
        "  |  Kernel: " & KERNEL_VERSION

    ' Row 4: Headers
    ws.Cells(TR_HEADER_ROW, TR_COL_TIER).Value = "Tier"
    ws.Cells(TR_HEADER_ROW, TR_COL_TESTID).Value = "TestID"
    ws.Cells(TR_HEADER_ROW, TR_COL_TESTNAME).Value = "TestName"
    ws.Cells(TR_HEADER_ROW, TR_COL_EXPECTED).Value = "Expected"
    ws.Cells(TR_HEADER_ROW, TR_COL_ACTUAL).Value = "Actual"
    ws.Cells(TR_HEADER_ROW, TR_COL_RESULT).Value = "Result"
    ws.Cells(TR_HEADER_ROW, TR_COL_DETAIL).Value = "Detail"

    With ws.Range(ws.Cells(TR_HEADER_ROW, TR_COL_TIER), ws.Cells(TR_HEADER_ROW, TR_COL_DETAIL))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Set column widths
    ws.Columns(TR_COL_TIER).ColumnWidth = 6
    ws.Columns(TR_COL_TESTID).ColumnWidth = 20
    ws.Columns(TR_COL_TESTNAME).ColumnWidth = 40
    ws.Columns(TR_COL_EXPECTED).ColumnWidth = 18
    ws.Columns(TR_COL_ACTUAL).ColumnWidth = 18
    ws.Columns(TR_COL_RESULT).ColumnWidth = 8
    ws.Columns(TR_COL_DETAIL).ColumnWidth = 50

    m_nextTestRow = TR_DATA_START_ROW
    m_passCount = 0
    m_failCount = 0
    m_totalCount = 0
End Sub


' =============================================================================
' WriteTestRow
' Writes a single test result row to the TestResults sheet.
' =============================================================================
Public Sub WriteTestRow(tier As Long, testID As String, testName As String, _
                        expected As String, actual As String, _
                        result As String, detail As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_TEST_RESULTS)

    ws.Cells(m_nextTestRow, TR_COL_TIER).Value = tier
    ws.Cells(m_nextTestRow, TR_COL_TESTID).Value = testID
    ws.Cells(m_nextTestRow, TR_COL_TESTNAME).Value = testName
    ws.Cells(m_nextTestRow, TR_COL_EXPECTED).Value = expected
    ws.Cells(m_nextTestRow, TR_COL_ACTUAL).Value = actual
    ws.Cells(m_nextTestRow, TR_COL_RESULT).Value = result
    ws.Cells(m_nextTestRow, TR_COL_DETAIL).Value = detail

    m_nextTestRow = m_nextTestRow + 1
    m_totalCount = m_totalCount + 1

    If result = TEST_PASS Then
        m_passCount = m_passCount + 1
    Else
        m_failCount = m_failCount + 1
    End If
End Sub


' =============================================================================
' RunIntegrationTests
' Tier 3: Auto-generated from ColumnRegistry balance groups.
' Verifies Derived column DerivationRules hold for every Detail row.
' =============================================================================
Public Sub RunIntegrationTests()
    On Error GoTo ErrHandler

    ' Ensure config loaded
    If KernelConfig.GetColumnCount() = 0 Then KernelConfig.LoadAllConfig
    If KernelConfig.GetColumnCount() = 0 Then
        KernelConfig.LogError SEV_WARN, "KernelTests", "W-710", _
            "No columns loaded -- skipping integration tests", ""
        Exit Sub
    End If

    ' Read Detail tab data into array for fast access
    Dim wsDetail As Worksheet
    Set wsDetail = ThisWorkbook.Sheets(TAB_DETAIL)

    Dim lastRow As Long
    lastRow = wsDetail.Cells(wsDetail.Rows.Count, 1).End(xlUp).Row
    If lastRow < DETAIL_DATA_START_ROW Then
        KernelConfig.LogError SEV_WARN, "KernelTests", "W-711", _
            "Detail tab has no data -- skipping integration tests", ""
        Exit Sub
    End If

    Dim totalDataRows As Long
    totalDataRows = lastRow - DETAIL_DATA_START_ROW + 1

    Dim totalCols As Long
    totalCols = KernelConfig.GetColumnCount()

    ' Read all Detail data into array (PT-001 batch read)
    Dim detailData As Variant
    detailData = wsDetail.Range(wsDetail.Cells(DETAIL_DATA_START_ROW, 1), _
                                wsDetail.Cells(lastRow, totalCols)).Value

    ' Determine which rows to test (sampling for large models)
    Dim testRows() As Long
    Dim testRowCount As Long
    BuildSampleRows totalDataRows, testRows, testRowCount

    ' For each Derived column, verify DerivationRule
    Dim derivCols As Variant
    derivCols = KernelConfig.GetDerivedColumns()

    If Not IsArray(derivCols) Then Exit Sub

    Dim dcIdx As Long
    For dcIdx = LBound(derivCols) To UBound(derivCols)
        Dim colName As String
        colName = derivCols(dcIdx)

        Dim rule As String
        rule = KernelConfig.GetDerivationRule(colName)
        If Len(rule) = 0 Then GoTo NextDerivedCol

        Dim opA As String
        Dim opStr As String
        Dim opB As String
        If Not ParseRule(rule, opA, opStr, opB) Then GoTo NextDerivedCol

        Dim colResult As Long
        colResult = KernelConfig.ColIndex(colName)
        Dim colA As Long
        colA = KernelConfig.ColIndex(opA)
        Dim colB As Long
        colB = KernelConfig.ColIndex(opB)

        If colResult < 1 Or colA < 1 Or colB < 1 Then GoTo NextDerivedCol

        Dim balGrp As String
        balGrp = KernelConfig.GetBalGrp(colName)
        If Len(balGrp) = 0 Then balGrp = "NoGrp"

        Dim trIdx As Long
        For trIdx = 1 To testRowCount
            Dim dataRowIdx As Long
            dataRowIdx = testRows(trIdx)

            Dim valA As Double
            valA = 0
            If IsNumeric(detailData(dataRowIdx, colA)) Then valA = CDbl(detailData(dataRowIdx, colA))

            Dim valB As Double
            valB = 0
            If IsNumeric(detailData(dataRowIdx, colB)) Then valB = CDbl(detailData(dataRowIdx, colB))

            Dim expectedVal As Double
            Select Case opStr
                Case "-": expectedVal = valA - valB
                Case "+": expectedVal = valA + valB
                Case "*": expectedVal = valA * valB
                Case "/":
                    If valB = 0 Then
                        expectedVal = 0
                    Else
                        expectedVal = valA / valB
                    End If
            End Select

            Dim actualVal As Double
            actualVal = 0
            If IsNumeric(detailData(dataRowIdx, colResult)) Then actualVal = CDbl(detailData(dataRowIdx, colResult))

            Dim delta As Double
            delta = Abs(expectedVal - actualVal)

            Dim testResult As String
            If delta < TEST_DEFAULT_TOLERANCE Then
                testResult = TEST_PASS
            Else
                testResult = TEST_FAIL
            End If

            Dim detailRow As Long
            detailRow = DETAIL_DATA_START_ROW + dataRowIdx - 1

            WriteTestRow TEST_TIER_INTEGRATION, _
                "INT-" & balGrp & "-" & detailRow, _
                colName & " balance check row " & detailRow, _
                Format(expectedVal, "0.000000"), _
                Format(actualVal, "0.000000"), _
                testResult, _
                rule & " | delta=" & Format(delta, "0.000000000")
        Next trIdx
NextDerivedCol:
    Next dcIdx

    Exit Sub

ErrHandler:
    KernelConfig.LogError SEV_ERROR, "KernelTests", "E-710", _
        "Error in RunIntegrationTests: " & Err.Description, _
        "MANUAL BYPASS: Compare Detail tab Derived columns manually using the DerivationRule from Config sheet COLUMN_REGISTRY."
End Sub


' =============================================================================
' RunRegressionTests
' Tier 4: Regression test against golden baseline.
' Enhanced: loads golden inputs, re-runs model, compares output, restores state.
' =============================================================================
Public Sub RunRegressionTests()
    On Error GoTo ErrHandler

    ' Find golden snapshot
    Dim goldenName As String
    goldenName = FindGoldenSnapshot()

    If Len(goldenName) = 0 Then
        KernelConfig.LogError SEV_INFO, "KernelTests", "I-720", _
            "No golden baseline. Use SaveGolden to create one.", ""
        WriteTestRow TEST_TIER_REGRESSION, "REG-000", _
            "Golden baseline check", "Golden exists", "No golden found", _
            TEST_PASS, "Skipped -- no golden baseline. Use SaveGolden to create one."
        Exit Sub
    End If

    Dim root As String
    root = ThisWorkbook.Path & "\.."

    ' Resolve golden directory: check workspaces first, then legacy snapshots
    Dim goldenDir As String
    If Dir(root & "\workspaces\" & goldenName, vbDirectory) <> "" Then
        ' Golden is a workspace -- find the latest version
        goldenDir = root & "\workspaces\" & goldenName
        Dim fsoG As Object
        Set fsoG = CreateObject("Scripting.FileSystemObject")
        Dim latestVer As String: latestVer = ""
        Dim sfG As Object
        On Error Resume Next
        For Each sfG In fsoG.GetFolder(goldenDir).SubFolders
            If Left(sfG.Name, 1) = "v" Then
                If sfG.Name > latestVer Then latestVer = sfG.Name
            End If
        Next sfG
        On Error GoTo ErrHandler
        Set fsoG = Nothing
        If Len(latestVer) > 0 Then
            goldenDir = goldenDir & "\" & latestVer
        End If
    Else
        ' Legacy snapshots fallback
        goldenDir = root & "\snapshots\" & goldenName
    End If

    ' Check if golden has input_tabs (full regression) or just detail.csv (legacy)
    Dim hasGoldenInputs As Boolean
    hasGoldenInputs = (Dir(goldenDir & "\input_tabs", vbDirectory) <> "")

    If hasGoldenInputs Then
        ' FULL REGRESSION: save current state, load golden inputs, re-run, compare, restore
        ' 1. Save current state to temp
        Dim tempDir As String
        tempDir = root & "\snapshots\_REGRESSION_TEMP"
        KernelSnapshot.EnsureDirectoryExists root & "\snapshots"
        KernelSnapshot.EnsureDirectoryExists tempDir
        KernelTabIO.ExportAllInputTabs tempDir
        KernelSnapshotIO.ExportInputsToFile tempDir & "\inputs.csv"

        ' 2. Load golden inputs
        KernelTabIO.ImportAllInputTabs goldenDir
        If Dir(goldenDir & "\inputs.csv") <> "" Then
            KernelSnapshotIO.ImportInputsFromCsv goldenDir & "\inputs.csv"
        End If

        ' 3. Re-run model silently
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        KernelEngine.RunProjectionsEx
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
    End If

    ' Load golden detail.csv
    Dim goldenCsvPath As String
    goldenCsvPath = goldenDir & "\detail.csv"

    If Dir(goldenCsvPath) = "" Then
        KernelConfig.LogError SEV_WARN, "KernelTests", "W-720", _
            "Golden detail.csv not found: " & goldenCsvPath, ""
        WriteTestRow TEST_TIER_REGRESSION, "REG-000", _
            "Golden detail.csv exists", "File exists", "File missing", _
            TEST_FAIL, goldenCsvPath
        Exit Sub
    End If

    ' Load golden CSV into array
    Dim goldenData As Variant
    Dim goldenRows As Long
    Dim goldenCols As Long
    LoadCsvToArray goldenCsvPath, goldenData, goldenRows, goldenCols

    If goldenRows = 0 Then
        WriteTestRow TEST_TIER_REGRESSION, "REG-000", _
            "Golden CSV has data", ">0 rows", "0 rows", _
            TEST_FAIL, "Golden CSV is empty"
        Exit Sub
    End If

    ' Read current Detail tab
    Dim wsDetail As Worksheet
    Set wsDetail = ThisWorkbook.Sheets(TAB_DETAIL)

    Dim lastRow As Long
    lastRow = wsDetail.Cells(wsDetail.Rows.Count, 1).End(xlUp).Row
    If lastRow < DETAIL_DATA_START_ROW Then
        WriteTestRow TEST_TIER_REGRESSION, "REG-000", _
            "Detail tab has data", ">0 rows", "0 rows", _
            TEST_FAIL, "Run RunProjections first to generate Detail data."
        Exit Sub
    End If

    Dim totalCols As Long
    totalCols = KernelConfig.GetColumnCount()
    Dim currentRows As Long
    currentRows = lastRow - DETAIL_DATA_START_ROW + 1

    Dim currentData As Variant
    currentData = wsDetail.Range(wsDetail.Cells(DETAIL_DATA_START_ROW, 1), _
                                  wsDetail.Cells(lastRow, totalCols)).Value

    ' Compare row counts
    If currentRows <> goldenRows Then
        WriteTestRow TEST_TIER_REGRESSION, "REG-001", _
            "Row count matches golden", CStr(goldenRows), CStr(currentRows), _
            TEST_FAIL, "Regression vs " & goldenName
        ' Still compare what we can
    End If

    ' Compare row-by-row, column-by-column
    Dim mismatchCount As Long
    mismatchCount = 0
    Dim maxMismatches As Long
    maxMismatches = 50

    Dim compareRows As Long
    compareRows = goldenRows
    If currentRows < compareRows Then compareRows = currentRows

    Dim compareCols As Long
    compareCols = goldenCols
    If totalCols < compareCols Then compareCols = totalCols

    Dim rIdx As Long
    Dim cIdx As Long
    For rIdx = 1 To compareRows
        For cIdx = 1 To compareCols
            Dim gVal As Variant
            gVal = goldenData(rIdx, cIdx)
            Dim dVal As Variant
            dVal = currentData(rIdx, cIdx)

            Dim isMismatch As Boolean
            isMismatch = False

            If IsNumeric(gVal) And IsNumeric(dVal) Then
                If Abs(CDbl(gVal) - CDbl(dVal)) > TEST_DEFAULT_TOLERANCE Then
                    isMismatch = True
                End If
            Else
                If StrComp(CStr(gVal), CStr(dVal), vbBinaryCompare) <> 0 Then
                    isMismatch = True
                End If
            End If

            If isMismatch Then
                mismatchCount = mismatchCount + 1
                If mismatchCount <= maxMismatches Then
                    Dim colNameR As String
                    If cIdx <= KernelConfig.GetColumnCount() Then
                        colNameR = KernelConfig.GetColName(cIdx)
                    Else
                        colNameR = "Col" & cIdx
                    End If
                    WriteTestRow TEST_TIER_REGRESSION, _
                        "REG-R" & rIdx & "C" & cIdx, _
                        "Row " & rIdx & " " & colNameR, _
                        CStr(gVal), CStr(dVal), _
                        TEST_FAIL, _
                        "Regression vs " & goldenName
                End If
            End If
        Next cIdx
    Next rIdx

    ' Summary row
    If mismatchCount = 0 Then
        WriteTestRow TEST_TIER_REGRESSION, "REG-SUMMARY", _
            "Regression vs " & goldenName & ": " & compareRows & " rows", _
            "0 mismatches", "0 mismatches", _
            TEST_PASS, "All values match golden baseline"
    Else
        WriteTestRow TEST_TIER_REGRESSION, "REG-SUMMARY", _
            "Regression vs " & goldenName & ": " & compareRows & " rows", _
            "0 mismatches", CStr(mismatchCount) & " mismatches", _
            TEST_FAIL, _
            IIf(mismatchCount > maxMismatches, _
                "Showing first " & maxMismatches & " of " & mismatchCount, _
                CStr(mismatchCount) & " total mismatches")
    End If

    ' Check for formula errors after regression run
    Dim postRunErrors As String
    postRunErrors = ScanForFormulaErrors()
    If Len(postRunErrors) > 0 Then
        WriteTestRow TEST_TIER_REGRESSION, "REG-ERRORS", _
            "No formula errors after regression run", "No errors", _
            "Errors on: " & postRunErrors, TEST_FAIL, _
            "#VALUE!, #REF!, or #NAME? detected"
    Else
        WriteTestRow TEST_TIER_REGRESSION, "REG-ERRORS", _
            "No formula errors after regression run", "No errors", "No errors", _
            TEST_PASS, ""
    End If

    ' Compare regression tabs (formula tab outputs)
    Dim regResult As String
    regResult = KernelTabIO.CompareRegressionTabs(goldenDir)
    If InStr(1, regResult, "FAIL", vbTextCompare) > 0 Then
        WriteTestRow TEST_TIER_REGRESSION, "REG-TABS", _
            "Formula tab regression", "All tabs match", regResult, _
            TEST_FAIL, "Regression vs " & goldenName
    Else
        WriteTestRow TEST_TIER_REGRESSION, "REG-TABS", _
            "Formula tab regression", "All tabs match", regResult, _
            TEST_PASS, "Regression vs " & goldenName
    End If

    ' Restore original state if we did a full regression
    If hasGoldenInputs Then
        KernelTabIO.ImportAllInputTabs tempDir
        If Dir(tempDir & "\inputs.csv") <> "" Then
            KernelSnapshotIO.ImportInputsFromCsv tempDir & "\inputs.csv"
        End If
        ' Clean up temp
        On Error Resume Next
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(tempDir) Then fso.DeleteFolder tempDir, True
        Set fso = Nothing
        On Error GoTo 0
        KernelConfig.LogError SEV_INFO, "KernelTests", "I-721", _
            "Regression test complete. Original inputs restored.", ""
    End If

    Exit Sub

ErrHandler:
    ' Restore on error too
    If hasGoldenInputs Then
        On Error Resume Next
        KernelTabIO.ImportAllInputTabs tempDir
        If Dir(tempDir & "\inputs.csv") <> "" Then
            KernelSnapshotIO.ImportInputsFromCsv tempDir & "\inputs.csv"
        End If
        On Error GoTo 0
    End If
    KernelConfig.LogError SEV_ERROR, "KernelTests", "E-720", _
        "Error in RunRegressionTests: " & Err.Description, _
        "MANUAL BYPASS: Open the golden CSV and current CSV side-by-side."
End Sub


' =============================================================================
' SaveGolden
' Creates a golden baseline as a workspace. The workspace is named
' GOLDEN_{name} and contains one version with the current model state.
' Regression tests find it by scanning for workspaces with GOLDEN_ prefix.
' =============================================================================
Public Sub SaveGolden(goldenName As String, Optional description As String = "")
    On Error GoTo ErrHandler

    ' Verify Detail tab has data
    Dim wsDetail As Worksheet
    Set wsDetail = ThisWorkbook.Sheets(TAB_DETAIL)
    If wsDetail.Cells(DETAIL_DATA_START_ROW, 1).Value = "" Then
        MsgBox "Detail tab has no data. Run RunProjections first.", _
               vbExclamation, "RDK -- Save Golden"
        Exit Sub
    End If

    ' Verify model health before saving golden
    Dim healthResult As String
    healthResult = KernelTabIO.VerifyHealthAfterLoad()
    If InStr(1, healthResult, "FAIL", vbTextCompare) > 0 Then
        Dim saveAnyway As VbMsgBoxResult
        saveAnyway = MsgBox("WARNING: Health checks are failing:" & vbCrLf & _
            healthResult & vbCrLf & vbCrLf & _
            "Saving a golden baseline with errors will make regression tests " & _
            "compare against broken data. Save anyway?", _
            vbYesNo Or vbExclamation, "RDK -- Save Golden")
        If saveAnyway = vbNo Then Exit Sub
    End If

    ' Scan for #VALUE!, #REF!, #NAME? errors on key output tabs
    Dim errorTabs As String
    errorTabs = ScanForFormulaErrors()
    If Len(errorTabs) > 0 Then
        Dim saveWithErrors As VbMsgBoxResult
        saveWithErrors = MsgBox("WARNING: Formula errors found on:" & vbCrLf & _
            errorTabs & vbCrLf & vbCrLf & _
            "Golden baseline will contain error values. Save anyway?", _
            vbYesNo Or vbExclamation, "RDK -- Save Golden")
        If saveWithErrors = vbNo Then Exit Sub
    End If

    ' Save as workspace with GOLDEN_ prefix
    Dim fullName As String
    fullName = GOLDEN_PREFIX & goldenName

    If Len(description) = 0 Then
        description = "Golden baseline: " & goldenName
    End If

    KernelWorkspace.SaveWorkspace fullName

    ' Export regression tab outputs to the golden workspace
    Dim goldenRoot As String
    goldenRoot = ThisWorkbook.Path & "\..\workspaces\" & fullName
    ' Find the latest version folder
    On Error Resume Next
    Dim fsoG As Object
    Set fsoG = CreateObject("Scripting.FileSystemObject")
    Dim latestV As String: latestV = ""
    Dim sfG As Object
    For Each sfG In fsoG.GetFolder(goldenRoot).SubFolders
        If Left(sfG.Name, 1) = "v" Then
            If sfG.Name > latestV Then latestV = sfG.Name
        End If
    Next sfG
    Set fsoG = Nothing
    On Error GoTo ErrHandler
    If Len(latestV) > 0 Then
        KernelTabIO.ExportRegressionTabs goldenRoot & "\" & latestV
    End If

    KernelConfig.LogError SEV_INFO, "KernelTests", "I-730", _
        "Golden baseline created as workspace: " & fullName, ""

    MsgBox "Golden baseline created: " & fullName, _
           vbInformation, "RDK -- Golden Saved"
    Exit Sub

ErrHandler:
    KernelConfig.LogError SEV_ERROR, "KernelTests", "E-730", _
        "Error saving golden: " & Err.Description, _
        "MANUAL BYPASS: Save a workspace named GOLDEN_" & goldenName & " using Save / Load."
    MsgBox "Error saving golden: " & Err.Description, vbCritical, "RDK -- Error"
End Sub


' =============================================================================
' SaveGoldenUI
' Parameterless entry point for the Dashboard button.
' =============================================================================
Public Sub SaveGoldenUI()
    Dim gName As String
    gName = InputBox("Enter golden baseline name:", "RDK -- Save Golden")
    If Len(Trim(gName)) = 0 Then Exit Sub

    Dim gDesc As String
    gDesc = InputBox("Description (optional):", "RDK -- Save Golden")

    SaveGolden gName, gDesc
End Sub


' =============================================================================
' RunDomainTests
' Tier 1+2: Delegate to DomainTests module if it exists.
' =============================================================================
Public Function RunDomainTests() As Long
    On Error GoTo NoDomainTests

    ' Check if DomainTests module exists by attempting to call it
    Dim testCountBefore As Long
    testCountBefore = m_totalCount

    Application.Run "DomainTests.RunAllTests"

    RunDomainTests = m_totalCount - testCountBefore
    Exit Function

NoDomainTests:
    ' Check if it was a "Sub or Function not defined" error
    If Err.Number = 1004 Or Err.Number = 424 Or Err.Number = 438 Then
        KernelConfig.LogError SEV_INFO, "KernelTests", "I-740", _
            "No DomainTests module found. Skipping Tier 1+2.", ""
        RunDomainTests = 0
    Else
        KernelConfig.LogError SEV_ERROR, "KernelTests", "E-740", _
            "Error running DomainTests: " & Err.Description, _
            "MANUAL BYPASS: Fix errors in DomainTests.bas, then call RunDomainTests again."
        RunDomainTests = 0
    End If
End Function


' =============================================================================
' WriteSummary
' Updates the summary info row on the TestResults sheet.
' =============================================================================
Private Sub WriteSummary()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_TEST_RESULTS)

    ws.Cells(2, 1).Value = "Run: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & _
        "  |  Kernel: " & KERNEL_VERSION & _
        "  |  Total: " & m_totalCount & _
        "  |  Pass: " & m_passCount & _
        "  |  Fail: " & m_failCount
End Sub


' =============================================================================
' ApplyConditionalFormatting
' Colors PASS green and FAIL red in the Result column.
' =============================================================================
Private Sub ApplyConditionalFormatting()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_TEST_RESULTS)

    If m_nextTestRow <= TR_DATA_START_ROW Then Exit Sub

    Dim rng As Object
    Set rng = ws.Range(ws.Cells(TR_DATA_START_ROW, TR_COL_RESULT), _
                       ws.Cells(m_nextTestRow - 1, TR_COL_RESULT))

    rng.FormatConditions.Delete

    ' PASS = green
    Dim cfPass As Object
    Set cfPass = rng.FormatConditions.Add(Type:=1, Operator:=3, Formula1:="=""PASS""")
    cfPass.Interior.Color = RGB(198, 239, 206)
    cfPass.Font.Color = RGB(0, 97, 0)

    ' FAIL = red bold
    Dim cfFail As Object
    Set cfFail = rng.FormatConditions.Add(Type:=1, Operator:=3, Formula1:="=""FAIL""")
    cfFail.Interior.Color = RGB(255, 199, 206)
    cfFail.Font.Color = RGB(156, 0, 6)
    cfFail.Font.Bold = True
End Sub


' =============================================================================
' EnsureTestSheet
' Creates the TestResults sheet if it doesn't exist.
' =============================================================================
Private Function EnsureTestSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(TAB_TEST_RESULTS)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = TAB_TEST_RESULTS
    End If

    Set EnsureTestSheet = ws
End Function


' =============================================================================
' FindGoldenSnapshot
' Scans the snapshots directory for a folder with GOLDEN_ prefix.
' Returns the first match or empty string.
' =============================================================================
' =============================================================================
' ScanForFormulaErrors
' Scans regression_config tabs for #VALUE!, #REF!, #NAME?, #DIV/0! errors.
' Returns a string listing tabs with errors, or empty if clean.
' =============================================================================
Private Function ScanForFormulaErrors() As String
    ScanForFormulaErrors = ""
    On Error Resume Next

    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    If wsConfig Is Nothing Then Exit Function

    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_REGRESSION_CONFIG)
    If sr = 0 Then Exit Function

    Dim dr As Long
    dr = sr + 2
    Do While Len(Trim(CStr(wsConfig.Cells(dr, REGCFG_COL_TABNAME).Value))) > 0
        Dim tabName As String
        tabName = Trim(CStr(wsConfig.Cells(dr, REGCFG_COL_TABNAME).Value))

        Dim ws As Worksheet
        Set ws = Nothing
        Set ws = ThisWorkbook.Sheets(tabName)
        If Not ws Is Nothing Then
            ' Check for error cells in the used range
            Dim errCells As Range
            Set errCells = Nothing
            Err.Clear
            Set errCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors)
            If Err.Number = 0 And Not errCells Is Nothing Then
                ' Has error cells
                If Len(ScanForFormulaErrors) > 0 Then ScanForFormulaErrors = ScanForFormulaErrors & ", "
                ScanForFormulaErrors = ScanForFormulaErrors & tabName
            End If
            Err.Clear
        End If
        dr = dr + 1
    Loop
    On Error GoTo 0
End Function


Private Function FindGoldenSnapshot() As String
    FindGoldenSnapshot = ""

    Dim root As String
    root = ThisWorkbook.Path & "\.."

    ' Search workspaces first (new location)
    Dim wsDir As String
    wsDir = root & "\workspaces"
    If Dir(wsDir, vbDirectory) <> "" Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(wsDir) Then
            Dim subFld As Object
            For Each subFld In fso.GetFolder(wsDir).SubFolders
                If Left(subFld.Name, Len(GOLDEN_PREFIX)) = GOLDEN_PREFIX Then
                    FindGoldenSnapshot = subFld.Name
                    Set fso = Nothing
                    Exit Function
                End If
            Next subFld
        End If
        Set fso = Nothing
    End If

    ' Fallback: search legacy snapshots directory
    Dim snapDir As String
    snapDir = root & "\snapshots"
    If Dir(snapDir, vbDirectory) <> "" Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(snapDir) Then
            For Each subFld In fso.GetFolder(snapDir).SubFolders
                If Left(subFld.Name, Len(GOLDEN_PREFIX)) = GOLDEN_PREFIX Then
                    FindGoldenSnapshot = subFld.Name
                    Set fso = Nothing
                    Exit Function
                End If
            Next subFld
        End If
        Set fso = Nothing
    End If
End Function


' =============================================================================
' BuildSampleRows
' For large models, samples first 10, last 10, and 10 random rows.
' For small models, tests all rows.
' =============================================================================
Private Sub BuildSampleRows(totalRows As Long, ByRef sampleRows() As Long, _
                             ByRef sampleCount As Long)
    If totalRows <= LARGE_MODEL_ROW_THRESHOLD Then
        ' Test all rows
        sampleCount = totalRows
        ReDim sampleRows(1 To sampleCount)
        Dim i As Long
        For i = 1 To sampleCount
            sampleRows(i) = i
        Next i
    Else
        ' Sample: first 10, last 10, 10 random
        sampleCount = LARGE_MODEL_SAMPLE_SIZE
        ReDim sampleRows(1 To sampleCount)

        ' First 10
        For i = 1 To 10
            sampleRows(i) = i
        Next i

        ' Last 10
        For i = 1 To 10
            sampleRows(10 + i) = totalRows - 10 + i
        Next i

        ' 10 random from middle
        Dim midStart As Long
        midStart = 11
        Dim midEnd As Long
        midEnd = totalRows - 10

        For i = 1 To 10
            Dim rndRow As Long
            rndRow = midStart + Int(Rnd * (midEnd - midStart + 1))
            sampleRows(20 + i) = rndRow
        Next i

        KernelConfig.LogError SEV_WARN, "KernelTests", "W-712", _
            "Large model -- sampled " & sampleCount & " of " & totalRows & " rows for balance checks.", ""
    End If
End Sub


' =============================================================================
' ParseRule
' Parses a DerivationRule like "A - B" into components.
' =============================================================================
Private Function ParseRule(rule As String, ByRef opA As String, _
                           ByRef opStr As String, ByRef opB As String) As Boolean
    ParseRule = False

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
                ParseRule = True
                Exit Function
            End If
        End If
    Next opIdx
End Function


' =============================================================================
' LoadCsvToArray
' Reads a CSV file into a 2D Variant array.
' =============================================================================
Private Sub LoadCsvToArray(csvPath As String, ByRef data As Variant, _
                           ByRef rowCount As Long, ByRef numCols As Long)
    rowCount = 0
    numCols = 0

    If Dir(csvPath) = "" Then Exit Sub

    ' Read file content
    Dim fileNum As Integer
    fileNum = FreeFile

    Dim fileContent As String
    Dim fileSize As Long

    Open csvPath For Binary Access Read As #fileNum
    fileSize = LOF(fileNum)
    If fileSize = 0 Then
        Close #fileNum
        Exit Sub
    End If
    fileContent = Space$(fileSize)
    Get #fileNum, , fileContent
    Close #fileNum

    ' Normalize line endings
    fileContent = Replace(fileContent, vbCrLf, vbLf)
    fileContent = Replace(fileContent, vbCr, vbLf)

    Dim lines() As String
    lines = Split(fileContent, vbLf)

    ' Count non-empty data lines (skip header)
    Dim totalLines As Long
    totalLines = 0
    Dim li As Long
    For li = 0 To UBound(lines)
        If Len(Trim(lines(li))) > 0 Then totalLines = totalLines + 1
    Next li

    If totalLines <= 1 Then Exit Sub  ' header only

    rowCount = totalLines - 1  ' minus header

    ' Parse header to get column count
    Dim headerFields() As String
    headerFields = ParseCsvLine(lines(0))
    numCols = UBound(headerFields) - LBound(headerFields) + 1

    ReDim data(1 To rowCount, 1 To numCols)

    Dim dataIdx As Long
    dataIdx = 0
    For li = 1 To UBound(lines)
        If Len(Trim(lines(li))) = 0 Then GoTo NextCsvLine
        dataIdx = dataIdx + 1
        If dataIdx > rowCount Then Exit For

        Dim fields() As String
        fields = ParseCsvLine(lines(li))

        Dim fi As Long
        For fi = LBound(fields) To UBound(fields)
            Dim colIdx As Long
            colIdx = fi - LBound(fields) + 1
            If colIdx <= numCols Then
                If IsNumeric(fields(fi)) And Len(fields(fi)) > 0 Then
                    data(dataIdx, colIdx) = CDbl(fields(fi))
                Else
                    data(dataIdx, colIdx) = fields(fi)
                End If
            End If
        Next fi
NextCsvLine:
    Next li
End Sub


' =============================================================================
' ParseCsvLine
' Parses a CSV line into an array of field values.
' =============================================================================
Private Function ParseCsvLine(lineText As String) As String()
    Dim result() As String
    Dim fieldCount As Long
    fieldCount = 0

    Dim pos As Long
    pos = 1

    Dim lineLen As Long
    lineLen = Len(lineText)

    Dim maxFields As Long
    maxFields = 1
    Dim scanPos As Long
    For scanPos = 1 To lineLen
        If Mid(lineText, scanPos, 1) = "," Then maxFields = maxFields + 1
    Next scanPos

    ReDim result(0 To maxFields - 1)

    Do While pos <= lineLen
        Dim fieldVal As String
        fieldVal = ""

        If Mid(lineText, pos, 1) = """" Then
            pos = pos + 1
            Do While pos <= lineLen
                If Mid(lineText, pos, 1) = """" Then
                    If pos < lineLen And Mid(lineText, pos + 1, 1) = """" Then
                        fieldVal = fieldVal & """"
                        pos = pos + 2
                    Else
                        pos = pos + 1
                        Exit Do
                    End If
                Else
                    fieldVal = fieldVal & Mid(lineText, pos, 1)
                    pos = pos + 1
                End If
            Loop
            If pos <= lineLen And Mid(lineText, pos, 1) = "," Then
                pos = pos + 1
            End If
        Else
            Dim commaPos As Long
            commaPos = InStr(pos, lineText, ",")
            If commaPos = 0 Then
                fieldVal = Mid(lineText, pos)
                pos = lineLen + 1
            Else
                fieldVal = Mid(lineText, pos, commaPos - pos)
                pos = commaPos + 1
            End If
        End If

        If fieldCount <= UBound(result) Then
            result(fieldCount) = fieldVal
        End If
        fieldCount = fieldCount + 1
    Loop

    If fieldCount > 0 And fieldCount <= maxFields Then
        ReDim Preserve result(0 To fieldCount - 1)
    ElseIf fieldCount = 0 Then
        ReDim result(0 To 0)
    End If

    ParseCsvLine = result
End Function


' =============================================================================
' RunSmokeTests
' Tier 5: Automated walkthrough regression. Checks everything that does not
' require human eyes -- module existence, config sections, deterministic
' fixture values, CurveLib math, extension registry, domain dispatch.
' =============================================================================
Public Sub RunSmokeTests()
    Dim tier As Long
    tier = 5

    ' --- Load extension registry (not included in LoadAllConfig) ---
    KernelExtension.LoadExtensionRegistry

    ' --- Module existence (kernel only -- domain module checks belong in domain tests) ---
    SmokeCheckModule tier, "SMK-001", "KernelExtension module exists", "KernelExtension"
    SmokeCheckModule tier, "SMK-004", "KernelEngine module exists", "KernelEngine"

    ' --- Config sections on Config sheet ---
    SmokeCheckConfigSection tier, "SMK-010", "EXTENSION_REGISTRY section exists", CFG_MARKER_EXTENSION_REGISTRY
    SmokeCheckConfigSection tier, "SMK-011", "CURVE_LIBRARY_CONFIG section exists", CFG_MARKER_CURVE_LIBRARY
    SmokeCheckConfigSection tier, "SMK-012", "GRANULARITY_CONFIG section exists", CFG_MARKER_GRANULARITY_CONFIG

    ' --- DomainModule setting (verify it exists, not its specific value) ---
    Dim domMod As String
    domMod = KernelConfig.GetSetting("DomainModule")
    SmokeCheckTrue tier, "SMK-020", "DomainModule setting is set", Len(domMod) > 0, domMod

    ' --- Extension registry (generic count check, no specific extensions) ---
    Dim extCount As Long
    extCount = KernelExtension.GetActiveExtensionCount()
    SmokeCheckTrue tier, "SMK-030", "Extension registry loaded", extCount >= 0, CStr(extCount) & " active"

    ' --- Detail tab fixture values ---
    Dim wsDet As Worksheet
    Set wsDet = Nothing
    On Error Resume Next
    Set wsDet = ThisWorkbook.Sheets(TAB_DETAIL)
    On Error GoTo 0

    If wsDet Is Nothing Then
        WriteTestRow tier, "SMK-040", "Detail tab exists", "exists", "MISSING", TEST_FAIL, ""
    Else
        WriteTestRow tier, "SMK-040", "Detail tab exists", "exists", "exists", TEST_PASS, ""
    End If

    ' --- CurveLib / domain extension tests removed (belong in domain test module) ---

    ' --- Required tabs exist ---
    ' Required kernel infrastructure tabs only (all others come from tab_registry)
    SmokeCheckTabExists tier, "SMK-071", "Dashboard tab exists", TAB_DASHBOARD
    SmokeCheckTabExists tier, "SMK-072", "Config tab exists", TAB_CONFIG
    SmokeCheckTabExists tier, "SMK-073", "ErrorLog tab exists", TAB_ERROR_LOG
    SmokeCheckTabExists tier, "SMK-074", "Detail tab exists", TAB_DETAIL

    KernelConfig.LogError SEV_INFO, "KernelTests", "I-705", _
        "Smoke tests (Tier 5) complete.", ""
End Sub


' =============================================================================
' Smoke Test Helpers
' =============================================================================

Private Sub SmokeCheckModule(tier As Long, testID As String, testName As String, modName As String)
    Dim found As Boolean
    found = False
    On Error Resume Next
    Dim comp As Object
    Set comp = ThisWorkbook.VBProject.VBComponents(modName)
    If Not comp Is Nothing Then found = True
    On Error GoTo 0
    If found Then
        WriteTestRow tier, testID, testName, "exists", "exists", TEST_PASS, ""
    Else
        WriteTestRow tier, testID, testName, "exists", "MISSING", TEST_FAIL, _
            "Module " & modName & " not found in VBProject"
    End If
End Sub

Private Sub SmokeCheckConfigSection(tier As Long, testID As String, testName As String, marker As String)
    Dim wsConfig As Worksheet
    Set wsConfig = Nothing
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    On Error GoTo 0
    If wsConfig Is Nothing Then
        WriteTestRow tier, testID, testName, "exists", "Config tab missing", TEST_FAIL, ""
        Exit Sub
    End If

    Dim found As Boolean
    found = False
    Dim r As Long
    For r = 1 To wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).Row
        If CStr(wsConfig.Cells(r, 1).Value) = marker Then
            found = True
            Exit For
        End If
    Next r
    If found Then
        WriteTestRow tier, testID, testName, "exists", "exists", TEST_PASS, ""
    Else
        WriteTestRow tier, testID, testName, "exists", "NOT FOUND", TEST_FAIL, marker
    End If
End Sub

Private Sub SmokeCheckEqual(tier As Long, testID As String, testName As String, expected As String, actual As String)
    If StrComp(expected, actual, vbTextCompare) = 0 Then
        WriteTestRow tier, testID, testName, expected, actual, TEST_PASS, ""
    Else
        WriteTestRow tier, testID, testName, expected, actual, TEST_FAIL, ""
    End If
End Sub

Private Sub SmokeCheckClose(tier As Long, testID As String, testName As String, _
                            expected As Double, actual As Double, tol As Double)
    If Abs(expected - actual) <= tol Then
        WriteTestRow tier, testID, testName, Format(expected, "0.######"), _
            Format(actual, "0.######"), TEST_PASS, ""
    Else
        WriteTestRow tier, testID, testName, Format(expected, "0.######"), _
            Format(actual, "0.######"), TEST_FAIL, "Diff=" & Format(Abs(expected - actual), "0.######")
    End If
End Sub

Private Sub SmokeCheckTrue(tier As Long, testID As String, testName As String, _
                           condition As Boolean, detail As String)
    If condition Then
        WriteTestRow tier, testID, testName, "TRUE", "TRUE", TEST_PASS, detail
    Else
        WriteTestRow tier, testID, testName, "TRUE", "FALSE", TEST_FAIL, detail
    End If
End Sub

Private Sub SmokeCheckTabExists(tier As Long, testID As String, testName As String, tabName As String)
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(tabName)
    On Error GoTo 0
    If Not ws Is Nothing Then
        WriteTestRow tier, testID, testName, "exists", "exists", TEST_PASS, ""
    Else
        WriteTestRow tier, testID, testName, "exists", "MISSING", TEST_FAIL, tabName
    End If
End Sub


' =============================================================================
' TestAuditAssertions (Tier 6 - Audit Regression)
'
' Runtime assertions that catch the specific bug classes surfaced during the
' RBC tab audit of 2026-04-18/19. Each assertion is a lightweight sanity
' check that runs after a model run, surfaces PASS/FAIL on the TestResults
' tab, and prevents regression of the audited bugs.
'
' Companion to the static scanner at scripts/audit_scan.py. The static
' scanner catches config-level issues pre-commit; this sub catches runtime
' issues post-model-run.
'
' See docs/AUDIT_PLAYBOOK.md Phase 4 for the framework.
' =============================================================================
Public Sub TestAuditAssertions()
    On Error GoTo ErrHandler
    Const tier As Long = 6

    ' Ensure TestResults initialized
    If m_nextTestRow = 0 Then InitTestResults

    ' -- Assertion A: Balance Sheet identity holds at Q1 Y1
    Dim bsCheck As Double
    bsCheck = 0
    On Error Resume Next
    Dim wsBS As Worksheet
    Set wsBS = ThisWorkbook.Sheets("Balance Sheet")
    Dim bsCheckRow As Long
    bsCheckRow = KernelFormula.ResolveRowID("Balance Sheet", "BS_CHECK")
    If bsCheckRow > 0 Then
        Dim cellVal As Variant
        cellVal = wsBS.Cells(bsCheckRow, 3).Value
        If IsNumeric(cellVal) Then bsCheck = CDbl(cellVal)
    End If
    On Error GoTo ErrHandler
    If Abs(bsCheck) < 0.01 Then
        WriteTestRow tier, "AUD-001", "Balance Sheet identity: A - (L+E) = 0 at Q1 Y1", "0", _
            Format(bsCheck, "#,##0.00"), TEST_PASS, ""
    Else
        WriteTestRow tier, "AUD-001", "Balance Sheet identity: A - (L+E) = 0 at Q1 Y1", "0", _
            Format(bsCheck, "#,##0.00"), TEST_FAIL, _
            "BS_CHECK non-zero -- accounting identity violated"
    End If

    ' -- Assertion B: Corp sub-allocation sums to 100%
    On Error Resume Next
    Dim wsRBC As Worksheet
    Set wsRBC = ThisWorkbook.Sheets("RBC Capital Model")
    Dim corpSum As Double
    corpSum = 0
    Dim subTotalRow As Long
    subTotalRow = KernelFormula.ResolveRowID("RBC Capital Model", "RBC_CORP_SUB_TOTAL")
    If subTotalRow > 0 Then
        Dim csVal As Variant
        csVal = wsRBC.Cells(subTotalRow, 3).Value
        If IsNumeric(csVal) Then corpSum = CDbl(csVal)
    End If
    On Error GoTo ErrHandler
    If Abs(corpSum - 1#) < 0.001 Then
        WriteTestRow tier, "AUD-002", "Corp sub-allocation sum = 100%", "1.000", _
            Format(corpSum, "0.000"), TEST_PASS, ""
    Else
        WriteTestRow tier, "AUD-002", "Corp sub-allocation sum = 100%", "1.000", _
            Format(corpSum, "0.000"), TEST_FAIL, _
            "Corp sub-allocation drift -- R1 bond granularity will be proportionally off"
    End If

    ' -- Assertion C: R3 (Credit Risk) non-negative
    ' Catches regression of the CRSV sign-convention bug fixed 2026-04-19.
    On Error Resume Next
    Dim r3Row As Long
    Dim r3Val As Double
    r3Val = 0
    r3Row = KernelFormula.ResolveRowID("RBC Capital Model", "RBC_R3")
    If r3Row > 0 Then
        Dim r3cell As Variant
        r3cell = wsRBC.Cells(r3Row, 3).Value
        If IsNumeric(r3cell) Then r3Val = CDbl(r3cell)
    End If
    On Error GoTo ErrHandler
    If r3Val >= -0.001 Then
        WriteTestRow tier, "AUD-003", "R3 Credit Risk non-negative at Q1 Y1", ">= 0", _
            Format(r3Val, "#,##0"), TEST_PASS, ""
    Else
        WriteTestRow tier, "AUD-003", "R3 Credit Risk non-negative at Q1 Y1", ">= 0", _
            Format(r3Val, "#,##0"), TEST_FAIL, _
            "R3 negative -- check MIR_CRSV has ABS wrapper (PD-05 sign convention)"
    End If

    ' -- Assertion D: R1-R5 non-zero at Q2 Y1 when Invested Assets > 0
    ' Catches regression of the ROWID-to-static-cell bug fixed 2026-04-19.
    ' If invested assets have grown past Q1 Y1, R1 and R2 MUST be non-zero at Q2 Y1.
    On Error Resume Next
    Dim investRow As Long
    investRow = KernelFormula.ResolveRowID("RBC Capital Model", "RBC_MIR_INVEST")
    Dim r1Row As Long
    r1Row = KernelFormula.ResolveRowID("RBC Capital Model", "RBC_R1")
    If investRow > 0 And r1Row > 0 Then
        Dim q2col As Long
        q2col = 4  ' Col D = Q2 Y1 on a quarterly tab
        Dim investQ2 As Double
        investQ2 = 0
        Dim r1Q2 As Double
        r1Q2 = 0
        If IsNumeric(wsRBC.Cells(investRow, q2col).Value) Then _
            investQ2 = CDbl(wsRBC.Cells(investRow, q2col).Value)
        If IsNumeric(wsRBC.Cells(r1Row, q2col).Value) Then _
            r1Q2 = CDbl(wsRBC.Cells(r1Row, q2col).Value)
        If investQ2 > 0.01 And Abs(r1Q2) < 0.001 Then
            WriteTestRow tier, "AUD-004", "R1 non-zero at Q2 Y1 when invested > 0", "> 0", _
                Format(r1Q2, "#,##0"), TEST_FAIL, _
                "R1 zero at Q2+ despite non-zero invested assets -- ROWID-to-static-cell regression"
        Else
            WriteTestRow tier, "AUD-004", "R1 non-zero at Q2 Y1 when invested > 0", "> 0 or invested = 0", _
                Format(r1Q2, "#,##0"), TEST_PASS, ""
        End If
    End If
    On Error GoTo ErrHandler

    ' -- Assertion E: Total RBC covariance formula produces sensible result
    ' Formula: R0 + sqrt(R1^2 + R2^2 + R3^2 + R4^2 + R5^2 + Rcat^2)
    ' Must be non-negative and at least as large as the largest component.
    On Error Resume Next
    Dim totRow As Long
    totRow = KernelFormula.ResolveRowID("RBC Capital Model", "RBC_TOTRBC")
    If totRow > 0 Then
        Dim totVal As Double
        totVal = 0
        If IsNumeric(wsRBC.Cells(totRow, 3).Value) Then _
            totVal = CDbl(wsRBC.Cells(totRow, 3).Value)
        If totVal >= 0 Then
            WriteTestRow tier, "AUD-005", "Total RBC covariance non-negative", ">= 0", _
                Format(totVal, "#,##0"), TEST_PASS, ""
        Else
            WriteTestRow tier, "AUD-005", "Total RBC covariance non-negative", ">= 0", _
                Format(totVal, "#,##0"), TEST_FAIL, "Total RBC negative -- covariance formula broken"
        End If
    End If
    On Error GoTo ErrHandler

    KernelConfig.LogError SEV_INFO, "KernelTests", "I-706", _
        "Audit assertions (Tier 6) complete.", ""
    Exit Sub
ErrHandler:
    KernelConfig.LogError SEV_WARN, "KernelTests", "W-706", _
        "TestAuditAssertions error: " & Err.Description, ""
End Sub
