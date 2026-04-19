Attribute VB_Name = "KernelEngine"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelEngine.bas
' Purpose: Computation lifecycle orchestration. Calls Domain module contract
'          functions in the correct order. Handles Derived field auto-computation.
'          Implements pipeline step markers and ResumeFrom (AP-45/MBP).
' =============================================================================

' Module-level pipeline tracking
Private m_currentStep As Long
Public SilentMode As Boolean
Private m_expectedRows As Long
Private m_runStartTime As Double
Private m_stepStartTimes(0 To 6) As Double
Private m_stepElapsed(0 To 6) As Double

' Phase 6A: DomainOutputs handoff for Application.Run dispatch (BUG-034 pattern)
Public DomainOutputs As Variant


' =============================================================================
' RunModel
' Generic entry point for the Dashboard "Run Model" button.
' Delegates to the kernel computation pipeline.
' =============================================================================
Public Sub RunModel()
    ' Check lock gate before running
    If KernelFormHelpers.CheckLockGate("RUN_MODEL") Then Exit Sub
    RunProjectionsEx
End Sub

' =============================================================================
' RunProjections
' Parameterless entry point visible in Alt+F8. Delegates to RunProjectionsEx.
' =============================================================================
Public Sub RunProjections()
    RunProjectionsEx
End Sub

' =============================================================================
' RunProjectionsEx
' The full computation pipeline. Optional seed for deterministic PRNG.
' =============================================================================
Public Sub RunProjectionsEx(Optional seed As Long = -1)
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    m_runStartTime = Timer
    KernelConfig.ResetRunErrorCounters

    ' --- Step 1: Load Config ---
    m_stepStartTimes(STEP_LOAD_CONFIG) = Timer
    m_currentStep = STEP_LOAD_CONFIG
    KernelConfig.LoadAllConfig
    WriteStepMarker STEP_LOAD_CONFIG, STEP_STATUS_COMPLETE
    m_stepElapsed(STEP_LOAD_CONFIG) = Timer - m_stepStartTimes(STEP_LOAD_CONFIG)
    If m_stepElapsed(STEP_LOAD_CONFIG) < 0 Then m_stepElapsed(STEP_LOAD_CONFIG) = m_stepElapsed(STEP_LOAD_CONFIG) + 86400

    ' --- Step 1.5: Initialize PRNG from ReproConfig (or explicit seed) ---
    InitializePRNG seed

    ' --- Step 1.6: Initialize Domain (dynamic dispatch via DomainModule setting) ---
    Dim domMod As String
    domMod = CStr(KernelConfig.GetSetting("DomainModule"))
    If Len(domMod) = 0 Then domMod = "SampleDomainEngine"
    If IsPipelineStepEnabled("DOMAIN_INIT") Then
        Application.Run domMod & ".Initialize"
    End If

    ' --- Step 2: Validate ---
    m_stepStartTimes(STEP_VALIDATE) = Timer
    m_currentStep = STEP_VALIDATE

    Dim entityCount As Long
    entityCount = DetectEntityCount()
    If entityCount = 0 Then
        KernelConfig.LogError SEV_FATAL, "KernelEngine", "E-200", _
                              "No entities found on Inputs sheet", _
                              "MANUAL BYPASS: Add entity names to the Inputs tab row 3, columns C onward. Then re-run RunProjections."
        WriteStepMarker STEP_VALIDATE, STEP_STATUS_FAILED
        GoTo Cleanup
    End If

    Dim periodCount As Long
    periodCount = KernelConfig.GetTimeHorizon()
    If periodCount <= 0 Then
        KernelConfig.LogError SEV_FATAL, "KernelEngine", "E-201", _
                              "Invalid TimeHorizon: " & CStr(periodCount), _
                              "MANUAL BYPASS: On the Config sheet, find the GRANULARITY_CONFIG section and set TimeHorizon to a positive integer. Then re-run RunProjections."
        WriteStepMarker STEP_VALIDATE, STEP_STATUS_FAILED
        GoTo Cleanup
    End If

    If Not Application.Run(domMod & ".Validate") Then
        KernelConfig.LogError SEV_FATAL, "KernelEngine", "E-202", _
                              "Domain validation failed", _
                              "MANUAL BYPASS: Check the ErrorLog tab for specific validation errors (E-300 series). Fix the flagged inputs on the Inputs tab. Then re-run RunProjections."
        WriteStepMarker STEP_VALIDATE, STEP_STATUS_FAILED
        If Not SilentMode Then
            KernelFormHelpers.ShowConfigMsgBox "VALIDATION_FAILED"
        End If
        GoTo Cleanup
    End If

    WriteStepMarker STEP_VALIDATE, STEP_STATUS_COMPLETE
    m_stepElapsed(STEP_VALIDATE) = Timer - m_stepStartTimes(STEP_VALIDATE)
    If m_stepElapsed(STEP_VALIDATE) < 0 Then m_stepElapsed(STEP_VALIDATE) = m_stepElapsed(STEP_VALIDATE) + 86400

    ' --- Step 3: Compute ---
    m_stepStartTimes(STEP_COMPUTE) = Timer
    m_currentStep = STEP_COMPUTE

    Application.Run domMod & ".Reset"

    ' Determine row count: try DomainEngine.GetRowCount() first (insurance model
    ' has development run-off beyond TimeHorizon). Fall back to entityCount x periodCount.
    Dim totalRows As Long
    On Error Resume Next
    totalRows = CLng(Application.Run(domMod & ".GetRowCount"))
    On Error GoTo ErrHandler
    If totalRows <= 0 Then
        totalRows = entityCount * periodCount
    End If
    m_expectedRows = totalRows

    ' Determine max period for Summary tab columns
    Dim maxPeriod As Long
    On Error Resume Next
    maxPeriod = CLng(Application.Run(domMod & ".GetMaxPeriod"))
    On Error GoTo ErrHandler
    If maxPeriod <= 0 Then maxPeriod = periodCount

    Dim totalCols As Long
    totalCols = KernelConfig.GetColumnCount()

    Dim outputs() As Variant
    ReDim outputs(1 To totalRows, 1 To totalCols)

    ' Phase 6A: PreCompute extensions
    If KernelExtension.GetActiveExtensionCount("PreCompute") > 0 Then
        KernelExtension.RunExtensions "PreCompute", outputs
    End If

    ' Domain fills Dimension and Incremental columns (AP-43 contract)
    ' DomainOutputs handoff pattern (Application.Run cannot pass arrays -- BUG-034)
    DomainOutputs = outputs
    Application.Run domMod & ".Execute"
    outputs = DomainOutputs
    DomainOutputs = Empty

    ' Compute Derived fields from DerivationRules (AP-42)
    ComputeDerivedFields outputs, totalRows

    ' Run any registered transforms (Phase 5B)
    If KernelTransform.GetTransformCount() > 0 Then
        KernelTransform.RunTransforms outputs
    End If

    ' Phase 6A: PostCompute extensions
    If KernelExtension.GetActiveExtensionCount("PostCompute") > 0 Then
        KernelExtension.RunExtensions "PostCompute", outputs
    End If

    WriteStepMarker STEP_COMPUTE, STEP_STATUS_COMPLETE
    m_stepElapsed(STEP_COMPUTE) = Timer - m_stepStartTimes(STEP_COMPUTE)
    If m_stepElapsed(STEP_COMPUTE) < 0 Then m_stepElapsed(STEP_COMPUTE) = m_stepElapsed(STEP_COMPUTE) + 86400

    ' --- Step 4: Write Detail ---
    If IsPipelineStepEnabled("WRITE_DETAIL") Then
        m_stepStartTimes(STEP_WRITE_DETAIL) = Timer
        m_currentStep = STEP_WRITE_DETAIL
        KernelOutput.WriteDetailTab outputs, totalRows
        WriteStepMarker STEP_WRITE_DETAIL, STEP_STATUS_COMPLETE
        m_stepElapsed(STEP_WRITE_DETAIL) = Timer - m_stepStartTimes(STEP_WRITE_DETAIL)
        If m_stepElapsed(STEP_WRITE_DETAIL) < 0 Then m_stepElapsed(STEP_WRITE_DETAIL) = m_stepElapsed(STEP_WRITE_DETAIL) + 86400
    Else
        WriteStepMarker STEP_WRITE_DETAIL, STEP_STATUS_SKIPPED
    End If

    ' --- Step 5: Write CSV ---
    Dim csvDir As String
    Dim csvPath As String
    If IsPipelineStepEnabled("WRITE_CSV") Then
        m_stepStartTimes(STEP_WRITE_CSV) = Timer
        m_currentStep = STEP_WRITE_CSV

        csvDir = ThisWorkbook.Path & "\..\scenarios"
        If Dir(csvDir, vbDirectory) = "" Then
            MkDir csvDir
        End If

        csvPath = csvDir & "\scenario_" & Format(Now, "yyyymmdd_hhnnss") & ".csv"
        KernelCsvIO.WriteCSV outputs, totalRows, csvPath

        WriteStepMarker STEP_WRITE_CSV, STEP_STATUS_COMPLETE
        m_stepElapsed(STEP_WRITE_CSV) = Timer - m_stepStartTimes(STEP_WRITE_CSV)
        If m_stepElapsed(STEP_WRITE_CSV) < 0 Then m_stepElapsed(STEP_WRITE_CSV) = m_stepElapsed(STEP_WRITE_CSV) + 86400
    Else
        WriteStepMarker STEP_WRITE_CSV, STEP_STATUS_SKIPPED
    End If

    ' --- Step 6: Write Summary ---
    If IsPipelineStepEnabled("WRITE_SUMMARY") Then
        m_stepStartTimes(STEP_WRITE_SUMMARY) = Timer
        m_currentStep = STEP_WRITE_SUMMARY
        KernelOutput.WriteSummaryFormulas entityCount, maxPeriod
        WriteStepMarker STEP_WRITE_SUMMARY, STEP_STATUS_COMPLETE
        m_stepElapsed(STEP_WRITE_SUMMARY) = Timer - m_stepStartTimes(STEP_WRITE_SUMMARY)
        If m_stepElapsed(STEP_WRITE_SUMMARY) < 0 Then m_stepElapsed(STEP_WRITE_SUMMARY) = m_stepElapsed(STEP_WRITE_SUMMARY) + 86400
    Else
        WriteStepMarker STEP_WRITE_SUMMARY, STEP_STATUS_SKIPPED
    End If

    WriteRunState

    ' --- Commit input state to Run Metadata tab ---
    WriteRunMetadataTab entityCount, periodCount

    ' --- Refresh data model if configured (Phase 5B) ---
    If IsPipelineStepEnabled("REFRESH_DATA_MODEL") Then
        KernelTransform.RefreshDataModel
    End If

    ' --- Prove-It Refresh (Phase 3) --- skip in silent mode (bootstrap)
    If IsPipelineStepEnabled("REFRESH_PROVEIT") And Not SilentMode Then
        RefreshProveItIfExists
    End If

    ' --- Phase 5A: Generate CumulativeView and refresh presentations ---
    If IsPipelineStepEnabled("REFRESH_PRESENTATIONS") And Not SilentMode Then
        RefreshPresentationsIfExist
    End If

    ' --- Phase 5C: Refresh formula tabs ---
    If IsPipelineStepEnabled("REFRESH_FORMULAS") Then
        KernelFormulaWriter.RefreshFormulaTabs
    End If

    ' --- Phase 6A: PostOutput extensions ---
    If IsPipelineStepEnabled("POST_OUTPUT_EXTENSIONS") Then
        If KernelExtension.GetActiveExtensionCount("PostOutput") > 0 Then
            KernelExtension.RunExtensions "PostOutput", outputs
        End If
    End If

    ' --- Force full recalculation before tests ---
    If IsPipelineStepEnabled("RECALCULATE") Then
        Application.Calculate
    End If

    ' --- Domain validation tests (if module exists) --- skip in silent mode
    Dim insTestResult As Boolean
    Dim domTestEntry As String
    If IsPipelineStepEnabled("DOMAIN_TESTS") And Not SilentMode Then
        insTestResult = True
        On Error Resume Next
        domTestEntry = KernelConfig.GetBrandingSetting("DomainTestEntry")
        If Len(domTestEntry) > 0 Then
            insTestResult = Application.Run(domTestEntry)
        End If
        If Err.Number <> 0 Then
            insTestResult = True  ' Module not present (sample config) -- skip
            Err.Clear
        End If
        On Error GoTo ErrHandler
        If Not insTestResult Then
            KernelConfig.LogError SEV_ERROR, "KernelEngine", "E-210", _
                "Domain validation tests failed. Check TestResults tab.", _
                "MANUAL BYPASS: Review TestResults tab for details. Pipeline results may be inconsistent."
        End If
    End If

    ' --- Done ---
    ' Write WAL entry if enabled
    If StrComp(KernelConfig.GetReproSetting("WALEnabled"), "TRUE", vbTextCompare) = 0 Then
        KernelSnapshot.WriteWAL "RUN_COMPLETE", _
            entityCount & " entities, " & periodCount & " periods, seed=" & KernelRandom.GetSeed()
    End If

    ' --- 12d: Write run metadata to Dashboard --- skip in silent mode
    If IsPipelineStepEnabled("WRITE_METADATA") And Not SilentMode Then
        WriteRunMetadata entityCount
    End If

    ' Regenerate Assumptions Register after run --- skip in silent mode
    If IsPipelineStepEnabled("REFRESH_ASSUMPTIONS") And Not SilentMode Then
        On Error Resume Next
        KernelAssumptions.GenerateAssumptionsRegister
        On Error GoTo ErrHandler
    End If

    ' AutoSaveOnRun: auto-save workspace if workspace_config says TRUE
    ' AutoSaveOnRun: skip in silent mode (bootstrap saves explicitly after)
    If Not SilentMode Then
        Dim autoSave As String
        autoSave = KernelConfig.GetWorkspaceSetting("AutoSaveOnRun")
        If StrComp(autoSave, "TRUE", vbTextCompare) = 0 Then
            On Error Resume Next
            KernelWorkspace.SaveWorkspace
            On Error GoTo ErrHandler
        End If
    End If

    ' Auto-lock after successful run if configured (skip in silent mode)
    If Not SilentMode Then
        Dim autoLock As String
        autoLock = KernelConfig.GetLockSetting("AutoLockOnRun")
        If StrComp(autoLock, "TRUE", vbTextCompare) = 0 Then
            On Error Resume Next
            KernelFormHelpers.SetModelLocked True
            On Error GoTo ErrHandler
        End If
    End If

    KernelConfig.LogError SEV_INFO, "KernelEngine", "I-200", _
                          "RunProjections completed successfully", _
                          entityCount & " entities, " & periodCount & " periods, " & totalRows & " rows"

    If Not SilentMode Then
        Dim totalElap As Double
        totalElap = Timer - m_runStartTime
        If totalElap < 0 Then totalElap = totalElap + 86400
        Dim runFatals As Long
        runFatals = KernelConfig.GetRunFatalCount()
        Dim runErrors As Long
        runErrors = KernelConfig.GetRunErrorCount()

        Dim doneMsg As String
        doneMsg = "RunProjections completed." & vbCrLf & _
            entityCount & " entities, " & periodCount & " periods, " & totalRows & " rows." & vbCrLf & _
            "Elapsed: " & Format(totalElap, "0.0") & "s"

        If runFatals > 0 Or runErrors > 0 Then
            doneMsg = doneMsg & vbCrLf & vbCrLf & _
                "WARNING: " & runFatals & " FATAL, " & runErrors & " ERROR entries logged during this run." & vbCrLf & _
                "Check the ErrorLog tab for details and manual bypass instructions."
            MsgBox doneMsg, vbExclamation, "RDK -- Complete (with errors)"
        Else
            MsgBox doneMsg, vbInformation, "RDK -- Complete"
        End If
    End If  ' Not SilentMode

Cleanup:
    Application.EnableEvents = True
    If Not SilentMode Then
        ' AutoFit BEFORE restoring automatic calc to avoid recalc during column sizing
        Application.ScreenUpdating = False
        AutoFitAllOutputTabs
        Application.ScreenUpdating = True
    End If
    Application.Calculation = xlCalculationAutomatic
    If Not SilentMode Then
        ' Navigate to post-run tab from branding config
        On Error Resume Next
        Dim postRunTab As String
        postRunTab = KernelConfig.GetBrandingSetting("PostRunActivateTab")
        If Len(postRunTab) > 0 Then ThisWorkbook.Sheets(postRunTab).Activate
        On Error GoTo 0
    End If
    Exit Sub

ErrHandler:
    Dim bypassMsg As String
    Select Case m_currentStep
        Case STEP_LOAD_CONFIG
            bypassMsg = "MANUAL BYPASS: Open the Config sheet and verify all section markers (=== COLUMN_REGISTRY === etc.) are present. Fix any missing data, then call ResumeFrom(" & STEP_LOAD_CONFIG & ")."
        Case STEP_VALIDATE
            bypassMsg = "MANUAL BYPASS: Check the ErrorLog tab for validation details. Fix the flagged inputs on the Inputs tab, then call ResumeFrom(" & STEP_VALIDATE & ")."
        Case STEP_COMPUTE
            bypassMsg = "MANUAL BYPASS: Verify inputs on the Inputs tab are numeric and in valid ranges. Then call ResumeFrom(" & STEP_COMPUTE & ")."
        Case STEP_WRITE_DETAIL
            bypassMsg = "MANUAL BYPASS: Detail tab may be partially written. Expected " & m_expectedRows & " data rows with headers: " & GetDetailHeaderList() & ". Paste data manually, then call ResumeFrom(" & STEP_WRITE_CSV & ")."
        Case STEP_WRITE_CSV
            bypassMsg = "MANUAL BYPASS: Copy the Detail tab data to a CSV file in the scenarios folder. Headers: " & GetDetailHeaderList() & ". Then call ResumeFrom(" & STEP_WRITE_SUMMARY & ")."
        Case STEP_WRITE_SUMMARY
            bypassMsg = "MANUAL BYPASS: On the Summary tab, create SUMIFS formulas referencing the Detail tab for each entity/metric combination. No further steps needed."
        Case Else
            bypassMsg = "MANUAL BYPASS: Re-run RunProjections from scratch after fixing the issue described above."
    End Select

    WriteStepMarker m_currentStep, STEP_STATUS_FAILED

    KernelConfig.LogError SEV_FATAL, "KernelEngine", "E-299", _
                          "Unhandled error in RunProjections at step " & m_currentStep & ": " & Err.Description, _
                          bypassMsg
    MsgBox "RunProjections failed at step " & m_currentStep & ":" & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "Check the ErrorLog tab for details.", _
           vbCritical, "RDK -- Error"
    KernelDiagnostic.AutoDump
    Resume Cleanup
End Sub


' =============================================================================
' ResumeFrom
' Re-enters the pipeline at the given step number.
' Validates that all prior step artifacts exist before resuming.
' =============================================================================
Public Sub ResumeFrom(stepNumber As Long)
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    KernelConfig.ResetRunErrorCounters

    ' Always reload config first
    KernelConfig.LoadAllConfig

    ' Initialize domain (registers transforms) -- dynamic dispatch
    Dim domMod As String
    domMod = CStr(KernelConfig.GetSetting("DomainModule"))
    If Len(domMod) = 0 Then domMod = "SampleDomainEngine"
    Application.Run domMod & ".Initialize"

    ' Validate prior artifacts based on requested step
    If stepNumber > STEP_VALIDATE Then
        ' Config must be loaded
        If KernelConfig.GetColumnCount() = 0 Then
            KernelConfig.LogError SEV_FATAL, "KernelEngine", "E-250", _
                                  "Cannot resume from step " & stepNumber & ": Config not loaded (column count = 0)", _
                                  "MANUAL BYPASS: Verify the Config sheet has a valid COLUMN_REGISTRY section. Then call ResumeFrom(" & STEP_LOAD_CONFIG & ")."
            GoTo Cleanup
        End If
    End If

    If stepNumber > STEP_COMPUTE Then
        ' Detail tab must have data
        Dim wsDetail As Worksheet
        Set wsDetail = ThisWorkbook.Sheets(TAB_DETAIL)
        If wsDetail.Cells(DETAIL_DATA_START_ROW, 1).Value = "" Then
            KernelConfig.LogError SEV_FATAL, "KernelEngine", "E-251", _
                                  "Cannot resume from step " & stepNumber & ": Detail tab has no data", _
                                  "MANUAL BYPASS: Populate the Detail tab with data rows starting at row " & DETAIL_DATA_START_ROW & ". Then call ResumeFrom(" & STEP_WRITE_DETAIL & ")."
            GoTo Cleanup
        End If
    End If

    ' Mark skipped steps
    Dim s As Long
    For s = STEP_LOAD_CONFIG To stepNumber - 1
        WriteStepMarker s, STEP_STATUS_SKIPPED
    Next s

    ' Detect entity/period counts
    Dim entityCount As Long
    entityCount = DetectEntityCount()

    Dim periodCount As Long
    periodCount = KernelConfig.GetTimeHorizon()

    ' Determine row count: try DomainEngine.GetRowCount() first
    Dim totalRows As Long
    On Error Resume Next
    totalRows = CLng(Application.Run(domMod & ".GetRowCount"))
    On Error GoTo ErrHandler
    If totalRows <= 0 Then
        totalRows = entityCount * periodCount
    End If
    m_expectedRows = totalRows

    Dim maxPeriod As Long
    On Error Resume Next
    maxPeriod = CLng(Application.Run(domMod & ".GetMaxPeriod"))
    On Error GoTo ErrHandler
    If maxPeriod <= 0 Then maxPeriod = periodCount

    ' Resume at the requested step
    Select Case stepNumber
        Case STEP_LOAD_CONFIG
            ' Just re-run from start
            GoTo Cleanup   ' Config already loaded; fall through to full run
        Case STEP_VALIDATE
            m_currentStep = STEP_VALIDATE
            If Not Application.Run(domMod & ".Validate") Then
                KernelConfig.LogError SEV_FATAL, "KernelEngine", "E-252", _
                                      "Domain validation failed on resume", _
                                      "MANUAL BYPASS: Check the ErrorLog tab for specific validation errors. Fix flagged inputs on the Inputs tab."
                WriteStepMarker STEP_VALIDATE, STEP_STATUS_FAILED
                GoTo Cleanup
            End If
            WriteStepMarker STEP_VALIDATE, STEP_STATUS_COMPLETE
            ' Fall through to compute
            GoTo ResumeCompute

        Case STEP_COMPUTE
            GoTo ResumeCompute

        Case STEP_WRITE_DETAIL
            GoTo ResumeWriteDetail

        Case STEP_WRITE_CSV
            GoTo ResumeWriteCSV

        Case STEP_WRITE_SUMMARY
            GoTo ResumeWriteSummary

        Case Else
            KernelConfig.LogError SEV_ERROR, "KernelEngine", "E-253", _
                                  "Invalid step number for ResumeFrom: " & stepNumber, _
                                  "MANUAL BYPASS: Valid step numbers are " & STEP_LOAD_CONFIG & " through " & STEP_WRITE_SUMMARY & "."
            GoTo Cleanup
    End Select

ResumeCompute:
    m_currentStep = STEP_COMPUTE
    Application.Run domMod & ".Reset"

    Dim totalCols As Long
    totalCols = KernelConfig.GetColumnCount()

    Dim outputs() As Variant
    ReDim outputs(1 To totalRows, 1 To totalCols)

    DomainOutputs = outputs
    Application.Run domMod & ".Execute"
    outputs = DomainOutputs
    DomainOutputs = Empty
    ComputeDerivedFields outputs, totalRows
    If KernelTransform.GetTransformCount() > 0 Then
        KernelTransform.RunTransforms outputs
    End If
    WriteStepMarker STEP_COMPUTE, STEP_STATUS_COMPLETE

    ' Write Detail
    m_currentStep = STEP_WRITE_DETAIL
    KernelOutput.WriteDetailTab outputs, totalRows
    WriteStepMarker STEP_WRITE_DETAIL, STEP_STATUS_COMPLETE

    ' Write CSV
    GoTo ResumeWriteCSVWithOutputs

ResumeWriteDetail:
    ' Re-compute to get outputs array
    m_currentStep = STEP_COMPUTE
    totalCols = KernelConfig.GetColumnCount()
    ReDim outputs(1 To totalRows, 1 To totalCols)
    Application.Run domMod & ".Reset"
    DomainOutputs = outputs
    Application.Run domMod & ".Execute"
    outputs = DomainOutputs
    DomainOutputs = Empty
    ComputeDerivedFields outputs, totalRows
    WriteStepMarker STEP_COMPUTE, STEP_STATUS_BYPASSED

    m_currentStep = STEP_WRITE_DETAIL
    KernelOutput.WriteDetailTab outputs, totalRows
    WriteStepMarker STEP_WRITE_DETAIL, STEP_STATUS_COMPLETE
    GoTo ResumeWriteCSVWithOutputs

ResumeWriteCSV:
    ' Re-compute to get outputs array for CSV
    m_currentStep = STEP_COMPUTE
    totalCols = KernelConfig.GetColumnCount()
    ReDim outputs(1 To totalRows, 1 To totalCols)
    Application.Run domMod & ".Reset"
    DomainOutputs = outputs
    Application.Run domMod & ".Execute"
    outputs = DomainOutputs
    DomainOutputs = Empty
    ComputeDerivedFields outputs, totalRows
    WriteStepMarker STEP_COMPUTE, STEP_STATUS_BYPASSED

ResumeWriteCSVWithOutputs:
    m_currentStep = STEP_WRITE_CSV
    Dim csvDir2 As String
    csvDir2 = ThisWorkbook.Path & "\..\scenarios"
    If Dir(csvDir2, vbDirectory) = "" Then
        MkDir csvDir2
    End If

    Dim csvPath2 As String
    csvPath2 = csvDir2 & "\scenario_" & Format(Now, "yyyymmdd_hhnnss") & ".csv"
    KernelCsvIO.WriteCSV outputs, totalRows, csvPath2
    WriteStepMarker STEP_WRITE_CSV, STEP_STATUS_COMPLETE

ResumeWriteSummary:
    m_currentStep = STEP_WRITE_SUMMARY
    KernelOutput.WriteSummaryFormulas entityCount, maxPeriod
    WriteStepMarker STEP_WRITE_SUMMARY, STEP_STATUS_COMPLETE

    WriteRunState
    RefreshPresentationsIfExist
    KernelFormulaWriter.RefreshFormulaTabs

    ' Phase 6A: PostOutput extensions (ResumeFrom path)
    If KernelExtension.GetActiveExtensionCount("PostOutput") > 0 Then
        KernelExtension.RunExtensions "PostOutput", outputs
    End If

    KernelConfig.LogError SEV_INFO, "KernelEngine", "I-201", _
                          "ResumeFrom(" & stepNumber & ") completed successfully", _
                          entityCount & " entities, " & periodCount & " periods"

    Dim runFatals As Long
    runFatals = KernelConfig.GetRunFatalCount()
    Dim runErrors As Long
    runErrors = KernelConfig.GetRunErrorCount()

    Dim resumeMsg As String
    resumeMsg = "ResumeFrom(" & stepNumber & ") completed." & vbCrLf & _
        entityCount & " entities, " & periodCount & " periods."
    If runFatals > 0 Or runErrors > 0 Then
        resumeMsg = resumeMsg & vbCrLf & vbCrLf & _
            "WARNING: " & runFatals & " FATAL, " & runErrors & " ERROR entries logged." & vbCrLf & _
            "Check the ErrorLog tab for details and manual bypass instructions."
        MsgBox resumeMsg, vbExclamation, "RDK -- Complete (with errors)"
    Else
        MsgBox resumeMsg, vbInformation, "RDK -- Complete"
    End If

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = False
    AutoFitAllOutputTabs
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    On Error Resume Next
    Dim postRunTab2 As String
    postRunTab2 = KernelConfig.GetBrandingSetting("PostRunActivateTab")
    If Len(postRunTab2) > 0 Then ThisWorkbook.Sheets(postRunTab2).Activate
    On Error GoTo 0
    Exit Sub

ErrHandler:
    WriteStepMarker m_currentStep, STEP_STATUS_FAILED
    KernelConfig.LogError SEV_FATAL, "KernelEngine", "E-259", _
                          "Unhandled error in ResumeFrom(" & stepNumber & ") at step " & m_currentStep & ": " & Err.Description, _
                          "MANUAL BYPASS: Check the ErrorLog tab for details. Fix the issue and call ResumeFrom(" & m_currentStep & ") again."
    KernelDiagnostic.AutoDump
    Resume Cleanup
End Sub


' =============================================================================
' WriteStepMarker
' Writes step status to the PIPELINE_STATE section on the Config sheet.
' =============================================================================
Private Sub WriteStepMarker(stepNum As Long, status As String)
    On Error Resume Next

    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)

    Dim markerRow As Long
    markerRow = FindPipelineStateRow(wsConfig)

    If markerRow = 0 Then
        ' No pipeline state section yet - create it at the bottom
        Dim lastRow As Long
        lastRow = wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).row + 2
        wsConfig.Cells(lastRow, 1).NumberFormat = "@"
        wsConfig.Cells(lastRow, 1).Value = PIPELINE_STATE_MARKER
        markerRow = lastRow
    End If

    ' Write step status: row = markerRow + stepNum + 1 (header row after marker)
    Dim dataRow As Long
    dataRow = markerRow + 1 + stepNum

    wsConfig.Cells(dataRow, 1).Value = "Step_" & stepNum
    wsConfig.Cells(dataRow, 2).Value = status
    wsConfig.Cells(dataRow, 3).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")

    On Error GoTo 0
End Sub


' =============================================================================
' FindPipelineStateRow
' Finds the row containing the PIPELINE_STATE marker on the Config sheet.
' =============================================================================
Private Function FindPipelineStateRow(wsConfig As Worksheet) As Long
    FindPipelineStateRow = 0

    Dim r As Long
    For r = 1 To wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).row
        If StrComp(CStr(wsConfig.Cells(r, 1).Value), PIPELINE_STATE_MARKER, vbTextCompare) = 0 Then
            FindPipelineStateRow = r
            Exit Function
        End If
    Next r
End Function


' =============================================================================
' GetDetailHeaderList
' Returns a comma-separated list of Detail column names for bypass messages.
' =============================================================================
Private Function GetDetailHeaderList() As String
    On Error GoTo Fallback

    Dim colCount As Long
    colCount = KernelConfig.GetColumnCount()

    If colCount = 0 Then GoTo Fallback

    Dim headers() As String
    ReDim headers(1 To colCount)

    Dim i As Long
    For i = 1 To colCount
        headers(i) = KernelConfig.GetColName(i)
    Next i

    GetDetailHeaderList = Join(headers, ", ")
    Exit Function

Fallback:
    GetDetailHeaderList = "(column list unavailable - check Config sheet COLUMN_REGISTRY)"
End Function


' =============================================================================
' DetectEntityCount
' Reads the Inputs sheet EntityName row, counts non-empty cells from column C.
' =============================================================================
Private Function DetectEntityCount() As Long
    ' Phase 11A: Use configurable inputs tab name
    Dim inputsTabName As String
    inputsTabName = CStr(KernelConfig.GetSetting("InputsTabName"))
    If Len(inputsTabName) = 0 Then inputsTabName = TAB_INPUTS

    Dim wsInputs As Worksheet
    On Error Resume Next
    Set wsInputs = ThisWorkbook.Sheets(inputsTabName)
    On Error GoTo 0
    If wsInputs Is Nothing Then
        Set wsInputs = ThisWorkbook.Sheets(TAB_INPUTS)
    End If

    ' Find the EntityName row from InputSchema
    Dim entityNameVal As Variant
    Dim colIdx As Long
    Dim cnt As Long
    cnt = 0

    Dim entityRow As Long
    entityRow = 0

    Dim paramCount As Long
    paramCount = KernelConfig.GetInputCount()

    Dim pidx As Long
    For pidx = 1 To paramCount
        If StrComp(KernelConfig.GetInputParam(pidx), "EntityName", vbTextCompare) = 0 Then
            entityRow = KernelConfig.GetInputRow(pidx)
            Exit For
        End If
    Next pidx

    If entityRow = 0 Then
        ' No EntityName in input_schema. Try EntitySourceTab/EntitySourceRow.
        Dim srcTab As String
        srcTab = CStr(KernelConfig.GetSetting("EntitySourceTab"))
        Dim srcRowStr As String
        srcRowStr = CStr(KernelConfig.GetSetting("EntitySourceRow"))
        Dim srcRow As Long
        If IsNumeric(srcRowStr) And Len(srcRowStr) > 0 Then srcRow = CLng(srcRowStr)

        If Len(srcTab) > 0 And srcRow > 0 Then
            Dim wsSrc As Worksheet
            On Error Resume Next
            Set wsSrc = ThisWorkbook.Sheets(srcTab)
            On Error GoTo 0
            If Not wsSrc Is Nothing Then
                ' Count non-empty entity names starting at srcRow, col C
                Dim maxEnt As Long
                maxEnt = CLng(KernelConfig.GetSetting("MaxEntities"))
                If maxEnt <= 0 Then maxEnt = 10
                Dim ei As Long
                For ei = 0 To maxEnt - 1
                    If Len(Trim(CStr(wsSrc.Cells(srcRow + ei, INPUT_ENTITY_START_COL).Value))) = 0 Then Exit For
                    cnt = cnt + 1
                Next ei
                DetectEntityCount = cnt
                Exit Function
            End If
        End If

        ' Final fallback: MaxEntities
        Dim maxEntSetting As String
        maxEntSetting = CStr(KernelConfig.GetSetting("MaxEntities"))
        If Len(maxEntSetting) > 0 And IsNumeric(maxEntSetting) Then
            DetectEntityCount = CLng(maxEntSetting)
        Else
            KernelConfig.LogError SEV_ERROR, "KernelEngine", "E-210", _
                                  "Cannot detect entity count. Set EntitySourceTab/EntitySourceRow or MaxEntities.", _
                                  "MANUAL BYPASS: Add EntitySourceTab and EntitySourceRow to granularity_config."
            DetectEntityCount = 0
        End If
        Exit Function
    End If

    ' Count non-empty cells starting at column C
    colIdx = INPUT_ENTITY_START_COL
    Do While colIdx < INPUT_ENTITY_START_COL + INPUT_MAX_ENTITIES
        entityNameVal = wsInputs.Cells(entityRow, colIdx).Value
        If Trim(CStr(entityNameVal)) = "" Then Exit Do
        cnt = cnt + 1
        colIdx = colIdx + 1
    Loop

    DetectEntityCount = cnt
End Function


' =============================================================================
' ComputeDerivedFields
' For each column where FieldClass = "Derived", parse the DerivationRule
' and compute the result for every row.
' =============================================================================
Private Sub ComputeDerivedFields(ByRef outputs() As Variant, totalRows As Long)
    Dim derivCols As Variant
    derivCols = KernelConfig.GetDerivedColumns()

    If Not IsArray(derivCols) Then Exit Sub

    Dim colIdx As Long
    For colIdx = LBound(derivCols) To UBound(derivCols)
        Dim colName As String
        colName = derivCols(colIdx)

        Dim rule As String
        rule = KernelConfig.GetDerivationRule(colName)

        If Len(rule) = 0 Then
            KernelConfig.LogError SEV_WARN, "KernelEngine", "E-220", _
                                  "Derived column has no DerivationRule: " & colName, _
                                  "MANUAL BYPASS: On the Config sheet COLUMN_REGISTRY, add a DerivationRule for column '" & colName & "' (e.g., 'Revenue - COGS')."
            GoTo NextDerived
        End If

        ' Parse the rule: find the operator
        Dim operatorStr As String
        Dim operandA As String
        Dim operandB As String
        operatorStr = ""

        If ParseDerivationRule(rule, operandA, operatorStr, operandB) Then
            Dim detailColResult As Long
            detailColResult = KernelConfig.ColIndex(colName)

            Dim detailColA As Long
            detailColA = KernelConfig.ColIndex(operandA)

            Dim detailColB As Long
            detailColB = KernelConfig.ColIndex(operandB)

            If detailColResult < 1 Or detailColA < 1 Or detailColB < 1 Then
                KernelConfig.LogError SEV_ERROR, "KernelEngine", "E-221", _
                                      "DerivationRule references unknown column: " & rule, _
                                      "MANUAL BYPASS: Verify that columns '" & operandA & "' and '" & operandB & "' exist in the COLUMN_REGISTRY on the Config sheet."
                GoTo NextDerived
            End If

            ' Compute for every row
            Dim rowIdx As Long
            For rowIdx = 1 To totalRows
                Dim valA As Double
                Dim valB As Double
                Dim result As Double

                valA = 0
                valB = 0
                If IsNumeric(outputs(rowIdx, detailColA)) Then valA = CDbl(outputs(rowIdx, detailColA))
                If IsNumeric(outputs(rowIdx, detailColB)) Then valB = CDbl(outputs(rowIdx, detailColB))

                Select Case operatorStr
                    Case "-"
                        result = valA - valB
                    Case "+"
                        result = valA + valB
                    Case "*"
                        result = valA * valB
                    Case "/"
                        If valB = 0 Then
                            result = 0
                        Else
                            result = valA / valB
                        End If
                End Select

                outputs(rowIdx, detailColResult) = result
            Next rowIdx
        Else
            ' Identity/copy rule: rule is a single column name (BUG-061)
            Dim identityCol As Long
            identityCol = KernelConfig.ColIndex(rule)
            If identityCol >= 1 Then
                Dim detailColCopy As Long
                detailColCopy = KernelConfig.ColIndex(colName)
                If detailColCopy >= 1 Then
                    Dim copyRow As Long
                    For copyRow = 1 To totalRows
                        outputs(copyRow, detailColCopy) = outputs(copyRow, identityCol)
                    Next copyRow
                End If
            Else
                KernelConfig.LogError SEV_ERROR, "KernelEngine", "E-222", _
                                      "Cannot parse DerivationRule: " & rule, _
                                      "MANUAL BYPASS: DerivationRule must be in format 'ColumnA <op> ColumnB' or a single column name for identity copy. Fix on the Config sheet COLUMN_REGISTRY."
            End If
        End If

NextDerived:
    Next colIdx
End Sub


' =============================================================================
' ParseDerivationRule
' Parses a simple rule like "A - B" into operandA, operator, operandB.
' Supports: +, -, *, / (single operator only).
' Returns True if parsed successfully.
' =============================================================================
Private Function ParseDerivationRule(rule As String, ByRef operandA As String, _
                                     ByRef operatorStr As String, _
                                     ByRef operandB As String) As Boolean
    ParseDerivationRule = False

    ' Try each operator with space delimiters
    Dim ops As Variant
    ops = Array(" - ", " + ", " * ", " / ")

    Dim opIdx As Long
    For opIdx = LBound(ops) To UBound(ops)
        Dim pos As Long
        pos = InStr(1, rule, ops(opIdx), vbTextCompare)
        If pos > 0 Then
            operandA = Trim(Mid(rule, 1, pos - 1))
            operatorStr = Trim(Mid(ops(opIdx), 1))
            operandB = Trim(Mid(rule, pos + Len(ops(opIdx))))

            If Len(operandA) > 0 And Len(operandB) > 0 Then
                ParseDerivationRule = True
                Exit Function
            End If
        End If
    Next opIdx
End Function


' =============================================================================
' InitializePRNG
' Reads ReproConfig to determine seed behavior and initializes KernelRandom.
' If an explicit seed is passed (>= 0), it overrides config entirely.
' Pass -1 (default) to use config-driven behavior.
' =============================================================================
Private Sub InitializePRNG(Optional explicitSeed As Long = -1)
    ' Explicit seed overrides everything
    If explicitSeed >= 0 Then
        KernelConfig.LogError SEV_INFO, "KernelEngine", "I-205", _
            "PRNG: explicit seed provided: " & explicitSeed, ""
        KernelRandom.InitSeed explicitSeed
        Exit Sub
    End If

    ' Read config-driven behavior
    Dim mode As String
    mode = UCase(Trim(KernelConfig.GetReproSetting("DeterministicMode")))

    Dim defaultSeed As String
    defaultSeed = KernelConfig.GetReproSetting("DefaultSeed")

    Dim seedVal As Long
    seedVal = 0
    If IsNumeric(defaultSeed) And Len(defaultSeed) > 0 Then seedVal = CLng(defaultSeed)

    KernelConfig.LogError SEV_INFO, "KernelEngine", "I-206", _
        "PRNG config: DeterministicMode=" & mode & ", DefaultSeed=" & defaultSeed & _
        " (resolved=" & seedVal & ")", ""

    Select Case mode
        Case "TRUE"
            ' Always require seed
            If seedVal = 0 Then
                KernelConfig.LogError SEV_WARN, "KernelEngine", "W-200", _
                    "DeterministicMode=TRUE but DefaultSeed=0; using fallback seed 5489", ""
                seedVal = 5489
            End If
            KernelRandom.InitSeed seedVal

        Case "FALSE"
            ' Always random
            KernelRandom.AutoSeed

        Case Else
            ' AUTO (default): use config seed if provided, keep existing if already
            ' initialized (e.g. user set seed via Immediate Window), else random
            If seedVal <> 0 Then
                KernelRandom.InitSeed seedVal
            ElseIf KernelRandom.IsInitialized Then
                KernelConfig.LogError SEV_INFO, "KernelEngine", "I-207", _
                    "PRNG already initialized with seed " & KernelRandom.GetSeed() & " -- keeping it", ""
            Else
                KernelRandom.AutoSeed
            End If
    End Select
End Sub


' =============================================================================
' WriteRunState
' Writes timing and hash data to the Config sheet RUN_STATE section.
' =============================================================================
Private Sub WriteRunState()
    On Error Resume Next
    Dim totalElapsed As Double
    totalElapsed = Timer - m_runStartTime
    If totalElapsed < 0 Then totalElapsed = totalElapsed + 86400
    KernelFormHelpers.WriteRunStateValue RS_KEY_TIMESTAMP, Format(Now, "yyyy-mm-dd hh:nn:ss")
    KernelFormHelpers.WriteRunStateValue RS_KEY_TOTAL_ELAPSED, Format(totalElapsed, "0.00")
    Dim s As Long
    For s = STEP_LOAD_CONFIG To STEP_WRITE_SUMMARY
        If m_stepElapsed(s) > 0 Then
            KernelFormHelpers.WriteRunStateValue "Step_" & s & "_Elapsed", Format(m_stepElapsed(s), "0.00")
        End If
    Next s
    Dim inputHash As String
    inputHash = KernelFormHelpers.BuildInputHash()
    KernelFormHelpers.WriteRunStateValue RS_KEY_INPUT_HASH, inputHash
    Dim configHash As String
    configHash = KernelSnapshot.BuildConfigHash()
    KernelFormHelpers.WriteRunStateValue RS_KEY_CONFIG_HASH, configHash
    KernelFormHelpers.WriteRunStateValue RS_KEY_STALE, "FALSE"
    On Error GoTo 0
End Sub


' =============================================================================
' RefreshPresentationsIfExist
' If presentation config exists, generate CumulativeView and refresh tabs.
' =============================================================================
Private Sub RefreshPresentationsIfExist()
    On Error Resume Next
    ' Only refresh if summary_config has data
    If KernelConfig.GetSummaryConfigCount() > 0 Then
        KernelTabs.GenerateCumulativeView
        KernelTabs.RefreshAllPresentations
    End If
    On Error GoTo 0
End Sub


' =============================================================================
' RefreshProveItIfExists
' If the ProveIt tab exists and has content, silently refresh it.
' =============================================================================
Private Sub RefreshProveItIfExists()
    On Error Resume Next
    Dim wsPI As Worksheet
    Set wsPI = ThisWorkbook.Sheets(TAB_PROVE_IT)
    If wsPI Is Nothing Then Exit Sub

    ' Only refresh if tab has existing content (row 5+ has data)
    If Len(Trim(CStr(wsPI.Cells(5, 1).Value))) > 0 Then
        KernelProveIt.RefreshProveIt
    End If
    On Error GoTo 0
End Sub


' =============================================================================
' AutoFitAllOutputTabs
' Auto-fits columns on all visible worksheets so values display without
' manual resizing. Called after Calculation is restored to xlCalculationAutomatic
' so formula results are evaluated before column widths are measured.
' =============================================================================
' =============================================================================
' WriteRunMetadata (12d)
' Writes key run metadata to the Dashboard tab (F2:G6) so the user knows
' when outputs were generated and from what inputs.
' =============================================================================
Private Sub WriteRunMetadata(entityCount As Long)
    On Error Resume Next
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets(TAB_DASHBOARD)
    If wsDash Is Nothing Then Exit Sub

    ' Write to MODEL OUTPUT panel (rows 7-10, cols 5-6)
    wsDash.Cells(7, 5).Value = Format(Now, "yyyy-mm-dd  hh:mm")
    wsDash.Cells(8, 5).Value = CStr(KernelConfig.InputValue("Global Assumptions", "ScenarioName", 1))
    ' Count programs with non-zero GWP on Detail tab
    Dim activeProgs As Long
    activeProgs = 0
    Dim wsDetCheck As Worksheet
    Set wsDetCheck = Nothing
    Set wsDetCheck = ThisWorkbook.Sheets(TAB_DETAIL)
    If Not wsDetCheck Is Nothing Then
        Dim detLastRow As Long
        detLastRow = wsDetCheck.Cells(wsDetCheck.Rows.Count, 1).End(xlUp).Row
        If detLastRow >= DETAIL_DATA_START_ROW Then
            ' Find G_WP column (Gross Written Premium)
            Dim gwpCol As Long
            gwpCol = KernelConfig.TryColIndex("G_WP")
            If gwpCol < 1 Then gwpCol = KernelConfig.TryColIndex("Revenue")
            If gwpCol > 0 Then
                Dim progDict As Object
                Set progDict = CreateObject("Scripting.Dictionary")
                progDict.CompareMode = vbTextCompare
                ' Batch read entity + GWP columns
                Dim entData As Variant
                entData = wsDetCheck.Range(wsDetCheck.Cells(DETAIL_DATA_START_ROW, 1), _
                    wsDetCheck.Cells(detLastRow, 1)).Value
                Dim gwpData As Variant
                gwpData = wsDetCheck.Range(wsDetCheck.Cells(DETAIL_DATA_START_ROW, gwpCol), _
                    wsDetCheck.Cells(detLastRow, gwpCol)).Value
                Dim dri As Long
                For dri = 1 To UBound(entData, 1)
                    Dim eName As String
                    eName = Trim(CStr(entData(dri, 1)))
                    If Len(eName) > 0 And Not progDict.Exists(eName) Then
                        If IsNumeric(gwpData(dri, 1)) Then
                            If CDbl(gwpData(dri, 1)) > 0 Then
                                progDict.Add eName, True
                            End If
                        End If
                    End If
                Next dri
                activeProgs = progDict.Count
            End If
        End If
    End If
    If activeProgs = 0 Then activeProgs = entityCount
    wsDash.Cells(9, 5).Value = CStr(activeProgs) & " active"

    ' Config-driven metrics from branding_config
    Dim metricCount As Long
    Dim mcStr As String
    mcStr = KernelConfig.GetBrandingSetting("DashMetricCount")
    If IsNumeric(mcStr) And Len(mcStr) > 0 Then metricCount = CLng(mcStr) Else metricCount = 0

    Dim mi As Long
    For mi = 1 To metricCount
        Dim mLabel As String
        mLabel = KernelConfig.GetBrandingSetting("DashMetric" & mi & "Label")
        Dim mSource As String
        mSource = KernelConfig.GetBrandingSetting("DashMetric" & mi & "Source")
        If Len(mLabel) > 0 Then
            ' GWP metric goes in row 10 (Total GWP row in MODEL OUTPUT panel)
            Dim metricRow As Long
            metricRow = 9 + mi
            If StrComp(mSource, "EntityCount", vbTextCompare) = 0 Then
                ' Already written as "N active" in row 9
                GoTo NextMetric
            ElseIf InStr(1, mSource, "!") > 0 Then
                ' Tab!RowID format -- read Y1 Total from that tab/row
                Dim bangPos As Long
                bangPos = InStr(1, mSource, "!")
                Dim mTab As String
                mTab = Left(mSource, bangPos - 1)
                Dim mRowID As String
                mRowID = Mid(mSource, bangPos + 1)
                Dim mWs As Worksheet
                Set mWs = ThisWorkbook.Sheets(mTab)
                If Not mWs Is Nothing Then
                    Dim mRow As Long
                    mRow = KernelFormula.ResolveRowID(mTab, mRowID)
                    If mRow > 0 Then
                        ' Read Grand Total column (full horizon, all years)
                        Dim nYrs As Long
                        Dim timeH As Long
                        timeH = KernelConfig.GetTimeHorizon()
                        If timeH <= 0 Then timeH = 12
                        nYrs = (timeH \ 3) \ QS_QUARTERS_PER_YEAR
                        If nYrs < 1 Then nYrs = 1
                        Dim gtCol As Long
                        gtCol = QS_DATA_START_COL + nYrs * QS_COLS_PER_YEAR
                        Dim gwpVal As Double
                        gwpVal = CDbl(mWs.Cells(mRow, gtCol).Value)
                        ' Format as $NNM or $NNN,NNN
                        If Abs(gwpVal) >= 1000000 Then
                            wsDash.Cells(10, 5).Value = "$" & Format(gwpVal / 1000000, "#,##0.0") & "M"
                        Else
                            wsDash.Cells(10, 5).Value = "$" & Format(gwpVal, "#,##0")
                        End If
                    End If
                End If
            End If
NextMetric:
        End If
    Next mi

    ' MODEL OUTPUT panel formatting already handled by SetupDashboardTab
    On Error GoTo 0
End Sub


' =============================================================================
' WriteRunMetadataTab
' Captures committed input state to the Run Metadata tab at end of pipeline.
' Downstream formula tabs reference this tab (not the input tab) so that
' changing inputs mid-edit doesn't corrupt displayed results.
' =============================================================================
Private Sub WriteRunMetadataTab(entityCount As Long, periodCount As Long)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets(TAB_RUN_METADATA)
    If ws Is Nothing Then Exit Sub

    ws.Cells.ClearContents

    ' Header
    ws.Cells(1, 1).Value = "Run Metadata"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 2).Value = "Committed input state from last Run Model"
    ws.Cells(1, 2).Font.Color = RGB(128, 128, 128)

    ' Run info
    ws.Cells(2, 1).Value = "RunTimestamp"
    ws.Cells(2, 2).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(3, 1).Value = "KernelVersion"
    ws.Cells(3, 2).Value = KERNEL_VERSION
    ws.Cells(4, 1).Value = "EntityCount"
    ws.Cells(4, 2).Value = entityCount
    ws.Cells(5, 1).Value = "TimeHorizon"
    ws.Cells(5, 2).Value = periodCount
    ws.Cells(6, 1).Value = "ScenarioName"
    ws.Cells(6, 2).Value = CStr(KernelConfig.InputValue("Global Assumptions", "ScenarioName", 1))

    ' Entity names (committed snapshot from input tab)
    ws.Cells(8, 1).Value = "=== ENTITY NAMES ==="
    ws.Cells(8, 1).Font.Bold = True
    ws.Cells(9, 1).Value = "EntityIndex"
    ws.Cells(9, 2).Value = "EntityName"
    ws.Cells(9, 3).Value = "BusinessUnit"
    ws.Cells(9, 1).Font.Bold = True
    ws.Cells(9, 2).Font.Bold = True
    ws.Cells(9, 3).Font.Bold = True

    ' Read entity names and BU from EntitySourceTab
    Dim srcTab As String
    srcTab = CStr(KernelConfig.GetSetting("EntitySourceTab"))
    Dim srcRowStr As String
    srcRowStr = CStr(KernelConfig.GetSetting("EntitySourceRow"))
    Dim srcRow As Long
    If IsNumeric(srcRowStr) And Len(srcRowStr) > 0 Then srcRow = CLng(srcRowStr) Else srcRow = 6

    Dim wsSource As Worksheet
    If Len(srcTab) > 0 Then Set wsSource = ThisWorkbook.Sheets(srcTab)

    Dim i As Long
    For i = 1 To entityCount
        ws.Cells(9 + i, 1).Value = i
        ws.Cells(9 + i, 2).Value = KernelConfig.GetEntityName(i)
        ' BU: column 2 of the entity source row
        If Not wsSource Is Nothing And srcRow > 0 Then
            ws.Cells(9 + i, 3).Value = Trim(CStr(wsSource.Cells(srcRow + i - 1, 2).Value))
        End If
    Next i

    ' Center-align all value cells
    ws.Columns(2).HorizontalAlignment = xlCenter
    ws.Columns(3).HorizontalAlignment = xlCenter

    On Error GoTo 0
End Sub


' =============================================================================
' IsPipelineStepEnabled
' Reads pipeline_config section on Config sheet. Returns True if StepID is
' enabled, or True if pipeline_config section is missing (backward compat).
' =============================================================================
Private Function IsPipelineStepEnabled(stepID As String) As Boolean
    IsPipelineStepEnabled = True  ' default: enabled if config missing
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_CONFIG)
    If ws Is Nothing Then Exit Function
    On Error GoTo 0
    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(ws, CFG_MARKER_PIPELINE_CONFIG)
    If sr = 0 Then Exit Function  ' no pipeline config = all enabled
    Dim dr As Long
    dr = sr + 2
    Do While Len(Trim(CStr(ws.Cells(dr, PLCFG_COL_STEPID).Value))) > 0
        If StrComp(Trim(CStr(ws.Cells(dr, PLCFG_COL_STEPID).Value)), stepID, vbTextCompare) = 0 Then
            IsPipelineStepEnabled = (StrComp(Trim(CStr(ws.Cells(dr, PLCFG_COL_ENABLED).Value)), "TRUE", vbTextCompare) = 0)
            Exit Function
        End If
        dr = dr + 1
    Loop
    ' StepID not found in config = enabled by default
End Function


Private Sub AutoFitAllOutputTabs()
    On Error Resume Next
    Dim ws As Worksheet
    Dim skipTabs As String
    skipTabs = "|" & TAB_DASHBOARD & "|" & TAB_COVER_PAGE & "|" & TAB_USER_GUIDE & "|" & _
               TAB_ERROR_LOG & "|" & TAB_TEST_RESULTS & "|" & TAB_RUN_METADATA & "|" & TAB_CONFIG & "|"

    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible <> xlSheetVisible Then GoTo NextWS
        If InStr(1, skipTabs, "|" & ws.Name & "|", vbTextCompare) > 0 Then GoTo NextWS

        ' AutoFit columns
        ws.Columns.AutoFit

        ' Hide Column A if it has RowIDs (formula tabs)
        If Not ws.Columns(1).Hidden Then
            If Len(Trim(CStr(ws.Cells(2, 1).Value))) > 0 Then
                If ws.Name <> TAB_ASSUMPTIONS And ws.Name <> TAB_INPUTS Then
                    ws.Columns(1).Hidden = True
                End If
            End If
        End If
NextWS:
    Next ws
    On Error GoTo 0
End Sub
