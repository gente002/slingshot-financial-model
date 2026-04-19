Attribute VB_Name = "KernelBootstrap"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelBootstrap.bas
' Purpose: Called by Setup.ps1 after VBA modules are imported.
'          Generates the workbook structure from config tables.
' =============================================================================

' Sub-step tracking moved to KernelBootstrapUI.m_subStep (Public)


' =============================================================================
' BootstrapWorkbook
' Called once by Setup.ps1 after all VBA modules are imported.
' =============================================================================
Public Sub BootstrapWorkbook()
    Dim bootstrapStep As String

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    bootstrapStep = "Step 1: LoadConfigFromDisk"
    LoadConfigFromDisk

    bootstrapStep = "Step 2: LoadAllConfig"
    KernelConfig.LoadAllConfig

    bootstrapStep = "Step 2b: LoadExtensionRegistry"
    KernelExtension.LoadExtensionRegistry

    bootstrapStep = "Step 3: CreateTabsFromRegistry"
    KernelBootstrapUI.CreateTabsFromRegistry

    bootstrapStep = "Step 4: GenerateInputsTab"
    On Error GoTo ErrHandler_Step4
    KernelBootstrapUI.GenerateInputsTab
    On Error GoTo ErrHandler

    bootstrapStep = "Step 5: SetupDetailHeaders"
    KernelBootstrapUI.SetupDetailHeaders

    bootstrapStep = "Step 6: SetupSummaryStructure"
    KernelBootstrapUI.SetupSummaryStructure

    bootstrapStep = "Step 7: SetupErrorLog"
    KernelBootstrapUI.SetupErrorLog

    bootstrapStep = "Step 8: SetupDashboardTab"
    KernelBootstrapUI.SetupDashboardTab

    bootstrapStep = "Step 9: ProtectOutputTabs"
    KernelBootstrapUI.ProtectOutputTabs

    bootstrapStep = "Step 9b: CreateAllForms"
    KernelFormSetup.CreateAllForms

    ' BUG-081: Workbook_Open injection removed. The on-open health check
    ' caused COM object hangs, crash loops, and safe mode prompts. Users can
    ' run RunHealthCheckFull from the Dashboard button instead.
    ' bootstrapStep = "Step 9c: InjectWorkbookOpen"
    ' InjectWorkbookOpen

    bootstrapStep = "Step 9d: CreateFormulaTabs"
    KernelFormulaWriter.CreateFormulaTabs True

    ' Named ranges created during first model run via RefreshFormulaTabs
    ' (QuarterlySummary is empty at bootstrap, so RowIDs cannot resolve yet)

    ' Generate Assumptions Register tab from assumptions_config
    bootstrapStep = "Step 9d2: GenerateAssumptionsRegister"
    On Error Resume Next
    KernelAssumptions.GenerateAssumptionsRegister
    On Error GoTo ErrHandler

    ' Run PostBootstrap extensions (e.g., Curve Reference tab)
    bootstrapStep = "Step 9e: PostBootstrap extensions"
    If KernelExtension.GetActiveExtensionCount("PostBootstrap") > 0 Then
        Dim bootOutputs() As Variant
        ReDim bootOutputs(0)
        On Error Resume Next
        KernelExtension.RunExtensions "PostBootstrap", bootOutputs
        On Error GoTo ErrHandler
    End If

    bootstrapStep = "Step 9f: ApplyDefaultDevMode"
    KernelBootstrapUI.ApplyDefaultDevMode

    bootstrapStep = "Step 9g: PopulateCoverPage"
    KernelBootstrapUI.PopulateCoverPage

    bootstrapStep = "Step 9h: PopulateUserGuide"
    ' Dispatch to domain module via branding_config
    Dim ugEntry As String
    ugEntry = KernelConfig.GetBrandingSetting("UserGuideEntry")
    If Len(ugEntry) > 0 Then
        On Error Resume Next
        Application.Run ugEntry
        Err.Clear
        On Error GoTo ErrHandler
    End If

    ' Create buttons on non-Dashboard tabs from button_config
    bootstrapStep = "Step 9i: CreateButtonsOnAllTabs"
    KernelButtons.CreateButtonsOnAllTabs

    ' Stamp build fingerprint on VeryHidden sheet
    bootstrapStep = "Step 9j: Fingerprint"
    KernelBootstrapUI.StampFingerprint

    bootstrapStep = "Step 10: Log completion"
    KernelConfig.LogError SEV_INFO, "KernelBootstrap", "I-500", _
                          "Bootstrap completed successfully", _
                          "Kernel v" & KERNEL_VERSION

    ' Step 10b: Seed starter workspaces (inputs-only, no Run Model needed)
    bootstrapStep = "Step 10b: Seed workspaces"
    On Error Resume Next
    KernelBootstrapUI.SeedStarterWorkspaces
    On Error GoTo ErrHandler

    ' Default to Dashboard tab on open
    On Error Resume Next
    ThisWorkbook.Sheets(TAB_DASHBOARD).Activate
    On Error GoTo 0

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ' Force recalculation so UDF formulas (CurveRefPct, etc.) evaluate
    Application.CalculateFull
    Exit Sub

ErrHandler_Step4:
    bootstrapStep = "Step 4: GenerateInputsTab >> " & KernelBootstrapUI.m_subStep
ErrHandler:
    Dim errMsg As String
    errMsg = "Bootstrap failed at [" & bootstrapStep & "]: " & _
             Err.Description & " (Error " & Err.Number & ")"

    Dim bypassMsg As String
    Select Case True
        Case InStr(1, bootstrapStep, "Step 1", vbTextCompare) > 0
            bypassMsg = "MANUAL BYPASS: Verify config CSVs exist in the config/ folder (column_registry.csv, input_schema.csv, granularity_config.csv, tab_registry.csv). Re-run Setup.bat."
        Case InStr(1, bootstrapStep, "Step 2", vbTextCompare) > 0
            bypassMsg = "MANUAL BYPASS: Config sheet may have malformed data. Open Config sheet (unhide via VBA: Sheets(""Config"").Visible=True) and verify section markers."
        Case InStr(1, bootstrapStep, "Step 3", vbTextCompare) > 0
            bypassMsg = "MANUAL BYPASS: Manually create the required tabs (Inputs, Detail, Summary, ErrorLog). Re-run Setup.bat to retry."
        Case InStr(1, bootstrapStep, "Step 4", vbTextCompare) > 0
            bypassMsg = "MANUAL BYPASS: Inputs tab generation failed at sub-step [" & m_subStep & "]. Manually populate the Inputs tab with entity data starting at row 3, column C."
        Case InStr(1, bootstrapStep, "Step 5", vbTextCompare) > 0
            bypassMsg = "MANUAL BYPASS: Detail tab headers failed. Manually add column headers from column_registry.csv to row 1 of the Detail tab."
        Case InStr(1, bootstrapStep, "Step 6", vbTextCompare) > 0
            bypassMsg = "MANUAL BYPASS: Summary tab setup failed. Add 'Entity' in A1 and 'Metric' in B1 on the Summary tab. RunProjections will generate formulas."
        Case InStr(1, bootstrapStep, "Step 7", vbTextCompare) > 0
            bypassMsg = "MANUAL BYPASS: ErrorLog tab setup failed. Create tab with headers: Timestamp, Severity, Source, Code, Message, Detail in row 1."
        Case InStr(1, bootstrapStep, "Step 8", vbTextCompare) > 0
            bypassMsg = "MANUAL BYPASS: Dashboard tab setup failed. Non-critical -- proceed with RunProjections."
        Case InStr(1, bootstrapStep, "Step 9d", vbTextCompare) > 0
            bypassMsg = "MANUAL BYPASS: Formula tab creation failed. Non-critical -- formula tabs can be refreshed later via Dashboard button. Proceed with RunProjections."
        Case InStr(1, bootstrapStep, "Step 9c", vbTextCompare) > 0
            bypassMsg = "MANUAL BYPASS: Workbook_Open injection failed. Non-critical -- health check on open will not run automatically. Proceed with RunProjections."
        Case InStr(1, bootstrapStep, "Step 9b", vbTextCompare) > 0
            bypassMsg = "MANUAL BYPASS: UserForm creation failed. Non-critical -- forms are convenience UI. Proceed with RunProjections."
        Case InStr(1, bootstrapStep, "Step 9", vbTextCompare) > 0
            bypassMsg = "MANUAL BYPASS: Tab protection failed. Non-critical -- proceed with RunProjections."
        Case Else
            bypassMsg = "MANUAL BYPASS: Re-run Setup.bat after fixing the issue described above."
    End Select

    errMsg = errMsg & " | " & bypassMsg

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ' Write error to Sheet1 A1 for diagnostic read-back (COM swallows Err.Raise details)
    On Error Resume Next
    ThisWorkbook.Sheets(1).Cells(1, 1).Value = errMsg
    On Error GoTo 0
    Err.Raise vbObjectError + 1000, "KernelBootstrap", errMsg
End Sub


' =============================================================================
' LoadConfigFromDisk
' Reads config CSVs from the config/ directory and writes to the Config sheet.
' =============================================================================
Public Sub LoadConfigFromDisk()
    Dim loadStep As String
    loadStep = "init"

    ' Clear section position cache (will rebuild after loading)
    KernelConfigLoader.ClearSectionCache

    ' Ensure Config sheet exists
    loadStep = "EnsureSheet"
    Dim wsConfig As Worksheet
    Set wsConfig = EnsureSheet(TAB_CONFIG)
    wsConfig.Cells.ClearContents

    loadStep = "ResolvePath"
    Dim configDir As String
    ' Config lives at project root, one level up from workbook/
    ' Resolve parent directory without FSO
    Dim wbPath As String
    wbPath = ThisWorkbook.Path
    Dim parentDir As String
    parentDir = Left(wbPath, InStrRev(wbPath, "\") - 1)
    configDir = parentDir & "\config"

    If Dir(configDir, vbDirectory) = "" Then
        Err.Raise vbObjectError + 1001, "KernelBootstrap", _
                  "Config directory not found: " & configDir
    End If

    Dim curRow As Long
    curRow = 1

    ' Load column_registry.csv
    loadStep = "Load column_registry.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\column_registry.csv", _
                                  CFG_MARKER_COLUMN_REGISTRY)

    ' Blank separator row
    curRow = curRow + 1

    ' Load input_schema.csv
    loadStep = "Load input_schema.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\input_schema.csv", _
                                  CFG_MARKER_INPUT_SCHEMA)

    ' Blank separator row
    curRow = curRow + 1

    ' Load granularity_config.csv
    loadStep = "Load granularity_config.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\granularity_config.csv", _
                                  CFG_MARKER_GRANULARITY_CONFIG)

    ' Blank separator row
    curRow = curRow + 1

    ' Load tab_registry.csv
    loadStep = "Load tab_registry.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\tab_registry.csv", _
                                  CFG_MARKER_TAB_REGISTRY)

    ' Blank separator row
    curRow = curRow + 1

    ' Load repro_config.csv (Phase 2)
    loadStep = "Load repro_config.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\repro_config.csv", _
                                  CFG_MARKER_REPRO_CONFIG)

    ' Blank separator row
    curRow = curRow + 1

    ' Load scale_limits.csv (Phase 2)
    loadStep = "Load scale_limits.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\scale_limits.csv", _
                                  CFG_MARKER_SCALE_LIMITS)

    ' Blank separator row
    curRow = curRow + 1

    ' Load prove_it_config.csv (Phase 3)
    loadStep = "Load prove_it_config.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\prove_it_config.csv", _
                                  CFG_MARKER_PROVE_IT_CONFIG)

    ' Blank separator row
    curRow = curRow + 1

    ' Load summary_config.csv (Phase 5A)
    loadStep = "Load summary_config.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\summary_config.csv", _
                                  CFG_MARKER_SUMMARY_CONFIG)

    ' Blank separator row
    curRow = curRow + 1

    ' Load chart_registry.csv (Phase 5A)
    loadStep = "Load chart_registry.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\chart_registry.csv", _
                                  CFG_MARKER_CHART_REGISTRY)

    ' Blank separator row
    curRow = curRow + 1

    ' Load exhibit_config.csv (Phase 5A)
    loadStep = "Load exhibit_config.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\exhibit_config.csv", _
                                  CFG_MARKER_EXHIBIT_CONFIG)

    ' Blank separator row
    curRow = curRow + 1

    ' Load display_mode_config.csv (Phase 5A)
    loadStep = "Load display_mode_config.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\display_mode_config.csv", _
                                  CFG_MARKER_DISPLAY_MODE_CONFIG)

    ' Blank separator row
    curRow = curRow + 1

    ' Load print_config.csv (Phase 5B)
    loadStep = "Load print_config.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\print_config.csv", _
                                  CFG_MARKER_PRINT_CONFIG)

    ' Blank separator row
    curRow = curRow + 1

    ' Load data_model_config.csv (Phase 5B)
    loadStep = "Load data_model_config.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\data_model_config.csv", _
                                  CFG_MARKER_DATA_MODEL_CONFIG)

    ' Blank separator row
    curRow = curRow + 1

    ' Load pivot_config.csv (Phase 5B)
    loadStep = "Load pivot_config.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\pivot_config.csv", _
                                  CFG_MARKER_PIVOT_CONFIG)

    ' Blank separator row
    curRow = curRow + 1

    ' Load formula_tab_config.csv (Phase 5C)
    loadStep = "Load formula_tab_config.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\formula_tab_config.csv", _
                                  CFG_MARKER_FORMULA_TAB_CONFIG)

    ' Blank separator row
    curRow = curRow + 1

    ' Load named_range_registry.csv (Phase 5C)
    loadStep = "Load named_range_registry.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\named_range_registry.csv", _
                                  CFG_MARKER_NAMED_RANGE_REGISTRY)

    ' Blank separator row
    curRow = curRow + 1

    ' Load extension_registry.csv (Phase 6A)
    loadStep = "Load extension_registry.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\extension_registry.csv", _
                                  CFG_MARKER_EXTENSION_REGISTRY)

    ' Blank separator row
    curRow = curRow + 1

    ' Load curve_library_config.csv (Phase 6A)
    loadStep = "Load curve_library_config.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\curve_library_config.csv", _
                                  CFG_MARKER_CURVE_LIBRARY)

    ' Blank separator row
    curRow = curRow + 1

    ' Load validation_config.csv (Phase 13)
    loadStep = "Load validation_config.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\validation_config.csv", _
                                  CFG_MARKER_VALIDATION_CONFIG)

    ' Blank separator row
    curRow = curRow + 1

    ' Load health_config.csv (Phase 13)
    loadStep = "Load health_config.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\health_config.csv", _
                                  CFG_MARKER_HEALTH_CONFIG)

    ' Blank separator row
    curRow = curRow + 1

    ' Load branding_config.csv
    loadStep = "Load branding_config.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\branding_config.csv", _
                                  CFG_MARKER_BRANDING_CONFIG)
    curRow = curRow + 1

    ' Load pipeline_config.csv
    loadStep = "Load pipeline_config.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\pipeline_config.csv", _
                                  CFG_MARKER_PIPELINE_CONFIG)
    curRow = curRow + 1

    ' Load msgbox_config.csv
    loadStep = "Load msgbox_config.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\msgbox_config.csv", _
                                  CFG_MARKER_MSGBOX_CONFIG)
    curRow = curRow + 1

    ' Load display_aliases.csv
    loadStep = "Load display_aliases.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\display_aliases.csv", _
                                  CFG_MARKER_DISPLAY_ALIASES)
    curRow = curRow + 1

    ' Load regression_config.csv
    loadStep = "Load regression_config.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\regression_config.csv", _
                                  CFG_MARKER_REGRESSION_CONFIG)
    curRow = curRow + 1

    ' Load diagnostic_config.csv
    loadStep = "Load diagnostic_config.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\diagnostic_config.csv", _
                                  CFG_MARKER_DIAGNOSTIC_CONFIG)
    curRow = curRow + 1

    ' Load config_version.csv
    loadStep = "Load config_version.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\config_version.csv", _
                                  CFG_MARKER_CONFIG_VERSION)
    curRow = curRow + 1

    ' Load button_config.csv
    loadStep = "Load button_config.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\button_config.csv", _
                                  CFG_MARKER_BUTTON_CONFIG)
    curRow = curRow + 1

    ' Load report_templates.csv
    loadStep = "Load report_templates.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\report_templates.csv", _
                                  CFG_MARKER_REPORT_TEMPLATES)
    curRow = curRow + 1

    ' Load workspace_config.csv
    loadStep = "Load workspace_config.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\workspace_config.csv", _
                                  CFG_MARKER_WORKSPACE_CONFIG)
    curRow = curRow + 1

    ' Load lock_config.csv
    loadStep = "Load lock_config.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\lock_config.csv", _
                                  CFG_MARKER_LOCK_CONFIG)
    curRow = curRow + 1

    ' Load config_schema.csv
    loadStep = "Load config_schema.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\config_schema.csv", _
                                  CFG_MARKER_CONFIG_SCHEMA)
    curRow = curRow + 1

    ' Load assumptions_config.csv
    loadStep = "Load assumptions_config.csv"
    curRow = LoadCsvToConfigSheet(wsConfig, curRow, configDir & "\assumptions_config.csv", _
                                  CFG_MARKER_ASSUMPTIONS_CONFIG)
    curRow = curRow + 1

    ' Write report_config settings section (Phase 6A)
    loadStep = "Write report_config"
    wsConfig.Cells(curRow, 1).NumberFormat = "@"
    wsConfig.Cells(curRow, 1).Value = CFG_MARKER_REPORT_CONFIG
    wsConfig.Cells(curRow, 1).Font.Bold = True
    curRow = curRow + 1
    wsConfig.Cells(curRow, 1).NumberFormat = "@"
    wsConfig.Cells(curRow, 1).Value = "Setting"
    wsConfig.Cells(curRow, 2).NumberFormat = "@"
    wsConfig.Cells(curRow, 2).Value = "Value"
    wsConfig.Cells(curRow, 3).NumberFormat = "@"
    wsConfig.Cells(curRow, 3).Value = "Description"
    curRow = curRow + 1
    wsConfig.Cells(curRow, 1).NumberFormat = "@"
    wsConfig.Cells(curRow, 1).Value = "ReportTitle"
    wsConfig.Cells(curRow, 2).NumberFormat = "@"
    wsConfig.Cells(curRow, 2).Value = "RDK Model Report"
    wsConfig.Cells(curRow, 3).NumberFormat = "@"
    wsConfig.Cells(curRow, 3).Value = "Title on cover page"
    curRow = curRow + 1
    wsConfig.Cells(curRow, 1).NumberFormat = "@"
    wsConfig.Cells(curRow, 1).Value = "IncludeCoverPage"
    wsConfig.Cells(curRow, 2).NumberFormat = "@"
    wsConfig.Cells(curRow, 2).Value = "TRUE"
    wsConfig.Cells(curRow, 3).NumberFormat = "@"
    wsConfig.Cells(curRow, 3).Value = "Generate a cover page with TOC"
    curRow = curRow + 1
    wsConfig.Cells(curRow, 1).NumberFormat = "@"
    wsConfig.Cells(curRow, 1).Value = "IncludeProveItSummary"
    wsConfig.Cells(curRow, 2).NumberFormat = "@"
    wsConfig.Cells(curRow, 2).Value = "TRUE"
    wsConfig.Cells(curRow, 3).NumberFormat = "@"
    wsConfig.Cells(curRow, 3).Value = "Add Prove-It check summary to cover page"
    curRow = curRow + 1
    wsConfig.Cells(curRow, 1).NumberFormat = "@"
    wsConfig.Cells(curRow, 1).Value = "IncludeTimestamp"
    wsConfig.Cells(curRow, 2).NumberFormat = "@"
    wsConfig.Cells(curRow, 2).Value = "TRUE"
    wsConfig.Cells(curRow, 3).NumberFormat = "@"
    wsConfig.Cells(curRow, 3).Value = "Add generation timestamp to filename"
    curRow = curRow + 1
    wsConfig.Cells(curRow, 1).NumberFormat = "@"
    wsConfig.Cells(curRow, 1).Value = "OutputDirectory"
    wsConfig.Cells(curRow, 2).NumberFormat = "@"
    wsConfig.Cells(curRow, 2).Value = "output"
    wsConfig.Cells(curRow, 3).NumberFormat = "@"
    wsConfig.Cells(curRow, 3).Value = "Directory for generated reports"
    curRow = curRow + 1

    ' Hide Config sheet - ensure at least one other visible sheet exists first
    loadStep = "Hide Config sheet"
    Dim visibleCount As Long
    visibleCount = 0
    Dim sheetObj As Object
    For Each sheetObj In ThisWorkbook.Sheets
        If sheetObj.Visible = xlSheetVisible And sheetObj.Name <> TAB_CONFIG Then
            visibleCount = visibleCount + 1
        End If
    Next sheetObj

    If visibleCount > 0 Then
        wsConfig.Visible = xlSheetHidden
    End If

    ' Build section position cache now that all sections are loaded
    KernelConfigLoader.BuildSectionCache wsConfig
End Sub


' =============================================================================
' LoadCsvToConfigSheet
' Reads a CSV file and writes its contents to the Config sheet starting at curRow.
' Returns the next available row after writing.
' =============================================================================
Private Function LoadCsvToConfigSheet(ws As Worksheet, startRow As Long, _
                                      csvPath As String, marker As String) As Long
    ' Write section marker (force text to avoid formula interpretation on "===" prefix)
    ws.Cells(startRow, 1).NumberFormat = "@"
    ws.Cells(startRow, 1).Value = marker
    ws.Cells(startRow, 1).Font.Bold = True

    Dim curRow As Long
    curRow = startRow + 1

    If Dir(csvPath) = "" Then
        LoadCsvToConfigSheet = curRow + 1
        Exit Function
    End If

    ' Read entire file as binary
    Dim fileNum As Integer
    fileNum = FreeFile
    Dim fileContent As String
    Dim fileSize As Long

    Open csvPath For Binary Access Read As #fileNum
    fileSize = LOF(fileNum)
    If fileSize = 0 Then
        Close #fileNum
        LoadCsvToConfigSheet = curRow
        Exit Function
    End If
    fileContent = Space$(fileSize)
    Get #fileNum, , fileContent
    Close #fileNum

    ' Normalize line endings
    fileContent = Replace(fileContent, vbCrLf, vbLf)
    fileContent = Replace(fileContent, vbCr, vbLf)

    Dim lines() As String
    lines = Split(fileContent, vbLf)

    ' First pass: count non-empty lines and find max columns
    Dim lineCount As Long
    lineCount = 0
    Dim maxCols As Long
    maxCols = 0
    Dim lineIdx As Long
    For lineIdx = 0 To UBound(lines)
        If Len(Trim(lines(lineIdx))) > 0 Then
            lineCount = lineCount + 1
            Dim tmpFields() As String
            tmpFields = ParseCsvLine(lines(lineIdx))
            If UBound(tmpFields) + 1 > maxCols Then maxCols = UBound(tmpFields) + 1
        End If
    Next lineIdx

    If lineCount = 0 Or maxCols = 0 Then
        LoadCsvToConfigSheet = curRow
        Exit Function
    End If

    ' Second pass: build 2D array
    Dim dataArr() As Variant
    ReDim dataArr(1 To lineCount, 1 To maxCols)
    Dim rowIdx As Long
    rowIdx = 0
    For lineIdx = 0 To UBound(lines)
        If Len(Trim(lines(lineIdx))) > 0 Then
            rowIdx = rowIdx + 1
            Dim fields() As String
            fields = ParseCsvLine(lines(lineIdx))
            Dim colIdx As Long
            For colIdx = 0 To UBound(fields)
                dataArr(rowIdx, colIdx + 1) = fields(colIdx)
            Next colIdx
        End If
    Next lineIdx

    ' Batch write: set entire range to text format, then write array
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(curRow, 1), ws.Cells(curRow + lineCount - 1, maxCols))
    rng.NumberFormat = "@"
    rng.Value = dataArr

    curRow = curRow + lineCount
    LoadCsvToConfigSheet = curRow
End Function


' =============================================================================
' WriteCsvLineToRow
' Parses a CSV line and writes values to the given row.
' Handles quoted fields properly.
' =============================================================================
Private Sub WriteCsvLineToRow(ws As Worksheet, row As Long, lineText As String)
    Dim fields() As String
    fields = ParseCsvLine(lineText)

    Dim colIdx As Long
    For colIdx = 0 To UBound(fields)
        Dim val As String
        val = fields(colIdx)
        ' Write as text to prevent Excel auto-interpretation (AP-07)
        ' Config sheet is storage only -- all values stored as text
        ' prevents "0.0%" being converted to number 0
        If Len(val) > 0 Then
            ws.Cells(row, colIdx + 1).NumberFormat = "@"
            ws.Cells(row, colIdx + 1).Value = val
        End If
    Next colIdx
End Sub


' =============================================================================
' ParseCsvLine
' Parses a CSV line into an array of field values.
' Handles double-quoted fields and embedded commas.
' =============================================================================
Private Function ParseCsvLine(lineText As String) As String()
    Dim result() As String
    Dim fieldCount As Long
    fieldCount = 0

    Dim pos As Long
    pos = 1

    Dim lineLen As Long
    lineLen = Len(lineText)

    ' Pre-count fields for array sizing (max = number of commas + 1)
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
            ' Quoted field
            pos = pos + 1
            Do While pos <= lineLen
                If Mid(lineText, pos, 1) = """" Then
                    If pos < lineLen And Mid(lineText, pos + 1, 1) = """" Then
                        ' Escaped quote
                        fieldVal = fieldVal & """"
                        pos = pos + 2
                    Else
                        ' End of quoted field
                        pos = pos + 1
                        Exit Do
                    End If
                Else
                    fieldVal = fieldVal & Mid(lineText, pos, 1)
                    pos = pos + 1
                End If
            Loop
            ' Skip comma after quoted field
            If pos <= lineLen And Mid(lineText, pos, 1) = "," Then
                pos = pos + 1
            End If
        Else
            ' Unquoted field
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

        ' Bounds check before assignment
        If fieldCount <= UBound(result) Then
            result(fieldCount) = fieldVal
        End If
        fieldCount = fieldCount + 1
    Loop

    ' Handle trailing comma (empty last field)
    If lineLen > 0 And Right(lineText, 1) = "," Then
        If fieldCount <= UBound(result) Then
            result(fieldCount) = ""
        End If
        fieldCount = fieldCount + 1
    End If

    ' Resize to actual count
    If fieldCount > 0 And fieldCount <= maxFields Then
        ReDim Preserve result(0 To fieldCount - 1)
    ElseIf fieldCount = 0 Then
        ReDim result(0 To 0)
    End If

    ParseCsvLine = result
End Function
