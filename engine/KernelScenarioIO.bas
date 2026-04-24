Attribute VB_Name = "KernelScenarioIO"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.

' =============================================================================
' KernelScenarioIO.bas
' Purpose: Export/Import user-editable input cells as portable scenario CSVs.
'          Replaces the workspace subsystem's save-state role. The schema of
'          exportable cells is declared in config/assumptions_schema.csv.
'
'          Export: walks the schema, reads each cell from the live workbook,
'                  writes a scenario CSV to <WorkbookDir>/scenarios/<Name>.csv
'          Import: reads a scenario CSV, validates against the schema, writes
'                  the values back to the cells identified by
'                  (TabName, AssumptionID, Address). RowID lookup on column A
'                  handles row-number shifts; Address disambiguates columns for
'                  RowIDs that span multiple cells (e.g. RBC LOB factor rows,
'                  Staffing Expense per-quarter salary/headcount inputs).
'
'          Every schema row maps to exactly one cell. Quarterly replicator
'          inputs (Col="C" in formula_tab_config) are expanded in the schema
'          CSV to one row per quarterly/annual-total cell across the horizon,
'          so no special runtime handling is needed.
'
' Scenario file format (one row per cell):
'     TabName,AssumptionID,Address,Value
'
' Safety rails:
'   - Cells NOT listed in assumptions_schema.csv are NEVER touched by Import.
'   - Import runs in dry-run preview first if invoked via UI.
'   - All diffs logged via KernelConfig.LogError at SEV_INFO.
'   - No On Error Resume Next wrapping loops. Fails loudly.
' =============================================================================

Private Const MODULE_NAME As String = "KernelScenarioIO"
Private Const SCENARIO_FILE_EXT As String = ".csv"
Private Const SCHEMA_FILE As String = "assumptions_schema.csv"
Private Const SCHEMA_META_FILE As String = "assumptions_schema.meta.csv"
Private Const SCENARIOS_SUBDIR As String = "scenarios"

' Validator thresholds (tunable)
Private Const PCT_MIN As Double = -2#     ' Pct values below this are flagged
Private Const PCT_MAX As Double = 2#      ' Pct values above this are flagged
Private Const COVERAGE_WARN_PCT As Double = 0.9  ' Warn if <90% of schema covered


' =============================================================================
' ExportScenarioUI  (Dashboard button entry point)
' Prompts the user for a scenario name, then calls ExportScenario.
' =============================================================================
Public Sub ExportScenarioUI()
    Dim defaultName As String
    defaultName = GetAssumptionsTabScenarioName()
    If Len(defaultName) = 0 Then defaultName = "Scenario_" & Format(Now, "yyyymmdd_hhnnss")

    Dim name As String
    name = InputBox( _
        "Name this scenario (letters, digits, underscore, hyphen only):" & vbCrLf & _
        "The CSV will be saved to: " & GetScenariosDir(), _
        "Export Scenario", defaultName)
    If Len(Trim(name)) = 0 Then Exit Sub

    name = SanitizeFileName(Trim(name))
    If Len(name) = 0 Then
        MsgBox "Scenario name contains no valid characters.", vbExclamation, "Export Scenario"
        Exit Sub
    End If

    ExportScenario name, silent:=False
End Sub


' =============================================================================
' ImportScenarioUI  (Dashboard button entry point)
' Validates the chosen scenario, shows the report, and imports on confirmation.
' Import blocks automatically if validation has ANY errors.
' =============================================================================
Public Sub ImportScenarioUI()
    Dim scenarioPath As String
    scenarioPath = PickScenarioFile()
    If Len(scenarioPath) = 0 Then Exit Sub  ' user cancelled

    Dim r As Object: Set r = ValidateScenario(scenarioPath)

    If CBool(r("HasErrors")) Then
        MsgBox r("Report") & vbCrLf & vbCrLf & _
               "Import BLOCKED. Fix the errors above (regenerate schema, " & _
               "rebuild tabs, or edit the scenario file) and try again.", _
               vbCritical, "Import Scenario -- Validation Failed"
        Exit Sub
    End If

    Dim prompt As String
    prompt = r("Report") & vbCrLf & vbCrLf
    If CLng(r("WarnCount")) > 0 Then
        prompt = prompt & "Import has " & r("WarnCount") & " warning(s). "
    End If
    prompt = prompt & "A snapshot of the current workbook will be archived " & _
             "before any cell is written." & vbCrLf & vbCrLf & _
             "Proceed with import?"

    Dim resp As VbMsgBoxResult
    resp = MsgBox(prompt, vbOKCancel + vbQuestion, "Import Scenario -- Preview")
    If resp <> vbOK Then Exit Sub

    ImportScenario scenarioPath, dryRun:=False, silent:=False
End Sub


' =============================================================================
' ExportScenario
' Writes the current workbook's assumption values to scenarios/<name>.csv
' =============================================================================
Public Sub ExportScenario(scenarioName As String, Optional silent As Boolean = False)
    Dim scenariosDir As String
    scenariosDir = GetScenariosDir()
    EnsureDirExists scenariosDir

    Dim outPath As String
    outPath = scenariosDir & Application.PathSeparator & _
              SanitizeFileName(scenarioName) & SCENARIO_FILE_EXT

    Dim schema As Collection
    Set schema = LoadSchemaFromDisk()
    If schema.Count = 0 Then
        If Not silent Then MsgBox "Schema is empty or unreadable: " & SchemaPath(), vbCritical
        Exit Sub
    End If

    ' Open for write (ASCII, CRLF)
    Dim fnum As Integer
    fnum = FreeFile
    Open outPath For Output As #fnum
    Print #fnum, """TabName"",""AssumptionID"",""Address"",""Value"""

    Dim tabCache As Object
    Set tabCache = CreateObject("Scripting.Dictionary")

    Dim written As Long: written = 0
    Dim missing As Long: missing = 0
    Dim s As Object

    Dim i As Long
    For i = 1 To schema.Count
        Set s = schema(i)
        Dim tabName As String: tabName = CStr(s("TabName"))
        Dim assumID As String: assumID = CStr(s("AssumptionID"))
        Dim addr As String:    addr = CStr(s("Address"))

        Dim ws As Worksheet
        Set ws = GetSheetCached(tabCache, tabName)
        If ws Is Nothing Then
            missing = missing + 1
            GoTo NextRow
        End If

        ' Resolve cell: prefer RowID lookup (survives row shifts);
        ' fall back to Address only if RowID not found.
        Dim cellRef As Range
        Set cellRef = ResolveCellForExport(ws, assumID, addr)
        If cellRef Is Nothing Then
            missing = missing + 1
            GoTo NextRow
        End If

        Dim val As Variant: val = cellRef.Value
        Dim valStr As String
        valStr = CoerceValueToString(val)

        Print #fnum, CsvQuote(tabName) & "," & CsvQuote(assumID) & "," & _
                     CsvQuote(addr) & "," & CsvQuote(valStr)
        written = written + 1
NextRow:
    Next i

    Close #fnum

    KernelConfig.LogError SEV_INFO, MODULE_NAME, "I-830", _
        "ExportScenario wrote " & written & " value(s) to " & outPath & _
        " (missing: " & missing & ")", ""

    If Not silent Then
        MsgBox "Exported " & written & " assumption value(s) to:" & vbCrLf & outPath & _
               IIf(missing > 0, vbCrLf & vbCrLf & "Skipped " & missing & _
                   " cell(s) not found (missing tab or RowID).", ""), _
               vbInformation, "Export Scenario"
    End If
End Sub


' =============================================================================
' ImportScenario
' Reads a scenario CSV and writes each value into the cell identified by
' (TabName, AssumptionID). Validates every scenario row against the schema:
' rows whose AssumptionID is not in the schema are SKIPPED (never written).
' =============================================================================
Public Sub ImportScenario(scenarioPath As String, _
                          Optional dryRun As Boolean = False, _
                          Optional silent As Boolean = False)
    If Len(Dir(scenarioPath)) = 0 Then
        If Not silent Then MsgBox "Scenario file not found: " & scenarioPath, vbCritical
        Exit Sub
    End If

    ' Safety rail: validate before any write. Block on errors.
    If Not dryRun Then
        Dim v As Object: Set v = ValidateScenario(scenarioPath)
        If CBool(v("HasErrors")) Then
            If Not silent Then
                MsgBox v("Report") & vbCrLf & vbCrLf & _
                       "Import BLOCKED due to validation errors.", _
                       vbCritical, "Import Scenario"
            End If
            KernelConfig.LogError SEV_ERROR, MODULE_NAME, "E-833", _
                "ImportScenario blocked by validation: " & v("ErrorCount") & _
                " error(s) in " & scenarioPath, ""
            Exit Sub
        End If

        ' Auto-archive before any write. Non-negotiable rollback rail.
        ' If archive fails (disk full, permissions, path issue), block the
        ' import rather than leaving the user without a restore point.
        Dim archiveOk As Boolean
        archiveOk = TryArchiveWorkbook()
        If Not archiveOk Then
            If Not silent Then
                MsgBox "Pre-import archive failed. Import aborted to preserve " & _
                       "rollback safety. Resolve the archive failure (check " & _
                       ThisWorkbook.Path & "\Archive folder permissions and " & _
                       "disk space), then retry.", _
                       vbCritical, "Import Scenario"
            End If
            KernelConfig.LogError SEV_ERROR, MODULE_NAME, "E-834", _
                "ImportScenario aborted: pre-import archive failed", scenarioPath
            Exit Sub
        End If
    End If

    Dim schemaIdx As Object
    Set schemaIdx = BuildSchemaIndex()  ' key: "TabName||AssumptionID||Address" -> schema row
    If schemaIdx.Count = 0 Then
        If Not silent Then MsgBox "Schema is empty or unreadable: " & SchemaPath(), vbCritical
        Exit Sub
    End If

    Dim rows As Collection
    Set rows = ReadScenarioFile(scenarioPath)

    Dim tabCache As Object
    Set tabCache = CreateObject("Scripting.Dictionary")

    Dim savedCalc As Long, savedEvents As Boolean
    savedCalc = Application.Calculation
    savedEvents = Application.EnableEvents
    If Not dryRun Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    End If

    Dim written As Long:     written = 0
    Dim skippedSchema As Long: skippedSchema = 0
    Dim skippedTab As Long:  skippedTab = 0
    Dim skippedRowID As Long: skippedRowID = 0
    Dim unchanged As Long:   unchanged = 0

    Dim i As Long
    For i = 1 To rows.Count
        Dim r As Object: Set r = rows(i)
        Dim tabName As String: tabName = CStr(r("TabName"))
        Dim assumID As String: assumID = CStr(r("AssumptionID"))
        Dim addr As String:    addr = CStr(r("Address"))
        Dim newVal As String:  newVal = CStr(r("Value"))

        ' Safety rail: (TabName, AssumptionID, Address) must be in schema. Some
        ' RowIDs span multiple cells (e.g. RBC LOB factor rows -- same RowID at
        ' 6 different columns), so AssumptionID alone is not unique.
        Dim key As String: key = tabName & "||" & assumID & "||" & addr
        If Not schemaIdx.Exists(key) Then
            skippedSchema = skippedSchema + 1
            GoTo NextImport
        End If

        Dim ws As Worksheet
        Set ws = GetSheetCached(tabCache, tabName)
        If ws Is Nothing Then
            skippedTab = skippedTab + 1
            GoTo NextImport
        End If

        Dim cellRef As Range
        Set cellRef = ResolveCellForImport(ws, assumID, addr)
        If cellRef Is Nothing Then
            skippedRowID = skippedRowID + 1
            GoTo NextImport
        End If

        Dim curVal As Variant: curVal = cellRef.Value
        If CStr(curVal) = newVal Then
            unchanged = unchanged + 1
            GoTo NextImport
        End If

        If Not dryRun Then
            WriteValueToCell cellRef, newVal, CStr(schemaIdx(key)("DataType"))
            written = written + 1
        Else
            written = written + 1  ' count planned writes
        End If
NextImport:
    Next i

    If Not dryRun Then
        Application.Calculation = savedCalc
        Application.EnableEvents = savedEvents
        Application.ScreenUpdating = True
        Application.CalculateFull
    End If

    Dim msg As String
    msg = "Import " & IIf(dryRun, "preview", "complete") & ":" & vbCrLf & _
          "  Scenario file:       " & scenarioPath & vbCrLf & _
          "  Rows in scenario:    " & rows.Count & vbCrLf & _
          "  Cells " & IIf(dryRun, "would write", "written") & ":  " & written & vbCrLf & _
          "  Unchanged:           " & unchanged & vbCrLf & _
          "  Skipped (not in schema): " & skippedSchema & vbCrLf & _
          "  Skipped (missing tab):   " & skippedTab & vbCrLf & _
          "  Skipped (RowID not found): " & skippedRowID

    KernelConfig.LogError SEV_INFO, MODULE_NAME, "I-831", _
        "ImportScenario " & IIf(dryRun, "dry-run ", "") & "result: " & _
        "written=" & written & " unchanged=" & unchanged & _
        " skippedSchema=" & skippedSchema & " skippedTab=" & skippedTab & _
        " skippedRowID=" & skippedRowID, scenarioPath

    If Not silent Then MsgBox msg, vbInformation, "Import Scenario"
End Sub


' =============================================================================
' ImportScenarioDryRun
' Thin wrapper around ValidateScenario -- returns just the formatted report
' string. Kept for backwards compatibility with any callers that expected a
' single-string return.
' =============================================================================
Public Function ImportScenarioDryRun(scenarioPath As String) As String
    Dim r As Object: Set r = ValidateScenario(scenarioPath)
    ImportScenarioDryRun = CStr(r("Report"))
End Function


' =============================================================================
' ValidateScenario
' Runs the full pre-import checklist and returns a structured result:
'   HasErrors     (Boolean) -- at least one ERROR-severity finding
'   ErrorCount    (Long)
'   WarnCount     (Long)
'   Report        (String)  -- multi-line human-readable report
'   RowCount      (Long)
'   WouldWrite    (Long)
'   WouldUnchanged (Long)
'
' Checks performed:
'   1. File exists and is readable
'   2. CSV has expected header (TabName, AssumptionID, Address, Value)
'   3. Every row has 4 fields (CSV-structure integrity)
'   4. Every row resolves on (TabName, AssumptionID, Address) in the schema
'   5. Target tab exists in current workbook
'   6. Target cell resolvable (RowID found OR Address valid)
'   7. DataType sanity per schema entry:
'      - Pct values within [PCT_MIN, PCT_MAX]
'      - Numeric values parse as numbers
'      - Text values don't have leading = (formula-injection protection)
'   8. Target cell is not locked on a protected sheet
'   9. Coverage check: scenario covers >=COVERAGE_WARN_PCT of schema
'  10. Schema drift: compare live Input-row count to schema metadata
' =============================================================================
Public Function ValidateScenario(scenarioPath As String) As Object
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    result("HasErrors") = False
    result("ErrorCount") = CLng(0)
    result("WarnCount") = CLng(0)
    result("RowCount") = CLng(0)
    result("WouldWrite") = CLng(0)
    result("WouldUnchanged") = CLng(0)
    result("Report") = ""

    Dim errs As New Collection, warns As New Collection

    ' (1) File exists
    If Len(Dir(scenarioPath)) = 0 Then
        errs.Add "File not found: " & scenarioPath
        result("ErrorCount") = 1
        result("HasErrors") = True
        result("Report") = FormatValidationReport(scenarioPath, errs, warns, result)
        Set ValidateScenario = result
        Exit Function
    End If

    ' (10) Schema drift check -- run once, non-fatal warning
    Dim driftMsg As String: driftMsg = CheckSchemaDrift()
    If Len(driftMsg) > 0 Then warns.Add "SCHEMA DRIFT: " & driftMsg

    ' Load schema index
    Dim schemaIdx As Object: Set schemaIdx = BuildSchemaIndex()
    If schemaIdx.Count = 0 Then
        errs.Add "Schema is empty or unreadable: " & SchemaPath()
        result("ErrorCount") = errs.Count
        result("WarnCount") = warns.Count
        result("HasErrors") = True
        result("Report") = FormatValidationReport(scenarioPath, errs, warns, result)
        Set ValidateScenario = result
        Exit Function
    End If

    ' (2-3) Header + row-shape check (stream-parse to catch malformed lines before
    ' loading full rows)
    Dim fnum As Integer: fnum = FreeFile
    Open scenarioPath For Input As #fnum
    Dim ln As String
    Dim lineNum As Long: lineNum = 0
    Dim headerValid As Boolean: headerValid = False
    Dim malformed As Long: malformed = 0
    Do While Not EOF(fnum)
        Line Input #fnum, ln
        lineNum = lineNum + 1
        If Len(Trim(ln)) = 0 Then GoTo NxtParseLine
        Dim flds As Variant: flds = ParseCsvLine(ln)
        Dim nf As Long: nf = UBound(flds) - LBound(flds) + 1
        If lineNum = 1 Then
            If nf = 4 And _
               StrComp(CStr(flds(0)), "TabName", vbTextCompare) = 0 And _
               StrComp(CStr(flds(1)), "AssumptionID", vbTextCompare) = 0 And _
               StrComp(CStr(flds(2)), "Address", vbTextCompare) = 0 And _
               StrComp(CStr(flds(3)), "Value", vbTextCompare) = 0 Then
                headerValid = True
            End If
        Else
            If nf <> 4 Then malformed = malformed + 1
        End If
NxtParseLine:
    Loop
    Close #fnum

    If Not headerValid Then
        errs.Add "Header row must be exactly: TabName, AssumptionID, Address, Value"
    End If
    If malformed > 0 Then
        errs.Add malformed & " malformed row(s) found (expected 4 fields each)"
    End If

    ' If structural errors, stop here
    If errs.Count > 0 Then
        result("ErrorCount") = errs.Count
        result("WarnCount") = warns.Count
        result("HasErrors") = True
        result("Report") = FormatValidationReport(scenarioPath, errs, warns, result)
        Set ValidateScenario = result
        Exit Function
    End If

    ' (4-9) Per-row semantic checks
    Dim rows As Collection: Set rows = ReadScenarioFile(scenarioPath)
    result("RowCount") = CLng(rows.Count)

    Dim tabCache As Object: Set tabCache = CreateObject("Scripting.Dictionary")

    Dim orphanCount As Long: orphanCount = 0
    Dim missingTab As Long:  missingTab = 0
    Dim missingRow As Long:  missingRow = 0
    Dim lockedCells As Long: lockedCells = 0
    Dim pctRangeBad As Long: pctRangeBad = 0
    Dim numParseBad As Long: numParseBad = 0
    Dim textFormula As Long: textFormula = 0
    Dim wouldWrite As Long:  wouldWrite = 0
    Dim wouldSame As Long:   wouldSame = 0

    Dim firstOrphanSample As String: firstOrphanSample = ""
    Dim firstTabSample As String:    firstTabSample = ""
    Dim firstRowSample As String:    firstRowSample = ""

    Dim i As Long
    For i = 1 To rows.Count
        Dim r As Object: Set r = rows(i)
        Dim tabName As String: tabName = CStr(r("TabName"))
        Dim assumID As String: assumID = CStr(r("AssumptionID"))
        Dim addr As String:    addr = CStr(r("Address"))
        Dim newVal As String:  newVal = CStr(r("Value"))

        Dim key As String: key = tabName & "||" & assumID & "||" & addr
        If Not schemaIdx.Exists(key) Then
            orphanCount = orphanCount + 1
            If Len(firstOrphanSample) = 0 Then
                firstOrphanSample = tabName & " / " & assumID & " / " & addr
            End If
            GoTo NextValRow
        End If

        Dim ws As Worksheet
        Set ws = GetSheetCached(tabCache, tabName)
        If ws Is Nothing Then
            missingTab = missingTab + 1
            If Len(firstTabSample) = 0 Then firstTabSample = tabName
            GoTo NextValRow
        End If

        Dim cellRef As Range
        Set cellRef = ResolveCellForImport(ws, assumID, addr)
        If cellRef Is Nothing Then
            missingRow = missingRow + 1
            If Len(firstRowSample) = 0 Then firstRowSample = tabName & "!" & assumID
            GoTo NextValRow
        End If

        ' Data-type sanity
        Dim s As Object: Set s = schemaIdx(key)
        Dim dt As String: dt = UCase(Trim(CStr(s("DataType"))))
        If dt = "PCT" And IsNumeric(newVal) Then
            If CDbl(newVal) < PCT_MIN Or CDbl(newVal) > PCT_MAX Then
                pctRangeBad = pctRangeBad + 1
            End If
        ElseIf (dt = "NUMBER" Or dt = "CURRENCY") And Len(newVal) > 0 Then
            If Not IsNumeric(newVal) Then numParseBad = numParseBad + 1
        ElseIf dt = "TEXT" And Len(newVal) > 0 Then
            Dim first As String: first = Left(newVal, 1)
            If first = "=" Or first = "+" Or first = "@" Then
                textFormula = textFormula + 1
            End If
        End If

        ' Cell-lock check
        If cellRef.Locked And ws.ProtectContents Then
            lockedCells = lockedCells + 1
        End If

        ' Change preview
        If CStr(cellRef.Value) = newVal Then
            wouldSame = wouldSame + 1
        Else
            wouldWrite = wouldWrite + 1
        End If
NextValRow:
    Next i

    result("WouldWrite") = CLng(wouldWrite)
    result("WouldUnchanged") = CLng(wouldSame)

    ' Accumulate findings by severity
    If orphanCount > 0 Then
        errs.Add orphanCount & " row(s) not in schema " & _
                 "(TabName/AssumptionID/Address triple not found). " & _
                 "First: " & firstOrphanSample
    End If
    If missingTab > 0 Then
        errs.Add missingTab & " row(s) reference missing tab(s). First: " & firstTabSample
    End If
    If missingRow > 0 Then
        errs.Add missingRow & " row(s) reference a RowID that can't be " & _
                 "found on column A of the target tab. First: " & firstRowSample
    End If
    If lockedCells > 0 Then
        errs.Add lockedCells & " target cell(s) are locked on a protected sheet"
    End If
    If numParseBad > 0 Then
        warns.Add numParseBad & " Number/Currency value(s) are not numeric"
    End If
    If pctRangeBad > 0 Then
        warns.Add pctRangeBad & " Pct value(s) outside [" & PCT_MIN & "," & PCT_MAX & "]"
    End If
    If textFormula > 0 Then
        warns.Add textFormula & " Text value(s) start with =, +, or @ " & _
                  "(will be written with NumberFormat=@ to prevent formula injection)"
    End If

    ' Coverage check
    Dim coverage As Double
    coverage = (rows.Count - orphanCount) / schemaIdx.Count
    If coverage < COVERAGE_WARN_PCT Then
        warns.Add "Coverage " & Format(coverage, "0.0%") & " of schema " & _
                  "(" & (rows.Count - orphanCount) & " of " & schemaIdx.Count & _
                  "). Cells not in the scenario will be LEFT AS-IS (not reset)."
    End If

    result("ErrorCount") = CLng(errs.Count)
    result("WarnCount") = CLng(warns.Count)
    result("HasErrors") = (errs.Count > 0)
    result("Report") = FormatValidationReport(scenarioPath, errs, warns, result)

    Set ValidateScenario = result
End Function


' =============================================================================
' CheckSchemaDrift
' Compares live Config-sheet input-row count against the SchemaRowCount /
' InputRowCount recorded in assumptions_schema.meta.csv at generation time.
' Returns empty string if no drift, or a short diagnostic message.
' =============================================================================
Private Function CheckSchemaDrift() As String
    Dim meta As Object: Set meta = LoadSchemaMeta()
    If meta.Count = 0 Then
        CheckSchemaDrift = "metadata file missing (" & SchemaMetaPath() & _
                           "). Run scripts/regen_assumptions_schema.py to generate."
        Exit Function
    End If

    Dim expectedSchemaCount As Long
    If meta.Exists("SchemaRowCount") Then
        expectedSchemaCount = CLng(meta("SchemaRowCount"))
    Else
        expectedSchemaCount = 0
    End If

    Dim actualSchemaCount As Long
    actualSchemaCount = CountSchemaRows()

    If expectedSchemaCount > 0 And _
       Abs(expectedSchemaCount - actualSchemaCount) > 0 Then
        CheckSchemaDrift = "schema row count mismatch (meta says " & _
                           expectedSchemaCount & ", actual file has " & _
                           actualSchemaCount & "). Re-run regen script."
    Else
        CheckSchemaDrift = ""
    End If
End Function


' =============================================================================
' HELPERS
' =============================================================================

Private Function GetScenariosDir() As String
    GetScenariosDir = ThisWorkbook.Path & Application.PathSeparator & SCENARIOS_SUBDIR
End Function

Private Function SchemaPath() As String
    SchemaPath = ThisWorkbook.Path & Application.PathSeparator & "config" & _
                 Application.PathSeparator & SCHEMA_FILE
End Function

Private Function SchemaMetaPath() As String
    SchemaMetaPath = ThisWorkbook.Path & Application.PathSeparator & "config" & _
                     Application.PathSeparator & SCHEMA_META_FILE
End Function

Private Function LoadSchemaMeta() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim path As String: path = SchemaMetaPath()
    If Len(Dir(path)) = 0 Then
        Set LoadSchemaMeta = d
        Exit Function
    End If

    Dim fnum As Integer: fnum = FreeFile
    Open path For Input As #fnum
    Dim ln As String
    Dim header As Variant: header = Empty
    Do While Not EOF(fnum)
        Line Input #fnum, ln
        If Len(Trim(ln)) = 0 Then GoTo NxtMeta
        Dim flds As Variant: flds = ParseCsvLine(ln)
        If IsEmpty(header) Then
            header = flds
            GoTo NxtMeta
        End If
        ' Header is Key,Value -- simple flat mapping
        If UBound(flds) >= 1 Then
            d(CStr(flds(0))) = CStr(flds(1))
        End If
NxtMeta:
    Loop
    Close #fnum
    Set LoadSchemaMeta = d
End Function

Private Function CountSchemaRows() As Long
    Dim path As String: path = SchemaPath()
    If Len(Dir(path)) = 0 Then CountSchemaRows = 0: Exit Function

    Dim n As Long: n = 0
    Dim fnum As Integer: fnum = FreeFile
    Open path For Input As #fnum
    Dim ln As String
    Dim isHeader As Boolean: isHeader = True
    Do While Not EOF(fnum)
        Line Input #fnum, ln
        If Len(Trim(ln)) = 0 Then GoTo NxtCnt
        If isHeader Then
            isHeader = False
            GoTo NxtCnt
        End If
        n = n + 1
NxtCnt:
    Loop
    Close #fnum
    CountSchemaRows = n
End Function

Private Function FormatValidationReport(scenarioPath As String, _
                                        errs As Collection, _
                                        warns As Collection, _
                                        result As Object) As String
    Dim out As String
    out = "Scenario validation -- " & scenarioPath & vbCrLf
    out = out & String(78, "-") & vbCrLf

    out = out & "  Rows in scenario:       " & result("RowCount") & vbCrLf
    out = out & "  Cells that would write: " & result("WouldWrite") & vbCrLf
    out = out & "  Cells unchanged:        " & result("WouldUnchanged") & vbCrLf
    out = out & "  Errors:                 " & errs.Count & vbCrLf
    out = out & "  Warnings:               " & warns.Count & vbCrLf

    If errs.Count > 0 Then
        out = out & vbCrLf & "ERRORS (import blocked):" & vbCrLf
        Dim i As Long
        For i = 1 To errs.Count
            out = out & "  - " & errs(i) & vbCrLf
        Next i
    End If

    If warns.Count > 0 Then
        out = out & vbCrLf & "WARNINGS (import allowed):" & vbCrLf
        Dim j As Long
        For j = 1 To warns.Count
            out = out & "  - " & warns(j) & vbCrLf
        Next j
    End If

    If errs.Count = 0 And warns.Count = 0 Then
        out = out & vbCrLf & "All checks passed." & vbCrLf
    End If

    FormatValidationReport = out
End Function

Private Sub EnsureDirExists(path As String)
    If Dir(path, vbDirectory) = "" Then MkDir path
End Sub

' Calls KernelWorkspaceExt.ArchiveWorkbookNow and returns True on success.
' Unlike ArchiveWorkbookNow (which shows a MsgBox on completion), this helper
' runs silently -- the calling Import flow handles user messaging itself.
Private Function TryArchiveWorkbook() As Boolean
    On Error GoTo ArchFail
    If ThisWorkbook.Path = "" Then
        TryArchiveWorkbook = False
        Exit Function
    End If

    Dim archiveDir As String
    archiveDir = ThisWorkbook.Path & Application.PathSeparator & "Archive"
    If Dir(archiveDir, vbDirectory) = "" Then MkDir archiveDir

    Dim base As String: base = ThisWorkbook.Name
    Dim ext As String: ext = ".xlsm"
    If InStrRev(base, ".") > 0 Then
        ext = Mid(base, InStrRev(base, "."))
        base = Left(base, InStrRev(base, ".") - 1)
    End If

    Dim target As String
    target = archiveDir & Application.PathSeparator & base & "_" & _
             Format(Now, "yyyymmdd_hhnnss") & ext

    Application.ScreenUpdating = False
    ThisWorkbook.SaveCopyAs target
    Application.ScreenUpdating = True

    KernelConfig.LogError SEV_INFO, MODULE_NAME, "I-835", _
        "Pre-import archive saved: " & target, ""
    TryArchiveWorkbook = True
    Exit Function

ArchFail:
    Application.ScreenUpdating = True
    KernelConfig.LogError SEV_ERROR, MODULE_NAME, "E-836", _
        "Pre-import archive failed: " & Err.Description, ""
    TryArchiveWorkbook = False
End Function

Private Function SanitizeFileName(s As String) As String
    Dim out As String: out = ""
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or _
           (ch >= "0" And ch <= "9") Or ch = "_" Or ch = "-" Then
            out = out & ch
        End If
    Next i
    SanitizeFileName = out
End Function

Private Function GetAssumptionsTabScenarioName() As String
    Dim v As Variant
    On Error Resume Next
    v = ThisWorkbook.Sheets("Assumptions").Range("$C$4").Value
    On Error GoTo 0
    If IsError(v) Or IsEmpty(v) Then GetAssumptionsTabScenarioName = "" _
    Else GetAssumptionsTabScenarioName = CStr(v)
End Function

Private Function PickScenarioFile() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select Scenario CSV"
    fd.InitialFileName = GetScenariosDir() & Application.PathSeparator
    fd.Filters.Clear
    fd.Filters.Add "Scenario CSV", "*.csv"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then
        PickScenarioFile = fd.SelectedItems(1)
    Else
        PickScenarioFile = ""
    End If
End Function

Private Function GetSheetCached(cache As Object, tabName As String) As Worksheet
    If cache.Exists(tabName) Then
        If TypeName(cache(tabName)) = "Nothing" Then
            Set GetSheetCached = Nothing
        Else
            Set GetSheetCached = cache(tabName)
        End If
        Exit Function
    End If
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(tabName)
    On Error GoTo 0
    If ws Is Nothing Then
        cache.Add tabName, Nothing
        Set GetSheetCached = Nothing
    Else
        Set cache(tabName) = ws
        Set GetSheetCached = ws
    End If
End Function

' Locate the cell that represents <assumID> on <ws>. First tries Column A
' RowID lookup (stable across row shifts); falls back to the Address hint
' if the RowID can't be found in column A.
Private Function ResolveCellForExport(ws As Worksheet, assumID As String, _
                                      addr As String) As Range
    Dim r As Long: r = FindRowIDOnTab(ws, assumID)
    If r > 0 Then
        ' RowID found -- derive column from the Address hint (same col letter)
        Set ResolveCellForExport = ResolveByRowAndAddressCol(ws, r, addr)
        Exit Function
    End If
    ' Fallback: use Address directly
    On Error Resume Next
    Set ResolveCellForExport = ws.Range(addr)
    On Error GoTo 0
End Function

Private Function ResolveCellForImport(ws As Worksheet, assumID As String, _
                                      addr As String) As Range
    Set ResolveCellForImport = ResolveCellForExport(ws, assumID, addr)
End Function

' Given a row number and an Address hint like "$C$123", return the cell at
' (row, col-of-address). If Address is invalid, fall back to col 3.
Private Function ResolveByRowAndAddressCol(ws As Worksheet, rowNum As Long, _
                                           addr As String) As Range
    Dim colNum As Long: colNum = ColNumFromAddress(addr)
    If colNum <= 0 Then colNum = 3  ' conventional "value" column
    Set ResolveByRowAndAddressCol = ws.Cells(rowNum, colNum)
End Function

Private Function ColNumFromAddress(addr As String) As Long
    ' Parses "$C$5" or "C5" -> 3. Returns 0 on failure.
    Dim s As String: s = Replace(addr, "$", "")
    Dim i As Long, letters As String: letters = ""
    For i = 1 To Len(s)
        Dim ch As String: ch = UCase(Mid(s, i, 1))
        If ch >= "A" And ch <= "Z" Then
            letters = letters & ch
        Else
            Exit For
        End If
    Next i
    If Len(letters) = 0 Then ColNumFromAddress = 0: Exit Function
    Dim total As Long: total = 0
    For i = 1 To Len(letters)
        total = total * 26 + (Asc(Mid(letters, i, 1)) - Asc("A") + 1)
    Next i
    ColNumFromAddress = total
End Function

Private Function FindRowIDOnTab(ws As Worksheet, rowID As String) As Long
    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 1 To lastR
        If StrComp(Trim(CStr(ws.Cells(r, 1).Value)), rowID, vbTextCompare) = 0 Then
            FindRowIDOnTab = r
            Exit Function
        End If
    Next r
    FindRowIDOnTab = 0
End Function

' Schema: returns Collection of Dictionary objects with keys TabName,
' AssumptionID, Address, Section, DataType, DefaultValue, Description
Private Function LoadSchemaFromDisk() As Collection
    Dim result As New Collection
    Dim path As String: path = SchemaPath()
    If Len(Dir(path)) = 0 Then
        Set LoadSchemaFromDisk = result
        Exit Function
    End If

    Dim fnum As Integer: fnum = FreeFile
    Open path For Input As #fnum
    Dim ln As String
    Dim header As Variant: header = Empty
    Do While Not EOF(fnum)
        Line Input #fnum, ln
        If Len(Trim(ln)) = 0 Then GoTo NxtLine
        Dim flds As Variant: flds = ParseCsvLine(ln)
        If IsEmpty(header) Then
            header = flds
            GoTo NxtLine
        End If
        Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
        Dim j As Long
        For j = LBound(header) To UBound(header)
            If j <= UBound(flds) Then
                d(CStr(header(j))) = CStr(flds(j))
            Else
                d(CStr(header(j))) = ""
            End If
        Next j
        result.Add d
NxtLine:
    Loop
    Close #fnum
    Set LoadSchemaFromDisk = result
End Function

Private Function BuildSchemaIndex() As Object
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    Dim schema As Collection: Set schema = LoadSchemaFromDisk()
    Dim i As Long
    For i = 1 To schema.Count
        Dim s As Object: Set s = schema(i)
        ' Key on (Tab, AssumptionID, Address). AssumptionID alone is NOT unique
        ' -- some RowIDs (e.g. RBC LOB factor rows) span multiple columns on
        ' the same tab-row. Address disambiguates.
        Dim key As String
        key = CStr(s("TabName")) & "||" & CStr(s("AssumptionID")) & "||" & CStr(s("Address"))
        If Not idx.Exists(key) Then Set idx(key) = s
    Next i
    Set BuildSchemaIndex = idx
End Function

Private Function ReadScenarioFile(path As String) As Collection
    Dim result As New Collection
    Dim fnum As Integer: fnum = FreeFile
    Open path For Input As #fnum
    Dim ln As String
    Dim header As Variant: header = Empty
    Do While Not EOF(fnum)
        Line Input #fnum, ln
        If Len(Trim(ln)) = 0 Then GoTo NxtLine
        Dim flds As Variant: flds = ParseCsvLine(ln)
        If IsEmpty(header) Then
            header = flds
            GoTo NxtLine
        End If
        Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
        Dim j As Long
        For j = LBound(header) To UBound(header)
            If j <= UBound(flds) Then
                d(CStr(header(j))) = CStr(flds(j))
            Else
                d(CStr(header(j))) = ""
            End If
        Next j
        result.Add d
NxtLine:
    Loop
    Close #fnum
    Set ReadScenarioFile = result
End Function

' Simple CSV parser -- handles "quoted fields", embedded commas, and escaped ""
Private Function ParseCsvLine(ln As String) As Variant
    Dim fields() As String
    ReDim fields(0 To 0)
    Dim cur As String: cur = ""
    Dim inQuote As Boolean: inQuote = False
    Dim i As Long, ch As String
    For i = 1 To Len(ln)
        ch = Mid(ln, i, 1)
        If inQuote Then
            If ch = """" Then
                If i < Len(ln) And Mid(ln, i + 1, 1) = """" Then
                    cur = cur & """"
                    i = i + 1
                Else
                    inQuote = False
                End If
            Else
                cur = cur & ch
            End If
        Else
            If ch = """" Then
                inQuote = True
            ElseIf ch = "," Then
                fields(UBound(fields)) = cur
                ReDim Preserve fields(0 To UBound(fields) + 1)
                cur = ""
            Else
                cur = cur & ch
            End If
        End If
    Next i
    fields(UBound(fields)) = cur
    ParseCsvLine = fields
End Function

Private Function CsvQuote(s As String) As String
    CsvQuote = """" & Replace(s, """", """""") & """"
End Function

Private Function CoerceValueToString(v As Variant) As String
    If IsError(v) Then
        CoerceValueToString = "#ERR"
    ElseIf IsNull(v) Or IsEmpty(v) Then
        CoerceValueToString = ""
    ElseIf IsNumeric(v) Then
        CoerceValueToString = CStr(v)
    Else
        CoerceValueToString = CStr(v)
    End If
End Function

Private Sub WriteValueToCell(cellRef As Range, valStr As String, dataType As String)
    ' AP-07 / AP-50 safety: strings starting with =,+,-,@ are treated by Excel as
    ' formulas unless NumberFormat is @. For Text types, force @ format.
    Dim dt As String: dt = UCase(Trim(dataType))
    If dt = "TEXT" Then
        cellRef.NumberFormat = "@"
        cellRef.Value = valStr
    ElseIf IsNumeric(valStr) Then
        cellRef.Value = CDbl(valStr)
    Else
        ' Fall through: write as-is; guard against formula-injection
        If Len(valStr) > 0 Then
            Dim first As String: first = Left(valStr, 1)
            If first = "=" Or first = "+" Or first = "-" Or first = "@" Then
                cellRef.NumberFormat = "@"
            End If
        End If
        cellRef.Value = valStr
    End If
End Sub

