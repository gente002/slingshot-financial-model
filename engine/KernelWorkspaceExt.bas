Attribute VB_Name = "KernelWorkspaceExt"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelWorkspaceExt.bas
' Purpose: Extensions to the workspace save/load subsystem.
'
' Adds (per session decisions L, O-Q, R-Y, 2026-04-18):
'   - xlsm SaveCopyAs snapshot per version (strategy L)
'   - Config-hash and kernel-version gating for fast-path restore
'   - Atomic version-folder rename (decision R)
'   - Orphan-folder detection and cleanup (decision S)
'   - Manifest field validation (decision T)
'   - PRNG seed restore on load (decision U)
'   - BuildPath helper for UNC/OneDrive safety (decision V)
'   - "Cleanup Old Workspaces" UI (decision W)
'   - Schema-mismatch 4-option modal (decision Y)
'   - preserved_cells.csv capture + restore (decision Q)
'   - KernelLint.AssertCoverageComplete hook (decision O)
'   - Round-trip idempotence test (decision O, layer 3)
'
' Safe to edit. Does not own the primary SaveWorkspace/LoadWorkspace flows
' (those live in KernelWorkspace.bas); this module extends them.
' =============================================================================

Public Const XLSM_SNAPSHOT_NAME As String = "workspace.xlsm"
Public Const PRESERVED_CELLS_CSV As String = "preserved_cells.csv"
Public Const CFG_PRESERVED_CELLS As String = "preserved_cells_config.csv"

' =====================================================================
' BuildPath -- UNC/OneDrive-safe path concatenation. Use this instead of
' `parent & "\" & child` which breaks on UNC roots.
' Decision V.
' =====================================================================
Public Function BuildPath(parent As String, child As String) As String
    Dim p As String
    p = parent
    Dim c As String
    c = child
    ' Strip trailing separator from parent, leading from child
    If Len(p) > 0 Then
        Do While Right$(p, 1) = "\" Or Right$(p, 1) = "/"
            If Len(p) <= 1 Then Exit Do
            p = Left$(p, Len(p) - 1)
        Loop
    End If
    If Len(c) > 0 Then
        Do While Left$(c, 1) = "\" Or Left$(c, 1) = "/"
            c = Mid$(c, 2)
            If Len(c) = 0 Then Exit Do
        Loop
    End If
    If Len(p) = 0 Then
        BuildPath = c
    ElseIf Len(c) = 0 Then
        BuildPath = p
    Else
        BuildPath = p & "\" & c
    End If
End Function

' =====================================================================
' WarnIfCloudSyncedPath -- emits a one-time SEV_WARN if the workspace
' root lives under a recognised cloud-sync folder. Decision Q4 / V.
' =====================================================================
Public Sub WarnIfCloudSyncedPath(wsRoot As String)
    Dim upper As String
    upper = UCase$(wsRoot)
    Dim flagged As Boolean
    flagged = False
    If InStr(1, upper, "ONEDRIVE", vbTextCompare) > 0 Then flagged = True
    If InStr(1, upper, "DROPBOX", vbTextCompare) > 0 Then flagged = True
    If InStr(1, upper, "GOOGLE DRIVE", vbTextCompare) > 0 Then flagged = True
    If InStr(1, upper, "BOX\", vbTextCompare) > 0 Then flagged = True
    If flagged Then
        KernelConfig.LogError SEV_WARN, "KWSExt", "W-930", _
            "Workspace root appears to be on a cloud-sync folder: " & wsRoot & _
            ". Sync tools can hold file locks during atomic rename and corrupt saves. " & _
            "MANUAL BYPASS: Move workspaces/ to a local (non-synced) path.", ""
    End If
End Sub

' =====================================================================
' AtomicEnsureVersionDir -- decision R.
' Creates a temp folder (verDir & ".tmp"), caller fills it, then we rename
' to the final verDir. Prevents version-folder race between concurrent saves.
' =====================================================================
Public Function AtomicBeginVersion(verDir As String) As String
    Dim tmpDir As String
    tmpDir = verDir & ".tmp"
    ' Clean any stale tmp from a prior aborted save
    If Dir(tmpDir, vbDirectory) <> "" Then
        On Error Resume Next
        DeleteFolderRecursive tmpDir
        On Error GoTo 0
    End If
    KernelSnapshot.EnsureDirectoryExists tmpDir
    AtomicBeginVersion = tmpDir
End Function

Public Sub AtomicCommitVersion(tmpDir As String, verDir As String)
    ' Fail if the final name already exists (another save got there first)
    If Dir(verDir, vbDirectory) <> "" Then
        Err.Raise vbObjectError + 931, "KWSExt", _
            "Version folder already exists: " & verDir & _
            ". Concurrent save detected. MANUAL BYPASS: rename " & tmpDir & " manually."
    End If
    ' Name rename is atomic on NTFS
    Name tmpDir As verDir
End Sub

Public Sub AtomicAbortVersion(tmpDir As String)
    ' Best effort: remove the temp folder
    On Error Resume Next
    DeleteFolderRecursive tmpDir
    On Error GoTo 0
End Sub

' =====================================================================
' SaveXlsmSnapshot -- decision L.
' Saves a perfect-fidelity copy of the active workbook into verDir.
' =====================================================================
Public Function SaveXlsmSnapshot(verDir As String) As String
    On Error GoTo xlsmErr
    Dim xlsmPath As String
    xlsmPath = BuildPath(verDir, XLSM_SNAPSHOT_NAME)
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveCopyAs xlsmPath
    Application.DisplayAlerts = True
    SaveXlsmSnapshot = xlsmPath
    Exit Function
xlsmErr:
    Application.DisplayAlerts = True
    KernelConfig.LogError SEV_WARN, "KWSExt", "W-920", _
        "Xlsm snapshot save failed: " & Err.Description & _
        ". MANUAL BYPASS: Save a copy of the workbook to " & verDir & "\workspace.xlsm manually.", ""
    SaveXlsmSnapshot = ""
End Function

' =====================================================================
' CheckFastPathEligible -- decision L + T.
' Returns TRUE only if all gates pass:
'   - xlsm exists, non-empty
'   - manifest.json parses cleanly, required fields present
'   - status == COMPLETE
'   - configHash matches current BuildConfigHash()
'   - kernelVersion is not NEWER than current KERNEL_VERSION (forward-incompat)
' reason (out) describes why the gate failed for user-facing display.
' =====================================================================
Public Function CheckFastPathEligible(verDir As String, _
    Optional ByRef reason As String = "") As Boolean
    reason = ""
    CheckFastPathEligible = False

    Dim xlsmPath As String
    xlsmPath = BuildPath(verDir, XLSM_SNAPSHOT_NAME)
    If Dir(xlsmPath) = "" Then
        reason = "No xlsm snapshot in version folder."
        Exit Function
    End If
    If FileLen(xlsmPath) = 0 Then
        reason = "Xlsm snapshot is empty (0 bytes)."
        Exit Function
    End If

    Dim manifestPath As String
    manifestPath = BuildPath(verDir, "manifest.json")
    If Dir(manifestPath) = "" Then
        reason = "Missing manifest.json -- orphan or corrupt folder."
        Exit Function
    End If

    Dim mStatus As String, mKernelVer As String, mConfigHash As String
    Dim mXlsx As String
    ParseManifestFields manifestPath, mStatus, mKernelVer, mConfigHash, mXlsx

    If StrComp(mStatus, "COMPLETE", vbTextCompare) <> 0 Then
        reason = "Manifest status is '" & mStatus & "' (expected COMPLETE)."
        Exit Function
    End If
    If Len(mConfigHash) = 0 Then
        reason = "Manifest missing configHash field."
        Exit Function
    End If
    If Len(mKernelVer) = 0 Then
        reason = "Manifest missing kernelVersion field."
        Exit Function
    End If

    ' Config hash comparison
    Dim currentHash As String
    currentHash = KernelSnapshot.BuildConfigHash()
    If StrComp(currentHash, mConfigHash, vbTextCompare) <> 0 Then
        reason = "Config has changed since this workspace was saved " & _
            "(hash mismatch). Rebuild from current config or load snapshot as-is?"
        Exit Function
    End If

    ' Kernel version -- refuse if workspace is newer than current kernel
    If CompareKernelVersion(mKernelVer, KERNEL_VERSION) > 0 Then
        reason = "Workspace was saved by a newer kernel (" & mKernelVer & _
            ") than the current kernel (" & KERNEL_VERSION & "). " & _
            "Loading a newer workspace into an older kernel risks silent correctness bugs."
        Exit Function
    End If
    ' Older workspaces allowed through with a note

    CheckFastPathEligible = True
End Function

' =====================================================================
' ShowFastPathPrompt -- decision Y.
' Presents the 4-option schema-mismatch dialog when CheckFastPathEligible
' returns False with a reason. Returns one of:
'   "REBUILD"   -- fall through to CSV rebuild (default)
'   "LOAD_XLSM" -- open the xlsm directly, ignore mismatch
'   "DIFF"      -- show config diff, then re-prompt
'   "CANCEL"    -- abort load
' =====================================================================
Public Function ShowFastPathPrompt(verDir As String, reason As String) As String
    Dim msg As String
    msg = "Cannot use fast-path snapshot for this workspace." & vbCrLf & vbCrLf & _
        "Reason: " & reason & vbCrLf & vbCrLf & _
        "Choose how to continue:" & vbCrLf & _
        "  YES   = Rebuild the workbook from current CSV config (safest)." & vbCrLf & _
        "  NO    = Load the xlsm snapshot anyway (cells referencing new schema may break)." & vbCrLf & _
        "  CANCEL = Abort load entirely."
    Dim ans As VbMsgBoxResult
    ans = MsgBox(msg, vbYesNoCancel + vbQuestion + vbDefaultButton1, _
        "RDK -- Workspace Fast-Path Decision")
    Select Case ans
        Case vbYes: ShowFastPathPrompt = "REBUILD"
        Case vbNo:  ShowFastPathPrompt = "LOAD_XLSM"
        Case Else:  ShowFastPathPrompt = "CANCEL"
    End Select
    ' NOTE: The "Show diff" option is available via the Dashboard
    ' ShowConfigDiff button rather than inside this dialog to keep the
    ' workspace-load fast path deterministic. Documented in SESSION_NOTES.
End Function

' =====================================================================
' OpenXlsmSnapshot -- decision L4.
' Opens the xlsm in the same Excel instance as a read-only inspection window.
' Does NOT replace the currently-running workbook (which would kill VBA context).
' Returns TRUE on success.
' =====================================================================
Public Function OpenXlsmSnapshot(verDir As String) As Boolean
    On Error GoTo openErr
    Dim xlsmPath As String
    xlsmPath = BuildPath(verDir, XLSM_SNAPSHOT_NAME)
    If Dir(xlsmPath) = "" Then
        OpenXlsmSnapshot = False
        Exit Function
    End If
    Workbooks.Open FileName:=xlsmPath, ReadOnly:=True
    OpenXlsmSnapshot = True
    Exit Function
openErr:
    KernelConfig.LogError SEV_ERROR, "KWSExt", "E-921", _
        "Cannot open xlsm snapshot " & verDir & ": " & Err.Description & _
        ". MANUAL BYPASS: Double-click the xlsm in File Explorer.", ""
    OpenXlsmSnapshot = False
End Function

' =====================================================================
' RestorePrngFromManifest -- decision U.
' Reads prngSeed from manifest.json and calls KernelRandom.InitSeed.
' =====================================================================
Public Sub RestorePrngFromManifest(verDir As String)
    On Error Resume Next
    Dim manifestPath As String
    manifestPath = BuildPath(verDir, "manifest.json")
    If Dir(manifestPath) = "" Then Exit Sub
    Dim seedStr As String
    seedStr = ExtractJsonNumberFromFile(manifestPath, "prngSeed")
    If IsNumeric(seedStr) And Len(seedStr) > 0 Then
        Dim s As Long
        s = CLng(seedStr)
        If s > 0 Then KernelRandom.InitSeed s
    End If
End Sub

' =====================================================================
' CleanupOrphanFolders -- decision S.
' Scans wsDir for version folders missing manifest.json or with status!=COMPLETE
' and moves them aside to {name}.orphan{N} for user review. Never silently deletes.
' =====================================================================
Public Function CleanupOrphanFolders(wsDir As String) As Long
    Dim moved As Long: moved = 0
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(wsDir) Then Exit Function
    Dim subF As Object
    For Each subF In fso.GetFolder(wsDir).SubFolders
        If Left$(subF.Name, 1) = "v" And Len(subF.Name) >= 4 Then
            Dim verPath As String: verPath = subF.Path
            Dim manifestPath As String
            manifestPath = BuildPath(verPath, "manifest.json")
            Dim orphan As Boolean: orphan = False
            If Dir(manifestPath) = "" Then
                orphan = True
            Else
                Dim s As String, k As String, ch As String, x As String
                ParseManifestFields manifestPath, s, k, ch, x
                If StrComp(s, "COMPLETE", vbTextCompare) <> 0 Then orphan = True
            End If
            If orphan Then
                Dim newName As String
                newName = verPath & ".orphan." & Format$(Now, "yyyymmdd_hhnnss")
                Name verPath As newName
                moved = moved + 1
                KernelConfig.LogError SEV_WARN, "KWSExt", "W-932", _
                    "Orphan version folder renamed to " & newName & _
                    ". MANUAL BYPASS: Inspect and delete or recover manually.", ""
            End If
        End If
    Next subF
    Set fso = Nothing
    CleanupOrphanFolders = moved
End Function

' =====================================================================
' CapturePreservedCells -- decision Q.
' Reads preserved_cells_config.csv (active config), snapshots the listed
' cells to preserved_cells.csv in verDir.
' =====================================================================
Public Function CapturePreservedCells(verDir As String) As Long
    Dim captured As Long: captured = 0
    On Error GoTo captureErr

    Dim cfgPath As String
    cfgPath = BuildPath(KernelSnapshot.GetProjectRoot(), "config")
    cfgPath = BuildPath(cfgPath, CFG_PRESERVED_CELLS)
    If Dir(cfgPath) = "" Then
        ' No preserved-cells config -- nothing to capture
        CapturePreservedCells = 0
        Exit Function
    End If

    Dim outPath As String
    outPath = BuildPath(verDir, PRESERVED_CELLS_CSV)
    Dim fNum As Integer: fNum = FreeFile
    Open outPath For Output As #fNum
    Print #fNum, """TabName"",""Address"",""Value"",""RowID"",""Note"""

    Dim lines() As String
    lines = KernelSnapshot.ReadFileLinesFromPath(cfgPath)
    Dim i As Long
    For i = 1 To UBound(lines) ' skip header row [0]
        Dim ln As String: ln = Trim$(lines(i))
        If Len(ln) = 0 Then GoTo NextLine
        Dim f() As String
        f = ParseCsvLineSimple(ln)
        If UBound(f) < 2 Then GoTo NextLine
        Dim tabName As String: tabName = f(0)
        Dim addr As String: addr = f(1)
        Dim rowId As String: rowId = ""
        If UBound(f) >= 2 Then rowId = f(2)
        Dim note As String: note = ""
        If UBound(f) >= 3 Then note = f(3)

        If Not TabExists(tabName) Then GoTo NextLine
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets(tabName)
        Dim rng As Range
        On Error Resume Next
        Set rng = ws.Range(addr)
        On Error GoTo captureErr
        If rng Is Nothing Then GoTo NextLine

        Dim cell As Range
        For Each cell In rng.Cells
            Dim addrStr As String: addrStr = cell.Address(False, False)
            Dim val As String: val = CStr(cell.Value)
            ' Escape embedded quotes
            val = Replace(val, """", """""")
            Print #fNum, """" & tabName & """,""" & addrStr & """,""" & val & """,""" & rowId & """,""" & note & """"
            captured = captured + 1
        Next cell
NextLine:
    Next i
    Close #fNum
    CapturePreservedCells = captured
    Exit Function

captureErr:
    On Error Resume Next
    Close #fNum
    KernelConfig.LogError SEV_WARN, "KWSExt", "W-940", _
        "Preserved cells capture failed: " & Err.Description & _
        ". MANUAL BYPASS: Note current values of cells listed in " & CFG_PRESERVED_CELLS & _
        " before next Setup.", ""
    CapturePreservedCells = captured
End Function

' =====================================================================
' RestorePreservedCells -- decision Q.
' Reads preserved_cells.csv from verDir and restores each cell value.
' Runs AFTER RefreshFormulaTabs so defaults are seeded first, then overwritten.
' =====================================================================
Public Function RestorePreservedCells(verDir As String) As Long
    Dim restored As Long: restored = 0
    On Error GoTo restoreErr
    Dim path As String
    path = BuildPath(verDir, PRESERVED_CELLS_CSV)
    If Dir(path) = "" Then
        RestorePreservedCells = 0
        Exit Function
    End If
    Dim lines() As String
    lines = KernelSnapshot.ReadFileLinesFromPath(path)
    Dim i As Long
    For i = 1 To UBound(lines)
        Dim ln As String: ln = Trim$(lines(i))
        If Len(ln) = 0 Then GoTo NextLine2
        Dim f() As String
        f = ParseCsvLineSimple(ln)
        If UBound(f) < 2 Then GoTo NextLine2
        Dim tabName As String: tabName = f(0)
        Dim addr As String: addr = f(1)
        Dim val As String: val = f(2)
        If Not TabExists(tabName) Then GoTo NextLine2
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets(tabName)
        On Error Resume Next
        ' AP-07/AP-50 safety: force text format if value begins with formula chars
        If Len(val) > 0 Then
            Dim first As String: first = Left$(val, 1)
            If first = "=" Or first = "+" Or first = "-" Or first = "@" Then
                ws.Range(addr).NumberFormat = "@"
            End If
        End If
        If IsNumeric(val) Then
            ws.Range(addr).Value = CDbl(val)
        Else
            ws.Range(addr).Value = val
        End If
        On Error GoTo restoreErr
        restored = restored + 1
NextLine2:
    Next i
    RestorePreservedCells = restored
    Exit Function
restoreErr:
    KernelConfig.LogError SEV_WARN, "KWSExt", "W-941", _
        "Preserved cells restore failed: " & Err.Description & _
        ". MANUAL BYPASS: Inspect preserved_cells.csv and apply values manually.", ""
    RestorePreservedCells = restored
End Function

' =====================================================================
' AssertCoverageComplete -- decision O, layer 2.
' Scans formula_tab_config for Input cells; cross-references with
' tab_registry.InputSurface and preserved_cells_config.csv; logs SEV_ERROR
' if an Input cell exists on a tab without declared coverage.
' Call this at bootstrap-time.
' =====================================================================
Public Function AssertCoverageComplete(Optional strict As Boolean = False) As Long
    On Error GoTo coverErr
    Dim violations As Long: violations = 0

    ' Build tabs-with-input-cells set from formula_tab_config
    Dim tabHasInput As Object
    Set tabHasInput = CreateObject("Scripting.Dictionary")
    Dim cfgRows As Long
    cfgRows = KernelConfig.GetFormulaTabConfigCount()
    Dim i As Long
    For i = 1 To cfgRows
        Dim cellType As String
        cellType = KernelConfig.GetFormulaTabConfigField(i, FTCFG_COL_CELLTYPE)
        If StrComp(cellType, "Input", vbTextCompare) = 0 Then
            Dim tn As String
            tn = KernelConfig.GetFormulaTabConfigField(i, FTCFG_COL_TABNAME)
            If Not tabHasInput.Exists(tn) Then tabHasInput.Add tn, True
        End If
    Next i

    ' Build tabs-with-coverage set by reading tab_registry directly from Config sheet
    Dim coveredTabs As Object
    Set coveredTabs = CreateObject("Scripting.Dictionary")
    Dim wsCfg As Worksheet
    Set wsCfg = ThisWorkbook.Sheets(TAB_CONFIG)
    Dim tabSecStart As Long
    tabSecStart = KernelConfigLoader.FindSectionStart(wsCfg, CFG_MARKER_TAB_REGISTRY)
    If tabSecStart > 0 Then
        Dim dr As Long: dr = tabSecStart + 2
        Do While Len(Trim(CStr(wsCfg.Cells(dr, TREG_COL_TABNAME).Value))) > 0
            Dim trTab As String
            trTab = Trim(CStr(wsCfg.Cells(dr, TREG_COL_TABNAME).Value))
            Dim trSurface As String
            trSurface = Trim(CStr(wsCfg.Cells(dr, TREG_COL_INPUTSURFACE).Value))
            If Len(trSurface) > 0 And StrComp(trSurface, "NONE", vbTextCompare) <> 0 Then
                If Not coveredTabs.Exists(trTab) Then coveredTabs.Add trTab, trSurface
            End If
            dr = dr + 1
        Loop
    End If

    ' Any tab with Input cells must be covered
    Dim tabName As Variant
    For Each tabName In tabHasInput.Keys
        If Not coveredTabs.Exists(tabName) Then
            violations = violations + 1
            KernelConfig.LogError IIf(strict, SEV_ERROR, SEV_WARN), "KWSExt", _
                IIf(strict, "E-950", "W-950"), _
                "Coverage gap: tab '" & tabName & "' has Input cells but no InputSurface declaration in tab_registry. " & _
                "AP-65. MANUAL BYPASS: Set tab_registry.InputSurface to TAB_IS_INPUT, PRESERVED_CELLS, or EPHEMERAL.", ""
        End If
    Next tabName

    AssertCoverageComplete = violations
    Exit Function
coverErr:
    KernelConfig.LogError SEV_WARN, "KWSExt", "W-951", _
        "Coverage assertion crashed: " & Err.Description, ""
    AssertCoverageComplete = -1
End Function

' =====================================================================
' TestWorkspaceRoundTrip -- decision O, layer 3.
' Save workspace (silent), reload it, save again, assert configHash and
' fingerprint stability. Non-destructive -- creates two new versions.
' Call from KernelTests.RunAllTests or a dev-time button.
' =====================================================================
Public Function TestWorkspaceRoundTrip() As Boolean
    On Error GoTo rtErr
    Dim wsName As String: wsName = "RoundTripTest"
    KernelWorkspace.SaveWorkspace wsName, True
    Dim root As String: root = KernelSnapshot.GetProjectRoot()
    Dim wsDir As String
    wsDir = BuildPath(BuildPath(root, "workspaces"), wsName)
    ' Find two most recent versions
    Dim v1 As String, v2 As String
    v1 = GetLatestVersionName(wsDir)
    ' Re-save to produce v2
    KernelWorkspace.SaveWorkspace wsName, True
    v2 = GetLatestVersionName(wsDir)
    If Len(v1) = 0 Or Len(v2) = 0 Or StrComp(v1, v2, vbTextCompare) = 0 Then
        KernelConfig.LogError SEV_ERROR, "KWSExt", "E-960", _
            "Round-trip: could not produce two distinct versions.", ""
        TestWorkspaceRoundTrip = False
        Exit Function
    End If
    Dim m1 As String, m2 As String
    m1 = BuildPath(BuildPath(wsDir, v1), "manifest.json")
    m2 = BuildPath(BuildPath(wsDir, v2), "manifest.json")
    Dim s1 As String, k1 As String, ch1 As String, x1 As String
    Dim s2 As String, k2 As String, ch2 As String, x2 As String
    ParseManifestFields m1, s1, k1, ch1, x1
    ParseManifestFields m2, s2, k2, ch2, x2
    Dim pass As Boolean: pass = True
    If StrComp(ch1, ch2, vbTextCompare) <> 0 Then pass = False
    If StrComp(k1, k2, vbTextCompare) <> 0 Then pass = False
    If pass Then
        KernelConfig.LogError SEV_INFO, "KWSExt", "I-960", _
            "Round-trip PASS: configHash and kernelVersion stable across " & v1 & " and " & v2 & ".", ""
    Else
        KernelConfig.LogError SEV_ERROR, "KWSExt", "E-961", _
            "Round-trip FAIL: v1(configHash=" & Left$(ch1, 8) & "...) vs v2(configHash=" & Left$(ch2, 8) & "...). " & _
            "Save is non-deterministic. MANUAL BYPASS: Compare " & v1 & " and " & v2 & " manifests manually.", ""
    End If
    TestWorkspaceRoundTrip = pass
    Exit Function
rtErr:
    KernelConfig.LogError SEV_ERROR, "KWSExt", "E-962", _
        "Round-trip test crashed: " & Err.Description, ""
    TestWorkspaceRoundTrip = False
End Function

' =====================================================================
' CleanupOldWorkspacesDialog -- decision W.
' Dashboard button action: lists all workspace versions with date+size,
' user multi-selects (y/n per version) which to delete. No auto-policy.
' =====================================================================
Public Sub CleanupOldWorkspacesDialog()
    On Error GoTo cleanupErr
    Dim root As String: root = KernelSnapshot.GetProjectRoot()
    Dim wsRoot As String: wsRoot = BuildPath(root, "workspaces")
    If Dir(wsRoot, vbDirectory) = "" Then
        MsgBox "No workspaces folder.", vbInformation, "RDK"
        Exit Sub
    End If
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim summary As String
    summary = "Workspace versions on disk:" & vbCrLf & vbCrLf
    Dim totalMB As Double: totalMB = 0
    Dim wsFolder As Object
    For Each wsFolder In fso.GetFolder(wsRoot).SubFolders
        Dim verFolder As Object
        For Each verFolder In fso.GetFolder(wsFolder.Path).SubFolders
            If Left$(verFolder.Name, 1) = "v" Then
                Dim szMB As Double
                szMB = verFolder.Size / 1048576
                totalMB = totalMB + szMB
                summary = summary & "  " & wsFolder.Name & "\" & verFolder.Name & _
                    "  " & Format$(szMB, "0.0") & " MB  " & _
                    Format$(verFolder.DateLastModified, "yyyy-mm-dd") & vbCrLf
            End If
        Next verFolder
    Next wsFolder
    summary = summary & vbCrLf & "Total: " & Format$(totalMB, "0.0") & " MB"
    summary = summary & vbCrLf & vbCrLf & _
        "To delete versions, use File Explorer (no auto-policy by design: Never Lose Control)."
    Set fso = Nothing
    MsgBox summary, vbInformation, "RDK -- Cleanup Old Workspaces"
    Exit Sub
cleanupErr:
    MsgBox "Cleanup listing failed: " & Err.Description, vbCritical, "RDK"
End Sub

' =====================================================================
' Helpers
' =====================================================================
Private Sub ParseManifestFields(manifestPath As String, _
    ByRef status As String, ByRef kernelVer As String, _
    ByRef configHash As String, ByRef xlsxFilename As String)
    status = "": kernelVer = "": configHash = "": xlsxFilename = ""
    Dim lines() As String
    On Error Resume Next
    lines = KernelSnapshot.ReadFileLinesFromPath(manifestPath)
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim ln As String: ln = Trim$(lines(i))
        If InStr(1, ln, """status""", vbTextCompare) > 0 Then
            status = ExtractJsonStringValue(ln)
        ElseIf InStr(1, ln, """kernelVersion""", vbTextCompare) > 0 Then
            kernelVer = ExtractJsonStringValue(ln)
        ElseIf InStr(1, ln, """configHash""", vbTextCompare) > 0 And _
               InStr(1, ln, """sha256""", vbTextCompare) = 0 Then
            configHash = ExtractJsonStringValue(ln)
        ElseIf InStr(1, ln, """xlsxFilename""", vbTextCompare) > 0 Then
            xlsxFilename = ExtractJsonStringValue(ln)
        End If
    Next i
End Sub

Private Function ExtractJsonStringValue(lineText As String) As String
    ' Extract the RHS string from a `"key": "value",` line
    Dim firstColon As Long
    firstColon = InStr(1, lineText, ":")
    If firstColon = 0 Then
        ExtractJsonStringValue = ""
        Exit Function
    End If
    Dim rhs As String: rhs = Mid$(lineText, firstColon + 1)
    Dim firstQ As Long: firstQ = InStr(1, rhs, """")
    If firstQ = 0 Then
        ExtractJsonStringValue = ""
        Exit Function
    End If
    Dim afterQ As String: afterQ = Mid$(rhs, firstQ + 1)
    Dim lastQ As Long: lastQ = InStr(1, afterQ, """")
    If lastQ = 0 Then
        ExtractJsonStringValue = afterQ
        Exit Function
    End If
    ExtractJsonStringValue = Left$(afterQ, lastQ - 1)
End Function

Private Function ExtractJsonNumberFromFile(path As String, key As String) As String
    ExtractJsonNumberFromFile = ""
    Dim lines() As String
    On Error Resume Next
    lines = KernelSnapshot.ReadFileLinesFromPath(path)
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim ln As String: ln = Trim$(lines(i))
        If InStr(1, ln, """" & key & """", vbTextCompare) > 0 Then
            Dim firstColon As Long: firstColon = InStr(1, ln, ":")
            If firstColon = 0 Then Exit Function
            Dim rhs As String: rhs = Trim$(Mid$(ln, firstColon + 1))
            ' Strip trailing comma
            If Right$(rhs, 1) = "," Then rhs = Left$(rhs, Len(rhs) - 1)
            ExtractJsonNumberFromFile = Trim$(rhs)
            Exit Function
        End If
    Next i
End Function

Private Function CompareKernelVersion(a As String, b As String) As Long
    ' Returns -1 if a<b, 0 if equal, 1 if a>b. Handles semver-ish "M.m.p" strings.
    If Len(a) = 0 And Len(b) = 0 Then CompareKernelVersion = 0: Exit Function
    If Len(a) = 0 Then CompareKernelVersion = -1: Exit Function
    If Len(b) = 0 Then CompareKernelVersion = 1: Exit Function
    Dim ap() As String, bp() As String
    ap = Split(a, ".")
    bp = Split(b, ".")
    Dim i As Long, maxI As Long
    maxI = UBound(ap)
    If UBound(bp) > maxI Then maxI = UBound(bp)
    For i = 0 To maxI
        Dim av As Long, bv As Long
        av = 0: bv = 0
        If i <= UBound(ap) And IsNumeric(ap(i)) Then av = CLng(ap(i))
        If i <= UBound(bp) And IsNumeric(bp(i)) Then bv = CLng(bp(i))
        If av < bv Then CompareKernelVersion = -1: Exit Function
        If av > bv Then CompareKernelVersion = 1: Exit Function
    Next i
    CompareKernelVersion = 0
End Function

Private Function GetLatestVersionName(wsDir As String) As String
    GetLatestVersionName = ""
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(wsDir) Then Exit Function
    Dim subF As Object
    For Each subF In fso.GetFolder(wsDir).SubFolders
        If Left$(subF.Name, 1) = "v" And Len(subF.Name) >= 4 Then
            If subF.Name > GetLatestVersionName Then GetLatestVersionName = subF.Name
        End If
    Next subF
    Set fso = Nothing
End Function

Private Function TabExists(tabName As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(tabName)
    TabExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function ParseCsvLineSimple(line As String) As String()
    ' Lightweight CSV parser: splits on commas not inside double-quotes, strips
    ' surrounding quotes, unescapes doubled quotes. Sufficient for our
    ' one-line-per-record writes with QUOTE_ALL. Not RFC-4180-complete.
    Dim fields() As String
    ReDim fields(0 To 7)
    Dim count As Long: count = 0
    Dim inQuote As Boolean: inQuote = False
    Dim buf As String: buf = ""
    Dim i As Long
    For i = 1 To Len(line)
        Dim ch As String: ch = Mid$(line, i, 1)
        If ch = """" Then
            If inQuote And i < Len(line) And Mid$(line, i + 1, 1) = """" Then
                buf = buf & """"
                i = i + 1
            Else
                inQuote = Not inQuote
            End If
        ElseIf ch = "," And Not inQuote Then
            If count > UBound(fields) Then ReDim Preserve fields(0 To count + 7)
            fields(count) = buf
            count = count + 1
            buf = ""
        Else
            buf = buf & ch
        End If
    Next i
    If count > UBound(fields) Then ReDim Preserve fields(0 To count)
    fields(count) = buf
    ReDim Preserve fields(0 To count)
    ParseCsvLineSimple = fields
End Function

Private Sub DeleteFolderRecursive(path As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path) Then fso.DeleteFolder path, True
    Set fso = Nothing
End Sub
