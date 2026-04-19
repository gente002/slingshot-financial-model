Attribute VB_Name = "KernelWorkspace"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelWorkspace.bas
' Purpose: Named workspaces with linear version history.
'          Delegates export/import to KernelSnapshot public helpers.
'
' Workspace layout on disk:
'   workspaces/
'     {Name}/
'       workspace.json       -- metadata
'       v001/                -- version folder (snapshot structure)
'       v002/
'
' Public API:
'   SaveWorkspace, LoadWorkspace, LoadWorkspaceInputsOnly,
'   ListWorkspaces, BranchWorkspace, RevertWorkspace
' =============================================================================

Private Const DIR_WORKSPACES As String = "workspaces"
Private Const WS_JSON As String = "workspace.json"
Private Const WS_NAME_MAX As Long = 50
Private Const WS_DEFAULT_NAME As String = "Main"
Private Const WS_VERSION_PREFIX As String = "v"
Private Const WS_VERSION_PAD As Long = 3

' =====================================================================
' SaveWorkspace
' Saves current state as a new version in the named workspace.
' =====================================================================
Public Sub SaveWorkspace(Optional workspaceName As String = "", Optional silent As Boolean = False)
    On Error GoTo ErrHandler

    ' Gate: workspace_config WorkspacesEnabled
    Dim enabled As String
    enabled = KernelConfig.GetWorkspaceSetting("WorkspacesEnabled")
    If StrComp(enabled, "FALSE", vbTextCompare) = 0 Then
        MsgBox "Workspaces are disabled.", vbInformation, "RDK"
        Exit Sub
    End If

    Dim wsName As String
    wsName = ResolveWorkspaceName(workspaceName)
    If Len(wsName) = 0 Then Exit Sub

    Dim staleAction As Long
    staleAction = KernelFormHelpers.CheckStalenessBeforeSave()
    If staleAction = 0 Then Exit Sub
    If staleAction = 1 Then KernelEngine.RunModel

    Dim root As String
    root = KernelSnapshot.GetProjectRoot()
    Dim wsRoot As String
    wsRoot = KernelWorkspaceExt.BuildPath(root, DIR_WORKSPACES)
    KernelSnapshot.EnsureDirectoryExists wsRoot
    KernelWorkspaceExt.WarnIfCloudSyncedPath wsRoot
    Dim wsDir As String
    wsDir = KernelWorkspaceExt.BuildPath(wsRoot, wsName)
    KernelSnapshot.EnsureDirectoryExists wsDir

    ' Decision S: move any incomplete/orphan version folders aside so they
    ' don't confuse GetNextVersionNumber or subsequent load-latest.
    KernelWorkspaceExt.CleanupOrphanFolders wsDir

    ' Decision O layer 2: coverage lint. SEV_WARN mode -- logs gaps without
    ' blocking the save. Strict mode is invoked at bootstrap / pre-package.
    KernelWorkspaceExt.AssertCoverageComplete False

    Dim nextVer As Long
    nextVer = GetNextVersionNumber(wsDir)

    ' Check MaxVersionsPerWorkspace
    Dim maxVerStr As String
    maxVerStr = KernelConfig.GetWorkspaceSetting("MaxVersionsPerWorkspace")
    If IsNumeric(maxVerStr) And Len(maxVerStr) > 0 Then
        Dim maxVer As Long
        maxVer = CLng(maxVerStr)
        If maxVer > 0 And (nextVer - 1) >= maxVer Then
            KernelConfig.LogError SEV_WARN, "KWS", "W-905", _
                "Workspace " & wsName & " has " & (nextVer - 1) & " versions (max " & maxVer & ").", ""
        End If
    End If

    Dim verName As String
    verName = FmtVer(nextVer)
    Dim verDir As String
    verDir = KernelWorkspaceExt.BuildPath(wsDir, verName)
    KernelSnapshot.EnsureDirectoryExists verDir

    If IsWALEnabled() Then KernelSnapshot.WriteWAL "WS_SAVE_START", wsName & "/" & verName

    ' Export state using KernelSnapshot helpers (includes regression tabs SC-09)
    ExportStateToFolder verDir

    ' Decision Q: capture preserved cells from Output-category tabs (e.g. RBC
    ' Capital Model LOB factors / Program map / NAIC charges). Runs AFTER
    ' ExportStateToFolder so input_tabs/ has already been written.
    Dim preservedCount As Long
    preservedCount = KernelWorkspaceExt.CapturePreservedCells(verDir)
    If preservedCount > 0 Then
        KernelConfig.LogError SEV_INFO, "KWS", "I-903", _
            "Captured " & preservedCount & " preserved cells to preserved_cells.csv", ""
    End If

    ' Decision L: perfect-fidelity xlsm snapshot alongside the CSV artifacts.
    Dim xlsmPath As String
    xlsmPath = KernelWorkspaceExt.SaveXlsmSnapshot(verDir)
    If Len(xlsmPath) > 0 Then
        KernelConfig.LogError SEV_INFO, "KWS", "I-904", _
            "Xlsm snapshot saved: " & xlsmPath, ""
    End If

    ' Write/update workspace.json
    Dim wsJsonPath As String
    wsJsonPath = wsDir & "\" & WS_JSON
    Dim isNew As Boolean
    isNew = (Dir(wsJsonPath) = "")
    Dim nowStamp As String
    nowStamp = KernelSnapshot.FormatISOTimestamp()
    Dim desc As String: desc = ""
    Dim parentWs As String: parentWs = ""
    Dim parentVer As Long: parentVer = 0

    If Not isNew Then
        desc = ReadJsonField(wsJsonPath, "description")
        parentWs = ReadJsonField(wsJsonPath, "parentWorkspace")
        Dim pvStr As String
        pvStr = ReadJsonField(wsJsonPath, "parentVersion")
        If IsNumeric(pvStr) And Len(pvStr) > 0 Then parentVer = CLng(pvStr)
    End If

    WriteWsJson wsJsonPath, wsName, nextVer, desc, _
                IIf(isNew, nowStamp, ReadJsonField(wsJsonPath, "created")), _
                nowStamp, parentWs, parentVer

    If IsWALEnabled() Then KernelSnapshot.WriteWAL "WS_SAVE_DONE", wsName & "/" & verName
    KernelConfig.LogError SEV_INFO, "KWS", "I-900", _
        "Workspace saved: " & wsName & " " & verName, verDir
    If Not silent Then
        MsgBox "Workspace saved: " & wsName & " " & verName, vbInformation, "RDK"
    End If
    Exit Sub
ErrHandler:
    If IsWALEnabled() Then
        On Error Resume Next
        KernelSnapshot.WriteWAL "WS_SAVE_FAIL", wsName & " | " & Err.Description
        On Error GoTo 0
    End If
    KernelConfig.LogError SEV_ERROR, "KWS", "E-900", _
        "Save workspace error: " & Err.Description, _
        "MANUAL BYPASS: Copy files to workspaces\" & wsName & "\vNNN manually."
    MsgBox "Save workspace error: " & Err.Description, vbCritical, "RDK"
End Sub

' =====================================================================
' LoadWorkspace
' Loads a workspace version. version=0 loads latest.
' =====================================================================
Public Sub LoadWorkspace(workspaceName As String, Optional version As Long = 0)
    On Error GoTo ErrHandler

    ' Gate: lock check
    If KernelFormHelpers.CheckLockGate("LOAD_WORKSPACE") Then Exit Sub

    ' Gate: workspace_config WorkspacesEnabled
    Dim wsEnabled As String
    wsEnabled = KernelConfig.GetWorkspaceSetting("WorkspacesEnabled")
    If StrComp(wsEnabled, "FALSE", vbTextCompare) = 0 Then
        MsgBox "Workspaces are disabled.", vbInformation, "RDK"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim verDir As String
    Dim verName As String
    verDir = ResolveVersionDir(workspaceName, version, verName)
    If Len(verDir) = 0 Then GoTo Cleanup

    If IsWALEnabled() Then KernelSnapshot.WriteWAL "WS_LOAD_START", workspaceName & "/" & verName

    ' Restore config
    KernelSnapshotIO.RestoreConfigFromSnapshot verDir
    KernelBootstrap.LoadConfigFromDisk
    KernelConfig.LoadAllConfig

    ' Restore inputs
    Dim ipPath As String
    ipPath = verDir & "\inputs.csv"
    If Dir(ipPath) <> "" Then
        KernelSnapshotIO.ImportInputsFromCsv ipPath
    Else
        KernelConfig.LogError SEV_WARN, "KWS", "W-920", _
            "inputs.csv missing in " & verName & "; skipped", ""
    End If

    ' Restore domain input tabs (UW Inputs, Assumptions, etc.)
    KernelTabIO.ImportAllInputTabs verDir

    ' Restore detail
    Dim dtPath As String
    dtPath = verDir & "\detail.csv"
    If Dir(dtPath) <> "" Then
        KernelSnapshotIO.ImportDetailFromCsv dtPath
    Else
        KernelConfig.LogError SEV_WARN, "KWS", "W-921", _
            "detail.csv missing in " & verName & "; skipped", ""
    End If

    ' Restore error log
    Dim elPath As String
    elPath = verDir & "\errorlog.csv"
    If Dir(elPath) <> "" Then
        KernelSnapshotIO.ImportErrorLogFromCsv elPath
    Else
        KernelConfig.LogError SEV_WARN, "KWS", "W-922", _
            "errorlog.csv missing in " & verName & "; skipped", ""
    End If

    ' Restore granular CSV if present
    Dim granSrc As String
    granSrc = verDir & "\granular_detail.csv"
    If Dir(granSrc) <> "" Then
        Dim granDest As String
        granDest = KernelFormHelpers.EnsureOutputDir() & "\granular_detail_restored.csv"
        On Error Resume Next
        FileCopy granSrc, granDest
        On Error GoTo ErrHandler
        KernelConfig.LogError SEV_INFO, "KWS", "I-925", _
            "Restored granular CSV: " & granDest, ""
    End If

    ' Regenerate Quarterly Summary SUMIFS and refresh formula tabs
    ' if Detail was restored (QS formulas reference Detail rows)
    If Dir(verDir & "\detail.csv") <> "" Then
        On Error Resume Next
        ' Re-run transforms (QuarterlyAgg writes SUMIFS, Triangles rebuild)
        Dim domMod As String
        domMod = KernelConfig.GetSetting("DomainModule")
        If Len(domMod) > 0 Then
            Application.Run domMod & ".Initialize"
        End If
        KernelExtension.LoadExtensionRegistry
        Dim loadOutputs() As Variant
        ReDim loadOutputs(0)
        KernelTransform.RunTransforms loadOutputs
        ' Refresh formula tabs so UWEX/IS/BS pick up restored data
        KernelFormulaWriter.RefreshFormulaTabs
        On Error GoTo ErrHandler
    End If

    ' Decision Q: restore preserved cells AFTER RefreshFormulaTabs so default
    ' values are seeded first, then overwritten with user edits from the save.
    On Error Resume Next
    Dim restoredCount As Long
    restoredCount = KernelWorkspaceExt.RestorePreservedCells(verDir)
    If restoredCount > 0 Then
        KernelConfig.LogError SEV_INFO, "KWS", "I-914", _
            "Restored " & restoredCount & " preserved cells from preserved_cells.csv", ""
    End If
    ' Decision U: restore PRNG seed from manifest so post-load model runs are
    ' deterministic against the saved state.
    KernelWorkspaceExt.RestorePrngFromManifest verDir
    On Error GoTo ErrHandler

    ' Recalculate
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    If IsWALEnabled() Then KernelSnapshot.WriteWAL "WS_LOAD_DONE", workspaceName & "/" & verName
    KernelConfig.LogError SEV_INFO, "KWS", "I-910", _
        "Loaded: " & workspaceName & " " & verName, verDir

    ' Update scenario name on Assumptions tab
    On Error Resume Next
    Dim wsAssume As Worksheet
    Set wsAssume = ThisWorkbook.Sheets(KernelConfig.GetSetting("InputsTabName"))
    If Not wsAssume Is Nothing Then wsAssume.Cells(4, 3).Value = workspaceName
    On Error GoTo ErrHandler

    ' Restore lock visual if model is locked
    KernelFormHelpers.RestoreLockVisual

    ' Check if detail was present -- if not, offer to Run & Save
    Dim hadDetail As Boolean
    hadDetail = (Dir(verDir & "\detail.csv") <> "")
    If Not hadDetail Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("Loaded: " & workspaceName & " " & verName & " (inputs only)." & vbCrLf & vbCrLf & _
            "No saved results found. Run Model now to compute" & vbCrLf & _
            "and save results to this workspace?", _
            vbYesNo Or vbQuestion, "RDK -- Workspace Loaded")
        If ans = vbYes Then
            KernelEngine.RunModel
            KernelSnapshotIO.ExportDetailToFile verDir & "\detail.csv"
            MsgBox "Results computed and saved.", vbInformation, "RDK"
        End If
    Else
        MsgBox "Loaded: " & workspaceName & " " & verName, vbInformation, "RDK"
    End If
    Exit Sub
Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    KernelConfig.LogError SEV_ERROR, "KWS", "E-919", _
        "Load workspace error: " & Err.Description, _
        "MANUAL BYPASS: Copy CSV files from version folder to respective tabs."
    MsgBox "Load workspace error: " & Err.Description, vbCritical, "RDK"
End Sub

' =====================================================================
' LoadWorkspaceInputsOnly
' Loads only inputs (and config) from a workspace version.
' =====================================================================
Public Sub LoadWorkspaceInputsOnly(workspaceName As String, Optional version As Long = 0)
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Dim verDir As String
    Dim verName As String
    verDir = ResolveVersionDir(workspaceName, version, verName)
    If Len(verDir) = 0 Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Restore config + inputs only
    KernelSnapshotIO.RestoreConfigFromSnapshot verDir
    KernelBootstrap.LoadConfigFromDisk
    KernelConfig.LoadAllConfig

    Dim ipPath As String
    ipPath = verDir & "\inputs.csv"
    If Dir(ipPath) <> "" Then
        KernelSnapshotIO.ImportInputsFromCsv ipPath
    Else
        KernelConfig.LogError SEV_WARN, "KWS", "W-930", _
            "inputs.csv not found in " & verName, _
            "MANUAL BYPASS: Place inputs.csv in version folder."
        Application.ScreenUpdating = True
        MsgBox "inputs.csv not found in " & verName, vbExclamation, "RDK"
        Exit Sub
    End If

    ' Restore domain input tabs (UW Inputs, Assumptions, etc.)
    KernelTabIO.ImportAllInputTabs verDir

    KernelFormHelpers.WriteRunStateValue RS_KEY_STALE, "TRUE"
    Application.ScreenUpdating = True

    If IsWALEnabled() Then KernelSnapshot.WriteWAL "WS_LOAD_INPUTS", workspaceName & "/" & verName
    KernelConfig.LogError SEV_INFO, "KWS", "I-930", _
        "Inputs loaded (stale): " & workspaceName & " " & verName, ""
    MsgBox "Inputs loaded: " & workspaceName & " " & verName & vbCrLf & _
           "Results now stale. Run model.", vbInformation, "RDK"
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    KernelConfig.LogError SEV_ERROR, "KWS", "E-939", _
        "Load inputs error: " & Err.Description, _
        "MANUAL BYPASS: Copy inputs.csv from version folder to Inputs tab."
    MsgBox "Load inputs error: " & Err.Description, vbCritical, "RDK"
End Sub

' =====================================================================
' ListWorkspaces
' Shows a MsgBox listing all workspaces with version count and last saved.
' =====================================================================
Public Sub ListWorkspaces()
    On Error GoTo ErrHandler
    Dim root As String
    root = KernelSnapshot.GetProjectRoot()
    Dim wsRoot As String
    wsRoot = root & "\" & DIR_WORKSPACES

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(wsRoot) Then
        Set fso = Nothing
        MsgBox "No workspaces found." & vbCrLf & vbCrLf & _
               "Use SaveWorkspace to create one.", vbInformation, "RDK"
        Exit Sub
    End If

    Dim parentFolder As Object
    Set parentFolder = fso.GetFolder(wsRoot)
    Dim cnt As Long: cnt = 0
    Dim subFolder As Object
    For Each subFolder In parentFolder.SubFolders
        If fso.FileExists(subFolder.Path & "\" & WS_JSON) Then cnt = cnt + 1
    Next subFolder

    If cnt = 0 Then
        Set fso = Nothing
        MsgBox "No workspaces found." & vbCrLf & _
               "Use SaveWorkspace to create one.", vbInformation, "RDK"
        Exit Sub
    End If

    Dim listing As String
    listing = "Workspaces (" & cnt & "):" & vbCrLf & String(40, "-") & vbCrLf

    For Each subFolder In parentFolder.SubFolders
        Dim jsonPath As String
        jsonPath = subFolder.Path & "\" & WS_JSON
        If fso.FileExists(jsonPath) Then
            Dim wsN As String: wsN = ReadJsonField(jsonPath, "name")
            If Len(wsN) = 0 Then wsN = subFolder.Name
            Dim curV As String: curV = ReadJsonField(jsonPath, "currentVersion")
            Dim lastS As String: lastS = ReadJsonField(jsonPath, "lastSaved")
            Dim dsc As String: dsc = ReadJsonField(jsonPath, "description")
            Dim pw As String: pw = ReadJsonField(jsonPath, "parentWorkspace")

            listing = listing & wsN & "  [" & curV & " versions]"
            If Len(pw) > 0 Then listing = listing & " (from " & pw & ")"
            listing = listing & vbCrLf
            If Len(lastS) > 0 Then listing = listing & "  Last saved: " & lastS & vbCrLf
            If Len(dsc) > 0 Then listing = listing & "  " & dsc & vbCrLf
            listing = listing & vbCrLf
        End If
    Next subFolder
    Set fso = Nothing
    MsgBox listing, vbInformation, "RDK -- Workspaces"
    Exit Sub
ErrHandler:
    KernelConfig.LogError SEV_ERROR, "KWS", "E-940", _
        "List workspaces error: " & Err.Description, _
        "MANUAL BYPASS: Browse workspaces\ folder manually."
    MsgBox "List error: " & Err.Description, vbCritical, "RDK"
End Sub

' =====================================================================
' BranchWorkspace
' Creates a new workspace by copying a version from an existing workspace.
' =====================================================================
Public Sub BranchWorkspace(newName As String, _
                           Optional fromWorkspace As String = "", _
                           Optional fromVersion As Long = 0)
    On Error GoTo ErrHandler
    Dim safeName As String
    safeName = SanitizeName(newName)
    If Len(safeName) = 0 Then
        KernelConfig.LogError SEV_ERROR, "KWS", "E-950", _
            "Invalid workspace name: " & newName, _
            "MANUAL BYPASS: Use alphanumeric, underscore, or hyphen only."
        MsgBox "Invalid workspace name.", vbExclamation, "RDK"
        Exit Sub
    End If

    Dim root As String
    root = KernelSnapshot.GetProjectRoot()
    Dim wsRoot As String
    wsRoot = root & "\" & DIR_WORKSPACES

    Dim newWsDir As String
    newWsDir = wsRoot & "\" & safeName
    If Dir(newWsDir, vbDirectory) <> "" Then
        KernelConfig.LogError SEV_ERROR, "KWS", "E-951", _
            "Workspace already exists: " & safeName, _
            "MANUAL BYPASS: Choose a different name or delete existing folder."
        MsgBox "Workspace already exists: " & safeName, vbExclamation, "RDK"
        Exit Sub
    End If

    ' Resolve source
    Dim srcName As String
    If Len(fromWorkspace) = 0 Then srcName = WS_DEFAULT_NAME Else srcName = SanitizeName(fromWorkspace)
    Dim srcDir As String
    srcDir = wsRoot & "\" & srcName
    If Dir(srcDir, vbDirectory) = "" Then
        KernelConfig.LogError SEV_ERROR, "KWS", "E-952", _
            "Source workspace not found: " & srcName, _
            "MANUAL BYPASS: Verify workspaces\" & srcName & " exists."
        MsgBox "Source workspace not found: " & srcName, vbExclamation, "RDK"
        Exit Sub
    End If

    Dim srcVer As Long
    If fromVersion = 0 Then srcVer = GetCurrentVersion(srcDir) Else srcVer = fromVersion
    If srcVer = 0 Then
        KernelConfig.LogError SEV_ERROR, "KWS", "E-953", _
            "No versions in source workspace: " & srcName, _
            "MANUAL BYPASS: Save a version in " & srcName & " first."
        MsgBox "No versions in source workspace.", vbExclamation, "RDK"
        Exit Sub
    End If

    Dim srcVerName As String
    srcVerName = FmtVer(srcVer)
    Dim srcVerDir As String
    srcVerDir = srcDir & "\" & srcVerName
    If Dir(srcVerDir, vbDirectory) = "" Then
        KernelConfig.LogError SEV_ERROR, "KWS", "E-954", _
            "Source version not found: " & srcName & "/" & srcVerName, _
            "MANUAL BYPASS: Verify version folder exists."
        MsgBox "Source version not found.", vbExclamation, "RDK"
        Exit Sub
    End If

    If IsWALEnabled() Then _
        KernelSnapshot.WriteWAL "WS_BRANCH_START", srcName & "/" & srcVerName & " -> " & safeName

    ' Create new workspace
    KernelSnapshot.EnsureDirectoryExists wsRoot
    KernelSnapshot.EnsureDirectoryExists newWsDir
    Dim newVerDir As String
    newVerDir = newWsDir & "\v001"
    KernelSnapshot.EnsureDirectoryExists newVerDir
    CopyFolderContents srcVerDir, newVerDir

    ' Write workspace.json
    Dim nowStamp As String
    nowStamp = KernelSnapshot.FormatISOTimestamp()
    WriteWsJson newWsDir & "\" & WS_JSON, safeName, 1, _
                "Branched from " & srcName & " " & srcVerName, _
                nowStamp, nowStamp, srcName, srcVer

    If IsWALEnabled() Then _
        KernelSnapshot.WriteWAL "WS_BRANCH_DONE", srcName & "/" & srcVerName & " -> " & safeName
    KernelConfig.LogError SEV_INFO, "KWS", "I-950", _
        "Branched: " & srcName & "/" & srcVerName & " -> " & safeName, newWsDir
    MsgBox "Branched: " & srcName & " " & srcVerName & " -> " & safeName & " v001", _
           vbInformation, "RDK"
    Exit Sub
ErrHandler:
    KernelConfig.LogError SEV_ERROR, "KWS", "E-959", _
        "Branch error: " & Err.Description, _
        "MANUAL BYPASS: Copy version folder manually, create workspace.json."
    MsgBox "Branch error: " & Err.Description, vbCritical, "RDK"
End Sub

' =====================================================================
' RevertWorkspace
' Loads an older version and saves it as a new version (preserves history).
' =====================================================================
Public Sub RevertWorkspace(workspaceName As String, toVersion As Long)
    On Error GoTo ErrHandler
    If toVersion < 1 Then
        KernelConfig.LogError SEV_ERROR, "KWS", "E-960", _
            "Invalid version number: " & toVersion, _
            "MANUAL BYPASS: Specify a positive version number."
        MsgBox "Invalid version number.", vbExclamation, "RDK"
        Exit Sub
    End If

    Dim root As String
    root = KernelSnapshot.GetProjectRoot()
    Dim wsDir As String
    wsDir = root & "\" & DIR_WORKSPACES & "\" & workspaceName
    If Dir(wsDir, vbDirectory) = "" Then
        KernelConfig.LogError SEV_ERROR, "KWS", "E-961", _
            "Workspace not found: " & workspaceName, _
            "MANUAL BYPASS: Verify workspaces\" & workspaceName & " exists."
        MsgBox "Workspace not found: " & workspaceName, vbExclamation, "RDK"
        Exit Sub
    End If

    Dim revertVerName As String
    revertVerName = FmtVer(toVersion)
    Dim revertVerDir As String
    revertVerDir = wsDir & "\" & revertVerName
    If Dir(revertVerDir, vbDirectory) = "" Then
        KernelConfig.LogError SEV_ERROR, "KWS", "E-962", _
            "Version not found: " & workspaceName & "/" & revertVerName, _
            "MANUAL BYPASS: Verify version folder exists."
        MsgBox "Version not found: " & workspaceName & "/" & revertVerName, vbExclamation, "RDK"
        Exit Sub
    End If

    Dim curVer As Long
    curVer = GetCurrentVersion(wsDir)
    If toVersion >= curVer Then
        Dim ans As Long
        ans = MsgBox("Version " & toVersion & " is current or newer (" & curVer & ")." & vbCrLf & _
                     "Continue anyway?", vbYesNo Or vbQuestion, "RDK")
        If ans = vbNo Then Exit Sub
    End If

    If IsWALEnabled() Then _
        KernelSnapshot.WriteWAL "WS_REVERT_START", workspaceName & " " & revertVerName

    Dim nextVer As Long
    nextVer = GetNextVersionNumber(wsDir)
    Dim newVerName As String
    newVerName = FmtVer(nextVer)
    Dim newVerDir As String
    newVerDir = wsDir & "\" & newVerName
    KernelSnapshot.EnsureDirectoryExists newVerDir
    CopyFolderContents revertVerDir, newVerDir

    ' Update workspace.json
    Dim wsJsonPath As String
    wsJsonPath = wsDir & "\" & WS_JSON
    Dim nowStamp As String
    nowStamp = KernelSnapshot.FormatISOTimestamp()
    WriteWsJson wsJsonPath, workspaceName, nextVer, _
                ReadJsonField(wsJsonPath, "description"), _
                ReadJsonField(wsJsonPath, "created"), _
                nowStamp, _
                ReadJsonField(wsJsonPath, "parentWorkspace"), _
                SafeCLng(ReadJsonField(wsJsonPath, "parentVersion"))

    If IsWALEnabled() Then _
        KernelSnapshot.WriteWAL "WS_REVERT_DONE", workspaceName & " " & revertVerName & " -> " & newVerName
    KernelConfig.LogError SEV_INFO, "KWS", "I-960", _
        "Reverted: " & workspaceName & " " & revertVerName & " -> " & newVerName, ""
    MsgBox "Reverted: " & workspaceName & vbCrLf & _
           revertVerName & " copied as " & newVerName & vbCrLf & vbCrLf & _
           "Use LoadWorkspace to load the reverted version.", vbInformation, "RDK"
    Exit Sub
ErrHandler:
    KernelConfig.LogError SEV_ERROR, "KWS", "E-969", _
        "Revert error: " & Err.Description, _
        "MANUAL BYPASS: Copy version folder manually, update workspace.json."
    MsgBox "Revert error: " & Err.Description, vbCritical, "RDK"
End Sub


' #############################################################################
' PRIVATE HELPERS
' #############################################################################

Private Function ResolveWorkspaceName(rawName As String) As String
    Dim wsName As String
    If Len(Trim(rawName)) > 0 Then
        wsName = SanitizeName(rawName)
    Else
        ' Read default name from workspace_config, fall back to constant
        Dim cfgDefault As String
        cfgDefault = KernelConfig.GetWorkspaceSetting("DefaultWorkspaceName")
        If Len(cfgDefault) > 0 Then
            wsName = SanitizeName(cfgDefault)
        Else
            wsName = WS_DEFAULT_NAME
        End If
    End If
    If Len(wsName) = 0 Then
        KernelConfig.LogError SEV_ERROR, "KWS", "E-901", _
            "Invalid workspace name: " & rawName, _
            "MANUAL BYPASS: Use alphanumeric, underscore, or hyphen only."
        MsgBox "Invalid workspace name.", vbExclamation, "RDK"
    End If
    ResolveWorkspaceName = wsName
End Function

' ResolveVersionDir -- validates workspace + version, returns folder path.
' Sets verName ByRef. Returns empty string on failure (with MsgBox).
Private Function ResolveVersionDir(workspaceName As String, version As Long, _
                                   ByRef verName As String) As String
    ResolveVersionDir = ""
    verName = ""
    Dim root As String
    root = KernelSnapshot.GetProjectRoot()
    Dim wsDir As String
    wsDir = root & "\" & DIR_WORKSPACES & "\" & workspaceName
    If Dir(wsDir, vbDirectory) = "" Then
        KernelConfig.LogError SEV_ERROR, "KWS", "E-910", _
            "Workspace not found: " & workspaceName, _
            "MANUAL BYPASS: Verify workspaces\" & workspaceName & " exists."
        MsgBox "Workspace not found: " & workspaceName, vbExclamation, "RDK"
        Exit Function
    End If

    Dim targetVer As Long
    If version = 0 Then targetVer = GetCurrentVersion(wsDir) Else targetVer = version
    If targetVer = 0 Then
        KernelConfig.LogError SEV_ERROR, "KWS", "E-911", _
            "No versions in workspace: " & workspaceName, _
            "MANUAL BYPASS: Save a version first with SaveWorkspace."
        MsgBox "No versions in workspace: " & workspaceName, vbExclamation, "RDK"
        Exit Function
    End If

    verName = FmtVer(targetVer)
    Dim verDir As String
    verDir = wsDir & "\" & verName
    If Dir(verDir, vbDirectory) = "" Then
        KernelConfig.LogError SEV_ERROR, "KWS", "E-912", _
            "Version not found: " & workspaceName & "/" & verName, _
            "MANUAL BYPASS: Verify version folder exists."
        MsgBox "Version not found: " & workspaceName & "/" & verName, vbExclamation, "RDK"
        Exit Function
    End If
    ResolveVersionDir = verDir
End Function

Private Function SanitizeName(rawName As String) As String
    Dim result As String
    result = Trim(rawName)
    result = Replace(result, " ", "_")
    Dim cleaned As String: cleaned = ""
    Dim i As Long
    For i = 1 To Len(result)
        Dim ch As String
        ch = Mid(result, i, 1)
        If ch Like "[A-Za-z0-9_-]" Then cleaned = cleaned & ch
    Next i
    If Len(cleaned) > WS_NAME_MAX Then cleaned = Left(cleaned, WS_NAME_MAX)
    SanitizeName = cleaned
End Function

Private Function GetCurrentVersion(wsDir As String) As Long
    GetCurrentVersion = 0
    Dim jsonPath As String
    jsonPath = wsDir & "\" & WS_JSON
    If Dir(jsonPath) = "" Then
        GetCurrentVersion = ScanMaxVersion(wsDir)
        Exit Function
    End If
    Dim verStr As String
    verStr = ReadJsonField(jsonPath, "currentVersion")
    If IsNumeric(verStr) And Len(verStr) > 0 Then
        GetCurrentVersion = CLng(verStr)
    Else
        GetCurrentVersion = ScanMaxVersion(wsDir)
    End If
End Function

Private Function GetNextVersionNumber(wsDir As String) As Long
    GetNextVersionNumber = ScanMaxVersion(wsDir) + 1
End Function

Private Function ScanMaxVersion(wsDir As String) As Long
    ScanMaxVersion = 0
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(wsDir) Then
        Set fso = Nothing
        Exit Function
    End If
    Dim subFolder As Object
    For Each subFolder In fso.GetFolder(wsDir).SubFolders
        Dim fn As String: fn = subFolder.Name
        If Len(fn) > 1 And Left(fn, 1) = "v" Then
            Dim numPart As String: numPart = Mid(fn, 2)
            If IsNumeric(numPart) Then
                Dim verNum As Long: verNum = CLng(numPart)
                If verNum > ScanMaxVersion Then ScanMaxVersion = verNum
            End If
        End If
    Next subFolder
    Set fso = Nothing
End Function

Private Function FmtVer(verNum As Long) As String
    FmtVer = WS_VERSION_PREFIX & Right(String(WS_VERSION_PAD, "0") & CStr(verNum), WS_VERSION_PAD)
End Function

Private Function IsWALEnabled() As Boolean
    Dim val As String
    val = KernelConfig.GetReproSetting("WALEnabled")
    IsWALEnabled = (StrComp(val, "TRUE", vbTextCompare) = 0)
End Function

Private Function SafeCLng(s As String) As Long
    If IsNumeric(s) And Len(s) > 0 Then SafeCLng = CLng(s) Else SafeCLng = 0
End Function

' ExportStateToFolder -- exports current workbook state to a version folder.
' Delegates to KernelSnapshot public helpers for each file.
' Detail export is optional (may not exist if model hasn't been run).
Private Sub ExportStateToFolder(verDir As String)
    On Error GoTo ExportErr
    ' Detail is optional -- skip if empty (model not yet run)
    On Error Resume Next
    KernelSnapshotIO.ExportDetailToFile verDir & "\detail.csv"
    On Error GoTo ExportErr
    KernelSnapshotIO.ExportInputsToFile verDir & "\inputs.csv"
    KernelSnapshotIO.ExportErrorLogToFile verDir & "\errorlog.csv"
    KernelSnapshotIO.ExportSettingsToFile verDir & "\settings.csv"
    ' Sync assumptions edits from Config sheet back to CSV before saving
    KernelAssumptions.SyncAssumptionsToCSV
    KernelSnapshotIO.SaveConfigToSnapshot verDir
    ' Export all input tabs (UW Inputs, Capital Activity, etc.)
    KernelTabIO.ExportAllInputTabs verDir
    ' SC-09: Capture regression tabs (formula tab outputs) in workspace version
    KernelTabIO.ExportRegressionTabs verDir

    ' SHA256 hashing is optional -- each hash spawns a PowerShell subprocess
    ' (~150-700ms each, 6 hashes = 1-4 seconds). Skip for fast saves.
    Dim hashEnabled As Boolean
    hashEnabled = (StrComp(KernelConfig.GetReproSetting("DeterministicMode"), "TRUE", vbTextCompare) = 0)

    Dim cfgHash As String: cfgHash = ""
    Dim hD As String: hD = ""
    Dim hI As String: hI = ""
    Dim hE As String: hE = ""
    Dim hS As String: hS = ""
    Dim hC As String: hC = ""
    Dim hG As String: hG = ""

    If hashEnabled Then
        cfgHash = KernelSnapshot.BuildConfigHash()
        Dim chp As String: chp = verDir & "\config_hash.txt"
        WriteTextFile chp, cfgHash
        hD = KernelSnapshot.ComputeSHA256(verDir & "\detail.csv")
        hI = KernelSnapshot.ComputeSHA256(verDir & "\inputs.csv")
        hE = KernelSnapshot.ComputeSHA256(verDir & "\errorlog.csv")
        hS = KernelSnapshot.ComputeSHA256(verDir & "\settings.csv")
        hC = KernelSnapshot.ComputeSHA256(chp)
    End If

    ' Granular CSV (domain-specific)
    Dim granPath As String: granPath = ""
    Dim domMod As String
    domMod = KernelConfig.GetSetting("DomainModule")
    On Error Resume Next
    If Len(domMod) > 0 Then granPath = Application.Run(domMod & ".GranularCSVPath")
    If Err.Number <> 0 Then granPath = "": Err.Clear
    On Error GoTo ExportErr
    If Len(granPath) > 0 Then
        If Dir(granPath) <> "" Then
            Dim granDest As String: granDest = verDir & "\granular_detail.csv"
            FileCopy granPath, granDest
            If hashEnabled Then hG = KernelSnapshot.ComputeSHA256(granDest)
        End If
    End If

    ' Manifest
    Dim fp As String
    fp = KernelSnapshot.BuildExecutionFingerprint()
    Dim manifestPath As String
    manifestPath = verDir & "\manifest.json"
    WriteVersionManifest manifestPath, cfgHash, fp, hD, hI, hE, hS, hC, hG

    ' Patch staleness
    Dim isStale As Boolean: isStale = KernelFormHelpers.IsResultsStale()
    Dim elapsed As String: elapsed = KernelFormHelpers.ReadRunStateValue(RS_KEY_TOTAL_ELAPSED)
    KernelFormHelpers.PatchManifestStaleAndElapsed manifestPath, isStale, elapsed
    Exit Sub
ExportErr:
    Err.Raise vbObjectError + 910, "KWS", _
        "Export state failed: " & Err.Description & ". " & _
        "MANUAL BYPASS: Copy data files to " & verDir & " manually."
End Sub

' WriteVersionManifest -- writes manifest.json for a workspace version.
Private Sub WriteVersionManifest(jsonPath As String, configHash As String, _
                                 fingerprint As String, hD As String, _
                                 hI As String, hE As String, hS As String, _
                                 hC As String, Optional hG As String = "")
    Dim prngSeed As Long
    If KernelRandom.IsInitialized() Then prngSeed = KernelRandom.GetSeed() Else prngSeed = 0
    Dim Q As String: Q = """"
    Dim mColCount As Long: mColCount = KernelConfig.GetColumnCount()

    Dim fileNum As Integer: fileNum = FreeFile
    Open jsonPath For Output As #fileNum
    Print #fileNum, "{"
    Print #fileNum, "  " & Q & "status" & Q & ": " & Q & SP_STATUS_COMPLETE & Q & ","
    Print #fileNum, "  " & Q & "created" & Q & ": " & Q & KernelSnapshot.FormatISOTimestamp() & Q & ","
    Print #fileNum, "  " & Q & "kernelVersion" & Q & ": " & Q & KERNEL_VERSION & Q & ","
    Print #fileNum, "  " & Q & "configHash" & Q & ": " & Q & configHash & Q & ","
    Print #fileNum, "  " & Q & "prngSeed" & Q & ": " & prngSeed & ","
    Print #fileNum, "  " & Q & "columnCount" & Q & ": " & mColCount & ","
    Dim hdrJson As String: hdrJson = "  " & Q & "columnHeaders" & Q & ": ["
    Dim ci As Long
    For ci = 1 To mColCount
        If ci > 1 Then hdrJson = hdrJson & ","
        hdrJson = hdrJson & Q & KernelConfig.GetColName(ci) & Q
    Next ci
    Print #fileNum, hdrJson & "],"
    Print #fileNum, "  " & Q & "entityCount" & Q & ": " & KernelSnapshotIO.DetectEntityCount() & ","
    Print #fileNum, "  " & Q & "timeHorizon" & Q & ": " & KernelConfig.GetTimeHorizon() & ","
    Print #fileNum, "  " & Q & "fingerprint" & Q & ": {"
    Print #fileNum, fingerprint
    Print #fileNum, "  },"
    Print #fileNum, "  " & Q & "files" & Q & ": ["
    Print #fileNum, "    {" & Q & "name" & Q & ":" & Q & "detail.csv" & Q & "," & Q & "sha256" & Q & ":" & Q & hD & Q & "},"
    Print #fileNum, "    {" & Q & "name" & Q & ":" & Q & "inputs.csv" & Q & "," & Q & "sha256" & Q & ":" & Q & hI & Q & "},"
    Print #fileNum, "    {" & Q & "name" & Q & ":" & Q & "errorlog.csv" & Q & "," & Q & "sha256" & Q & ":" & Q & hE & Q & "},"
    Print #fileNum, "    {" & Q & "name" & Q & ":" & Q & "settings.csv" & Q & "," & Q & "sha256" & Q & ":" & Q & hS & Q & "},"
    If Len(hG) > 0 Then
        Print #fileNum, "    {" & Q & "name" & Q & ":" & Q & "config_hash.txt" & Q & "," & Q & "sha256" & Q & ":" & Q & hC & Q & "},"
        Print #fileNum, "    {" & Q & "name" & Q & ":" & Q & "granular_detail.csv" & Q & "," & Q & "sha256" & Q & ":" & Q & hG & Q & "}"
    Else
        Print #fileNum, "    {" & Q & "name" & Q & ":" & Q & "config_hash.txt" & Q & "," & Q & "sha256" & Q & ":" & Q & hC & Q & "}"
    End If
    Print #fileNum, "  ]"
    Print #fileNum, "}"
    Close #fileNum
End Sub

' WriteWsJson -- writes workspace.json (atomic: temp then rename).
Private Sub WriteWsJson(jsonPath As String, wsName As String, _
                        currentVersion As Long, description As String, _
                        created As String, lastSaved As String, _
                        parentWorkspace As String, parentVersion As Long)
    On Error GoTo JsonErr
    Dim Q As String: Q = """"
    Dim tempPath As String: tempPath = jsonPath & ".tmp"
    Dim fileNum As Integer: fileNum = FreeFile
    Open tempPath For Output As #fileNum
    Print #fileNum, "{"
    Print #fileNum, "  " & Q & "name" & Q & ": " & Q & EscJson(wsName) & Q & ","
    Print #fileNum, "  " & Q & "currentVersion" & Q & ": " & currentVersion & ","
    Print #fileNum, "  " & Q & "description" & Q & ": " & Q & EscJson(description) & Q & ","
    Print #fileNum, "  " & Q & "created" & Q & ": " & Q & EscJson(created) & Q & ","
    Print #fileNum, "  " & Q & "lastSaved" & Q & ": " & Q & EscJson(lastSaved) & Q & ","
    Print #fileNum, "  " & Q & "parentWorkspace" & Q & ": " & Q & EscJson(parentWorkspace) & Q & ","
    Print #fileNum, "  " & Q & "parentVersion" & Q & ": " & parentVersion & ","
    Print #fileNum, "  " & Q & "kernelVersion" & Q & ": " & Q & KERNEL_VERSION & Q & ","
    Dim wsHash As String: wsHash = ""
    On Error Resume Next: wsHash = KernelSnapshot.BuildConfigHash(): On Error GoTo JsonErr
    Print #fileNum, "  " & Q & "configHash" & Q & ": " & Q & wsHash & Q & ","
    Print #fileNum, "  " & Q & "pinned" & Q & ": false"
    Print #fileNum, "}"
    Close #fileNum
    If Dir(jsonPath) <> "" Then Kill jsonPath
    Name tempPath As jsonPath
    Exit Sub
JsonErr:
    On Error Resume Next
    Close #fileNum
    If Dir(tempPath) <> "" Then Kill tempPath
    On Error GoTo 0
    Err.Raise vbObjectError + 903, "KWS", _
        "Failed to write workspace.json: " & Err.Description & ". " & _
        "MANUAL BYPASS: Create workspace.json manually."
End Sub

Private Function EscJson(s As String) As String
    EscJson = Replace(Replace(s, "\", "\\"), """", "\""")
End Function

' CopyFolderContents -- recursively copies files and subdirs.
Private Sub CopyFolderContents(srcDir As String, destDir As String)
    On Error GoTo CopyErr
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(srcDir) Then
        Set fso = Nothing
        Err.Raise vbObjectError + 920, "KWS", "Source folder not found: " & srcDir
    End If
    Dim srcFile As Object
    For Each srcFile In fso.GetFolder(srcDir).Files
        fso.CopyFile srcFile.Path, destDir & "\" & srcFile.Name, True
    Next srcFile
    Dim subFolder As Object
    For Each subFolder In fso.GetFolder(srcDir).SubFolders
        Dim destSub As String: destSub = destDir & "\" & subFolder.Name
        KernelSnapshot.EnsureDirectoryExists destSub
        CopyFolderContents subFolder.Path, destSub
    Next subFolder
    Set fso = Nothing
    Exit Sub
CopyErr:
    Set fso = Nothing
    Err.Raise vbObjectError + 921, "KWS", _
        "Copy failed: " & srcDir & " -> " & destDir & ". " & _
        "MANUAL BYPASS: Copy files manually."
End Sub

' ReadJsonField -- reads a simple field from a JSON file.
Private Function ReadJsonField(jsonPath As String, fieldName As String) As String
    ReadJsonField = ""
    On Error GoTo ReadErr
    If Dir(jsonPath) = "" Then Exit Function
    Dim fileNum As Integer: fileNum = FreeFile
    Dim fileSize As Long
    Open jsonPath For Binary Access Read As #fileNum
    fileSize = LOF(fileNum)
    If fileSize = 0 Then Close #fileNum: Exit Function
    Dim content As String: content = Space$(fileSize)
    Get #fileNum, , content
    Close #fileNum
    content = Replace(Replace(content, vbCrLf, vbLf), vbCr, vbLf)
    If Right(content, 1) = vbLf Then content = Left(content, Len(content) - 1)
    Dim lines() As String: lines = Split(content, vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim lineText As String: lineText = Trim(lines(i))
        If InStr(1, lineText, """" & fieldName & """", vbTextCompare) > 0 Then
            Dim colonPos As Long: colonPos = InStr(1, lineText, ":")
            If colonPos > 0 Then
                Dim afterColon As String: afterColon = Trim(Mid(lineText, colonPos + 1))
                If Right(afterColon, 1) = "," Then afterColon = Left(afterColon, Len(afterColon) - 1)
                afterColon = Trim(afterColon)
                If Left(afterColon, 1) = """" And Right(afterColon, 1) = """" Then
                    afterColon = Mid(afterColon, 2, Len(afterColon) - 2)
                End If
                ReadJsonField = afterColon
            End If
            Exit For
        End If
    Next i
    Exit Function
ReadErr:
    ReadJsonField = ""
End Function

' WriteTextFile -- writes a string to a file.
Private Sub WriteTextFile(filePath As String, content As String)
    Dim fileNum As Integer: fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, content;
    Close #fileNum
End Sub


' =============================================================================
' PinWorkspace / UnpinWorkspace
' Toggles the pinned flag in workspace.json.
' =============================================================================
Public Sub PinWorkspace(workspaceName As String)
    SetPinnedFlag workspaceName, True
End Sub

Public Sub UnpinWorkspace(workspaceName As String)
    SetPinnedFlag workspaceName, False
End Sub

Private Sub SetPinnedFlag(workspaceName As String, pinned As Boolean)
    On Error Resume Next
    Dim root As String: root = KernelSnapshot.GetProjectRoot()
    Dim jsonPath As String
    jsonPath = root & "\" & DIR_WORKSPACES & "\" & workspaceName & "\" & WS_JSON
    If Dir(jsonPath) = "" Then Exit Sub
    Dim content As String
    content = KernelSnapshot.ReadEntireFile(jsonPath)
    If InStr(1, content, """pinned""") > 0 Then
        If pinned Then
            content = Replace(content, """pinned"": false", """pinned"": true")
        Else
            content = Replace(content, """pinned"": true", """pinned"": false")
        End If
    End If
    WriteTextFile jsonPath, content
    On Error GoTo 0
End Sub

Public Function IsWorkspacePinned(workspaceName As String) As Boolean
    On Error Resume Next
    Dim root As String: root = KernelSnapshot.GetProjectRoot()
    Dim jsonPath As String
    jsonPath = root & "\" & DIR_WORKSPACES & "\" & workspaceName & "\" & WS_JSON
    If Dir(jsonPath) = "" Then
        IsWorkspacePinned = False
        Exit Function
    End If
    Dim val As String
    val = ReadJsonField(jsonPath, "pinned")
    IsWorkspacePinned = (StrComp(val, "true", vbTextCompare) = 0)
    On Error GoTo 0
End Function


' =============================================================================
' CheckWorkspaceCompleteness
' Compares a workspace version against current tab_registry.
' Returns a human-readable report.
' =============================================================================
Public Function CheckWorkspaceCompleteness(workspaceName As String, _
    Optional version As Long = 0) As String
    On Error GoTo CompErr
    Dim root As String: root = KernelSnapshot.GetProjectRoot()
    Dim wsDir As String: wsDir = root & "\" & DIR_WORKSPACES & "\" & workspaceName
    If Dir(wsDir, vbDirectory) = "" Then
        CheckWorkspaceCompleteness = "Workspace not found."
        Exit Function
    End If
    Dim targetVer As Long
    If version = 0 Then targetVer = GetCurrentVersion(wsDir) Else targetVer = version
    Dim verDir As String: verDir = wsDir & "\" & FmtVer(targetVer)
    If Dir(verDir, vbDirectory) = "" Then
        CheckWorkspaceCompleteness = "Version not found."
        Exit Function
    End If

    Dim report As String
    report = "Workspace: " & workspaceName & " " & FmtVer(targetVer) & vbCrLf
    report = report & String(40, "-") & vbCrLf

    ' Check detail
    If Dir(verDir & "\detail.csv") <> "" Then
        report = report & "Detail: Present" & vbCrLf
    Else
        report = report & "Detail: MISSING (Run Model needed)" & vbCrLf
    End If

    ' Check input_tabs
    Dim itDir As String: itDir = verDir & "\input_tabs"
    Dim tabCount As Long: tabCount = 0
    If Dir(itDir, vbDirectory) <> "" Then
        Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
        Dim csvF As Object
        For Each csvF In fso.GetFolder(itDir).Files
            If StrComp(fso.GetExtensionName(csvF.Name), "csv", vbTextCompare) = 0 Then
                tabCount = tabCount + 1
            End If
        Next csvF
        Set fso = Nothing
    End If
    report = report & "Input tabs: " & tabCount & " CSV file(s)" & vbCrLf

    ' Config hash check
    Dim wsJsonPath As String: wsJsonPath = wsDir & "\" & WS_JSON
    If Dir(wsJsonPath) <> "" Then
        Dim mHash As String
        mHash = ReadJsonField(wsJsonPath, "configHash")
        If Len(mHash) > 0 Then
            Dim curHash As String: curHash = KernelSnapshot.BuildConfigHash()
            If StrComp(mHash, curHash, vbTextCompare) = 0 Then
                report = report & "Config: matches current model" & vbCrLf
            Else
                report = report & "Config: CHANGED since save" & vbCrLf
            End If
        End If
    End If

    ' Pinned status
    If IsWorkspacePinned(workspaceName) Then
        report = report & "Pinned: Yes (reference scenario)" & vbCrLf
    End If

    CheckWorkspaceCompleteness = report
    Exit Function
CompErr:
    CheckWorkspaceCompleteness = "Error: " & Err.Description
End Function
