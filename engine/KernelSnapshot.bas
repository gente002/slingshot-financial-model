Attribute VB_Name = "KernelSnapshot"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.
Private Const DIR_SNAPSHOTS As String = "snapshots"
Private Const SNAP_NAME_MAX As Long = 50

Public Sub SaveSnapshot(snapshotName As String, Optional description As String = "")
    On Error GoTo ErrHandler
    Dim safeName As String
    safeName = SanitizeName(snapshotName)
    If Len(safeName) = 0 Then
        KernelConfig.LogError SEV_ERROR, "KSnap", "E-800", _
            "Invalid name: " & snapshotName, _
            "MANUAL BYPASS: Use alphanumeric/underscore/hyphen only."
        MsgBox "Invalid snapshot name.", vbExclamation, "RDK"
        Exit Sub
    End If
    Dim staleAction As Long
    staleAction = KernelFormHelpers.CheckStalenessBeforeSave()
    If staleAction = 0 Then Exit Sub
    If staleAction = 1 Then KernelEngine.RunModel
    Dim root As String
    root = GetProjectRoot()
    Dim snapDir As String
    snapDir = root & "\" & DIR_SNAPSHOTS & "\" & safeName
    EnsureDirectoryExists root & "\" & DIR_SNAPSHOTS

    ' Name clash handling: if snapshot already exists, prompt user
    If Dir(snapDir, vbDirectory) <> "" Then
        Dim overwriteChoice As VbMsgBoxResult
        overwriteChoice = MsgBox("Snapshot '" & safeName & "' already exists." & vbCrLf & vbCrLf & _
            "Overwrite? The existing snapshot will be archived automatically.", _
            vbYesNo Or vbQuestion, "RDK -- Snapshot Exists")
        If overwriteChoice = vbNo Then
            MsgBox "Save cancelled.", vbInformation, "RDK"
            Exit Sub
        End If
        ' Archive existing snapshot with timestamp
        Dim archiveDir As String
        archiveDir = root & "\" & DIR_SNAPSHOTS & "\archive"
        EnsureDirectoryExists archiveDir
        Dim archiveName As String
        archiveName = safeName & "_" & Format(Now, "yyyymmdd_hhnnss")
        Dim archiveDest As String
        archiveDest = archiveDir & "\" & archiveName
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(snapDir) Then
            fso.MoveFolder snapDir, archiveDest
        End If
        Set fso = Nothing
        KernelConfig.LogError SEV_INFO, "KSnap", "I-805", _
            "Archived existing snapshot: " & safeName & " -> archive/" & archiveName, ""
    End If

    EnsureDirectoryExists snapDir
    If IsWALEnabled() Then WriteWAL "SNAP_START", safeName
    Dim dp As String, ip As String, ep As String, sp As String, cp As String
    dp = snapDir & "\detail.csv"
    KernelSnapshotIO.ExportDetailToFile dp
    ip = snapDir & "\inputs.csv"
    KernelSnapshotIO.ExportInputsToFile ip
    ' Export all input tabs (UW Inputs, Capital Activity, etc.)
    KernelTabIO.ExportAllInputTabs snapDir
    ep = snapDir & "\errorlog.csv"
    KernelSnapshotIO.ExportErrorLogToFile ep
    sp = snapDir & "\settings.csv"
    KernelSnapshotIO.ExportSettingsToFile sp
    Dim cfgHash As String
    cfgHash = BuildConfigHash()
    cp = snapDir & "\config_hash.txt"
    WriteTextFile cp, cfgHash
    ' Save full config directory so snapshot captures model definition
    KernelSnapshotIO.SaveConfigToSnapshot snapDir
    ' Hash files (skip if SnapshotHashEnabled=FALSE for faster saves)
    Dim hashEnabled As Boolean
    Dim hashSetting As String
    hashSetting = KernelConfig.GetReproSetting("SnapshotHashEnabled")
    hashEnabled = (StrComp(hashSetting, "FALSE", vbTextCompare) <> 0)
    Dim hD As String, hI As String, hE As String, hS As String, hC As String
    If hashEnabled Then
        hD = ComputeSHA256(dp)
        hI = ComputeSHA256(ip)
        hE = ComputeSHA256(ep)
        hS = ComputeSHA256(sp)
        hC = ComputeSHA256(cp)
    End If
    ' Phase 11A: Copy granular CSV to snapshot if it exists
    Dim granPath As String
    Dim hG As String
    hG = ""
    ' Config-driven domain module dispatch (no hardcoded InsuranceDomainEngine reference)
    Dim domMod As String
    domMod = KernelConfig.GetSetting("DomainModule")
    On Error Resume Next
    If Len(domMod) > 0 Then granPath = Application.Run(domMod & ".GranularCSVPath")
    If Err.Number <> 0 Then granPath = "": Err.Clear
    On Error GoTo ErrHandler
    If Len(granPath) > 0 Then
        If Dir(granPath) <> "" Then
            Dim granDest As String
            granDest = snapDir & "\granular_detail.csv"
            FileCopy granPath, granDest
            If hashEnabled Then hG = ComputeSHA256(granDest)
        End If
    End If
    Dim fp As String
    fp = BuildExecutionFingerprint()
    Dim manifestPath As String
    manifestPath = snapDir & "\manifest.json"
    WriteManifestJson manifestPath, safeName, description, cfgHash, fp, _
                      hD, hI, hE, hS, hC, hG
    Dim isStale As Boolean
    isStale = KernelFormHelpers.IsResultsStale()
    Dim elapsed As String
    elapsed = KernelFormHelpers.ReadRunStateValue(RS_KEY_TOTAL_ELAPSED)
    KernelFormHelpers.PatchManifestStaleAndElapsed manifestPath, isStale, elapsed
    If IsWALEnabled() Then WriteWAL "SNAP_DONE", safeName
    KernelConfig.LogError SEV_INFO, "KSnap", "I-800", _
        "Saved: " & safeName, snapDir
    KernelFormHelpers.ShowConfigMsgBox "SNAPSHOT_SAVED", "{NAME}", safeName
    Exit Sub
ErrHandler:
    If IsWALEnabled() Then
        On Error Resume Next
        WriteWAL "SNAP_FAIL", safeName & " | " & Err.Description
        On Error GoTo 0
    End If
    KernelConfig.LogError SEV_ERROR, "KSnap", "E-899", _
        "Save error: " & Err.Description, _
        "MANUAL BYPASS: Copy files to snapshot folder, create manifest.json."
    MsgBox "Save error: " & Err.Description, vbCritical, "RDK"
End Sub

Public Sub LoadSnapshot(snapshotName As String, Optional forensic As Boolean = False)
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Dim root As String
    root = GetProjectRoot()
    Dim snapDir As String
    snapDir = root & "\" & DIR_SNAPSHOTS & "\" & snapshotName
    If Dir(snapDir, vbDirectory) = "" Then
        KernelConfig.LogError SEV_ERROR, "KSnap", "E-810", _
            "Not found: " & snapshotName, _
            "MANUAL BYPASS: Verify snapshot folder exists."
        KernelFormHelpers.ShowConfigMsgBox "SNAPSHOT_NOT_FOUND", "{NAME}", snapshotName
        Exit Sub
    End If
    Dim manifestPath As String
    manifestPath = snapDir & "\manifest.json"
    If Dir(manifestPath) = "" Then
        If Not forensic Then
            KernelConfig.LogError SEV_ERROR, "KSnap", "E-811", _
                "No manifest.json", _
                "MANUAL BYPASS: Create manifest.json or use forensic."
            MsgBox "No manifest. Use forensic mode.", vbExclamation, "RDK"
            Exit Sub
        Else
            KernelConfig.LogError SEV_WARN, "KSnap", "W-811", _
                "FORENSIC: no manifest; unverified", ""
            GoTo ForensicRestore
        End If
    End If
    Dim mSt As String, mKV As String, mCH As String
    Dim mFN() As String, mFH() As String
    Dim mFC As Long
    ReadManifest manifestPath, mSt, mKV, mCH, mFN, mFH, mFC
    If StrComp(mSt, SP_STATUS_COMPLETE, vbTextCompare) <> 0 Then
        If Not forensic Then
            KernelConfig.LogError SEV_ERROR, "KSnap", "E-812", _
                "Status='" & mSt & "', expected COMPLETE", _
                "MANUAL BYPASS: Set status=COMPLETE or use forensic."
            MsgBox "Not COMPLETE (status=" & mSt & ").", vbExclamation, "RDK"
            Exit Sub
        Else
            KernelConfig.LogError SEV_WARN, "KSnap", "W-812", _
                "FORENSIC: status=" & mSt & "; continuing", ""
        End If
    End If
    Dim fIdx As Long
    For fIdx = 1 To mFC
        Dim fPath As String
        fPath = snapDir & "\" & mFN(fIdx)
        If Dir(fPath) = "" Then
            If Not forensic Then
                KernelConfig.LogError SEV_ERROR, "KSnap", "E-813", _
                    "Missing file: " & mFN(fIdx), _
                    "MANUAL BYPASS: Restore file or use forensic."
                MsgBox "CORRUPTED: Missing " & mFN(fIdx), vbCritical, "RDK"
                Exit Sub
            Else
                KernelConfig.LogError SEV_WARN, "KSnap", "W-813", _
                    "FORENSIC: missing " & mFN(fIdx) & "; skip", ""
                GoTo NextManifestFile
            End If
        End If
        If Not forensic Then
            Dim actualHash As String
            actualHash = ComputeSHA256(fPath)
            If Len(actualHash) > 0 And Len(mFH(fIdx)) > 0 Then
                If StrComp(actualHash, mFH(fIdx), vbTextCompare) <> 0 Then
                    KernelConfig.LogError SEV_ERROR, "KSnap", "E-814", _
                        "CORRUPTED: " & mFN(fIdx) & " checksum mismatch", _
                        "MANUAL BYPASS: Replace file or use forensic."
                    KernelFormHelpers.ShowConfigMsgBox "SNAPSHOT_CORRUPTED", "{DETAIL}", mFN(fIdx) & " checksum mismatch"
                    Exit Sub
                End If
            End If
        End If
NextManifestFile:
    Next fIdx
    If Len(mKV) > 0 And mKV <> KERNEL_VERSION Then
        KernelConfig.LogError SEV_WARN, "KSnap", "W-815", _
            "Ver mismatch: " & mKV & " vs " & KERNEL_VERSION, ""
    End If
    Dim cch As String
    cch = BuildConfigHash()
    If Len(mCH) > 0 And Len(cch) > 0 Then
        If StrComp(mCH, cch, vbTextCompare) <> 0 Then
            KernelConfig.LogError SEV_WARN, "KSnap", "W-816", _
                "Config hash differs from snapshot.", ""
        End If
    End If
ForensicRestore:
    ' Restore config directory from snapshot (if present), then reload from disk
    KernelSnapshotIO.RestoreConfigFromSnapshot snapDir
    ' Reload: disk CSVs -> Config sheet -> memory arrays
    KernelBootstrap.LoadConfigFromDisk
    KernelConfig.LoadAllConfig
    Dim ipPath As String
    ipPath = snapDir & "\inputs.csv"
    If Dir(ipPath) <> "" Then
        KernelSnapshotIO.ImportInputsFromCsv ipPath
    End If
    ' Restore all input tabs (UW Inputs, Capital Activity, etc.)
    KernelTabIO.ImportAllInputTabs snapDir
    Dim dtPath As String
    dtPath = snapDir & "\detail.csv"
    If Dir(dtPath) <> "" Then
        KernelSnapshotIO.ImportDetailFromCsv dtPath
    Else
        KernelConfig.LogError SEV_WARN, "KSnap", "W-821", _
            "detail.csv missing; skipped", ""
    End If
    Dim elPath As String
    elPath = snapDir & "\errorlog.csv"
    If Dir(elPath) <> "" Then
        KernelSnapshotIO.ImportErrorLogFromCsv elPath
    Else
        KernelConfig.LogError SEV_WARN, "KSnap", "W-822", _
            "errorlog.csv missing; skipped", ""
    End If
    Dim stPath As String
    stPath = snapDir & "\settings.csv"
    If Dir(stPath) <> "" Then
        KernelConfig.LogError SEV_INFO, "KSnap", "I-822", _
            "settings.csv found; review for config diffs.", ""
    End If
    ' Phase 11A: Restore granular CSV if present in snapshot
    Dim granSrc As String
    granSrc = snapDir & "\granular_detail.csv"
    If Dir(granSrc) <> "" Then
        Dim granRestDest As String
        granRestDest = KernelFormHelpers.EnsureOutputDir() & "\granular_detail_restored.csv"
        On Error Resume Next
        FileCopy granSrc, granRestDest
        On Error GoTo ErrHandler
        KernelConfig.LogError SEV_INFO, "KSnap", "I-825", _
            "Restored granular CSV: " & granRestDest, ""
    End If
    Dim entCnt As Long
    entCnt = KernelSnapshotIO.DetectEntityCount()
    Dim perCnt As Long
    perCnt = KernelConfig.GetTimeHorizon()
    If entCnt > 0 And perCnt > 0 Then
        KernelOutput.WriteSummaryFormulas entCnt, perCnt
    End If
    ' Refresh formula tabs so formulas recalculate against restored data
    On Error Resume Next
    KernelFormulaWriter.RefreshFormulaTabs
    On Error GoTo ErrHandler
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate

    ' Post-load verification: check health indicators
    Dim verifyResult As String
    verifyResult = KernelTabIO.VerifyHealthAfterLoad()

    If IsWALEnabled() Then
        Dim mode As String
        If forensic Then mode = "FORENSIC" Else mode = "NORMAL"
        WriteWAL "LOAD_COMPLETE", snapshotName & " | mode=" & mode & " | " & verifyResult
    End If
    KernelConfig.LogError SEV_INFO, "KSnap", "I-810", _
        "Loaded: " & snapshotName, verifyResult
    Dim snapStale As String
    snapStale = KernelFormHelpers.ReadJsonField(manifestPath, "resultsStale")

    Dim loadMsg As String
    loadMsg = "Loaded: " & snapshotName
    If StrComp(snapStale, "true", vbTextCompare) = 0 Then
        loadMsg = loadMsg & vbCrLf & "Warning: Stale when saved. Run model."
    End If
    If Len(verifyResult) > 0 Then
        loadMsg = loadMsg & vbCrLf & vbCrLf & "Health Check: " & verifyResult
    End If

    If InStr(1, verifyResult, "FAIL", vbTextCompare) > 0 Then
        MsgBox loadMsg, vbExclamation, "RDK -- Snapshot Loaded (with warnings)"
    Else
        MsgBox loadMsg, vbInformation, "RDK -- Snapshot Loaded"
    End If
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    KernelConfig.LogError SEV_ERROR, "KSnap", "E-819", _
        "Load error: " & Err.Description, _
        "MANUAL BYPASS: Copy CSV files to respective tabs."
    MsgBox "Load error: " & Err.Description, vbCritical, "RDK"
End Sub

Public Sub LoadSnapshotInputsOnly(snapshotName As String)
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Dim root As String
    root = GetProjectRoot()
    Dim snapDir As String
    snapDir = root & "\" & DIR_SNAPSHOTS & "\" & snapshotName
    If Dir(snapDir, vbDirectory) = "" Then
        KernelConfig.LogError SEV_ERROR, "KSnap", "E-830", _
            "Not found: " & snapshotName, _
            "MANUAL BYPASS: Verify snapshot folder exists."
        MsgBox "Not found: " & snapshotName, vbExclamation, "RDK"
        Exit Sub
    End If
    If Dir(snapDir & "\manifest.json") = "" Then
        KernelConfig.LogError SEV_ERROR, "KSnap", "E-831", _
            "No manifest.json", _
            "MANUAL BYPASS: Create manifest.json or copy inputs.csv."
        MsgBox "No manifest.json.", vbExclamation, "RDK"
        Exit Sub
    End If
    ' Restore config directory from snapshot (if present), then reload from disk
    KernelSnapshotIO.RestoreConfigFromSnapshot snapDir
    KernelBootstrap.LoadConfigFromDisk
    KernelConfig.LoadAllConfig
    Dim ipPath As String
    ipPath = snapDir & "\inputs.csv"
    If Dir(ipPath) <> "" Then
        KernelSnapshotIO.ImportInputsFromCsv ipPath
    End If
    ' Restore all input tabs
    KernelTabIO.ImportAllInputTabs snapDir
    If Dir(ipPath) = "" And Dir(snapDir & "\input_tabs", vbDirectory) = "" Then
        KernelConfig.LogError SEV_WARN, "KSnap", "W-830", _
            "No input data found in snapshot", _
            "MANUAL BYPASS: Place inputs.csv or input_tabs/ in snapshot folder."
        Application.ScreenUpdating = True
        MsgBox "No input data found.", vbExclamation, "RDK"
        Exit Sub
    End If
    KernelFormHelpers.WriteRunStateValue RS_KEY_STALE, "TRUE"
    Application.ScreenUpdating = True
    If IsWALEnabled() Then WriteWAL "LOAD_INPUTS", snapshotName
    KernelConfig.LogError SEV_INFO, "KSnap", "I-830", _
        "Inputs loaded (stale): " & snapshotName, ""
    MsgBox "Inputs loaded: " & snapshotName & vbCrLf & _
           "Results now stale. Run model.", vbInformation, "RDK"
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    KernelConfig.LogError SEV_ERROR, "KSnap", "E-839", _
        "Load inputs error: " & Err.Description, _
        "MANUAL BYPASS: Copy inputs.csv to Inputs tab."
    MsgBox "Load error: " & Err.Description, vbCritical, "RDK"
End Sub

Public Sub DeleteSnapshot(snapshotName As String)
    On Error GoTo ErrHandler
    Dim root As String
    root = GetProjectRoot()
    Dim snapDir As String
    snapDir = root & "\" & DIR_SNAPSHOTS & "\" & snapshotName
    If Dir(snapDir, vbDirectory) = "" Then
        MsgBox "Not found: " & snapshotName, vbExclamation, "RDK"
        Exit Sub
    End If
    If MsgBox("Delete '" & snapshotName & "'? Cannot be undone.", _
              vbYesNo Or vbQuestion, "RDK") = vbNo Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(snapDir) Then
        Dim f As Object
        For Each f In fso.GetFolder(snapDir).Files
            SetAttr f.Path, vbNormal
        Next f
        fso.DeleteFolder snapDir, True
    End If
    If IsWALEnabled() Then WriteWAL "SNAP_DEL", snapshotName
    KernelConfig.LogError SEV_INFO, "KSnap", "I-840", _
        "Deleted: " & snapshotName, ""
    MsgBox "Deleted: " & snapshotName, vbInformation, "RDK"
    Exit Sub
ErrHandler:
    KernelConfig.LogError SEV_ERROR, "KSnap", "E-849", _
        "Delete error: " & Err.Description, _
        "MANUAL BYPASS: Delete snapshot folder."
    MsgBox "Delete error: " & Err.Description, vbCritical, "RDK"
End Sub

Public Sub RenameSnapshot(oldName As String, newName As String)
    On Error GoTo ErrHandler
    Dim safeName As String
    safeName = SanitizeName(newName)
    If Len(safeName) = 0 Then
        KernelConfig.LogError SEV_ERROR, "KSnap", "E-870", _
            "Invalid name: " & newName, _
            "MANUAL BYPASS: Use alphanumeric/underscore/hyphen only."
        MsgBox "Invalid snapshot name.", vbExclamation, "RDK"
        Exit Sub
    End If
    Dim root As String
    root = GetProjectRoot()
    Dim oldDir As String
    oldDir = root & "\" & DIR_SNAPSHOTS & "\" & oldName
    Dim newDir As String
    newDir = root & "\" & DIR_SNAPSHOTS & "\" & safeName
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(oldDir) Then
        MsgBox "Not found: " & oldName, vbExclamation, "RDK"
        Exit Sub
    End If
    If fso.FolderExists(newDir) Then
        MsgBox "Already exists: " & safeName, vbExclamation, "RDK"
        Exit Sub
    End If
    fso.MoveFolder oldDir, newDir
    Dim manPath As String
    manPath = newDir & "\manifest.json"
    If fso.FileExists(manPath) Then
        Dim content As String
        content = ReadEntireFile(manPath)
        Dim oldField As String
        oldField = """snapshotName"": """ & Replace(oldName, """", "\""") & """"
        Dim newField As String
        newField = """snapshotName"": """ & Replace(safeName, """", "\""") & """"
        content = Replace(content, oldField, newField)
        Dim fileNum As Integer
        fileNum = FreeFile
        Open manPath For Output As #fileNum
        Print #fileNum, content
        Close #fileNum
    End If
    If IsWALEnabled() Then WriteWAL "SNAP_RENAME", oldName & " -> " & safeName
    KernelConfig.LogError SEV_INFO, "KSnap", "I-870", _
        "Renamed: " & oldName & " -> " & safeName, ""
    MsgBox "Renamed: " & oldName & " -> " & safeName, vbInformation, "RDK"
    Exit Sub
ErrHandler:
    KernelConfig.LogError SEV_ERROR, "KSnap", "E-879", _
        "Rename error: " & Err.Description, _
        "MANUAL BYPASS: Rename folder, update manifest.json."
    MsgBox "Rename error: " & Err.Description, vbCritical, "RDK"
End Sub

Public Function ListSnapshots() As String()
    Dim root As String
    root = GetProjectRoot()
    Dim snapRoot As String
    snapRoot = root & "\" & DIR_SNAPSHOTS
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(snapRoot) Then
        ListSnapshots = Split("", ",")
        Exit Function
    End If
    Dim parentFolder As Object
    Set parentFolder = fso.GetFolder(snapRoot)
    Dim cnt As Long
    cnt = 0
    Dim subFolder As Object
    For Each subFolder In parentFolder.SubFolders
        If fso.FileExists(subFolder.Path & "\manifest.json") Then
            cnt = cnt + 1
        End If
    Next subFolder
    If cnt = 0 Then
        ListSnapshots = Split("", ",")
        Exit Function
    End If
    Dim result() As String
    ReDim result(1 To cnt)
    Dim pos As Long
    pos = 0
    For Each subFolder In parentFolder.SubFolders
        If fso.FileExists(subFolder.Path & "\manifest.json") Then
            pos = pos + 1
            result(pos) = subFolder.Name
        End If
    Next subFolder
    ListSnapshots = result
End Function

Public Function ListArchivedSnapshots() As String()
    Dim root As String
    root = GetProjectRoot()
    Dim archRoot As String
    archRoot = root & "\" & DIR_ARCHIVE & "\snapshots"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(archRoot) Then
        ListArchivedSnapshots = Split("", ",")
        Exit Function
    End If
    Dim parentFolder As Object
    Set parentFolder = fso.GetFolder(archRoot)
    Dim cnt As Long
    cnt = 0
    Dim subFolder As Object
    For Each subFolder In parentFolder.SubFolders
        If fso.FileExists(subFolder.Path & "\manifest.json") Then
            cnt = cnt + 1
        End If
    Next subFolder
    If cnt = 0 Then
        ListArchivedSnapshots = Split("", ",")
        Exit Function
    End If
    Dim result() As String
    ReDim result(1 To cnt)
    Dim pos As Long
    pos = 0
    For Each subFolder In parentFolder.SubFolders
        If fso.FileExists(subFolder.Path & "\manifest.json") Then
            pos = pos + 1
            result(pos) = subFolder.Name
        End If
    Next subFolder
    ListArchivedSnapshots = result
End Function

Public Sub ArchiveSnapshot(snapshotName As String)
    On Error GoTo ErrHandler
    Dim root As String
    root = GetProjectRoot()
    Dim snapDir As String
    snapDir = root & "\" & DIR_SNAPSHOTS & "\" & snapshotName
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(snapDir) Then
        MsgBox "Not found: " & snapshotName, vbExclamation, "RDK"
        Exit Sub
    End If
    If Not fso.FileExists(snapDir & "\manifest.json") Then
        KernelConfig.LogError SEV_ERROR, "KSnap", "E-850", _
            "No manifest.json", _
            "MANUAL BYPASS: Create manifest.json then retry."
        MsgBox "No manifest.json.", vbExclamation, "RDK"
        Exit Sub
    End If
    Dim archiveDir As String
    archiveDir = root & "\" & DIR_ARCHIVE & "\snapshots"
    EnsureDirectoryExists root & "\" & DIR_ARCHIVE
    EnsureDirectoryExists archiveDir
    Dim archiveDest As String
    archiveDest = archiveDir & "\" & snapshotName
    If fso.FolderExists(archiveDest) Then
        Dim oldFile As Object
        For Each oldFile In fso.GetFolder(archiveDest).Files
            SetAttr oldFile.Path, vbNormal
        Next oldFile
        fso.DeleteFolder archiveDest, True
    End If
    EnsureDirectoryExists archiveDest
    Dim srcFile As Object
    For Each srcFile In fso.GetFolder(snapDir).Files
        fso.CopyFile srcFile.Path, archiveDest & "\" & srcFile.Name
        SetAttr archiveDest & "\" & srcFile.Name, vbReadOnly
    Next srcFile
    Dim origFile As Object
    For Each origFile In fso.GetFolder(snapDir).Files
        Kill origFile.Path
    Next origFile
    RmDir snapDir
    If IsWALEnabled() Then WriteWAL "ARCHIVE", snapshotName
    KernelConfig.LogError SEV_INFO, "KSnap", "I-850", _
        "Archived: " & snapshotName, archiveDest
    MsgBox "Archived: " & snapshotName, vbInformation, "RDK"
    Exit Sub
ErrHandler:
    KernelConfig.LogError SEV_ERROR, "KSnap", "E-859", _
        "Archive error: " & Err.Description, _
        "MANUAL BYPASS: Move snapshot to archive\snapshots\."
    MsgBox "Archive error: " & Err.Description, vbCritical, "RDK"
End Sub

Public Sub RestoreFromArchive(snapshotName As String)
    On Error GoTo ErrHandler
    Dim root As String
    root = GetProjectRoot()
    Dim archiveSrc As String
    archiveSrc = root & "\" & DIR_ARCHIVE & "\snapshots\" & snapshotName
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(archiveSrc) Then
        KernelConfig.LogError SEV_ERROR, "KSnap", "E-860", _
            "Archive not found: " & snapshotName, _
            "MANUAL BYPASS: Verify archived snapshot folder exists."
        MsgBox "Not found: " & snapshotName, vbExclamation, "RDK"
        Exit Sub
    End If
    Dim snapDir As String
    snapDir = root & "\" & DIR_SNAPSHOTS & "\" & snapshotName
    If fso.FolderExists(snapDir) Then
        KernelConfig.LogError SEV_ERROR, "KSnap", "E-861", _
            "Already exists: " & snapshotName, _
            "MANUAL BYPASS: Delete or rename existing snapshot."
        MsgBox "Already exists.", vbExclamation, "RDK"
        Exit Sub
    End If
    EnsureDirectoryExists root & "\" & DIR_SNAPSHOTS
    EnsureDirectoryExists snapDir
    Dim srcFile As Object
    For Each srcFile In fso.GetFolder(archiveSrc).Files
        fso.CopyFile srcFile.Path, snapDir & "\" & srcFile.Name
        SetAttr snapDir & "\" & srcFile.Name, vbNormal
    Next srcFile
    Dim archFile As Object
    For Each archFile In fso.GetFolder(archiveSrc).Files
        SetAttr archFile.Path, vbNormal
    Next archFile
    fso.DeleteFolder archiveSrc, True
    If IsWALEnabled() Then WriteWAL "RESTORE_ARCHIVE", snapshotName
    KernelConfig.LogError SEV_INFO, "KSnap", "I-860", _
        "Restored: " & snapshotName, ""
    MsgBox "Restored: " & snapshotName, vbInformation, "RDK"
    Exit Sub
ErrHandler:
    KernelConfig.LogError SEV_ERROR, "KSnap", "E-869", _
        "Restore error: " & Err.Description, _
        "MANUAL BYPASS: Copy archive to snapshots\."
    MsgBox "Restore error: " & Err.Description, vbCritical, "RDK"
End Sub

Public Sub WriteWAL(operation As String, detail As String)
    On Error Resume Next
    Dim root As String
    root = GetProjectRoot()
    Dim walDir As String
    walDir = root & "\" & DIR_WAL
    EnsureDirectoryExists walDir
    Dim walPath As String
    walPath = walDir & "\wal.log"
    Dim fileNum As Integer
    fileNum = FreeFile
    Open walPath For Append As #fileNum
    Print #fileNum, FormatISOTimestamp() & "," & operation & "," & detail
    Close #fileNum
    On Error GoTo 0
End Sub

Public Sub PurgeWAL(Optional daysToKeep As Long = 0)
    On Error GoTo PurgeError
    If daysToKeep = 0 Then
        Dim retDays As String
        retDays = KernelConfig.GetReproSetting("WALRetentionDays")
        If IsNumeric(retDays) And Len(retDays) > 0 Then
            daysToKeep = CLng(retDays)
        Else
            daysToKeep = 30
        End If
    End If
    Dim root As String
    root = GetProjectRoot()
    Dim walPath As String
    walPath = root & "\" & DIR_WAL & "\wal.log"
    If Dir(walPath) = "" Then
        MsgBox "No WAL file found.", vbInformation, "RDK"
        Exit Sub
    End If
    Dim lines() As String
    lines = ReadFileLinesFromPath(walPath)
    Dim cutoffDate As Date
    cutoffDate = DateAdd("d", -daysToKeep, Now)
    Dim kept As String
    kept = ""
    Dim purgedCount As Long
    purgedCount = 0
    Dim i As Long
    For i = 0 To UBound(lines)
        If Len(Trim(lines(i))) = 0 Then GoTo NextWALLine
        Dim commaPos As Long
        commaPos = InStr(1, lines(i), ",")
        If commaPos > 0 Then
            Dim tsStr As String
            tsStr = Left(lines(i), commaPos - 1)
            Dim entryDate As Date
            On Error Resume Next
            entryDate = CDate(Replace(tsStr, "T", " "))
            If Err.Number = 0 Then
                On Error GoTo PurgeError
                If entryDate >= cutoffDate Then
                    kept = kept & lines(i) & vbCrLf
                Else
                    purgedCount = purgedCount + 1
                End If
            Else
                On Error GoTo PurgeError
                kept = kept & lines(i) & vbCrLf
            End If
        Else
            kept = kept & lines(i) & vbCrLf
        End If
NextWALLine:
    Next i
    WriteTextFile walPath, kept
    KernelConfig.LogError SEV_INFO, "KSnap", "I-845", _
        "WAL purged: " & purgedCount & " (>" & daysToKeep & "d)", ""
    MsgBox "WAL purged: " & purgedCount & " removed.", vbInformation, "RDK"
    Exit Sub
PurgeError:
    KernelConfig.LogError SEV_ERROR, "KSnap", "E-845", _
        "Purge error: " & Err.Description, _
        "MANUAL BYPASS: Edit wal\wal.log manually."
    MsgBox "Purge error: " & Err.Description, vbCritical, "RDK"
End Sub

Public Function ComputeSHA256(filePath As String) As String
    On Error GoTo HashError
    ComputeSHA256 = ""
    If Dir(filePath) = "" Then Exit Function
    Dim tempResult As String
    tempResult = Environ("TEMP") & "\rdk_sha256_" & Format(Timer * 1000, "0") & ".txt"
    Dim psCmd As String
    psCmd = "powershell -NoProfile -Command ""(Get-FileHash '" & Replace(filePath, "'", "''") & _
            "' -Algorithm SHA256).Hash|Out-File '" & Replace(tempResult, "'", "''") & "' -Encoding ASCII -NoNewline"""
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run "cmd /c " & psCmd, 0, True
    Set wsh = Nothing
    If Dir(tempResult) <> "" Then
        Dim fileNum As Integer
        fileNum = FreeFile
        Dim hashVal As String
        Open tempResult For Input As #fileNum
        If Not EOF(fileNum) Then
            Line Input #fileNum, hashVal
        End If
        Close #fileNum
        Kill tempResult
        ComputeSHA256 = Trim(hashVal)
    End If
    Exit Function
HashError:
    KernelConfig.LogError SEV_WARN, "KSnap", "W-850", _
        "SHA256 failed: " & filePath & ": " & Err.Description, ""
    ComputeSHA256 = ""
End Function

Public Function BuildConfigHash() As String
    On Error GoTo HashError
    Dim root As String
    root = GetProjectRoot()
    Dim configDir As String
    configDir = root & "\config"
    If Dir(configDir, vbDirectory) = "" Then
        BuildConfigHash = ""
        Exit Function
    End If
    ' Hash ALL CSVs in config/ (not a hardcoded subset)
    Dim combined As String
    combined = ""
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim csvFile As Object
    For Each csvFile In fso.GetFolder(configDir).Files
        If StrComp(fso.GetExtensionName(csvFile.Name), "csv", vbTextCompare) = 0 Then
            Dim fc As String
            fc = ReadEntireFile(csvFile.Path)
            combined = combined & fc
        End If
    Next csvFile
    Set fso = Nothing
    Dim tempConcat As String
    tempConcat = Environ("TEMP") & "\rdk_confighash_" & Format(Timer * 1000, "0") & ".txt"
    WriteTextFile tempConcat, combined
    BuildConfigHash = ComputeSHA256(tempConcat)
    On Error Resume Next
    Kill tempConcat
    On Error GoTo 0
    Exit Function
HashError:
    KernelConfig.LogError SEV_WARN, "KSnap", "W-851", _
        "Config hash failed: " & Err.Description, ""
    BuildConfigHash = ""
End Function

Public Function BuildExecutionFingerprint() As String
    On Error Resume Next
    Dim fp As String
    fp = ""
    Dim Q As String
    Q = """"
    fp = "    ""excelVersion"": " & Q & Application.Version & Q & "," & vbCrLf
    fp = fp & "    ""excelBuild"": " & Q & Application.Build & Q & "," & vbCrLf
    #If Win64 Then
        fp = fp & "    ""excelBitness"": ""64""," & vbCrLf
    #Else
        fp = fp & "    ""excelBitness"": ""32""," & vbCrLf
    #End If
    fp = fp & "    ""osVersion"": " & Q & Environ("OS") & Q & "," & vbCrLf
    fp = fp & "    ""machineName"": " & Q & Environ("COMPUTERNAME") & Q & "," & vbCrLf
    fp = fp & "    ""userName"": " & Q & Environ("USERNAME") & Q & "," & vbCrLf
    fp = fp & "    ""calculationMode"": " & Q & CStr(Application.Calculation) & Q & "," & vbCrLf
    fp = fp & "    ""decimalSeparator"": " & Q & Application.DecimalSeparator & Q & "," & vbCrLf
    fp = fp & "    ""thousandsSeparator"": " & Q & Application.ThousandsSeparator & Q & "," & vbCrLf
    fp = fp & "    ""useSystemSeparators"": " & IIf(Application.UseSystemSeparators, "true", "false") & "," & vbCrLf
    fp = fp & "    ""precisionAsDisplayed"": " & IIf(ThisWorkbook.PrecisionAsDisplayed, "true", "false")
    On Error GoTo 0
    BuildExecutionFingerprint = fp
End Function

Public Function AcquireLock(filePath As String) As Boolean
    On Error GoTo LockError
    Dim lockPath As String
    lockPath = filePath & ".lock"
    If Dir(lockPath) <> "" Then
        Dim lockContent As String
        lockContent = ReadEntireFile(lockPath)
        Dim tsLine As String
        Dim lockLines() As String
        lockLines = Split(lockContent, vbCrLf)
        Dim lockIdx As Long
        For lockIdx = 0 To UBound(lockLines)
            If InStr(1, lockLines(lockIdx), "TIMESTAMP=", vbTextCompare) > 0 Then
                tsLine = Mid(lockLines(lockIdx), Len("TIMESTAMP=") + 1)
            End If
        Next lockIdx
        If Len(tsLine) > 0 Then
            Dim lockTime As Date
            On Error Resume Next
            lockTime = CDate(tsLine)
            On Error GoTo LockError
            If DateDiff("s", lockTime, Now) > LOCK_TIMEOUT_SECONDS Then
                KernelConfig.LogError SEV_WARN, "KSnap", "W-860", _
                    "Stale lock (>" & LOCK_TIMEOUT_SECONDS & "s); removing", lockPath
                Kill lockPath
            Else
                AcquireLock = False
                Exit Function
            End If
        Else
            Kill lockPath
        End If
    End If
    Dim fileNum As Integer
    fileNum = FreeFile
    Open lockPath For Output As #fileNum
    Print #fileNum, "OWNER=" & Environ("USERNAME")
    Print #fileNum, "MACHINE=" & Environ("COMPUTERNAME")
    Print #fileNum, "TIMESTAMP=" & FormatISOTimestamp()
    Close #fileNum
    AcquireLock = True
    Exit Function
LockError:
    KernelConfig.LogError SEV_WARN, "KSnap", "W-861", _
        "Lock failed: " & filePath & ": " & Err.Description, ""
    AcquireLock = False
End Function

Public Sub ReleaseLock(filePath As String)
    On Error Resume Next
    Dim lockPath As String
    lockPath = filePath & ".lock"
    If Dir(lockPath) <> "" Then
        Dim lockContent As String
        lockContent = ReadEntireFile(lockPath)
        If InStr(1, lockContent, "OWNER=" & Environ("USERNAME"), vbTextCompare) > 0 And _
           InStr(1, lockContent, "MACHINE=" & Environ("COMPUTERNAME"), vbTextCompare) > 0 Then
            Kill lockPath
        End If
    End If
    On Error GoTo 0
End Sub

Private Function SanitizeName(rawName As String) As String
    Dim result As String
    result = Trim(rawName)
    result = Replace(result, " ", "_")
    Dim cleaned As String
    cleaned = ""
    Dim i As Long
    For i = 1 To Len(result)
        Dim ch As String
        ch = Mid(result, i, 1)
        If ch Like "[A-Za-z0-9_-]" Then
            cleaned = cleaned & ch
        End If
    Next i
    If Len(cleaned) > SNAP_NAME_MAX Then
        cleaned = Left(cleaned, SNAP_NAME_MAX)
    End If
    SanitizeName = cleaned
End Function

' GetInputsSheet -- returns the inputs tab using the configured name
Public Function GetInputsSheet() As Worksheet
    On Error Resume Next
    Set GetInputsSheet = ThisWorkbook.Sheets(KernelConfig.GetInputsTabName())
    On Error GoTo 0
End Function

Public Function GetProjectRoot() As String
    Dim wbPath As String
    wbPath = ThisWorkbook.Path
    If Len(wbPath) = 0 Then
        Err.Raise vbObjectError + 800, "KSnap", _
            "Workbook not saved to disk. Save it first. " & _
            "MANUAL BYPASS: Save workbook, then retry."
    End If
    Dim lastSep As Long
    lastSep = InStrRev(wbPath, "\")
    If lastSep <= 1 Then
        Err.Raise vbObjectError + 801, "KSnap", _
            "Workbook at drive root. Move to a project subfolder."
    End If
    GetProjectRoot = Left(wbPath, lastSep - 1)
End Function

' SaveConfigToSnapshot / RestoreConfigFromSnapshot moved to KernelSnapshotIO.bas


Public Sub EnsureDirectoryExists(dirPath As String)
    On Error GoTo DirErr
    If Dir(dirPath, vbDirectory) = "" Then
        MkDir dirPath
    End If
    Exit Sub
DirErr:
    Err.Raise vbObjectError + 802, "KSnap", _
        "Cannot create folder: " & dirPath & " (" & Err.Description & "). " & _
        "MANUAL BYPASS: Create the folder manually and retry."
End Sub

Public Function FormatISOTimestamp() As String
    FormatISOTimestamp = CStr(Year(Now)) & "-" & _
        Right("0" & CStr(Month(Now)), 2) & "-" & _
        Right("0" & CStr(Day(Now)), 2) & "T" & _
        Right("0" & CStr(Hour(Now)), 2) & ":" & _
        Right("0" & CStr(Minute(Now)), 2) & ":" & _
        Right("0" & CStr(Second(Now)), 2)
End Function

Private Function IsWALEnabled() As Boolean
    Dim val As String
    val = KernelConfig.GetReproSetting("WALEnabled")
    IsWALEnabled = (StrComp(val, "TRUE", vbTextCompare) = 0)
End Function

' Export/import functions moved to KernelSnapshotIO.bas (AD-09 split)

Private Sub WriteManifestJson(jsonPath As String, snapName As String, _
                              description As String, configHash As String, _
                              fingerprint As String, hashDetail As String, _
                              hashInputs As String, hashErrorlog As String, _
                              hashSettings As String, hashConfigHash As String, _
                              Optional hashGranular As String = "")
    Dim prngSeed As Long
    If KernelRandom.IsInitialized() Then
        prngSeed = KernelRandom.GetSeed()
    Else
        prngSeed = 0
    End If
    Dim fileNum As Integer
    fileNum = FreeFile
    Open jsonPath For Output As #fileNum
    Print #fileNum, "{"
    Print #fileNum, "  ""snapshotName"": """ & Replace(snapName, """", "\""") & ""","
    Print #fileNum, "  ""status"": """ & SP_STATUS_COMPLETE & ""","
    Print #fileNum, "  ""created"": """ & FormatISOTimestamp() & ""","
    Print #fileNum, "  ""description"": """ & Replace(description, """", "\""") & ""","
    Print #fileNum, "  ""kernelVersion"": """ & KERNEL_VERSION & ""","
    Print #fileNum, "  ""xlsxFilename"": ""workspace.xlsm"","
    Print #fileNum, "  ""configHash"": """ & configHash & ""","
    Print #fileNum, "  ""prngSeed"": " & prngSeed & ","
    Dim mColCount As Long
    mColCount = KernelConfig.GetColumnCount()
    Print #fileNum, "  ""columnCount"": " & mColCount & ","
    Dim hdrJson As String
    hdrJson = "  ""columnHeaders"": ["
    Dim ci As Long
    For ci = 1 To mColCount
        If ci > 1 Then hdrJson = hdrJson & ","
        hdrJson = hdrJson & """" & KernelConfig.GetColName(ci) & """"
    Next ci
    hdrJson = hdrJson & "],"
    Print #fileNum, hdrJson
    Print #fileNum, "  ""entityCount"": " & KernelSnapshotIO.DetectEntityCount() & ","
    Print #fileNum, "  ""timeHorizon"": " & KernelConfig.GetTimeHorizon() & ","
    ' Run metadata
    Dim lastRunAt As String
    lastRunAt = KernelFormHelpers.ReadRunStateValue(RS_KEY_TIMESTAMP)
    Dim runDuration As String
    runDuration = KernelFormHelpers.ReadRunStateValue(RS_KEY_TOTAL_ELAPSED)
    Print #fileNum, "  ""lastRunAt"": """ & lastRunAt & ""","
    Print #fileNum, "  ""runDuration"": """ & runDuration & "s"","
    Print #fileNum, "  ""fingerprint"": {"
    Print #fileNum, fingerprint
    Print #fileNum, "  },"
    Print #fileNum, "  ""files"": ["
    Print #fileNum, "    {""name"":""detail.csv"",""sha256"":""" & hashDetail & """},"
    Print #fileNum, "    {""name"":""inputs.csv"",""sha256"":""" & hashInputs & """},"
    Print #fileNum, "    {""name"":""errorlog.csv"",""sha256"":""" & hashErrorlog & """},"
    Print #fileNum, "    {""name"":""settings.csv"",""sha256"":""" & hashSettings & """},"
    If Len(hashGranular) > 0 Then
        Print #fileNum, "    {""name"":""config_hash.txt"",""sha256"":""" & hashConfigHash & """},"
        Print #fileNum, "    {""name"":""granular_detail.csv"",""sha256"":""" & hashGranular & """}"
    Else
        Print #fileNum, "    {""name"":""config_hash.txt"",""sha256"":""" & hashConfigHash & """}"
    End If
    Print #fileNum, "  ]"
    Print #fileNum, "}"
    Close #fileNum
End Sub

Private Sub ReadManifest(jsonPath As String, ByRef status As String, _
                         ByRef kernelVer As String, ByRef configHash As String, _
                         ByRef fileNames() As String, ByRef fileHashes() As String, _
                         ByRef fileCount As Long)
    status = ""
    kernelVer = ""
    configHash = ""
    fileCount = 0
    Dim lines() As String
    lines = ReadFileLinesFromPath(jsonPath)
    ReDim fileNames(1 To 10)
    ReDim fileHashes(1 To 10)
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim lineText As String
        lineText = Trim(lines(i))
        If InStr(1, lineText, """status""", vbTextCompare) > 0 Then
            status = ExtractJsonString(lineText)
        ElseIf InStr(1, lineText, """kernelVersion""", vbTextCompare) > 0 Then
            kernelVer = ExtractJsonString(lineText)
        ElseIf InStr(1, lineText, """configHash""", vbTextCompare) > 0 And _
               InStr(1, lineText, """sha256""", vbTextCompare) = 0 Then
            configHash = ExtractJsonString(lineText)
        ElseIf InStr(1, lineText, """name""", vbTextCompare) > 0 And _
               InStr(1, lineText, """sha256""", vbTextCompare) > 0 Then
            fileCount = fileCount + 1
            If fileCount > UBound(fileNames) Then
                ReDim Preserve fileNames(1 To fileCount + 5)
                ReDim Preserve fileHashes(1 To fileCount + 5)
            End If
            Dim nameStart As Long
            nameStart = InStr(1, lineText, """name""", vbTextCompare)
            fileNames(fileCount) = ExtractJsonStringAt(lineText, nameStart)
            Dim hashStart As Long
            hashStart = InStr(1, lineText, """sha256""", vbTextCompare)
            fileHashes(fileCount) = ExtractJsonStringAt(lineText, hashStart)
        End If
    Next i
End Sub

Public Function ReadEntireFile(filePath As String) As String
    Dim fileNum As Integer
    fileNum = FreeFile
    Dim fileSize As Long
    Open filePath For Binary Access Read As #fileNum
    fileSize = LOF(fileNum)
    If fileSize = 0 Then
        Close #fileNum
        ReadEntireFile = ""
        Exit Function
    End If
    ReadEntireFile = Space$(fileSize)
    Get #fileNum, , ReadEntireFile
    Close #fileNum
End Function

Public Function ReadFileLinesFromPath(filePath As String) As String()
    Dim content As String
    content = ReadEntireFile(filePath)
    content = Replace(content, vbCrLf, vbLf)
    content = Replace(content, vbCr, vbLf)
    If Right(content, 1) = vbLf Then
        content = Left(content, Len(content) - 1)
    End If
    ReadFileLinesFromPath = Split(content, vbLf)
End Function

Public Sub WriteTextFile(filePath As String, content As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, content;
    Close #fileNum
End Sub

Public Function ParseCsvLine(lineText As String) As String()
    Dim result() As String
    Dim fieldCount As Long
    fieldCount = 0
    Dim maxFields As Long
    maxFields = 1
    Dim scanPos As Long
    For scanPos = 1 To Len(lineText)
        If Mid(lineText, scanPos, 1) = "," Then maxFields = maxFields + 1
    Next scanPos
    ReDim result(0 To maxFields - 1)
    Dim pos As Long
    pos = 1
    Dim lineLen As Long
    lineLen = Len(lineText)
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

Private Function ExtractJsonString(lineText As String) As String
    Dim colonPos As Long
    colonPos = InStr(1, lineText, ":")
    If colonPos = 0 Then
        ExtractJsonString = ""
        Exit Function
    End If
    Dim afterColon As String
    afterColon = Trim(Mid(lineText, colonPos + 1))
    If Right(afterColon, 1) = "," Then
        afterColon = Left(afterColon, Len(afterColon) - 1)
    End If
    afterColon = Trim(afterColon)
    If Left(afterColon, 1) = """" And Right(afterColon, 1) = """" Then
        afterColon = Mid(afterColon, 2, Len(afterColon) - 2)
    End If
    ExtractJsonString = afterColon
End Function

Private Function ExtractJsonStringAt(lineText As String, startPos As Long) As String
    Dim colonPos As Long
    colonPos = InStr(startPos, lineText, ":")
    If colonPos = 0 Then
        ExtractJsonStringAt = ""
        Exit Function
    End If
    Dim quoteStart As Long
    quoteStart = InStr(colonPos + 1, lineText, """")
    If quoteStart = 0 Then
        ExtractJsonStringAt = ""
        Exit Function
    End If
    Dim quoteEnd As Long
    quoteEnd = InStr(quoteStart + 1, lineText, """")
    If quoteEnd = 0 Then
        ExtractJsonStringAt = ""
        Exit Function
    End If
    ExtractJsonStringAt = Mid(lineText, quoteStart + 1, quoteEnd - quoteStart - 1)
End Function
