Attribute VB_Name = "KernelHealth"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelHealth.bas
' Purpose: Proactive checks on workbook open. Catches drift, stale state,
'          and environmental issues before the user runs anything.
' Phase 4: Observability + Hardening
' =============================================================================

' =============================================================================
' RunHealthCheck
' Runs health checks and reports results.
' mode: "LIGHTWEIGHT" (default on open) or "FULL" (all checks).
' =============================================================================
Public Sub RunHealthCheck(Optional mode As String = "LIGHTWEIGHT", Optional silent As Boolean = False)
    On Error GoTo ErrHandler

    Dim issues() As String
    Dim issueCount As Long
    issueCount = 0
    ReDim issues(1 To 50)

    Dim isLightweight As Boolean
    isLightweight = (StrComp(mode, HEALTH_LIGHTWEIGHT, vbTextCompare) = 0)

    ' HC-05: Kernel version mismatch (in-memory only -- safe on open)
    CheckKernelVersion issues, issueCount

    If Not isLightweight Then
        ' HC-01: Config drift (shells out to PowerShell for SHA256)
        CheckConfigDrift issues, issueCount

        ' HC-02: Incomplete snapshots (creates FileSystemObject COM)
        CheckIncompleteSnapshots issues, issueCount

        ' HC-03: Stale lock files (creates FileSystemObject COM)
        CheckStaleLocks issues, issueCount

        ' HC-04: Module size
        CheckModuleSizes issues, issueCount

        ' HC-06: Detail tab staleness
        CheckDetailStaleness issues, issueCount

        ' HC-07: ErrorLog overflow
        CheckErrorLogOverflow issues, issueCount

        ' HC-08: WAL overflow
        CheckWALOverflow issues, issueCount

        ' HC-09: Prove-It status
        CheckProveItStatus issues, issueCount
    End If

    ' Build results message
    Dim summary As String
    If issueCount = 0 Then
        summary = "Health Check: All clear."
        KernelConfig.LogError SEV_INFO, "KernelHealth", "I-760", summary, "Mode=" & mode
    Else
        summary = "Health Check: " & issueCount & " issue(s) found" & vbCrLf & vbCrLf
        Dim i As Long
        For i = 1 To issueCount
            summary = summary & issues(i) & vbCrLf
        Next i
        KernelConfig.LogError SEV_WARN, "KernelHealth", "W-760", _
            issueCount & " health check issues found", "Mode=" & mode
        ' Log each issue
        For i = 1 To issueCount
            KernelConfig.LogError SEV_INFO, "KernelHealth", "I-761", issues(i), ""
        Next i
    End If

    If Not silent Then
        MsgBox summary, IIf(issueCount = 0, vbInformation, vbExclamation), "RDK -- Health Check"
    End If
    Exit Sub

ErrHandler:
    KernelConfig.LogError SEV_ERROR, "KernelHealth", "E-769", _
        "Health check failed: " & Err.Description, _
        "MANUAL BYPASS: Check config/ directory timestamps against last run. " & _
        "Inspect snapshots/ for directories without manifest.json. " & _
        "Delete any .lock files older than 1 minute."
    If Not silent Then
        MsgBox "Health check failed: " & Err.Description, vbCritical, "RDK -- Health Check"
    End If
End Sub


' =============================================================================
' RunHealthCheckFull
' Parameterless entry point for Full health check (for Dashboard button).
' =============================================================================
Public Sub RunHealthCheckFull()
    RunHealthCheck HEALTH_FULL
End Sub


' =============================================================================
' RunHealthCheckOnOpen
' Called from Workbook_Open event (if configured).
' Runs LIGHTWEIGHT mode only. Can be disabled via ReproConfig.
' =============================================================================
Public Sub RunHealthCheckOnOpen()
    On Error Resume Next

    ' BUG-081: Lightweight on-open must NOT create COM objects, shell out
    ' to external processes, or load config arrays. These operations can
    ' hang Workbook_Open, causing Excel to flag the file and prompt for
    ' safe mode on next open. Only HC-05 (kernel version) runs here --
    ' it reads in-memory CustomDocumentProperties only.
    RunHealthCheck HEALTH_LIGHTWEIGHT, True

    On Error GoTo 0
End Sub


' =============================================================================
' HC-01: Config drift
' Compare config hash stored in last pipeline state vs current config hash.
' =============================================================================
Private Sub CheckConfigDrift(ByRef issues() As String, ByRef issueCount As Long)
    On Error Resume Next

    ' Read stored config hash from RUN_STATE
    Dim storedHash As String
    storedHash = KernelFormHelpers.ReadRunStateValue(RS_KEY_CONFIG_HASH)

    If Len(storedHash) = 0 Then
        ' No stored hash -- no prior run to compare against
        Exit Sub
    End If

    ' Compute current config hash
    Dim currentHash As String
    currentHash = KernelSnapshot.BuildConfigHash()

    If Len(currentHash) = 0 Then Exit Sub

    If StrComp(storedHash, currentHash, vbTextCompare) <> 0 Then
        AddIssue issues, issueCount, "[WARN] HC-01: Config files changed since last run. " & _
            "Results may be stale. Consider re-running the model."
    End If

    On Error GoTo 0
End Sub


' =============================================================================
' HC-02: Incomplete snapshots
' Scan snapshots/ for directories without manifest.json or incomplete status.
' =============================================================================
Private Sub CheckIncompleteSnapshots(ByRef issues() As String, ByRef issueCount As Long)
    On Error Resume Next

    Dim projectRoot As String
    projectRoot = ThisWorkbook.Path & "\.."
    Dim snapRoot As String
    snapRoot = projectRoot & "\snapshots"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(snapRoot) Then
        Set fso = Nothing
        Exit Sub
    End If

    Dim folder As Object
    Set folder = fso.GetFolder(snapRoot)
    Dim subFolder As Object

    For Each subFolder In folder.SubFolders
        Dim manifestPath As String
        manifestPath = subFolder.Path & "\manifest.json"
        If Not fso.FileExists(manifestPath) Then
            AddIssue issues, issueCount, "[WARN] HC-02: Incomplete snapshot '" & _
                subFolder.Name & "' -- no manifest.json"
        Else
            ' Check status field in manifest
            Dim snapStatus As String
            snapStatus = KernelFormHelpers.ReadJsonField(manifestPath, "status")
            If Len(snapStatus) > 0 And StrComp(snapStatus, "COMPLETE", vbTextCompare) <> 0 Then
                AddIssue issues, issueCount, "[WARN] HC-02: Snapshot '" & _
                    subFolder.Name & "' has status: " & snapStatus
            End If
        End If
    Next subFolder

    Set folder = Nothing
    Set fso = Nothing
    On Error GoTo 0
End Sub


' =============================================================================
' HC-03: Stale lock files
' Scan for .lock files older than 60 seconds.
' =============================================================================
Private Sub CheckStaleLocks(ByRef issues() As String, ByRef issueCount As Long)
    On Error Resume Next

    Dim projectRoot As String
    projectRoot = ThisWorkbook.Path & "\.."

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(projectRoot) Then
        Set fso = Nothing
        Exit Sub
    End If

    ' Scan project root and key subdirectories for .lock files
    ScanForStaleLocks fso, projectRoot, issues, issueCount
    If fso.FolderExists(projectRoot & "\snapshots") Then
        ScanForStaleLocks fso, projectRoot & "\snapshots", issues, issueCount
    End If
    If fso.FolderExists(projectRoot & "\wal") Then
        ScanForStaleLocks fso, projectRoot & "\wal", issues, issueCount
    End If

    Set fso = Nothing
    On Error GoTo 0
End Sub

Private Sub ScanForStaleLocks(fso As Object, folderPath As String, _
                               ByRef issues() As String, ByRef issueCount As Long)
    On Error Resume Next
    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)
    Dim fileObj As Object

    For Each fileObj In folder.Files
        If LCase(fso.GetExtensionName(fileObj.Name)) = "lock" Then
            ' Check age
            Dim ageSec As Long
            ageSec = DateDiff("s", fileObj.DateLastModified, Now)
            If ageSec > LOCK_TIMEOUT_SECONDS Then
                AddIssue issues, issueCount, "[WARN] HC-03: Stale lock file '" & _
                    fileObj.Name & "' in " & folderPath & " (age: " & ageSec & "s)"
            End If
        End If
    Next fileObj

    Set folder = Nothing
    On Error GoTo 0
End Sub


' =============================================================================
' HC-04: Module size check
' =============================================================================
Private Sub CheckModuleSizes(ByRef issues() As String, ByRef issueCount As Long)
    On Error Resume Next

    Dim enginePath As String
    enginePath = ThisWorkbook.Path & "\..\engine"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(enginePath) Then
        Set fso = Nothing
        Exit Sub
    End If

    Dim folder As Object
    Set folder = fso.GetFolder(enginePath)
    Dim fileObj As Object

    For Each fileObj In folder.Files
        If LCase(fso.GetExtensionName(fileObj.Name)) = "bas" Then
            If fileObj.Size > MODULE_SIZE_ERROR Then
                AddIssue issues, issueCount, "[ERROR] HC-04: " & fileObj.Name & _
                    " is " & Format(fileObj.Size / 1024, "0.0") & " KB (exceeds 64KB limit)"
            ElseIf fileObj.Size > MODULE_SIZE_WARN Then
                AddIssue issues, issueCount, "[WARN] HC-04: " & fileObj.Name & _
                    " is " & Format(fileObj.Size / 1024, "0.0") & " KB (exceeds 50KB WARN threshold)"
            End If
        End If
    Next fileObj

    Set folder = Nothing
    Set fso = Nothing
    On Error GoTo 0
End Sub


' =============================================================================
' HC-05: Kernel version mismatch
' =============================================================================
Private Sub CheckKernelVersion(ByRef issues() As String, ByRef issueCount As Long)
    On Error Resume Next

    ' Check if workbook has a stored version custom property
    Dim storedVersion As String
    storedVersion = ""
    Dim prop As Object
    For Each prop In ThisWorkbook.CustomDocumentProperties
        If prop.Name = "KernelVersion" Then
            storedVersion = CStr(prop.Value)
            Exit For
        End If
    Next prop

    If Len(storedVersion) > 0 Then
        If StrComp(storedVersion, KERNEL_VERSION, vbTextCompare) <> 0 Then
            AddIssue issues, issueCount, "[WARN] HC-05: Kernel version mismatch. " & _
                "Workbook property: " & storedVersion & ", Code: " & KERNEL_VERSION
        End If
    End If

    On Error GoTo 0
End Sub


' =============================================================================
' HC-06: Detail tab staleness
' =============================================================================
Private Sub CheckDetailStaleness(ByRef issues() As String, ByRef issueCount As Long)
    On Error Resume Next

    If KernelFormHelpers.IsResultsStale() Then
        AddIssue issues, issueCount, "[INFO] HC-06: Detail tab results are stale. " & _
            "Inputs have changed since last run."
    End If

    On Error GoTo 0
End Sub


' =============================================================================
' HC-07: ErrorLog overflow
' =============================================================================
Private Sub CheckErrorLogOverflow(ByRef issues() As String, ByRef issueCount As Long)
    On Error Resume Next

    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets(TAB_ERROR_LOG)
    If wsLog Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).row
    If lastRow > HC_ERRORLOG_WARN_ROWS Then
        AddIssue issues, issueCount, "[WARN] HC-07: ErrorLog has " & lastRow & _
            " rows (>" & HC_ERRORLOG_WARN_ROWS & "). Consider purging old entries."
    End If

    On Error GoTo 0
End Sub


' =============================================================================
' HC-08: WAL overflow
' =============================================================================
Private Sub CheckWALOverflow(ByRef issues() As String, ByRef issueCount As Long)
    On Error Resume Next

    Dim walPath As String
    walPath = ThisWorkbook.Path & "\..\wal\wal.log"

    If Dir(walPath) = "" Then Exit Sub

    ' Count lines
    Dim fileNum As Integer
    fileNum = FreeFile
    Dim lineCount As Long
    lineCount = 0
    Dim lineText As String

    Open walPath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineCount = lineCount + 1
    Loop
    Close #fileNum

    If lineCount > HC_WAL_WARN_LINES Then
        AddIssue issues, issueCount, "[WARN] HC-08: WAL has " & lineCount & _
            " lines (>" & HC_WAL_WARN_LINES & "). Consider purging with KernelSnapshot.PurgeWAL."
    End If

    On Error GoTo 0
End Sub


' =============================================================================
' HC-09: Prove-It status
' =============================================================================
Private Sub CheckProveItStatus(ByRef issues() As String, ByRef issueCount As Long)
    On Error Resume Next

    Dim wsPI As Worksheet
    Set wsPI = ThisWorkbook.Sheets(TAB_PROVE_IT)
    If wsPI Is Nothing Then Exit Sub

    ' Only check if tab has content
    If Len(Trim(CStr(wsPI.Cells(5, 1).Value))) = 0 Then Exit Sub

    Dim result As Boolean
    result = KernelProveIt.ValidateProveIt()
    If Not result Then
        AddIssue issues, issueCount, "[WARN] HC-09: Prove-It has FALSE checks. " & _
            "Run 'Generate Prove-It' to investigate."
    End If

    On Error GoTo 0
End Sub


' =============================================================================
' Helper: AddIssue
' =============================================================================
Private Sub AddIssue(ByRef issues() As String, ByRef issueCount As Long, msg As String)
    issueCount = issueCount + 1
    If issueCount > UBound(issues) Then
        ReDim Preserve issues(1 To UBound(issues) * 2)
    End If
    issues(issueCount) = msg
End Sub
