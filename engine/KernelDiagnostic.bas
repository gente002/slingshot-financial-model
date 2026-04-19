Attribute VB_Name = "KernelDiagnostic"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelDiagnostic.bas
' Purpose: One-paste diagnostic context for AI-assisted debugging.
'          When something goes wrong, DiagnosticDump produces a single text
'          file that an AI can read and diagnose without follow-up questions.
' Phase 4: Observability + Hardening
' =============================================================================

' Recursion guard
Private m_dumpInProgress As Boolean

' =============================================================================
' GenerateDiagnosticDump
' Creates the diagnostic dump file with all 10 sections.
' This function MUST NOT throw errors itself -- each section is wrapped
' in On Error Resume Next so a partial dump is still useful.
' =============================================================================
Public Sub GenerateDiagnosticDump(Optional errContext As String = "")
    If m_dumpInProgress Then Exit Sub
    m_dumpInProgress = True

    Dim dump As String
    Dim sep As String
    sep = String(47, "=")

    dump = sep & vbCrLf
    dump = dump & "RDK DIAGNOSTIC DUMP" & vbCrLf
    dump = dump & "Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    dump = dump & sep & vbCrLf

    ' Section 1: Error Context
    dump = dump & vbCrLf & "--- ERROR CONTEXT ---" & vbCrLf
    dump = dump & BuildErrorContext(errContext)

    ' Section 2: System Info
    dump = dump & vbCrLf & "--- SYSTEM INFO ---" & vbCrLf
    dump = dump & BuildSystemInfo()

    ' Section 3: Kernel State
    dump = dump & vbCrLf & "--- KERNEL STATE ---" & vbCrLf
    dump = dump & BuildKernelState()

    ' Section 4: Pipeline State
    dump = dump & vbCrLf & "--- PIPELINE STATE ---" & vbCrLf
    dump = dump & BuildPipelineState()

    ' Section 5: Config Summary
    dump = dump & vbCrLf & "--- CONFIG SUMMARY ---" & vbCrLf
    dump = dump & BuildConfigSummary()

    ' Section 6: Module Inventory
    dump = dump & vbCrLf & "--- MODULE INVENTORY ---" & vbCrLf
    dump = dump & BuildModuleInventory()

    ' Section 7: Lint Quick
    dump = dump & vbCrLf & "--- LINT (Quick) ---" & vbCrLf
    dump = dump & BuildLintQuick()

    ' Section 8: Recent ErrorLog
    dump = dump & vbCrLf & "--- RECENT ERRORLOG (last " & DIAG_DUMP_MAX_LOG_ENTRIES & ") ---" & vbCrLf
    dump = dump & BuildRecentErrorLog()

    ' Section 9: Recent WAL
    dump = dump & vbCrLf & "--- RECENT WAL (last " & DIAG_DUMP_MAX_WAL_ENTRIES & ") ---" & vbCrLf
    dump = dump & BuildRecentWAL()

    ' Section 10: Snapshot State
    dump = dump & vbCrLf & "--- SNAPSHOT STATE ---" & vbCrLf
    dump = dump & BuildSnapshotState()

    dump = dump & vbCrLf & sep & vbCrLf
    dump = dump & "END OF DIAGNOSTIC DUMP" & vbCrLf
    dump = dump & sep & vbCrLf

    ' Write to file (atomic: temp then rename, PT-002)
    Dim fileName As String
    fileName = "diagnostic_dump_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"

    Dim filePath As String
    filePath = ThisWorkbook.Path & "\..\" & fileName

    On Error Resume Next
    Dim tmpPath As String
    tmpPath = filePath & ".tmp"
    Dim fileNum As Integer
    fileNum = FreeFile
    Open tmpPath For Output As #fileNum
    Print #fileNum, dump
    Close #fileNum

    ' Verify temp file exists, then rename
    If Dir(tmpPath) <> "" Then
        If Dir(filePath) <> "" Then Kill filePath
        Name tmpPath As filePath
    End If
    On Error GoTo 0

    ' Log it
    On Error Resume Next
    KernelConfig.LogError SEV_INFO, "KernelDiagnostic", "I-750", _
        "Diagnostic dump saved: " & fileName, filePath
    On Error GoTo 0

    MsgBox "Diagnostic dump saved: " & fileName & vbCrLf & vbCrLf & _
           "Paste this file into an AI chat for diagnosis.", _
           vbInformation, "RDK -- Diagnostic Dump"

    m_dumpInProgress = False
End Sub


' =============================================================================
' AutoDump
' Called from error handlers when severity = ERROR or FATAL.
' Checks a flag to prevent recursive dumps.
' =============================================================================
Public Sub AutoDump()
    If m_dumpInProgress Then Exit Sub
    On Error Resume Next
    GenerateDiagnosticDump "Auto-triggered from error handler"
    On Error GoTo 0
End Sub


' =============================================================================
' Section Builders -- each wrapped in On Error Resume Next
' =============================================================================

Private Function BuildErrorContext(errContext As String) As String
    On Error Resume Next
    Dim s As String
    If Len(errContext) > 0 Then
        s = errContext & vbCrLf
    Else
        s = "No error context provided (manual dump)." & vbCrLf
    End If
    If Err.Number <> 0 Then
        s = s & "Error: " & Err.Number & " - " & Err.Description & vbCrLf
        s = s & "Source: " & Err.Source & vbCrLf
    End If
    On Error GoTo 0
    BuildErrorContext = s
End Function


Private Function BuildSystemInfo() As String
    On Error Resume Next
    Dim s As String
    s = "Excel: " & Application.Version
    #If Win64 Then
        s = s & " (64-bit)" & vbCrLf
    #Else
        s = s & " (32-bit)" & vbCrLf
    #End If
    s = s & "OS: " & Environ("OS") & vbCrLf
    s = s & "Machine: " & Environ("COMPUTERNAME") & vbCrLf
    s = s & "User: " & Environ("USERNAME") & vbCrLf
    s = s & "Calculation: "
    Select Case Application.Calculation
        Case xlCalculationAutomatic: s = s & "xlCalculationAutomatic"
        Case xlCalculationManual: s = s & "xlCalculationManual"
        Case Else: s = s & CStr(Application.Calculation)
    End Select
    s = s & vbCrLf
    s = s & "Separators: " & Application.DecimalSeparator & " (decimal), " & _
        Application.ThousandsSeparator & " (thousands)" & vbCrLf
    On Error GoTo 0
    BuildSystemInfo = s
End Function


Private Function BuildKernelState() As String
    On Error Resume Next
    Dim s As String
    s = "Version: " & KERNEL_VERSION & vbCrLf

    ' Check if config is loaded
    Dim colCount As Long
    colCount = KernelConfig.GetColumnCount()
    If colCount > 0 Then
        s = s & "Config Loaded: Yes" & vbCrLf
    Else
        s = s & "Config Loaded: No (or empty)" & vbCrLf
    End If

    ' Entity count from Inputs tab
    Dim entityCount As Long
    entityCount = 0
    Dim wsInputs As Worksheet
    Set wsInputs = ThisWorkbook.Sheets(TAB_INPUTS)
    If Not wsInputs Is Nothing Then
        Dim col As Long
        col = INPUT_ENTITY_START_COL
        Do While col < INPUT_ENTITY_START_COL + INPUT_MAX_ENTITIES
            If Len(Trim(CStr(wsInputs.Cells(3, col).Value))) = 0 Then Exit Do
            entityCount = entityCount + 1
            col = col + 1
        Loop
    End If
    s = s & "Entities: " & entityCount & vbCrLf

    ' Period count
    Dim periodCount As Long
    periodCount = KernelConfig.GetTimeHorizon()
    s = s & "Periods: " & periodCount & vbCrLf
    s = s & "Columns: " & colCount & vbCrLf
    On Error GoTo 0
    BuildKernelState = s
End Function


Private Function BuildPipelineState() As String
    On Error Resume Next
    Dim s As String
    s = ""
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    If wsConfig Is Nothing Then
        s = "(Config sheet not found)" & vbCrLf
        BuildPipelineState = s
        Exit Function
    End If

    ' Find PIPELINE_STATE marker
    Dim markerRow As Long
    markerRow = 0
    Dim r As Long
    For r = 1 To wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).row
        If StrComp(CStr(wsConfig.Cells(r, 1).Value), PIPELINE_STATE_MARKER, vbTextCompare) = 0 Then
            markerRow = r
            Exit For
        End If
    Next r

    If markerRow = 0 Then
        s = "(No pipeline state recorded)" & vbCrLf
        BuildPipelineState = s
        Exit Function
    End If

    Dim stepNames As Variant
    stepNames = Array("Bootstrap", "LoadConfig", "Validate", "Compute", "WriteDetail", "WriteCSV", "WriteSummary")

    Dim stepIdx As Long
    For stepIdx = 0 To 6
        Dim dataRow As Long
        dataRow = markerRow + 1 + stepIdx
        Dim stepName As String
        stepName = CStr(wsConfig.Cells(dataRow, 1).Value)
        Dim stepStatus As String
        stepStatus = CStr(wsConfig.Cells(dataRow, 2).Value)
        Dim stepTime As String
        stepTime = CStr(wsConfig.Cells(dataRow, 3).Value)
        If Len(stepName) > 0 Then
            s = s & "Step " & stepIdx & " (" & stepNames(stepIdx) & "): " & _
                stepStatus & " at " & stepTime & vbCrLf
        End If
    Next stepIdx

    On Error GoTo 0
    BuildPipelineState = s
End Function


Private Function BuildConfigSummary() As String
    On Error Resume Next
    Dim s As String
    s = ""

    ' Column summary
    Dim colCount As Long
    colCount = KernelConfig.GetColumnCount()
    If colCount > 0 Then
        s = s & "Columns: "
        Dim cidx As Long
        For cidx = 1 To colCount
            If cidx > 1 Then s = s & ", "
            Dim colName As String
            colName = KernelConfig.GetColName(cidx)
            s = s & colName & "(" & _
                Left(KernelConfig.GetFieldClass(colName), 3) & ")"
        Next cidx
        s = s & vbCrLf
    Else
        s = s & "Columns: (not loaded)" & vbCrLf
    End If

    ' Input summary
    Dim inputCount As Long
    inputCount = KernelConfig.GetInputCount()
    If inputCount > 0 Then
        s = s & "Inputs: "
        Dim lastSect As String
        lastSect = ""
        Dim pidx As Long
        For pidx = 1 To inputCount
            Dim sect As String
            sect = KernelConfig.GetInputSection(pidx)
            If StrComp(sect, lastSect, vbTextCompare) <> 0 Then
                If pidx > 1 Then s = s & " | "
                s = s & "[" & sect & "] "
                lastSect = sect
            ElseIf pidx > 1 Then
                s = s & ", "
            End If
            s = s & KernelConfig.GetInputParam(pidx)
        Next pidx
        s = s & vbCrLf
    Else
        s = s & "Inputs: (not loaded)" & vbCrLf
    End If

    ' Granularity
    Dim th As Long
    th = KernelConfig.GetTimeHorizon()
    s = s & "Granularity: TimeHorizon=" & th & vbCrLf

    On Error GoTo 0
    BuildConfigSummary = s
End Function


Private Function BuildModuleInventory() As String
    On Error Resume Next
    Dim s As String
    s = ""

    Dim enginePath As String
    enginePath = ThisWorkbook.Path & "\..\engine"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(enginePath) Then
        s = "(engine/ directory not found)" & vbCrLf
        Set fso = Nothing
        BuildModuleInventory = s
        Exit Function
    End If

    Dim folder As Object
    Set folder = fso.GetFolder(enginePath)
    Dim fileObj As Object

    For Each fileObj In folder.Files
        If LCase(fso.GetExtensionName(fileObj.Name)) = "bas" Then
            Dim sizeStr As String
            Dim sizeKB As Double
            sizeKB = fileObj.Size / 1024
            sizeStr = Format(sizeKB, "0.0") & " KB"

            ' Check CRLF
            Dim crlfStatus As String
            crlfStatus = "CRLF=OK"
            Dim fc As String
            Dim ts As Object
            Set ts = fso.OpenTextFile(fileObj.Path, 1, False)
            fc = ts.ReadAll
            ts.Close
            If InStr(1, fc, vbCrLf) = 0 And InStr(1, fc, vbLf) > 0 Then
                crlfStatus = "CRLF=FAIL"
            End If

            Dim warnTag As String
            warnTag = ""
            If fileObj.Size > MODULE_SIZE_ERROR Then
                warnTag = "  [ERROR: >64KB]"
            ElseIf fileObj.Size > MODULE_SIZE_WARN Then
                warnTag = "  [WARN: >50KB]"
            End If

            ' Pad filename for alignment
            Dim paddedName As String
            paddedName = fileObj.Name & String(30 - Len(fileObj.Name), " ")
            If Len(fileObj.Name) >= 30 Then paddedName = fileObj.Name & " "

            s = s & paddedName & sizeStr & "  " & crlfStatus & warnTag & vbCrLf
        End If
    Next fileObj

    Set folder = Nothing
    Set fso = Nothing
    On Error GoTo 0
    BuildModuleInventory = s
End Function


Private Function BuildLintQuick() As String
    On Error Resume Next
    Dim s As String
    ' Run lint quick silently (no MsgBox or ErrorLog writes)
    KernelLint.RunLintQuickSilent
    s = KernelLint.GetQuickViolationSummary()
    On Error GoTo 0
    BuildLintQuick = s
End Function


Private Function BuildRecentErrorLog() As String
    On Error Resume Next
    Dim s As String
    s = ""

    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets(TAB_ERROR_LOG)
    If wsLog Is Nothing Then
        s = "(ErrorLog sheet not found)" & vbCrLf
        BuildRecentErrorLog = s
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).row
    If lastRow < 2 Then
        s = "(No entries)" & vbCrLf
        BuildRecentErrorLog = s
        Exit Function
    End If

    Dim startRow As Long
    startRow = lastRow - DIAG_DUMP_MAX_LOG_ENTRIES + 1
    If startRow < 2 Then startRow = 2

    Dim r As Long
    For r = startRow To lastRow
        Dim ts As String
        ts = CStr(wsLog.Cells(r, 1).Value)
        Dim sev As String
        sev = CStr(wsLog.Cells(r, 2).Value)
        Dim src As String
        src = CStr(wsLog.Cells(r, 3).Value)
        Dim errCode As String
        errCode = CStr(wsLog.Cells(r, 4).Value)
        Dim errMsg As String
        errMsg = CStr(wsLog.Cells(r, 5).Value)
        s = s & ts & "  " & sev & "  " & src & "  " & errCode & "  " & errMsg & vbCrLf
    Next r

    On Error GoTo 0
    BuildRecentErrorLog = s
End Function


Private Function BuildRecentWAL() As String
    On Error Resume Next
    Dim s As String
    s = ""

    Dim walPath As String
    walPath = ThisWorkbook.Path & "\..\wal\wal.log"

    If Dir(walPath) = "" Then
        s = "(No WAL file)" & vbCrLf
        BuildRecentWAL = s
        Exit Function
    End If

    ' Read WAL file
    Dim fileNum As Integer
    fileNum = FreeFile
    Dim fileContent As String
    Dim fSize As Long

    Open walPath For Binary Access Read As #fileNum
    fSize = LOF(fileNum)
    If fSize = 0 Then
        Close #fileNum
        s = "(WAL empty)" & vbCrLf
        BuildRecentWAL = s
        Exit Function
    End If
    fileContent = Space$(fSize)
    Get #fileNum, , fileContent
    Close #fileNum

    ' Normalize line endings and get last N lines
    fileContent = Replace(fileContent, vbCrLf, vbLf)
    Dim allLines() As String
    allLines = Split(fileContent, vbLf)

    Dim totalLines As Long
    totalLines = UBound(allLines) + 1
    Dim startIdx As Long
    startIdx = totalLines - DIAG_DUMP_MAX_WAL_ENTRIES
    If startIdx < 0 Then startIdx = 0

    Dim idx As Long
    For idx = startIdx To UBound(allLines)
        If Len(Trim(allLines(idx))) > 0 Then
            s = s & allLines(idx) & vbCrLf
        End If
    Next idx

    On Error GoTo 0
    BuildRecentWAL = s
End Function


Private Function BuildSnapshotState() As String
    On Error Resume Next
    Dim s As String
    s = ""

    Dim snapshots() As String
    snapshots = KernelSnapshot.ListSnapshots()

    ' Check if array is valid
    If Not IsArrayValid(snapshots) Then
        s = "(No snapshots)" & vbCrLf
        BuildSnapshotState = s
        Exit Function
    End If

    Dim i As Long
    For i = LBound(snapshots) To UBound(snapshots)
        If Len(snapshots(i)) > 0 Then
            Dim snapDir As String
            snapDir = ThisWorkbook.Path & "\..\snapshots\" & snapshots(i)
            Dim manifestPath As String
            manifestPath = snapDir & "\manifest.json"

            Dim snapStatus As String
            snapStatus = "UNKNOWN"
            If Dir(manifestPath) <> "" Then
                snapStatus = KernelFormHelpers.ReadJsonField(manifestPath, "status")
                If Len(snapStatus) = 0 Then snapStatus = "COMPLETE"
            Else
                snapStatus = "INCOMPLETE (no manifest)"
            End If

            Dim snapTime As String
            snapTime = ""
            If Dir(manifestPath) <> "" Then
                snapTime = KernelFormHelpers.ReadJsonField(manifestPath, "timestamp")
            End If

            s = s & snapshots(i) & ": " & snapStatus
            If Len(snapTime) > 0 Then s = s & " (" & snapTime & ")"
            s = s & vbCrLf
        End If
    Next i

    On Error GoTo 0
    BuildSnapshotState = s
End Function


' =============================================================================
' IsArrayValid - checks if a string array has been properly initialized
' =============================================================================
Private Function IsArrayValid(arr() As String) As Boolean
    On Error GoTo NotValid
    Dim test As Long
    test = UBound(arr)
    ' Check it is not the empty Split result
    If test = 0 And Len(arr(0)) = 0 Then
        IsArrayValid = False
    Else
        IsArrayValid = True
    End If
    Exit Function
NotValid:
    IsArrayValid = False
End Function
