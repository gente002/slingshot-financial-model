Attribute VB_Name = "KernelExtension"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelExtension.bas
' Purpose: Extension activation, deactivation, and execution lifecycle manager.
'          Reads extension_registry.csv from Config sheet. Runs extensions at
'          pipeline hook points (PreCompute, PostCompute, PostOutput).
'          Standalone extensions provide utility functions called directly.
' Phase 6A module. All errors include manual bypass instructions (AP-46).
' =============================================================================

' Extension data arrays (loaded from Config sheet)
Private m_extIDs() As String
Private m_extModules() As String
Private m_extEntries() As String
Private m_extHooks() As String
Private m_extSortOrders() As Long
Private m_extActive() As Boolean
Private m_extMutates() As Boolean
Private m_extSeeds() As Boolean
Private m_extDescs() As String
Private m_extCount As Long
Private m_loaded As Boolean

' Extension outputs handoff (same pattern as KernelTransform.TransformOutputs)
Public ExtensionOutputs As Variant


' =============================================================================
' LoadExtensionRegistry
' Reads extension_registry from Config sheet. Called during LoadAllConfig.
' =============================================================================
Public Sub LoadExtensionRegistry()
    m_extCount = 0
    m_loaded = False

    Dim cnt As Long
    cnt = KernelConfig.GetExtensionCount()
    If cnt = 0 Then
        m_loaded = True
        Exit Sub
    End If

    ReDim m_extIDs(1 To cnt)
    ReDim m_extModules(1 To cnt)
    ReDim m_extEntries(1 To cnt)
    ReDim m_extHooks(1 To cnt)
    ReDim m_extSortOrders(1 To cnt)
    ReDim m_extActive(1 To cnt)
    ReDim m_extMutates(1 To cnt)
    ReDim m_extSeeds(1 To cnt)
    ReDim m_extDescs(1 To cnt)

    Dim idx As Long
    For idx = 1 To cnt
        m_extIDs(idx) = KernelConfig.GetExtensionField(idx, EXTCFG_COL_ID)
        m_extModules(idx) = KernelConfig.GetExtensionField(idx, EXTCFG_COL_MODULE)
        m_extEntries(idx) = KernelConfig.GetExtensionField(idx, EXTCFG_COL_ENTRY)
        m_extHooks(idx) = KernelConfig.GetExtensionField(idx, EXTCFG_COL_HOOK)

        Dim sortVal As String
        sortVal = KernelConfig.GetExtensionField(idx, EXTCFG_COL_SORT)
        If IsNumeric(sortVal) And Len(sortVal) > 0 Then
            m_extSortOrders(idx) = CLng(sortVal)
        Else
            m_extSortOrders(idx) = 999
        End If

        m_extActive(idx) = (StrComp(KernelConfig.GetExtensionField(idx, EXTCFG_COL_ACTIVE), "TRUE", vbTextCompare) = 0)
        m_extMutates(idx) = (StrComp(KernelConfig.GetExtensionField(idx, EXTCFG_COL_MUTATES), "TRUE", vbTextCompare) = 0)
        m_extSeeds(idx) = (StrComp(KernelConfig.GetExtensionField(idx, EXTCFG_COL_SEED), "TRUE", vbTextCompare) = 0)
        m_extDescs(idx) = KernelConfig.GetExtensionField(idx, EXTCFG_COL_DESC)
    Next idx

    m_extCount = cnt
    m_loaded = True

    KernelConfig.LogError SEV_INFO, "KernelExtension", "I-800", _
        "Extension registry loaded: " & cnt & " extension(s) registered", ""
End Sub


' =============================================================================
' RunExtensions
' Called from KernelEngine at the appropriate pipeline stage.
' Filters extensions by hookType and Activated=TRUE.
' Sorts by SortOrder ascending. Two-pass execution:
'   Pass 1: Extensions with MutatesOutputs=FALSE (read-only copy)
'   Pass 2: Extensions with MutatesOutputs=TRUE (writes persist)
' =============================================================================
Public Sub RunExtensions(hookType As String, ByRef outputs() As Variant)
    If m_extCount = 0 Then Exit Sub

    ' Build sorted index of matching active extensions
    Dim matchIdx() As Long
    Dim matchCount As Long
    matchCount = 0

    Dim idx As Long
    For idx = 1 To m_extCount
        If m_extActive(idx) And StrComp(m_extHooks(idx), hookType, vbTextCompare) = 0 Then
            matchCount = matchCount + 1
        End If
    Next idx

    If matchCount = 0 Then Exit Sub

    ReDim matchIdx(1 To matchCount)
    Dim pos As Long
    pos = 0
    For idx = 1 To m_extCount
        If m_extActive(idx) And StrComp(m_extHooks(idx), hookType, vbTextCompare) = 0 Then
            pos = pos + 1
            matchIdx(pos) = idx
        End If
    Next idx

    ' Sort by SortOrder (simple bubble sort)
    Dim i As Long
    Dim j As Long
    Dim tmp As Long
    For i = 1 To matchCount - 1
        For j = 1 To matchCount - i
            If m_extSortOrders(matchIdx(j)) > m_extSortOrders(matchIdx(j + 1)) Then
                tmp = matchIdx(j)
                matchIdx(j) = matchIdx(j + 1)
                matchIdx(j + 1) = tmp
            End If
        Next j
    Next i

    KernelConfig.LogError SEV_INFO, "KernelExtension", "I-801", _
        "Running " & matchCount & " " & hookType & " extension(s)", ""

    ' Pass 1: Read-only extensions (MutatesOutputs=FALSE)
    For i = 1 To matchCount
        idx = matchIdx(i)
        If Not m_extMutates(idx) Then
            RunSingleExtension idx, outputs, False
        End If
    Next i

    ' Pass 2: Mutating extensions (MutatesOutputs=TRUE)
    For i = 1 To matchCount
        idx = matchIdx(i)
        If m_extMutates(idx) Then
            RunSingleExtension idx, outputs, True
        End If
    Next i
End Sub


' =============================================================================
' RunSingleExtension
' Executes a single extension via Application.Run with error handling.
' =============================================================================
Private Sub RunSingleExtension(idx As Long, ByRef outputs() As Variant, mutates As Boolean)
    Dim qualName As String
    qualName = m_extModules(idx) & "." & m_extEntries(idx)

    KernelConfig.LogError SEV_INFO, "KernelExtension", "I-802", _
        "Executing extension: " & m_extIDs(idx) & " (" & qualName & ")", ""

    ' Initialize PRNG if extension requires seed and not yet initialized
    If m_extSeeds(idx) Then
        If Not KernelRandom.IsInitialized() Then
            KernelRandom.AutoSeed
        End If
    End If

    ' Copy outputs to handoff variable
    ExtensionOutputs = outputs

    On Error Resume Next
    Application.Run qualName
    Dim runErr As Long
    runErr = Err.Number
    Dim runDesc As String
    runDesc = Err.Description
    On Error GoTo 0

    If runErr <> 0 Then
        KernelConfig.LogError SEV_ERROR, "KernelExtension", "E-800", _
            "Extension failed: " & m_extIDs(idx) & " -- " & runDesc, _
            "MANUAL BYPASS: Deactivate the failing extension by setting " & _
            "Activated=FALSE for '" & m_extIDs(idx) & "' in extension_registry.csv " & _
            "and re-running Setup.bat."
        ExtensionOutputs = Empty
        Exit Sub
    End If

    ' Copy back if mutating
    If mutates Then
        Dim r As Long
        Dim c As Long
        For r = LBound(outputs, 1) To UBound(outputs, 1)
            For c = LBound(outputs, 2) To UBound(outputs, 2)
                outputs(r, c) = ExtensionOutputs(r, c)
            Next c
        Next r
    End If

    ExtensionOutputs = Empty
End Sub


' =============================================================================
' GetActiveExtensionCount
' Returns count of active extensions, optionally filtered by hookType.
' =============================================================================
Public Function GetActiveExtensionCount(Optional hookType As String = "") As Long
    If Not m_loaded Or m_extCount = 0 Then
        GetActiveExtensionCount = 0
        Exit Function
    End If

    Dim cnt As Long
    cnt = 0
    Dim idx As Long
    For idx = 1 To m_extCount
        If m_extActive(idx) Then
            If Len(hookType) = 0 Then
                cnt = cnt + 1
            ElseIf StrComp(m_extHooks(idx), hookType, vbTextCompare) = 0 Then
                cnt = cnt + 1
            End If
        End If
    Next idx

    GetActiveExtensionCount = cnt
End Function


' =============================================================================
' IsExtensionActive
' Returns TRUE if the named extension is Activated=TRUE.
' =============================================================================
Public Function IsExtensionActive(extensionID As String) As Boolean
    IsExtensionActive = False
    If Not m_loaded Or m_extCount = 0 Then Exit Function

    Dim idx As Long
    For idx = 1 To m_extCount
        If StrComp(m_extIDs(idx), extensionID, vbTextCompare) = 0 Then
            IsExtensionActive = m_extActive(idx)
            Exit Function
        End If
    Next idx
End Function


' =============================================================================
' ListExtensions
' Writes a summary of all extensions (active/inactive) to ErrorLog.
' Called during bootstrap for diagnostic visibility.
' =============================================================================
Public Sub ListExtensions()
    If Not m_loaded Then
        LoadExtensionRegistry
    End If

    If m_extCount = 0 Then
        KernelConfig.LogError SEV_INFO, "KernelExtension", "I-810", _
            "No extensions registered.", ""
        MsgBox "No extensions registered in extension_registry.csv.", _
               vbInformation, "RDK -- Extensions"
        Exit Sub
    End If

    Dim activeCount As Long
    activeCount = 0

    Dim idx As Long
    For idx = 1 To m_extCount
        Dim statusStr As String
        If m_extActive(idx) Then
            statusStr = "ACTIVE"
            activeCount = activeCount + 1
        Else
            statusStr = "INACTIVE"
        End If

        KernelConfig.LogError SEV_INFO, "KernelExtension", "I-811", _
            "[" & statusStr & "] " & m_extIDs(idx) & " (" & m_extModules(idx) & _
            "." & m_extEntries(idx) & ") -- Hook: " & m_extHooks(idx) & _
            ", Sort: " & m_extSortOrders(idx), _
            m_extDescs(idx)
    Next idx

    MsgBox m_extCount & " extension(s) registered, " & activeCount & " active." & vbCrLf & vbCrLf & _
           "See ErrorLog tab for full details.", _
           vbInformation, "RDK -- Extensions"
End Sub
