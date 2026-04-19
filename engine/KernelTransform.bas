Attribute VB_Name = "KernelTransform"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelTransform.bas
' Purpose: Post-computation transform framework + Power Pivot data model +
'          pivot tables. Phase 5B module.
' Transforms run between DomainExecute and WriteDetail (AP-47 compliant:
' they modify the output array, never the Detail tab directly).
' Power Pivot degrades gracefully if unavailable.
' =============================================================================

' Transform registry entry
Private Type TransformEntry
    transformName As String
    moduleName As String
    functionName As String
    sortOrder As Long
    mutatesOutputs As Boolean
    enabled As Boolean
End Type

Private m_transforms() As TransformEntry
Private m_transformCount As Long
Private m_transformsInitialized As Boolean

' Public handoff for transform functions.
' Application.Run cannot pass arrays, so transforms read/write this directly.
Public TransformOutputs As Variant


' =============================================================================
' RegisterTransform
' Registers a transform in the module-level collection.
' =============================================================================
Public Sub RegisterTransform(transformName As String, moduleName As String, _
                             functionName As String, sortOrder As Long, _
                             Optional mutatesOutputs As Boolean = True)
    If Not m_transformsInitialized Then
        m_transformCount = 0
        m_transformsInitialized = True
    End If

    ' Check for duplicate name
    Dim i As Long
    For i = 1 To m_transformCount
        If StrComp(m_transforms(i).transformName, transformName, vbTextCompare) = 0 Then
            KernelConfig.LogError SEV_WARN, "KernelTransform", "W-700", _
                "Transform already registered: " & transformName, _
                "Overwriting existing registration."
            m_transforms(i).moduleName = moduleName
            m_transforms(i).functionName = functionName
            m_transforms(i).sortOrder = sortOrder
            m_transforms(i).mutatesOutputs = mutatesOutputs
            m_transforms(i).enabled = True
            Exit Sub
        End If
    Next i

    m_transformCount = m_transformCount + 1
    ReDim Preserve m_transforms(1 To m_transformCount)
    With m_transforms(m_transformCount)
        .transformName = transformName
        .moduleName = moduleName
        .functionName = functionName
        .sortOrder = sortOrder
        .mutatesOutputs = mutatesOutputs
        .enabled = True
    End With

    KernelConfig.LogError SEV_INFO, "KernelTransform", "I-700", _
        "Transform registered: " & transformName & " (order=" & sortOrder & ")", _
        moduleName & "." & functionName
End Sub


' =============================================================================
' RunTransforms
' Executes all registered transforms in sortOrder.
' Called from KernelEngine after DomainExecute, before WriteDetail.
' MANUAL BYPASS: Skip transforms by calling WriteDetail directly after DomainExecute.
' =============================================================================
Public Sub RunTransforms(ByRef outputs() As Variant)
    If m_transformCount = 0 Then Exit Sub

    ' Sort transforms by sortOrder (simple bubble sort)
    SortTransforms

    KernelConfig.LogError SEV_INFO, "KernelTransform", "I-701", _
        "Running " & m_transformCount & " registered transform(s)", ""

    ' Copy outputs to module-level handoff.
    ' Application.Run cannot pass arrays -- transforms access
    ' KernelTransform.TransformOutputs directly instead.
    TransformOutputs = outputs

    Dim idx As Long
    For idx = 1 To m_transformCount
        If Not m_transforms(idx).enabled Then GoTo NextTransform

        Dim qualName As String
        qualName = m_transforms(idx).moduleName & "." & m_transforms(idx).functionName

        KernelConfig.LogError SEV_INFO, "KernelTransform", "I-702", _
            "Executing transform: " & m_transforms(idx).transformName, qualName

        On Error Resume Next
        Application.Run qualName
        If Err.Number <> 0 Then
            KernelConfig.LogError SEV_ERROR, "KernelTransform", "E-700", _
                "Transform failed: " & m_transforms(idx).transformName & " -- " & Err.Description, _
                "MANUAL BYPASS: Skip transforms by calling WriteDetail directly. " & _
                "Disable this transform or fix the function: " & qualName
            Err.Clear
        End If
        On Error GoTo 0
NextTransform:
    Next idx

    ' Copy handoff back to caller's array
    Dim r As Long
    Dim c As Long
    For r = LBound(outputs, 1) To UBound(outputs, 1)
        For c = LBound(outputs, 2) To UBound(outputs, 2)
            outputs(r, c) = TransformOutputs(r, c)
        Next c
    Next r

    ' Clear handoff
    TransformOutputs = Empty
End Sub


' =============================================================================
' ClearTransforms
' Removes all registered transforms.
' =============================================================================
Public Sub ClearTransforms()
    m_transformCount = 0
    m_transformsInitialized = True
    Erase m_transforms
End Sub


' =============================================================================
' GetTransformCount
' =============================================================================
Public Function GetTransformCount() As Long
    If m_transformsInitialized Then
        GetTransformCount = m_transformCount
    Else
        GetTransformCount = 0
    End If
End Function


' =============================================================================
' SortTransforms (Private)
' Simple bubble sort by sortOrder ascending.
' =============================================================================
Private Sub SortTransforms()
    If m_transformCount <= 1 Then Exit Sub
    Dim i As Long
    Dim j As Long
    For i = 1 To m_transformCount - 1
        For j = 1 To m_transformCount - i
            If m_transforms(j).sortOrder > m_transforms(j + 1).sortOrder Then
                Dim tmp As TransformEntry
                tmp = m_transforms(j)
                m_transforms(j) = m_transforms(j + 1)
                m_transforms(j + 1) = tmp
            End If
        Next j
    Next i
End Sub


' =============================================================================
' CreateDataModel
' Creates or refreshes Power Pivot data model from Detail + Inputs.
' Gracefully degrades if Power Pivot is unavailable.
' =============================================================================
Public Sub CreateDataModel()
    On Error GoTo ErrHandler

    ' Check if Power Pivot is enabled
    Dim ppSetting As String
    ppSetting = KernelConfig.GetDataModelSetting("PowerPivotEnabled")
    If StrComp(ppSetting, "FALSE", vbTextCompare) = 0 Then
        KernelConfig.LogError SEV_INFO, "KernelTransform", "I-710", _
            "Power Pivot disabled in config. Skipping data model creation.", ""
        Exit Sub
    End If

    ' Check if Power Pivot / data model is available
    Dim pm As Object
    On Error Resume Next
    Set pm = ThisWorkbook.Model
    On Error GoTo ErrHandler

    If pm Is Nothing Then
        KernelConfig.LogError SEV_INFO, "KernelTransform", "I-711", _
            "Power Pivot not available. Skipping data model creation.", _
            "MANUAL BYPASS: Create a Power Pivot data model manually. " & _
            "Add Detail as a table. Add relationships via the diagram view."
        Exit Sub
    End If

    ' Verify Detail tab has data
    Dim wsDetail As Worksheet
    Set wsDetail = ThisWorkbook.Sheets(TAB_DETAIL)
    Dim lastRow As Long
    lastRow = wsDetail.Cells(wsDetail.Rows.Count, 1).End(xlUp).row
    If lastRow < DETAIL_DATA_START_ROW Then
        KernelConfig.LogError SEV_WARN, "KernelTransform", "W-710", _
            "Detail tab has no data. Cannot create data model.", _
            "MANUAL BYPASS: Run projections first, then create data model."
        Exit Sub
    End If

    ' Determine data range for Detail
    Dim totalCols As Long
    totalCols = KernelConfig.GetColumnCount()
    If totalCols = 0 Then totalCols = wsDetail.Cells(DETAIL_HEADER_ROW, wsDetail.Columns.Count).End(xlToLeft).Column

    Dim detailRange As Range
    Set detailRange = wsDetail.Range(wsDetail.Cells(DETAIL_HEADER_ROW, 1), _
                                     wsDetail.Cells(lastRow, totalCols))

    ' Create or refresh FactDetail table
    Dim tblCount As Long
    tblCount = 0

    On Error Resume Next
    ' Try to add as a model table
    Dim mdlTbl As Object
    Set mdlTbl = pm.ModelTables.Add(detailRange)
    If Err.Number <> 0 Then
        Err.Clear
        ' Table may already exist -- try to refresh
        pm.Refresh
        KernelConfig.LogError SEV_INFO, "KernelTransform", "I-712", _
            "Data model refreshed (tables may already exist).", ""
    Else
        tblCount = tblCount + 1
    End If
    On Error GoTo ErrHandler

    KernelConfig.LogError SEV_INFO, "KernelTransform", "I-713", _
        "Data model created/refreshed.", _
        tblCount & " table(s) added."

    MsgBox "Data model created/refreshed." & vbCrLf & _
           tblCount & " table(s) processed.", _
           vbInformation, "RDK -- Data Model"
    Exit Sub

ErrHandler:
    KernelConfig.LogError SEV_ERROR, "KernelTransform", "E-710", _
        "Data model creation error: " & Err.Description, _
        "MANUAL BYPASS: Create a Power Pivot data model manually. " & _
        "Add Detail as a table. Add relationships via the diagram view."
    MsgBox "Data model error: " & Err.Description & vbCrLf & vbCrLf & _
           "MANUAL BYPASS: Create a Power Pivot data model manually." & vbCrLf & _
           "Add Detail as a table. Add relationships via the diagram view.", _
           vbExclamation, "RDK -- Data Model Error"
End Sub


' =============================================================================
' CreatePivotTables
' Reads pivot_config and creates PivotTables on the Analysis tab.
' Falls back to Detail range if no data model available.
' =============================================================================
Public Sub CreatePivotTables()
    On Error GoTo ErrHandler

    Dim cnt As Long
    cnt = KernelConfig.GetPivotConfigCount()
    If cnt = 0 Then
        KernelConfig.LogError SEV_INFO, "KernelTransform", "I-720", _
            "No pivot config entries found. Skipping pivot table creation.", ""
        MsgBox "No pivot configuration found." & vbCrLf & vbCrLf & _
               "MANUAL BYPASS: Insert -> PivotTable. Select Detail range. Drag fields.", _
               vbInformation, "RDK -- Pivot Tables"
        Exit Sub
    End If

    ' Ensure Analysis tab exists
    ' DEPRECATED: TAB_ANALYSIS removed from tab_registry; no-op if tab absent
    Dim wsAnalysis As Worksheet
    Set wsAnalysis = Nothing
    On Error Resume Next
    Set wsAnalysis = ThisWorkbook.Sheets(TAB_ANALYSIS)
    On Error GoTo ErrHandler
    If wsAnalysis Is Nothing Then Exit Sub

    ' Clear existing content
    wsAnalysis.Cells.ClearContents
    ' Delete existing pivot tables
    Dim pt As PivotTable
    For Each pt In wsAnalysis.PivotTables
        pt.TableRange2.Clear
    Next pt

    ' Get Detail data range
    Dim wsDetail As Worksheet
    Set wsDetail = ThisWorkbook.Sheets(TAB_DETAIL)
    Dim lastRow As Long
    lastRow = wsDetail.Cells(wsDetail.Rows.Count, 1).End(xlUp).row
    If lastRow < DETAIL_DATA_START_ROW Then
        KernelConfig.LogError SEV_WARN, "KernelTransform", "W-720", _
            "Detail tab has no data. Cannot create pivot tables.", _
            "MANUAL BYPASS: Run projections first, then create pivot tables."
        MsgBox "Detail tab has no data. Run projections first.", _
               vbExclamation, "RDK -- Pivot Tables"
        Exit Sub
    End If

    Dim totalCols As Long
    totalCols = KernelConfig.GetColumnCount()
    If totalCols = 0 Then totalCols = wsDetail.Cells(DETAIL_HEADER_ROW, wsDetail.Columns.Count).End(xlToLeft).Column

    Dim sourceRange As Range
    Set sourceRange = wsDetail.Range(wsDetail.Cells(DETAIL_HEADER_ROW, 1), _
                                     wsDetail.Cells(lastRow, totalCols))

    Dim created As Long
    created = 0
    Dim topRow As Long
    topRow = 2

    Dim idx As Long
    For idx = 1 To cnt
        Dim enVal As String
        enVal = KernelConfig.GetPivotConfigField(idx, PVTCFG_COL_ENABLED)
        If StrComp(enVal, "TRUE", vbTextCompare) <> 0 Then GoTo NextPivot

        Dim pvtName As String
        pvtName = KernelConfig.GetPivotConfigField(idx, PVTCFG_COL_NAME)
        Dim pvtId As String
        pvtId = KernelConfig.GetPivotConfigField(idx, PVTCFG_COL_ID)

        ' Label
        wsAnalysis.Cells(topRow, 1).Value = pvtName
        wsAnalysis.Cells(topRow, 1).Font.Bold = True
        topRow = topRow + 1

        ' Create PivotCache from Detail range
        Dim pvtCache As PivotCache
        Set pvtCache = ThisWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=sourceRange)

        ' Create PivotTable
        Dim pvtTable As PivotTable
        Set pvtTable = pvtCache.CreatePivotTable( _
            TableDestination:=wsAnalysis.Cells(topRow, 1), _
            TableName:=pvtId)

        ' Configure fields
        Dim rowField As String
        rowField = KernelConfig.GetPivotConfigField(idx, PVTCFG_COL_ROWFIELD)
        Dim colField As String
        colField = KernelConfig.GetPivotConfigField(idx, PVTCFG_COL_COLFIELD)
        Dim valField As String
        valField = KernelConfig.GetPivotConfigField(idx, PVTCFG_COL_VALUEFIELD)
        Dim aggFunc As String
        aggFunc = KernelConfig.GetPivotConfigField(idx, PVTCFG_COL_AGGFUNC)

        On Error Resume Next
        ' Add row field
        If Len(rowField) > 0 Then
            pvtTable.PivotFields(rowField).Orientation = xlRowField
        End If

        ' Add column field
        If Len(colField) > 0 Then
            pvtTable.PivotFields(colField).Orientation = xlColumnField
        End If

        ' Add value field with aggregation
        If Len(valField) > 0 Then
            Dim dataFld As PivotField
            Set dataFld = pvtTable.PivotFields(valField)
            dataFld.Orientation = xlDataField
            Select Case UCase(aggFunc)
                Case "SUM": dataFld.Function = xlSum
                Case "AVERAGE": dataFld.Function = xlAverage
                Case "COUNT": dataFld.Function = xlCount
                Case Else: dataFld.Function = xlSum
            End Select
        End If

        If Err.Number <> 0 Then
            KernelConfig.LogError SEV_ERROR, "KernelTransform", "E-720", _
                "Error configuring pivot " & pvtId & ": " & Err.Description, _
                "MANUAL BYPASS: Insert -> PivotTable. Select Detail range. Drag fields."
            Err.Clear
        End If
        On Error GoTo ErrHandler

        ' Move topRow past the pivot table area
        topRow = wsAnalysis.Cells(wsAnalysis.Rows.Count, 1).End(xlUp).row + 3
        created = created + 1
NextPivot:
    Next idx

    MsgBox "Created " & created & " pivot table(s) on the Analysis tab.", _
           vbInformation, "RDK -- Pivot Tables"
    Exit Sub

ErrHandler:
    KernelConfig.LogError SEV_ERROR, "KernelTransform", "E-721", _
        "Pivot table creation error: " & Err.Description, _
        "MANUAL BYPASS: Insert -> PivotTable. Select Detail range. Drag fields."
    MsgBox "Pivot table error: " & Err.Description & vbCrLf & vbCrLf & _
           "MANUAL BYPASS: Insert -> PivotTable. Select Detail range. Drag fields.", _
           vbExclamation, "RDK -- Pivot Table Error"
End Sub


' =============================================================================
' RefreshDataModel
' Refreshes the Power Pivot data model after new data.
' Called from KernelEngine after WriteDetail if RefreshOnRun=TRUE.
' =============================================================================
Public Sub RefreshDataModel()
    On Error Resume Next

    Dim refreshSetting As String
    refreshSetting = KernelConfig.GetDataModelSetting("RefreshOnRun")
    If StrComp(refreshSetting, "TRUE", vbTextCompare) <> 0 Then Exit Sub

    Dim ppSetting As String
    ppSetting = KernelConfig.GetDataModelSetting("PowerPivotEnabled")
    If StrComp(ppSetting, "FALSE", vbTextCompare) = 0 Then Exit Sub

    Dim pm As Object
    Set pm = ThisWorkbook.Model
    If pm Is Nothing Then Exit Sub

    pm.Refresh

    KernelConfig.LogError SEV_INFO, "KernelTransform", "I-730", _
        "Data model refreshed after run.", ""

    On Error GoTo 0
End Sub
