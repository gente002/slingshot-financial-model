Attribute VB_Name = "KernelConfig"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelConfig.bas
' Purpose: Reads config CSVs from the Config sheet. Provides runtime lookup
'          functions that all other modules use.
' Supports fallback mode (AP-45): if arrays fail to load, reads directly
' from the Config sheet via KernelConfigLoader fallback helpers.
' Phase 5B: Loaders split to KernelConfigLoader.bas. Data arrays are Public
'           so KernelConfigLoader can write to them (AP-13 exception for split).
' =============================================================================

' Column registry (Public for KernelConfigLoader)
Public m_colNames() As String
Public m_detailCols() As Long
Public m_csvCols() As Long
Public m_blocks() As String
Public m_fieldClasses() As String
Public m_defaultViews() As String
Public m_formats() As String
Public m_balGrps() As String
Public m_derivRules() As String
Public m_balanceTypes() As String
Public m_colCount As Long
Public m_colDict As Object
Public m_colDictLoaded As Boolean

' Input schema
Public m_inputSections() As String
Public m_inputParams() As String
Public m_inputRows() As Long
Public m_inputTypes() As String
Public m_inputDefaults() As String
Public m_inputCount As Long

' Granularity config
Public m_timeHorizon As Long
Public m_maxEntities As Long
Public m_defaultView As String
Public m_settingKeys() As String
Public m_settingValues() As String
Public m_settingCount As Long

' ReproConfig
Public m_reproKeys() As String
Public m_reproValues() As String
Public m_reproCount As Long

' ScaleLimits
Public m_scaleKeys() As String
Public m_scaleValues() As String
Public m_scaleCount As Long

' ProveItConfig
Public m_piCheckIDs() As String
Public m_piCheckTypes() As String
Public m_piCheckNames() As String
Public m_piMetricA() As String
Public m_piMetricB() As String
Public m_piMetricC() As String
Public m_piOperators() As String
Public m_piTolerances() As Double
Public m_piEnabled() As Boolean
Public m_piCount As Long

' Phase 5A: Summary config (2D)
Public m_sumCfg() As String
Public m_sumCfgCount As Long

' Phase 5A: Chart registry (2D)
Public m_chtCfg() As String
Public m_chtCfgCount As Long

' Phase 5A: Exhibit config (2D)
Public m_exhCfg() As String
Public m_exhCfgCount As Long

' Phase 5A: Display mode config (key-value)
Public m_dispKeys() As String
Public m_dispValues() As String
Public m_dispCount As Long

' Phase 5B: Print config (2D, 11 cols)
Public m_prtCfg() As String
Public m_prtCfgCount As Long

' Phase 5B: Data model config (key-value)
Public m_dmKeys() As String
Public m_dmValues() As String
Public m_dmCount As Long

' Phase 5B: Pivot config (2D, 8 cols)
Public m_pvtCfg() As String
Public m_pvtCfgCount As Long

' Phase 5C: Formula tab config (2D, 15 cols)
Public m_ftCfg() As String
Public m_ftCfgCount As Long

' Phase 5C: Named range registry (2D, 6 cols)
Public m_nrCfg() As String
Public m_nrCfgCount As Long

' Phase 6A: Extension registry (2D, 9 cols)
Public m_extCfg() As String
Public m_extCfgCount As Long

' Phase 6A: Curve library config (2D, 10 cols)
Public m_clCfg() As String
Public m_clCfgCount As Long

' Phase 6A: Report config (key-value)
Public m_rptKeys() As String
Public m_rptValues() As String
Public m_rptCount As Long

' Operational state (Private - not shared)
Private m_fallbackMode As Boolean
Private m_fallbackLogged As Boolean
Private m_runFatalCount As Long
Private m_runErrorCount As Long


' --- LoadAllConfig ---
Public Sub LoadAllConfig()
    On Error GoTo LoadFailed
    m_fallbackMode = False
    m_fallbackLogged = False
    KernelConfigLoader.LoadColumnRegistry
    KernelConfigLoader.LoadInputSchema
    KernelConfigLoader.LoadGranularityConfig
    KernelConfigLoader.LoadReproConfig
    KernelConfigLoader.LoadScaleLimits
    KernelConfigLoader.LoadProveItConfig
    KernelConfigLoader.LoadSummaryConfig
    KernelConfigLoader.LoadChartRegistry
    KernelConfigLoader.LoadExhibitConfig
    KernelConfigLoader.LoadDisplayModeConfig
    KernelConfigLoader.LoadPrintConfig
    KernelConfigLoader.LoadDataModelConfig
    KernelConfigLoader.LoadPivotConfig
    KernelConfigLoader.LoadFormulaTabConfig
    KernelConfigLoader.LoadNamedRangeRegistry
    KernelConfigLoader.LoadExtensionRegistry
    KernelConfigLoader.LoadCurveLibraryConfig
    KernelConfigLoader.LoadReportConfig
    ' Validate cross-references between config tables
    KernelConfigLoader.ValidateConfigReferences
    Exit Sub
LoadFailed:
    m_fallbackMode = True
    LogError SEV_WARN, "KernelConfig", "W-100", _
        "Config arrays failed to load. Operating in fallback mode (reading from Config sheet).", _
        Err.Description
End Sub


' --- ColIndex ---
Public Function ColIndex(metricName As String) As Long
    If Not m_fallbackMode Then
        If Not m_colDictLoaded Then
            ColIndex = -1
            LogError SEV_FATAL, "KernelConfig", "E-100", _
                "ColIndex called before LoadAllConfig", metricName
            Exit Function
        End If
        If m_colDict.Exists(metricName) Then
            ColIndex = m_detailCols(m_colDict(metricName))
        Else
            ColIndex = -1
            LogError SEV_FATAL, "KernelConfig", "E-101", _
                "Column not found in registry: " & metricName, _
                "MANUAL BYPASS: Check Config sheet '=== COLUMN_REGISTRY ===' section for available column names."
        End If
        Exit Function
    End If
    If Not m_fallbackLogged Then
        LogError SEV_WARN, "KernelConfig", "W-101", _
            "ColIndex operating in fallback mode (sheet read)", metricName
        m_fallbackLogged = True
    End If
    Dim val As Variant
    val = KernelConfigLoader.FallbackColLookup(metricName, CREG_COL_DETAIL)
    If IsEmpty(val) Then
        ColIndex = -1
        LogError SEV_FATAL, "KernelConfig", "E-101", _
            "Column not found: " & metricName, _
            "MANUAL BYPASS: Check Config sheet COLUMN_REGISTRY for available names."
    Else
        ColIndex = CLng(val)
    End If
End Function


' --- TryColIndex ---
' Non-logging probe: returns DetailCol if column exists, -1 otherwise.
' Use for column name fallback patterns (e.g., Period vs CalPeriod).
Public Function TryColIndex(metricName As String) As Long
    If Not m_fallbackMode Then
        If Not m_colDictLoaded Then
            TryColIndex = -1
            Exit Function
        End If
        If m_colDict.Exists(metricName) Then
            TryColIndex = m_detailCols(m_colDict(metricName))
        Else
            TryColIndex = -1
        End If
        Exit Function
    End If
    Dim val As Variant
    val = KernelConfigLoader.FallbackColLookup(metricName, CREG_COL_DETAIL)
    If IsEmpty(val) Then
        TryColIndex = -1
    Else
        TryColIndex = CLng(val)
    End If
End Function


' --- CsvIndex ---
Public Function CsvIndex(metricName As String) As Long
    If Not m_fallbackMode Then
        If Not m_colDictLoaded Then
            CsvIndex = -1
            LogError SEV_FATAL, "KernelConfig", "E-112", "CsvIndex called before LoadAllConfig", metricName
            Exit Function
        End If
        If m_colDict.Exists(metricName) Then
            CsvIndex = m_csvCols(m_colDict(metricName))
        Else
            CsvIndex = -1
            LogError SEV_FATAL, "KernelConfig", "E-113", _
                "Column not found in registry: " & metricName, _
                "MANUAL BYPASS: Check Config sheet COLUMN_REGISTRY for available column names."
        End If
        Exit Function
    End If
    Dim val As Variant
    val = KernelConfigLoader.FallbackColLookup(metricName, CREG_COL_CSV)
    If IsEmpty(val) Then
        CsvIndex = -1
    Else
        CsvIndex = CLng(val)
    End If
End Function


' --- InputValue ---
Public Function InputValue(section As String, param As String, entityIdx As Long) As Variant
    Dim targetRow As Long
    If Not m_fallbackMode Then
        Dim paramIdx As Long
        paramIdx = FindInputParam(section, param)
        If paramIdx = 0 Then
            LogError SEV_ERROR, "KernelConfig", "E-120", _
                     "Input parameter not found: " & section & "." & param, _
                     "MANUAL BYPASS: Check Inputs tab for parameter '" & param & "' in section '" & section & "'."
            InputValue = Empty
            Exit Function
        End If
        targetRow = m_inputRows(paramIdx)
    Else
        targetRow = KernelConfigLoader.FallbackInputRow(section, param)
        If targetRow = 0 Then
            LogError SEV_ERROR, "KernelConfig", "E-120", _
                     "Input parameter not found (fallback): " & section & "." & param, _
                     "MANUAL BYPASS: Check Config sheet INPUT_SCHEMA section."
            InputValue = Empty
            Exit Function
        End If
    End If
    Dim wsInputs As Worksheet
    Set wsInputs = ThisWorkbook.Sheets(GetInputsTabName())
    Dim targetCol As Long
    targetCol = INPUT_ENTITY_START_COL + entityIdx - 1
    InputValue = wsInputs.Cells(targetRow, targetCol).Value
End Function

Private Function FindInputParam(section As String, param As String) As Long
    Dim idx As Long
    For idx = 1 To m_inputCount
        If StrComp(m_inputSections(idx), section, vbTextCompare) = 0 And _
           StrComp(m_inputParams(idx), param, vbTextCompare) = 0 Then
            FindInputParam = idx
            Exit Function
        End If
    Next idx
    FindInputParam = 0
End Function


' --- GetFieldClass ---
Public Function GetFieldClass(metricName As String) As String
    If Not m_fallbackMode Then
        If m_colDictLoaded And m_colDict.Exists(metricName) Then
            GetFieldClass = m_fieldClasses(m_colDict(metricName))
        Else
            LogError SEV_ERROR, "KernelConfig", "E-130", "GetFieldClass: column not found: " & metricName, ""
            GetFieldClass = ""
        End If
        Exit Function
    End If
    Dim val As Variant
    val = KernelConfigLoader.FallbackColLookup(metricName, CREG_COL_FIELDCLASS)
    GetFieldClass = IIf(IsEmpty(val), "", CStr(val))
End Function


' --- GetDerivationRule ---
Public Function GetDerivationRule(metricName As String) As String
    If Not m_fallbackMode Then
        If m_colDictLoaded And m_colDict.Exists(metricName) Then
            GetDerivationRule = m_derivRules(m_colDict(metricName))
        Else
            GetDerivationRule = ""
        End If
        Exit Function
    End If
    Dim val As Variant
    val = KernelConfigLoader.FallbackColLookup(metricName, CREG_COL_DERIVRULE)
    GetDerivationRule = IIf(IsEmpty(val), "", CStr(val))
End Function


' --- GetDefaultView ---
Public Function GetDefaultView(metricName As String) As String
    If Not m_fallbackMode Then
        If m_colDictLoaded And m_colDict.Exists(metricName) Then
            GetDefaultView = m_defaultViews(m_colDict(metricName))
        Else
            GetDefaultView = ""
        End If
        Exit Function
    End If
    Dim val As Variant
    val = KernelConfigLoader.FallbackColLookup(metricName, CREG_COL_DEFAULTVIEW)
    GetDefaultView = IIf(IsEmpty(val), "", CStr(val))
End Function


' --- GetFormat ---
Public Function GetFormat(metricName As String) As String
    If Not m_fallbackMode Then
        If m_colDictLoaded And m_colDict.Exists(metricName) Then
            GetFormat = m_formats(m_colDict(metricName))
        Else
            GetFormat = ""
        End If
        Exit Function
    End If
    Dim val As Variant
    val = KernelConfigLoader.FallbackColLookup(metricName, CREG_COL_FORMAT)
    GetFormat = IIf(IsEmpty(val), "", CStr(val))
End Function


' --- GetBalGrp ---
Public Function GetBalGrp(metricName As String) As String
    If Not m_fallbackMode Then
        If m_colDictLoaded And m_colDict.Exists(metricName) Then
            GetBalGrp = m_balGrps(m_colDict(metricName))
        Else
            GetBalGrp = ""
        End If
        Exit Function
    End If
    Dim val As Variant
    val = KernelConfigLoader.FallbackColLookup(metricName, CREG_COL_BALGRP)
    GetBalGrp = IIf(IsEmpty(val), "", CStr(val))
End Function


' --- GetSetting ---
Public Function GetSetting(key As String) As Variant
    If Not m_fallbackMode Then
        Dim idx As Long
        For idx = 1 To m_settingCount
            If StrComp(m_settingKeys(idx), key, vbTextCompare) = 0 Then
                GetSetting = m_settingValues(idx)
                Exit Function
            End If
        Next idx
        LogError SEV_WARN, "KernelConfig", "E-140", "Setting not found: " & key, ""
        GetSetting = ""
        Exit Function
    End If
    GetSetting = KernelConfigLoader.FallbackKeyValue("GRANULARITY_CONFIG", key, GCFG_COL_KEY, GCFG_COL_VALUE)
End Function


' --- GetEntityName ---
' Returns entity name by index. Reads from EntitySourceTab/EntitySourceRow
' in granularity_config (e.g., "UW Inputs" row 6). Falls back to input tab row 3.
Public Function GetEntityName(entityIdx As Long) As String
    GetEntityName = ""
    If entityIdx < 1 Then Exit Function

    On Error Resume Next
    ' Read source tab and row from granularity_config
    Dim srcTab As String
    srcTab = CStr(GetSetting("EntitySourceTab"))
    Dim srcRowStr As String
    srcRowStr = CStr(GetSetting("EntitySourceRow"))
    Dim srcRow As Long
    If IsNumeric(srcRowStr) And Len(srcRowStr) > 0 Then srcRow = CLng(srcRowStr)

    ' Resolve source worksheet
    Dim ws As Worksheet
    If Len(srcTab) > 0 Then
        Set ws = ThisWorkbook.Sheets(srcTab)
    End If
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets(GetInputsTabName())
    End If
    If ws Is Nothing Then Exit Function
    On Error GoTo 0

    ' Read entity name: col C + (entityIdx - 1)
    ' For UW Inputs, entities are in rows 6-15 col 3 (Program name column)
    ' Each entity is one ROW, not one column
    If srcRow > 0 Then
        GetEntityName = Trim(CStr(ws.Cells(srcRow + entityIdx - 1, INPUT_ENTITY_START_COL).Value))
    Else
        ' Fallback: entities as columns on row 3 (sample model pattern)
        GetEntityName = Trim(CStr(ws.Cells(3, INPUT_ENTITY_START_COL + entityIdx - 1).Value))
    End If
End Function


' --- GetInputsTabName (BUG-062b) ---
' Returns the resolved Inputs tab name (configurable via InputsTabName setting).
Public Function GetInputsTabName() As String
    Dim tabName As String
    tabName = CStr(GetSetting("InputsTabName"))
    If Len(tabName) = 0 Then tabName = TAB_INPUTS
    GetInputsTabName = tabName
End Function


' --- GetColumnCount ---
Public Function GetColumnCount() As Long
    If Not m_fallbackMode Then
        GetColumnCount = m_colCount
    Else
        GetColumnCount = KernelConfigLoader.FallbackColCount()
    End If
End Function


Public Function GetIncrementalColumns() As Variant
    GetIncrementalColumns = GetColumnsByClass("Incremental")
End Function

Public Function GetDerivedColumns() As Variant
    GetDerivedColumns = GetColumnsByClass("Derived")
End Function

Public Function GetDimensionColumns() As Variant
    GetDimensionColumns = GetColumnsByClass("Dimension")
End Function

Private Function GetColumnsByClass(fieldClass As String) As Variant
    Dim totalCols As Long
    totalCols = GetColumnCount()
    If totalCols = 0 Then
        GetColumnsByClass = Array()
        Exit Function
    End If
    Dim cnt As Long
    cnt = 0
    Dim idx As Long
    For idx = 1 To totalCols
        Dim colName As String
        colName = GetColName(idx)
        If StrComp(GetFieldClass(colName), fieldClass, vbTextCompare) = 0 Then
            cnt = cnt + 1
        End If
    Next idx
    If cnt = 0 Then
        GetColumnsByClass = Array()
        Exit Function
    End If
    Dim result() As String
    ReDim result(1 To cnt)
    Dim pos As Long
    pos = 0
    For idx = 1 To totalCols
        colName = GetColName(idx)
        If StrComp(GetFieldClass(colName), fieldClass, vbTextCompare) = 0 Then
            pos = pos + 1
            result(pos) = colName
        End If
    Next idx
    GetColumnsByClass = result
End Function


' --- GetTimeHorizon / GetMaxEntities ---
Public Function GetTimeHorizon() As Long
    If Not m_fallbackMode Then
        GetTimeHorizon = m_timeHorizon
    Else
        Dim val As Variant
        val = GetSetting("TimeHorizon")
        If IsNumeric(val) Then
            GetTimeHorizon = CLng(val)
        Else
            GetTimeHorizon = 0
        End If
    End If
End Function

Public Function GetMaxEntities() As Long
    If Not m_fallbackMode Then
        GetMaxEntities = m_maxEntities
    Else
        Dim val As Variant
        val = GetSetting("MaxEntities")
        If IsNumeric(val) Then
            GetMaxEntities = CLng(val)
        Else
            GetMaxEntities = 10
        End If
    End If
End Function


' --- LogError ---
Public Sub LogError(severity As Long, source As String, code As String, _
                    msg As String, detail As String)
    On Error Resume Next
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets(TAB_ERROR_LOG)
    If wsLog Is Nothing Then Exit Sub
    Dim nextRow As Long
    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).row + 1
    If nextRow < 2 Then nextRow = 2
    Dim sevLabel As String
    Select Case severity
        Case SEV_FATAL: sevLabel = "FATAL"
        Case SEV_ERROR: sevLabel = "ERROR"
        Case SEV_WARN: sevLabel = "WARN"
        Case SEV_INFO: sevLabel = "INFO"
        Case Else: sevLabel = "UNKNOWN"
    End Select
    wsLog.Cells(nextRow, 1).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    wsLog.Cells(nextRow, 2).Value = sevLabel
    wsLog.Cells(nextRow, 3).Value = source
    wsLog.Cells(nextRow, 4).Value = code
    wsLog.Cells(nextRow, 5).Value = msg
    wsLog.Cells(nextRow, 6).Value = detail
    If severity = SEV_FATAL Then m_runFatalCount = m_runFatalCount + 1
    If severity = SEV_ERROR Then m_runErrorCount = m_runErrorCount + 1
    Select Case severity
        Case SEV_FATAL
            wsLog.Cells(nextRow, 2).Interior.Color = RGB(192, 0, 0)
            wsLog.Cells(nextRow, 2).Font.Color = RGB(255, 255, 255)
            wsLog.Cells(nextRow, 2).Font.Bold = True
        Case SEV_ERROR
            wsLog.Cells(nextRow, 2).Interior.Color = RGB(255, 199, 206)
            wsLog.Cells(nextRow, 2).Font.Color = RGB(156, 0, 6)
            wsLog.Cells(nextRow, 2).Font.Bold = True
        Case SEV_WARN
            wsLog.Cells(nextRow, 2).Interior.Color = RGB(255, 235, 156)
            wsLog.Cells(nextRow, 2).Font.Color = RGB(156, 101, 0)
        Case SEV_INFO
            wsLog.Cells(nextRow, 2).Interior.Color = RGB(198, 239, 206)
            wsLog.Cells(nextRow, 2).Font.Color = RGB(0, 97, 0)
    End Select
    On Error GoTo 0
End Sub

Public Sub ResetRunErrorCounters()
    m_runFatalCount = 0
    m_runErrorCount = 0
End Sub

Public Function GetRunFatalCount() As Long
    GetRunFatalCount = m_runFatalCount
End Function

Public Function GetRunErrorCount() As Long
    GetRunErrorCount = m_runErrorCount
End Function


' --- Input schema getters ---
Public Function GetInputCount() As Long
    If Not m_fallbackMode Then
        GetInputCount = m_inputCount
        Exit Function
    End If
    GetInputCount = KernelConfigLoader.FallbackSectionCount("INPUT_SCHEMA")
End Function

Public Function GetInputSection(idx As Long) As String
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_inputCount Then
            GetInputSection = m_inputSections(idx)
        Else
            GetInputSection = ""
        End If
        Exit Function
    End If
    GetInputSection = KernelConfigLoader.FallbackInputField(idx, ISCH_COL_SECTION)
End Function

Public Function GetInputParam(idx As Long) As String
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_inputCount Then
            GetInputParam = m_inputParams(idx)
        Else
            GetInputParam = ""
        End If
        Exit Function
    End If
    GetInputParam = KernelConfigLoader.FallbackInputField(idx, ISCH_COL_PARAM)
End Function

Public Function GetInputRow(idx As Long) As Long
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_inputCount Then
            GetInputRow = m_inputRows(idx)
        Else
            GetInputRow = 0
        End If
        Exit Function
    End If
    Dim val As String
    val = KernelConfigLoader.FallbackInputField(idx, ISCH_COL_ROW)
    If IsNumeric(val) Then
        GetInputRow = CLng(val)
    Else
        GetInputRow = 0
    End If
End Function

Public Function GetInputType(idx As Long) As String
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_inputCount Then
            GetInputType = m_inputTypes(idx)
        Else
            GetInputType = ""
        End If
        Exit Function
    End If
    GetInputType = KernelConfigLoader.FallbackInputField(idx, ISCH_COL_TYPE)
End Function

Public Function GetInputDefault(idx As Long) As String
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_inputCount Then
            GetInputDefault = m_inputDefaults(idx)
        Else
            GetInputDefault = ""
        End If
        Exit Function
    End If
    GetInputDefault = KernelConfigLoader.FallbackInputField(idx, ISCH_COL_DEFAULT)
End Function


' --- GetReproSetting ---
Public Function GetReproSetting(key As String) As String
    If Not m_fallbackMode Then
        Dim idx As Long
        For idx = 1 To m_reproCount
            If StrComp(m_reproKeys(idx), key, vbTextCompare) = 0 Then
                GetReproSetting = m_reproValues(idx)
                Exit Function
            End If
        Next idx
        GetReproSetting = ""
        Exit Function
    End If
    GetReproSetting = KernelConfigLoader.FallbackKeyValue("REPRO_CONFIG", key, RCFG_COL_KEY, RCFG_COL_VALUE)
End Function


' --- GetScaleSetting ---
Public Function GetScaleSetting(key As String) As String
    If Not m_fallbackMode Then
        Dim idx As Long
        For idx = 1 To m_scaleCount
            If StrComp(m_scaleKeys(idx), key, vbTextCompare) = 0 Then
                GetScaleSetting = m_scaleValues(idx)
                Exit Function
            End If
        Next idx
        GetScaleSetting = ""
        Exit Function
    End If
    GetScaleSetting = KernelConfigLoader.FallbackKeyValue("SCALE_LIMITS", key, SCFG_COL_KEY, SCFG_COL_VALUE)
End Function


' --- GetColName / GetDetailCol ---
Public Function GetColName(idx As Long) As String
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_colCount Then
            GetColName = m_colNames(idx)
        Else
            GetColName = ""
        End If
        Exit Function
    End If
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    Dim startRow As Long
    startRow = KernelConfigLoader.FindConfigSection(wsConfig, "COLUMN_REGISTRY")
    If startRow = 0 Then
        GetColName = ""
        Exit Function
    End If
    GetColName = Trim(CStr(wsConfig.Cells(startRow + 1 + idx, CREG_COL_NAME).Value))
End Function

Public Function GetDetailCol(idx As Long) As Long
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_colCount Then
            GetDetailCol = m_detailCols(idx)
        Else
            GetDetailCol = 0
        End If
        Exit Function
    End If
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    Dim startRow As Long
    startRow = KernelConfigLoader.FindConfigSection(wsConfig, "COLUMN_REGISTRY")
    If startRow = 0 Then
        GetDetailCol = 0
        Exit Function
    End If
    Dim val As Variant
    val = wsConfig.Cells(startRow + 1 + idx, CREG_COL_DETAIL).Value
    If IsNumeric(val) Then
        GetDetailCol = CLng(val)
    Else
        GetDetailCol = 0
    End If
End Function


' --- ProveIt getters ---
Public Function GetProveItCheckCount() As Long
    If Not m_fallbackMode Then
        GetProveItCheckCount = m_piCount
        Exit Function
    End If
    GetProveItCheckCount = KernelConfigLoader.FallbackSectionCount("PROVE_IT_CONFIG")
End Function

Public Function GetProveItCheckID(idx As Long) As String
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_piCount Then GetProveItCheckID = m_piCheckIDs(idx)
        Exit Function
    End If
    GetProveItCheckID = KernelConfigLoader.FallbackProveItField(idx, PCFG_COL_CHECKID)
End Function

Public Function GetProveItCheckType(idx As Long) As String
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_piCount Then GetProveItCheckType = m_piCheckTypes(idx)
        Exit Function
    End If
    GetProveItCheckType = KernelConfigLoader.FallbackProveItField(idx, PCFG_COL_CHECKTYPE)
End Function

Public Function GetProveItCheckName(idx As Long) As String
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_piCount Then GetProveItCheckName = m_piCheckNames(idx)
        Exit Function
    End If
    GetProveItCheckName = KernelConfigLoader.FallbackProveItField(idx, PCFG_COL_CHECKNAME)
End Function

Public Function GetProveItMetricA(idx As Long) As String
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_piCount Then GetProveItMetricA = m_piMetricA(idx)
        Exit Function
    End If
    GetProveItMetricA = KernelConfigLoader.FallbackProveItField(idx, PCFG_COL_METRICA)
End Function

Public Function GetProveItMetricB(idx As Long) As String
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_piCount Then GetProveItMetricB = m_piMetricB(idx)
        Exit Function
    End If
    GetProveItMetricB = KernelConfigLoader.FallbackProveItField(idx, PCFG_COL_METRICB)
End Function

Public Function GetProveItMetricC(idx As Long) As String
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_piCount Then GetProveItMetricC = m_piMetricC(idx)
        Exit Function
    End If
    GetProveItMetricC = KernelConfigLoader.FallbackProveItField(idx, PCFG_COL_METRICC)
End Function

Public Function GetProveItOperator(idx As Long) As String
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_piCount Then GetProveItOperator = m_piOperators(idx)
        Exit Function
    End If
    GetProveItOperator = KernelConfigLoader.FallbackProveItField(idx, PCFG_COL_OPERATOR)
End Function

Public Function GetProveItTolerance(idx As Long) As Double
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_piCount Then
            GetProveItTolerance = m_piTolerances(idx)
        Else
            GetProveItTolerance = 0.000001
        End If
        Exit Function
    End If
    Dim val As String
    val = KernelConfigLoader.FallbackProveItField(idx, PCFG_COL_TOLERANCE)
    If IsNumeric(val) And Len(val) > 0 Then
        GetProveItTolerance = CDbl(val)
    Else
        GetProveItTolerance = 0.000001
    End If
End Function

Public Function GetProveItEnabled(idx As Long) As Boolean
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_piCount Then
            GetProveItEnabled = m_piEnabled(idx)
        Else
            GetProveItEnabled = False
        End If
        Exit Function
    End If
    Dim val As String
    val = KernelConfigLoader.FallbackProveItField(idx, PCFG_COL_ENABLED)
    GetProveItEnabled = (StrComp(val, "TRUE", vbTextCompare) = 0)
End Function


' --- Phase 5A: Summary/Chart/Exhibit count + field getters ---
Public Function GetSummaryConfigCount() As Long
    If Not m_fallbackMode Then
        GetSummaryConfigCount = m_sumCfgCount
    Else
        GetSummaryConfigCount = KernelConfigLoader.FallbackSectionCount("SUMMARY_CONFIG")
    End If
End Function

Public Function GetChartRegistryCount() As Long
    If Not m_fallbackMode Then
        GetChartRegistryCount = m_chtCfgCount
    Else
        GetChartRegistryCount = KernelConfigLoader.FallbackSectionCount("CHART_REGISTRY")
    End If
End Function

Public Function GetExhibitConfigCount() As Long
    If Not m_fallbackMode Then
        GetExhibitConfigCount = m_exhCfgCount
    Else
        GetExhibitConfigCount = KernelConfigLoader.FallbackSectionCount("EXHIBIT_CONFIG")
    End If
End Function

Public Function GetSummaryConfigField(idx As Long, col As Long) As String
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_sumCfgCount And col >= 1 And col <= 5 Then
            GetSummaryConfigField = m_sumCfg(idx, col)
        End If
        Exit Function
    End If
    GetSummaryConfigField = KernelConfigLoader.FallbackSectionField("SUMMARY_CONFIG", idx, col)
End Function

Public Function GetChartRegistryField(idx As Long, col As Long) As String
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_chtCfgCount And col >= 1 And col <= 8 Then
            GetChartRegistryField = m_chtCfg(idx, col)
        End If
        Exit Function
    End If
    GetChartRegistryField = KernelConfigLoader.FallbackSectionField("CHART_REGISTRY", idx, col)
End Function

Public Function GetExhibitConfigField(idx As Long, col As Long) As String
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_exhCfgCount And col >= 1 And col <= 7 Then
            GetExhibitConfigField = m_exhCfg(idx, col)
        End If
        Exit Function
    End If
    GetExhibitConfigField = KernelConfigLoader.FallbackSectionField("EXHIBIT_CONFIG", idx, col)
End Function


' --- Phase 5A: Display mode getters ---
Public Function GetDisplayModeSetting(key As String) As String
    If Not m_fallbackMode Then
        Dim i As Long
        For i = 1 To m_dispCount
            If StrComp(m_dispKeys(i), key, vbTextCompare) = 0 Then
                GetDisplayModeSetting = m_dispValues(i)
                Exit Function
            End If
        Next i
        GetDisplayModeSetting = ""
        Exit Function
    End If
    GetDisplayModeSetting = KernelConfigLoader.FallbackKeyValue("DISPLAY_MODE_CONFIG", key, DMCFG_COL_KEY, DMCFG_COL_VALUE)
End Function

Public Function GetCurrentDisplayMode() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_CONFIG)
    Dim sr As Long
    sr = KernelConfigLoader.FindConfigSection(ws, "DISPLAY_MODE_CONFIG")
    If sr > 0 Then
        Dim r As Long
        r = sr + 2
        Do While Len(Trim(CStr(ws.Cells(r, 1).Value))) > 0
            If StrComp(Trim(CStr(ws.Cells(r, DMCFG_COL_KEY).Value)), "DefaultMode", vbTextCompare) = 0 Then
                Dim v As String
                v = Trim(CStr(ws.Cells(r, DMCFG_COL_VALUE).Value))
                If v = DISPLAY_CUMULATIVE Or v = DISPLAY_INCREMENTAL Then
                    GetCurrentDisplayMode = v
                    Exit Function
                End If
            End If
            r = r + 1
        Loop
    End If
    GetCurrentDisplayMode = DISPLAY_INCREMENTAL
End Function

Public Sub SetCurrentDisplayMode(mode As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_CONFIG)
    Dim sr As Long
    sr = KernelConfigLoader.FindConfigSection(ws, "DISPLAY_MODE_CONFIG")
    If sr = 0 Then Exit Sub
    Dim r As Long
    r = sr + 2
    Do While Len(Trim(CStr(ws.Cells(r, 1).Value))) > 0
        If StrComp(Trim(CStr(ws.Cells(r, DMCFG_COL_KEY).Value)), "DefaultMode", vbTextCompare) = 0 Then
            ws.Cells(r, DMCFG_COL_VALUE).Value = mode
            Exit Sub
        End If
        r = r + 1
    Loop
    On Error GoTo 0
End Sub


' --- Dev Mode state ---
Public Function GetDevMode() As String
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_CONFIG)
    Dim sr As Long
    sr = KernelConfigLoader.FindConfigSection(ws, "DEV_MODE")
    If sr > 0 Then
        Dim r As Long
        r = sr + 2
        Do While Len(Trim(CStr(ws.Cells(r, 1).Value))) > 0
            If StrComp(Trim(CStr(ws.Cells(r, DEVCFG_COL_KEY).Value)), "DevMode", vbTextCompare) = 0 Then
                Dim v As String
                v = Trim(CStr(ws.Cells(r, DEVCFG_COL_VALUE).Value))
                If v = DEV_MODE_ON Or v = DEV_MODE_OFF Then
                    GetDevMode = v
                    Exit Function
                End If
            End If
            r = r + 1
        Loop
    End If
    GetDevMode = DEV_MODE_OFF
    On Error GoTo 0
End Function

Public Function GetBrandingSetting(key As String) As String
    GetBrandingSetting = ""
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_CONFIG)
    If ws Is Nothing Then Exit Function
    On Error GoTo 0
    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(ws, CFG_MARKER_BRANDING_CONFIG)
    If sr = 0 Then Exit Function
    Dim r As Long
    r = sr + 2
    Do While Len(Trim(CStr(ws.Cells(r, BRCFG_COL_KEY).Value))) > 0
        If StrComp(Trim(CStr(ws.Cells(r, BRCFG_COL_KEY).Value)), key, vbTextCompare) = 0 Then
            GetBrandingSetting = Trim(CStr(ws.Cells(r, BRCFG_COL_VALUE).Value))
            Exit Function
        End If
        r = r + 1
    Loop
End Function

' --- GetMsgBox ---
' Returns a message template from msgbox_config by MsgBoxID.
' Caller resolves placeholders like {ENTITIES}, {ELAPSED}, etc.
Public Function GetMsgBox(msgBoxID As String) As Variant
    ' Returns Array(Title, Message, Icon, Buttons) or Empty if not found
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_CONFIG)
    If ws Is Nothing Then
        GetMsgBox = Empty
        Exit Function
    End If
    On Error GoTo 0
    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(ws, CFG_MARKER_MSGBOX_CONFIG)
    If sr = 0 Then
        GetMsgBox = Empty
        Exit Function
    End If
    Dim r As Long
    r = sr + 2
    Do While Len(Trim(CStr(ws.Cells(r, MBCFG_COL_ID).Value))) > 0
        If StrComp(Trim(CStr(ws.Cells(r, MBCFG_COL_ID).Value)), msgBoxID, vbTextCompare) = 0 Then
            GetMsgBox = Array( _
                Trim(CStr(ws.Cells(r, MBCFG_COL_TITLE).Value)), _
                Trim(CStr(ws.Cells(r, MBCFG_COL_MESSAGE).Value)), _
                Trim(CStr(ws.Cells(r, MBCFG_COL_ICON).Value)), _
                Trim(CStr(ws.Cells(r, MBCFG_COL_BUTTONS).Value)))
            Exit Function
        End If
        r = r + 1
    Loop
    GetMsgBox = Empty
End Function


' --- GetDisplayAlias ---
' Returns the user-facing display name for an internal ID, or the ID itself if no alias.
Public Function GetDisplayAlias(internalID As String) As String
    GetDisplayAlias = internalID  ' default: return ID itself
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_CONFIG)
    If ws Is Nothing Then Exit Function
    On Error GoTo 0
    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(ws, CFG_MARKER_DISPLAY_ALIASES)
    If sr = 0 Then Exit Function
    Dim r As Long
    r = sr + 2
    Do While Len(Trim(CStr(ws.Cells(r, DACFG_COL_ID).Value))) > 0
        If StrComp(Trim(CStr(ws.Cells(r, DACFG_COL_ID).Value)), internalID, vbTextCompare) = 0 Then
            Dim displayName As String
            displayName = Trim(CStr(ws.Cells(r, DACFG_COL_DISPLAY).Value))
            If Len(displayName) > 0 Then GetDisplayAlias = displayName
            Exit Function
        End If
        r = r + 1
    Loop
End Function


' --- GetWorkspaceSetting ---
Public Function GetWorkspaceSetting(key As String) As String
    GetWorkspaceSetting = ""
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_CONFIG)
    If ws Is Nothing Then Exit Function
    On Error GoTo 0
    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(ws, CFG_MARKER_WORKSPACE_CONFIG)
    If sr = 0 Then Exit Function
    Dim r As Long
    r = sr + 2
    Do While Len(Trim(CStr(ws.Cells(r, WSCFG_COL_KEY).Value))) > 0
        If StrComp(Trim(CStr(ws.Cells(r, WSCFG_COL_KEY).Value)), key, vbTextCompare) = 0 Then
            GetWorkspaceSetting = Trim(CStr(ws.Cells(r, WSCFG_COL_VALUE).Value))
            Exit Function
        End If
        r = r + 1
    Loop
End Function

Public Function GetLockSetting(key As String) As String
    GetLockSetting = ""
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_CONFIG)
    If ws Is Nothing Then Exit Function
    On Error GoTo 0
    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(ws, CFG_MARKER_LOCK_CONFIG)
    If sr = 0 Then Exit Function
    Dim r As Long
    r = sr + 2
    Do While Len(Trim(CStr(ws.Cells(r, 1).Value))) > 0
        If StrComp(Trim(CStr(ws.Cells(r, 1).Value)), key, vbTextCompare) = 0 Then
            GetLockSetting = Trim(CStr(ws.Cells(r, 2).Value))
            Exit Function
        End If
        r = r + 1
    Loop
End Function


' --- GetConfigVersion ---
' Returns a key from config_version section (ConfigVersion, ConfigName, etc.)
Public Function GetConfigVersion(key As String) As String
    GetConfigVersion = ""
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_CONFIG)
    If ws Is Nothing Then Exit Function
    On Error GoTo 0
    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(ws, CFG_MARKER_CONFIG_VERSION)
    If sr = 0 Then Exit Function
    Dim r As Long
    r = sr + 2
    Do While Len(Trim(CStr(ws.Cells(r, 1).Value))) > 0
        If StrComp(Trim(CStr(ws.Cells(r, 1).Value)), key, vbTextCompare) = 0 Then
            GetConfigVersion = Trim(CStr(ws.Cells(r, 2).Value))
            Exit Function
        End If
        r = r + 1
    Loop
End Function


Public Sub SetDevMode(mode As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_CONFIG)
    Dim sr As Long
    sr = KernelConfigLoader.FindConfigSection(ws, "DEV_MODE")
    If sr = 0 Then
        ' Section doesn't exist yet -- create it at next available row
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 2
        ws.Cells(lastRow, 1).NumberFormat = "@"
        ws.Cells(lastRow, 1).Value = CFG_MARKER_DEV_MODE
        ws.Cells(lastRow, 1).Font.Bold = True
        ws.Cells(lastRow + 1, 1).NumberFormat = "@"
        ws.Cells(lastRow + 1, 1).Value = "Setting"
        ws.Cells(lastRow + 1, 2).NumberFormat = "@"
        ws.Cells(lastRow + 1, 2).Value = "Value"
        ws.Cells(lastRow + 2, 1).NumberFormat = "@"
        ws.Cells(lastRow + 2, 1).Value = "DevMode"
        ws.Cells(lastRow + 2, 2).NumberFormat = "@"
        ws.Cells(lastRow + 2, 2).Value = mode
        Exit Sub
    End If
    Dim r As Long
    r = sr + 2
    Do While Len(Trim(CStr(ws.Cells(r, 1).Value))) > 0
        If StrComp(Trim(CStr(ws.Cells(r, DEVCFG_COL_KEY).Value)), "DevMode", vbTextCompare) = 0 Then
            ws.Cells(r, DEVCFG_COL_VALUE).Value = mode
            Exit Sub
        End If
        r = r + 1
    Loop
    ' Key not found -- append it
    ws.Cells(r, 1).NumberFormat = "@"
    ws.Cells(r, 1).Value = "DevMode"
    ws.Cells(r, 2).NumberFormat = "@"
    ws.Cells(r, 2).Value = mode
    On Error GoTo 0
End Sub


' --- Phase 5B: Print config getters ---
Public Function GetPrintConfigCount() As Long
    If Not m_fallbackMode Then
        GetPrintConfigCount = m_prtCfgCount
    Else
        GetPrintConfigCount = KernelConfigLoader.FallbackSectionCount("PRINT_CONFIG")
    End If
End Function

Public Function GetPrintConfigField(idx As Long, col As Long) As String
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_prtCfgCount And col >= 1 And col <= 14 Then
            GetPrintConfigField = m_prtCfg(idx, col)
        End If
        Exit Function
    End If
    GetPrintConfigField = KernelConfigLoader.FallbackSectionField("PRINT_CONFIG", idx, col)
End Function


' --- Phase 5B: Data model config getter ---
Public Function GetDataModelSetting(key As String) As String
    If Not m_fallbackMode Then
        Dim i As Long
        For i = 1 To m_dmCount
            If StrComp(m_dmKeys(i), key, vbTextCompare) = 0 Then
                GetDataModelSetting = m_dmValues(i)
                Exit Function
            End If
        Next i
        GetDataModelSetting = ""
        Exit Function
    End If
    GetDataModelSetting = KernelConfigLoader.FallbackKeyValue("DATA_MODEL_CONFIG", key, DMCFG_COL_KEY, DMCFG_COL_VALUE)
End Function


' --- Phase 5B: Pivot config getters ---
Public Function GetPivotConfigCount() As Long
    If Not m_fallbackMode Then
        GetPivotConfigCount = m_pvtCfgCount
    Else
        GetPivotConfigCount = KernelConfigLoader.FallbackSectionCount("PIVOT_CONFIG")
    End If
End Function

Public Function GetPivotConfigField(idx As Long, col As Long) As String
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_pvtCfgCount And col >= 1 And col <= 8 Then
            GetPivotConfigField = m_pvtCfg(idx, col)
        End If
        Exit Function
    End If
    GetPivotConfigField = KernelConfigLoader.FallbackSectionField("PIVOT_CONFIG", idx, col)
End Function


' --- Phase 5C: BalanceType getter ---
Public Function GetBalanceType(metricName As String) As String
    If Not m_fallbackMode Then
        If m_colDictLoaded And m_colDict.Exists(metricName) Then
            GetBalanceType = m_balanceTypes(m_colDict(metricName))
        Else
            GetBalanceType = BALANCE_TYPE_FLOW
        End If
        Exit Function
    End If
    Dim val As Variant
    val = KernelConfigLoader.FallbackColLookup(metricName, CREG_COL_BALANCETYPE)
    If IsEmpty(val) Or Len(Trim(CStr(val))) = 0 Then
        GetBalanceType = BALANCE_TYPE_FLOW
    Else
        GetBalanceType = CStr(val)
    End If
End Function


' --- Phase 5C: Formula tab config getters ---
Public Function GetFormulaTabConfigCount() As Long
    If Not m_fallbackMode Then
        GetFormulaTabConfigCount = m_ftCfgCount
    Else
        GetFormulaTabConfigCount = KernelConfigLoader.FallbackSectionCount("FORMULA_TAB_CONFIG")
    End If
End Function

Public Function GetFormulaTabConfigField(idx As Long, col As Long) As String
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_ftCfgCount And col >= 1 And col <= FTCFG_COL_BALANCEITEM Then
            GetFormulaTabConfigField = m_ftCfg(idx, col)
        End If
        Exit Function
    End If
    GetFormulaTabConfigField = KernelConfigLoader.FallbackSectionField("FORMULA_TAB_CONFIG", idx, col)
End Function


' --- Phase 5C: Named range registry getters ---
Public Function GetNamedRangeCount() As Long
    If Not m_fallbackMode Then
        GetNamedRangeCount = m_nrCfgCount
    Else
        GetNamedRangeCount = KernelConfigLoader.FallbackSectionCount("NAMED_RANGE_REGISTRY")
    End If
End Function

Public Function GetNamedRangeField(idx As Long, col As Long) As String
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_nrCfgCount And col >= 1 And col <= 6 Then
            GetNamedRangeField = m_nrCfg(idx, col)
        End If
        Exit Function
    End If
    GetNamedRangeField = KernelConfigLoader.FallbackSectionField("NAMED_RANGE_REGISTRY", idx, col)
End Function


' --- Phase 6A: Extension registry getters ---
Public Function GetExtensionCount() As Long
    If Not m_fallbackMode Then
        GetExtensionCount = m_extCfgCount
    Else
        GetExtensionCount = KernelConfigLoader.FallbackSectionCount("EXTENSION_REGISTRY")
    End If
End Function

Public Function GetExtensionField(idx As Long, col As Long) As String
    If Not m_fallbackMode Then
        If idx >= 1 And idx <= m_extCfgCount And col >= 1 And col <= 9 Then
            GetExtensionField = m_extCfg(idx, col)
        End If
        Exit Function
    End If
    GetExtensionField = KernelConfigLoader.FallbackSectionField("EXTENSION_REGISTRY", idx, col)
End Function


' --- Phase 6A: Curve library config getters ---
Public Function GetCurveLibraryCount() As Long
    If Not m_fallbackMode And m_clCfgCount > 0 Then
        GetCurveLibraryCount = m_clCfgCount
    Else
        ' Fallback: read from Config sheet (BUG-167: UDFs evaluate before LoadAllConfig)
        GetCurveLibraryCount = KernelConfigLoader.FallbackSectionCount("CURVE_LIBRARY_CONFIG")
    End If
End Function

Public Function GetCurveLibraryField(idx As Long, col As Long) As String
    If Not m_fallbackMode And m_clCfgCount > 0 Then
        If idx >= 1 And idx <= m_clCfgCount And col >= 1 And col <= 21 Then
            GetCurveLibraryField = m_clCfg(idx, col)
        End If
        Exit Function
    End If
    ' Fallback: read from Config sheet (BUG-167)
    GetCurveLibraryField = KernelConfigLoader.FallbackSectionField("CURVE_LIBRARY_CONFIG", idx, col)
End Function


' --- Phase 6A: Report config getter (key-value) ---
Public Function GetReportSetting(key As String) As String
    If Not m_fallbackMode Then
        Dim i As Long
        For i = 1 To m_rptCount
            If StrComp(m_rptKeys(i), key, vbTextCompare) = 0 Then
                GetReportSetting = m_rptValues(i)
                Exit Function
            End If
        Next i
        GetReportSetting = ""
        Exit Function
    End If
    GetReportSetting = KernelConfigLoader.FallbackKeyValue("REPORT_CONFIG", key, 1, 2)
End Function
