Attribute VB_Name = "KernelConfigLoader"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelConfigLoader.bas
' Purpose: All config section loading functions, split from KernelConfig.bas
'          at Phase 5B. Called once at startup by KernelConfig.LoadAllConfig.
'          Also contains fallback helpers used by KernelConfig getters (AP-45).
' Note: Writes to KernelConfig Public m_ arrays (AP-13 exception for split).
'
' CONFIG-MISSING CONVENTION (TD-06):
'   Domain config missing = FAIL FAST with LogError (E-level error code).
'     - formula_tab_config.csv rows referencing nonexistent RowIDs
'     - named_range_registry referencing nonexistent tabs
'     - curve_library_config missing required LOB/CurveType combinations
'     These are authored data; if they reference something absent, it is
'     a build error that must surface immediately.
'
'   Infrastructure config missing = FALLBACK with LogEvent (W-level warning).
'     - Snapshot files not found -> skip restore, log warning
'     - Optional config tables empty -> use defaults, log warning
'     - Print/chart/exhibit config empty -> skip feature, log warning
'     These are optional subsystems; graceful degradation is correct behavior.
' =============================================================================


' Section position cache: Dictionary of marker -> row number
' Built on first call, cleared by ClearSectionCache (called from LoadConfigFromDisk)
Private m_sectionCache As Object
Private m_sectionCacheBuilt As Boolean

Public Sub ClearSectionCache()
    Set m_sectionCache = Nothing
    m_sectionCacheBuilt = False
End Sub

Public Sub BuildSectionCache(ws As Worksheet)
    Set m_sectionCache = CreateObject("Scripting.Dictionary")
    m_sectionCache.CompareMode = vbTextCompare
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastRow < 1 Then lastRow = 1
    Dim scanRow As Long
    For scanRow = 1 To lastRow
        Dim cv As String
        cv = Trim(StripBOM(CStr(ws.Cells(scanRow, 1).Value)))
        If Left(cv, 4) = "=== " And Right(cv, 4) = " ===" Then
            If Not m_sectionCache.Exists(cv) Then
                m_sectionCache.Add cv, scanRow
            End If
        End If
    Next scanRow
    m_sectionCacheBuilt = True
End Sub


' =============================================================================
' FindSectionStart
' Returns the cached row number for a section marker, or 0 if not found.
' BUG-194: Lazy-builds the section cache on first call. Prior implementation
' fell through to a linear scan when the cache was not yet built, which caused
' catastrophic slowdowns when UDFs (e.g. Ext_CurveLib.CurveRefPct) fired
' before Bootstrap had run BuildSectionCache. Building the cache once makes
' all subsequent lookups O(1) dictionary hits.
' =============================================================================
Public Function FindSectionStart(ws As Worksheet, marker As String) As Long
    ' BUG-194: Lazy-build cache on first call so UDF callers get O(1) after the
    ' first hit instead of re-scanning the Config sheet on every invocation.
    If Not m_sectionCacheBuilt Then BuildSectionCache ws

    If Not m_sectionCache Is Nothing Then
        If m_sectionCache.Exists(marker) Then
            FindSectionStart = m_sectionCache(marker)
            Exit Function
        End If
    End If
    FindSectionStart = 0
End Function


' =============================================================================
' FindConfigSection (AP-45 fallback helper)
' Looks up row of "=== SECTION_NAME ===" marker. Returns row number or 0.
' BUG-194: Now delegates to FindSectionStart so it reuses the section cache.
' Prior implementation did a fresh linear scan on every call, causing UDF
' callers (CurveRefPct et al.) to re-scan thousands of Config rows per cell.
' =============================================================================
Public Function FindConfigSection(ws As Worksheet, sectionName As String) As Long
    FindConfigSection = FindSectionStart(ws, "=== " & sectionName & " ===")
End Function


' =============================================================================
' StripBOM (TD-05)
' Strips leading UTF-8 BOM character (Chr(65279) / U+FEFF) from a string.
' Called on the first cell of any CSV section to prevent BOM corruption.
' =============================================================================
Public Function StripBOM(val As String) As String
    If Len(val) > 0 Then
        If AscW(Left(val, 1)) = 65279 Then
            StripBOM = Mid(val, 2)
        Else
            StripBOM = val
        End If
    Else
        StripBOM = val
    End If
End Function


' =============================================================================
' ValidateConfigReferences (T1-2)
' Checks cross-references between config tables after all are loaded.
' Logs warnings for broken references (does not halt -- graceful degradation).
' =============================================================================
Public Sub ValidateConfigReferences()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_CONFIG)
    If ws Is Nothing Then Exit Sub

    ' Check named_range_registry: verify referenced tabs exist in tab_registry
    Dim nrSr As Long
    nrSr = FindSectionStart(ws, CFG_MARKER_NAMED_RANGE_REGISTRY)
    If nrSr > 0 Then
        Dim nrDr As Long
        nrDr = nrSr + 2
        Do While Len(Trim(CStr(ws.Cells(nrDr, NRCFG_COL_NAME).Value))) > 0
            Dim nrTab As String
            nrTab = Trim(CStr(ws.Cells(nrDr, NRCFG_COL_TABNAME).Value))
            If Len(nrTab) > 0 Then
                If Not TabExistsInRegistry(ws, nrTab) Then
                    KernelConfig.LogError SEV_WARN, "ConfigValidator", "W-150", _
                        "named_range_registry references tab '" & nrTab & "' not in tab_registry", _
                        "Range: " & Trim(CStr(ws.Cells(nrDr, NRCFG_COL_NAME).Value))
                End If
            End If
            nrDr = nrDr + 1
        Loop
    End If

    ' Check formula_tab_config: verify referenced tabs exist in tab_registry
    Dim ftSr As Long
    ftSr = FindSectionStart(ws, CFG_MARKER_FORMULA_TAB_CONFIG)
    If ftSr > 0 Then
        Dim ftDr As Long
        ftDr = ftSr + 2
        Dim checkedTabs As Object
        Set checkedTabs = CreateObject("Scripting.Dictionary")
        checkedTabs.CompareMode = vbTextCompare
        Do While Len(Trim(CStr(ws.Cells(ftDr, FTCFG_COL_TABNAME).Value))) > 0
            Dim ftTab As String
            ftTab = Trim(CStr(ws.Cells(ftDr, FTCFG_COL_TABNAME).Value))
            If Len(ftTab) > 0 And Not checkedTabs.Exists(ftTab) Then
                checkedTabs.Add ftTab, True
                If Not TabExistsInRegistry(ws, ftTab) Then
                    KernelConfig.LogError SEV_WARN, "ConfigValidator", "W-151", _
                        "formula_tab_config references tab '" & ftTab & "' not in tab_registry", ""
                End If
            End If
            ftDr = ftDr + 1
        Loop
    End If

    ' Check DomainModule setting is not empty
    Dim domMod As String
    domMod = KernelConfig.GetSetting("DomainModule")
    If Len(domMod) = 0 Then
        KernelConfig.LogError SEV_WARN, "ConfigValidator", "W-152", _
            "DomainModule not set in granularity_config. Default will be used.", ""
    End If

    On Error GoTo 0
End Sub


' TabExistsInRegistry -- checks if a tab name is listed in tab_registry
Private Function TabExistsInRegistry(ws As Worksheet, tabName As String) As Boolean
    TabExistsInRegistry = False
    Dim sr As Long
    sr = FindSectionStart(ws, CFG_MARKER_TAB_REGISTRY)
    If sr = 0 Then Exit Function
    Dim dr As Long
    dr = sr + 2
    Do While Len(Trim(CStr(ws.Cells(dr, TREG_COL_TABNAME).Value))) > 0
        If StrComp(Trim(CStr(ws.Cells(dr, TREG_COL_TABNAME).Value)), tabName, vbTextCompare) = 0 Then
            TabExistsInRegistry = True
            Exit Function
        End If
        dr = dr + 1
    Loop
End Function


' =============================================================================
' LoadColumnRegistry
' =============================================================================
Public Sub LoadColumnRegistry()
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)

    Dim startRow As Long
    startRow = FindSectionStart(wsConfig, CFG_MARKER_COLUMN_REGISTRY)
    If startRow = 0 Then
        KernelConfig.LogError SEV_FATAL, "KernelConfig", "E-101", _
            "COLUMN_REGISTRY section not found on Config sheet", _
            "MANUAL BYPASS: Verify Config sheet contains '=== COLUMN_REGISTRY ===' marker row."
        Exit Sub
    End If

    Dim headerRow As Long
    headerRow = startRow + 1
    Dim dataRow As Long
    dataRow = headerRow + 1
    Dim cnt As Long
    cnt = 0
    Do While Trim(CStr(wsConfig.Cells(dataRow + cnt, 1).Value)) <> ""
        cnt = cnt + 1
    Loop

    KernelConfig.m_colCount = cnt
    If cnt = 0 Then
        KernelConfig.LogError SEV_WARN, "KernelConfig", "E-102", "COLUMN_REGISTRY has no data rows", ""
        Exit Sub
    End If

    ReDim KernelConfig.m_colNames(1 To cnt)
    ReDim KernelConfig.m_detailCols(1 To cnt)
    ReDim KernelConfig.m_csvCols(1 To cnt)
    ReDim KernelConfig.m_blocks(1 To cnt)
    ReDim KernelConfig.m_fieldClasses(1 To cnt)
    ReDim KernelConfig.m_defaultViews(1 To cnt)
    ReDim KernelConfig.m_formats(1 To cnt)
    ReDim KernelConfig.m_balGrps(1 To cnt)
    ReDim KernelConfig.m_derivRules(1 To cnt)
    ReDim KernelConfig.m_balanceTypes(1 To cnt)

    Dim idx As Long
    For idx = 1 To cnt
        Dim r As Long
        r = dataRow + idx - 1
        KernelConfig.m_colNames(idx) = Trim(CStr(wsConfig.Cells(r, CREG_COL_NAME).Value))
        KernelConfig.m_detailCols(idx) = CLng(wsConfig.Cells(r, CREG_COL_DETAIL).Value)
        KernelConfig.m_csvCols(idx) = CLng(wsConfig.Cells(r, CREG_COL_CSV).Value)
        KernelConfig.m_blocks(idx) = Trim(CStr(wsConfig.Cells(r, CREG_COL_BLOCK).Value))
        KernelConfig.m_fieldClasses(idx) = Trim(CStr(wsConfig.Cells(r, CREG_COL_FIELDCLASS).Value))
        KernelConfig.m_defaultViews(idx) = Trim(CStr(wsConfig.Cells(r, CREG_COL_DEFAULTVIEW).Value))
        KernelConfig.m_formats(idx) = Trim(CStr(wsConfig.Cells(r, CREG_COL_FORMAT).Value))
        KernelConfig.m_balGrps(idx) = Trim(CStr(wsConfig.Cells(r, CREG_COL_BALGRP).Value))
        KernelConfig.m_derivRules(idx) = Trim(CStr(wsConfig.Cells(r, CREG_COL_DERIVRULE).Value))
        Dim btVal As String
        btVal = Trim(CStr(wsConfig.Cells(r, CREG_COL_BALANCETYPE).Value))
        If Len(btVal) = 0 Then btVal = BALANCE_TYPE_FLOW
        KernelConfig.m_balanceTypes(idx) = btVal
    Next idx

    Set KernelConfig.m_colDict = CreateObject("Scripting.Dictionary")
    KernelConfig.m_colDict.CompareMode = vbTextCompare
    For idx = 1 To cnt
        If Not KernelConfig.m_colDict.Exists(KernelConfig.m_colNames(idx)) Then
            KernelConfig.m_colDict.Add KernelConfig.m_colNames(idx), idx
        End If
    Next idx
    KernelConfig.m_colDictLoaded = True
End Sub


' =============================================================================
' LoadInputSchema
' =============================================================================
Public Sub LoadInputSchema()
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)

    Dim startRow As Long
    startRow = FindSectionStart(wsConfig, CFG_MARKER_INPUT_SCHEMA)
    If startRow = 0 Then
        KernelConfig.LogError SEV_FATAL, "KernelConfig", "E-103", _
            "INPUT_SCHEMA section not found on Config sheet", _
            "MANUAL BYPASS: Verify Config sheet contains '=== INPUT_SCHEMA ===' marker row."
        Exit Sub
    End If

    Dim dataRow As Long
    dataRow = startRow + 2
    Dim cnt As Long
    cnt = 0
    Do While Trim(CStr(wsConfig.Cells(dataRow + cnt, 1).Value)) <> ""
        cnt = cnt + 1
    Loop

    KernelConfig.m_inputCount = cnt
    If cnt = 0 Then
        KernelConfig.LogError SEV_WARN, "KernelConfig", "E-104", "INPUT_SCHEMA has no data rows", ""
        Exit Sub
    End If

    ReDim KernelConfig.m_inputSections(1 To cnt)
    ReDim KernelConfig.m_inputParams(1 To cnt)
    ReDim KernelConfig.m_inputRows(1 To cnt)
    ReDim KernelConfig.m_inputTypes(1 To cnt)
    ReDim KernelConfig.m_inputDefaults(1 To cnt)

    Dim idx As Long
    For idx = 1 To cnt
        Dim r As Long
        r = dataRow + idx - 1
        KernelConfig.m_inputSections(idx) = Trim(CStr(wsConfig.Cells(r, ISCH_COL_SECTION).Value))
        KernelConfig.m_inputParams(idx) = Trim(CStr(wsConfig.Cells(r, ISCH_COL_PARAM).Value))
        KernelConfig.m_inputRows(idx) = CLng(wsConfig.Cells(r, ISCH_COL_ROW).Value)
        KernelConfig.m_inputTypes(idx) = Trim(CStr(wsConfig.Cells(r, ISCH_COL_TYPE).Value))
        KernelConfig.m_inputDefaults(idx) = Trim(CStr(wsConfig.Cells(r, ISCH_COL_DEFAULT).Value))
    Next idx
End Sub


' =============================================================================
' LoadGranularityConfig
' =============================================================================
Public Sub LoadGranularityConfig()
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)

    Dim startRow As Long
    startRow = FindSectionStart(wsConfig, CFG_MARKER_GRANULARITY_CONFIG)
    If startRow = 0 Then
        KernelConfig.LogError SEV_FATAL, "KernelConfig", "E-105", _
            "GRANULARITY_CONFIG section not found on Config sheet", _
            "MANUAL BYPASS: Verify Config sheet contains '=== GRANULARITY_CONFIG ===' marker row."
        Exit Sub
    End If

    Dim dataRow As Long
    dataRow = startRow + 2
    Dim cnt As Long
    cnt = 0
    Do While Trim(CStr(wsConfig.Cells(dataRow + cnt, 1).Value)) <> ""
        cnt = cnt + 1
    Loop

    KernelConfig.m_settingCount = cnt
    If cnt = 0 Then
        KernelConfig.LogError SEV_WARN, "KernelConfig", "E-106", "GRANULARITY_CONFIG has no data rows", ""
        Exit Sub
    End If

    ReDim KernelConfig.m_settingKeys(1 To cnt)
    ReDim KernelConfig.m_settingValues(1 To cnt)

    Dim idx As Long
    For idx = 1 To cnt
        Dim r As Long
        r = dataRow + idx - 1
        KernelConfig.m_settingKeys(idx) = Trim(CStr(wsConfig.Cells(r, GCFG_COL_KEY).Value))
        KernelConfig.m_settingValues(idx) = Trim(CStr(wsConfig.Cells(r, GCFG_COL_VALUE).Value))
    Next idx

    KernelConfig.m_timeHorizon = CLng(KernelConfig.GetSetting("TimeHorizon"))
    KernelConfig.m_maxEntities = CLng(KernelConfig.GetSetting("MaxEntities"))
    KernelConfig.m_defaultView = CStr(KernelConfig.GetSetting("DefaultSummaryView"))
End Sub


' =============================================================================
' LoadReproConfig
' =============================================================================
Public Sub LoadReproConfig()
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)

    Dim startRow As Long
    startRow = FindSectionStart(wsConfig, CFG_MARKER_REPRO_CONFIG)
    If startRow = 0 Then
        KernelConfig.m_reproCount = 0
        Exit Sub
    End If

    Dim dataRow As Long
    dataRow = startRow + 2
    Dim cnt As Long
    cnt = 0
    Do While Trim(CStr(wsConfig.Cells(dataRow + cnt, 1).Value)) <> ""
        cnt = cnt + 1
    Loop

    KernelConfig.m_reproCount = cnt
    If cnt = 0 Then Exit Sub

    ReDim KernelConfig.m_reproKeys(1 To cnt)
    ReDim KernelConfig.m_reproValues(1 To cnt)

    Dim idx As Long
    For idx = 1 To cnt
        Dim r As Long
        r = dataRow + idx - 1
        KernelConfig.m_reproKeys(idx) = Trim(CStr(wsConfig.Cells(r, RCFG_COL_KEY).Value))
        KernelConfig.m_reproValues(idx) = Trim(CStr(wsConfig.Cells(r, RCFG_COL_VALUE).Value))
    Next idx
End Sub


' =============================================================================
' LoadScaleLimits
' =============================================================================
Public Sub LoadScaleLimits()
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)

    Dim startRow As Long
    startRow = FindSectionStart(wsConfig, CFG_MARKER_SCALE_LIMITS)
    If startRow = 0 Then
        KernelConfig.m_scaleCount = 0
        Exit Sub
    End If

    Dim dataRow As Long
    dataRow = startRow + 2
    Dim cnt As Long
    cnt = 0
    Do While Trim(CStr(wsConfig.Cells(dataRow + cnt, 1).Value)) <> ""
        cnt = cnt + 1
    Loop

    KernelConfig.m_scaleCount = cnt
    If cnt = 0 Then Exit Sub

    ReDim KernelConfig.m_scaleKeys(1 To cnt)
    ReDim KernelConfig.m_scaleValues(1 To cnt)

    Dim idx As Long
    For idx = 1 To cnt
        Dim r As Long
        r = dataRow + idx - 1
        KernelConfig.m_scaleKeys(idx) = Trim(CStr(wsConfig.Cells(r, SCFG_COL_KEY).Value))
        KernelConfig.m_scaleValues(idx) = Trim(CStr(wsConfig.Cells(r, SCFG_COL_VALUE).Value))
    Next idx
End Sub


' =============================================================================
' LoadProveItConfig
' =============================================================================
Public Sub LoadProveItConfig()
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)

    Dim startRow As Long
    startRow = FindSectionStart(wsConfig, CFG_MARKER_PROVE_IT_CONFIG)
    If startRow = 0 Then
        KernelConfig.m_piCount = 0
        Exit Sub
    End If

    Dim dataRow As Long
    dataRow = startRow + 2
    Dim cnt As Long
    cnt = 0
    Do While Trim(CStr(wsConfig.Cells(dataRow + cnt, 1).Value)) <> ""
        cnt = cnt + 1
    Loop

    KernelConfig.m_piCount = cnt
    If cnt = 0 Then Exit Sub

    ReDim KernelConfig.m_piCheckIDs(1 To cnt)
    ReDim KernelConfig.m_piCheckTypes(1 To cnt)
    ReDim KernelConfig.m_piCheckNames(1 To cnt)
    ReDim KernelConfig.m_piMetricA(1 To cnt)
    ReDim KernelConfig.m_piMetricB(1 To cnt)
    ReDim KernelConfig.m_piMetricC(1 To cnt)
    ReDim KernelConfig.m_piOperators(1 To cnt)
    ReDim KernelConfig.m_piTolerances(1 To cnt)
    ReDim KernelConfig.m_piEnabled(1 To cnt)

    Dim idx As Long
    For idx = 1 To cnt
        Dim r As Long
        r = dataRow + idx - 1
        KernelConfig.m_piCheckIDs(idx) = Trim(CStr(wsConfig.Cells(r, PCFG_COL_CHECKID).Value))
        KernelConfig.m_piCheckTypes(idx) = Trim(CStr(wsConfig.Cells(r, PCFG_COL_CHECKTYPE).Value))
        KernelConfig.m_piCheckNames(idx) = Trim(CStr(wsConfig.Cells(r, PCFG_COL_CHECKNAME).Value))
        KernelConfig.m_piMetricA(idx) = Trim(CStr(wsConfig.Cells(r, PCFG_COL_METRICA).Value))
        KernelConfig.m_piMetricB(idx) = Trim(CStr(wsConfig.Cells(r, PCFG_COL_METRICB).Value))
        KernelConfig.m_piMetricC(idx) = Trim(CStr(wsConfig.Cells(r, PCFG_COL_METRICC).Value))
        KernelConfig.m_piOperators(idx) = Trim(CStr(wsConfig.Cells(r, PCFG_COL_OPERATOR).Value))
        Dim tolVal As String
        tolVal = Trim(CStr(wsConfig.Cells(r, PCFG_COL_TOLERANCE).Value))
        If IsNumeric(tolVal) And Len(tolVal) > 0 Then
            KernelConfig.m_piTolerances(idx) = CDbl(tolVal)
        Else
            KernelConfig.m_piTolerances(idx) = 0.000001
        End If
        Dim enVal As String
        enVal = Trim(CStr(wsConfig.Cells(r, PCFG_COL_ENABLED).Value))
        KernelConfig.m_piEnabled(idx) = (StrComp(enVal, "TRUE", vbTextCompare) = 0)
    Next idx
End Sub


' =============================================================================
' LoadSection2D - generic 2D section loader
' =============================================================================
Public Sub LoadSection2D(marker As String, colCnt As Long, ByRef arr() As String, ByRef cnt As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_CONFIG)
    Dim sr As Long
    sr = FindSectionStart(ws, marker)
    If sr = 0 Then
        cnt = 0
        Exit Sub
    End If
    Dim dr As Long
    dr = sr + 2
    cnt = 0
    Do While Trim(CStr(ws.Cells(dr + cnt, 1).Value)) <> ""
        cnt = cnt + 1
    Loop
    If cnt = 0 Then Exit Sub
    ReDim arr(1 To cnt, 1 To colCnt)
    Dim i As Long
    For i = 1 To cnt
        Dim c As Long
        For c = 1 To colCnt
            arr(i, c) = Trim(CStr(ws.Cells(dr + i - 1, c).Value))
        Next c
    Next i
End Sub


Public Sub LoadSummaryConfig()
    LoadSection2D CFG_MARKER_SUMMARY_CONFIG, 5, KernelConfig.m_sumCfg, KernelConfig.m_sumCfgCount
End Sub

Public Sub LoadChartRegistry()
    LoadSection2D CFG_MARKER_CHART_REGISTRY, 8, KernelConfig.m_chtCfg, KernelConfig.m_chtCfgCount
End Sub

Public Sub LoadExhibitConfig()
    LoadSection2D CFG_MARKER_EXHIBIT_CONFIG, 7, KernelConfig.m_exhCfg, KernelConfig.m_exhCfgCount
End Sub


' =============================================================================
' LoadDisplayModeConfig
' =============================================================================
Public Sub LoadDisplayModeConfig()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_CONFIG)
    Dim sr As Long
    sr = FindSectionStart(ws, CFG_MARKER_DISPLAY_MODE_CONFIG)
    If sr = 0 Then
        KernelConfig.m_dispCount = 0
        Exit Sub
    End If
    Dim dr As Long
    dr = sr + 2
    Dim cnt As Long
    cnt = 0
    Do While Trim(CStr(ws.Cells(dr + cnt, 1).Value)) <> ""
        cnt = cnt + 1
    Loop
    KernelConfig.m_dispCount = cnt
    If cnt = 0 Then Exit Sub
    ReDim KernelConfig.m_dispKeys(1 To cnt)
    ReDim KernelConfig.m_dispValues(1 To cnt)
    Dim i As Long
    For i = 1 To cnt
        KernelConfig.m_dispKeys(i) = Trim(CStr(ws.Cells(dr + i - 1, DMCFG_COL_KEY).Value))
        KernelConfig.m_dispValues(i) = Trim(CStr(ws.Cells(dr + i - 1, DMCFG_COL_VALUE).Value))
    Next i
End Sub


' =============================================================================
' Phase 5B: LoadPrintConfig
' =============================================================================
Public Sub LoadPrintConfig()
    LoadSection2D CFG_MARKER_PRINT_CONFIG, 14, KernelConfig.m_prtCfg, KernelConfig.m_prtCfgCount
End Sub


' =============================================================================
' Phase 5B: LoadDataModelConfig
' =============================================================================
Public Sub LoadDataModelConfig()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_CONFIG)
    Dim sr As Long
    sr = FindSectionStart(ws, CFG_MARKER_DATA_MODEL_CONFIG)
    If sr = 0 Then
        KernelConfig.m_dmCount = 0
        Exit Sub
    End If
    Dim dr As Long
    dr = sr + 2
    Dim cnt As Long
    cnt = 0
    Do While Trim(CStr(ws.Cells(dr + cnt, 1).Value)) <> ""
        cnt = cnt + 1
    Loop
    KernelConfig.m_dmCount = cnt
    If cnt = 0 Then Exit Sub
    ReDim KernelConfig.m_dmKeys(1 To cnt)
    ReDim KernelConfig.m_dmValues(1 To cnt)
    Dim i As Long
    For i = 1 To cnt
        KernelConfig.m_dmKeys(i) = Trim(CStr(ws.Cells(dr + i - 1, DMCFG_COL_KEY).Value))
        KernelConfig.m_dmValues(i) = Trim(CStr(ws.Cells(dr + i - 1, DMCFG_COL_VALUE).Value))
    Next i
End Sub


' =============================================================================
' Phase 5B: LoadPivotConfig
' =============================================================================
Public Sub LoadPivotConfig()
    LoadSection2D CFG_MARKER_PIVOT_CONFIG, 8, KernelConfig.m_pvtCfg, KernelConfig.m_pvtCfgCount
End Sub


' =============================================================================
' Phase 5C: LoadFormulaTabConfig
' =============================================================================
Public Sub LoadFormulaTabConfig()
    LoadSection2D CFG_MARKER_FORMULA_TAB_CONFIG, FTCFG_COL_BALANCEITEM, KernelConfig.m_ftCfg, KernelConfig.m_ftCfgCount
End Sub


' =============================================================================
' Phase 5C: LoadNamedRangeRegistry
' =============================================================================
Public Sub LoadNamedRangeRegistry()
    LoadSection2D CFG_MARKER_NAMED_RANGE_REGISTRY, 6, KernelConfig.m_nrCfg, KernelConfig.m_nrCfgCount
End Sub


' =============================================================================
' Phase 6A: LoadExtensionRegistry
' =============================================================================
Public Sub LoadExtensionRegistry()
    LoadSection2D CFG_MARKER_EXTENSION_REGISTRY, 9, KernelConfig.m_extCfg, KernelConfig.m_extCfgCount
End Sub


' =============================================================================
' Phase 6A: LoadCurveLibraryConfig
' =============================================================================
Public Sub LoadCurveLibraryConfig()
    LoadSection2D CFG_MARKER_CURVE_LIBRARY, 21, KernelConfig.m_clCfg, KernelConfig.m_clCfgCount
End Sub


' =============================================================================
' Phase 6A: LoadReportConfig
' =============================================================================
Public Sub LoadReportConfig()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_CONFIG)
    Dim sr As Long
    sr = FindSectionStart(ws, CFG_MARKER_REPORT_CONFIG)
    If sr = 0 Then
        KernelConfig.m_rptCount = 0
        Exit Sub
    End If
    Dim dr As Long
    dr = sr + 2
    Dim cnt As Long
    cnt = 0
    Do While Trim(CStr(ws.Cells(dr + cnt, 1).Value)) <> ""
        cnt = cnt + 1
    Loop
    KernelConfig.m_rptCount = cnt
    If cnt = 0 Then Exit Sub
    ReDim KernelConfig.m_rptKeys(1 To cnt)
    ReDim KernelConfig.m_rptValues(1 To cnt)
    Dim i As Long
    For i = 1 To cnt
        KernelConfig.m_rptKeys(i) = Trim(CStr(ws.Cells(dr + i - 1, 1).Value))
        KernelConfig.m_rptValues(i) = Trim(CStr(ws.Cells(dr + i - 1, 2).Value))
    Next i
End Sub


' =============================================================================
' FALLBACK HELPERS (AP-45) - used by KernelConfig getters in fallback mode
' =============================================================================

Public Function FallbackColLookup(metricName As String, targetCol As Long) As Variant
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    Dim startRow As Long
    startRow = FindConfigSection(wsConfig, "COLUMN_REGISTRY")
    If startRow = 0 Then
        FallbackColLookup = Empty
        Exit Function
    End If
    Dim r As Long
    r = startRow + 2
    Do While Len(Trim(CStr(wsConfig.Cells(r, 1).Value))) > 0
        If StrComp(Trim(CStr(wsConfig.Cells(r, 1).Value)), metricName, vbTextCompare) = 0 Then
            FallbackColLookup = wsConfig.Cells(r, targetCol).Value
            Exit Function
        End If
        r = r + 1
    Loop
    FallbackColLookup = Empty
End Function


Public Function FallbackColCount() As Long
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    Dim startRow As Long
    startRow = FindConfigSection(wsConfig, "COLUMN_REGISTRY")
    If startRow = 0 Then
        FallbackColCount = 0
        Exit Function
    End If
    Dim cnt As Long
    cnt = 0
    Dim r As Long
    r = startRow + 2
    Do While Len(Trim(CStr(wsConfig.Cells(r, 1).Value))) > 0
        cnt = cnt + 1
        r = r + 1
    Loop
    FallbackColCount = cnt
End Function


Public Function FallbackInputRow(section As String, param As String) As Long
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    Dim startRow As Long
    startRow = FindConfigSection(wsConfig, "INPUT_SCHEMA")
    If startRow = 0 Then
        FallbackInputRow = 0
        Exit Function
    End If
    Dim r As Long
    r = startRow + 2
    Do While Len(Trim(CStr(wsConfig.Cells(r, 1).Value))) > 0
        If StrComp(Trim(CStr(wsConfig.Cells(r, ISCH_COL_SECTION).Value)), section, vbTextCompare) = 0 And _
           StrComp(Trim(CStr(wsConfig.Cells(r, ISCH_COL_PARAM).Value)), param, vbTextCompare) = 0 Then
            FallbackInputRow = CLng(wsConfig.Cells(r, ISCH_COL_ROW).Value)
            Exit Function
        End If
        r = r + 1
    Loop
    FallbackInputRow = 0
End Function


Public Function FallbackInputField(idx As Long, col As Long) As String
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    Dim startRow As Long
    startRow = FindConfigSection(wsConfig, "INPUT_SCHEMA")
    If startRow = 0 Then
        FallbackInputField = ""
        Exit Function
    End If
    Dim targetRow As Long
    targetRow = startRow + 1 + idx
    If Len(Trim(CStr(wsConfig.Cells(targetRow, 1).Value))) = 0 Then
        FallbackInputField = ""
    Else
        FallbackInputField = Trim(CStr(wsConfig.Cells(targetRow, col).Value))
    End If
End Function


Public Function FallbackProveItField(idx As Long, col As Long) As String
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    Dim startRow As Long
    startRow = FindConfigSection(wsConfig, "PROVE_IT_CONFIG")
    If startRow = 0 Then
        FallbackProveItField = ""
        Exit Function
    End If
    Dim targetRow As Long
    targetRow = startRow + 1 + idx
    If Len(Trim(CStr(wsConfig.Cells(targetRow, 1).Value))) = 0 Then
        FallbackProveItField = ""
    Else
        FallbackProveItField = Trim(CStr(wsConfig.Cells(targetRow, col).Value))
    End If
End Function


Public Function FallbackSectionCount(sn As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_CONFIG)
    Dim sr As Long
    sr = FindConfigSection(ws, sn)
    If sr = 0 Then
        FallbackSectionCount = 0
        Exit Function
    End If
    Dim cnt As Long
    cnt = 0
    Dim r As Long
    r = sr + 2
    Do While Len(Trim(CStr(ws.Cells(r, 1).Value))) > 0
        cnt = cnt + 1
        r = r + 1
    Loop
    FallbackSectionCount = cnt
End Function


Public Function FallbackSectionField(sn As String, idx As Long, col As Long) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_CONFIG)
    Dim sr As Long
    sr = FindConfigSection(ws, sn)
    If sr = 0 Then
        FallbackSectionField = ""
        Exit Function
    End If
    Dim tr As Long
    tr = sr + 1 + idx
    If Len(Trim(CStr(ws.Cells(tr, 1).Value))) = 0 Then
        FallbackSectionField = ""
    Else
        FallbackSectionField = Trim(CStr(ws.Cells(tr, col).Value))
    End If
End Function


' =============================================================================
' FallbackKeyValue - generic key-value fallback for Repro/Scale/DataModel/etc
' =============================================================================
Public Function FallbackKeyValue(sectionName As String, key As String, keyCol As Long, valCol As Long) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_CONFIG)
    Dim sr As Long
    sr = FindConfigSection(ws, sectionName)
    If sr = 0 Then
        FallbackKeyValue = ""
        Exit Function
    End If
    Dim r As Long
    r = sr + 2
    Do While Len(Trim(CStr(ws.Cells(r, 1).Value))) > 0
        If StrComp(Trim(CStr(ws.Cells(r, keyCol).Value)), key, vbTextCompare) = 0 Then
            FallbackKeyValue = Trim(CStr(ws.Cells(r, valCol).Value))
            Exit Function
        End If
        r = r + 1
    Loop
    FallbackKeyValue = ""
End Function
