Attribute VB_Name = "KernelFormula"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelFormula.bas
' Purpose: Named range management, formula placeholder resolution, RowID cache,
'          and shared helpers (ColLetter, GetDataHorizonYears, GetQuarterlyHorizon).
'          Tab-writing logic moved to KernelFormulaWriter.bas (TD-01).
'          Phase 5C module.
' =============================================================================

' RowID lookup cache: Dictionary of "TabName|RowID" -> row number (Long)
' Built on first lookup per tab, avoids repeated linear scans (perf optimization)
Private m_rowIdCache As Object
Private m_rowIdCachedTabs As Object

' GetDataHorizonYears cache (stable within a pipeline run)
Private m_cachedDataHorizonYears As Long
Private m_dataHorizonCached As Boolean


' =============================================================================
' CreateNamedRanges
' Reads named_range_registry from Config sheet. For each entry:
'   1. Resolve RowID to actual row number (scan Column A of target tab)
'   2. Build range address (Single/Quarterly/Row)
'   3. Delete existing name and recreate (idempotent)
'
' MsgBox: "Created [N] named ranges."
' MANUAL BYPASS: "Create named ranges via Formulas > Name Manager.
'   See named_range_registry.csv for the full list."
' =============================================================================
Public Sub CreateNamedRanges(Optional ByVal silent As Boolean = False)
    On Error GoTo ErrHandler

    Dim cnt As Long
    cnt = KernelConfig.GetNamedRangeCount()
    If cnt = 0 Then
        KernelConfig.LogError SEV_INFO, "KernelFormula", "I-810", _
            "No named range entries found. Skipping named range creation.", ""
        If Not silent Then
            KernelFormHelpers.ShowConfigMsgBox "NO_NAMED_RANGES"
        End If
        Exit Sub
    End If

    Dim numYears As Long
    Dim timeHorizon As Long
    timeHorizon = KernelConfig.GetTimeHorizon()
    If timeHorizon <= 0 Then timeHorizon = 12
    Dim writingYearsNR As Long
    writingYearsNR = (timeHorizon \ 3) \ QS_QUARTERS_PER_YEAR
    If writingYearsNR < 1 Then writingYearsNR = 1

    Dim created As Long
    created = 0

    Dim idx As Long
    For idx = 1 To cnt
        Dim rangeName As String
        rangeName = KernelConfig.GetNamedRangeField(idx, NRCFG_COL_NAME)

        Dim tabName As String
        tabName = KernelConfig.GetNamedRangeField(idx, NRCFG_COL_TABNAME)

        Dim rowID As String
        rowID = KernelConfig.GetNamedRangeField(idx, NRCFG_COL_ROWID)

        Dim cellAddr As String
        cellAddr = KernelConfig.GetNamedRangeField(idx, NRCFG_COL_CELLADDR)

        Dim rangeType As String
        rangeType = KernelConfig.GetNamedRangeField(idx, NRCFG_COL_RANGETYPE)

        If Len(rangeName) = 0 Or Len(tabName) = 0 Then GoTo NextRange

        ' Per-tab numYears based on QuarterlyHorizon
        numYears = writingYearsNR
        If StrComp(GetQuarterlyHorizon(tabName), "Data", vbTextCompare) = 0 Then
            Dim dataYearsNR As Long
            dataYearsNR = GetDataHorizonYears()
            If dataYearsNR > 0 Then numYears = dataYearsNR
        End If

        ' Verify tab exists
        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(tabName)
        On Error GoTo ErrHandler
        If ws Is Nothing Then
            KernelConfig.LogError SEV_WARN, "KernelFormula", "W-810", _
                "Tab not found for named range: " & tabName, _
                "MANUAL BYPASS: Create tab '" & tabName & "' then re-run CreateNamedRanges."
            GoTo NextRange
        End If

        ' Resolve row number
        Dim targetRow As Long
        targetRow = 0
        If Len(rowID) > 0 Then
            targetRow = ResolveRowID(tabName, rowID)
            If targetRow = 0 Then
                KernelConfig.LogError SEV_WARN, "KernelFormula", "W-811", _
                    "RowID not found: " & rowID & " on tab " & tabName, _
                    "MANUAL BYPASS: Check Column A of tab '" & tabName & "' for RowID '" & rowID & "'."
                GoTo NextRange
            End If
        End If

        ' Build range address
        Dim refStr As String
        refStr = ""
        Dim quotedTab As String
        quotedTab = "'" & tabName & "'"

        Select Case UCase(rangeType)
            Case "SINGLE"
                If targetRow > 0 And Len(cellAddr) > 0 Then
                    ' RowID resolved row + cellAddr column letter
                    refStr = "=" & quotedTab & "!$" & UCase(cellAddr) & "$" & targetRow
                ElseIf targetRow > 0 Then
                    ' Default to column C
                    refStr = "=" & quotedTab & "!$C$" & targetRow
                ElseIf Len(cellAddr) > 0 Then
                    ' Direct cell address -- make fully absolute (BUG-064)
                    refStr = "=" & quotedTab & "!" & ws.Range(cellAddr).Address(True, True)
                End If

            Case "QUARTERLY"
                If targetRow > 0 Then
                    ' Build non-contiguous range spanning quarterly columns (skip annual totals)
                    Dim qParts As String
                    qParts = ""
                    Dim qyr As Long
                    For qyr = 1 To numYears
                        Dim q1c As Long
                        q1c = QS_DATA_START_COL + (qyr - 1) * QS_COLS_PER_YEAR
                        Dim q4c As Long
                        q4c = q1c + QS_QUARTERS_PER_YEAR - 1
                        If Len(qParts) > 0 Then qParts = qParts & ","
                        qParts = qParts & quotedTab & "!$" & ColLetter(q1c) & "$" & targetRow & _
                                 ":$" & ColLetter(q4c) & "$" & targetRow
                    Next qyr
                    refStr = "=" & qParts
                End If

            Case "ROW"
                If targetRow > 0 Then
                    ' Full row range including annual subtotals
                    Dim lastDataCol As Long
                    lastDataCol = QS_DATA_START_COL + numYears * QS_COLS_PER_YEAR - 1
                    refStr = "=" & quotedTab & "!$" & ColLetter(QS_DATA_START_COL) & "$" & targetRow & _
                             ":$" & ColLetter(lastDataCol) & "$" & targetRow
                End If

            Case Else
                KernelConfig.LogError SEV_WARN, "KernelFormula", "W-812", _
                    "Unknown RangeType: " & rangeType & " for range " & rangeName, ""
                GoTo NextRange
        End Select

        If Len(refStr) = 0 Then GoTo NextRange

        ' Delete existing name if any (idempotent)
        On Error Resume Next
        ThisWorkbook.Names(rangeName).Delete
        On Error GoTo ErrHandler

        ' Add named range
        ThisWorkbook.Names.Add Name:=rangeName, RefersTo:=refStr
        created = created + 1

NextRange:
    Next idx

    KernelConfig.LogError SEV_INFO, "KernelFormula", "I-811", _
        "Created " & created & " named range(s).", ""

    If Not silent Then
        MsgBox "Created " & created & " named range(s).", _
               vbInformation, "RDK -- Named Ranges"
    End If
    Exit Sub

ErrHandler:
    KernelConfig.LogError SEV_ERROR, "KernelFormula", "E-810", _
        "Error creating named ranges: " & Err.Description, _
        "MANUAL BYPASS: Create named ranges via Formulas > Name Manager." & _
        " See named_range_registry.csv for the full list."
    If Not silent Then
        MsgBox "Error creating named ranges: " & Err.Description & vbCrLf & vbCrLf & _
               "MANUAL BYPASS: Create named ranges via Formulas > Name Manager." & vbCrLf & _
               "See named_range_registry.csv for the full list.", _
               vbExclamation, "RDK -- Named Range Error"
    End If
End Sub


' =============================================================================
' ResolveRowID
' Scans Column A of the given tab for the RowID string.
' Returns the row number, or 0 if not found.
' =============================================================================
Public Function ResolveRowID(tabName As String, rowID As String) As Long
    ResolveRowID = 0

    ' Initialize cache dictionaries on first call
    If m_rowIdCache Is Nothing Then
        Set m_rowIdCache = CreateObject("Scripting.Dictionary")
        m_rowIdCache.CompareMode = vbTextCompare
    End If
    If m_rowIdCachedTabs Is Nothing Then
        Set m_rowIdCachedTabs = CreateObject("Scripting.Dictionary")
        m_rowIdCachedTabs.CompareMode = vbTextCompare
    End If

    ' Build cache for this tab if not yet cached
    If Not m_rowIdCachedTabs.Exists(tabName) Then
        On Error Resume Next
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets(tabName)
        On Error GoTo 0
        If ws Is Nothing Then Exit Function

        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, FTAB_COL_ROWID).End(xlUp).row
        If lastRow > 2000 Then lastRow = 2000

        ' Read column A into array for fast scan (PT-001)
        ' BUG-075: Ensure at least 2 rows so VBA returns a 2D array,
        ' not a scalar (single-cell .Value returns scalar, causing Type mismatch)
        Dim colData As Variant
        If lastRow >= 1 Then
            Dim readToRow As Long
            If lastRow < 2 Then readToRow = 2 Else readToRow = lastRow
            colData = ws.Range(ws.Cells(1, FTAB_COL_ROWID), ws.Cells(readToRow, FTAB_COL_ROWID)).Value
        End If

        Dim r As Long
        For r = 1 To lastRow
            Dim cellVal As String
            cellVal = Trim(CStr(colData(r, 1)))
            If Len(cellVal) > 0 Then
                Dim cacheKey As String
                cacheKey = tabName & "|" & cellVal
                If Not m_rowIdCache.Exists(cacheKey) Then
                    m_rowIdCache.Add cacheKey, r
                End If
            End If
        Next r
        m_rowIdCachedTabs.Add tabName, True
    End If

    ' Lookup from cache
    Dim lookupKey As String
    lookupKey = tabName & "|" & rowID
    If m_rowIdCache.Exists(lookupKey) Then
        ResolveRowID = m_rowIdCache(lookupKey)
    End If
End Function


' =============================================================================
' ClearRowIDCache
' Clears the RowID lookup cache. Call when tabs are rebuilt.
' =============================================================================
Public Sub ClearRowIDCache()
    Set m_rowIdCache = Nothing
    Set m_rowIdCachedTabs = Nothing
    m_dataHorizonCached = False
    m_cachedDataHorizonYears = 0
End Sub


' =============================================================================
' ResolveFormulaPlaceholders
' Replaces placeholders in a formula template:
'   {Q} -- column letter for this quarter position
'   {TAB} -- tab name (quoted if contains spaces)
'   {ROWID:xxx} -- cell address of RowID "xxx" on same tab, same column
'   {REF:Tab!RowID} -- cell address of RowID on another tab, same column
'   {NAMED:RangeName} -- the named range reference
'   {PREV_Q:RowID} -- prior quarter column cell for RowID (col-1); resolves to 0 for Q1Y1
' Returns the resolved Excel formula string.
' =============================================================================
Public Function ResolveFormulaPlaceholders(template As String, _
    tabName As String, row As Long, col As Long) As String

    Dim result As String
    Dim pqRowID As String
    Dim pqRow As Long
    result = template

    ' Replace {Q} with current column letter
    result = Replace(result, "{Q}", ColLetter(col))

    ' Replace {TAB} with tab name
    Dim quotedTab As String
    If InStr(1, tabName, " ") > 0 Then
        quotedTab = "'" & tabName & "'"
    Else
        quotedTab = tabName
    End If
    result = Replace(result, "{TAB}", quotedTab)

    ' Replace {ROWID:xxx} -- same tab, same column
    Dim pos As Long
    pos = InStr(1, result, "{ROWID:")
    Do While pos > 0
        Dim endPos As Long
        endPos = InStr(pos, result, "}")
        If endPos = 0 Then Exit Do
        Dim rid As String
        rid = Mid(result, pos + 7, endPos - pos - 7)
        Dim resolvedRow As Long
        resolvedRow = ResolveRowID(tabName, rid)
        Dim cellRef As String
        If resolvedRow > 0 Then
            cellRef = ColLetter(col) & resolvedRow
        Else
            cellRef = "#REF!"
        End If
        result = Left(result, pos - 1) & cellRef & Mid(result, endPos + 1)
        pos = InStr(pos + Len(cellRef), result, "{ROWID:")
    Loop

    ' Replace {REF:Tab!RowID} -- other tab, same column
    pos = InStr(1, result, "{REF:")
    Do While pos > 0
        endPos = InStr(pos, result, "}")
        If endPos = 0 Then Exit Do
        Dim refSpec As String
        refSpec = Mid(result, pos + 5, endPos - pos - 5)
        Dim bangPos As Long
        bangPos = InStr(1, refSpec, "!")
        If bangPos > 0 Then
            Dim refTab As String
            refTab = Left(refSpec, bangPos - 1)
            Dim refRowID As String
            refRowID = Mid(refSpec, bangPos + 1)
            Dim refResolvedRow As Long
            refResolvedRow = ResolveRowID(refTab, refRowID)
            If refResolvedRow > 0 Then
                Dim refQuotedTab As String
                If InStr(1, refTab, " ") > 0 Then
                    refQuotedTab = "'" & refTab & "'"
                Else
                    refQuotedTab = refTab
                End If
                cellRef = refQuotedTab & "!" & ColLetter(col) & refResolvedRow
            Else
                cellRef = "#REF!"
            End If
        Else
            cellRef = "#REF!"
        End If
        result = Left(result, pos - 1) & cellRef & Mid(result, endPos + 1)
        pos = InStr(pos + Len(cellRef), result, "{REF:")
    Loop

    ' Replace {NAMED:RangeName}
    pos = InStr(1, result, "{NAMED:")
    Do While pos > 0
        endPos = InStr(pos, result, "}")
        If endPos = 0 Then Exit Do
        Dim namedRef As String
        namedRef = Mid(result, pos + 7, endPos - pos - 7)
        result = Left(result, pos - 1) & namedRef & Mid(result, endPos + 1)
        pos = InStr(pos + Len(namedRef), result, "{NAMED:")
    Loop

    ' Replace {PREV_Q:RowID} -- prior quarter column, same row
    ' Must skip annual total columns (every QS_COLS_PER_YEAR-th col from start)
    Dim prevCol As Long
    pos = InStr(1, result, "{PREV_Q:")
    Do While pos > 0
        endPos = InStr(pos, result, "}")
        If endPos = 0 Then Exit Do
        pqRowID = Mid(result, pos + 8, endPos - pos - 8)
        pqRow = ResolveRowID(tabName, pqRowID)
        If pqRow > 0 And col > QS_DATA_START_COL Then
            prevCol = col - 1
            ' If prev col is an annual total column, skip to Q4 of prior year
            If prevCol >= QS_DATA_START_COL Then
                If ((prevCol - QS_DATA_START_COL) Mod QS_COLS_PER_YEAR) = QS_QUARTERS_PER_YEAR Then
                    prevCol = prevCol - 1
                End If
            End If
            If prevCol >= QS_DATA_START_COL Then
                cellRef = ColLetter(prevCol) & pqRow
            Else
                cellRef = "0"
            End If
        Else
            cellRef = "0"
        End If
        result = Left(result, pos - 1) & cellRef & Mid(result, endPos + 1)
        pos = InStr(pos + Len(cellRef), result, "{PREV_Q:")
    Loop

    ResolveFormulaPlaceholders = result
End Function


' =============================================================================
' SHARED HELPERS (Public -- called by KernelFormulaWriter)
' =============================================================================


' ColLetter -- converts 1-based column number to letter (1=A, 26=Z, 27=AA)
Public Function ColLetter(colNum As Long) As String
    Dim n As Long
    n = colNum
    ColLetter = ""
    Do While n > 0
        Dim remainder As Long
        remainder = (n - 1) Mod 26
        ColLetter = Chr(65 + remainder) & ColLetter
        n = (n - 1) \ 26
    Loop
End Function


' GetQuarterlyHorizon -- returns "Data" or "Writing" (default) for a tab
Public Function GetQuarterlyHorizon(tabName As String) As String
    GetQuarterlyHorizon = "Writing"
    On Error Resume Next
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    If wsConfig Is Nothing Then
        KernelConfig.LogError SEV_WARN, "KernelFormula", "W-820", _
            "Config sheet not found in GetQuarterlyHorizon, defaulting to Writing", ""
        Exit Function
    End If
    On Error GoTo 0

    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_TAB_REGISTRY)
    If sr = 0 Then
        KernelConfig.LogError SEV_WARN, "KernelFormula", "W-821", _
            "Tab registry section not found in GetQuarterlyHorizon, defaulting to Writing", ""
        Exit Function
    End If

    Dim dr As Long
    dr = sr + 2
    Do While Len(Trim(CStr(wsConfig.Cells(dr, TREG_COL_TABNAME).Value))) > 0
        If StrComp(Trim(CStr(wsConfig.Cells(dr, TREG_COL_TABNAME).Value)), tabName, vbTextCompare) = 0 Then
            Dim hVal As String
            hVal = Trim(CStr(wsConfig.Cells(dr, TREG_COL_QTR_HORIZON).Value))
            If StrComp(hVal, "Data", vbTextCompare) = 0 Then
                GetQuarterlyHorizon = "Data"
            End If
            Exit Function
        End If
        dr = dr + 1
    Loop
End Function


' GetDataHorizonYears -- scans Detail tab for max Year value
' Returns 0 if Detail tab or Year column not found
' Cached within a pipeline run (cleared by ClearRowIDCache)
Public Function GetDataHorizonYears() As Long
    If m_dataHorizonCached Then
        GetDataHorizonYears = m_cachedDataHorizonYears
        Exit Function
    End If

    GetDataHorizonYears = 0
    Dim wsDet As Worksheet
    Set wsDet = Nothing
    On Error Resume Next
    Set wsDet = ThisWorkbook.Sheets(TAB_DETAIL)
    On Error GoTo 0
    If wsDet Is Nothing Then Exit Function

    Dim yrCol As Long
    yrCol = KernelConfig.TryColIndex("Year")
    If yrCol < 1 Then yrCol = KernelConfig.ColIndex("CalYear")
    If yrCol = 0 Then Exit Function

    ' Batch-read year column into array (PT-001) instead of cell-by-cell
    Dim lastRow As Long
    lastRow = wsDet.Cells(wsDet.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Dim yrData As Variant
    yrData = wsDet.Range(wsDet.Cells(2, yrCol), wsDet.Cells(lastRow, yrCol)).Value

    Dim maxYr As Long
    maxYr = 0
    Dim r As Long
    For r = 1 To UBound(yrData, 1)
        If IsNumeric(yrData(r, 1)) Then
            If CLng(yrData(r, 1)) > maxYr Then maxYr = CLng(yrData(r, 1))
        End If
    Next r
    GetDataHorizonYears = maxYr

    m_cachedDataHorizonYears = maxYr
    m_dataHorizonCached = True
End Function
