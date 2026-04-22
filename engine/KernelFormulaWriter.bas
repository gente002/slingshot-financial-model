Attribute VB_Name = "KernelFormulaWriter"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelFormulaWriter.bas
' Purpose: Tab-writing logic split from KernelFormula.bas (TD-01).
'          CreateFormulaTabs, RefreshFormulaTabs, RefreshFormulaTabsUI,
'          and all Private helpers for tab generation.
'          Calls KernelFormula for shared helpers (ColLetter, ResolveRowID,
'          ResolveFormulaPlaceholders, GetDataHorizonYears, GetQuarterlyHorizon).
' =============================================================================

' Module-level flag to suppress MsgBox during pipeline calls
Private m_silent As Boolean

' Phase 11A: When True, skip ClearContents to preserve user-entered data on input tabs
Private m_preserveOnRefresh As Boolean


' =============================================================================
' CreateFormulaTabs
' Reads formula_tab_config from Config sheet. For each unique TabName:
'   1. EnsureSheet
'   2. Write RowIDs in Column A (hidden)
'   3. Write labels and formulas
'   4. If QuarterlyColumns=TRUE: generate quarterly column headers
'   5. Apply formatting (fonts, fills, borders, indentation, comments)
'   6. AutoFit columns, freeze panes
'
' MsgBox: "Created [N] formula-driven tabs with [M] formulas."
' MANUAL BYPASS: "Create tabs manually. Add formulas referencing
'   the named ranges listed in named_range_registry.csv."
' =============================================================================
Public Sub CreateFormulaTabs(Optional silent As Boolean = False)
    On Error GoTo ErrHandler
    m_silent = m_silent Or silent

    Dim cnt As Long
    cnt = KernelConfig.GetFormulaTabConfigCount()
    If cnt = 0 Then
        KernelConfig.LogError SEV_INFO, "KernelFormula", "I-800", _
            "No formula tab config entries found. Skipping formula tab creation.", ""
        If Not m_silent Then
            KernelFormHelpers.ShowConfigMsgBox "NO_FORMULA_CONFIG"
        End If
        Exit Sub
    End If

    ' Clear RowID cache so it rebuilds from fresh tab data
    KernelFormula.ClearRowIDCache

    ' Save/restore ScreenUpdating state so pipeline callers stay suppressed (BUG-058)
    Dim savedScreenUpdating As Boolean
    savedScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    ' Pre-index: build tab -> config row indices map (eliminates O(n^2) scan)
    Dim tabIndex As Object
    Set tabIndex = CreateObject("Scripting.Dictionary")
    tabIndex.CompareMode = vbTextCompare

    Dim idx As Long
    For idx = 1 To cnt
        Dim tn As String
        tn = KernelConfig.GetFormulaTabConfigField(idx, FTCFG_COL_TABNAME)
        If Len(tn) > 0 Then
            If Not tabIndex.Exists(tn) Then
                tabIndex.Add tn, Array(idx)
            Else
                Dim existing As Variant
                existing = tabIndex(tn)
                Dim newArr() As Variant
                ReDim newArr(0 To UBound(existing) + 1)
                Dim ci As Long
                For ci = 0 To UBound(existing)
                    newArr(ci) = existing(ci)
                Next ci
                newArr(UBound(existing) + 1) = idx
                tabIndex(tn) = newArr
            End If
        End If
    Next idx

    ' Collect unique tab names in order
    Dim tabNames() As String
    Dim tabCount As Long
    tabCount = 0
    ReDim tabNames(1 To cnt)
    For idx = 1 To cnt
        tn = KernelConfig.GetFormulaTabConfigField(idx, FTCFG_COL_TABNAME)
        If Len(tn) > 0 And Not IsInArray(tn, tabNames, tabCount) Then
            tabCount = tabCount + 1
            tabNames(tabCount) = tn
        End If
    Next idx

    If tabCount = 0 Then
        KernelConfig.LogError SEV_INFO, "KernelFormulaWriter", "I-802", _
            "No unique tab names found in formula tab config after filtering.", ""
        Application.ScreenUpdating = savedScreenUpdating
        Exit Sub
    End If

    Dim totalFormulas As Long
    totalFormulas = 0

    ' Determine quarterly layout parameters
    Dim numYears As Long
    Dim numQuarters As Long
    Dim timeHorizon As Long
    timeHorizon = KernelConfig.GetTimeHorizon()
    If timeHorizon <= 0 Then timeHorizon = 12
    numQuarters = timeHorizon \ 3
    If numQuarters < 1 Then numQuarters = 1
    Dim writingYears As Long
    writingYears = numQuarters \ QS_QUARTERS_PER_YEAR
    If writingYears < 1 Then writingYears = 1
    If numQuarters > writingYears * QS_QUARTERS_PER_YEAR Then
        writingYears = writingYears + 1
    End If

    ' GLOBAL pre-write: write ALL RowIDs across ALL tabs before any formula
    ' resolution. This ensures cross-tab {REF:} references can find RowIDs
    ' on tabs that haven't been processed yet (BUG-182).
    '
    ' BUG-191: On Output-category tabs in preserve mode, stale legacy cells
    ' (from a prior CSV layout) survive because preserve mode skips
    ' ClearContents. Fix: always clear Output tabs, but first capture their
    ' preserved-cells values to an in-memory dict and restore at the end of
    ' the build. Input tabs keep their existing preserve-mode behavior
    ' (user-entered data preserved without capture/restore).
    Dim preservedCellDicts As Object
    Set preservedCellDicts = CreateObject("Scripting.Dictionary")
    Dim preTabIdx As Long
    For preTabIdx = 1 To tabCount
        Dim preTabName As String: preTabName = tabNames(preTabIdx)
        Dim preWs As Worksheet
        Set preWs = EnsureSheetFormula(preTabName)

        ' BUG-191: determine if this is an Output-category tab
        Dim preTabCategory As String
        preTabCategory = KernelWorkspaceExt.GetTabCategory(preTabName)
        Dim preIsOutputTab As Boolean
        preIsOutputTab = (StrComp(preTabCategory, "Output", vbTextCompare) = 0)

        ' BUG-191: if in preserve mode AND this is an Output tab, capture
        ' preserved-cells values to memory BEFORE clearing.
        If m_preserveOnRefresh And preIsOutputTab Then
            Dim preCapDict As Object
            Set preCapDict = KernelWorkspaceExt.CapturePreservedCellsInMemory(preTabName)
            If Not preCapDict Is Nothing Then
                If preCapDict.Count > 0 Then
                    preservedCellDicts(preTabName) = preCapDict
                End If
            End If
        End If

        ' BUG-191: clear if NOT in preserve mode, OR if this is an Output
        ' tab in preserve mode (we captured user overrides above).
        If (Not m_preserveOnRefresh) Or preIsOutputTab Then
            preWs.Cells.ClearContents
        End If

        Dim preRows As Variant: preRows = tabIndex(preTabName)
        Dim preTri As Long
        For preTri = LBound(preRows) To UBound(preRows)
            Dim preCfgIdx As Long: preCfgIdx = CLng(preRows(preTri))
            Dim gRowID As String
            gRowID = KernelConfig.GetFormulaTabConfigField(preCfgIdx, FTCFG_COL_ROWID)
            If Len(gRowID) > 0 Then
                Dim gRowStr As String
                gRowStr = KernelConfig.GetFormulaTabConfigField(preCfgIdx, FTCFG_COL_ROW)
                If IsNumeric(gRowStr) And Len(gRowStr) > 0 Then
                    preWs.Cells(CLng(gRowStr), FTAB_COL_ROWID).Value = gRowID
                End If
            End If
        Next preTri
    Next preTabIdx
    KernelFormula.ClearRowIDCache

    ' Process each unique tab
    Dim tabIdx As Long
    For tabIdx = 1 To tabCount
        Dim tabName As String
        tabName = tabNames(tabIdx)

        ' EnsureSheet (already created + cleared in global pre-write)
        Dim ws As Worksheet
        Set ws = EnsureSheetFormula(tabName)

        ' Check if this tab has QuarterlyColumns=TRUE
        Dim hasQuarterly As Boolean
        hasQuarterly = HasQuarterlyColumns(tabName)

        ' Check if this tab has GrandTotal=TRUE (Y1-Yn sum column)
        Dim wantsGrandTotal As Boolean
        wantsGrandTotal = hasQuarterly And HasGrandTotal(tabName)

        ' Tail column: read HasTailCol flag from tab_registry (config-driven)
        Dim wantsTail As Boolean
        wantsTail = GetTabRegistryFlag(tabName, TREG_COL_HASTAILCOL)

        ' Per-tab numYears: Data horizon scans Detail for max Year,
        ' Writing horizon (default) uses TimeHorizon
        numYears = writingYears
        If hasQuarterly And StrComp(KernelFormula.GetQuarterlyHorizon(tabName), "Data", vbTextCompare) = 0 Then
            Dim dataYears As Long
            dataYears = KernelFormula.GetDataHorizonYears()
            If dataYears > 0 Then numYears = dataYears
        End If

        ' Process rows for this tab using pre-indexed map (O(n) not O(n^2))
        Dim tabRows As Variant
        tabRows = tabIndex(tabName)
        Dim cfgIdx As Long
        Dim tri As Long

        ' RowIDs already pre-written by global pass above (BUG-176/BUG-182)
        For tri = LBound(tabRows) To UBound(tabRows)
            cfgIdx = CLng(tabRows(tri))

            ' BUG-115: Resume Next within config row loop so one error
            ' does not abort all formula tab creation
            On Error Resume Next

            Dim rowID As String
            rowID = KernelConfig.GetFormulaTabConfigField(cfgIdx, FTCFG_COL_ROWID)

            Dim cfgRow As Long
            cfgRow = 0
            Dim rowStr As String
            rowStr = KernelConfig.GetFormulaTabConfigField(cfgIdx, FTCFG_COL_ROW)
            If IsNumeric(rowStr) And Len(rowStr) > 0 Then cfgRow = CLng(rowStr)

            Dim colStr As String
            colStr = KernelConfig.GetFormulaTabConfigField(cfgIdx, FTCFG_COL_COL)

            Dim cellType As String
            cellType = KernelConfig.GetFormulaTabConfigField(cfgIdx, FTCFG_COL_CELLTYPE)

            Dim content As String
            content = KernelConfig.GetFormulaTabConfigField(cfgIdx, FTCFG_COL_CONTENT)

            Dim fmt As String
            fmt = KernelConfig.GetFormulaTabConfigField(cfgIdx, FTCFG_COL_FORMAT)

            Dim fontStyle As String
            fontStyle = KernelConfig.GetFormulaTabConfigField(cfgIdx, FTCFG_COL_FONTSTYLE)

            Dim fillColor As String
            fillColor = KernelConfig.GetFormulaTabConfigField(cfgIdx, FTCFG_COL_FILLCOLOR)

            Dim fontColor As String
            fontColor = KernelConfig.GetFormulaTabConfigField(cfgIdx, FTCFG_COL_FONTCOLOR)

            Dim mergeSpan As Long
            Dim csStr As String
            csStr = KernelConfig.GetFormulaTabConfigField(cfgIdx, FTCFG_COL_COLSPAN)
            mergeSpan = 1
            If IsNumeric(csStr) And Len(csStr) > 0 Then
                If CLng(csStr) > 1 Then mergeSpan = CLng(csStr)
            End If

            Dim borderBot As String
            borderBot = KernelConfig.GetFormulaTabConfigField(cfgIdx, FTCFG_COL_BORDERBOTTOM)

            Dim borderTop As String
            borderTop = KernelConfig.GetFormulaTabConfigField(cfgIdx, FTCFG_COL_BORDERTOP)

            Dim indentLvl As Long
            Dim indStr As String
            indStr = KernelConfig.GetFormulaTabConfigField(cfgIdx, FTCFG_COL_INDENT)
            indentLvl = 0
            If IsNumeric(indStr) And Len(indStr) > 0 Then indentLvl = CLng(indStr)

            Dim commentText As String
            commentText = KernelConfig.GetFormulaTabConfigField(cfgIdx, FTCFG_COL_COMMENT)

            Dim hAlign As String
            hAlign = KernelConfig.GetFormulaTabConfigField(cfgIdx, FTCFG_COL_HALIGN)

            If cfgRow = 0 Then GoTo NextCfgRow

            ' Determine target column
            Dim colNum As Long
            Dim isQuarterlyFormula As Boolean
            isQuarterlyFormula = False
            Dim isQuarterlyInput As Boolean
            isQuarterlyInput = False

            If IsNumeric(colStr) Then
                colNum = CLng(colStr)
            ElseIf Len(colStr) > 0 Then
                colNum = ColLetterToNum(colStr)
                If hasQuarterly Then
                    If StrComp(cellType, "Formula", vbTextCompare) = 0 Then
                        isQuarterlyFormula = True
                    ElseIf StrComp(cellType, "Input", vbTextCompare) = 0 Then
                        isQuarterlyInput = True
                    End If
                End If
            Else
                colNum = FTAB_COL_ROWID
            End If

            ' Write RowID in Column A (always)
            If Len(rowID) > 0 Then
                ws.Cells(cfgRow, FTAB_COL_ROWID).Value = rowID
            End If

            Select Case UCase(cellType)
                Case "LABEL"
                    ws.Cells(cfgRow, colNum).Value = content
                    ApplyCellFormatting ws, cfgRow, colNum, fontStyle, fillColor, _
                        fontColor, mergeSpan, borderBot, borderTop, indentLvl, commentText, fmt, hAlign

                Case "SECTION"
                    ' Default navy fill + white bold unless overridden
                    If Len(fillColor) = 0 Then fillColor = "1F3864"
                    If Len(fontColor) = 0 Then fontColor = "FFFFFF"
                    If Len(fontStyle) = 0 Then fontStyle = "Bold"
                    ws.Cells(cfgRow, colNum).NumberFormat = "@"
                    ws.Cells(cfgRow, colNum).Value = content
                    ApplyCellFormatting ws, cfgRow, colNum, fontStyle, fillColor, _
                        fontColor, mergeSpan, borderBot, borderTop, indentLvl, commentText, fmt, hAlign

                Case "INPUT"
                    ' Blue font for input cells
                    If Len(fontColor) = 0 Then fontColor = "0000FF"
                    If isQuarterlyInput Then
                        ' BUG-073: Replicate default value across all quarterly columns
                        Dim iYr As Long
                        For iYr = 1 To numYears
                            Dim iQtr As Long
                            For iQtr = 1 To QS_QUARTERS_PER_YEAR
                                Dim iCol As Long
                                iCol = QS_DATA_START_COL + (iYr - 1) * QS_COLS_PER_YEAR + (iQtr - 1)
                                If Len(content) > 0 And Not m_preserveOnRefresh Then
                                    If IsNumeric(content) Then
                                        ws.Cells(cfgRow, iCol).Value = CDbl(content)
                                    Else
                                        ws.Cells(cfgRow, iCol).Value = content
                                    End If
                                End If
                            Next iQtr
                            ' Annual total column: SUM(Q1:Q4) or Q4 for balance items
                            Dim iAnnCol As Long
                            iAnnCol = QS_DATA_START_COL + (iYr - 1) * QS_COLS_PER_YEAR + QS_QUARTERS_PER_YEAR
                            If IsNumeric(content) Or InStr(1, fmt, "#") > 0 Or InStr(1, fmt, "0") > 0 Then
                                Dim iQ1Col As Long
                                iQ1Col = QS_DATA_START_COL + (iYr - 1) * QS_COLS_PER_YEAR
                                Dim iQ4Col As Long
                                iQ4Col = iQ1Col + QS_QUARTERS_PER_YEAR - 1
                                ' Check BalanceItem flag for Input rows (e.g., headcount)
                                Dim iBiVal As String
                                iBiVal = KernelConfig.GetFormulaTabConfigField(cfgIdx, FTCFG_COL_BALANCEITEM)
                                If StrComp(iBiVal, "TRUE", vbTextCompare) = 0 Then
                                    ws.Cells(cfgRow, iAnnCol).formula = "=" & _
                                        KernelFormula.ColLetter(iQ4Col) & cfgRow
                                Else
                                    ws.Cells(cfgRow, iAnnCol).formula = "=SUM(" & _
                                        KernelFormula.ColLetter(iQ1Col) & cfgRow & ":" & KernelFormula.ColLetter(iQ4Col) & cfgRow & ")"
                                End If
                            End If
                            ws.Cells(cfgRow, iAnnCol).Interior.Color = RGB(217, 217, 217)
                        Next iYr
                        ' Batch format: apply NumberFormat and Font.Color to entire data range at once
                        Dim inFirstCol As Long
                        inFirstCol = QS_DATA_START_COL
                        Dim inLastCol As Long
                        inLastCol = QS_DATA_START_COL + numYears * QS_COLS_PER_YEAR - 1
                        Dim inDataRng As Range
                        Set inDataRng = ws.Range(ws.Cells(cfgRow, inFirstCol), ws.Cells(cfgRow, inLastCol))
                        If Len(fmt) > 0 Then inDataRng.NumberFormat = fmt
                        If Len(fontColor) = 6 Then inDataRng.Font.Color = HexToRGB(fontColor)
                        ' Apply label column formatting
                        ApplyCellFormatting ws, cfgRow, 2, fontStyle, "", _
                            "", 1, borderBot, borderTop, 0, "", "", hAlign
                        ' Grand total column: SUM of annual totals for numeric inputs
                        If wantsGrandTotal Then
                            Dim inGtCol As Long
                            inGtCol = QS_DATA_START_COL + numYears * QS_COLS_PER_YEAR
                            If IsNumeric(content) Or InStr(1, fmt, "#") > 0 Or InStr(1, fmt, "0") > 0 Then
                                Dim inGtFormula As String
                                inGtFormula = "=SUM("
                                Dim inGtYr As Long
                                For inGtYr = 1 To numYears
                                    Dim inAnnTotCol As Long
                                    inAnnTotCol = QS_DATA_START_COL + (inGtYr - 1) * QS_COLS_PER_YEAR + QS_QUARTERS_PER_YEAR
                                    If inGtYr > 1 Then inGtFormula = inGtFormula & ","
                                    inGtFormula = inGtFormula & KernelFormula.ColLetter(inAnnTotCol) & cfgRow
                                Next inGtYr
                                inGtFormula = inGtFormula & ")"
                                WriteFormula ws, cfgRow, inGtCol, inGtFormula
                            End If
                            If Len(fmt) > 0 Then ws.Cells(cfgRow, inGtCol).NumberFormat = fmt
                            ws.Cells(cfgRow, inGtCol).Interior.Color = RGB(189, 215, 238)
                            ws.Cells(cfgRow, inGtCol).Font.Bold = True
                        End If
                    Else
                        If Len(content) > 0 And Not m_preserveOnRefresh Then
                            If IsNumeric(content) Then
                                ws.Cells(cfgRow, colNum).Value = CDbl(content)
                            Else
                                ws.Cells(cfgRow, colNum).Value = content
                            End If
                        End If
                        ApplyCellFormatting ws, cfgRow, colNum, fontStyle, fillColor, _
                            fontColor, mergeSpan, borderBot, borderTop, indentLvl, commentText, fmt, hAlign
                    End If

                Case "FORMULA"
                    If isQuarterlyFormula Then
                        ' Replicate formula across quarterly columns
                        Dim yr As Long
                        For yr = 1 To numYears
                            Dim qtr As Long
                            For qtr = 1 To QS_QUARTERS_PER_YEAR
                                Dim qColNum As Long
                                qColNum = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + (qtr - 1)
                                Dim resolved As String
                                resolved = KernelFormula.ResolveFormulaPlaceholders(content, tabName, cfgRow, qColNum)
                                WriteFormula ws, cfgRow, qColNum, resolved
                                totalFormulas = totalFormulas + 1
                            Next qtr
                            ' Annual total column
                            Dim annCol As Long
                            annCol = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + QS_QUARTERS_PER_YEAR

                            ' Point-in-time tabs: annual totals left blank (config-driven)
                            Dim skipAnnual As Boolean
                            skipAnnual = GetTabRegistryFlag(tabName, TREG_COL_SKIPANNUAL)

                            If Not skipAnnual Then
                                Dim annFormula As String

                                ' Check BalanceItem flag from config (replaces PREV_Q heuristic)
                                Dim isBalanceItem As Boolean
                                Dim biVal As String
                                biVal = KernelConfig.GetFormulaTabConfigField(cfgIdx, FTCFG_COL_BALANCEITEM)
                                isBalanceItem = (StrComp(biVal, "TRUE", vbTextCompare) = 0)

                                If isBalanceItem Then
                                    ' Balance: annual = Q4 value (EOP)
                                    Dim fQ4Col As Long
                                    fQ4Col = annCol - 1
                                    annFormula = "=" & KernelFormula.ColLetter(fQ4Col) & cfgRow
                                ElseIf InStr(1, content, "{PREV_Q:", vbTextCompare) > 0 And _
                                       InStr(1, UCase(content), "IFERROR", vbTextCompare) = 0 Then
                                    ' Cross-ref {PREV_Q:} without IFERROR: Flow -- SUM(Q1:Q4)
                                    Dim fQ1Col As Long
                                    fQ1Col = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR
                                    annFormula = "=SUM(" & KernelFormula.ColLetter(fQ1Col) & cfgRow & ":" & _
                                        KernelFormula.ColLetter(fQ1Col + QS_QUARTERS_PER_YEAR - 1) & cfgRow & ")"
                                ElseIf InStr(1, UCase(content), "IFERROR", vbTextCompare) > 0 Then
                                    ' BUG-102: Ratio formulas (IFERROR): resolve at annual
                                    ' column so annual = totalA/totalB (weighted ratio)
                                    annFormula = KernelFormula.ResolveFormulaPlaceholders(content, tabName, cfgRow, annCol)
                                Else
                                    ' BUG-102: All other {ROWID:} formulas: SUM(Q1:Q4).
                                    ' Resolving multiplicative formulas at annual column
                                    ' multiplies summed operands (wrong). SUM of correct
                                    ' quarterly values is always right for non-ratio formulas.
                                    fQ1Col = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR
                                    annFormula = "=SUM(" & KernelFormula.ColLetter(fQ1Col) & cfgRow & ":" & _
                                        KernelFormula.ColLetter(fQ1Col + QS_QUARTERS_PER_YEAR - 1) & cfgRow & ")"
                                End If
                                WriteFormula ws, cfgRow, annCol, annFormula
                            End If
                            ' Grey shading for annual total (even if blank)
                            ws.Cells(cfgRow, annCol).Interior.Color = RGB(217, 217, 217)
                        Next yr

                        ' Batch format: apply NumberFormat to entire quarterly data range at once
                        If Len(fmt) > 0 Then
                            Dim fmFirstCol As Long
                            fmFirstCol = QS_DATA_START_COL
                            Dim fmLastCol As Long
                            fmLastCol = QS_DATA_START_COL + numYears * QS_COLS_PER_YEAR - 1
                            ws.Range(ws.Cells(cfgRow, fmFirstCol), ws.Cells(cfgRow, fmLastCol)).NumberFormat = fmt
                        End If

                        ' Grand total column: sums annual totals (Y1..Yn)
                        ' Balance items: GT = last year's annual (EOP), not SUM
                        If wantsGrandTotal And Not skipAnnual Then
                            Dim gtCol As Long
                            gtCol = QS_DATA_START_COL + numYears * QS_COLS_PER_YEAR

                            If isBalanceItem Then
                                ' Balance: GT = last year's annual column (EOP)
                                Dim gtLastAnnCol As Long
                                gtLastAnnCol = QS_DATA_START_COL + (numYears - 1) * QS_COLS_PER_YEAR + QS_QUARTERS_PER_YEAR
                                WriteFormula ws, cfgRow, gtCol, "=" & KernelFormula.ColLetter(gtLastAnnCol) & cfgRow

                            Else
                            Dim hasPlaceholder As Boolean
                            hasPlaceholder = (InStr(1, content, "{REF:") > 0 Or _
                                InStr(1, content, "{ROWID:") > 0 Or _
                                InStr(1, content, "{PREV_Q:") > 0)

                            If hasPlaceholder And _
                               (InStr(1, content, "{ROWID:") > 0 Or InStr(1, content, "{PREV_Q:") > 0) Then
                                ' BUG-102: Split into ratio vs non-ratio paths.
                                ' Ratios and {PREV_Q:} formulas: resolve at GT column.
                                ' All others: SUM of annual totals (safe for additive
                                ' and multiplicative formulas alike).
                                Dim gtNeedsResolve As Boolean
                                gtNeedsResolve = (InStr(1, content, "{PREV_Q:", vbTextCompare) > 0) Or _
                                    (InStr(1, UCase(content), "IFERROR", vbTextCompare) > 0)
                                If gtNeedsResolve Then
                                    Dim gtResolved As String
                                    gtResolved = KernelFormula.ResolveFormulaPlaceholders(content, tabName, cfgRow, gtCol)
                                    WriteFormula ws, cfgRow, gtCol, gtResolved
                                Else
                                    Dim gtSumFormula As String
                                    gtSumFormula = "=SUM("
                                    Dim gtSumYr As Long
                                    For gtSumYr = 1 To numYears
                                        Dim gtSumAnnCol As Long
                                        gtSumAnnCol = QS_DATA_START_COL + (gtSumYr - 1) * QS_COLS_PER_YEAR + QS_QUARTERS_PER_YEAR
                                        If gtSumYr > 1 Then gtSumFormula = gtSumFormula & ","
                                        gtSumFormula = gtSumFormula & KernelFormula.ColLetter(gtSumAnnCol) & cfgRow
                                    Next gtSumYr
                                    gtSumFormula = gtSumFormula & ")"
                                    WriteFormula ws, cfgRow, gtCol, gtSumFormula
                                End If
                            Else
                                ' Pure cross-tab ref OR standalone formula (e.g. COLUMN()-based):
                                ' SUM of annual total columns
                                Dim gtFormula As String
                                gtFormula = "=SUM("
                                Dim gtYr As Long
                                For gtYr = 1 To numYears
                                    Dim annTotCol As Long
                                    annTotCol = QS_DATA_START_COL + (gtYr - 1) * QS_COLS_PER_YEAR + QS_QUARTERS_PER_YEAR
                                    If gtYr > 1 Then gtFormula = gtFormula & ","
                                    gtFormula = gtFormula & KernelFormula.ColLetter(annTotCol) & cfgRow
                                Next gtYr
                                gtFormula = gtFormula & ")"
                                WriteFormula ws, cfgRow, gtCol, gtFormula
                            End If
                            End If ' End Else (non-balance) block
                            If Len(fmt) > 0 Then ws.Cells(cfgRow, gtCol).NumberFormat = fmt
                            ws.Cells(cfgRow, gtCol).Interior.Color = RGB(189, 215, 238)
                            ws.Cells(cfgRow, gtCol).Font.Bold = True
                            totalFormulas = totalFormulas + 1

                            ' Tail column: development beyond writing horizon
                            Dim qsDataYears As Long
                            qsDataYears = KernelFormula.GetDataHorizonYears()
                            If qsDataYears > numYears And wantsTail Then
                                Dim fmTailCol As Long
                                fmTailCol = gtCol + 1
                                Dim qsTailCol As Long
                                qsTailCol = QS_DATA_START_COL + qsDataYears * QS_COLS_PER_YEAR
                                If InStr(1, content, "{REF:") > 0 And InStr(1, content, "{ROWID:") = 0 Then
                                    ' Cross-tab reference: point to QS Tail column
                                    Dim refParts() As String
                                    Dim tailRefTab As String
                                    Dim tailRefRowID As String
                                    Dim tailRefRow As Long
                                    ' Parse {REF:Tab!RowID}
                                    Dim refStart As Long
                                    refStart = InStr(1, content, "{REF:")
                                    Dim refEnd As Long
                                    refEnd = InStr(refStart, content, "}")
                                    Dim refBody As String
                                    refBody = Mid(content, refStart + 5, refEnd - refStart - 5)
                                    Dim bangPos As Long
                                    bangPos = InStr(1, refBody, "!")
                                    If bangPos > 0 Then
                                        tailRefTab = Left(refBody, bangPos - 1)
                                        tailRefRowID = Mid(refBody, bangPos + 1)
                                        tailRefRow = KernelFormula.ResolveRowID(tailRefTab, tailRefRowID)
                                        If tailRefRow > 0 Then
                                            Dim qsTailLtr As String
                                            qsTailLtr = KernelFormula.ColLetter(qsTailCol)
                                            ' BUG-101: Preserve prefix (e.g. =-) and suffix
                                            ' (e.g. *0.5) around {REF:} so PD-05 negation
                                            ' and multipliers are retained in Tail column.
                                            Dim tailPrefix As String
                                            tailPrefix = Left(content, refStart - 1)
                                            Dim tailSuffix As String
                                            tailSuffix = Mid(content, refEnd + 1)
                                            Dim tailQuotedTab As String
                                            If InStr(1, tailRefTab, " ") > 0 Then
                                                tailQuotedTab = "'" & tailRefTab & "'"
                                            Else
                                                tailQuotedTab = tailRefTab
                                            End If
                                            Dim tailFormula As String
                                            tailFormula = tailPrefix & tailQuotedTab & "!" & _
                                                qsTailLtr & tailRefRow & tailSuffix
                                            WriteFormula ws, cfgRow, fmTailCol, tailFormula
                                        End If
                                    End If
                                Else
                                    ' Same-tab {ROWID:} formula: resolve at tail column
                                    Dim tailResolved As String
                                    tailResolved = KernelFormula.ResolveFormulaPlaceholders(content, tabName, cfgRow, fmTailCol)
                                    WriteFormula ws, cfgRow, fmTailCol, tailResolved
                                End If
                                If Len(fmt) > 0 Then ws.Cells(cfgRow, fmTailCol).NumberFormat = fmt
                                ws.Cells(cfgRow, fmTailCol).Interior.Color = RGB(198, 224, 180)
                                ws.Cells(cfgRow, fmTailCol).Font.Bold = True
                                totalFormulas = totalFormulas + 1
                            End If
                        End If
                        ' Apply formatting to the label column
                        ApplyCellFormatting ws, cfgRow, 2, fontStyle, fillColor, _
                            fontColor, 1, borderBot, borderTop, 0, "", "", hAlign
                        ' Apply row-level formatting (fill, bold, borders) to all data columns
                        If Len(fillColor) > 0 Or Len(fontStyle) > 0 Then
                            fmFirstCol = QS_DATA_START_COL
                            fmLastCol = QS_DATA_START_COL + numYears * QS_COLS_PER_YEAR - 1
                            If wantsTail Then fmLastCol = fmLastCol + 1
                            Dim fmDataRange As Range
                            Set fmDataRange = ws.Range(ws.Cells(cfgRow, fmFirstCol), ws.Cells(cfgRow, fmLastCol))
                            If Len(fillColor) > 0 Then fmDataRange.Interior.Color = HexToRGB(fillColor)
                            If InStr(1, fontStyle, "Bold", vbTextCompare) > 0 Then fmDataRange.Font.Bold = True
                            If Len(borderBot) > 0 Then
                                Select Case LCase(borderBot)
                                    Case "thin": fmDataRange.Borders(xlEdgeBottom).LineStyle = xlContinuous: fmDataRange.Borders(xlEdgeBottom).Weight = xlThin
                                    Case "double": fmDataRange.Borders(xlEdgeBottom).LineStyle = xlDouble
                                End Select
                            End If
                            If Len(borderTop) > 0 Then
                                Select Case LCase(borderTop)
                                    Case "thin": fmDataRange.Borders(xlEdgeTop).LineStyle = xlContinuous: fmDataRange.Borders(xlEdgeTop).Weight = xlThin
                                    Case "double": fmDataRange.Borders(xlEdgeTop).LineStyle = xlDouble
                                End Select
                            End If
                        End If
                    Else
                        ' Fixed position formula
                        Dim fixRes As String
                        fixRes = KernelFormula.ResolveFormulaPlaceholders(content, tabName, cfgRow, colNum)
                        WriteFormula ws, cfgRow, colNum, fixRes
                        If Len(fmt) > 0 Then ws.Cells(cfgRow, colNum).NumberFormat = fmt
                        ApplyCellFormatting ws, cfgRow, colNum, fontStyle, fillColor, _
                            fontColor, mergeSpan, borderBot, borderTop, indentLvl, commentText, "", hAlign
                        totalFormulas = totalFormulas + 1
                    End If

                Case "SPACER"
                    ' Blank row -- nothing to write

            End Select

NextCfgRow:
            ' BUG-115: Per-row error recovery so one bad cell
            ' does not abort all formula tab creation
            If Err.Number <> 0 Then
                KernelConfig.LogError SEV_WARN, "KernelFormula", "W-800", _
                    "Skipped config row: " & Err.Description, _
                    "Tab=" & tabName & " RowID=" & rowID
                Err.Clear
            End If
        Next tri
        On Error GoTo ErrHandler

        ' Write quarterly column headers if applicable (after config rows so merges are done)
        If hasQuarterly Then
            ' Find the first formula row for this tab
            Dim hdrRow As Long
            hdrRow = 0
            Dim scanIdx As Long
            ' Use pre-indexed rows (no full-scan needed)
            For tri = LBound(tabRows) To UBound(tabRows)
                scanIdx = CLng(tabRows(tri))
                ' BUG-073: Include quarterly Input cells (non-numeric Col)
                Dim scanCellType As String
                scanCellType = KernelConfig.GetFormulaTabConfigField(scanIdx, FTCFG_COL_CELLTYPE)
                Dim scanColStr As String
                scanColStr = KernelConfig.GetFormulaTabConfigField(scanIdx, FTCFG_COL_COL)
                Dim isQtrCell As Boolean
                isQtrCell = ((StrComp(scanCellType, "Formula", vbTextCompare) = 0) Or _
                            (StrComp(scanCellType, "Input", vbTextCompare) = 0)) And _
                            Len(scanColStr) > 0 And Not IsNumeric(scanColStr)
                If isQtrCell Then
                    Dim hdrRowStr As String
                    hdrRowStr = KernelConfig.GetFormulaTabConfigField(scanIdx, FTCFG_COL_ROW)
                    If IsNumeric(hdrRowStr) And Len(hdrRowStr) > 0 Then
                        Dim hdrRowNum As Long
                        hdrRowNum = CLng(hdrRowStr)
                        If hdrRow = 0 Or hdrRowNum < hdrRow Then hdrRow = hdrRowNum
                    End If
                End If
            Next tri
            ' Write headers one row above the first quarterly data row
            If hdrRow > 1 Then
                WriteQuarterlyHeaders ws, numYears, hdrRow - 1, wantsGrandTotal, wantsTail
            End If
        End If

        ' AutoFit visible columns (skip during pipeline -- AutoFitAllOutputTabs handles it)
        If Not m_silent Then
            On Error Resume Next
            ws.Columns.AutoFit
            On Error GoTo ErrHandler
        End If

        ' Hide Column A (RowID column)
        ws.Columns(FTAB_COL_ROWID).Hidden = True

        ' Freeze panes: row-only for non-quarterly tabs (BUG-066),
        ' row+col for quarterly tabs that need horizontal scrolling
        ws.Activate
        If hasQuarterly Then
            ws.Cells(3, QS_DATA_START_COL).Select
        Else
            ws.Cells(3, 1).Select
        End If
        ActiveWindow.FreezePanes = False
        ActiveWindow.FreezePanes = True
    Next tabIdx

    ' BUG-191: restore preserved-cells values for Output tabs that had
    ' their user-overrides captured to memory before the clear. This
    ' writes captured values AFTER the formula-writing pass so the user
    ' overrides overwrite the seeded formula defaults.
    If Not preservedCellDicts Is Nothing Then
        If preservedCellDicts.Count > 0 Then
            Dim restTabKey As Variant
            Dim restTotal As Long: restTotal = 0
            For Each restTabKey In preservedCellDicts.Keys
                Dim restDict As Object
                Set restDict = preservedCellDicts(restTabKey)
                If Not restDict Is Nothing Then
                    KernelWorkspaceExt.RestorePreservedCellsInMemory CStr(restTabKey), restDict
                    restTotal = restTotal + restDict.Count
                End If
            Next restTabKey
            KernelConfig.LogError SEV_INFO, "KernelFormula", "I-803", _
                "BUG-191 restore: " & restTotal & " preserved cell value(s) restored on " & _
                preservedCellDicts.Count & " Output tab(s).", ""
        End If
    End If

    Application.ScreenUpdating = savedScreenUpdating

    KernelConfig.LogError SEV_INFO, "KernelFormula", "I-801", _
        "Created " & tabCount & " formula-driven tab(s) with " & totalFormulas & " formulas.", ""

    If Not m_silent Then
        MsgBox "Created " & tabCount & " formula-driven tab(s) with " & totalFormulas & " formulas.", _
               vbInformation, "RDK -- Formula Tabs"
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = savedScreenUpdating
    KernelConfig.LogError SEV_ERROR, "KernelFormula", "E-800", _
        "Error creating formula tabs: " & Err.Description, _
        "MANUAL BYPASS: Create tabs manually. Add formulas referencing" & _
        " the named ranges listed in named_range_registry.csv."
    If Not m_silent Then
        MsgBox "Error creating formula tabs: " & Err.Description & vbCrLf & vbCrLf & _
               "MANUAL BYPASS: Create tabs manually. Add formulas referencing" & vbCrLf & _
               "the named ranges listed in named_range_registry.csv.", _
               vbExclamation, "RDK -- Formula Tab Error"
    End If
End Sub




' =============================================================================
' RefreshFormulaTabs
' Called after RunProjections completes. Silent pipeline call.
' =============================================================================
Public Sub RefreshFormulaTabs()
    On Error Resume Next

    ' Suppress MsgBox during pipeline refresh
    m_silent = True
    ' Phase 11A: Preserve user-entered data on input tabs (BUG-056)
    m_preserveOnRefresh = True

    ' Clear RowID and data horizon caches so tail column detects run-off data
    KernelFormula.ClearRowIDCache

    ' Recreate formula tabs and named ranges (data may have changed since bootstrap)
    CreateFormulaTabs
    KernelFormula.CreateNamedRanges silent:=True

    ' 12f/12g: Health formatting and input validation
    ApplyHealthFormatting
    ApplyInputValidation

    ' Recalculation deferred to Cleanup block in KernelEngine.RunProjectionsEx
    ' (xlCalculationAutomatic triggers a single recalc). Standalone UI path
    ' uses RefreshFormulaTabsUI which retains its own CalculateFull call.

    m_preserveOnRefresh = False
    m_silent = False
    On Error GoTo 0
End Sub


' =============================================================================
' RefreshFormulaTabsUI
' Dashboard button entry point. Shows MsgBox feedback.
' =============================================================================
Public Sub RefreshFormulaTabsUI()
    m_silent = False
    CreateFormulaTabs
    KernelFormula.CreateNamedRanges
    ApplyHealthFormatting
    ApplyInputValidation
    Application.CalculateFull
End Sub




' =============================================================================
' PRIVATE HELPERS
' =============================================================================


' WriteFormula -- writes a formula string to a cell safely (AP-50)
Private Sub WriteFormula(ws As Worksheet, row As Long, col As Long, formula As String)
    On Error Resume Next
    If Left(formula, 1) = "=" Then
        ws.Cells(row, col).formula = formula
    Else
        ws.Cells(row, col).formula = "=" & formula
    End If
    If Err.Number <> 0 Then
        Err.Clear
        ws.Cells(row, col).NumberFormat = "@"
        ws.Cells(row, col).Value = formula
        KernelConfig.LogError SEV_WARN, "KernelFormula", "W-800", _
            "Formula write failed at row " & row & " col " & col & ": " & formula, _
            "Cell written as text instead."
    End If
    On Error GoTo 0
End Sub


' ApplyCellFormatting -- applies font, fill, border, indent, comment, format
Private Sub ApplyCellFormatting(ws As Worksheet, row As Long, col As Long, _
    fontStyle As String, fillColor As String, fontColor As String, _
    mergeSpan As Long, borderBot As String, borderTopStr As String, _
    indentLvl As Long, commentText As String, fmt As String, _
    Optional hAlign As String = "")

    On Error Resume Next

    ' Font style
    If InStr(1, fontStyle, "Bold", vbTextCompare) > 0 Then
        ws.Cells(row, col).Font.Bold = True
    End If
    If InStr(1, fontStyle, "Italic", vbTextCompare) > 0 Then
        ws.Cells(row, col).Font.Italic = True
    End If

    ' Fill color (RGB hex)
    If Len(fillColor) = 6 Then
        ws.Cells(row, col).Interior.Color = HexToRGB(fillColor)
    End If

    ' Font color (RGB hex)
    If Len(fontColor) = 6 Then
        ws.Cells(row, col).Font.Color = HexToRGB(fontColor)
    End If

    ' Column span (merge)
    If mergeSpan > 1 Then
        ws.Range(ws.Cells(row, col), ws.Cells(row, col + mergeSpan - 1)).Merge
    End If

    ' Borders
    If StrComp(borderBot, "Thin", vbTextCompare) = 0 Then
        ws.Cells(row, col).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ws.Cells(row, col).Borders(xlEdgeBottom).Weight = xlThin
    ElseIf StrComp(borderBot, "Double", vbTextCompare) = 0 Then
        ws.Cells(row, col).Borders(xlEdgeBottom).LineStyle = xlDouble
    End If

    If StrComp(borderTopStr, "Thin", vbTextCompare) = 0 Then
        ws.Cells(row, col).Borders(xlEdgeTop).LineStyle = xlContinuous
        ws.Cells(row, col).Borders(xlEdgeTop).Weight = xlThin
    ElseIf StrComp(borderTopStr, "Double", vbTextCompare) = 0 Then
        ws.Cells(row, col).Borders(xlEdgeTop).LineStyle = xlDouble
    End If

    ' Indent
    If indentLvl > 0 Then
        ws.Cells(row, col).IndentLevel = indentLvl
    End If

    ' Number format
    If Len(fmt) > 0 Then
        ws.Cells(row, col).NumberFormat = fmt
    End If

    ' Horizontal alignment
    If StrComp(hAlign, "Center", vbTextCompare) = 0 Then
        ws.Cells(row, col).HorizontalAlignment = xlCenter
    ElseIf StrComp(hAlign, "Right", vbTextCompare) = 0 Then
        ws.Cells(row, col).HorizontalAlignment = xlRight
    ElseIf StrComp(hAlign, "Left", vbTextCompare) = 0 Then
        ws.Cells(row, col).HorizontalAlignment = xlLeft
    End If

    ' Comment
    If Len(commentText) > 0 Then
        ws.Cells(row, col).ClearComments
        ws.Cells(row, col).AddComment commentText
    End If

    On Error GoTo 0
End Sub


' WriteQuarterlyHeaders -- writes Q1 Y1, Q2 Y1, ..., Y1 Total column headers
' Uses batch array write + range formatting for performance
Private Sub WriteQuarterlyHeaders(ws As Worksheet, numYears As Long, headerRow As Long, _
    Optional grandTotal As Boolean = False, Optional writeTail As Boolean = False)

    ' Calculate total columns and build header array
    Dim totalCols As Long
    totalCols = numYears * QS_COLS_PER_YEAR
    Dim hdrArr() As Variant
    ReDim hdrArr(1 To 1, 1 To totalCols)

    Dim yr As Long
    For yr = 1 To numYears
        Dim qtr As Long
        For qtr = 1 To QS_QUARTERS_PER_YEAR
            Dim arrIdx As Long
            arrIdx = (yr - 1) * QS_COLS_PER_YEAR + qtr
            hdrArr(1, arrIdx) = "Q" & qtr & " Y" & yr
        Next qtr
        ' Annual total header
        hdrArr(1, (yr - 1) * QS_COLS_PER_YEAR + QS_QUARTERS_PER_YEAR + 1) = "Y" & yr & " Total"
    Next yr

    ' Batch write all headers at once
    Dim hdrRange As Range
    Set hdrRange = ws.Range(ws.Cells(headerRow, QS_DATA_START_COL), _
                            ws.Cells(headerRow, QS_DATA_START_COL + totalCols - 1))
    hdrRange.Value = hdrArr
    hdrRange.Font.Bold = True
    hdrRange.HorizontalAlignment = xlCenter

    ' Grey shading for annual total columns
    For yr = 1 To numYears
        Dim annCol As Long
        annCol = QS_DATA_START_COL + (yr - 1) * QS_COLS_PER_YEAR + QS_QUARTERS_PER_YEAR
        ws.Cells(headerRow, annCol).Interior.Color = RGB(217, 217, 217)
    Next yr

    ' Grand total header (Y1-Yn sum)
    If grandTotal Then
        Dim gtCol As Long
        gtCol = QS_DATA_START_COL + numYears * QS_COLS_PER_YEAR
        ws.Cells(headerRow, gtCol).Value = "Y1-Y" & numYears & " Total"
        ws.Cells(headerRow, gtCol).Font.Bold = True
        ws.Cells(headerRow, gtCol).HorizontalAlignment = xlCenter
        ws.Cells(headerRow, gtCol).Interior.Color = RGB(189, 215, 238)
        ' Tail header (development beyond writing horizon) -- UW Exec Summary only
        Dim hdrDataYears As Long
        hdrDataYears = KernelFormula.GetDataHorizonYears()
        If hdrDataYears > numYears And writeTail Then
            Dim tailHdrCol As Long
            tailHdrCol = gtCol + 1
            ws.Cells(headerRow, tailHdrCol).Value = "Tail"
            ws.Cells(headerRow, tailHdrCol).Font.Bold = True
            ws.Cells(headerRow, tailHdrCol).HorizontalAlignment = xlCenter
            ws.Cells(headerRow, tailHdrCol).Interior.Color = RGB(198, 224, 180)
        End If
    End If
End Sub


' HasQuarterlyColumns -- checks tab_registry for QuarterlyColumns=TRUE
Private Function HasQuarterlyColumns(tabName As String) As Boolean
    HasQuarterlyColumns = False
    On Error Resume Next
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    If wsConfig Is Nothing Then
        KernelConfig.LogError SEV_WARN, "KernelFormulaWriter", "W-830", _
            "Config sheet not found in HasQuarterlyColumns", ""
        Exit Function
    End If
    On Error GoTo 0

    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_TAB_REGISTRY)
    If sr = 0 Then
        KernelConfig.LogError SEV_WARN, "KernelFormulaWriter", "W-831", _
            "Tab registry section not found in HasQuarterlyColumns", ""
        Exit Function
    End If

    Dim dr As Long
    dr = sr + 2
    Do While Len(Trim(CStr(wsConfig.Cells(dr, TREG_COL_TABNAME).Value))) > 0
        If StrComp(Trim(CStr(wsConfig.Cells(dr, TREG_COL_TABNAME).Value)), tabName, vbTextCompare) = 0 Then
            Dim qcVal As String
            qcVal = Trim(CStr(wsConfig.Cells(dr, TREG_COL_QUARTERLY).Value))
            HasQuarterlyColumns = (StrComp(qcVal, "TRUE", vbTextCompare) = 0)
            Exit Function
        End If
        dr = dr + 1
    Loop
End Function


' HasGrandTotal -- checks tab_registry for GrandTotal=TRUE
Private Function HasGrandTotal(tabName As String) As Boolean
    HasGrandTotal = False
    On Error Resume Next
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    If wsConfig Is Nothing Then
        KernelConfig.LogError SEV_WARN, "KernelFormulaWriter", "W-832", _
            "Config sheet not found in HasGrandTotal", ""
        Exit Function
    End If
    On Error GoTo 0

    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_TAB_REGISTRY)
    If sr = 0 Then
        KernelConfig.LogError SEV_WARN, "KernelFormulaWriter", "W-833", _
            "Tab registry section not found in HasGrandTotal", ""
        Exit Function
    End If

    Dim dr As Long
    dr = sr + 2
    Do While Len(Trim(CStr(wsConfig.Cells(dr, TREG_COL_TABNAME).Value))) > 0
        If StrComp(Trim(CStr(wsConfig.Cells(dr, TREG_COL_TABNAME).Value)), tabName, vbTextCompare) = 0 Then
            Dim gtVal As String
            gtVal = Trim(CStr(wsConfig.Cells(dr, TREG_COL_GRANDTOTAL).Value))
            HasGrandTotal = (StrComp(gtVal, "TRUE", vbTextCompare) = 0)
            Exit Function
        End If
        dr = dr + 1
    Loop
End Function




' EnsureSheetFormula -- creates a sheet if it does not exist
Private Function EnsureSheetFormula(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    End If
    Set EnsureSheetFormula = ws
End Function




' ColLetterToNum -- converts column letter to 1-based number (A=1, C=3, AA=27)
Private Function ColLetterToNum(colLetter As String) As Long
    Dim result As Long
    result = 0
    Dim i As Long
    For i = 1 To Len(colLetter)
        result = result * 26 + (Asc(UCase(Mid(colLetter, i, 1))) - 64)
    Next i
    ColLetterToNum = result
End Function


' HexToRGB -- converts 6-char hex to Long RGB
Private Function HexToRGB(hexStr As String) As Long
    If Len(hexStr) <> 6 Then
        HexToRGB = 0
        Exit Function
    End If
    Dim r As Long
    Dim g As Long
    Dim b As Long
    r = Val("&H" & Left(hexStr, 2))
    g = Val("&H" & Mid(hexStr, 3, 2))
    b = Val("&H" & Right(hexStr, 2))
    HexToRGB = RGB(r, g, b)
End Function


' GetTabRegistryFlag -- reads a Boolean flag column from tab_registry on Config sheet
Private Function GetTabRegistryFlag(tabName As String, colIdx As Long) As Boolean
    GetTabRegistryFlag = False
    On Error Resume Next
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    If wsConfig Is Nothing Then Exit Function
    On Error GoTo 0
    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_TAB_REGISTRY)
    If sr = 0 Then Exit Function
    Dim dr As Long
    dr = sr + 2
    Do While Len(Trim(CStr(wsConfig.Cells(dr, TREG_COL_TABNAME).Value))) > 0
        If StrComp(Trim(CStr(wsConfig.Cells(dr, TREG_COL_TABNAME).Value)), tabName, vbTextCompare) = 0 Then
            GetTabRegistryFlag = (StrComp(Trim(CStr(wsConfig.Cells(dr, colIdx).Value)), "TRUE", vbTextCompare) = 0)
            Exit Function
        End If
        dr = dr + 1
    Loop
End Function


' IsInArray -- checks if a string exists in a partially filled array
Private Function IsInArray(val As String, arr() As String, cnt As Long) As Boolean
    IsInArray = False
    Dim i As Long
    For i = 1 To cnt
        If StrComp(arr(i), val, vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next i
End Function


' =============================================================================
' ApplyHealthFormatting (12f)
' Config-driven conditional formatting. Reads health_config from Config sheet.
' Each row specifies TabName, RowID, column range, check type, and thresholds.
' Zero domain-specific references.
' =============================================================================
Public Sub ApplyHealthFormatting()
    On Error Resume Next
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    If wsConfig Is Nothing Then Exit Sub
    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_HEALTH_CONFIG)
    If sr = 0 Then Exit Sub

    Dim timeHorizon As Long
    timeHorizon = KernelConfig.GetTimeHorizon()
    If timeHorizon <= 0 Then timeHorizon = 12
    Dim numYears As Long
    numYears = (timeHorizon \ 3) \ QS_QUARTERS_PER_YEAR
    If numYears < 1 Then numYears = 1
    Dim lastQCol As Long
    lastQCol = QS_DATA_START_COL + numYears * QS_COLS_PER_YEAR

    Dim dr As Long
    dr = sr + 2
    Do While Len(Trim(CStr(wsConfig.Cells(dr, HLCFG_COL_TABNAME).Value))) > 0
        Dim hTabName As String
        hTabName = Trim(CStr(wsConfig.Cells(dr, HLCFG_COL_TABNAME).Value))
        Dim hRowID As String
        hRowID = Trim(CStr(wsConfig.Cells(dr, HLCFG_COL_ROWID).Value))
        Dim hColStart As String
        hColStart = Trim(CStr(wsConfig.Cells(dr, HLCFG_COL_COLSTART).Value))
        Dim hColEnd As String
        hColEnd = Trim(CStr(wsConfig.Cells(dr, HLCFG_COL_COLEND).Value))
        Dim hCheckType As String
        hCheckType = Trim(CStr(wsConfig.Cells(dr, HLCFG_COL_CHECKTYPE).Value))
        Dim hGoodValue As String
        hGoodValue = Trim(CStr(wsConfig.Cells(dr, HLCFG_COL_GOODVALUE).Value))
        Dim hThreshold As String
        hThreshold = Trim(CStr(wsConfig.Cells(dr, HLCFG_COL_THRESHOLD).Value))

        Dim hWs As Worksheet
        Set hWs = Nothing
        Set hWs = ThisWorkbook.Sheets(hTabName)
        If Not hWs Is Nothing Then
            Dim hRow As Long
            hRow = KernelFormula.ResolveRowID(hTabName, hRowID)
            If hRow > 0 Then
                Dim startC As Long
                Dim endC As Long
                If StrComp(hColStart, "Q", vbTextCompare) = 0 Then
                    startC = QS_DATA_START_COL
                    endC = lastQCol
                Else
                    startC = CLng(hColStart)
                    endC = CLng(hColEnd)
                End If
                Dim hRng As Range
                Set hRng = hWs.Range(hWs.Cells(hRow, startC), hWs.Cells(hRow, endC))
                hRng.FormatConditions.Delete

                If StrComp(hCheckType, "NumericZero", vbTextCompare) = 0 Then
                    Dim thresh As String
                    thresh = hThreshold
                    If Len(thresh) = 0 Then thresh = "0.01"
                    hRng.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
                        Formula1:="-" & thresh, Formula2:=thresh
                    hRng.FormatConditions(hRng.FormatConditions.Count).Interior.Color = RGB(198, 239, 206)
                    hRng.FormatConditions(hRng.FormatConditions.Count).Font.Color = RGB(0, 97, 0)
                    hRng.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotBetween, _
                        Formula1:="-" & thresh, Formula2:=thresh
                    hRng.FormatConditions(hRng.FormatConditions.Count).Interior.Color = RGB(255, 199, 206)
                    hRng.FormatConditions(hRng.FormatConditions.Count).Font.Color = RGB(156, 0, 6)
                ElseIf StrComp(hCheckType, "TextMatch", vbTextCompare) = 0 Then
                    hRng.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
                        Formula1:="=""" & hGoodValue & """"
                    hRng.FormatConditions(hRng.FormatConditions.Count).Interior.Color = RGB(198, 239, 206)
                    hRng.FormatConditions(hRng.FormatConditions.Count).Font.Color = RGB(0, 97, 0)
                    hRng.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
                        Formula1:="=""" & hGoodValue & """"
                    hRng.FormatConditions(hRng.FormatConditions.Count).Interior.Color = RGB(255, 199, 206)
                    hRng.FormatConditions(hRng.FormatConditions.Count).Font.Color = RGB(156, 0, 6)
                End If
            End If
        End If
        dr = dr + 1
    Loop
    On Error GoTo 0
End Sub


' =============================================================================
' ApplyInputValidation (12g)
' Config-driven data validation. Reads validation_config from Config sheet.
' Each row specifies TabName, RowID pattern, column range, validation type,
' operator, min/max, alert style, and error message.
' Zero domain-specific references.
' =============================================================================
Public Sub ApplyInputValidation()
    On Error Resume Next
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    If wsConfig Is Nothing Then Exit Sub
    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_VALIDATION_CONFIG)
    If sr = 0 Then Exit Sub

    Dim lastQCol As Long
    lastQCol = GetLastQuarterlyCol()

    Dim dr As Long
    dr = sr + 2
    Do While Len(Trim(CStr(wsConfig.Cells(dr, VALCFG_COL_TABNAME).Value))) > 0
        Dim vTabName As String
        vTabName = Trim(CStr(wsConfig.Cells(dr, VALCFG_COL_TABNAME).Value))
        Dim vPattern As String
        vPattern = Trim(CStr(wsConfig.Cells(dr, VALCFG_COL_PATTERN).Value))
        Dim vColStart As String
        vColStart = Trim(CStr(wsConfig.Cells(dr, VALCFG_COL_COLSTART).Value))
        Dim vColEnd As String
        vColEnd = Trim(CStr(wsConfig.Cells(dr, VALCFG_COL_COLEND).Value))
        Dim vValType As String
        vValType = Trim(CStr(wsConfig.Cells(dr, VALCFG_COL_VALTYPE).Value))
        Dim vOperator As String
        vOperator = Trim(CStr(wsConfig.Cells(dr, VALCFG_COL_OPERATOR).Value))
        Dim vMin As String
        vMin = Trim(CStr(wsConfig.Cells(dr, VALCFG_COL_MIN).Value))
        Dim vMax As String
        vMax = Trim(CStr(wsConfig.Cells(dr, VALCFG_COL_MAX).Value))
        Dim vAlert As String
        vAlert = Trim(CStr(wsConfig.Cells(dr, VALCFG_COL_ALERTSTYLE).Value))
        Dim vErrMsg As String
        vErrMsg = Trim(CStr(wsConfig.Cells(dr, VALCFG_COL_ERRMSG).Value))

        ' Resolve Excel constants
        Dim xlValType As Long
        If StrComp(vValType, "WholeNumber", vbTextCompare) = 0 Then
            xlValType = xlValidateWholeNumber
        Else
            xlValType = xlValidateDecimal
        End If
        Dim xlOp As Long
        If StrComp(vOperator, "Between", vbTextCompare) = 0 Then
            xlOp = xlBetween
        ElseIf StrComp(vOperator, "GreaterEqual", vbTextCompare) = 0 Then
            xlOp = xlGreaterEqual
        ElseIf StrComp(vOperator, "LessEqual", vbTextCompare) = 0 Then
            xlOp = xlLessEqual
        Else
            xlOp = xlBetween
        End If
        Dim xlAlert As Long
        If StrComp(vAlert, "Stop", vbTextCompare) = 0 Then
            xlAlert = xlValidAlertStop
        Else
            xlAlert = xlValidAlertWarning
        End If

        ' Resolve column range
        Dim startC As Long
        Dim endC As Long
        If StrComp(vColStart, "Q", vbTextCompare) = 0 Then
            startC = QS_DATA_START_COL
            endC = lastQCol
        Else
            startC = CLng(vColStart)
            endC = CLng(vColEnd)
        End If

        ' Find the tab
        Dim vWs As Worksheet
        Set vWs = Nothing
        Set vWs = ThisWorkbook.Sheets(vTabName)
        If Not vWs Is Nothing Then
            ' Check if pattern contains wildcard
            Dim hasWild As Boolean
            hasWild = (Right(vPattern, 1) = "*")
            Dim prefix As String
            If hasWild Then
                prefix = Left(vPattern, Len(vPattern) - 1)
            End If

            If hasWild Then
                ' Scan column A for matching RowIDs
                Dim vLast As Long
                vLast = vWs.Cells(vWs.Rows.Count, 1).End(xlUp).Row
                If vLast > 500 Then vLast = 500
                Dim vr As Long
                For vr = 1 To vLast
                    Dim rid As String
                    rid = Trim(CStr(vWs.Cells(vr, 1).Value))
                    If Len(rid) >= Len(prefix) Then
                        If StrComp(Left(rid, Len(prefix)), prefix, vbTextCompare) = 0 Then
                            ' Only validate input cells (blue font)
                            If vWs.Cells(vr, startC).Font.Color = RGB(0, 0, 255) Then
                                ApplyVal vWs, vr, startC, endC, xlValType, xlAlert, xlOp, vMin, vMax, vErrMsg
                            End If
                        End If
                    End If
                Next vr
            Else
                ' Exact RowID match
                Dim rowNum As Long
                rowNum = KernelFormula.ResolveRowID(vTabName, vPattern)
                If rowNum > 0 Then
                    ApplyVal vWs, rowNum, startC, endC, xlValType, xlAlert, xlOp, vMin, vMax, vErrMsg
                End If
            End If
        End If
        dr = dr + 1
    Loop
    On Error GoTo 0
End Sub


' ApplyVal -- helper to apply data validation to a row range
Private Sub ApplyVal(ws As Worksheet, rowNum As Long, startCol As Long, endCol As Long, _
    valType As Long, alertStyle As Long, op As Long, _
    formula1 As String, formula2 As String, errMsg As String)

    On Error Resume Next
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(rowNum, startCol), ws.Cells(rowNum, endCol))
    rng.Validation.Delete
    If Len(formula2) > 0 Then
        rng.Validation.Add Type:=valType, AlertStyle:=alertStyle, _
            Operator:=op, Formula1:=formula1, Formula2:=formula2
    Else
        rng.Validation.Add Type:=valType, AlertStyle:=alertStyle, _
            Operator:=op, Formula1:=formula1
    End If
    rng.Validation.ErrorTitle = "Invalid Input"
    rng.Validation.ErrorMessage = errMsg
    On Error GoTo 0
End Sub


' GetLastQuarterlyCol -- returns the last quarterly data column index
Private Function GetLastQuarterlyCol() As Long
    Dim timeHorizon As Long
    timeHorizon = KernelConfig.GetTimeHorizon()
    If timeHorizon <= 0 Then timeHorizon = 12
    Dim numYears As Long
    numYears = (timeHorizon \ 3) \ QS_QUARTERS_PER_YEAR
    If numYears < 1 Then numYears = 1
    GetLastQuarterlyCol = QS_DATA_START_COL + numYears * QS_COLS_PER_YEAR - 1
End Function
