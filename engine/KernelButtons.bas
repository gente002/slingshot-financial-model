Attribute VB_Name = "KernelButtons"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelButtons.bas
' Purpose: Config-driven button creation for any tab. Reads button_config.csv
'          from the Config sheet and creates/positions Excel form control buttons.
'          Supports vertical stacking (Dashboard) and cell-positioned placement.
' =============================================================================


' =============================================================================
' CreateButtonsFromConfig
' Reads button_config, filters by tabName, sorts by SortOrder, creates buttons.
' Each shape is named "btn_" & ButtonID for dev mode toggling.
' =============================================================================
Public Sub CreateButtonsFromConfig(ws As Worksheet, tabName As String, curDevMode As String)
    Dim wsConfig As Worksheet
    Set wsConfig = Nothing
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    On Error GoTo 0
    If wsConfig Is Nothing Then Exit Sub

    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_BUTTON_CONFIG)
    If sr = 0 Then Exit Sub

    ' First pass: count matching buttons
    Dim dr As Long
    dr = sr + 2
    Dim totalCount As Long
    totalCount = 0
    Do While Len(Trim(CStr(wsConfig.Cells(dr, BTNCFG_COL_TABNAME).Value))) > 0
        Dim cellTab As String
        cellTab = Trim(CStr(wsConfig.Cells(dr, BTNCFG_COL_TABNAME).Value))
        Dim cellEnabled As String
        cellEnabled = Trim(CStr(wsConfig.Cells(dr, BTNCFG_COL_ENABLED).Value))
        If StrComp(cellTab, tabName, vbTextCompare) = 0 And _
           StrComp(cellEnabled, "TRUE", vbTextCompare) = 0 Then
            totalCount = totalCount + 1
        End If
        dr = dr + 1
    Loop

    If totalCount = 0 Then Exit Sub

    ' Collect matching buttons into parallel arrays
    Dim arrID() As String
    Dim arrCaption() As String
    Dim arrMacro() As String
    Dim arrDevOnly() As Boolean
    Dim arrSort() As Long
    Dim arrRow() As Long
    Dim arrCol() As Long
    ReDim arrID(1 To totalCount)
    ReDim arrCaption(1 To totalCount)
    ReDim arrMacro(1 To totalCount)
    ReDim arrDevOnly(1 To totalCount)
    ReDim arrSort(1 To totalCount)
    ReDim arrRow(1 To totalCount)
    ReDim arrCol(1 To totalCount)

    Dim idx As Long
    idx = 0
    dr = sr + 2
    Do While Len(Trim(CStr(wsConfig.Cells(dr, BTNCFG_COL_TABNAME).Value))) > 0
        Dim cTab As String
        cTab = Trim(CStr(wsConfig.Cells(dr, BTNCFG_COL_TABNAME).Value))
        Dim cEn As String
        cEn = Trim(CStr(wsConfig.Cells(dr, BTNCFG_COL_ENABLED).Value))
        If StrComp(cTab, tabName, vbTextCompare) = 0 And _
           StrComp(cEn, "TRUE", vbTextCompare) = 0 Then
            idx = idx + 1
            arrID(idx) = Trim(CStr(wsConfig.Cells(dr, BTNCFG_COL_ID).Value))
            arrCaption(idx) = Trim(CStr(wsConfig.Cells(dr, BTNCFG_COL_CAPTION).Value))
            arrMacro(idx) = Trim(CStr(wsConfig.Cells(dr, BTNCFG_COL_MACRO).Value))
            Dim devVal As String
            devVal = Trim(CStr(wsConfig.Cells(dr, BTNCFG_COL_DEVONLY).Value))
            arrDevOnly(idx) = (StrComp(devVal, "TRUE", vbTextCompare) = 0)
            Dim sortVal As String
            sortVal = Trim(CStr(wsConfig.Cells(dr, BTNCFG_COL_SORT).Value))
            If IsNumeric(sortVal) And Len(sortVal) > 0 Then
                arrSort(idx) = CLng(sortVal)
            Else
                arrSort(idx) = 9999
            End If
            Dim rowVal As String
            rowVal = Trim(CStr(wsConfig.Cells(dr, BTNCFG_COL_ROW).Value))
            Dim colVal As String
            colVal = Trim(CStr(wsConfig.Cells(dr, BTNCFG_COL_COL).Value))
            If IsNumeric(rowVal) And Len(rowVal) > 0 Then arrRow(idx) = CLng(rowVal)
            If IsNumeric(colVal) And Len(colVal) > 0 Then arrCol(idx) = CLng(colVal)
        End If
        dr = dr + 1
    Loop

    ' Bubble sort by SortOrder
    If totalCount > 1 Then
        Dim i As Long
        Dim j As Long
        Dim swapped As Boolean
        Dim tmpStr As String
        Dim tmpBool As Boolean
        Dim tmpLng As Long
        For i = 1 To totalCount - 1
            swapped = False
            For j = 1 To totalCount - i
                If arrSort(j) > arrSort(j + 1) Then
                    tmpStr = arrID(j): arrID(j) = arrID(j + 1): arrID(j + 1) = tmpStr
                    tmpStr = arrCaption(j): arrCaption(j) = arrCaption(j + 1): arrCaption(j + 1) = tmpStr
                    tmpStr = arrMacro(j): arrMacro(j) = arrMacro(j + 1): arrMacro(j + 1) = tmpStr
                    tmpBool = arrDevOnly(j): arrDevOnly(j) = arrDevOnly(j + 1): arrDevOnly(j + 1) = tmpBool
                    tmpLng = arrSort(j): arrSort(j) = arrSort(j + 1): arrSort(j + 1) = tmpLng
                    tmpLng = arrRow(j): arrRow(j) = arrRow(j + 1): arrRow(j + 1) = tmpLng
                    tmpLng = arrCol(j): arrCol(j) = arrCol(j + 1): arrCol(j + 1) = tmpLng
                    swapped = True
                End If
            Next j
            If Not swapped Then Exit For
        Next i
    End If

    ' Create buttons
    Dim btnLeft As Double
    btnLeft = ws.Cells(1, 1).Left + 12
    Dim btnWidth As Double
    btnWidth = 180
    Dim btnTop As Double
    btnTop = ws.Cells(4, 1).Top

    Dim devMode As Boolean
    devMode = (StrComp(curDevMode, DEV_MODE_ON, vbTextCompare) = 0)

    Dim inDevSection As Boolean
    inDevSection = False
    Dim isFirstBtn As Boolean
    isFirstBtn = True

    Dim btnIdx As Long
    For btnIdx = 1 To totalCount
        ' Dev section spacing
        If arrDevOnly(btnIdx) And Not inDevSection Then
            If Not isFirstBtn Then btnTop = btnTop + 52
            inDevSection = True
        ElseIf Not isFirstBtn Then
            btnTop = btnTop + 42
        End If

        ' Special caption for DEV_MODE button
        Dim caption As String
        caption = arrCaption(btnIdx)
        If StrComp(arrID(btnIdx), "DEV_MODE", vbTextCompare) = 0 Then
            If devMode Then
                caption = "Dev Mode: ON (click to hide)"
            Else
                caption = "Dev Mode: OFF (click to show)"
            End If
        End If

        ' Button sizing: primary action tall, others standard. All same width.
        Dim btnHeight As Double
        Dim useBtnWidth As Double
        useBtnWidth = btnWidth  ' all buttons same width (matches col B)
        If arrSort(btnIdx) <= 10 Then
            btnHeight = 34
        ElseIf arrDevOnly(btnIdx) Then
            btnHeight = 24
        Else
            btnHeight = 28
        End If

        ' Position: Row/Col if specified, else vertical stack
        Dim useBtnLeft As Double
        Dim useBtnTop As Double
        If arrRow(btnIdx) > 0 And arrCol(btnIdx) > 0 Then
            useBtnLeft = ws.Cells(arrRow(btnIdx), arrCol(btnIdx)).Left
            useBtnTop = ws.Cells(arrRow(btnIdx), arrCol(btnIdx)).Top
        Else
            useBtnLeft = btnLeft
            useBtnTop = btnTop
        End If
        AddMacroButton ws, useBtnLeft, useBtnTop, useBtnWidth, btnHeight, _
            caption, arrMacro(btnIdx), "btn_" & arrID(btnIdx)

        ' Hide dev-only buttons when dev mode is off
        If arrDevOnly(btnIdx) And Not devMode Then
            On Error Resume Next
            ws.Shapes("btn_" & arrID(btnIdx)).Visible = msoFalse
            On Error GoTo 0
        End If

        isFirstBtn = False
    Next btnIdx
End Sub


' =============================================================================
' CreateButtonsOnAllTabs
' Iterates button_config for unique tab names other than Dashboard and
' creates buttons on each. Dashboard buttons are handled by SetupDashboardTab.
' =============================================================================
Public Sub CreateButtonsOnAllTabs()
    On Error Resume Next
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    If wsConfig Is Nothing Then Exit Sub

    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_BUTTON_CONFIG)
    If sr = 0 Then Exit Sub

    Dim tabDict As Object
    Set tabDict = CreateObject("Scripting.Dictionary")
    tabDict.CompareMode = vbTextCompare
    Dim dr As Long
    dr = sr + 2
    Do While Len(Trim(CStr(wsConfig.Cells(dr, BTNCFG_COL_TABNAME).Value))) > 0
        Dim tn As String
        tn = Trim(CStr(wsConfig.Cells(dr, BTNCFG_COL_TABNAME).Value))
        If StrComp(tn, TAB_DASHBOARD, vbTextCompare) <> 0 Then
            If Not tabDict.Exists(tn) Then tabDict.Add tn, True
        End If
        dr = dr + 1
    Loop

    Dim devMode As String
    devMode = KernelConfig.GetDevMode()
    If Len(devMode) = 0 Then devMode = DEV_MODE_OFF

    Dim k As Variant
    For Each k In tabDict.Keys
        Dim ws As Worksheet
        Set ws = Nothing
        Set ws = ThisWorkbook.Sheets(CStr(k))
        If Not ws Is Nothing Then
            CreateButtonsFromConfig ws, CStr(k), devMode
        End If
    Next k
    On Error GoTo 0
End Sub


' =============================================================================
' AddMacroButton
' Adds a Form Control button to a worksheet and assigns a macro.
' =============================================================================
Public Sub AddMacroButton(ws As Worksheet, btnLeft As Double, btnTop As Double, _
                            btnWidth As Double, btnHeight As Double, _
                            caption As String, macroName As String, _
                            Optional shapeName As String = "")
    Dim btn As Shape
    Set btn = ws.Shapes.AddFormControl(xlButtonControl, _
              btnLeft, btnTop, btnWidth, btnHeight)
    btn.TextFrame.Characters.Text = caption
    btn.OnAction = macroName
    If Len(shapeName) > 0 Then
        btn.Name = shapeName
    End If
End Sub
