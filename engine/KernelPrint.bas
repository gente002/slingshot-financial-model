Attribute VB_Name = "KernelPrint"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelPrint.bas
' Purpose: Print preview and PDF export configured from print_config.csv.
' Phase 5B module. All errors include manual bypass instructions (AP-46).
' =============================================================================


' =============================================================================
' ConfigurePrintSettings
' Reads print_config from Config sheet and applies PageSetup to each tab.
' =============================================================================
Public Sub ConfigurePrintSettings(Optional silent As Boolean = False)
    On Error GoTo ErrHandler

    Dim cnt As Long
    cnt = KernelConfig.GetPrintConfigCount()

    If cnt = 0 Then
        KernelConfig.LogError SEV_INFO, "KernelPrint", "I-600", _
            "No print config entries found. Skipping print setup.", ""
        If Not silent Then
            MsgBox "No print configuration found." & vbCrLf & vbCrLf & _
                   "MANUAL BYPASS: Add entries to print_config.csv and re-run Setup.bat.", _
                   vbInformation, "RDK -- Print Setup"
        End If
        Exit Sub
    End If

    Dim configured As Long
    configured = 0

    ' Batch printer communication to avoid per-property round-trips (perf)
    On Error Resume Next
    Application.PrintCommunication = False
    On Error GoTo ErrHandler

    Dim idx As Long
    For idx = 1 To cnt
        Dim tabName As String
        tabName = KernelConfig.GetPrintConfigField(idx, PRTCFG_COL_TABNAME)
        If Len(tabName) = 0 Then GoTo NextTab

        ' Check if tab exists
        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(tabName)
        On Error GoTo ErrHandler
        If ws Is Nothing Then
            KernelConfig.LogError SEV_WARN, "KernelPrint", "W-600", _
                "Print config references non-existent tab: " & tabName, _
                "MANUAL BYPASS: Create the '" & tabName & "' tab or remove it from print_config.csv."
            GoTo NextTab
        End If

        ' Apply PageSetup
        With ws.PageSetup
            ' Orientation
            Dim orient As String
            orient = KernelConfig.GetPrintConfigField(idx, PRTCFG_COL_ORIENT)
            If StrComp(orient, "Landscape", vbTextCompare) = 0 Then
                .Orientation = xlLandscape
            Else
                .Orientation = xlPortrait
            End If

            ' FitToPages (wide and tall)
            Dim fitPages As String
            fitPages = KernelConfig.GetPrintConfigField(idx, PRTCFG_COL_FITPAGES)
            Dim fitTall As String
            fitTall = KernelConfig.GetPrintConfigField(idx, PRTCFG_COL_FITPAGESTALL)
            If (IsNumeric(fitPages) And Len(fitPages) > 0) Or _
               (IsNumeric(fitTall) And Len(fitTall) > 0) Then
                .Zoom = False
                If IsNumeric(fitPages) And Len(fitPages) > 0 Then
                    .FitToPagesWide = CLng(fitPages)
                Else
                    .FitToPagesWide = 1
                End If
                If IsNumeric(fitTall) And Len(fitTall) > 0 Then
                    .FitToPagesTall = CLng(fitTall)
                Else
                    .FitToPagesTall = False
                End If
            End If

            ' Margins (Narrow/Normal/Wide or custom L,R,T,B in inches)
            Dim margins As String
            margins = KernelConfig.GetPrintConfigField(idx, PRTCFG_COL_MARGINS)
            If Len(margins) > 0 Then
                Select Case LCase(margins)
                    Case "narrow"
                        .LeftMargin = Application.InchesToPoints(0.25)
                        .RightMargin = Application.InchesToPoints(0.25)
                        .TopMargin = Application.InchesToPoints(0.75)
                        .BottomMargin = Application.InchesToPoints(0.75)
                        .HeaderMargin = Application.InchesToPoints(0.3)
                        .FooterMargin = Application.InchesToPoints(0.3)
                    Case "normal"
                        .LeftMargin = Application.InchesToPoints(0.7)
                        .RightMargin = Application.InchesToPoints(0.7)
                        .TopMargin = Application.InchesToPoints(0.75)
                        .BottomMargin = Application.InchesToPoints(0.75)
                        .HeaderMargin = Application.InchesToPoints(0.3)
                        .FooterMargin = Application.InchesToPoints(0.3)
                    Case "wide"
                        .LeftMargin = Application.InchesToPoints(1#)
                        .RightMargin = Application.InchesToPoints(1#)
                        .TopMargin = Application.InchesToPoints(1#)
                        .BottomMargin = Application.InchesToPoints(1#)
                        .HeaderMargin = Application.InchesToPoints(0.5)
                        .FooterMargin = Application.InchesToPoints(0.5)
                    Case Else
                        ' Custom: L,R,T,B in inches (e.g. "0.5,0.5,0.75,0.75")
                        Dim parts() As String
                        parts = Split(margins, ",")
                        If UBound(parts) >= 3 Then
                            .LeftMargin = Application.InchesToPoints(CDbl(Trim(parts(0))))
                            .RightMargin = Application.InchesToPoints(CDbl(Trim(parts(1))))
                            .TopMargin = Application.InchesToPoints(CDbl(Trim(parts(2))))
                            .BottomMargin = Application.InchesToPoints(CDbl(Trim(parts(3))))
                        End If
                End Select
            End If

            ' PaperSize
            Dim paper As String
            paper = KernelConfig.GetPrintConfigField(idx, PRTCFG_COL_PAPER)
            Select Case UCase(paper)
                Case "LEGAL": .PaperSize = xlPaperLegal
                Case "A4": .PaperSize = xlPaperA4
                Case Else: .PaperSize = xlPaperLetter
            End Select

            ' PrintArea
            Dim printArea As String
            printArea = KernelConfig.GetPrintConfigField(idx, PRTCFG_COL_PRINTAREA)
            If Len(printArea) > 0 Then
                .PrintArea = printArea
            End If

            ' Headers
            Dim hdrLeft As String
            hdrLeft = KernelConfig.GetPrintConfigField(idx, PRTCFG_COL_HDRLEFT)
            If Len(hdrLeft) > 0 Then .LeftHeader = hdrLeft

            Dim hdrCenter As String
            hdrCenter = KernelConfig.GetPrintConfigField(idx, PRTCFG_COL_HDRCENTER)
            If Len(hdrCenter) > 0 Then .CenterHeader = hdrCenter

            Dim hdrRight As String
            hdrRight = KernelConfig.GetPrintConfigField(idx, PRTCFG_COL_HDRRIGHT)
            If Len(hdrRight) > 0 Then .RightHeader = hdrRight

            ' Footer
            Dim ftrCenter As String
            ftrCenter = KernelConfig.GetPrintConfigField(idx, PRTCFG_COL_FTRCENTER)
            If Len(ftrCenter) > 0 Then
                .CenterFooter = ftrCenter
            Else
                .CenterFooter = "Page &P of &N"
            End If

            ' CenterHorizontally
            Dim centerH As String
            centerH = KernelConfig.GetPrintConfigField(idx, PRTCFG_COL_CENTERH)
            If UCase(Trim(centerH)) = "TRUE" Then
                .CenterHorizontally = True
            Else
                .CenterHorizontally = False
            End If
        End With

        configured = configured + 1
NextTab:
    Next idx

    ' Resume printer communication
    On Error Resume Next
    Application.PrintCommunication = True
    On Error GoTo 0

    If Not silent Then
        MsgBox "Print settings configured for " & configured & " tab(s).", _
               vbInformation, "RDK -- Print Setup"
    End If
    Exit Sub

ErrHandler:
    On Error Resume Next
    Application.PrintCommunication = True
    On Error GoTo 0
    KernelConfig.LogError SEV_ERROR, "KernelPrint", "E-600", _
        "Error configuring print settings: " & Err.Description, _
        "MANUAL BYPASS: Use File -> Page Setup to configure print settings manually for each tab."
    If Not silent Then
        MsgBox "Error configuring print settings: " & Err.Description & vbCrLf & vbCrLf & _
               "MANUAL BYPASS: Use File -> Page Setup manually.", _
               vbExclamation, "RDK -- Print Error"
    End If
End Sub


' =============================================================================
' PrintPreview
' Opens print preview for configured tabs.
' =============================================================================
Public Sub PrintPreview(Optional tabName As String = "")
    On Error GoTo ErrHandler

    ConfigurePrintSettings True

    If Len(tabName) > 0 Then
        ' Preview specific tab
        Dim wsSingle As Worksheet
        Set wsSingle = Nothing
        On Error Resume Next
        Set wsSingle = ThisWorkbook.Sheets(tabName)
        On Error GoTo ErrHandler
        If wsSingle Is Nothing Then
            MsgBox "Tab '" & tabName & "' not found.", vbExclamation, "RDK -- Print Preview"
            Exit Sub
        End If
        wsSingle.PrintPreview
        Exit Sub
    End If

    ' Preview all IncludeInPDF tabs in PrintOrder
    Dim tabNames() As String
    Dim tabCount As Long
    tabCount = GetPdfTabNames(tabNames)
    If tabCount = 0 Then
        MsgBox "No tabs configured for PDF/print output." & vbCrLf & vbCrLf & _
               "MANUAL BYPASS: Set IncludeInPDF=TRUE in print_config.csv.", _
               vbInformation, "RDK -- Print Preview"
        Exit Sub
    End If

    ' Select all tabs then preview
    SelectMultipleTabs tabNames, tabCount
    ActiveWindow.SelectedSheets.PrintPreview
    ThisWorkbook.Sheets(tabNames(1)).Select
    Exit Sub

ErrHandler:
    KernelConfig.LogError SEV_ERROR, "KernelPrint", "E-601", _
        "Print preview error: " & Err.Description, _
        "MANUAL BYPASS: Select tabs manually then use File -> Print Preview."
    MsgBox "Print preview error: " & Err.Description, vbExclamation, "RDK -- Error"
End Sub


' =============================================================================
' ExportPDF
' Exports all IncludeInPDF tabs to a single PDF file.
' =============================================================================
Public Sub ExportPDF(Optional outputPath As String = "")
    On Error GoTo ErrHandler

    ConfigurePrintSettings True

    Dim tabNames() As String
    Dim tabCount As Long
    tabCount = GetPdfTabNames(tabNames)
    If tabCount = 0 Then
        MsgBox "No tabs configured for PDF output." & vbCrLf & vbCrLf & _
               "MANUAL BYPASS: Set IncludeInPDF=TRUE in print_config.csv, " & _
               "or use File -> Print -> Save as PDF.", _
               vbInformation, "RDK -- Export PDF"
        Exit Sub
    End If

    ' Determine output path
    If Len(outputPath) = 0 Then
        Dim parentDir As String
        Dim wbPath As String
        wbPath = ThisWorkbook.Path
        parentDir = Left(wbPath, InStrRev(wbPath, "\") - 1)
        Dim outDir As String
        outDir = parentDir & "\output"
        ' Create output directory if needed
        If Dir(outDir, vbDirectory) = "" Then
            MkDir outDir
        End If
        Dim modelName As String
        modelName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
        Dim ts As String
        ts = Format(Now, "yyyymmdd_hhnnss")
        outputPath = outDir & "\" & modelName & "_" & ts & ".pdf"
    End If

    ' Select all PDF tabs
    SelectMultipleTabs tabNames, tabCount

    ' Export as PDF
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=outputPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ' Restore selection
    ThisWorkbook.Sheets(tabNames(1)).Select

    KernelConfig.LogError SEV_INFO, "KernelPrint", "I-601", _
        "PDF exported successfully", outputPath

    MsgBox "PDF exported: " & outputPath, vbInformation, "RDK -- Export PDF"
    Exit Sub

ErrHandler:
    KernelConfig.LogError SEV_ERROR, "KernelPrint", "E-602", _
        "PDF export error: " & Err.Description, _
        "MANUAL BYPASS: Use File -> Print -> Save as PDF. Select the tabs you want."
    MsgBox "PDF export error: " & Err.Description & vbCrLf & vbCrLf & _
           "MANUAL BYPASS: Use File -> Print -> Save as PDF. Select the tabs you want.", _
           vbExclamation, "RDK -- Export PDF Error"
End Sub


' =============================================================================
' GetPdfTabNames (Private)
' Returns array of tab names where IncludeInPDF=TRUE, sorted by PrintOrder.
' =============================================================================
Private Function GetPdfTabNames(ByRef tabNames() As String) As Long
    Dim cnt As Long
    cnt = KernelConfig.GetPrintConfigCount()
    If cnt = 0 Then
        GetPdfTabNames = 0
        Exit Function
    End If

    ' Collect tabs with IncludeInPDF=TRUE
    Dim tmpNames() As String
    Dim tmpOrders() As Long
    ReDim tmpNames(1 To cnt)
    ReDim tmpOrders(1 To cnt)
    Dim found As Long
    found = 0

    Dim idx As Long
    For idx = 1 To cnt
        Dim incPdf As String
        incPdf = KernelConfig.GetPrintConfigField(idx, PRTCFG_COL_INCLUDEPDF)
        If StrComp(incPdf, "TRUE", vbTextCompare) = 0 Then
            Dim tName As String
            tName = KernelConfig.GetPrintConfigField(idx, PRTCFG_COL_TABNAME)
            ' Verify tab exists
            Dim wsCheck As Worksheet
            Set wsCheck = Nothing
            On Error Resume Next
            Set wsCheck = ThisWorkbook.Sheets(tName)
            On Error GoTo 0
            If Not wsCheck Is Nothing Then
                found = found + 1
                tmpNames(found) = tName
                Dim ordStr As String
                ordStr = KernelConfig.GetPrintConfigField(idx, PRTCFG_COL_PRINTORDER)
                If IsNumeric(ordStr) Then
                    tmpOrders(found) = CLng(ordStr)
                Else
                    tmpOrders(found) = 999
                End If
            End If
        End If
    Next idx

    If found = 0 Then
        GetPdfTabNames = 0
        Exit Function
    End If

    ' Simple bubble sort by PrintOrder
    Dim i As Long
    Dim j As Long
    For i = 1 To found - 1
        For j = 1 To found - i
            If tmpOrders(j) > tmpOrders(j + 1) Then
                Dim swpN As String
                swpN = tmpNames(j)
                tmpNames(j) = tmpNames(j + 1)
                tmpNames(j + 1) = swpN
                Dim swpO As Long
                swpO = tmpOrders(j)
                tmpOrders(j) = tmpOrders(j + 1)
                tmpOrders(j + 1) = swpO
            End If
        Next j
    Next i

    ReDim tabNames(1 To found)
    For i = 1 To found
        tabNames(i) = tmpNames(i)
    Next i
    GetPdfTabNames = found
End Function


' =============================================================================
' SelectMultipleTabs (Private)
' Selects multiple tabs for grouped print/export operations.
' =============================================================================
Private Sub SelectMultipleTabs(tabNames() As String, tabCount As Long)
    Dim sheetNames() As Variant
    ReDim sheetNames(1 To tabCount)
    Dim i As Long
    For i = 1 To tabCount
        sheetNames(i) = tabNames(i)
    Next i
    ThisWorkbook.Sheets(sheetNames).Select
End Sub
