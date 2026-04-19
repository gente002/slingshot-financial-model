Attribute VB_Name = "Ext_ReportGen"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' Ext_ReportGen.bas
' Purpose: PDF report generation from model results. Creates a cover page
'          with TOC and Prove-It summary, then exports selected tabs to PDF.
'          HookType: PostOutput -- runs after all output is written.
'          Leverages KernelPrint.ExportPDF for PDF mechanics. ReportGen adds
'          cover page, TOC, and Prove-It summary orchestration.
' Phase 6A extension. All errors include manual bypass instructions (AP-46).
' =============================================================================

' Module-level silent flag (PT-027: pipeline vs UI entry points)
Private m_silent As Boolean


' =============================================================================
' ReportGen_Execute
' Entry point called by KernelExtension.RunExtensions("PostOutput").
' Silent during pipeline -- no MsgBox. Generates PDF report.
' =============================================================================
Public Sub ReportGen_Execute()
    m_silent = True

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    ' Check if report generation is enabled
    Dim includeCover As String
    includeCover = KernelConfig.GetReportSetting("IncludeCoverPage")
    If Len(includeCover) = 0 Then includeCover = "TRUE"

    ' Determine output path
    Dim outputPath As String
    outputPath = BuildOutputPath()
    If Len(outputPath) = 0 Then
        KernelConfig.LogError SEV_WARN, "Ext_ReportGen", "W-900", _
            "Could not determine output path for report.", _
            "MANUAL BYPASS: Use File -> Print -> Save as PDF. Select tabs manually."
        Exit Sub
    End If

    ' Create cover page if configured
    Dim wsCover As Worksheet
    Set wsCover = Nothing
    If StrComp(includeCover, "TRUE", vbTextCompare) = 0 Then
        Set wsCover = CreateCoverSheet()
    End If

    ' Export PDF using KernelPrint pattern
    ExportReportPDF outputPath, wsCover

    ' Clean up cover page
    If Not wsCover Is Nothing Then
        Application.DisplayAlerts = False
        wsCover.Delete
        Application.DisplayAlerts = True
    End If

    KernelConfig.LogError SEV_INFO, "Ext_ReportGen", "I-900", _
        "Report generated: " & outputPath, ""

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    ' Clean up cover page on error
    If Not wsCover Is Nothing Then
        On Error Resume Next
        Application.DisplayAlerts = False
        wsCover.Delete
        Application.DisplayAlerts = True
        On Error GoTo 0
    End If

    Application.ScreenUpdating = True
    KernelConfig.LogError SEV_ERROR, "Ext_ReportGen", "E-900", _
        "Report generation failed: " & Err.Description, _
        "MANUAL BYPASS: Use File -> Print -> Save as PDF. Select tabs manually."
End Sub


' =============================================================================
' GenerateReport
' Public entry point for manual report generation (Dashboard button).
' Shows MsgBox with report path on completion (AP-53).
' =============================================================================
Public Sub GenerateReport(Optional outputPath As String = "")
    m_silent = False

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    ' Ensure config is loaded
    If KernelConfig.GetColumnCount() = 0 Then
        KernelConfig.LoadAllConfig
    End If
    KernelExtension.LoadExtensionRegistry

    ' Check if report generation is enabled
    Dim includeCover As String
    includeCover = KernelConfig.GetReportSetting("IncludeCoverPage")
    If Len(includeCover) = 0 Then includeCover = "TRUE"

    ' Determine output path
    If Len(outputPath) = 0 Then
        outputPath = BuildOutputPath()
    End If
    If Len(outputPath) = 0 Then
        MsgBox "Could not determine output path for report." & vbCrLf & vbCrLf & _
               "MANUAL BYPASS: Use File -> Print -> Save as PDF.", _
               vbExclamation, "RDK -- Report Generation"
        Exit Sub
    End If

    ' Create cover page if configured
    Dim wsCover As Worksheet
    Set wsCover = Nothing
    If StrComp(includeCover, "TRUE", vbTextCompare) = 0 Then
        Set wsCover = CreateCoverSheet()
    End If

    ' Export PDF
    ExportReportPDF outputPath, wsCover

    ' Clean up cover page
    If Not wsCover Is Nothing Then
        Application.DisplayAlerts = False
        wsCover.Delete
        Application.DisplayAlerts = True
    End If

    KernelConfig.LogError SEV_INFO, "Ext_ReportGen", "I-901", _
        "Report generated via Dashboard: " & outputPath, ""

    Application.ScreenUpdating = True

    MsgBox "Report generated successfully." & vbCrLf & vbCrLf & _
           "Path: " & outputPath, _
           vbInformation, "RDK -- Report Generated"

    Exit Sub

ErrHandler:
    ' Clean up cover page on error
    If Not wsCover Is Nothing Then
        On Error Resume Next
        Application.DisplayAlerts = False
        wsCover.Delete
        Application.DisplayAlerts = True
        On Error GoTo 0
    End If

    Application.ScreenUpdating = True

    KernelConfig.LogError SEV_ERROR, "Ext_ReportGen", "E-901", _
        "Report generation failed: " & Err.Description, _
        "MANUAL BYPASS: Use File -> Print -> Save as PDF. Select tabs manually."

    MsgBox "Report generation failed:" & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "MANUAL BYPASS: Use File -> Print -> Save as PDF.", _
           vbCritical, "RDK -- Report Error"
End Sub


' =============================================================================
' GenerateCoverPage
' Creates a formatted cover page on a temporary sheet.
' Public so it can be called independently.
' =============================================================================
Public Sub GenerateCoverPage(ws As Worksheet)
    ws.Cells.ClearContents

    ' Title
    Dim reportTitle As String
    reportTitle = KernelConfig.GetReportSetting("ReportTitle")
    If Len(reportTitle) = 0 Then reportTitle = "RDK Model Report"

    ws.Cells(1, 1).Value = reportTitle
    With ws.Cells(1, 1)
        .Font.Size = 24
        .Font.Bold = True
        .Font.Color = RGB(68, 114, 196)
    End With

    ' Separator line
    ws.Range("A2:F2").Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range("A2:F2").Borders(xlEdgeBottom).Weight = xlMedium
    ws.Range("A2:F2").Borders(xlEdgeBottom).Color = RGB(68, 114, 196)

    ' Report metadata
    Dim curRow As Long
    curRow = 4

    ws.Cells(curRow, 1).Value = "Report Generated:"
    ws.Cells(curRow, 1).Font.Bold = True
    ws.Cells(curRow, 2).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    curRow = curRow + 1

    ws.Cells(curRow, 1).Value = "Kernel Version:"
    ws.Cells(curRow, 1).Font.Bold = True
    ws.Cells(curRow, 2).Value = KERNEL_VERSION
    curRow = curRow + 1

    ' Active scenario
    ws.Cells(curRow, 1).Value = "Scenario:"
    ws.Cells(curRow, 1).Font.Bold = True
    ws.Cells(curRow, 2).Value = "Base"
    curRow = curRow + 2

    ' Table of Contents
    ws.Cells(curRow, 1).Value = "Table of Contents"
    With ws.Cells(curRow, 1)
        .Font.Size = 14
        .Font.Bold = True
    End With
    curRow = curRow + 1

    ' List tabs included in report
    Dim prtCnt As Long
    prtCnt = KernelConfig.GetPrintConfigCount()

    Dim tocNum As Long
    tocNum = 0

    Dim pIdx As Long
    For pIdx = 1 To prtCnt
        Dim incPdf As String
        incPdf = KernelConfig.GetPrintConfigField(pIdx, PRTCFG_COL_INCLUDEPDF)
        If StrComp(incPdf, "TRUE", vbTextCompare) = 0 Then
            Dim tabName As String
            tabName = KernelConfig.GetPrintConfigField(pIdx, PRTCFG_COL_TABNAME)
            ' Check tab exists
            Dim wsCheck As Worksheet
            Set wsCheck = Nothing
            On Error Resume Next
            Set wsCheck = ThisWorkbook.Sheets(tabName)
            On Error GoTo 0
            If Not wsCheck Is Nothing Then
                tocNum = tocNum + 1
                ws.Cells(curRow, 1).Value = CStr(tocNum) & "."
                ws.Cells(curRow, 2).Value = tabName
                ws.Cells(curRow, 2).IndentLevel = 1
                curRow = curRow + 1
            End If
        End If
    Next pIdx

    If tocNum = 0 Then
        ws.Cells(curRow, 1).Value = "(No tabs configured for PDF output)"
        ws.Cells(curRow, 1).Font.Italic = True
        curRow = curRow + 1
    End If

    curRow = curRow + 1

    ' Prove-It summary
    Dim includePI As String
    includePI = KernelConfig.GetReportSetting("IncludeProveItSummary")
    If Len(includePI) = 0 Then includePI = "TRUE"

    If StrComp(includePI, "TRUE", vbTextCompare) = 0 Then
        ws.Cells(curRow, 1).Value = "Prove-It Summary"
        With ws.Cells(curRow, 1)
            .Font.Size = 14
            .Font.Bold = True
        End With
        curRow = curRow + 1

        ' Count passing/total checks
        Dim piTotal As Long
        piTotal = KernelConfig.GetProveItCheckCount()
        Dim piPass As Long
        piPass = 0

        If piTotal > 0 Then
            Dim wsPI As Worksheet
            Set wsPI = Nothing
            On Error Resume Next
            Set wsPI = ThisWorkbook.Sheets(TAB_PROVE_IT)
            On Error GoTo 0

            If Not wsPI Is Nothing Then
                ' Count PASS results in ProveIt tab (column E, starting row 5)
                Dim piRow As Long
                For piRow = 5 To 5 + piTotal - 1
                    If StrComp(CStr(wsPI.Cells(piRow, 5).Value), "TRUE", vbTextCompare) = 0 Then
                        piPass = piPass + 1
                    End If
                Next piRow
            End If

            ws.Cells(curRow, 1).Value = piPass & " of " & piTotal & " checks passing"
            If piPass = piTotal Then
                ws.Cells(curRow, 1).Font.Color = RGB(0, 128, 0)
                ws.Cells(curRow, 1).Font.Bold = True
            Else
                ws.Cells(curRow, 1).Font.Color = RGB(192, 0, 0)
                ws.Cells(curRow, 1).Font.Bold = True
            End If
        Else
            ws.Cells(curRow, 1).Value = "No Prove-It checks configured"
            ws.Cells(curRow, 1).Font.Italic = True
        End If
        curRow = curRow + 2
    End If

    ' Footer
    curRow = curRow + 2
    ws.Cells(curRow, 1).Value = "Generated by Insurance NewCo RDK v" & KERNEL_VERSION
    ws.Cells(curRow, 1).Font.Color = RGB(128, 128, 128)
    ws.Cells(curRow, 1).Font.Italic = True

    ' Format column widths
    ws.Columns(1).ColumnWidth = 25
    ws.Columns(2).ColumnWidth = 40

    ' Constrain to single page to prevent blank page 2 (BUG-046)
    With ws.PageSetup
        .PrintArea = "A1:B" & curRow
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = Application.InchesToPoints(1#)
        .RightMargin = Application.InchesToPoints(1#)
        .TopMargin = Application.InchesToPoints(1#)
        .BottomMargin = Application.InchesToPoints(1#)
    End With
End Sub


' =============================================================================
' Private Helpers
' =============================================================================

' -----------------------------------------------------------------------------
' BuildOutputPath
' Constructs the PDF output path using report_config settings.
' -----------------------------------------------------------------------------
Private Function BuildOutputPath() As String
    Dim customDir As String
    customDir = KernelConfig.GetReportSetting("OutputDirectory")

    Dim fullDir As String
    If Len(customDir) > 0 Then
        ' Custom output directory from report_config
        Dim wbPath As String
        wbPath = ThisWorkbook.Path
        Dim parentDir As String
        parentDir = Left(wbPath, InStrRev(wbPath, "\") - 1)
        fullDir = parentDir & "\" & customDir
        On Error Resume Next
        If Dir(fullDir, vbDirectory) = "" Then MkDir fullDir
        On Error GoTo 0
    Else
        ' Default: use shared output directory
        fullDir = KernelFormHelpers.EnsureOutputDir()
    End If

    ' Build filename
    Dim modelName As String
    modelName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)

    Dim includeTs As String
    includeTs = KernelConfig.GetReportSetting("IncludeTimestamp")
    If Len(includeTs) = 0 Then includeTs = "TRUE"

    Dim fileName As String
    If StrComp(includeTs, "TRUE", vbTextCompare) = 0 Then
        fileName = modelName & "_Report_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    Else
        fileName = modelName & "_Report.pdf"
    End If

    BuildOutputPath = fullDir & "\" & fileName
End Function


' -----------------------------------------------------------------------------
' CreateCoverSheet
' Creates a temporary cover page sheet and populates it.
' Returns the worksheet reference.
' -----------------------------------------------------------------------------
Private Function CreateCoverSheet() As Worksheet
    Dim ws As Worksheet
    Application.DisplayAlerts = False

    ' Delete existing cover sheet if present
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("_ReportCover")
    If Not ws Is Nothing Then
        ws.Delete
    End If
    On Error GoTo 0

    Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    ws.Name = "_ReportCover"

    GenerateCoverPage ws

    Application.DisplayAlerts = True
    Set CreateCoverSheet = ws
End Function


' -----------------------------------------------------------------------------
' ExportReportPDF
' Exports selected tabs to PDF. Cover page is first if provided.
' Uses ExportAsFixedFormat on selected sheets.
' -----------------------------------------------------------------------------
Private Sub ExportReportPDF(outputPath As String, Optional wsCover As Worksheet = Nothing)
    ' Build tab list for PDF
    Dim tabNames() As String
    Dim tabCount As Long
    tabCount = 0

    ' Start with cover page if provided
    If Not wsCover Is Nothing Then
        tabCount = 1
        ReDim tabNames(1 To 1)
        tabNames(1) = wsCover.Name
    End If

    ' Add tabs from print_config where IncludeInPDF=TRUE, sorted by PrintOrder
    Dim prtCnt As Long
    prtCnt = KernelConfig.GetPrintConfigCount()

    ' Pre-count eligible tabs to avoid ReDim Preserve in loop (AP-18)
    Dim eligTabs() As String
    Dim eligOrders() As Long
    Dim eligCount As Long
    eligCount = 0

    Dim pIdx As Long
    Dim incPdf As String
    Dim tn As String
    Dim wsChk As Worksheet
    Dim poStr As String

    For pIdx = 1 To prtCnt
        incPdf = KernelConfig.GetPrintConfigField(pIdx, PRTCFG_COL_INCLUDEPDF)
        If StrComp(incPdf, "TRUE", vbTextCompare) = 0 Then
            tn = KernelConfig.GetPrintConfigField(pIdx, PRTCFG_COL_TABNAME)
            Set wsChk = Nothing
            On Error Resume Next
            Set wsChk = ThisWorkbook.Sheets(tn)
            On Error GoTo 0
            If Not wsChk Is Nothing Then eligCount = eligCount + 1
        End If
    Next pIdx

    If eligCount = 0 And tabCount = 0 Then
        KernelConfig.LogError SEV_WARN, "Ext_ReportGen", "W-901", _
            "No tabs eligible for PDF export.", _
            "MANUAL BYPASS: Export PDF manually via File > Save As > PDF."
        Exit Sub
    End If

    ' Allocate once, fill in second pass
    If eligCount > 0 Then
        ReDim eligTabs(1 To eligCount)
        ReDim eligOrders(1 To eligCount)
    End If

    Dim fillIdx As Long
    fillIdx = 0
    For pIdx = 1 To prtCnt
        incPdf = KernelConfig.GetPrintConfigField(pIdx, PRTCFG_COL_INCLUDEPDF)
        If StrComp(incPdf, "TRUE", vbTextCompare) = 0 Then
            tn = KernelConfig.GetPrintConfigField(pIdx, PRTCFG_COL_TABNAME)
            Set wsChk = Nothing
            On Error Resume Next
            Set wsChk = ThisWorkbook.Sheets(tn)
            On Error GoTo 0
            If Not wsChk Is Nothing Then
                fillIdx = fillIdx + 1
                eligTabs(fillIdx) = tn
                poStr = KernelConfig.GetPrintConfigField(pIdx, PRTCFG_COL_PRINTORDER)
                If IsNumeric(poStr) Then
                    eligOrders(fillIdx) = CLng(poStr)
                Else
                    eligOrders(fillIdx) = 999
                End If
            End If
        End If
    Next pIdx

    ' Sort by PrintOrder (bubble sort)
    If eligCount > 1 Then
        Dim i As Long
        Dim j As Long
        For i = 1 To eligCount - 1
            For j = 1 To eligCount - i
                If eligOrders(j) > eligOrders(j + 1) Then
                    Dim tmpS As String
                    tmpS = eligTabs(j)
                    eligTabs(j) = eligTabs(j + 1)
                    eligTabs(j + 1) = tmpS
                    Dim tmpL As Long
                    tmpL = eligOrders(j)
                    eligOrders(j) = eligOrders(j + 1)
                    eligOrders(j + 1) = tmpL
                End If
            Next j
        Next i
    End If

    ' Append eligible tabs to tabNames (single allocation, no ReDim Preserve in loop)
    If eligCount > 0 Then
        Dim newCount As Long
        newCount = tabCount + eligCount
        If tabCount = 0 Then
            ReDim tabNames(1 To newCount)
        Else
            ReDim Preserve tabNames(1 To newCount)
        End If
        Dim k As Long
        For k = 1 To eligCount
            tabNames(tabCount + k) = eligTabs(k)
        Next k
        tabCount = newCount
    End If

    If tabCount = 0 Then
        KernelConfig.LogError SEV_WARN, "Ext_ReportGen", "W-901", _
            "No tabs available for PDF export.", _
            "MANUAL BYPASS: Set IncludeInPDF=TRUE in print_config.csv for tabs to include."
        If Not m_silent Then
            MsgBox "No tabs configured for PDF output." & vbCrLf & vbCrLf & _
                   "MANUAL BYPASS: Set IncludeInPDF=TRUE in print_config.csv.", _
                   vbInformation, "RDK -- Report"
        End If
        Exit Sub
    End If

    ' Apply print config (orientation, fit-to-page, margins) to all tabs (BUG-048)
    KernelPrint.ConfigurePrintSettings True

    ' Select all tabs for PDF export
    Dim sheetArr() As Variant
    ReDim sheetArr(1 To tabCount)
    For k = 1 To tabCount
        sheetArr(k) = tabNames(k)
    Next k

    ThisWorkbook.Sheets(sheetArr).Select

    ' Export as PDF
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=outputPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ' Restore selection to first real tab (skip cover)
    If Not wsCover Is Nothing And tabCount > 1 Then
        ThisWorkbook.Sheets(tabNames(2)).Select
    ElseIf tabCount > 0 Then
        ThisWorkbook.Sheets(tabNames(1)).Select
    End If
End Sub
