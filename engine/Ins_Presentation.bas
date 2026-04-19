Attribute VB_Name = "Ins_Presentation"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' Ins_Presentation.bas
' Purpose: Insurance domain presentation layer. User Guide content, any other
'          domain-specific display logic. Called via branding_config dispatch.
' =============================================================================


' =============================================================================
' PopulateUserGuide
' Writes a static user guide for investors / co-founders receiving the
' insurance carrier financial model workbook cold. Six steps plus tips.
' Called from KernelBootstrap via Application.Run (branding_config UserGuideEntry).
' =============================================================================
Public Sub PopulateUserGuide()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_USER_GUIDE)
    If ws Is Nothing Then Exit Sub
    On Error GoTo 0

    ws.Cells.ClearContents

    ' Header
    ws.Cells(1, 2).Value = "User Guide"
    ws.Cells(1, 2).Font.Bold = True
    ws.Cells(1, 2).Font.Size = 14
    ws.Cells(1, 2).Font.Color = RGB(31, 56, 100)
    ws.Cells(2, 2).Value = "How to Use the Insurance NewCo Carrier Financial Model"
    ws.Cells(2, 2).Font.Size = 11

    ' Step 1
    ws.Cells(4, 2).Value = "STEP 1: Enter Your Programs"
    ws.Cells(4, 2).Font.Bold = True
    ws.Cells(4, 2).Interior.Color = RGB(217, 225, 242)
    ws.Cells(5, 2).Value = "Navigate to the UW Inputs tab. Enter up to 10 insurance programs."
    ws.Cells(6, 2).Value = "For each program, enter: name, line of business, policy term, gross written"
    ws.Cells(7, 2).Value = "premium by quarter (Y1-Y5), commission rates, QS cession rates, ELR, and"
    ws.Cells(8, 2).Value = "trend levels for loss and count development patterns."

    ' Step 2
    ws.Cells(10, 2).Value = "STEP 2: Enter Capital"
    ws.Cells(10, 2).Font.Bold = True
    ws.Cells(10, 2).Interior.Color = RGB(217, 225, 242)
    ws.Cells(11, 2).Value = "Navigate to the Capital Activity tab. Enter equity raises and/or surplus"
    ws.Cells(12, 2).Value = "note draws by quarter. Enter interest rates for debt instruments."

    ' Step 3
    ws.Cells(14, 2).Value = "STEP 3: Enter Operating Expenses"
    ws.Cells(14, 2).Font.Bold = True
    ws.Cells(14, 2).Interior.Color = RGB(217, 225, 242)
    ws.Cells(15, 2).Value = "Navigate to the Staffing Expense tab. Enter headcount and salary by"
    ws.Cells(16, 2).Value = "department for each year. Navigate to the Other Expense Detail tab."
    ws.Cells(17, 2).Value = "Enter non-staffing expenses (benefits, rent, travel, tech, etc.) by year."

    ' Step 4
    ws.Cells(19, 2).Value = "STEP 4: Enter Revenue Assumptions"
    ws.Cells(19, 2).Font.Bold = True
    ws.Cells(19, 2).Interior.Color = RGB(217, 225, 242)
    ws.Cells(20, 2).Value = "Navigate to the Other Revenue Detail tab. Enter software revenue by type,"
    ws.Cells(21, 2).Value = "fee income, and consulting revenue by quarter."
    ws.Cells(22, 2).Value = "Navigate to the Investments tab. Set asset allocation percentages and yields."

    ' Step 5
    ws.Cells(24, 2).Value = "STEP 5: Run the Model"
    ws.Cells(24, 2).Font.Bold = True
    ws.Cells(24, 2).Interior.Color = RGB(217, 225, 242)
    ws.Cells(25, 2).Value = "Return to the Dashboard tab. Click 'Run Model'. The model will compute"
    ws.Cells(26, 2).Value = "loss development, quarterly aggregation, and all financial statements."
    ws.Cells(27, 2).Value = "This typically takes 10-30 seconds depending on the number of programs."

    ' Step 6
    ws.Cells(29, 2).Value = "STEP 6: Review Results"
    ws.Cells(29, 2).Font.Bold = True
    ws.Cells(29, 2).Interior.Color = RGB(217, 225, 242)
    ws.Cells(30, 2).Value = "UW Exec Summary -- Portfolio underwriting P&L waterfall"
    ws.Cells(31, 2).Value = "UW Program Detail -- Per-program breakdown with loss development"
    ws.Cells(32, 2).Value = "Revenue Summary -- All revenue sources (UW + investment + software + fees)"
    ws.Cells(33, 2).Value = "Expense Summary -- UW expenses + operating expenses from detail tabs"
    ws.Cells(34, 2).Value = "Income Statement -- Full P&L with key ratios and growth rates"
    ws.Cells(35, 2).Value = "Balance Sheet -- Assets, liabilities, equity with balance check"
    ws.Cells(36, 2).Value = "Cash Flow Statement -- Indirect method with reconciliation check"

    ' Tips
    ws.Cells(38, 2).Value = "TIPS"
    ws.Cells(38, 2).Font.Bold = True
    ws.Cells(38, 2).Interior.Color = RGB(217, 225, 242)
    ws.Cells(39, 2).Value = "- Blue cells are inputs. Grey cells are computed. Do not edit grey cells."
    ws.Cells(40, 2).Value = "- Use Snapshots (Dashboard) to save/restore different scenarios."
    ws.Cells(41, 2).Value = "- Use Export PDF (Dashboard) to create a shareable report."
    ws.Cells(42, 2).Value = "- BS Balance Check and CFS Reconciliation should always show 0 (green)."
    ws.Cells(43, 2).Value = "- The Sales Funnel tab helps plan your pipeline before entering programs."
    ws.Cells(44, 2).Value = "- The Curve Reference tab shows loss development patterns by trend level."

    ' Formatting
    ws.Columns(1).ColumnWidth = 5
    ws.Columns(2).ColumnWidth = 90
    Dim rw As Long
    For rw = 5 To 44
        ws.Cells(rw, 2).WrapText = True
    Next rw

    ' No gridlines
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
End Sub
