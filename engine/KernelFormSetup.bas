Attribute VB_Name = "KernelFormSetup"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelFormSetup.bas
' Purpose: Programmatically creates UserForms during bootstrap.
'          Called from BootstrapWorkbook when VBProject access is available.
'          Forms are saved into the .xlsm permanently after creation.
' =============================================================================

' =============================================================================
' CreateAllForms -- Main entry point. Idempotent: skips existing forms.
' =============================================================================
Public Sub CreateAllForms()
    On Error GoTo ErrHandler

    ' Clean up legacy forms that have been replaced by SnapshotExplorer
    If FormExists("ScenarioExplorer") Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents("ScenarioExplorer")
    End If
    If FormExists("SavepointExplorer") Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents("SavepointExplorer")
    End If
    If FormExists("ArchiveExplorer") Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents("ArchiveExplorer")
    End If

    ' Recreate SnapshotExplorer if outdated (4-col -> 7-col upgrade)
    If FormExists("SnapshotExplorer") Then
        If Not FormHasControl("SnapshotExplorer", "btnArchiveAll") Then
            ThisWorkbook.VBProject.VBComponents.Remove _
                ThisWorkbook.VBProject.VBComponents("SnapshotExplorer")
        End If
    End If

    ' SnapshotExplorer deprecated -- workspaces replace snapshots
    ' If Not FormExists("SnapshotExplorer") Then CreateSnapshotExplorer
    If Not FormExists("CompareExplorer") Then CreateCompareExplorer
    If Not FormExists("ReportExplorer") Then CreateReportExplorer
    If Not FormExists("WorkspaceExplorer") Then KernelFormSetup2.CreateWorkspaceExplorer
    If Not FormExists("MoreActions") Then KernelFormSetup2.CreateMoreActions
    If Not FormExists("AssumptionExplorer") Then KernelFormSetup2.CreateAssumptionExplorer

    Exit Sub
ErrHandler:
    Debug.Print "KernelFormSetup.CreateAllForms error: " & Err.Description
End Sub

Public Function FormExists(ByVal formName As String) As Boolean
    Dim comp As Object
    On Error Resume Next
    Set comp = ThisWorkbook.VBProject.VBComponents(formName)
    FormExists = (Not comp Is Nothing)
    On Error GoTo 0
End Function

Private Function FormHasControl(ByVal formName As String, ByVal ctrlName As String) As Boolean
    Dim comp As Object
    On Error Resume Next
    Set comp = ThisWorkbook.VBProject.VBComponents(formName)
    If comp Is Nothing Then
        FormHasControl = False
        Exit Function
    End If
    Dim ctrl As Object
    Set ctrl = comp.designer.Controls(ctrlName)
    FormHasControl = (Not ctrl Is Nothing)
    On Error GoTo 0
End Function

' --- Layout helpers --------------------------------------------------------

Public Sub AddButton(ByVal dsgn As Object, ByVal ctrlName As String, _
                      ByVal cap As String, ByVal L As Single, _
                      ByVal T As Single, ByVal W As Single, ByVal H As Single)
    Dim ctrl As Object
    Set ctrl = dsgn.Controls.Add("Forms.CommandButton.1", ctrlName)
    ctrl.Left = L: ctrl.Top = T: ctrl.Width = W: ctrl.Height = H
    ctrl.caption = cap
End Sub

Public Sub AddListBox(ByVal dsgn As Object, ByVal ctrlName As String, _
                       ByVal L As Single, ByVal T As Single, _
                       ByVal W As Single, ByVal H As Single, _
                       ByVal colCount As Long, ByVal colWidths As String)
    Dim ctrl As Object
    Set ctrl = dsgn.Controls.Add("Forms.ListBox.1", ctrlName)
    ctrl.Left = L: ctrl.Top = T: ctrl.Width = W: ctrl.Height = H
    ctrl.ColumnCount = colCount
    ctrl.ColumnWidths = colWidths
    ctrl.ColumnHeads = False
End Sub

Public Sub AddLabel(ByVal dsgn As Object, ByVal ctrlName As String, _
                     ByVal cap As String, ByVal L As Single, _
                     ByVal T As Single, ByVal W As Single, ByVal H As Single)
    Dim ctrl As Object
    Set ctrl = dsgn.Controls.Add("Forms.Label.1", ctrlName)
    ctrl.Left = L: ctrl.Top = T: ctrl.Width = W: ctrl.Height = H
    ctrl.caption = cap
End Sub

Public Sub AddOptionButton(ByVal dsgn As Object, ByVal ctrlName As String, _
                            ByVal cap As String, ByVal L As Single, _
                            ByVal T As Single, ByVal W As Single, _
                            ByVal H As Single)
    Dim ctrl As Object
    Set ctrl = dsgn.Controls.Add("Forms.OptionButton.1", ctrlName)
    ctrl.Left = L: ctrl.Top = T: ctrl.Width = W: ctrl.Height = H
    ctrl.caption = cap
End Sub

Public Sub AddCheckBox(ByVal dsgn As Object, ByVal ctrlName As String, _
                        ByVal cap As String, ByVal L As Single, _
                        ByVal T As Single, ByVal W As Single, _
                        ByVal H As Single)
    Dim ctrl As Object
    Set ctrl = dsgn.Controls.Add("Forms.CheckBox.1", ctrlName)
    ctrl.Left = L: ctrl.Top = T: ctrl.Width = W: ctrl.Height = H
    ctrl.caption = cap
End Sub

Public Sub AddTextBox(ByVal dsgn As Object, ByVal ctrlName As String, _
                       ByVal L As Single, ByVal T As Single, _
                       ByVal W As Single, ByVal H As Single)
    Dim ctrl As Object
    Set ctrl = dsgn.Controls.Add("Forms.TextBox.1", ctrlName)
    ctrl.Left = L: ctrl.Top = T: ctrl.Width = W: ctrl.Height = H
End Sub

Public Sub AddComboBox(ByVal dsgn As Object, ByVal ctrlName As String, _
                        ByVal L As Single, ByVal T As Single, _
                        ByVal W As Single, ByVal H As Single)
    Dim ctrl As Object
    Set ctrl = dsgn.Controls.Add("Forms.ComboBox.1", ctrlName)
    ctrl.Left = L: ctrl.Top = T: ctrl.Width = W: ctrl.Height = H
End Sub

Public Function IL(ByVal codeMod As Object, ByVal n As Long, _
                    ByVal codeLine As String) As Long
    codeMod.InsertLines n, codeLine
    IL = n + 1
End Function

' =============================================================================
' FORM 1: SnapshotExplorer
' Unified snapshot management: Active/Archived toggle, filter, sort,
' Save New, Load All, Load Inputs, Delete, Archive/Restore, Close
' =============================================================================
Private Sub CreateSnapshotExplorer()
    Dim comp As Object
    Set comp = ThisWorkbook.VBProject.VBComponents.Add(3)
    comp.Name = "SnapshotExplorer"
    Dim d As Object: Set d = comp.designer
    comp.Properties("Caption") = "RDK -- Snapshots"
    comp.Properties("Width") = 620
    comp.Properties("Height") = 450

    ' View toggle
    AddOptionButton d, "optActive", "Active", 12, 8, 80, 18
    AddOptionButton d, "optArchived", "Archived", 100, 8, 80, 18

    ' Filter
    AddLabel d, "lblFilter", "Filter:", 260, 10, 40, 18
    AddTextBox d, "txtFilter", 304, 8, 200, 20

    ' Sort buttons
    AddButton d, "btnSortName", "Sort: Name", 12, 32, 80, 22
    AddButton d, "btnSortDate", "Sort: Date", 100, 32, 80, 22

    ' Column headers
    Dim hx As Single: hx = 14
    Dim hw() As Variant: hw = Array(160, 110, 50, 120, 35, 50, 35)
    Dim hn() As Variant: hn = Array("Name", "Date", "Elapsed", "Description", "Stale", "Status", "Files")
    Dim hi As Long
    For hi = 0 To 6
        AddLabel d, "lblH" & hi, CStr(hn(hi)), hx, 58, CSng(hw(hi)), 14
        Dim hCtrl As Object
        Set hCtrl = d.Controls("lblH" & hi)
        hCtrl.Font.Size = 8
        hCtrl.Font.Bold = True
        hx = hx + CSng(hw(hi))
    Next hi
    ' List box
    AddListBox d, "lstItems", 12, 73, 590, 225, 7, "160;110;50;120;35;50;35"

    ' Action buttons (row 1)
    AddButton d, "btnSaveNew", "Save New", 12, 306, 80, 28
    AddButton d, "btnLoadAll", "Load All", 100, 306, 80, 28
    AddButton d, "btnLoadInputs", "Load Inputs", 188, 306, 90, 28
    AddButton d, "btnRename", "Rename", 286, 306, 80, 28
    AddButton d, "btnDelete", "Delete", 374, 306, 80, 28
    ' Action buttons (row 2)
    AddButton d, "btnArchive", "Archive", 12, 340, 80, 28
    AddButton d, "btnArchiveAll", "Archive All", 100, 340, 80, 28
    AddButton d, "btnEditDesc", "Edit Desc", 188, 340, 80, 28

    AddSnapshotExplorerCode comp.CodeModule
End Sub

Private Sub AddSnapshotExplorerCode(ByVal cm As Object)
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    Dim n As Long: n = 1
    n = IL(cm, n, "Option Explicit")
    n = IL(cm, n, "Private m_sortByDate As Boolean")
    n = IL(cm, n, "Private m_sortAsc As Boolean")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub UserForm_Initialize()")
    n = IL(cm, n, "    m_sortAsc = True")
    n = IL(cm, n, "    Me.optActive.Value = True")
    n = IL(cm, n, "    RefreshList")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub optActive_Click()")
    n = IL(cm, n, "    RefreshList")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub optArchived_Click()")
    n = IL(cm, n, "    RefreshList")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub txtFilter_Change()")
    n = IL(cm, n, "    RefreshList")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub btnSortName_Click()")
    n = IL(cm, n, "    If m_sortByDate = False Then")
    n = IL(cm, n, "        m_sortAsc = Not m_sortAsc")
    n = IL(cm, n, "    Else")
    n = IL(cm, n, "        m_sortByDate = False")
    n = IL(cm, n, "        m_sortAsc = True")
    n = IL(cm, n, "    End If")
    n = IL(cm, n, "    RefreshList")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub btnSortDate_Click()")
    n = IL(cm, n, "    If m_sortByDate = True Then")
    n = IL(cm, n, "        m_sortAsc = Not m_sortAsc")
    n = IL(cm, n, "    Else")
    n = IL(cm, n, "        m_sortByDate = True")
    n = IL(cm, n, "        m_sortAsc = True")
    n = IL(cm, n, "    End If")
    n = IL(cm, n, "    RefreshList")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub RefreshList()")
    n = IL(cm, n, "    Dim isActive As Boolean")
    n = IL(cm, n, "    isActive = Me.optActive.Value")
    n = IL(cm, n, "    Me.btnSaveNew.Visible = isActive")
    n = IL(cm, n, "    Me.btnLoadAll.Visible = isActive")
    n = IL(cm, n, "    Me.btnLoadInputs.Visible = isActive")
    n = IL(cm, n, "    Me.btnRename.Visible = isActive")
    n = IL(cm, n, "    Me.btnEditDesc.Visible = isActive")
    n = IL(cm, n, "    Me.btnArchiveAll.Visible = isActive")
    n = IL(cm, n, "    Me.btnDelete.Visible = isActive")
    n = IL(cm, n, "    If isActive Then")
    n = IL(cm, n, "        Me.btnArchive.Caption = ""Archive""")
    n = IL(cm, n, "        KernelFormHelpers.PopulateSnapshotListBox Me.lstItems, m_sortByDate, m_sortAsc")
    n = IL(cm, n, "    Else")
    n = IL(cm, n, "        Me.btnArchive.Caption = ""Restore""")
    n = IL(cm, n, "        KernelFormHelpers.PopulateArchivedSnapshotListBox Me.lstItems, m_sortByDate, m_sortAsc")
    n = IL(cm, n, "    End If")
    n = IL(cm, n, "    ApplyFilter")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub ApplyFilter()")
    n = IL(cm, n, "    Dim filterText As String")
    n = IL(cm, n, "    filterText = LCase(Trim(Me.txtFilter.Text))")
    n = IL(cm, n, "    If Len(filterText) = 0 Then Exit Sub")
    n = IL(cm, n, "    Dim i As Long")
    n = IL(cm, n, "    For i = Me.lstItems.ListCount - 1 To 0 Step -1")
    n = IL(cm, n, "        If InStr(1, LCase(Me.lstItems.List(i, 0)), filterText, vbTextCompare) = 0 Then")
    n = IL(cm, n, "            Me.lstItems.RemoveItem i")
    n = IL(cm, n, "        End If")
    n = IL(cm, n, "    Next i")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Function SelectedName() As String")
    n = IL(cm, n, "    If Me.lstItems.ListIndex < 0 Then")
    n = IL(cm, n, "        MsgBox ""Select an item first."", vbExclamation, ""RDK""")
    n = IL(cm, n, "        SelectedName = """"")
    n = IL(cm, n, "    Else")
    n = IL(cm, n, "        SelectedName = Me.lstItems.List(Me.lstItems.ListIndex, 0)")
    n = IL(cm, n, "    End If")
    n = IL(cm, n, "End Function")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub btnSaveNew_Click()")
    n = IL(cm, n, "    Dim baseName As String")
    n = IL(cm, n, "    baseName = InputBox(""Enter snapshot name:"", ""RDK -- Save Snapshot"")")
    n = IL(cm, n, "    If Len(Trim(baseName)) = 0 Then Exit Sub")
    n = IL(cm, n, "    Dim addTS As VbMsgBoxResult")
    n = IL(cm, n, "    addTS = MsgBox(""Append date/time to name?"", vbYesNo Or vbQuestion Or vbDefaultButton2, ""RDK"")")
    n = IL(cm, n, "    If addTS = vbYes Then baseName = KernelFormHelpers.AppendTimestamp(baseName)")
    n = IL(cm, n, "    Dim descr As String")
    n = IL(cm, n, "    descr = InputBox(""Description (optional):"", ""RDK -- Save Snapshot"")")
    n = IL(cm, n, "    KernelSnapshot.SaveSnapshot baseName, descr")
    n = IL(cm, n, "    RefreshList")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub btnLoadAll_Click()")
    n = IL(cm, n, "    Dim nm As String: nm = SelectedName()")
    n = IL(cm, n, "    If Len(nm) = 0 Then Exit Sub")
    n = IL(cm, n, "    KernelSnapshot.LoadSnapshot nm")
    n = IL(cm, n, "    MsgBox ""Loaded: "" & nm, vbInformation, ""RDK""")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub btnLoadInputs_Click()")
    n = IL(cm, n, "    Dim nm As String: nm = SelectedName()")
    n = IL(cm, n, "    If Len(nm) = 0 Then Exit Sub")
    n = IL(cm, n, "    KernelSnapshot.LoadSnapshotInputsOnly nm")
    n = IL(cm, n, "    MsgBox ""Inputs loaded: "" & nm, vbInformation, ""RDK""")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub btnDelete_Click()")
    n = IL(cm, n, "    Dim nm As String: nm = SelectedName()")
    n = IL(cm, n, "    If Len(nm) = 0 Then Exit Sub")
    n = IL(cm, n, "    KernelSnapshot.DeleteSnapshot nm")
    n = IL(cm, n, "    RefreshList")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub btnRename_Click()")
    n = IL(cm, n, "    Dim nm As String: nm = SelectedName()")
    n = IL(cm, n, "    If Len(nm) = 0 Then Exit Sub")
    n = IL(cm, n, "    Dim newName As String")
    n = IL(cm, n, "    newName = InputBox(""Enter new name for '"" & nm & ""':"", ""RDK -- Rename Snapshot"", nm)")
    n = IL(cm, n, "    If Len(Trim(newName)) = 0 Then Exit Sub")
    n = IL(cm, n, "    If newName = nm Then Exit Sub")
    n = IL(cm, n, "    KernelSnapshot.RenameSnapshot nm, newName")
    n = IL(cm, n, "    RefreshList")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub btnEditDesc_Click()")
    n = IL(cm, n, "    Dim nm As String: nm = SelectedName()")
    n = IL(cm, n, "    If Len(nm) = 0 Then Exit Sub")
    n = IL(cm, n, "    Dim curDesc As String")
    n = IL(cm, n, "    Dim mp As String")
    n = IL(cm, n, "    mp = KernelFormHelpers.GetProjectRootPublic() & ""\snapshots\"" & nm & ""\manifest.json""")
    n = IL(cm, n, "    curDesc = KernelFormHelpers.ReadJsonField(mp, ""description"")")
    n = IL(cm, n, "    Dim newDesc As String")
    n = IL(cm, n, "    newDesc = InputBox(""Edit description:"", ""RDK -- Edit Description"", curDesc)")
    n = IL(cm, n, "    If StrPtr(newDesc) = 0 Then Exit Sub")
    n = IL(cm, n, "    KernelFormHelpers.EditSnapshotDescription nm, newDesc")
    n = IL(cm, n, "    RefreshList")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub btnArchive_Click()")
    n = IL(cm, n, "    Dim nm As String: nm = SelectedName()")
    n = IL(cm, n, "    If Len(nm) = 0 Then Exit Sub")
    n = IL(cm, n, "    If Me.optActive.Value Then")
    n = IL(cm, n, "        KernelSnapshot.ArchiveSnapshot nm")
    n = IL(cm, n, "    Else")
    n = IL(cm, n, "        KernelSnapshot.RestoreFromArchive nm")
    n = IL(cm, n, "    End If")
    n = IL(cm, n, "    RefreshList")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub btnArchiveAll_Click()")
    n = IL(cm, n, "    If Not Me.optActive.Value Then Exit Sub")
    n = IL(cm, n, "    Dim names() As String")
    n = IL(cm, n, "    names = KernelSnapshot.ListSnapshots()")
    n = IL(cm, n, "    Dim cnt As Long: cnt = 0")
    n = IL(cm, n, "    Dim i As Long")
    n = IL(cm, n, "    For i = LBound(names) To UBound(names)")
    n = IL(cm, n, "        If Len(names(i)) > 0 Then cnt = cnt + 1")
    n = IL(cm, n, "    Next i")
    n = IL(cm, n, "    If cnt = 0 Then")
    n = IL(cm, n, "        MsgBox ""No snapshots to archive."", vbInformation, ""RDK""")
    n = IL(cm, n, "        Exit Sub")
    n = IL(cm, n, "    End If")
    n = IL(cm, n, "    If MsgBox(""Archive all "" & cnt & "" snapshot(s)?"", vbYesNo Or vbQuestion, ""RDK"") = vbNo Then Exit Sub")
    n = IL(cm, n, "    Dim ok As Long: ok = 0")
    n = IL(cm, n, "    For i = LBound(names) To UBound(names)")
    n = IL(cm, n, "        If Len(names(i)) > 0 Then")
    n = IL(cm, n, "            On Error Resume Next")
    n = IL(cm, n, "            KernelSnapshot.ArchiveSnapshot names(i)")
    n = IL(cm, n, "            If Err.Number = 0 Then ok = ok + 1")
    n = IL(cm, n, "            On Error GoTo 0")
    n = IL(cm, n, "        End If")
    n = IL(cm, n, "    Next i")
    n = IL(cm, n, "    MsgBox ""Archived "" & ok & "" of "" & cnt & "" snapshot(s)."", vbInformation, ""RDK""")
    n = IL(cm, n, "    RefreshList")
    n = IL(cm, n, "End Sub")
End Sub

' =============================================================================
' FORM 2: CompareExplorer
' Snapshot-based comparison: Base/Variant selection, current-state and
' inputs-only checkboxes, filter, sort, Compare, Clear Comparisons, Close
' =============================================================================
Private Sub CreateCompareExplorer()
    Dim comp As Object
    Set comp = ThisWorkbook.VBProject.VBComponents.Add(3)
    comp.Name = "CompareExplorer"
    Dim d As Object: Set d = comp.designer
    comp.Properties("Caption") = "RDK -- Compare"
    comp.Properties("Width") = 510
    comp.Properties("Height") = 420

    ' Filter
    AddLabel d, "lblFilter", "Filter:", 12, 8, 40, 18
    AddTextBox d, "txtFilter", 56, 6, 200, 20

    ' Sort buttons
    AddButton d, "btnSortName", "Sort: Name", 280, 6, 80, 22
    AddButton d, "btnSortDate", "Sort: Date", 368, 6, 80, 22

    ' List box
    AddListBox d, "lstItems", 12, 32, 480, 170, 4, "200;130;80;80"

    ' Selection labels
    AddLabel d, "lblBase", "Base: (none)", 12, 210, 230, 18
    AddLabel d, "lblVariant", "Variant: (none)", 250, 210, 230, 18

    ' Set Base / Set Variant
    AddButton d, "btnSetBase", "Set as Base", 12, 232, 110, 28
    AddButton d, "btnSetVariant", "Set as Variant", 132, 232, 110, 28

    ' Checkboxes
    AddCheckBox d, "chkCurrentState", "Compare current state vs Base", 12, 270, 230, 18
    AddCheckBox d, "chkInputsOnly", "Inputs only", 250, 270, 120, 18

    ' Action buttons
    AddButton d, "btnCompare", "Compare", 12, 302, 110, 28
    AddButton d, "btnClear", "Clear Comparisons", 132, 302, 130, 28
    AddButton d, "btnClose", "Close", 402, 302, 90, 28

    AddCompareExplorerCode comp.CodeModule
End Sub

Private Sub AddCompareExplorerCode(ByVal cm As Object)
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    Dim n As Long: n = 1
    n = IL(cm, n, "Option Explicit")
    n = IL(cm, n, "Private m_baseName As String")
    n = IL(cm, n, "Private m_variantName As String")
    n = IL(cm, n, "Private m_sortByDate As Boolean")
    n = IL(cm, n, "Private m_sortAsc As Boolean")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub UserForm_Initialize()")
    n = IL(cm, n, "    m_sortAsc = True")
    n = IL(cm, n, "    RefreshList")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub txtFilter_Change()")
    n = IL(cm, n, "    RefreshList")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub btnSortName_Click()")
    n = IL(cm, n, "    If m_sortByDate = False Then")
    n = IL(cm, n, "        m_sortAsc = Not m_sortAsc")
    n = IL(cm, n, "    Else")
    n = IL(cm, n, "        m_sortByDate = False")
    n = IL(cm, n, "        m_sortAsc = True")
    n = IL(cm, n, "    End If")
    n = IL(cm, n, "    RefreshList")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub btnSortDate_Click()")
    n = IL(cm, n, "    If m_sortByDate = True Then")
    n = IL(cm, n, "        m_sortAsc = Not m_sortAsc")
    n = IL(cm, n, "    Else")
    n = IL(cm, n, "        m_sortByDate = True")
    n = IL(cm, n, "        m_sortAsc = True")
    n = IL(cm, n, "    End If")
    n = IL(cm, n, "    RefreshList")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub RefreshList()")
    n = IL(cm, n, "    KernelFormHelpers.PopulateSnapshotListBox Me.lstItems, m_sortByDate, m_sortAsc")
    n = IL(cm, n, "    ApplyFilter")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub ApplyFilter()")
    n = IL(cm, n, "    Dim filterText As String")
    n = IL(cm, n, "    filterText = LCase(Trim(Me.txtFilter.Text))")
    n = IL(cm, n, "    If Len(filterText) = 0 Then Exit Sub")
    n = IL(cm, n, "    Dim i As Long")
    n = IL(cm, n, "    For i = Me.lstItems.ListCount - 1 To 0 Step -1")
    n = IL(cm, n, "        If InStr(1, LCase(Me.lstItems.List(i, 0)), filterText, vbTextCompare) = 0 Then")
    n = IL(cm, n, "            Me.lstItems.RemoveItem i")
    n = IL(cm, n, "        End If")
    n = IL(cm, n, "    Next i")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub btnSetBase_Click()")
    n = IL(cm, n, "    If Me.lstItems.ListIndex < 0 Then")
    n = IL(cm, n, "        MsgBox ""Select an item first."", vbExclamation, ""RDK""")
    n = IL(cm, n, "        Exit Sub")
    n = IL(cm, n, "    End If")
    n = IL(cm, n, "    m_baseName = Me.lstItems.List(Me.lstItems.ListIndex, 0)")
    n = IL(cm, n, "    Me.lblBase.Caption = ""Base: "" & m_baseName")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub btnSetVariant_Click()")
    n = IL(cm, n, "    If Me.lstItems.ListIndex < 0 Then")
    n = IL(cm, n, "        MsgBox ""Select an item first."", vbExclamation, ""RDK""")
    n = IL(cm, n, "        Exit Sub")
    n = IL(cm, n, "    End If")
    n = IL(cm, n, "    m_variantName = Me.lstItems.List(Me.lstItems.ListIndex, 0)")
    n = IL(cm, n, "    Me.lblVariant.Caption = ""Variant: "" & m_variantName")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub chkCurrentState_Click()")
    n = IL(cm, n, "    Me.btnSetVariant.Enabled = Not Me.chkCurrentState.Value")
    n = IL(cm, n, "    If Me.chkCurrentState.Value Then")
    n = IL(cm, n, "        Me.lblVariant.Caption = ""Variant: (current state)""")
    n = IL(cm, n, "        m_variantName = """"")
    n = IL(cm, n, "    Else")
    n = IL(cm, n, "        Me.lblVariant.Caption = ""Variant: (none)""")
    n = IL(cm, n, "    End If")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub btnCompare_Click()")
    n = IL(cm, n, "    If Me.chkCurrentState.Value Then")
    n = IL(cm, n, "        If Len(m_baseName) = 0 Then")
    n = IL(cm, n, "            MsgBox ""Set a Base first."", vbExclamation, ""RDK""")
    n = IL(cm, n, "            Exit Sub")
    n = IL(cm, n, "        End If")
    n = IL(cm, n, "        KernelCompare.CompareCurrentToSnapshot m_baseName")
    n = IL(cm, n, "    Else")
    n = IL(cm, n, "        If Len(m_baseName) = 0 Or Len(m_variantName) = 0 Then")
    n = IL(cm, n, "            MsgBox ""Set both Base and Variant."", vbExclamation, ""RDK""")
    n = IL(cm, n, "            Exit Sub")
    n = IL(cm, n, "        End If")
    n = IL(cm, n, "        If Me.chkInputsOnly.Value Then")
    n = IL(cm, n, "            KernelCompare.CompareInputsOnly m_baseName, m_variantName")
    n = IL(cm, n, "        Else")
    n = IL(cm, n, "            KernelCompare.CompareSnapshots m_baseName, m_variantName")
    n = IL(cm, n, "        End If")
    n = IL(cm, n, "    End If")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub btnClear_Click()")
    n = IL(cm, n, "    KernelCompare.RemoveComparisonTabs")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub btnClose_Click()")
    n = IL(cm, n, "    Unload Me")
    n = IL(cm, n, "End Sub")
End Sub


' =============================================================================
' CreateReportExplorer
' Builds a UserForm for selecting and exporting reports from report_templates.
' =============================================================================
Private Sub CreateReportExplorer()
    Dim comp As Object
    Set comp = ThisWorkbook.VBProject.VBComponents.Add(3)
    comp.Name = "ReportExplorer"
    Dim d As Object: Set d = comp.designer
    comp.Properties("Caption") = "RDK -- Export Report"
    comp.Properties("Width") = 480
    comp.Properties("Height") = 340

    ' Report template list
    AddLabel d, "lblTemplates", "Select a report template:", 12, 10, 200, 16
    AddListBox d, "lstTemplates", 12, 28, 450, 140, 3, "180;200;60"

    ' Format options
    AddLabel d, "lblFormat", "Format:", 12, 178, 50, 16
    AddOptionButton d, "optPDF", "PDF", 70, 178, 50, 16
    AddOptionButton d, "optPrint", "Print", 130, 178, 50, 16
    AddOptionButton d, "optPreview", "Preview", 190, 178, 65, 16

    ' Output info
    AddLabel d, "lblOutput", "", 12, 210, 450, 16
    Dim oCtrl As Object
    Set oCtrl = d.Controls("lblOutput")
    oCtrl.Font.Size = 8
    oCtrl.ForeColor = &H808080

    ' Buttons
    AddButton d, "btnExport", "Export", 12, 260, 100, 32
    AddButton d, "btnClose", "Close", 390, 260, 80, 28

    AddReportExplorerCode comp.CodeModule
End Sub


Private Sub AddReportExplorerCode(ByVal cm As Object)
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    Dim n As Long: n = 1
    n = IL(cm, n, "Option Explicit")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub UserForm_Initialize()")
    n = IL(cm, n, "    optPDF.Value = True")
    n = IL(cm, n, "    LoadTemplates")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub LoadTemplates()")
    n = IL(cm, n, "    lstTemplates.Clear")
    n = IL(cm, n, "    On Error Resume Next")
    n = IL(cm, n, "    Dim ws As Worksheet")
    n = IL(cm, n, "    Set ws = ThisWorkbook.Sheets(TAB_CONFIG)")
    n = IL(cm, n, "    If ws Is Nothing Then Exit Sub")
    n = IL(cm, n, "    Dim sr As Long")
    n = IL(cm, n, "    sr = KernelConfigLoader.FindSectionStart(ws, CFG_MARKER_REPORT_TEMPLATES)")
    n = IL(cm, n, "    If sr = 0 Then Exit Sub")
    n = IL(cm, n, "    Dim dr As Long: dr = sr + 2")
    n = IL(cm, n, "    Do While Len(Trim(CStr(ws.Cells(dr, RPTCFG_COL_ID).Value))) > 0")
    n = IL(cm, n, "        lstTemplates.AddItem Trim(CStr(ws.Cells(dr, RPTCFG_COL_NAME).Value))")
    n = IL(cm, n, "        lstTemplates.List(lstTemplates.ListCount - 1, 1) = Trim(CStr(ws.Cells(dr, RPTCFG_COL_DESC).Value))")
    n = IL(cm, n, "        lstTemplates.List(lstTemplates.ListCount - 1, 2) = Trim(CStr(ws.Cells(dr, RPTCFG_COL_FORMAT).Value))")
    n = IL(cm, n, "        dr = dr + 1")
    n = IL(cm, n, "    Loop")
    n = IL(cm, n, "    If lstTemplates.ListCount > 0 Then lstTemplates.ListIndex = 0")
    n = IL(cm, n, "    On Error GoTo 0")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub btnExport_Click()")
    n = IL(cm, n, "    If lstTemplates.ListIndex < 0 Then")
    n = IL(cm, n, "        MsgBox ""Select a report template."", vbExclamation, ""RDK""")
    n = IL(cm, n, "        Exit Sub")
    n = IL(cm, n, "    End If")
    n = IL(cm, n, "    Dim reportName As String")
    n = IL(cm, n, "    reportName = lstTemplates.List(lstTemplates.ListIndex, 0)")
    n = IL(cm, n, "    On Error Resume Next")
    n = IL(cm, n, "    Dim ws As Worksheet")
    n = IL(cm, n, "    Set ws = ThisWorkbook.Sheets(TAB_CONFIG)")
    n = IL(cm, n, "    Dim sr As Long")
    n = IL(cm, n, "    sr = KernelConfigLoader.FindSectionStart(ws, CFG_MARKER_REPORT_TEMPLATES)")
    n = IL(cm, n, "    If sr = 0 Then Exit Sub")
    n = IL(cm, n, "    Dim dr As Long: dr = sr + 2")
    n = IL(cm, n, "    Dim tabList As String: tabList = """"")
    n = IL(cm, n, "    Do While Len(Trim(CStr(ws.Cells(dr, RPTCFG_COL_ID).Value))) > 0")
    n = IL(cm, n, "        If StrComp(Trim(CStr(ws.Cells(dr, RPTCFG_COL_NAME).Value)), reportName, vbTextCompare) = 0 Then")
    n = IL(cm, n, "            tabList = Trim(CStr(ws.Cells(dr, RPTCFG_COL_TABS).Value))")
    n = IL(cm, n, "            Exit Do")
    n = IL(cm, n, "        End If")
    n = IL(cm, n, "        dr = dr + 1")
    n = IL(cm, n, "    Loop")
    n = IL(cm, n, "    On Error GoTo 0")
    n = IL(cm, n, "    If Len(tabList) = 0 Then Exit Sub")
    n = IL(cm, n, "    Dim outDir As String")
    n = IL(cm, n, "    outDir = KernelFormHelpers.EnsureOutputDir()")
    n = IL(cm, n, "    Dim safeName As String")
    n = IL(cm, n, "    safeName = Replace(reportName, "" "", ""_"")")
    n = IL(cm, n, "    Dim pdfPath As String")
    n = IL(cm, n, "    pdfPath = outDir & ""\"" & safeName & ""_"" & Format(Now, ""yyyymmdd_hhnnss"") & "".pdf""")
    n = IL(cm, n, "    lblOutput.Caption = pdfPath")
    n = IL(cm, n, "    If optPDF.Value Then")
    n = IL(cm, n, "        ExportPDF tabList, pdfPath")
    n = IL(cm, n, "    ElseIf optPreview.Value Then")
    n = IL(cm, n, "        MsgBox ""Preview: select tabs and use File > Print Preview."", vbInformation, ""RDK""")
    n = IL(cm, n, "    ElseIf optPrint.Value Then")
    n = IL(cm, n, "        MsgBox ""Print: select tabs and use File > Print."", vbInformation, ""RDK""")
    n = IL(cm, n, "    End If")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub ExportPDF(tabList As String, pdfPath As String)")
    n = IL(cm, n, "    On Error Resume Next")
    n = IL(cm, n, "    Dim tabs() As String")
    n = IL(cm, n, "    If tabList = ""*"" Then")
    n = IL(cm, n, "        Dim wk As Worksheet")
    n = IL(cm, n, "        Dim cnt As Long: cnt = 0")
    n = IL(cm, n, "        For Each wk In ThisWorkbook.Worksheets")
    n = IL(cm, n, "            If wk.Visible = xlSheetVisible Then cnt = cnt + 1")
    n = IL(cm, n, "        Next wk")
    n = IL(cm, n, "        ReDim tabs(0 To cnt - 1): cnt = 0")
    n = IL(cm, n, "        For Each wk In ThisWorkbook.Worksheets")
    n = IL(cm, n, "            If wk.Visible = xlSheetVisible Then tabs(cnt) = wk.Name: cnt = cnt + 1")
    n = IL(cm, n, "        Next wk")
    n = IL(cm, n, "    Else")
    n = IL(cm, n, "        tabs = Split(tabList, "","")")
    n = IL(cm, n, "    End If")
    n = IL(cm, n, "    Dim names() As String")
    n = IL(cm, n, "    ReDim names(0 To UBound(tabs))")
    n = IL(cm, n, "    Dim vc As Long: vc = 0")
    n = IL(cm, n, "    Dim i As Long")
    n = IL(cm, n, "    For i = 0 To UBound(tabs)")
    n = IL(cm, n, "        Dim tn As String: tn = Trim(tabs(i))")
    n = IL(cm, n, "        Dim wt As Worksheet: Set wt = Nothing")
    n = IL(cm, n, "        Set wt = ThisWorkbook.Sheets(tn)")
    n = IL(cm, n, "        If Not wt Is Nothing Then names(vc) = tn: vc = vc + 1")
    n = IL(cm, n, "    Next i")
    n = IL(cm, n, "    If vc = 0 Then Exit Sub")
    n = IL(cm, n, "    ReDim Preserve names(0 To vc - 1)")
    n = IL(cm, n, "    ThisWorkbook.Sheets(names).Select")
    n = IL(cm, n, "    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, _")
    n = IL(cm, n, "        Quality:=xlQualityStandard, IncludeDocProperties:=True, _")
    n = IL(cm, n, "        IgnorePrintAreas:=False, OpenAfterPublish:=True")
    n = IL(cm, n, "    ThisWorkbook.Sheets(TAB_DASHBOARD).Select")
    n = IL(cm, n, "    On Error GoTo 0")
    n = IL(cm, n, "    MsgBox ""Exported: "" & pdfPath, vbInformation, ""RDK""")
    n = IL(cm, n, "End Sub")
    n = IL(cm, n, "")
    n = IL(cm, n, "Private Sub btnClose_Click()")
    n = IL(cm, n, "    Unload Me")
    n = IL(cm, n, "End Sub")
End Sub


' =============================================================================

' WorkspaceExplorer moved to KernelFormSetup2.bas (AD-09 split)
