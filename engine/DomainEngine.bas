Attribute VB_Name = "DomainEngine"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' DomainEngine.bas -- Domain Engine Contract Stub
' =============================================================================
' To build a new model on the RDK:
'   1. Copy this file and rename it (e.g., MonteCarloDomainEngine.bas)
'   2. Implement the required functions below
'   3. Set DomainModule=YourModuleName in granularity_config.csv
'   4. Create your config CSVs (column_registry, input_schema, formula_tab_config, etc.)
'   5. Run Setup.bat, bootstrap, and Run Model
'
' REQUIRED FUNCTIONS (kernel calls these via Application.Run):
'   Initialize()              Called once after config is loaded
'   Validate() As Boolean     Called before computation; return False to halt
'   Reset()                   Called before Execute to clear state
'   Execute()                 Fill DomainOutputs array (Incremental + Dimension cols)
'
' OPTIONAL FUNCTIONS (kernel calls via Application.Run with error trapping):
'   GetRowCount() As Long     Total Detail rows (default: entities x periods)
'   GetMaxPeriod() As Long    Max period for Summary columns
'   GetEntityCount() As Long  Entity count for dashboard metadata
'   RunDomainTests() As Boolean  Domain validation tests (called after run)
'   GranularCSVPath() As String  Path to granular CSV for snapshot capture
'
' RULES:
'   - Read inputs via KernelConfig.InputValue("Section", "Param", entityIdx)
'   - Reference columns via KernelConfig.ColIndex("MetricName") only (AP-08)
'   - Never compute or store cumulative values (AP-42, kernel handles)
'   - Never compute Derived fields (AP-42, kernel uses DerivationRule)
'   - Write to outputs via Public DomainOutputs array (BUG-034 workaround)
'   - Log errors via KernelConfig.LogError (AP-46: include bypass instructions)
'
' DATA FLOW:
'   KernelEngine allocates outputs() array -> copies to DomainOutputs ->
'   calls Execute() -> copies DomainOutputs back to outputs() ->
'   kernel computes Derived fields and writes to Detail tab.
' =============================================================================

Public Sub Initialize()
    ' Store any config references needed for Execute.
    ' Read from KernelConfig.InputValue, ColIndex, GetSetting, etc.
End Sub

Public Function Validate() As Boolean
    Validate = True
    ' Add domain-specific input validation here.
    ' Return False to halt RunProjections.
    ' Log failures via KernelConfig.LogError SEV_ERROR, ...
End Function

Public Sub Reset()
    ' Erase all module-level computation state.
    ' Called before each computation run.
End Sub

Public Sub Execute()
    ' Fill the DomainOutputs array with Dimension and Incremental values.
    '
    ' Example:
    '   Dim row As Long
    '   Dim colRevenue As Long
    '   colRevenue = KernelConfig.ColIndex("Revenue")
    '   DomainOutputs(row, colRevenue) = someCalculation
    '
    ' Use KernelConfig.InputValue("Section", "Param", entityIdx) for inputs.
    ' Use KernelConfig.ColIndex("MetricName") for column indices.
    ' Write ONLY Dimension and Incremental fields.
End Sub
