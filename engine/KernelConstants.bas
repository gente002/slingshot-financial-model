Attribute VB_Name = "KernelConstants"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelConstants.bas
' Purpose: Kernel-only constants. Version string. Error severity constants.
' =============================================================================

Public Const KERNEL_VERSION As String = "1.3.1"
Public Const CONFIG_SCHEMA_VERSION As String = "1"

' Build fingerprint (authorship proof -- do not remove)
Public Const KERNEL_AUTHOR As String = "Ethan Genteman"
Public Const KERNEL_BUILD_ID As String = "0398a2b3234eec80"
Public Const KERNEL_BUILD_DATE As String = "2026-04-03"

' Internal validation seed (location 4 of 4)
Private Const CFG_VALIDATION_SEED As String = "39ee874dd4a5df1a6547dbaa06ad94ce"

' Error severity levels
Public Const SEV_FATAL As Long = 0
Public Const SEV_ERROR As Long = 1
Public Const SEV_WARN As Long = 2
Public Const SEV_INFO As Long = 3

' Kernel tab names (must match tab_registry.csv)
Public Const TAB_CONFIG As String = "Config"
Public Const TAB_INPUTS As String = "Inputs"
Public Const TAB_DASHBOARD As String = "Dashboard"
Public Const TAB_DETAIL As String = "Detail"
Public Const TAB_SUMMARY As String = "Summary"  ' DEPRECATED -- tab removed from registry
Public Const TAB_ERROR_LOG As String = "Error Log"
Public Const TAB_TEST_RESULTS As String = "Test Results"

' Detail tab layout
Public Const DETAIL_HEADER_ROW As Long = 1
Public Const DETAIL_DATA_START_ROW As Long = 2

' Input tab layout
Public Const INPUT_ENTITY_START_COL As Long = 3  ' Column C = first entity
Public Const INPUT_MAX_ENTITIES As Long = 10

' Config sheet section markers
Public Const CFG_MARKER_COLUMN_REGISTRY As String = "=== COLUMN_REGISTRY ==="
Public Const CFG_MARKER_INPUT_SCHEMA As String = "=== INPUT_SCHEMA ==="
Public Const CFG_MARKER_GRANULARITY_CONFIG As String = "=== GRANULARITY_CONFIG ==="
Public Const CFG_MARKER_TAB_REGISTRY As String = "=== TAB_REGISTRY ==="

' Config sheet: ColumnRegistry CSV column positions (1-indexed on Config sheet)
Public Const CREG_COL_NAME As Long = 1
Public Const CREG_COL_DETAIL As Long = 2
Public Const CREG_COL_CSV As Long = 3
Public Const CREG_COL_BLOCK As Long = 4
Public Const CREG_COL_FIELDCLASS As Long = 5
Public Const CREG_COL_DEFAULTVIEW As Long = 6
Public Const CREG_COL_FORMAT As Long = 7
Public Const CREG_COL_BALGRP As Long = 8
Public Const CREG_COL_DERIVRULE As Long = 9

' Config sheet: InputSchema CSV column positions (1-indexed on Config sheet)
Public Const ISCH_COL_SECTION As Long = 1
Public Const ISCH_COL_PARAM As Long = 2
Public Const ISCH_COL_ROW As Long = 3
Public Const ISCH_COL_TYPE As Long = 4
Public Const ISCH_COL_DEFAULT As Long = 5

' Config sheet: GranularityConfig CSV column positions
Public Const GCFG_COL_KEY As Long = 1
Public Const GCFG_COL_VALUE As Long = 2

' Config sheet: TabRegistry CSV column positions
Public Const TREG_COL_TABNAME As Long = 1
Public Const TREG_COL_PROTECTED As Long = 4
Public Const TREG_COL_VISIBLE As Long = 5
Public Const TREG_COL_SORTORDER As Long = 6
Public Const TREG_COL_TABCOLOR As Long = 11

' Summary tab layout
Public Const SUMMARY_COL_ENTITY As Long = 1
Public Const SUMMARY_COL_METRIC As Long = 2
Public Const SUMMARY_DATA_START_COL As Long = 3

' Pipeline step numbers (MBP)
Public Const STEP_BOOTSTRAP As Long = 0
Public Const STEP_LOAD_CONFIG As Long = 1
Public Const STEP_VALIDATE As Long = 2
Public Const STEP_COMPUTE As Long = 3
Public Const STEP_WRITE_DETAIL As Long = 4
Public Const STEP_WRITE_CSV As Long = 5
Public Const STEP_WRITE_SUMMARY As Long = 6

' Pipeline status values (MBP)
Public Const STEP_STATUS_COMPLETE As String = "COMPLETE"
Public Const STEP_STATUS_FAILED As String = "FAILED"
Public Const STEP_STATUS_PENDING As String = "PENDING"
Public Const STEP_STATUS_SKIPPED As String = "SKIPPED"
Public Const STEP_STATUS_BYPASSED As String = "BYPASSED"

' Pipeline state location on Config sheet
Public Const PIPELINE_STATE_MARKER As String = "=== PIPELINE_STATE ==="

' Config sheet section markers (Phase 2)
Public Const CFG_MARKER_REPRO_CONFIG As String = "=== REPRO_CONFIG ==="
Public Const CFG_MARKER_SCALE_LIMITS As String = "=== SCALE_LIMITS ==="

' Config sheet: ReproConfig CSV column positions (1-indexed on Config sheet)
Public Const RCFG_COL_KEY As Long = 1
Public Const RCFG_COL_VALUE As Long = 2

' Config sheet: ScaleLimits CSV column positions (1-indexed on Config sheet)
Public Const SCFG_COL_KEY As Long = 1
Public Const SCFG_COL_VALUE As Long = 2

' Phase 2: Directory names (relative to project root)
Public Const DIR_SNAPSHOTS As String = "snapshots"
Public Const DIR_ARCHIVE As String = "archive"
Public Const DIR_WAL As String = "wal"

' Phase 11A: Output directory
Public Const DIR_OUTPUT As String = "output"

' Phase 2: Savepoint status
Public Const SP_STATUS_COMPLETE As String = "COMPLETE"

' Phase 2: Snapshot name constraints
Public Const SNAPSHOT_NAME_MAX_LEN As Long = 50

' Phase 2: Lock timeout (seconds)
Public Const LOCK_TIMEOUT_SECONDS As Long = 60

' Phase 2: Run state section on Config sheet
Public Const CFG_MARKER_RUN_STATE As String = "=== RUN_STATE ==="
Public Const RS_KEY_TIMESTAMP As String = "LastRunTimestamp"
Public Const RS_KEY_TOTAL_ELAPSED As String = "TotalElapsedSec"
Public Const RS_KEY_INPUT_HASH As String = "LastRunInputHash"
Public Const RS_KEY_CONFIG_HASH As String = "LastRunConfigHash"
Public Const RS_KEY_STALE As String = "ResultsStale"
Public Const SCALE_LARGE_MODEL_SEC As String = "LargeModelThresholdSec"

' Phase 2: Comparison defaults (AD-40: KernelCompare)
Public Const COMPARE_DEFAULT_THRESHOLD As Double = 0.000001
Public Const COMPARE_TAB_PREFIX As String = "Compare_"
Public Const COMPARE_INPUT_PREFIX As String = "InputComp_"
Public Const COMPARE_SUMMARY_PREFIX As String = "SummComp_"
Public Const COMPARE_COLS_PER_METRIC As Long = 4

' Phase 3: Test tiers
Public Const TEST_TIER_UNIT As Long = 1
Public Const TEST_TIER_EDGE As Long = 2
Public Const TEST_TIER_INTEGRATION As Long = 3
Public Const TEST_TIER_REGRESSION As Long = 4
Public Const TEST_TIER_EXHIBIT As Long = 5

' Phase 3: Test results
Public Const TEST_PASS As String = "PASS"
Public Const TEST_FAIL As String = "FAIL"

' Phase 3: Test result tolerance
Public Const TEST_DEFAULT_TOLERANCE As Double = 0.000001

' Phase 3: Prove-It check types
Public Const PROVEIT_IDENTITY As String = "Identity"
Public Const PROVEIT_ACCUMULATE As String = "Accumulate"
Public Const PROVEIT_RECONCILE As String = "Reconcile"

' Phase 3: Tabs
Public Const TAB_PROVE_IT As String = "Prove It"

' Phase 3: Golden prefix
Public Const GOLDEN_PREFIX As String = "GOLDEN_"

' Phase 3: Config sheet section marker for prove_it_config
Public Const CFG_MARKER_PROVE_IT_CONFIG As String = "=== PROVE_IT_CONFIG ==="

' Phase 3: ProveItConfig CSV column positions (1-indexed on Config sheet)
Public Const PCFG_COL_CHECKID As Long = 1
Public Const PCFG_COL_CHECKTYPE As Long = 2
Public Const PCFG_COL_CHECKNAME As Long = 3
Public Const PCFG_COL_METRICA As Long = 4
Public Const PCFG_COL_METRICB As Long = 5
Public Const PCFG_COL_METRICC As Long = 6
Public Const PCFG_COL_OPERATOR As Long = 7
Public Const PCFG_COL_TOLERANCE As Long = 8
Public Const PCFG_COL_ENABLED As Long = 9

' Phase 3: TestResults sheet layout
Public Const TR_HEADER_ROW As Long = 4
Public Const TR_DATA_START_ROW As Long = 5
Public Const TR_COL_TIER As Long = 1
Public Const TR_COL_TESTID As Long = 2
Public Const TR_COL_TESTNAME As Long = 3
Public Const TR_COL_EXPECTED As Long = 4
Public Const TR_COL_ACTUAL As Long = 5
Public Const TR_COL_RESULT As Long = 6
Public Const TR_COL_DETAIL As Long = 7

' Phase 3: Large model sampling threshold
Public Const LARGE_MODEL_ROW_THRESHOLD As Long = 1000
Public Const LARGE_MODEL_SAMPLE_SIZE As Long = 30

' Phase 4: Health check modes
Public Const HEALTH_LIGHTWEIGHT As String = "LIGHTWEIGHT"
Public Const HEALTH_FULL As String = "FULL"

' Phase 4: Lint check IDs
Public Const LINT_NON_ASCII As String = "LINT-01"
Public Const LINT_FORMULA_INJECT As String = "LINT-02"
Public Const LINT_MAGIC_NUMBER As String = "LINT-03"
Public Const LINT_REDIM_IN_LOOP As String = "LINT-04"
Public Const LINT_SIZED_DIM_AFTER_CODE As String = "LINT-05"
Public Const LINT_TWO_CHAR_VAR As String = "LINT-06"
Public Const LINT_CUMULATIVE_DOMAIN As String = "LINT-07"
Public Const LINT_MISSING_CONTRACT As String = "LINT-08"
Public Const LINT_EQ_NO_FORMAT As String = "LINT-09"
Public Const LINT_SUB_FUNC_BALANCE As String = "LINT-10"
Public Const LINT_MODULE_SIZE As String = "LINT-11"
Public Const LINT_CRLF As String = "LINT-12"

' Phase 4: Lint severity constants
Public Const LINT_SEV_ERROR As String = "ERROR"
Public Const LINT_SEV_WARN As String = "WARN"

' Phase 4: Module size thresholds (bytes)
Public Const MODULE_SIZE_WARN As Long = 50000
Public Const MODULE_SIZE_ERROR As Long = 65536

' Phase 4: Dashboard button names
Public Const BTN_RUN_LINT As String = "Run Lint"
Public Const BTN_HEALTH_CHECK As String = "Health Check"
Public Const BTN_DIAGNOSTIC_DUMP As String = "Diagnostic Dump"

' Phase 4: Diagnostic dump recursion guard
Public Const DIAG_DUMP_MAX_LOG_ENTRIES As Long = 20
Public Const DIAG_DUMP_MAX_WAL_ENTRIES As Long = 20

' Phase 4: Health check thresholds
Public Const HC_ERRORLOG_WARN_ROWS As Long = 1000
Public Const HC_WAL_WARN_LINES As Long = 10000

' Phase 5A: Config sheet section markers
Public Const CFG_MARKER_SUMMARY_CONFIG As String = "=== SUMMARY_CONFIG ==="
Public Const CFG_MARKER_CHART_REGISTRY As String = "=== CHART_REGISTRY ==="
Public Const CFG_MARKER_EXHIBIT_CONFIG As String = "=== EXHIBIT_CONFIG ==="
Public Const CFG_MARKER_DISPLAY_MODE_CONFIG As String = "=== DISPLAY_MODE_CONFIG ==="

' Phase 5A: Summary config columns (1-indexed on Config sheet)
Public Const SUMCFG_COL_METRIC As Long = 1
Public Const SUMCFG_COL_SECTION As Long = 2
Public Const SUMCFG_COL_SORT As Long = 3
Public Const SUMCFG_COL_FORMAT As Long = 4
Public Const SUMCFG_COL_SHOW As Long = 5

' Phase 5A: Chart registry columns
Public Const CHTCFG_COL_ID As Long = 1
Public Const CHTCFG_COL_NAME As Long = 2
Public Const CHTCFG_COL_TYPE As Long = 3
Public Const CHTCFG_COL_METRIC As Long = 4
Public Const CHTCFG_COL_GROUPBY As Long = 5
Public Const CHTCFG_COL_WIDTH As Long = 6
Public Const CHTCFG_COL_HEIGHT As Long = 7
Public Const CHTCFG_COL_ENABLED As Long = 8

' Phase 5A: Exhibit config columns
Public Const EXHCFG_COL_ID As Long = 1
Public Const EXHCFG_COL_NAME As Long = 2
Public Const EXHCFG_COL_METRICS As Long = 3
Public Const EXHCFG_COL_GROUPBY As Long = 4
Public Const EXHCFG_COL_TOTAL As Long = 5
Public Const EXHCFG_COL_FORMAT As Long = 6
Public Const EXHCFG_COL_ENABLED As Long = 7

' Phase 5A: Display mode config columns (key-value)
Public Const DMCFG_COL_KEY As Long = 1
Public Const DMCFG_COL_VALUE As Long = 2

' Phase 5A: Display mode constants
Public Const DISPLAY_INCREMENTAL As String = "Incremental"
Public Const DISPLAY_CUMULATIVE As String = "Cumulative"

' Phase 5A: Tab names
Public Const TAB_CUMULATIVE_VIEW As String = "CumulativeView"  ' DEPRECATED -- tab removed from registry
Public Const TAB_EXHIBITS As String = "Exhibits"  ' DEPRECATED -- tab removed from registry
Public Const TAB_CHARTS As String = "Charts"  ' DEPRECATED -- tab removed from registry

' Phase 5A: Chart type strings
Public Const CHART_LINE As String = "Line"
Public Const CHART_BAR As String = "Bar"
Public Const CHART_STACKED As String = "StackedBar"
Public Const CHART_PIE As String = "Pie"
Public Const CHART_AREA As String = "Area"

' Phase 5A: Dashboard button names
Public Const BTN_TOGGLE_DISPLAY As String = "Toggle Display Mode"
Public Const BTN_REFRESH_EXHIBITS As String = "Refresh Exhibits"

' Phase 5B: Config sheet section markers
Public Const CFG_MARKER_PRINT_CONFIG As String = "=== PRINT_CONFIG ==="
Public Const CFG_MARKER_DATA_MODEL_CONFIG As String = "=== DATA_MODEL_CONFIG ==="
Public Const CFG_MARKER_PIVOT_CONFIG As String = "=== PIVOT_CONFIG ==="

' Phase 5B: Print config columns (1-indexed on Config sheet)
Public Const PRTCFG_COL_TABNAME As Long = 1
Public Const PRTCFG_COL_ORIENT As Long = 2
Public Const PRTCFG_COL_FITPAGES As Long = 3
Public Const PRTCFG_COL_PAPER As Long = 4
Public Const PRTCFG_COL_PRINTAREA As Long = 5
Public Const PRTCFG_COL_HDRLEFT As Long = 6
Public Const PRTCFG_COL_HDRCENTER As Long = 7
Public Const PRTCFG_COL_HDRRIGHT As Long = 8
Public Const PRTCFG_COL_FTRCENTER As Long = 9
Public Const PRTCFG_COL_INCLUDEPDF As Long = 10
Public Const PRTCFG_COL_PRINTORDER As Long = 11
Public Const PRTCFG_COL_FITPAGESTALL As Long = 12
Public Const PRTCFG_COL_MARGINS As Long = 13
Public Const PRTCFG_COL_CENTERH As Long = 14

' Phase 5B: Pivot config columns (1-indexed on Config sheet)
Public Const PVTCFG_COL_ID As Long = 1
Public Const PVTCFG_COL_NAME As Long = 2
Public Const PVTCFG_COL_SOURCE As Long = 3
Public Const PVTCFG_COL_ROWFIELD As Long = 4
Public Const PVTCFG_COL_COLFIELD As Long = 5
Public Const PVTCFG_COL_VALUEFIELD As Long = 6
Public Const PVTCFG_COL_AGGFUNC As Long = 7
Public Const PVTCFG_COL_ENABLED As Long = 8

' Phase 5B: Tab names
Public Const TAB_ANALYSIS As String = "Analysis"  ' DEPRECATED -- tab removed from registry

' Phase 5B: Dashboard button names
Public Const BTN_PRINT_PREVIEW As String = "Print Preview"
Public Const BTN_EXPORT_PDF As String = "Export PDF"

' Phase 5B: Pipeline step for transforms
Public Const STEP_TRANSFORMS As Long = 35

' Phase 5C: Config sheet section markers
Public Const CFG_MARKER_FORMULA_TAB_CONFIG As String = "=== FORMULA_TAB_CONFIG ==="
Public Const CFG_MARKER_NAMED_RANGE_REGISTRY As String = "=== NAMED_RANGE_REGISTRY ==="

' Phase 5C: Formula tab config columns (1-indexed on Config sheet)
Public Const FTCFG_COL_TABNAME As Long = 1
Public Const FTCFG_COL_ROWID As Long = 2
Public Const FTCFG_COL_ROW As Long = 3
Public Const FTCFG_COL_COL As Long = 4
Public Const FTCFG_COL_CELLTYPE As Long = 5
Public Const FTCFG_COL_CONTENT As Long = 6
Public Const FTCFG_COL_FORMAT As Long = 7
Public Const FTCFG_COL_FONTSTYLE As Long = 8
Public Const FTCFG_COL_FILLCOLOR As Long = 9
Public Const FTCFG_COL_FONTCOLOR As Long = 10
Public Const FTCFG_COL_COLSPAN As Long = 11
Public Const FTCFG_COL_BORDERBOTTOM As Long = 12
Public Const FTCFG_COL_BORDERTOP As Long = 13
Public Const FTCFG_COL_INDENT As Long = 14
Public Const FTCFG_COL_COMMENT As Long = 15
Public Const FTCFG_COL_HALIGN As Long = 16

' Phase 5C: Named range registry columns
Public Const NRCFG_COL_NAME As Long = 1
Public Const NRCFG_COL_TABNAME As Long = 2
Public Const NRCFG_COL_ROWID As Long = 3
Public Const NRCFG_COL_CELLADDR As Long = 4
Public Const NRCFG_COL_RANGETYPE As Long = 5
Public Const NRCFG_COL_DESC As Long = 6

' Phase 5C: Tab names
Public Const TAB_QUARTERLY_SUMMARY As String = "Quarterly Summary"
Public Const TAB_FINANCIAL_SUMMARY As String = "FinancialSummary"

' Phase 5C: BalanceType values
Public Const BALANCE_TYPE_FLOW As String = "Flow"
Public Const BALANCE_TYPE_BALANCE As String = "Balance"

' Phase 5C: Column registry BalanceType column
Public Const CREG_COL_BALANCETYPE As Long = 10

' Phase 5C: Tab registry QuarterlyColumns column
Public Const TREG_COL_QUARTERLY As Long = 8
' Phase 11A: Tab registry QuarterlyHorizon column (Data or Writing)
Public Const TREG_COL_QTR_HORIZON As Long = 9
' Phase 11B: Tab registry GrandTotal column (Y1-Yn sum column after last year)
Public Const TREG_COL_GRANDTOTAL As Long = 10
' Domain-separation: tab_registry SkipAnnual and HasTailColumn columns
Public Const TREG_COL_SKIPANNUAL As Long = 12
Public Const TREG_COL_HASTAILCOL As Long = 13
' Coverage contract column (session 2026-04-18 decision O): declares
' how a tab's user-editable cells are preserved across workspace save/load.
' Values: TAB_IS_INPUT, PRESERVED_CELLS, EPHEMERAL, NONE
Public Const TREG_COL_INPUTSURFACE As Long = 14

' Phase 5C: Dashboard button name
Public Const BTN_REFRESH_FORMULAS As String = "Refresh Formula Tabs"

' Phase 5C: Formula tab fixed column positions
Public Const FTAB_COL_ROWID As Long = 1
Public Const FTAB_COL_LABEL As Long = 2

' Phase 5C: Quarterly layout constants
Public Const QS_DATA_START_COL As Long = 3
Public Const QS_QUARTERS_PER_YEAR As Long = 4
Public Const QS_COLS_PER_YEAR As Long = 5

' Phase 6A: Config sheet section markers
Public Const CFG_MARKER_EXTENSION_REGISTRY As String = "=== EXTENSION_REGISTRY ==="
Public Const CFG_MARKER_CURVE_LIBRARY As String = "=== CURVE_LIBRARY_CONFIG ==="
Public Const CFG_MARKER_REPORT_CONFIG As String = "=== REPORT_CONFIG ==="

' Domain-separation config section markers
Public Const CFG_MARKER_VALIDATION_CONFIG As String = "=== VALIDATION_CONFIG ==="
Public Const CFG_MARKER_HEALTH_CONFIG As String = "=== HEALTH_CONFIG ==="
Public Const CFG_MARKER_BRANDING_CONFIG As String = "=== BRANDING_CONFIG ==="

' Validation config columns (1-indexed)
Public Const VALCFG_COL_TABNAME As Long = 1
Public Const VALCFG_COL_PATTERN As Long = 2
Public Const VALCFG_COL_COLSTART As Long = 3
Public Const VALCFG_COL_COLEND As Long = 4
Public Const VALCFG_COL_VALTYPE As Long = 5
Public Const VALCFG_COL_OPERATOR As Long = 6
Public Const VALCFG_COL_MIN As Long = 7
Public Const VALCFG_COL_MAX As Long = 8
Public Const VALCFG_COL_ALERTSTYLE As Long = 9
Public Const VALCFG_COL_ERRMSG As Long = 10

' Health config columns (1-indexed)
Public Const HLCFG_COL_TABNAME As Long = 1
Public Const HLCFG_COL_ROWID As Long = 2
Public Const HLCFG_COL_COLSTART As Long = 3
Public Const HLCFG_COL_COLEND As Long = 4
Public Const HLCFG_COL_CHECKTYPE As Long = 5
Public Const HLCFG_COL_GOODVALUE As Long = 6
Public Const HLCFG_COL_THRESHOLD As Long = 7

' Branding config columns (key-value)
Public Const BRCFG_COL_KEY As Long = 1
Public Const BRCFG_COL_VALUE As Long = 2

' Pipeline config section marker and columns
Public Const CFG_MARKER_PIPELINE_CONFIG As String = "=== PIPELINE_CONFIG ==="
Public Const PLCFG_COL_STEPID As Long = 1
Public Const PLCFG_COL_STEPORDER As Long = 2
Public Const PLCFG_COL_ENABLED As Long = 3
Public Const PLCFG_COL_DESC As Long = 4

' Config schema section marker
Public Const CFG_MARKER_CONFIG_SCHEMA As String = "=== CONFIG_SCHEMA ==="

' MsgBox config section marker and columns
Public Const CFG_MARKER_MSGBOX_CONFIG As String = "=== MSGBOX_CONFIG ==="
Public Const MBCFG_COL_ID As Long = 1
Public Const MBCFG_COL_TITLE As Long = 2
Public Const MBCFG_COL_MESSAGE As Long = 3
Public Const MBCFG_COL_ICON As Long = 4
Public Const MBCFG_COL_BUTTONS As Long = 5

' Display aliases section marker and columns
Public Const CFG_MARKER_DISPLAY_ALIASES As String = "=== DISPLAY_ALIASES ==="
Public Const DACFG_COL_ID As Long = 1
Public Const DACFG_COL_DISPLAY As Long = 2
Public Const DACFG_COL_CATEGORY As Long = 3
Public Const DACFG_COL_DESC As Long = 4

' Formula tab config BalanceItem column
Public Const FTCFG_COL_BALANCEITEM As Long = 17

' Button config section marker and columns (any tab can have buttons)
Public Const CFG_MARKER_BUTTON_CONFIG As String = "=== BUTTON_CONFIG ==="
Public Const BTNCFG_COL_TABNAME As Long = 1
Public Const BTNCFG_COL_ID As Long = 2
Public Const BTNCFG_COL_CAPTION As Long = 3
Public Const BTNCFG_COL_MACRO As Long = 4
Public Const BTNCFG_COL_DEVONLY As Long = 5
Public Const BTNCFG_COL_SORT As Long = 6
Public Const BTNCFG_COL_ENABLED As Long = 7
Public Const BTNCFG_COL_ROW As Long = 8
Public Const BTNCFG_COL_COL As Long = 9

' Report templates section marker and columns
Public Const CFG_MARKER_REPORT_TEMPLATES As String = "=== REPORT_TEMPLATES ==="
Public Const RPTCFG_COL_ID As Long = 1
Public Const RPTCFG_COL_NAME As Long = 2
Public Const RPTCFG_COL_DESC As Long = 3
Public Const RPTCFG_COL_TABS As Long = 4
Public Const RPTCFG_COL_FORMAT As Long = 5
Public Const RPTCFG_COL_ORIENT As Long = 6
Public Const RPTCFG_COL_COVER As Long = 7

' Diagnostic config section marker and columns
Public Const CFG_MARKER_DIAGNOSTIC_CONFIG As String = "=== DIAGNOSTIC_CONFIG ==="
Public Const DIAGCFG_COL_ID As Long = 1
Public Const DIAGCFG_COL_NAME As Long = 2
Public Const DIAGCFG_COL_MODULE As Long = 3
Public Const DIAGCFG_COL_ENTRY As Long = 4
Public Const DIAGCFG_COL_ENABLED As Long = 5
Public Const DIAGCFG_COL_DESC As Long = 6

' Regression config section marker and columns
Public Const CFG_MARKER_REGRESSION_CONFIG As String = "=== REGRESSION_CONFIG ==="
Public Const REGCFG_COL_TABNAME As Long = 1
Public Const REGCFG_COL_TYPE As Long = 2
Public Const REGCFG_COL_TOLERANCE As Long = 3
Public Const REGCFG_COL_DESC As Long = 4

' Config version section marker
Public Const CFG_MARKER_CONFIG_VERSION As String = "=== CONFIG_VERSION ==="

' Assumptions config section marker
Public Const CFG_MARKER_ASSUMPTIONS_CONFIG As String = "=== ASSUMPTIONS_CONFIG ==="

' Lock config section marker
Public Const CFG_MARKER_LOCK_CONFIG As String = "=== LOCK_CONFIG ==="

' Workspace config section marker
Public Const CFG_MARKER_WORKSPACE_CONFIG As String = "=== WORKSPACE_CONFIG ==="
Public Const WSCFG_COL_KEY As Long = 1
Public Const WSCFG_COL_VALUE As Long = 2

' Phase 6A: Extension registry columns (1-indexed on Config sheet)
Public Const EXTCFG_COL_ID As Long = 1
Public Const EXTCFG_COL_MODULE As Long = 2
Public Const EXTCFG_COL_ENTRY As Long = 3
Public Const EXTCFG_COL_HOOK As Long = 4
Public Const EXTCFG_COL_SORT As Long = 5
Public Const EXTCFG_COL_ACTIVE As Long = 6
Public Const EXTCFG_COL_MUTATES As Long = 7
Public Const EXTCFG_COL_SEED As Long = 8
Public Const EXTCFG_COL_DESC As Long = 9

' Phase 6A: Curve library config columns (1-indexed on Config sheet)
Public Const CLCFG_COL_ID As Long = 1
Public Const CLCFG_COL_CATEGORY As Long = 2
Public Const CLCFG_COL_TYPE As Long = 3
Public Const CLCFG_COL_DIST As Long = 4
Public Const CLCFG_COL_P1 As Long = 5
Public Const CLCFG_COL_P2 As Long = 6
Public Const CLCFG_COL_P3 As Long = 7
Public Const CLCFG_COL_MAXAGE As Long = 8
Public Const CLCFG_COL_INTERP As Long = 9
Public Const CLCFG_COL_DESC As Long = 10

' Anchor-based curve library config columns (5-point: TL=1,25,50,75,100)
Public Const CLA_COL_LOB As Long = 1
Public Const CLA_COL_TYPE As Long = 2
Public Const CLA_COL_DIST As Long = 3
Public Const CLA_COL_P1METHOD As Long = 4
Public Const CLA_COL_P1_TL1 As Long = 5
Public Const CLA_COL_P1_TL25 As Long = 6
Public Const CLA_COL_P1_TL50 As Long = 7
Public Const CLA_COL_P1_TL75 As Long = 8
Public Const CLA_COL_P1_TL100 As Long = 9
Public Const CLA_COL_P2_TL1 As Long = 10
Public Const CLA_COL_P2_TL25 As Long = 11
Public Const CLA_COL_P2_TL50 As Long = 12
Public Const CLA_COL_P2_TL75 As Long = 13
Public Const CLA_COL_P2_TL100 As Long = 14
Public Const CLA_COL_P3 As Long = 15
Public Const CLA_COL_MA_TL1 As Long = 16
Public Const CLA_COL_MA_TL25 As Long = 17
Public Const CLA_COL_MA_TL50 As Long = 18
Public Const CLA_COL_MA_TL75 As Long = 19
Public Const CLA_COL_MA_TL100 As Long = 20
Public Const CLA_COL_DESC As Long = 21

' Phase 6A: Dashboard button names
Public Const BTN_LIST_EXTENSIONS As String = "List Extensions"

' Dev Mode constants
Public Const DEV_MODE_ON As String = "ON"
Public Const DEV_MODE_OFF As String = "OFF"
Public Const CFG_MARKER_DEV_MODE As String = "=== DEV_MODE ==="
Public Const DEVCFG_COL_KEY As Long = 1
Public Const DEVCFG_COL_VALUE As Long = 2

' Kernel infrastructure tabs
Public Const TAB_RUN_METADATA As String = "Run Metadata"

' Phase 12: Investor demo tabs
Public Const TAB_COVER_PAGE As String = "Cover Page"
Public Const TAB_USER_GUIDE As String = "User Guide"

' Phase 11A: Configurable input tab alias
Public Const TAB_ASSUMPTIONS As String = "Assumptions"

' Dev mode tab list -- derived from tab_registry (Hidden tabs)
Public Function GetDevModeTabs() As Variant
    On Error Resume Next
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(TAB_CONFIG)
    If wsConfig Is Nothing Then
        GetDevModeTabs = Array()
        Exit Function
    End If
    On Error GoTo 0

    Dim sr As Long
    sr = KernelConfigLoader.FindSectionStart(wsConfig, CFG_MARKER_TAB_REGISTRY)
    If sr = 0 Then
        GetDevModeTabs = Array()
        Exit Function
    End If

    ' First pass: count hidden tabs
    Dim dr As Long
    dr = sr + 2
    Dim cnt As Long
    cnt = 0
    Do While Len(Trim(CStr(wsConfig.Cells(dr, TREG_COL_TABNAME).Value))) > 0
        If StrComp(Trim(CStr(wsConfig.Cells(dr, TREG_COL_VISIBLE).Value)), _
                   "Hidden", vbTextCompare) = 0 Then
            cnt = cnt + 1
        End If
        dr = dr + 1
    Loop

    If cnt = 0 Then
        GetDevModeTabs = Array()
        Exit Function
    End If

    ' Second pass: collect names
    Dim result() As String
    ReDim result(0 To cnt - 1)
    Dim idx As Long
    idx = 0
    dr = sr + 2
    Do While Len(Trim(CStr(wsConfig.Cells(dr, TREG_COL_TABNAME).Value))) > 0
        If StrComp(Trim(CStr(wsConfig.Cells(dr, TREG_COL_VISIBLE).Value)), _
                   "Hidden", vbTextCompare) = 0 Then
            result(idx) = Trim(CStr(wsConfig.Cells(dr, TREG_COL_TABNAME).Value))
            idx = idx + 1
        End If
        dr = dr + 1
    Loop
    GetDevModeTabs = result
End Function
