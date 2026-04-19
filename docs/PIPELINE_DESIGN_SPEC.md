# Pipeline Full Design Specification

**Status:** DEFERRED — Build after FM ships
**Date:** 2026-03-26
**Context:** Designed during Phase 6A planning. Pipeline Light (Dashboard summary block) ships in Phase 6A. This document preserves the full design for a future phase.

---

## Locked Decisions

| ID | Decision | Answer |
|---|---|---|
| P1 | When to build | Post-FM phase (Pipeline Light ships in 6A) |
| P2 | Where status lives | Both — summary on Dashboard, detail on Pipeline tab |
| P3 | Configurability | Two-layer: kernel steps hardcoded, domain/extension steps config-driven |
| P4 | Execution model | Ordered list with dependency validation (extensible to full DAG — same prerequisite schema, smarter executor) |
| P5 | Run history | Configurable — default last-run-only, opt-in to full history |

---

## Architecture: Two-Layer Pipeline

### Layer 1 — Kernel Pipeline (hardcoded, invariant)

Steps 0-6 that every model needs. NOT configurable — they always run in order. KernelEngine owns this layer.

| Step | Name | Output |
|---|---|---|
| 0 | Bootstrap | Workbook with tabs |
| 1 | LoadConfig | Runtime arrays |
| 2 | Validate | Pass/fail |
| 3 | Compute | Output arrays |
| 4 | WriteDetail | Detail sheet |
| 5 | WriteCSV | CSV file |
| 6 | WriteSummary | Summary sheet |

### Layer 2 — Domain Pipeline (config-driven)

Additional steps registered by domain code and extensions. Defined in `pipeline_config.csv`. Each step declares prerequisites, an executor, and a SortOrder.

The kernel executes Layer 2 by walking SortOrder ascending, checking prerequisites before each step. If a prerequisite is not COMPLETE, the step is SKIPPED with WARN.

---

## pipeline_config.csv Schema

| Column | Type | Required | Description |
|---|---|---|---|
| StepID | String | YES | Unique identifier (e.g., "DOM_QUARTERLY_AGG") |
| StepName | String | YES | Human-readable name |
| ModuleName | String | YES | VBA module containing the executor |
| FunctionName | String | YES | Function to call via Application.Run |
| Prerequisites | String | NO | Comma-separated StepIDs that must be COMPLETE. Kernel step names allowed. |
| SortOrder | Long | YES | Execution order (ascending) |
| Enabled | Boolean | YES | TRUE = runs, FALSE = skipped |
| Category | String | NO | "Transform", "Extension", "Domain", "ETL" |
| Description | String | NO | Human-readable description |

### Sample

```csv
"StepID","StepName","ModuleName","FunctionName","Prerequisites","SortOrder","Enabled","Category","Description"
"DOM_QUARTERLY_AGG","Quarterly Aggregation","SampleDomainEngine","AggregateToQuarterly","STEP_COMPUTE","100","TRUE","Transform","Aggregate monthly Detail to quarterly"
"DOM_FORMULA_REFRESH","Refresh Formula Tabs","KernelFormula","RefreshFormulaTabs","DOM_QUARTERLY_AGG","200","TRUE","Transform","Update formula tabs and named ranges"
"EXT_REPORTGEN","Generate Report","Ext_ReportGen","ReportGen_Execute","STEP_SUMMARY,DOM_FORMULA_REFRESH","900","TRUE","Extension","PDF report generation"
```

---

## KernelPipeline.bas Module Design

```vba
' --- Execution ---
Public Sub RunDomainPipeline(ByRef outputs() As Variant)
  ' Reads pipeline_config. Filters Enabled=TRUE. Sorts by SortOrder.
  ' For each step: check prerequisites → execute → record status.
  ' Failures logged SEV_ERROR with bypass. Pipeline continues.

Public Sub RecordKernelStepStatus(stepName As String, status As String, duration As Double)
  ' Called by KernelEngine after each kernel step. Stores for prerequisite checking.

' --- Display ---
Public Sub RenderPipelineTab()
  ' Full step-by-step view on Pipeline tab.
  ' Kernel steps + domain steps. Status/duration/prerequisites/notes.
  ' Color coding: green=COMPLETE, red=FAILED, grey=SKIPPED, yellow=BYPASSED.

Public Sub RenderDashboardSummary(ws As Worksheet, startRow As Long)
  ' Compact 8-10 row summary on Dashboard.

' --- History ---
Public Sub PersistRunHistory()
  ' If PipelineHistoryEnabled=TRUE: append to PipelineHistory sheet.
  ' Columns: RunID, RunTimestamp, StepID, StepName, Status, Duration, ErrorDetail.
  ' One row per step per run. Append-only.

Public Function GetPipelineHistoryEnabled() As Boolean
```

---

## Pipeline Tab Layout

```
Row 1:  [Pipeline Status]                              (navy header)
Row 2:  Last Run: 2026-03-26 14:30:05  (3.2 seconds)
Row 3:  (blank)
Row 4:  Step | Category | Status | Duration | Prerequisites | Notes
Row 5:  Bootstrap        | Kernel    | ✓ COMPLETE | 0.1s | —
Row 6:  LoadConfig       | Kernel    | ✓ COMPLETE | 0.3s | Bootstrap
Row 7:  Validate         | Kernel    | ✓ COMPLETE | 0.0s | LoadConfig
Row 8:  Compute          | Kernel    | ✓ COMPLETE | 1.8s | Validate
Row 9:  WriteDetail      | Kernel    | ✓ COMPLETE | 0.4s | Compute
Row 10: WriteCSV         | Kernel    | ✓ COMPLETE | 0.2s | WriteDetail
Row 11: WriteSummary     | Kernel    | ✓ COMPLETE | 0.2s | WriteCSV
Row 12: (separator)
Row 13: Quarterly Agg    | Transform | ✓ COMPLETE | 0.2s | Compute
Row 14: Formula Refresh  | Transform | ✓ COMPLETE | 0.1s | Quarterly Agg
Row 15: Generate Report  | Extension | ✓ COMPLETE | 0.5s | WriteSummary, Formula Refresh
Row 16: (blank)
Row 17: Total: 10 steps — 10 passed, 0 failed, 0 skipped
```

---

## Extensibility to Full DAG

The ordered-list-with-prerequisites model is a subset of a full DAG. To upgrade:

1. **Same config schema.** pipeline_config.csv Prerequisites column already declares the dependency graph.
2. **Replace linear executor with topological sort.** Instead of walking SortOrder, build the dependency graph from Prerequisites, compute a topological order, and execute in that order.
3. **Add conditional execution.** New column `Condition` with expressions like `StepID.RowCount > 0` or `ExtensionActive("MonteCarlo")`.
4. **Add parallel execution.** Steps with no mutual dependencies can run concurrently (limited value in VBA's single-threaded model, but useful for future platform migration).

The prerequisite declarations carry forward unchanged. Only the executor gets smarter.

---

## ETL Coverage

For a complex ETL process, pipeline_config.csv would look like:

```csv
"ETL_INGEST_LOSS","Ingest Loss File","Ext_ETL","IngestLoss","","10","TRUE","ETL","Read loss bordereaux from inbox"
"ETL_INGEST_PREM","Ingest Premium File","Ext_ETL","IngestPremium","","10","TRUE","ETL","Read premium schedule from inbox"
"ETL_VALIDATE_LOSS","Validate Loss Schema","Ext_ETL","ValidateLoss","ETL_INGEST_LOSS","20","TRUE","ETL","Schema validation on loss data"
"ETL_VALIDATE_PREM","Validate Premium Schema","Ext_ETL","ValidatePremium","ETL_INGEST_PREM","20","TRUE","ETL","Schema validation on premium data"
"ETL_JOIN","Join Loss-Premium","Ext_ETL","JoinSources","ETL_VALIDATE_LOSS,ETL_VALIDATE_PREM","30","TRUE","ETL","PolicyID join across sources"
"ETL_STAGE","Write Staging","Ext_ETL","WriteStaging","ETL_JOIN","40","TRUE","ETL","Land joined data on staging sheets"
"STEP_COMPUTE","Run Projections","KernelEngine","ComputeCore","ETL_STAGE","50","TRUE","Kernel","Standard kernel compute pipeline"
```

This gives: parallel file ingestion → independent validation → join → staging → compute. Each step has clear prerequisites. A failure in loss ingestion doesn't block premium ingestion (they're parallel at SortOrder=10). The join step waits for both validations.

---

## granularity_config.csv Addition

```csv
"PipelineHistoryEnabled","FALSE","Set TRUE to persist all pipeline runs to PipelineHistory sheet"
```

---

## Implementation Estimate

| Component | Effort |
|---|---|
| KernelPipeline.bas | ~400 lines |
| pipeline_config.csv schema + loaders | ~100 lines |
| Pipeline tab rendering | ~150 lines |
| Dashboard summary enhancement | ~50 lines (upgrade from Pipeline Light) |
| PipelineHistory sheet | ~100 lines |
| KernelEngine integration | ~50 lines |
| **Total** | **~850 lines, 1 CC session** |
