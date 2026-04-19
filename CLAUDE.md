# CLAUDE.md — RDK Project Context

**Read this file completely before doing anything. This is your project bible.**
**Then read SESSION_NOTES.md for what changed since your last session.**

## What This Is

The Phronex Rapid Development Kit (RDK) is a zero-dependency rapid prototyping framework for building Excel/VBA models with CSV persistence. Blessed stack: Excel + VBA + CSV + PowerShell. Nothing else. No Python. No npm. No SQLite. No internet.

**Current state:** Phase 12B (Expense, Staffing, Sales Funnel) COMPLETE at v1.3.0. CLEANUP sprint + RBC Capital Model + Compute Cache in progress.

### Three God Rules (Foundational Constraints)
1. **Never Get Stuck** -- Always produce output. Use best available assumptions. Downstream shape is fixed.
2. **Never Lose Control** -- Users can override any value at any layer, with tracking, without breaking the system.
3. **Never Work Alone** -- Preview, annotate, converge, confirm. Multi-party collaboration with real-time feedback.

Full text in `docs/RDK_Developer_Flywheel_v1.2.md`. Every design decision must satisfy all three.

### Key Design Documents
- `docs/RDK_Phase_Roadmap_v5.0.md` -- current roadmap (BCAR, DATA phases, decision log)
- `docs/Insurance_Data_Model_Architecture_v0.1.md` -- data architecture design
- `docs/MODEL_REQUIREMENTS_TEMPLATE.md` -- standardized model requirements template

## Architecture

### Four-Layer Stack
1. **Kernel** (hardcoded VBA) — knows HOW. Never edited by developer.
2. **Configuration** (CSV files) — tells kernel WHAT. Adding a metric = one config row.
3. **Domain Logic** (Domain*.bas) — only code the developer writes.
4. **UI** (Excel workbook) — generated from configuration by KernelBootstrap.

### Domain Contract (Canonical)
Every Domain*.bas must implement:
```
Initialize()              — called once at bootstrap
Validate() As Boolean     — called pre-run; return False to halt
Reset()                   — called before and after computation
Execute(outputs())        — write Incremental + Dimension columns only
```
Domain reads inputs via `KernelConfig.InputValue()`. No separate inputs array.
Domain references columns via `KernelConfig.ColIndex()` only. No magic numbers.
Domain NEVER computes Derived or cumulative values — kernel handles both.

### Pipeline Steps
0=Bootstrap, 1=LoadConfig, 2=Validate, 3=Compute, 4=WriteDetail, 5=WriteCSV, 6=WriteSummary

### Manual Bypass Protocol (AD-39: Hard Requirement)
Every pipeline step has a manual bypass. Every error message includes bypass instructions (AP-46).
ResumeFrom(stepNumber) validates prior artifacts and resumes the pipeline.
KernelConfig has fallback mode — reads from Config sheet if arrays fail to load (AP-45).
ExportSchemaTemplate writes blank Detail templates for manual data entry.

## Files That Matter

| File | What It Does |
|------|-------------|
| CLAUDE.md | Project bible. Read FIRST. |
| SESSION_NOTES.md | What changed since your last session. Read SECOND. |
| data/anti_patterns.csv | 63 rules (AP-01 through AP-63). READ BEFORE CODING. |
| data/patterns.csv | 31 established patterns (PT-001 through PT-031). FOLLOW THESE. |
| data/bug_log.csv | 87 bugs logged (BUG-001 through BUG-087). APPEND when you fix something. |
| config_sample/ | Working toy model config (3 entities, 12 periods) |
| config_insurance/ | Insurance model config (10 programs, 60 periods) |
| config_blank/ | Headers + kernel infrastructure (bootable but no domain data) |
| config/ | Runtime copy (seeded from sample or blank by Setup.bat) |
| docs/archive/ | Historical files from prior phases. Reference only. NEVER modify archived files. |

### Protected Files — NEVER Delete, Move, or Archive
These files MUST remain at repo root after every phase. Do NOT move them to docs/archive/:
- **CLAUDE.md** — you read this first every session
- **SESSION_NOTES.md** — you read this second every session
- **config_sample/** and **config_blank/** — required by Setup.bat
- **data/*.csv** — institutional knowledge, append-only
- **engine/*.bas** — source modules, modify in-place only
- **scripts/** — bootstrap and utility scripts
- **docs/*.md** (non-archive) — design specs, roadmaps, flywheel. These are LIVING documents. Do NOT delete or move to archive. Only docs/archive/ files are frozen historical snapshots.

See Developer Flywheel v1.2 §F.4 for the full protected files doctrine and documented incidents.

## Critical Rules

### Before Writing Any Code
1. Read CLAUDE.md (this file)
2. Read SESSION_NOTES.md — what changed since your last session
3. Read data/anti_patterns.csv — all 63 entries
4. Read data/patterns.csv — all 31 entries
5. Understand what you're changing and why

### Anti-Patterns (Top 12 — Most Likely to Hit)
- **AP-06:** No non-ASCII in .bas files (VBA import fails silently)
- **AP-07:** No strings starting with =,+,-,@ in .Value (Excel interprets as formula)
- **AP-08:** All column indices via ColIndex() only. NO MAGIC NUMBERS. EVER.
- **AP-18:** No ReDim Preserve in loops
- **AP-34:** No sized-array Dim after executable code
- **AP-42:** Domain code never stores cumulative values
- **AP-43:** Every Domain*.bas implements all 4 contract functions
- **AP-45:** No step depends on in-memory state without fallback to visible artifact
- **AP-46:** Every error message includes manual bypass instructions
- **AP-50:** All cell writes of strings starting with = must set NumberFormat=@ first
- **AP-51:** Bootstrap must ensure required kernel tabs exist regardless of tab_registry contents
- **AP-53:** Pipeline entry points must show MsgBox for all outcomes
- **AP-54:** Bootstrap zombie-kill and /automation must not corrupt COM add-in state

### Code Patterns (Must Follow)
- **PT-001:** Array batch write to Excel (never cell-by-cell)
- **PT-002:** Atomic file write (temp → verify → rename)
- **PT-003:** ColIndex() for all column references
- **PT-013:** Config fallback mode (read from Config sheet if arrays fail)
- **PT-014:** Pipeline step markers on Config sheet
- **PT-015:** Manual bypass instructions in every error at SEV_ERROR or SEV_FATAL
- **PT-016:** File-lock check before overwrite
- **PT-017:** Severity color-coding in error logs
- **PT-018:** Required-tab fallback in bootstrap

### Before Packaging Any Delivery
- [ ] All .bas files have CRLF line endings
- [ ] No non-ASCII characters in any .bas file
- [ ] Sub/Function count = End Sub/End Function count in every module
- [ ] No magic numbers (grep for numeric array indexing)
- [ ] All modules under 64KB (AD-09). Modules over 50KB are at WARN threshold — monitor but allowed.
- [ ] ZIP is flat (no wrapper folder)
- [ ] Filename includes timestamp: `RDK_PhaseN_YYYYMMDD_HHMMSS.zip`
- [ ] Bug log updated with any bugs found and fixed
- [ ] Anti-patterns updated if new patterns discovered
- [ ] DELIVERY_SUMMARY.md or AMENDMENT_SUMMARY.md updated
- [ ] CLAUDE.md counts updated

## Config Tables

### Phase 1 (active): 4
| Table | Purpose |
|-------|---------|
| column_registry.csv | Defines every column: Name, DetailCol, CsvCol, FieldClass, DerivationRule |
| input_schema.csv | Defines every input parameter: Section, ParamName, Row, DataType, Default |
| tab_registry.csv | Defines every workbook tab: TabName, Protected, Visible, SortOrder |
| granularity_config.csv | Time axis and scale: TimeHorizon, MaxEntities, DefaultSummaryView |

### Phase 2 (active): 2
| Table | Purpose |
|-------|---------|
| repro_config.csv | Reproducibility: PRNG seed, float precision, deterministic mode |
| scale_limits.csv | Max row/column limits, CSV-only mode thresholds |

### Phase 3 (active): 1
| Table | Purpose |
|-------|---------|
| prove_it_config.csv | Prove-It audit checks: CheckID, CheckType, MetricA/B/C, Operator, Tolerance |

### Phase 5A (active): 4
| Table | Purpose |
|-------|---------|
| summary_config.csv | Summary tab metrics: MetricName, SectionName, SortOrder, Format, ShowInSummary |
| chart_registry.csv | Dashboard charts: ChartID, ChartName, ChartType, MetricName, GroupBy, Width, Height, Enabled |
| exhibit_config.csv | Exhibit tables: ExhibitID, ExhibitName, MetricList, GroupBy, IncludeTotal, Format, Enabled |
| display_mode_config.csv | Display mode toggle: Setting, Value (DefaultMode, ToggleEnabled, labels) |

### Phase 5B (active): 3
| Table | Purpose |
|-------|---------|
| print_config.csv | Print/PDF settings: TabName, Orientation, FitToPages, PaperSize, Headers, IncludeInPDF, PrintOrder |
| data_model_config.csv | Power Pivot data model: PowerPivotEnabled, FactTableSource, RefreshOnRun |
| pivot_config.csv | Pivot tables: PivotID, PivotName, SourceTable, RowField, ColField, ValueField, AggFunction, Enabled |

### Phase 5C (active): 2
| Table | Purpose |
|-------|---------|
| formula_tab_config.csv | Formula tab layout: TabName, RowID, Row, Col, CellType, Content, Format, FontStyle, FillColor, FontColor, ColSpan, BorderBottom, BorderTop, Indent, Comment |
| named_range_registry.csv | Named ranges: RangeName, TabName, RowID, CellAddress, RangeType, Description |

### Phase 6A (active): 2
| Table | Purpose |
|-------|---------|
| extension_registry.csv | Extension definitions: ID, Module, EntryPoint, HookType, SortOrder, Activated, MutatesOutputs, RequiresSeed, Description |
| curve_library_config.csv | Curve library parameters: CurveID, Category, CurveType, Distribution, Param1-3, MaxAge, InterpMethod, Description |

### Full kit: 18 total
column_registry, input_schema, tab_registry, granularity_config, summary_config, chart_registry, exhibit_config, print_config, prove_it_config, data_model_config, pivot_config, repro_config, display_mode_config, scale_limits, formula_tab_config, named_range_registry, extension_registry, curve_library_config.

**CRITICAL:** config_blank/ must be a valid bootable config, not an empty skeleton. It must contain all kernel infrastructure (tab definitions, granularity settings) — only domain-specific entity data should be empty (AP-51).

## Deterministic Fixture (Verified by Python + Excel)

Formula: Revenue = Units × UnitPrice × (1 + MonthlyGrowth)^(Period-1)

| Entity | Period | Revenue | COGS | GrossProfit | GPMargin |
|--------|--------|---------|------|-------------|----------|
| Product A | 1 | 25000.000000 | 15000.000000 | 10000.000000 | 0.400000 |
| Product A | 12 | 27589.436941 | 16553.662165 | 11035.774777 | 0.400000 |
| Product B | 12 | 26409.895818 | 15845.937491 | 10563.958327 | 0.400000 |
| Product C | 12 | 28505.301981 | 17103.181188 | 11402.120792 | 0.400000 |

Total Revenue (all 36 rows): **944,307.512360** (verified in Excel CSV output)
GPMargin: **0.400000** for every row (constant because COGS is fixed % of Revenue)
Verify against CSV output (6-decimal precision), not Detail tab display format.

## Development Process

### Developer Triangle
- **Claude Online (Opus):** Architecture, spec, review, triage. Never writes final code.
- **Claude Code (you):** Build, test, fix, package. Never makes architecture decisions.
- **Ethan:** Validates in Excel. Makes arbiter decisions. Domain expert.
- **ChatGPT:** Adversarial reviewer. Reviews specs before build. Never builds.

### 5-Step Phase Cycle
1. **Spec** — Claude Online produces phase build prompt + configs + contracts
2. **Adversarial Review** — ChatGPT reviews (0 P0, 0 P1 to proceed)
3. **Build** — Claude Code builds (`claude --dangerously-skip-permissions --model claude-opus-4-6 --max-turns 200`)
4. **Validate** — Ethan walks validation gates in Excel
5. **Review + Capture** — Claude Online updates knowledge artifacts, declares phase complete

### Bug Logging Format
Append to data/bug_log.csv:
```
"BUG-NNN","Description","RootCause","Fix","AntiPattern","Phase","DateFixed"
```

### Delivery Format
- Flat ZIP with date-time stamp
- Include DELIVERY_SUMMARY.md or AMENDMENT_SUMMARY.md
- No nested folders in the ZIP
- No old ZIP artifacts in the repo

## Current Counts

| Item | Count |
|------|-------|
| Kernel modules | 32 (KernelConstants, KernelConfig, KernelConfigLoader, KernelEngine, KernelOutput, KernelCsvIO, KernelBootstrap, KernelBootstrapUI, KernelRandom, KernelSnapshot, KernelSnapshotIO, KernelCompare, KernelFormHelpers, KernelFormSetup, KernelFormSetup2, KernelTests, KernelTestHarness, KernelProveIt, KernelLint, KernelDiagnostic, KernelHealth, KernelTabs, KernelPrint, KernelTransform, KernelFormula, KernelFormulaWriter, KernelExtension, KernelWorkspace, KernelWorkspaceExt, KernelTabIO, KernelButtons, KernelAssumptions) |
| Domain companion modules | 5 (Ins_GranularCSV, Ins_QuarterlyAgg, Ins_Triangles, Ins_Tests, Ins_Presentation) |
| Config tables (full kit) | 32 |
| Extension modules | 2 (Ext_CurveLib.bas, Ext_ReportGen.bas) |
| Domain modules | 3 (DomainEngine.bas stub, SampleDomainEngine.bas, InsuranceDomainEngine.bas) |
| Domain companion modules | 5 (Ins_GranularCSV.bas, Ins_QuarterlyAgg.bas, Ins_Triangles.bas, Ins_Tests.bas, Ins_Presentation.bas) |
| Domain test modules | 0 (SampleDomainTests.bas deleted in BUG-035 — was corrupt duplicate) |
| Anti-patterns | 66 (AP-01 through AP-66) |
| Patterns | 34 (PT-001 through PT-034) |
| Bugs logged | 160 (BUG-001 through BUG-160) |
| Config tables (Phases 1-10) | 18 |
| Config tables (full kit) | 32 |
| Config directories | 3 (config_sample, config_insurance, config_blank) |
| Arbiter decisions | 51 (all locked) |

## Phase Roadmap

See `docs/RDK_Phase_Roadmap_v5.0.md` for full roadmap with dependency chain and cross-reference.

| Phase | Status | What |
|-------|--------|------|
| 1 | **COMPLETE** (v1.0.2) | Foundation |
| 2 | **COMPLETE** (v1.0.2) | Persistence |
| 3 | **COMPLETE** (v1.0.3) | Testing |
| 4 | **COMPLETE** (v1.0.4) | Observability |
| 5 | **COMPLETE** (v1.0.5) | Presentation |
| 6 | **COMPLETE** (v1.0.6) | Output + Transforms |
| 7 | **COMPLETE** (v1.0.7) | Formula Infrastructure |
| 8 | **COMPLETE** (v1.0.8) | Extension Infrastructure |
| 9 | **COMPLETE** (v1.0.8) | CurveLib |
| 10 | **COMPLETE** (v1.0.8) | ReportGen |
| **11A** | **COMPLETE** (v1.1.0) | Insurance DomainEngine + Core Tabs |
| **11B** | **COMPLETE** (v1.2.0) | Financial Model (remaining FM tabs) |
| **12A** | **COMPLETE** (v1.3.0) | Detail Tabs (UW Program, Other Revenue, Software Income) |
| **12B** | **COMPLETE** (v1.3.0) | Expense, Staffing, Sales Funnel + PD-05 Sign Convention |
