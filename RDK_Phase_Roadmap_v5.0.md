# RDK Phase Roadmap

**Version:** 5.0
**Date:** 2026-04-03
**Status:** Cleanup & Config Wiring sprint in progress. RBC next.
**Convention:** Each phase = one spec → build → validate → audit cycle.
**Maintenance:** Roadmap updated after each phase completion.

---

## The Three God Rules

Every system, module, and workflow must satisfy all three:

1. **Never Get Stuck** — use what you have, make the best assumption, always produce output
2. **Never Lose Control** — any value, any layer, always overridable, always tracked
3. **Never Work Alone** — preview, annotate, converge, confirm. Collaboration is native.

---

## Completed Phases

| Phase | Name | Version | Key Delivery |
|---|---|---|---|
| 1 | Foundation | v1.0.2 | Bootstrap, compute pipeline, Detail/CSV/Summary, MBP |
| 2 | Persistence | v1.0.2 | Snapshots, PRNG, compare, UserForms, COM protection |
| 3 | Testing | v1.0.3 | 5-tier tests, golden baselines, Prove-It |
| 4 | Observability | v1.0.4 | KernelLint, DiagnosticDump, HealthCheck |
| 5 | Presentation | v1.0.5 | Summary from config, charts, exhibits, display mode toggle |
| 6 | Output + Transforms | v1.0.6 | Print/PDF, Power Pivot, transforms, KernelConfig split |
| 7 | Formula Infrastructure | v1.0.7 | KernelFormula, formula tabs, named ranges, quarterly aggregation |
| 8 | Extension Infrastructure | v1.0.8 | KernelExtension, dynamic domain dispatch, extension registry |
| 9 | CurveLib | v1.0.8 | Ext_CurveLib (CDF functions, interpolation, config-driven params) |
| 10 | ReportGen | v1.0.8 | Ext_ReportGen (PDF generation, cover page, PostOutput hook) |
| 11 | Financial Model | v1.2.0 | 10 FM tabs, InsuranceDomainEngine, 52-column registry, EP-based loss development, QS/XOL reinsurance |
| 12 | FM Detail Tabs | v1.3.0 | UW Program Detail, Other Revenue Detail, Staffing Expense, Other Expense Detail, Sales Funnel, Curve Reference, Loss/Count Triangles, Ins_Tests |
| TD-1 | Technical Debt Sprint | v1.3.0 | KernelFormula split, negative pool guard, config-driven buttons/validation/branding |
| INV | Investor Demo Items | v1.3.0 | Cover Page, Run Metadata, User Guide, Conditional Formatting, Input Validation |
| INFRA | Infrastructure Overhaul | v1.3.0 | KernelWorkspace, KernelTabIO, KernelButtons, 13 new config CSVs, deprecated tabs removed, architecture docs, kernel/domain separation |

---

## Active

| Phase | Name | Key Delivery |
|---|---|---|
| **CLEANUP** | **Cleanup & Config Wiring** | **KernelSnapshot split (→ +KernelSnapshotIO). Deprecated tab ref cleanup. Wire workspace_config, msgbox_config (12 entries), display_aliases. Regression tab capture in workspace saves. Compare Workspaces button stub. Sample config update. RBC Capital Model tab.** |

---

## Execution Queue (In Order)

| Phase | Name | Effort | Key Delivery |
|---|---|---|---|
| BCAR | AM Best Capital Model | ~1 week | Factor-based BCAR tab (formula tab, same pattern as RBC). Stochastic BCAR using Monte Carlo extension (Ext_MonteCarlo.bas). Compare factor vs stochastic results on the same portfolio. First extension build since Phase 9-10. |
| WS-COMPARE | Workspace Comparison | ~1.5 days | Rendering module for side-by-side workspace comparison. Reads two workspace versions' regression_tabs/ CSVs. Annual totals, Base/Variant/Delta/Delta%. Uses existing workspace + KernelTabIO infrastructure. |
| DATA-DESIGN | Data Model Deep Dive | ~1 week | Expand Insurance_Data_Model_Architecture_v0.1.md into full build-ready spec. Adversarial review. Resolve 10 open questions. Define canonical schema, DQ rules config, source mapping config, segmentation rules config, materiality config. Define RDK prototype tab layout and module architecture. |
| DATA-BRONZE | Bronze Layer Prototype | ~3 days | CSV ingestion to Raw Data tab. DQ checks from dq_rules_config.csv. Quarantine tab + DQ Log. Source system mapping from source_mapping_config.csv. Comment log (God Rule #3). |
| DATA-SILVER | Silver Layer Prototype | ~1 week | Resolution engine (Ins_DataResolution.bas). Granularity/timing/type normalization. Assumption engine with assumption log. Adjustment layer (God Rule #2). Segmentation rules engine. Preview-before-commit. |
| DATA-GOLD | Gold Layer Prototype | ~1 week | Queryable warehouse views. Hierarchy drill-down. Multi-basis transformations (SAP/GAAP/Management). Materiality tracking. Convergence dashboard. Integration with FM formula tabs and loss development pipeline. |
| DATA-PIPELINE | Pipeline Orchestration | ~3 days | Run history. Restatement tracking. Point-in-time snapshots vs restated history. Close cycle management. Convergence velocity tracking. |

---

## Module Size Remediation (As Needed)

| Module | Current Size | Split Plan | Urgency |
|---|---|---|---|
| KernelSnapshot | 62.9KB | → KernelSnapshot + KernelSnapshotIO (in CLEANUP) | **Critical — in progress** |
| KernelFormSetup | 60.6KB | Extract input-tab rendering into KernelFormSetupIO | High |
| KernelBootstrap | 60.8KB | Extract cover-page builder + branding into KernelBootstrapUI | High |
| KernelFormulaWriter | 58.9KB | Extract ApplyHealthFormatting + ApplyInputValidation into KernelFormulaValidation | High |
| InsuranceDomainEngine | 54.7KB | Extract actuarial math into Ins_Actuarial (TD-07) | Medium |
| KernelEngine | 53.4KB | Extract run metadata + dashboard updates into KernelRunMetadata | Medium |
| KernelTabs | 53.2KB | Remove deprecated tab code (in CLEANUP), should drop to ~45KB | Medium |

---

## Kernel Evolution (Backlog)

| Item | Dependencies | Key Delivery |
|---|---|---|
| KE-01 | Second model build | Evolve Prove-It to support cross-tab cell references (e.g., BS_CHECK=0). Evolve Exhibits to reference any tab (not just Detail). |

---

## Future Phases

| Phase | Name | Dependencies | Key Delivery |
|---|---|---|---|
| 13 | Pipeline Full | PIPELINE_DESIGN_SPEC.md | KernelPipeline.bas, pipeline_config.csv, Pipeline tab, prerequisite validation, run history |
| 14 | Remaining Extensions | Phase 8 infrastructure | Sensitivity, Correlation, Bootstrap, Optimization, DistFit, TimeSeries, ScenarioGen |
| 15 | Onboarding + Repo Reorg | Phase 12 FM as test case | Questionnaire, config generator, DomainEngine scaffold, integration checklist |
| 16 | Data Model + Reserving | DATA-GOLD complete | Insurance data hub integration, actuarial reserving module |
| 17 | Remaining Domain Models | Phase 15 onboarding | ETL Explorer, Surplus Taxes, Licensing |

---

## Future Considerations

| Item | Notes |
|---|---|
| Kernel Add-in Split (12i) | Move kernel modules into RDK_Kernel.xlam. Deferred indefinitely. Revisit if single-project VBA load causes issues. |
| Config-Driven Tab Lifecycle | Discover all tabs from tab_registry. Remove hardcoded TAB_* constants. Partially addressed in CLEANUP. |

---

## Decision Log

| ID | Decision | Date |
|---|---|---|
| DB-01 | Dev Mode gates Dashboard button visibility | 2026-04-02 |
| DB-02 | User buttons: Run Model, Save/Load, Export Report, Compare Workspaces, Dev Mode Toggle | 2026-04-03 |
| SC-08 | Drop standalone 12b scenario comparison. Extend workspaces instead. | 2026-04-03 |
| SC-09 | Workspace saves automatically capture regression tabs | 2026-04-03 |
| RBC-01 | RBC only — defer AM Best and Internal Capital models to BCAR phase | 2026-04-03 |
| RBC-02 | Formula tab via formula_tab_config.csv — no new VBA | 2026-04-03 |
| RBC-03 | Quarterly computation | 2026-04-03 |
| KE-02 | Removed — FM domain modules are correctly domain-specific, not kernel migration candidates | 2026-04-03 |
| 12i | Deferred indefinitely — add-in split not needed | 2026-04-03 |
| GOD-01 | Three God Rules adopted: Never Get Stuck, Never Lose Control, Never Work Alone | 2026-04-03 |
| DATA-01 | Medallion architecture (Bronze/Silver/Gold) with Resolution Engine and Adjustment Layer | 2026-04-03 |
| DATA-02 | Prototype in Excel/VBA (RDK), then port to AIOS (PostgreSQL) | 2026-04-03 |
| DATA-03 | Annotation via comment_log.csv in prototype, web-based in AIOS | 2026-04-03 |
| DATA-04 | Arbiter context-dependent: data uploader for raw data, business owner for logic | 2026-04-03 |
| DATA-05 | Materiality configurable per metric with fallback default | 2026-04-03 |
| DATA-06 | Preview shows immediate next layer only (v1) | 2026-04-03 |

---

## Codebase Stats (Current)

```
Version:           1.3.0
VBA Modules:       37 (27 kernel + 2 extension + 3 domain + 5 companion)
Total VBA:         1.1MB
Config CSVs:       31
Formula rows:      2,487
Tabs (registry):   28
Named ranges:      97
Bugs logged:       134
Anti-patterns:     63
Patterns:          31
Modules >50KB:     7 (target: 0 after all splits)
Adversarial score: 7.8 (ChatGPT R2) / 8.0 (arbiter-adjusted)
God Rules:         3
```
