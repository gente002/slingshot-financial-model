# Insurance Data Model — Architecture Specification

**Version:** 0.1 (Initial Design)
**Date:** 2026-04-03
**Status:** DRAFT — to be expanded via adversarial review
**Implementation:** RDK prototype (Excel/VBA), then AIOS (PostgreSQL)

---

## Design Philosophy

**The God Rule:** Use the information you have and make the best assumption available to get to the downstream structure and shape so we can report the same every time.

**Corollaries:**
1. Never get stuck — always produce output, even with incomplete data
2. Always document what was assumed vs what was known
3. Every assumption is reversible — when real data arrives, it replaces the assumption
4. The downstream shape is fixed — upstream variability is resolved before it reaches Gold
5. Data quality issues quarantine rows, they don't halt the pipeline
6. **The Backdoor Principle:** Users always have a sanctioned path to manually override any value at any layer — temporarily, with tracking, without breaking the pipeline
7. **Two paths to the same destination:** Automated assumptions (system-generated) and manual adjustments (user-entered) both produce the same downstream shape. The Gold layer is indifferent to how a value was determined — but lineage always traces the source.

---

## Adjustment Layer — The "Backdoor" Principle

### Core Principle

**Every layer of the architecture supports manual overrides.** Users always have the ability to layer in a temporary adjustment to fix something they know is wrong — quickly, without waiting for the upstream data to be corrected. This is not a workaround. It is a first-class architectural feature.

Insurance operations have hard deadlines (monthly close, quarterly filings, board reporting). When data is wrong and the close is tomorrow, the user needs a sanctioned path to "just get it done" — not a shadow spreadsheet.

### Two Types of Assumptions

The architecture treats all non-source-data values as **assumptions**, regardless of how they were created:

| Type | Created By | Example | Lifecycle |
|---|---|---|---|
| **Automated Assumption** | Resolution Engine | Mid-month extrapolated to EOM, aggregate allocated to policies, IBNR allocated by EP weight | Replaced automatically when better data arrives |
| **Manual Adjustment** | User (the "backdoor") | Override a reserve amount, adjust a premium allocation, correct a misclassified claim, plug a known gap | Replaced manually when user confirms the fix is permanent OR when upstream data corrects it |

**Both produce the same result:** downstream Gold data has the correct shape, and reporting works identically regardless of whether the value came from source data, an automated assumption, or a manual adjustment.

### Adjustment Registry

Every manual adjustment is tracked in an **Adjustment Log**:

```
adjustment_id          Unique ID (ADJ-001, ADJ-002, ...)
layer                  Bronze / Silver / Gold
entity_type            Policy / Claim / Transaction / Program / Aggregate
entity_id              FK to the adjusted record
field_name             Which field was adjusted
original_value         What the system computed (or NULL if the field was empty)
adjusted_value         What the user entered
adjustment_reason      Free text — why the adjustment was made
adjustment_type        Override / Plug / Reclass / Correction / Estimate
adjusted_by            User name
adjusted_date          When the adjustment was made
expiration_date        When this adjustment should be reviewed (NULL = permanent)
status                 Active / Expired / Replaced / Investigated
replaced_by            Source data that superseded this adjustment (NULL if still active)
replaced_date          When source data replaced this adjustment
batch_id               Which close cycle this belongs to (e.g., "2026-Q1-Close")
```

### Where Adjustments Can Be Made

| Layer | What Can Be Adjusted | Example |
|---|---|---|
| **Bronze** | Raw data corrections before normalization | "This premium should be $1.2M, not $12M (decimal error in source)" |
| **Silver** | Normalized values, assumption overrides | "Override the EOM extrapolation — I know the actual EOM number" |
| **Gold** | Reporting aggregates, top-side adjustments | "Add $500K to this program's IBNR — actuarial review determined it's under-reserved" |

### Adjustment Precedence Rules

When multiple values exist for the same field:

```
Priority 1: Manual adjustment (Active) — user knows best during close
Priority 2: Source data (latest batch) — real data beats assumptions
Priority 3: Automated assumption — system-generated estimate
Priority 4: Prior period carry-forward — stale but better than nothing
```

If a manual adjustment exists AND new source data arrives:
1. Flag the adjustment as "Potentially Superseded"
2. Show the user: "Source data now says X, your adjustment says Y. Keep adjustment or accept source?"
3. User decides → adjustment status changes to Replaced or Confirmed

### Close Cycle Management

During a financial close:

1. **Pre-close:** Pipeline runs normally. Automated assumptions fill gaps.
2. **Close window opens:** Users can enter manual adjustments. Each is tagged with the batch_id (e.g., "2026-Q1-Close").
3. **Close review:** Management reviews all active adjustments for the period. Dashboard shows: "12 manual adjustments, 3 need investigation."
4. **Close finalized:** Snapshot taken. Adjustments are frozen for this period.
5. **Post-close:** Source data corrections flow in. Adjustments are flagged for review. Next period starts clean.

### Adjustment Hygiene

The architecture tracks adjustment debt:

**Adjustment Dashboard:**
- Count of active adjustments by layer, by type, by age
- Adjustments older than 90 days → flagged for investigation
- Adjustments that have been carried forward >2 close cycles → escalated
- Trend: are adjustments increasing or decreasing over time?

**Cleanup workflow:**
1. After each close, review active adjustments
2. For each: was the upstream data fixed? → Replace adjustment with source data
3. For each: is this a recurring issue? → Create a permanent rule in the Resolution Engine
4. Goal: zero adjustments carried forward indefinitely. Every adjustment is either replaced by source data or converted to a permanent rule.

### How This Changes the God Rule

The God Rule ("use what you have, make the best assumption available") now has two mechanisms:

```
                    ┌─────────────────────┐
                    │   Source Data        │
                    │   (Bronze)           │
                    └──────────┬──────────┘
                               │
                    ┌──────────▼──────────┐
                    │  Resolution Engine   │
                    │  (Automated          │
                    │   Assumptions)       │
                    └──────────┬──────────┘
                               │
                    ┌──────────▼──────────┐
                    │  Adjustment Layer    │◄── User overrides
                    │  (Manual             │    (the "backdoor")
                    │   Assumptions)       │
                    └──────────┬──────────┘
                               │
                    ┌──────────▼──────────┐
                    │  Gold Warehouse      │
                    │  (Same shape,        │
                    │   always queryable)  │
                    └─────────────────────┘
```

The Gold layer doesn't know or care whether a value came from source data, an automated assumption, or a manual adjustment. It always has the same shape. But the **lineage** is always traceable — for any Gold value, you can answer: "Where did this number come from? Was it source data, a system assumption, or a user override? When? Why?"

### RDK Prototype Implementation

| Component | RDK Implementation |
|---|---|
| Adjustment Log | adjustment_log.csv in data/ directory |
| Bronze adjustments | "Adjustments" section on Raw Data tab — user enters corrections with reason |
| Silver adjustments | "Overrides" section on Normalized Data tab — user overrides Resolution Engine output |
| Gold adjustments | "Top-Side Adjustments" tab — user enters aggregate adjustments with reason |
| Adjustment Dashboard | Section on Dashboard (or dedicated tab) showing active adjustment count, age, trend |
| Close cycle | Managed via workspace versions — save workspace at close, adjustments frozen |

---

## Medallion Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                        BRONZE                                │
│  Raw ingestion. Source-faithful. No transformations.          │
│  Schema: source-specific (each source has its own shape)     │
│  Grain: as-received (could be transaction, policy, or agg)   │
│  Timing: as-received (could be mid-month, EOM, ad hoc)       │
│                                                              │
│  + Data Quality checks → Quarantine + DQ Log                │
│  + Source system metadata (source, received_at, batch_id)    │
└──────────────────────────┬──────────────────────────────────┘
                           │
                    Resolution Engine
                    (Granularity, Timing, Type, Assumptions)
                           │
┌──────────────────────────▼──────────────────────────────────┐
│                        SILVER                                │
│  Normalized. Canonical schema. Assumptions applied.          │
│  Schema: canonical insurance data model (fixed)              │
│  Grain: standardized (policy-period or exposure-month)       │
│  Timing: standardized (end-of-month)                         │
│  Type: standardized (incremental monthly)                    │
│                                                              │
│  + Assumption log (what was assumed, when, why)              │
│  + Lineage (every Silver row traces to Bronze source)        │
│  + Rules engine output (segmentation tags applied)           │
│  + Preview capability (staged Silver, not yet committed)     │
└──────────────────────────┬──────────────────────────────────┘
                           │
                    Aggregation + Multi-Basis Views
                           │
┌──────────────────────────▼──────────────────────────────────┐
│                         GOLD                                 │
│  Queryable warehouse. Reporting-ready. Multiple views.       │
│  Schema: dimensional (facts + dimensions)                    │
│  Grain: flexible (query at any level of hierarchy)           │
│  Views: Calendar Year, Accident Year, Report Year            │
│  Basis: Statutory, GAAP, Management                          │
│                                                              │
│  + Segmentation hierarchies (drill-down)                     │
│  + Point-in-time snapshots vs restated history               │
│  + Run history (which pipeline run produced this Gold state) │
└─────────────────────────────────────────────────────────────┘
```

---

## Core Entities

### 1. Program
The top-level business unit. An MGA or insurance program with defined terms.

| Field | Description |
|---|---|
| program_id | Unique identifier |
| program_name | Display name (via alias config) |
| line_of_business | GL, Property, Casualty, Specialty, etc. |
| effective_date | Program inception |
| expiration_date | Program expiration |
| policy_term_months | 12, 6, 3, etc. |
| business_unit | Organizational grouping |
| status | Active, Run-off, Cancelled |

### 2. Policy
Individual insurance policy within a program.

| Field | Description |
|---|---|
| policy_id | Unique identifier |
| program_id | FK to Program |
| policy_number | Source system policy number |
| insured_name | Policyholder |
| effective_date | Policy inception |
| expiration_date | Policy expiration |
| premium_written | Total written premium |
| state | Domicile state |
| coverage_type | Occurrence vs Claims-Made |

### 3. Claim
Loss event tied to a policy.

| Field | Description |
|---|---|
| claim_id | Unique identifier |
| policy_id | FK to Policy |
| claim_number | Source system claim number |
| accident_date | Date of loss |
| report_date | Date claim reported |
| close_date | Date claim closed (NULL if open) |
| reopen_date | Date claim reopened (NULL if never reopened) |
| status | Open, Closed, Reopened |
| claimant_name | Claimant |
| loss_type | Indemnity, ALAE, ULAE |

### 4. Transaction
Financial movement on a claim or policy. This is the atomic unit.

| Field | Description |
|---|---|
| transaction_id | Unique identifier |
| claim_id | FK to Claim (NULL for premium transactions) |
| policy_id | FK to Policy |
| transaction_type | Payment, Reserve Change, Premium, Recovery, Salvage, Subrogation |
| transaction_date | Date of transaction |
| amount | Dollar amount (positive = outflow, negative = recovery) |
| category | Indemnity, ALAE, ULAE, Premium, Commission, Fee |
| accounting_period | Calendar month (YYYYMM) |

### 5. Reinsurance Treaty
Reinsurance structure applied to programs/policies.

| Field | Description |
|---|---|
| treaty_id | Unique identifier |
| treaty_type | QS, XOL (per-occurrence), XOL (aggregate), Facultative |
| program_id | FK to Program (or NULL for portfolio-wide) |
| effective_date | Treaty inception |
| expiration_date | Treaty expiration |
| cession_pct | QS cession percentage (NULL for XOL) |
| attachment | XOL attachment point (NULL for QS) |
| limit | XOL limit (NULL for QS) |
| reinstatements | Number of reinstatements |
| commission_pct | Ceding commission rate |
| fronting_fee_pct | Fronting fee rate |
| reinsurer_name | Counterparty |
| reinsurer_rating | AM Best rating (for R5 credit risk) |

### 6. Exposure
Earned/written/in-force exposure at the policy-month level.

| Field | Description |
|---|---|
| exposure_id | Unique identifier |
| policy_id | FK to Policy |
| accounting_period | YYYYMM |
| written_premium | Premium written in this period |
| earned_premium | Premium earned in this period |
| in_force_premium | Premium in-force at period end |
| written_exposure | Exposure units written |
| earned_exposure | Exposure units earned |

---

## Resolution Engine

The Resolution Engine transforms Bronze data (variable shape) into Silver data (canonical shape). It handles six types of variability:

### R1: Granularity Resolution

| Source Grain | Target Grain | Resolution Rule |
|---|---|---|
| Transaction-level | Policy-period | Aggregate transactions by policy + period |
| Policy-level | Policy-period | Already at target grain |
| Program-aggregate | Policy-period | **Assumption:** allocate pro-rata by earned premium weight. Log assumption. |
| Bordereau (policy-level but no claims) | Policy-period | Map directly; claims will be zero until claim data arrives |

**Tiebreaker:** When both aggregate and transaction data exist for the same program-period, use transaction-level (higher fidelity). Flag the aggregate as "superseded."

### R2: Timing Resolution

| Source Timing | Target Timing | Resolution Rule |
|---|---|---|
| End-of-month | End-of-month | No transformation |
| Mid-month (e.g., as-of 15th) | End-of-month | **Assumption:** extrapolate to EOM using bulk-load method. Paid losses: carry forward (assume no additional payments in remaining days). Reserves: carry forward. Premium: pro-rate earned portion. Log assumption. |
| Quarterly | Monthly | **Assumption:** allocate evenly across 3 months (or use earned premium weights if available). Log assumption. |
| Annual | Monthly | **Assumption:** allocate using 1/12 or seasonal pattern if configured. Log assumption. |
| Ad hoc / irregular | Monthly | Interpolate between known points. Log assumption. |

**Tiebreaker:** When both monthly and quarterly data exist for the same period, use monthly. Flag quarterly as "superseded."

### R3: Type Resolution (Point-in-Time vs Incremental)

| Source Type | Target Type | Resolution Rule |
|---|---|---|
| Incremental (MTD) | Incremental (MTD) | No transformation (target is incremental monthly) |
| QTD | Incremental (MTD) | Month 1 of quarter = QTD value. Month 2 = QTD - Month 1. Month 3 = QTD - Month 1 - Month 2. |
| YTD | Incremental (MTD) | Same pattern — subtract prior months' cumulative. |
| ITD (cumulative) | Incremental (MTD) | Current ITD - Prior month ITD = monthly incremental. |
| Point-in-time balance | Incremental (MTD) | Current balance - Prior balance = monthly change. |

**Tiebreaker:** When both incremental and cumulative exist, use incremental (direct observation). Validate against cumulative (should reconcile). If they don't reconcile, log discrepancy and use incremental.

### R4: Assumption Engine

When data is missing or incomplete, the Assumption Engine fills gaps:

| Scenario | Assumption | Method |
|---|---|---|
| Premium known, losses unknown | Apply expected loss ratio | ELR × Earned Premium = Expected Ultimate |
| Aggregate known, policy detail unknown | Allocate pro-rata by premium weight | Aggregate × (Policy Premium / Total Premium) |
| Mid-month data, need EOM | Extrapolate using bulk-load | Carry known values forward |
| Current month missing, prior month available | Carry forward prior month | Copy prior values, flag as "assumed" |
| IBNR allocation to programs | Allocate by earned premium weight | Total IBNR × (Program EP / Total EP) |
| Operating expense allocation to cost centers | Allocate by configured rules | Rule-based: headcount, premium, equal split |

**Every assumption is tagged:**
```
assumption_id, source_field, assumed_value, assumption_method, 
assumption_date, replaced_by (NULL until real data arrives), 
replaced_date, confidence_level (High/Medium/Low)
```

### R5: Segmentation Rules Engine

Meta-tags applied to Silver data for custom grouping:

| Rule Type | Example |
|---|---|
| LOB mapping | "GL" + "Premises" → "General Liability - Premises/Operations" |
| Geography | State → Region → Territory |
| Size band | Premium < $500K → "Small", $500K-$2M → "Medium", > $2M → "Large" |
| Profitability | Loss Ratio < 60% → "Profitable", 60-80% → "Marginal", > 80% → "Unprofitable" |
| RI applicability | QS treaty applies to Programs 1-5, XOL applies to all |
| Regulatory | Surplus lines vs admitted, NRRA-eligible |
| Custom | User-defined segmentation rules via config |

Rules are config-driven (segmentation_rules_config.csv):
```csv
"RuleID","FieldName","Operator","Value","TagName","TagValue","Priority"
"SEG-001","LOB","=","GL","Segment","General Liability",1
"SEG-002","State","IN","CA,NY,FL,TX","Territory","Tier 1",1
"SEG-003","GWP",">=","2000000","SizeBand","Large",1
```

### R6: Referential Integrity & Orphan Handling

| Scenario | Resolution |
|---|---|
| Claim references unknown policy | Quarantine claim. Log: "Orphan claim — policy not found." Create placeholder policy if configured. |
| Policy references unknown program | Assign to "Unallocated" program. Log assumption. |
| Transaction references unknown claim | Quarantine transaction. Log: "Orphan transaction." |
| Duplicate transactions (same amount, date, claim) | Flag as potential duplicate. Quarantine second occurrence. |

---

## Temporal Management

### Point-in-Time Snapshots

Every pipeline run produces a **snapshot** — the complete state of Gold as of that run. Snapshots are immutable.

```
snapshot_id, run_date, as_of_date, run_parameters, 
record_count, status (Complete/Partial/Failed)
```

### Restated History

When new information arrives that would have changed prior answers:

1. The new data is ingested into Bronze with a `restatement_flag = TRUE`
2. Silver is recomputed for affected periods
3. Gold is updated with restated values
4. The prior snapshot is preserved (immutable)
5. A new snapshot is created with `restatement_of = prior_snapshot_id`
6. Delta between snapshots is computed and logged

### Run History

Every pipeline execution is tracked:

```
run_id, started_at, completed_at, status, 
bronze_records_ingested, silver_records_produced, 
gold_records_updated, quarantine_count, 
assumption_count, error_count, run_parameters
```

---

## Preview Before Commit

Silver data is staged in a **preview area** before committing to Gold:

1. Pipeline runs Bronze → Silver (staged)
2. User reviews staged Silver: row counts, totals, delta vs prior run
3. Data propagation preview: "If committed, these Gold aggregates will change by X"
4. User approves → staged Silver moves to committed Silver → Gold is rebuilt
5. User rejects → staged Silver is discarded, prior state preserved

---

## Hierarchies & Drill-Down

Dimensional hierarchies for Gold-layer queries:

```
Program Hierarchy:
  Company → Business Unit → Program → Policy → Claim → Transaction

Geographic Hierarchy:
  Country → Region → State → County → ZIP

Time Hierarchy:
  Year → Quarter → Month → Week → Day

LOB Hierarchy:
  Department → LOB Group → LOB → Sub-LOB → Coverage

Reinsurance Hierarchy:
  Structure → Treaty → Layer → Segment
```

Each hierarchy is config-driven (hierarchy_config.csv) so users can define custom roll-ups.

---

## Multi-Basis Views

Gold data supports multiple accounting/reporting bases:

| Basis | Description |
|---|---|
| Statutory (SAP) | State regulatory basis. Gross reserves on B/S, RI recoverables as asset. Conservative recognition. |
| GAAP | US GAAP basis. DAC amortization, different reserve recognition. |
| Management | Internal basis. May include IBNR adjustments, ELR selections, management overrides. |
| Economic | Market-value basis. Discounted reserves, fair-value assets. |

The Silver layer stores **one set of facts**. The Gold layer applies **basis-specific transformations** (e.g., DAC calculation for GAAP, undiscounted reserves for SAP).

---

## Currency Handling

| Scenario | Resolution |
|---|---|
| Single currency (USD) | No transformation |
| Multi-currency program | Store original currency + exchange rate at transaction date. Convert to functional currency (USD) using rate. Store both amounts. |
| Exchange rate updates | Revalue outstanding reserves at current rates. Book FX gain/loss. |

---

## Source System Mapping

Each source system has a mapping config:

```csv
"SourceSystem","SourceField","CanonicalField","TransformRule","DefaultValue"
"PolicyAdmin","pol_num","policy_number","TRIM(UPPER())","" 
"PolicyAdmin","eff_dt","effective_date","PARSE_DATE(MM/DD/YYYY)",""
"ClaimsSystem","clm_amt","amount","ABS()","0"
"ClaimsSystem","clm_typ","loss_type","MAP(IND→Indemnity,ALC→ALAE,DCC→ULAE)",""
"Bordereau","written_prem","premium_written","","0"
```

---

## Schema Evolution

| Change Type | Handling |
|---|---|
| New field added | Add to canonical schema with DEFAULT. Backfill if possible. Old records get default. |
| Field renamed | Add mapping in source_system_mapping. Old name → new name. Both work during transition. |
| Field deprecated | Mark as deprecated. Stop writing. Keep in schema for historical queries. |
| Field type changed | Version the schema. Transform on read for old records. |
| New source system added | Add mapping config. No code changes. |

Schema version tracked in `data_model_version.csv`.

---

## RDK Prototype Approach

For the Excel/VBA prototype:

| Layer | RDK Implementation |
|---|---|
| Bronze | CSV import to a "Raw Data" tab. Source-faithful columns. |
| DQ / Quarantine | Quarantine tab + DQ Log tab. Rules from dq_rules_config.csv. |
| Resolution Engine | VBA module (Ins_DataResolution.bas or kernel extension). Config-driven rules. |
| Silver | "Normalized Data" tab. Canonical schema. Assumption log tab. |
| Gold | Queryable via existing formula tabs, QuarterlyAgg, Triangles. |
| Preview | "Staged Data" tab. User approves before commit. |
| Snapshots | KernelWorkspace / KernelSnapshot (already built). |
| Run History | Run Metadata tab (already built). Extend with pipeline-specific fields. |

---

## Collaboration Architecture (God Rule #3: Never Work Alone)

### Annotation & Threading

Any value at any granularity can be tagged with comments and threaded conversations:

| Granularity | Example |
|---|---|
| Cell | "This premium looks wrong — should be $1.2M not $12M" |
| Row | "This entire claim record is duplicated — see claim #4523" |
| Column | "All Q3 reserves look understated by ~10%" |
| Block (rows × columns) | "These 50 policies were all reclassed from GL to Property — correct?" |
| Dataset | "The September bordereau has a systematic decimal error in the paid column" |

**RDK Prototype:** comment_log.csv in data/ directory:
```csv
"CommentID","Layer","EntityType","EntityID","FieldName","CommentText","Author","CreatedAt","ThreadParent","Status","ResolvedBy","ResolvedAt"
```

**AIOS Implementation:** Web-based threading with @mentions, resolution tracking, and notification.

### Arbiter Roles

When Partner A and Partner B disagree on an assumption:

| Layer | Arbiter | Rationale |
|---|---|---|
| Bronze (raw data) | Original data uploader | They know the source system and intended values |
| Silver (logic/normalization) | Business owner at Partner A | They understand the business rules and are the responsible party |
| Gold (reporting/aggregates) | Actuarial or finance team | They own the reported numbers |

Disputed assumptions are tracked separately with status = "Disputed" until the arbiter resolves.

### Materiality Tracking

Every metric has a configurable materiality threshold:

```csv
"MetricName","MaterialityThreshold","DefaultThreshold","Description"
"GWP","0.001","0.001","0.1% tolerance on gross written premium"
"NetReserves","0.01","0.001","1% tolerance on net reserves"
"RBCRatio","0.005","0.001","0.5% tolerance on RBC ratio"
```

Default threshold applies to any metric without a specific override.

The system computes **assumption impact** at all times:
- For each key metric: `|Value with assumptions - Value with actuals only| / |Value with actuals only|`
- If impact < threshold → green (within materiality)
- If impact ≥ threshold → yellow/red (material, needs investigation)

### Convergence Tracking

During a close cycle, the system tracks how quickly data converges from assumed to actual:

```
Day 1:  45% actual, 55% assumed (raw data just arrived)
Day 3:  65% actual, 35% assumed (group 1 fixes applied)
Day 5:  85% actual, 15% assumed (group 2 fixes applied)
Day 7:  95% actual, 5% assumed (group 3 fixes applied)
Day 10: 100% actual, 0% assumed (close finalized)
```

This is tracked per close cycle and displayed as a convergence dashboard. Velocity metrics: "at this rate, we'll reach 100% actual by day X."

### Preview Scope

When previewing a change before commit:

- **v1:** Show impact on the **immediate next layer** only (Bronze change → Silver impact, Silver change → Gold impact)
- **Future:** "Show full cascade" toggle that traces impact through all downstream layers to final reporting outputs

Preview displays: count of affected records, delta on key aggregates, before/after comparison.

### Collaborative Data Convergence — Full Pattern

```
1.  Partner A sends raw data
2.  Partner B uploads into Bronze layer
3.  System runs DQ checks → quarantine report generated
4.  Partner B reviews quarantine:
    - Makes automated assumptions (God Rule #1)
    - Makes manual overrides where needed (God Rule #2)
    - Tags issues with comments at appropriate granularity
    - Groups issues into patterns (e.g., 3 groups)
5.  Partner B runs pipeline → downstream outputs produced with assumptions
6.  Partner B sends Partner A questions about the data (threaded comments)
7.  Partner B continues downstream work using assumed values (God Rule #1)
8.  Partner A provides group 1 + 2 fixes
9.  Partner B previews changes → sees immediate next-layer impact
10. Partner B confirms with Partner A that impacts are expected
11. Partner B commits → full pipeline reruns → all values tagged actual vs assumed
    System shows: "assumptions have 0.08% impact on key metrics (within 0.1% threshold)"
12. Partner A provides group 3 fixes → same preview-confirm-commit cycle
    System shows: "0% assumption impact — full convergence achieved"
```

---

## Locked Design Decisions

| # | Decision | Details |
|---|---|---|
| 1 | Prototype in Excel/VBA (RDK), then port to AIOS (PostgreSQL) | RDK is proof-of-concept; business requirements derived from prototype |
| 2 | Full architecture design first, then build Bronze as first implementation phase | Need the full picture to avoid painting into a corner |
| 3 | Annotation via comment_log.csv in RDK prototype | Web-based threading in AIOS |
| 4 | Arbiter is context-dependent | Data uploader for raw data, business owner for logic/assumptions |
| 5 | Materiality configurable per metric with fallback default | materiality_config.csv |
| 6 | Convergence velocity tracked during close cycle | % actual vs % assumed by day |
| 7 | Preview shows immediate next layer only (v1) | Full cascade as future enhancement |
| 8 | Three God Rules apply to all layers | Never Get Stuck, Never Lose Control, Never Work Alone |

---

## Open Questions (For Deep-Dive Spec)

1. What's the minimum viable Bronze → Silver pipeline for the first RDK prototype?
2. Should DQ rules be soft (warn) or hard (quarantine) by default?
3. How should the rules engine handle conflicting segmentation tags?
4. What's the restatement workflow — automatic or manual approval?
5. Should the Gold layer support ad-hoc queries (pivot tables) or pre-defined views only?
6. How does this integrate with the existing InsuranceDomainEngine loss development pipeline?
7. What's the data volume expectation for the prototype? (100 rows? 10,000? 100,000?)
8. How many source systems will the prototype need to handle simultaneously?
9. What close cycle cadence should the prototype support (monthly, quarterly, both)?
10. Should the convergence dashboard be a dedicated tab or a section on Dashboard?
