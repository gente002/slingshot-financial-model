# RDK Repo Update — Data Model Spec, Roadmap, God Rules

Read `CLAUDE.md`, then `SESSION_NOTES.md`. This is an addendum to the current session.

## 1. Add Insurance Data Model Architecture spec

Copy the file `docs/Insurance_Data_Model_Architecture_v0.1.md` into the docs/ directory. This is a new design document — do not modify it, just include it in the repo.

The full content is provided as an attached file. If you cannot read the attached file, create `docs/Insurance_Data_Model_Architecture_v0.1.md` with a placeholder noting: "Full spec to be added — see Insurance_Data_Model_Architecture_v0.1.md deliverable."

## 2. Replace roadmap

Replace `docs/RDK_Phase_Roadmap_v2.0.md` with `docs/RDK_Phase_Roadmap_v5.0.md`. The v5.0 roadmap reflects:
- All completed phases through INFRA
- Active CLEANUP sprint with RBC
- Execution queue: BCAR → WS-COMPARE → DATA-DESIGN → DATA-BRONZE → DATA-SILVER → DATA-GOLD → DATA-PIPELINE
- Three God Rules at the top
- Module Size Remediation table
- KE-01 simplified (Prove-It + Exhibits evolution only)
- KE-02 removed
- 12i deferred indefinitely
- Full decision log (16 entries)
- Data architecture phases (DATA-DESIGN through DATA-PIPELINE)

## 3. Add Three God Rules to Developer Flywheel

In `docs/RDK_Developer_Flywheel_v1.2.md`, add a new section near the top (after the existing preamble/introduction, before the first numbered section):

```markdown
## The Three God Rules

These are foundational design rules that apply to every system, every module, and every workflow in the RDK. They are not guidelines — they are constraints. Every design decision must satisfy all three.

### God Rule #1: Never Get Stuck

Use the information you have and make the best assumption available to get to the downstream structure and shape so we can report the same every time.

**Design implications:**
- Always produce output, even with incomplete data
- Always document what was assumed vs what was known
- Every assumption is reversible — when real data arrives, it replaces the assumption
- The downstream shape is fixed — upstream variability is resolved before it reaches the output
- Quality issues quarantine rows, they don't halt the pipeline
- Applies to: data ingestion, model computation, formula tabs, reporting, configuration loading

### God Rule #2: Never Lose Control

Users always have a sanctioned path to manually override any value at any layer — temporarily, with tracking, without breaking the system.

**Design implications:**
- Any value can be overridden: raw data, computed output, system-generated assumptions, configuration-driven defaults
- Overrides are tracked: who, when, why, what was the original value, what was changed
- Overrides persist until explicitly replaced by corrected source data or removed by the user — they never silently expire
- Overrides that persist beyond a configurable threshold (e.g., 90 days) are flagged for investigation, not auto-removed
- The system doesn't distinguish between source data, automated assumptions, and manual overrides at the output layer — all produce the same shape
- Override granularity: individual record overrides, bulk overrides (apply one correction to many records), and aggregate overrides (top-side adjustment that allocates down to detail)
- Override hygiene: recurring overrides should be converted to permanent rules or configuration changes
- Applies to: formula tabs (blue input cells), data pipelines (adjustment log), assumptions (user-entered vs system-generated), raw data corrections, any workflow with hard deadlines

### God Rule #3: Never Work Alone

Preview, annotate, converge, confirm. Every data workflow is a multi-party collaboration with real-time feedback, threaded discussion, and progressive convergence toward truth.

**Design implications:**
- Any value at any granularity (cell, row, column, block of rows/columns, or entire dataset) can be tagged with comments and conversation threads
- Changes can be previewed before committing — including the downstream impact on all consuming systems
- Partners exchange feedback in real-time without waiting for full pipeline reruns
- Data quality issues are grouped into patterns, discussed collaboratively, and resolved incrementally
- Every value is tagged as "actual" (from source) or "assumption" (automated or manual) — the materiality of assumptions on key metrics is measurable at all times
- Materiality thresholds are configurable per metric (e.g., 0.1% tolerance on premium, 1% on reserves)
- The system tracks convergence: what percentage of values are actual vs assumed, and how that improves over the close cycle
- Applies to: data ingestion workflows, close cycles, partner data exchanges, assumption resolution, any process where two or more parties need to align on data
```

## 4. Update CLAUDE.md

Add to the project description or principles section:
- Reference the Three God Rules (Never Get Stuck, Never Lose Control, Never Work Alone)
- Note that docs/Insurance_Data_Model_Architecture_v0.1.md exists as the data architecture design
- Note that docs/RDK_Phase_Roadmap_v5.0.md is the current roadmap

## 5. Append to SESSION_NOTES.md

Document:
- Three God Rules added to Developer Flywheel
- Insurance Data Model Architecture v0.1 added to docs/
- Roadmap updated to v5.0 (added BCAR, DATA phases, God Rules, decision log)
