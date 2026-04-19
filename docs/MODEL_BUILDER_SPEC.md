# RDK Model Builder — Specification

## Purpose

A standardized, templated process for configuring a new financial model on the RDK kernel. The user answers business questions; the system generates a complete config directory ready for Setup.bat → Bootstrap → Run Model.

## Design Principles

1. **Progressive disclosure.** Three tiers. Tier 1 (2 min) produces a bootable skeleton. Tier 2 (10 min) adds structure. Tier 3 (15 min) adds polish. The user can stop at any tier and have a working model.

2. **Business language, not technical.** The questionnaire asks "What line items does your P&L have?" not "Define formula_tab_config rows with CellType=Formula and Content={ROWID:...} placeholders."

3. **Config as output.** The process generates CSV files, not VBA code. The user never touches kernel modules. If the model needs custom computation (actuarial math, Monte Carlo sampling), that's a separate domain module — the builder generates the config that WRAPS the computation.

4. **Idempotent.** Running the builder twice with the same answers produces the same config. Updating answers regenerates config without manual cleanup.

## Tier 1: Model Identity (2 minutes)

### Questions

| # | Question | Type | Default | Maps to |
|---|----------|------|---------|---------|
| 1 | What is your company name? | Text | "Acme Corp" | branding_config.CompanyName |
| 2 | What is the model title? | Text | "Financial Model" | branding_config.ModelTitle |
| 3 | One-line tagline (optional) | Text | "" | branding_config.Tagline |
| 4 | What are you modeling? | Choice: Financial Model, Monte Carlo, Actuarial, ETL, Custom | Financial Model | Template selection |
| 5 | What is your time horizon? | Number + unit: months/quarters/years | 60 months | granularity_config.TimeHorizon |
| 6 | What do you call your entities? | Text | "Products" | display_aliases.Entity |
| 7 | How many entities (max)? | Number 1-50 | 10 | granularity_config.MaxEntities |
| 8 | What is the scenario name? | Text | "Base" | input_schema ScenarioName default |

### Output: Bootable skeleton

- `granularity_config.csv` (TimeHorizon, MaxEntities, DomainModule)
- `tab_registry.csv` (Cover Page, Dashboard, Assumptions, Detail, Error Log, Config + one Input tab + one Output tab)
- `input_schema.csv` (ScenarioName only)
- `column_registry.csv` (Entity, Period, CalYear, CalQuarter — dimension columns only)
- `branding_config.csv` (CompanyName, ModelTitle, Tagline)
- `pipeline_config.csv` (all steps enabled)
- `formula_tab_config.csv` (empty — no formulas yet)
- `named_range_registry.csv` (empty)
- `validation_config.csv` (empty)
- `health_config.csv` (empty)
- Stub `DomainEngine.bas` (copy from template)

**Result:** User can run Setup.bat and bootstrap. Gets a workbook with Cover Page, Dashboard, one empty Input tab, one empty Output tab, and Detail. Run Model produces empty Detail (no computation yet).

---

## Tier 2: Model Structure (10 minutes)

### Questions

**2.1 Input Tabs**

"List the tabs where users enter data. For each, give a name and describe what's entered."

| # | Tab Name | Description | Example entries |
|---|----------|-------------|-----------------|
| 1 | _______ | _______ | e.g., "Revenue by product by quarter" |
| 2 | _______ | _______ | |
| ... | | | |

Maps to: `tab_registry.csv` rows (Type=Domain, Category=Input, QuarterlyColumns based on "by quarter")

**2.2 Output Tabs**

"List the tabs that show computed results. For each, give a name and what it displays."

| # | Tab Name | Description | Has quarterly columns? | Has grand total? |
|---|----------|-------------|----------------------|-------------------|
| 1 | _______ | _______ | Yes/No | Yes/No |
| 2 | _______ | _______ | Yes/No | Yes/No |

Maps to: `tab_registry.csv` rows (Type=Domain, Category=Output)

**2.3 Metrics**

"What numbers does your model compute? For each, is it a FLOW (activity during a period, like revenue) or a BALANCE (level at a point in time, like cash balance)?"

| # | Metric Name | Flow or Balance | Unit | Formula (plain English) |
|---|------------|----------------|------|------------------------|
| 1 | Revenue | Flow | Currency | (entered by user) |
| 2 | COGS | Flow | Currency | (entered by user) |
| 3 | Gross Profit | Flow | Currency | Revenue - COGS |
| 4 | Cash Balance | Balance | Currency | Previous cash + net income |
| ... | | | | |

Maps to:
- `column_registry.csv` (Name, FieldClass=Incremental or Derived, DerivationRule, BalanceType)
- `formula_tab_config.csv` (rows on output tabs)

**2.4 Global Assumptions**

"What global parameters affect the entire model?"

| # | Parameter | Type | Default | Range |
|---|-----------|------|---------|-------|
| 1 | Tax Rate | Percentage | 21% | 0-50% |
| 2 | Discount Rate | Percentage | 10% | 0-30% |
| ... | | | | |

Maps to:
- `input_schema.csv` (Section="Global Assumptions")
- `validation_config.csv` (range constraints)

**2.5 Tab Layout**

For each output tab from 2.2, "What sections and line items does this tab show?"

```
Tab: Income Statement
  Section: Revenue
    - Total Revenue (formula: sum of revenue sources)
    - Cost of Revenue
  Section: Gross Profit
    - Gross Profit (formula: Revenue - Cost of Revenue)
    - Gross Margin % (formula: Gross Profit / Revenue)
  Section: Operating Expenses
    - Staffing
    - Other Expenses
  Section: Net Income
    - Operating Income (formula: Gross Profit - OpEx)
    - Tax (formula: Operating Income * Tax Rate)
    - Net Income (formula: Operating Income - Tax)
```

Maps to: `formula_tab_config.csv` (rows with Section/Label/Formula/Spacer cell types, row numbers, formatting)

### Output: Structured model

All Tier 1 files updated plus:
- Full `tab_registry.csv` with all input and output tabs
- `column_registry.csv` with all metrics
- `formula_tab_config.csv` with tab layouts (sections, labels, formulas)
- `input_schema.csv` with global assumptions
- `validation_config.csv` with range constraints from 2.4

**Result:** Bootable model with real tabs, real formulas, real inputs. Run Model computes derived fields. Output tabs show formatted financial statements.

---

## Tier 3: Relationships & Polish (15 minutes)

### Questions

**3.1 Cross-Tab References**

"Which output tabs reference data from other tabs?"

| Source Tab | Source Metric | Target Tab | Target Line Item |
|-----------|--------------|-----------|-----------------|
| Revenue Detail | Total Revenue | Income Statement | Total Revenue |
| Expense Detail | Total Expenses | Income Statement | Operating Expenses |

Maps to: `{REF:SourceTab!RowID}` in formula_tab_config, named ranges in `named_range_registry.csv`

**3.2 Balance Checks**

"What invariants should always be true? (These show green when correct, red when broken.)"

| Tab | What should equal zero? |
|-----|------------------------|
| Balance Sheet | Total Assets - Total Liabilities & Equity |
| Cash Flow | Ending Cash - Balance Sheet Cash |

Maps to: `health_config.csv`

**3.3 Input Validation**

"What limits should inputs have?"

| Tab | Input | Min | Max | Required? |
|-----|-------|-----|-----|-----------|
| Assumptions | Tax Rate | 0% | 50% | Yes |
| Revenue Detail | Revenue amounts | 0 | - | No |

Maps to: `validation_config.csv`

**3.4 Prove-It Checks**

"What audit checks verify the model is working correctly?"

| Check | What it verifies |
|-------|-----------------|
| Revenue reconciliation | Sum of Detail revenue = Income Statement revenue |
| Balance identity | Assets = Liabilities + Equity |

Maps to: `prove_it_config.csv`

**3.5 Extensions**

"Does the model need any of these capabilities?"

- [ ] PDF report generation
- [ ] Development curve library (actuarial)
- [ ] Custom Dashboard buttons
- [ ] Named range references for external tools

Maps to: `extension_registry.csv`

**3.6 Branding**

"Customize the investor/board presentation."

| Setting | Value |
|---------|-------|
| Post-run tab (what shows after Run Model) | _______ |
| Dashboard metrics (what shows on Dashboard) | Metric 1: _______ , Metric 2: _______ |

Maps to: `branding_config.csv`

### Output: Production-ready model

All Tier 1+2 files updated plus:
- `named_range_registry.csv` with cross-tab references
- `health_config.csv` with balance checks
- `prove_it_config.csv` with audit checks
- `extension_registry.csv` with enabled extensions
- Updated `branding_config.csv`

**Result:** Complete, investor-ready model. All validation, health checks, and cross-tab references in place.

---

## Implementation Options

### Option A: Document-Based (Now)

The Model Builder is a **markdown template** (`MODEL_BUILDER_QUESTIONNAIRE.md`) that the user fills out. A developer (or Claude) reads the answers and generates config CSVs.

**Pros:** Zero code needed. Works today. Flexible for unusual models.
**Cons:** Manual translation from answers to config. Error-prone.

### Option B: Script-Based (Phase 15)

A **PowerShell/Python script** (`scripts/ModelBuilder.ps1`) that prompts the user interactively and generates config CSVs automatically.

**Pros:** Automated, repeatable, validates inputs.
**Cons:** Requires scripting. Less flexible for edge cases.

### Option C: Excel-Based (Phase 15)

A **Model Builder workbook** with input forms (VBA UserForms or data validation dropdowns). The user fills in the workbook, clicks "Generate Config", and the VBA writes CSV files.

**Pros:** Stays in Excel ecosystem. Visual. Validations built-in.
**Cons:** Complex to build. Another workbook to maintain.

### Option D: AI-Assisted (Future)

The user describes their model in natural language. Claude reads the description, asks clarifying questions, and generates config CSVs. The questionnaire becomes a conversation.

**Pros:** Fastest for the user. Handles ambiguity. Can suggest best practices.
**Cons:** Requires AI tooling. Non-deterministic.

### Recommendation

**Start with Option A (document template) now.** It's free, works today, and validates the questionnaire design. Upgrade to Option B or C when building the second model reveals friction points. Option D is the long-term vision.

---

## What's Excluded

1. **Domain module code.** The builder generates CONFIG, not VBA. If the model needs custom computation (actuarial curves, Monte Carlo sampling), the developer writes a DomainEngine module separately. The builder generates the config that wraps it.

2. **Data migration.** The builder creates empty input tabs. Populating them with actual data is a separate step.

3. **UW Inputs tab layout.** For the insurance FM, UW Inputs has a complex multi-section layout (program definitions, premium schedule, commission rates). This is domain-specific and defined in formula_tab_config by the domain author, not auto-generated.

4. **Extension development.** The builder can enable existing extensions (CurveLib, ReportGen) but doesn't create new ones.

---

## Questions the User Should Be Asking

1. **Can the builder UPDATE an existing model?** Yes — re-running with updated answers should regenerate config without losing user-entered data. The builder operates on config CSVs, not on the workbook.

2. **What about config versioning?** Each run of the builder should produce a dated config directory (e.g., `config_mymodel_20260403/`). The user can diff two versions to see what changed.

3. **What about testing?** Tier 2 output should include basic prove_it_config checks (e.g., if you defined Gross Profit = Revenue - COGS, auto-generate a reconciliation check). The builder should be opinionated about testing.

4. **What about the domain module?** The builder should generate a STUB domain module (`MyModelDomainEngine.bas`) with the 4 required functions pre-scaffolded. The developer fills in the computation logic. For models without custom computation (pure formula-driven), the stub's Execute() can be empty — the kernel handles everything via formula_tab_config.

5. **Can a model be 100% config with no domain module?** Almost. If all computation is expressible as Excel formulas (no VBA loops, no PRNG, no external data), the model can use the generic `DomainEngine.bas` stub with empty functions. The formula_tab_config does all the work. This is the ideal for simple models.

---

## Decisions Needed

None right now. The spec is self-contained. Implementation starts when building the second model (Phase 15 in roadmap).
