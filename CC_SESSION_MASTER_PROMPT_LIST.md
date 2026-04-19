# Current CC Session — All Prompts (In Order)

This session carries 7 work packages. Paste each as an addendum to the active CC session.

---

## PROMPT 1: CLEANUP & CONFIG WIRING (7 items)
**File:** `CLEANUP_WIRING_BUILD_PROMPT.md`
**What it does:**
1. KernelSnapshot split → +KernelSnapshotIO (62.9→<40KB)
2. Deprecated tab reference cleanup (Summary/Exhibits/Charts/CumulativeView/Analysis)
3. Wire workspace_config.csv (4 settings into KernelWorkspace)
4. Wire msgbox_config.csv (12 entries → ShowConfigMsgBox helper)
5. Wire display_aliases.csv (QuarterlyAgg + Triangles use friendly metric names)
6. Regression tab capture in workspace saves (SC-09)
7. Compare Workspaces button stub

---

## PROMPT 2: RBC CAPITAL MODEL
**File:** `RBC_CAPITAL_MODEL_SPEC.md`
**What it does:**
- New formula tab: RBC Capital Model (~50 rows in formula_tab_config.csv)
- Risk charge inputs (R0-R5 factors), quarterly ACL computation, TAC from Balance Sheet
- RBC Ratio with 4-tier conditional formatting (green/yellow/orange/red)
- Capital Surplus/Deficit vs target ratio
- Named ranges for RBC_Q_Ratio, RBC_Q_ACL, RBC_Q_TAC, RBC_Q_Surplus
- Tab registry, regression_config, health_config entries
- No new VBA — pure config

**CC kick-off line:** "Read CLAUDE.md, then SESSION_NOTES.md, then RBC_CAPITAL_MODEL_SPEC.md. Build the RBC Capital Model tab."

---

## PROMPT 3: REPO UPDATES (docs + roadmap + God Rules)
**File:** `REPO_UPDATE_PROMPT.md`
**What it does:**
- Add `docs/Insurance_Data_Model_Architecture_v0.1.md` to repo
- Replace roadmap with `docs/RDK_Phase_Roadmap_v5.0.md`
- Add Three God Rules to Developer Flywheel
- Update CLAUDE.md references
- Add missing display alias entries for QuarterlyAgg/Triangles metrics

**Supporting files to include:**
- `docs/Insurance_Data_Model_Architecture_v0.1.md` (31KB)
- `docs/RDK_Phase_Roadmap_v5.0.md` (8.7KB)

---

## PROMPT 4: SAMPLE CONFIG UPDATE
**CC kick-off line (inline — no separate file):**

"Read CLAUDE.md, then SESSION_NOTES.md. Update config_sample to showcase current kernel capabilities. Add missing 13 config CSVs with correct schema headers. Update tab_registry to remove deprecated tabs (Summary, Exhibits, Charts, CumulativeView, Analysis) and add infrastructure tabs (Cover Page, User Guide, Run Metadata). Verify formula_tab_config FinancialSummary references still work. Sync config_blank with all 31+ CSV headers. Do NOT modify engine/*.bas or config_insurance/."

---

## PROMPT 5: EXPENSE EXPANSION + SEED MANAGER
**File:** `EXPENSE_EXPANSION_BUILD_PROMPT.md`
**What it does:**

**Expense expansion:**
- Staffing Expense: +Bonus section (7 dept inputs), +Equity comp section (7 dept inputs), updated total formulas
- Other Expense Detail: restructured from 8 to 22 input lines across 8 categories
- Investments: +Management fee (25 bps), +Net investment income line
- Revenue Summary: reference net investment income
- Cash Flow Statement: +equity comp non-cash add-back
- Seeded defaults for 12-person startup (~$3.9M/year opex)

**Seed Manager:**
- New module: KernelSeedManager.bas (Save/Load/Delete/List/CheckCompleteness)
- seeds/ directory with manifest.json + completeness schema
- 4-state completeness check on load
- "Run & Save" flow for seeds missing detail
- Dashboard button (user-visible)
- Two starter seeds: Base_Case and Clean_Slate

---

## PROMPT 6: REVENUE EXPANSION
**File:** `REVENUE_EXPANSION_BUILD_PROMPT.md`
**What it does:**
- UW Exec Summary: +UWEX_PROG_COUNT (active program count formula + named range)
- Other Revenue Detail: full restructure from generic SW1-SW5/FEE/CON to purpose-built lines
  - Technology Revenue: 7 lines (Platform, API, Analytics, Implementation, License, Custom, Other)
  - Fee Income: 6 lines (Carrier Access % of GWP, Oversight, Onboarding, Policy Fees, Admin, Other)
  - Consulting Revenue: 4 lines (Actuarial, Risk, Regulatory, Other)
  - Model parameter: Average Premium per Policy
- Revenue Summary: label rename ("Software" → "Technology")
- Formula-driven scaling (% of GWP, per-program, per-policy)
- Seeded defaults (~$139K/quarter other revenue with 3 programs)

---

## PROMPT 7: ASSUMPTIONS REGISTER
**File:** `ASSUMPTIONS_BUILD_PROMPT.md`
**What it does:**
- New config: assumptions_config.csv (21 seed entries with confidence, sensitivity, history)
- New module: KernelAssumptions.bas (GenerateAssumptionsRegister, ShowAssumptionManager, Add/Edit/Archive/ReviewStale)
- Assumptions Register tab: visible, grouped by category, hyperlinked to input cells, color-coded confidence/sensitivity
- Manager panel: View / Add / Edit / Archive / Review Stale via MsgBox/InputBox
- Dashboard button: "Manage Assumptions" (user-visible)
- Auto-regenerates after Run Model

---

## TOTAL SCOPE

| Category | Items |
|---|---|
| New VBA modules | 3 (KernelSnapshotIO, KernelSeedManager, KernelAssumptions) |
| New config CSVs | 1 (assumptions_config) + 13 in config_sample |
| Modified VBA modules | ~10 (KernelSnapshot, KernelWorkspace, KernelFormHelpers, KernelEngine, KernelConfig, Ins_QuarterlyAgg, Ins_Triangles, KernelTabs, KernelBootstrap, KernelFormulaWriter) |
| Modified configs | formula_tab_config (~100 new rows), tab_registry (+3 tabs), button_config (+3 buttons), named_range_registry (+5 ranges), regression_config (+1 tab) |
| New tabs | 2 (RBC Capital Model, Assumptions Register) |
| Seeded data | 21 assumptions, expense defaults, revenue defaults, 2 seed scenarios |
| Validation gates | 42 |

**Estimated CC turns:** This is a large session. If CC hits the 200-turn limit, prioritize in this order:
1. CLEANUP (structural — unlocks everything else)
2. Expense + Revenue expansion (investor-facing)
3. RBC (investor-facing)
4. Assumptions Register (trust-building artifact)
5. Seed Manager (demo convenience)
6. Repo updates (docs, can be done manually)
7. Sample config (lowest priority)

**CC command:**
```
claude --dangerously-skip-permissions --model claude-opus-4-6 --max-turns 200
```

---

## CRITICAL INSTRUCTION: DO NOT REVERT EXISTING WORK

Some of the items in these prompts may have already been partially or fully implemented in a prior CC session. Before making changes to any file:

1. **Read the current state first.** Check if the feature, config row, module, or tab already exists.
2. **If it already exists and works**, do not overwrite it — even if the implementation differs slightly from what this prompt describes. The existing implementation may have been intentionally adjusted.
3. **If it partially exists**, extend or complete it rather than replacing it. Merge the prompt's requirements with what's already there.
4. **If it doesn't exist**, build it per the prompt.
5. **Never delete or revert** VBA code, config rows, tabs, or functionality that already works unless explicitly instructed to do so.
6. **When in doubt, preserve the existing implementation** and add what's missing rather than rebuilding from scratch.

This applies to all 7 prompts. The goal is additive — fill gaps and complete unfinished work, not replace work that's already done.
