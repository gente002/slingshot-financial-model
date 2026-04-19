# Expanded Expense Assumptions — Spec & CC Build Prompt

**Date:** April 2026
**Depends on:** CLEANUP sprint complete
**Scope:** Expand Staffing Expense and Other Expense Detail tabs from 14 to ~28 expense lines. Seed with startup carrier defaults. Wire into Expense Summary.

## Locked Decisions

| ID | Decision |
|---|---|
| EXP-01 | ~25-30 expense lines, grouped categories |
| EXP-02 | Department-specific per-head scaling for key items |
| EXP-04 | Remove surplus lines tax (insured pays) |
| EXP-05 | LAE stays in loss development (ELR covers it) |
| EXP-06 | Investment management fees shown separately on Investments tab |
| EXP-07 | One-time costs expensed as incurred |
| EXP-08 | Bonus accrued quarterly (annual ÷ 4). Equity comp accrued quarterly, hits IS not CFS. |
| EXP-09 | Seed with startup carrier defaults |

---

## CC Build Prompt

Read `CLAUDE.md`, then `SESSION_NOTES.md`. This is an addendum to the current session.

### Overview

Expand the expense assumptions across Staffing Expense and Other Expense Detail. Add investment management fees to Investments. Wire everything into Expense Summary. Seed with realistic startup carrier defaults.

### STAFFING EXPENSE TAB — Changes

**Add two new sections after Section 3 (Benefits):**

**Section 3B: Bonus / Incentive Compensation**

| RowID | Label | CellType | Default | Notes |
|---|---|---|---|---|
| STF_SEC_BONUS | Bonus & Incentive (% of Salary) | Label | | Section header |
| STF_BONUS_UW | Underwriting | Input | 0.15 | 15% of salary |
| STF_BONUS_CLAIMS | Claims | Input | 0.10 | |
| STF_BONUS_ACTUARY | Actuarial | Input | 0.15 | |
| STF_BONUS_FINANCE | Finance | Input | 0.15 | |
| STF_BONUS_TECH | Technology | Input | 0.15 | |
| STF_BONUS_EXEC | Executive | Input | 0.30 | 30% — executive bonus is higher |
| STF_BONUS_OTHER | Other | Input | 0.10 | |

**Section 3C: Equity / Stock Compensation**

| RowID | Label | CellType | Default | Notes |
|---|---|---|---|---|
| STF_SEC_EQUITY | Equity Compensation (Annual $ per Person) | Label | | Section header |
| STF_EQ_UW | Underwriting | Input | 0 | |
| STF_EQ_CLAIMS | Claims | Input | 0 | |
| STF_EQ_ACTUARY | Actuarial | Input | 0 | |
| STF_EQ_FINANCE | Finance | Input | 0 | |
| STF_EQ_TECH | Technology | Input | 15000 | Tech gets equity |
| STF_EQ_EXEC | Executive | Input | 50000 | Exec gets most equity |
| STF_EQ_OTHER | Other | Input | 0 | |

**Update Section 4 formula for each department:**

Current: `HC × Salary / 4 × (1 + Benefits)`
New: `HC × (Salary × (1 + Benefits + Bonus) + EquityComp) / 4`

Example for UW:
```
={ROWID:STF_HC_UW} * ({ROWID:STF_SAL_UW} * (1 + {ROWID:STF_BENEFITS} + {ROWID:STF_BONUS_UW}) + {ROWID:STF_EQ_UW}) / 4
```

**Add a new Section 5 showing the equity comp subtotal (for CFS exclusion):**

| RowID | Label | Formula | Notes |
|---|---|---|---|
| STF_SEC_EQTOT | Equity Compensation Subtotal (Non-Cash) | Label | |
| STF_EQ_TOTAL | Total Equity Compensation | Formula | Sum of HC × EQ for all departments ÷ 4 |

This total is needed so the Cash Flow Statement can add back equity comp as a non-cash expense.

**Seed headcount and salary defaults (Y1):**

| Department | HC | Salary |
|---|---|---|
| Underwriting | 2 | 120000 |
| Claims | 1 | 95000 |
| Actuarial | 2 | 150000 |
| Finance | 1 | 110000 |
| Technology | 2 | 140000 |
| Executive | 3 | 200000 |
| Other | 1 | 85000 |

Benefits factor: 0.30 (30%)

These represent a 12-person startup team. Total loaded compensation ≈ $2.2M/year.

---

### OTHER EXPENSE DETAIL TAB — Restructure

Replace the current 8-line structure with ~20 grouped categories. Keep the Personnel / Non-Personnel split.

**Personnel Expenses (expanded):**

| RowID | Label | CellType | Scale Type | Default (Annual) | Notes |
|---|---|---|---|---|---|
| OED_PER_CONTRACT | Contractors & Consultants | Input | Fixed | 150000 | Project-based, quarterly ÷ 4 |
| OED_PER_RECRUIT | Recruiting & Talent | Input | Fixed | 75000 | ~25% of first-year salary per hire, heavy Y1 |
| OED_PER_TRAINING | Training & Development | Input | Per Head (total) | 3000 | Per person per year. Formula: `={REF:Staffing Expense!STF_HC_TOTAL} * $C$X / 4` |
| OED_PER_TOTAL | Total Personnel | Formula | | | Sum of above |

**Technology & Infrastructure:**

| RowID | Label | CellType | Scale Type | Default (Annual) | Notes |
|---|---|---|---|---|---|
| OED_NP_CLOUD | Cloud & Infrastructure | Input | Fixed | 60000 | AWS/Azure baseline |
| OED_NP_SAAS | SaaS & Productivity | Input | Per Head (total) | 5000 | Per person/year. Office 365, Slack, Zoom, etc. Formula: `={REF:Staffing Expense!STF_HC_TOTAL} * $C$X / 4` |
| OED_NP_INSTECH | Insurance Platforms (PAS/Claims/Billing) | Input | Per Dollar (GWP) | 0 | For now 0 — add when vendor costs known. Could be % of GWP or flat. |
| OED_NP_CYBER | Cybersecurity | Input | Fixed | 25000 | Endpoint, SIEM, pen testing |
| OED_NP_HARDWARE | Hardware & Equipment | Input | Per Head (total) | 2500 | Per new hire/year. Formula: `={REF:Staffing Expense!STF_HC_TOTAL} * $C$X / 4` |
| OED_NP_TECH_TOTAL | Total Technology | Formula | | | Sum of above |

**Facilities:**

| RowID | Label | CellType | Scale Type | Default (Annual) | Notes |
|---|---|---|---|---|---|
| OED_NP_RENT | Rent & Facilities | Input | Fixed | 60000 | Coworking or small office |
| OED_NP_FACILITIES_TOTAL | Total Facilities | Formula | | | Just rent for now, extensible |

**Professional Services:**

| RowID | Label | CellType | Scale Type | Default (Annual) | Notes |
|---|---|---|---|---|---|
| OED_NP_AUDIT | External Audit | Input | Fixed | 150000 | Statutory audit |
| OED_NP_ACTUARIAL | Actuarial Consulting | Input | Fixed | 100000 | Reserve opinions, rate filings |
| OED_NP_LEGAL | Legal | Input | Fixed | 200000 | Corporate, regulatory, compliance. Heavy Y1. |
| OED_NP_TAX | Tax & Accounting | Input | Fixed | 50000 | |
| OED_NP_PROF_TOTAL | Total Professional Services | Formula | | | Sum of above |

**Insurance (Company's Own Coverage):**

| RowID | Label | CellType | Scale Type | Default (Annual) | Notes |
|---|---|---|---|---|---|
| OED_NP_INSUR | D&O / E&O / Cyber Insurance | Input | Fixed | 75000 | Bundled carrier coverage |

**Travel & Business Development:**

| RowID | Label | CellType | Scale Type | Default (Annual) | Notes |
|---|---|---|---|---|---|
| OED_NP_TRAVEL | Travel & Entertainment | Input | Fixed | 60000 | MGA visits, conferences, investor meetings |
| OED_NP_MARKETING | Marketing & Branding | Input | Fixed | 25000 | Website, collateral |
| OED_NP_TRAVEL_TOTAL | Total Travel & BD | Formula | | | Sum of above |

**Regulatory & Compliance:**

| RowID | Label | CellType | Scale Type | Default (Annual) | Notes |
|---|---|---|---|---|---|
| OED_NP_REG | Regulatory & Filing Fees | Input | Fixed | 30000 | State filings, license renewals, bureau assessments |
| OED_NP_CATMODEL | Cat Modeling | Input | Fixed | 0 | Add when needed. RMS/AIR/CoreLogic. |

**Board & Governance:**

| RowID | Label | CellType | Scale Type | Default (Annual) | Notes |
|---|---|---|---|---|---|
| OED_NP_BOARD | Board Compensation & Costs | Input | Fixed | 50000 | Director fees, meeting costs |
| OED_NP_INVESTOR | Investor Relations | Input | Fixed | 10000 | Reporting, IR tools |
| OED_NP_GOV_TOTAL | Total Board & Governance | Formula | | | Sum of above |

**Startup / One-Time Costs:**

| RowID | Label | CellType | Scale Type | Default (Annual) | Notes |
|---|---|---|---|---|---|
| OED_NP_STARTUP | One-Time Startup Costs | Input | Fixed | 100000 | Regulatory approval, initial buildout, branding. Y1 only. Expensed as incurred. |

**Non-Personnel Total and Grand Total:**

| RowID | Label | Formula |
|---|---|---|
| OED_NP_TOTAL | Total Non-Personnel | Sum of all sub-totals: OED_NP_TECH_TOTAL + OED_NP_FACILITIES_TOTAL + OED_NP_PROF_TOTAL + OED_NP_INSUR + OED_NP_TRAVEL_TOTAL + OED_NP_REG + OED_NP_CATMODEL + OED_NP_GOV_TOTAL + OED_NP_STARTUP |
| OED_Q_TOTAL | Total Other Operating Expense | OED_PER_TOTAL + OED_NP_TOTAL |

**Total: 22 input lines + 7 subtotals + 1 grand total = 30 rows** (plus headers/spacers).

---

### PER-HEAD FORMULA PATTERN

For per-head costs, use this pattern:

**Scales with total headcount:**
```
={REF:Staffing Expense!STF_HC_TOTAL} * $C$[row] / 4
```
Where $C$ holds the annual per-person cost. The user enters the annual rate, the formula multiplies by headcount and divides by 4 for quarterly.

**Scales with specific department headcount (EXP-02):**
```
={REF:Staffing Expense!STF_HC_TECH} * $C$[row] / 4
```
Use department-specific HC only for: OED_NP_INSTECH (if it becomes per-UW-head). For v1, most per-head items use total HC.

**Important:** The per-head inputs in column C are ANNUAL per-person costs. The formula divides by 4 for quarterly output. This matches how Staffing Expense works (annual salary ÷ 4).

---

### INVESTMENTS TAB — Add Investment Management Fees

Add one new row after INV_Q_INC_TOTAL:

| RowID | Label | CellType | Default | Formula |
|---|---|---|---|---|
| INV_MGMT_FEE_BPS | Management Fee (bps) | Input | 0.0025 | 25 bps annually. Static input in col C. |
| INV_MGMT_FEE | Investment Management Fee | Formula | | `={REF:Investments!INV_Q_INVESTED} * $C$[fee_row] / 4` |
| INV_Q_INC_NET | Net Investment Income | Formula | | `={ROWID:INV_Q_INC_TOTAL} - {ROWID:INV_MGMT_FEE}` |

Update Revenue Summary to reference INV_Q_INC_NET instead of INV_Q_INC_TOTAL (so investment income is net of fees).

---

### EXPENSE SUMMARY TAB — No Changes Needed

The Expense Summary already references:
- `{REF:Staffing Expense!STF_Q_TOTAL}` → picks up bonus + equity comp automatically
- `{REF:Other Expense Detail!OED_Q_TOTAL}` → picks up all new categories automatically

No changes to Expense Summary formulas. The wiring is already correct because the detail tabs feed through their totals.

---

### CASH FLOW STATEMENT — Add Equity Comp Add-Back

If Cash Flow Statement has an add-back section for non-cash items, add:

| RowID | Label | Formula |
|---|---|---|
| CFS_EQCOMP | Add: Equity Compensation (non-cash) | `={REF:Staffing Expense!STF_EQ_TOTAL}` |

This ensures equity comp hits the Income Statement (expense) but is added back in Cash from Operations (non-cash). Check the current CFS structure to find the right insertion point — it should be in the operating section alongside depreciation or other non-cash adjustments.

---

### SEEDED DEFAULTS SUMMARY

With the defaults above, a 12-person startup carrier in Y1 would show:

| Category | Annual | Quarterly |
|---|---|---|
| Staffing (salary + benefits + bonus) | ~$2.4M | ~$600K |
| Equity compensation (non-cash) | ~$180K | ~$45K |
| Personnel (contractors, recruiting, training) | ~$264K | ~$66K |
| Technology | ~$183K | ~$46K |
| Facilities | ~$60K | ~$15K |
| Professional services | ~$500K | ~$125K |
| Insurance (D&O/E&O) | ~$75K | ~$19K |
| Travel & BD | ~$85K | ~$21K |
| Regulatory | ~$30K | ~$8K |
| Board & governance | ~$60K | ~$15K |
| Startup one-time | ~$100K | ~$25K |
| **Total Operating Expense** | **~$3.9M** | **~$985K** |

This is realistic for a startup E&S carrier in its first year. The model starts with these defaults so it "makes sense" immediately after Run Model.

---

### VALIDATION GATES

1. All new RowIDs resolve correctly — no #REF! errors
2. Per-head formulas multiply by HC and divide by 4
3. OED_Q_TOTAL = OED_PER_TOTAL + OED_NP_TOTAL
4. Expense Summary picks up new totals automatically (no Expense Summary changes needed)
5. IS Total Expenses reflects the expanded costs
6. BS still balances (BS_CHECK = 0)
7. CFS still reconciles (CFS_CHECK = 0)
8. CFS adds back equity comp as non-cash
9. Investment income on Revenue Summary reflects net (after management fee)
10. Default seed values produce reasonable quarterly outputs (~$985K/quarter total opex)
11. SESSION_NOTES.md appended only

---

## SEED MANAGER — Build on Snapshot Architecture

### Overview

Build a kernel-level Seed Manager that provides curated, named scenario snapshots with one-click restore. Uses existing KernelSnapshot/KernelTabIO infrastructure. Seeds are read-only reference scenarios stored in a `seeds/` folder.

### Architecture

```
seeds/
  Base_Case/
    inputs.csv
    input_tabs/
      Assumptions.csv
      UW_Inputs.csv
      Capital_Activity.csv
      Staffing_Expense.csv
      Other_Expense_Detail.csv
      Sales_Funnel.csv
      Investments.csv
      Other_Revenue_Detail.csv
    detail.csv              (pre-computed — optional, enables "no Run Model" load)
    quarterly_summary.csv   (pre-computed — optional)
    manifest.json
  Full_Portfolio/
    ...same structure...
```

### New Module: KernelSeedManager.bas

**Public API:**

```vba
Public Sub SaveSeed(seedName As String)
' 1. Run staleness check (prompt to run model if stale)
' 2. Create seeds/{seedName}/ directory
' 3. Export using KernelSnapshot helpers:
'    - KernelSnapshot.ExportInputsToFile → inputs.csv
'    - KernelTabIO.ExportAllInputTabs → input_tabs/
'    - KernelSnapshot.ExportDetailToFile → detail.csv
'    - Export QuarterlySummary if it exists
' 4. Write manifest.json with completeness schema
' 5. Log and confirm

Public Sub LoadSeed(seedName As String)
' 1. Read manifest.json — check completeness
' 2. Compare captured tabs against current tab_registry input tabs
' 3. Flag any missing tabs
' 4. Import: KernelTabIO.ImportAllInputTabs → restores all input tabs
' 5. Import: KernelSnapshot.ImportInputsFromCsv if inputs.csv exists
' 6. Check if detail.csv exists:
'    - If YES: import detail, refresh formula tabs (no Run Model needed)
'    - If NO: prompt user with three options (see below)
' 7. Update Assumptions tab scenario name to seed name
' 8. Refresh Cover Page, Dashboard metadata

Public Sub DeleteSeed(seedName As String)
' Remove seeds/{seedName}/ directory after confirmation

Public Sub ListSeeds() As String()
' Enumerate seeds/ subfolders, read manifest.json for each

Public Function CheckSeedCompleteness(seedName As String) As String
' Compare seed contents against current tab_registry
' Return: "Complete", "Missing: TabA, TabB", or "Schema mismatch"
```

### Manifest Schema (manifest.json)

```json
{
  "name": "Base Case",
  "description": "Single GL program, $1M/quarter, 3-year horizon",
  "createdAt": "2026-04-03T16:30:00",
  "kernelVersion": "1.3.0",
  "completeness": {
    "captured": ["Assumptions", "UW Inputs", "Capital Activity", "Staffing Expense",
                  "Other Expense Detail", "Sales Funnel", "Investments", "Other Revenue Detail"],
    "hasDetail": true,
    "detailRows": 2410,
    "hasQuarterlySummary": true,
    "configHash": "abc123def456"
  }
}
```

### Completeness Check on Load

When loading a seed, compare `completeness.captured` against all tabs in `tab_registry.csv` where `Category = "Input"`:

1. **All input tabs present + detail present:** Load everything, refresh formula tabs, no Run Model needed. Message: "Scenario loaded: {name}. Results restored."

2. **All input tabs present + detail MISSING:** Prompt:
   ```
   "This scenario has no saved results. Would you like to run the model now?"
   [Run & Save]  [Load Inputs Only]  [Cancel]
   ```
   - **Run & Save:** Loads inputs, runs model, exports detail.csv back to the seed folder so next load is instant.
   - **Load Inputs Only:** Loads inputs only. User can inspect before running.
   - **Cancel:** Aborts.

3. **Some input tabs missing (new tabs added since seed was created):** Warn:
   ```
   "This scenario is missing data for: {TabA}, {TabB}.
   These tabs were added after this scenario was saved.
   Missing tabs will keep their current values."
   [Load Anyway]  [Cancel]
   ```

4. **Config hash mismatch:** Warn:
   ```
   "This scenario was saved with a different model configuration.
   Results may differ from the original. Run Model recommended after loading."
   [Load Anyway]  [Cancel]
   ```

### Dashboard Button

Add to `config_insurance/button_config.csv`:
```csv
"Dashboard","SEED_MANAGER","Seed Manager","KernelFormHelpers.ShowSeedManager","FALSE","15","TRUE","7","2"
```

SortOrder 15 places it between Run Model (10) and Save/Load (20). DevOnly=FALSE — user-visible.

Add to `KernelFormHelpers.bas`:
```vba
Public Sub ShowSeedManager()
    ' List available seeds
    ' Prompt: [Load Seed] [Save Current as Seed] [Delete Seed] [Cancel]
    ' Delegate to KernelSeedManager
End Sub
```

### Seed Manager UI Flow

**Save Current as Seed:**
1. Prompt for seed name (InputBox)
2. Prompt for description (InputBox)
3. Check staleness — offer to run model first
4. Export everything to seeds/{name}/
5. Confirm: "Seed saved: {name}"

**Load Seed:**
1. List available seeds (name + description from manifest)
2. User picks one
3. Run completeness check
4. Load based on completeness state (see above)
5. Confirm: "Scenario loaded: {name}"

**Delete Seed:**
1. List available seeds
2. User picks one
3. Confirm: "Delete seed '{name}'? This cannot be undone." [Yes] [No]
4. Delete folder

### Seeds Are Read-Only After Creation

Loading a seed copies data INTO the workbook. It does not modify the seed folder — EXCEPT when the user chooses "Run & Save" on a seed missing detail.csv. In that case, the detail.csv is written back to the seed to make future loads instant.

### Create Two Starter Seeds

After building the Seed Manager, create two seeds from the current model configuration:

**Seed 1: "Base_Case"**
- Description: "Startup carrier base case with seeded expense and UW assumptions"
- Use the seeded defaults from this prompt (12-person team, ~$3.9M opex)
- Include whatever UW programs are currently configured
- Run model and save with detail

**Seed 2: "Clean_Slate"**
- Description: "Empty model — all inputs zeroed, ready for fresh configuration"
- Zero out all UW Inputs, Capital Activity, Revenue, etc.
- Run model (will produce zero outputs)
- Save with detail

### Wire Into Existing Infrastructure

- `KernelSeedManager.bas` calls `KernelSnapshot.ExportDetailToFile`, `KernelSnapshot.ExportInputsToFile`, `KernelTabIO.ExportAllInputTabs`, `KernelTabIO.ImportAllInputTabs`
- Reuse `KernelSnapshot.EnsureDirectoryExists`, `KernelSnapshot.GetProjectRoot`
- Reuse `KernelSnapshot.BuildConfigHash` for manifest configHash
- Seed folder is at `{project_root}/seeds/` (same level as `snapshots/` and `workspaces/`)

### Update CLAUDE.md

- Kernel modules 28 → 29 (add KernelSeedManager)
- Note seeds/ directory in project structure

### Validation Gates (Seed Manager)

12. Seed Manager button appears on Dashboard (user-visible, SortOrder 15)
13. "Save Current as Seed" creates seeds/{name}/ with manifest.json, inputs.csv, input_tabs/, detail.csv
14. "Load Seed" restores all input tabs and detail — formula tabs refresh correctly
15. Completeness check identifies missing tabs correctly
16. "Run & Save" on a seed without detail.csv runs model and writes detail.csv back to seed
17. Loading a seed updates scenario name on Assumptions tab
18. Two starter seeds exist: Base_Case and Clean_Slate
19. Seeds are read-only — loading does not modify seed folder (except Run & Save for missing detail)
