# Revenue Expansion — CC Build Prompt Addendum

Read `CLAUDE.md`, then `SESSION_NOTES.md`. This is an addendum to the current session.

## Overview

Restructure Other Revenue Detail from generic SW1-SW5 / FEE / CON slots to purpose-built revenue lines. Add program count metric to UW Exec Summary. Add formula-driven scaling (% of GWP, per-program). Seed with modest defaults. Update Revenue Summary wiring.

---

## 1. UW EXEC SUMMARY — Add Program Count

Add one new formula row to UW Exec Summary, near the top (after UWEX_GWP or in a supporting metrics section):

| RowID | Label | CellType | Formula |
|---|---|---|---|
| UWEX_PROG_COUNT | Active Program Count | Formula | Count of non-blank program names on UW Inputs |

The formula should count how many programs have a non-blank name. The exact implementation depends on how UW Inputs stores program names — likely counting non-blank cells in the program name column across all 10 program blocks.

Add a named range:
```csv
"UWEX_Q_ProgCount","UW Exec Summary","UWEX_PROG_COUNT","","Quarterly","Active program count"
```

**Note:** Program count is the same for all quarters (programs don't change quarter to quarter in the current model). The value in every quarterly column should be the same — it's a static count, not a time-varying metric. Use `$C$[row]` style reference or repeat the count formula across all quarters.

---

## 2. OTHER REVENUE DETAIL — Full Restructure

**Delete all existing ORD_SW1 through ORD_SW5, ORD_FEE_MGA/ADMIN/OTHER, ORD_CON_ACT/RISK/OTHER rows.** Replace with the structure below. Keep the tab header (ORD_HEADER, ORD_BASIS).

### Section 1: Technology Revenue

| RowID | Label | CellType | Scale Type | Default | Formula/Notes |
|---|---|---|---|---|---|
| ORD_SEC_TECH | Technology Revenue | Section | | | D9E1F2 fill |
| ORD_SW_PLATFORM_RATE | Platform Subscription ($/program/year) | Input | Per Program | 25000 | Annual rate per program. Col C only (static). |
| ORD_SW_PLATFORM | Platform Subscriptions | Formula | Per Program | | `={REF:UW Exec Summary!UWEX_PROG_COUNT} * $C$[rate_row] / 4` |
| ORD_SW_API_RATE | API / Transaction Fee ($/policy) | Input | Per Policy | 15 | Per policy per year. Col C only. |
| ORD_SW_API | API & Transaction Fees | Formula | Per Policy | | `={REF:UW Exec Summary!UWEX_GWP} / 50000 * $C$[rate_row] / 4` (assume avg premium $50K to estimate policy count) |
| ORD_SW_IMPLEMENT | Implementation Fees | Input | Fixed | 0 | One-time per new program. User enters when programs onboard. |
| ORD_SW_ANALYTICS_RATE | Analytics Products ($/program/year) | Input | Per Program | 10000 | Annual rate per program. |
| ORD_SW_ANALYTICS | Analytics & Data Products | Formula | Per Program | | `={REF:UW Exec Summary!UWEX_PROG_COUNT} * $C$[rate_row] / 4` |
| ORD_SW_LICENSE | Licensing & White-Label | Input | Fixed | 0 | |
| ORD_SW_CUSTOM | Custom Development | Input | Fixed | 0 | |
| ORD_SW_OTHER | Other Technology Revenue | Input | Fixed | 0 | |
| ORD_SW_TOTAL | Total Technology Revenue | Formula | | | Sum of ORD_SW_PLATFORM + ORD_SW_API + ORD_SW_IMPLEMENT + ORD_SW_ANALYTICS + ORD_SW_LICENSE + ORD_SW_CUSTOM + ORD_SW_OTHER |

### Section 2: Fee Income

| RowID | Label | CellType | Scale Type | Default | Formula/Notes |
|---|---|---|---|---|---|
| ORD_SEC_FEE | Fee Income | Section | | | D9E1F2 fill |
| ORD_FEE_CARRIER_RATE | Carrier Access Fee Rate (% of GWP) | Input | % of GWP | 0.02 | 2% of GWP. Col C only (static rate). |
| ORD_FEE_CARRIER | Carrier Access / Program Fee | Formula | % of GWP | | `={REF:UW Exec Summary!UWEX_GWP} * $C$[rate_row]` |
| ORD_FEE_OVERSIGHT_RATE | Oversight Fee ($/program/year) | Input | Per Program | 15000 | Annual per program. |
| ORD_FEE_OVERSIGHT | Oversight & Monitoring | Formula | Per Program | | `={REF:UW Exec Summary!UWEX_PROG_COUNT} * $C$[rate_row] / 4` |
| ORD_FEE_ONBOARD | Program Onboarding (one-time) | Input | Fixed | 0 | User enters when new programs launch. |
| ORD_FEE_POLICY_RATE | Policy Fee ($/policy) | Input | Per Policy | 50 | Per policy. |
| ORD_FEE_POLICY | Policy & Binding Fees | Formula | Per Policy | | `={REF:UW Exec Summary!UWEX_GWP} / 50000 * $C$[rate_row] / 4` |
| ORD_FEE_ADMIN | Administrative Fees | Input | Fixed | 0 | Endorsements, cancellations, misc. |
| ORD_FEE_OTHER | Other Fee Income | Input | Fixed | 0 | |
| ORD_FEE_TOTAL | Total Fee Income | Formula | | | Sum of ORD_FEE_CARRIER + ORD_FEE_OVERSIGHT + ORD_FEE_ONBOARD + ORD_FEE_POLICY + ORD_FEE_ADMIN + ORD_FEE_OTHER |

### Section 3: Consulting Revenue

| RowID | Label | CellType | Scale Type | Default | Formula/Notes |
|---|---|---|---|---|---|
| ORD_SEC_CON | Consulting Revenue | Section | | | D9E1F2 fill |
| ORD_CON_ACT | Actuarial Consulting | Input | Fixed | 0 | |
| ORD_CON_RISK | Risk Management | Input | Fixed | 0 | |
| ORD_CON_REG | Regulatory / Compliance | Input | Fixed | 0 | |
| ORD_CON_OTHER | Other Consulting | Input | Fixed | 0 | |
| ORD_CON_TOTAL | Total Consulting Revenue | Formula | | | Sum of all CON lines |

### Grand Total

| RowID | Label | Formula |
|---|---|---|
| ORD_TOTAL | Total Other Revenue | `={ROWID:ORD_SW_TOTAL}+{ROWID:ORD_FEE_TOTAL}+{ROWID:ORD_CON_TOTAL}` |

**Total: 15 input lines (7 rates + 8 flat inputs) + 8 formula lines + 3 section headers + 1 grand total = ~27 rows** (plus spacers).

---

## 3. REVENUE SUMMARY — Update Wiring

Revenue Summary currently references:
- `ORD_SW_TOTAL` → keep (same RowID, different content)
- `ORD_FEE_TOTAL` → keep (same RowID)
- `ORD_CON_TOTAL` → keep (same RowID)

**The wiring doesn't change** because the total RowIDs are preserved. Revenue Summary already picks up Software, Fee, and Consulting totals by reference.

However, rename the Revenue Summary labels to match:
- REV_SOFTWARE label: change from "Software Revenue" to "Technology Revenue"
- Keep REV_FEE as "Fee Income"
- Keep REV_CONSULT as "Consulting Revenue"

---

## 4. POLICY COUNT ESTIMATION

Several formulas use `GWP / 50000` to estimate policy count (assuming $50K average premium). This is a simplification.

**Add a configurable average premium input** on the Other Revenue Detail tab:

| RowID | Label | CellType | Default | Notes |
|---|---|---|---|---|
| ORD_AVG_PREMIUM | Average Premium per Policy | Input | 50000 | Used to estimate policy count for per-policy fee calculations. Col C only. |

Then the per-policy formulas become:
```
={REF:UW Exec Summary!UWEX_GWP} / $C$[avg_prem_row] * $C$[fee_rate_row] / 4
```

Place this input at the top of the tab, before the Technology section, as a "Model Parameters" section:

| RowID | Label |
|---|---|
| ORD_SEC_PARAMS | Revenue Model Parameters |
| ORD_AVG_PREMIUM | Average Premium per Policy |

---

## 5. SEEDED DEFAULTS SUMMARY

With the defaults above, a startup carrier with 3 programs and $5M GWP/quarter would show:

| Category | Quarterly | Annual |
|---|---|---|
| Platform Subscriptions (3 × $25K/yr) | $18,750 | $75,000 |
| API Fees (100 policies × $15) | $375 | $1,500 |
| Analytics (3 × $10K/yr) | $7,500 | $30,000 |
| Carrier Access Fee (2% × $5M) | $100,000 | $400,000 |
| Oversight Fee (3 × $15K/yr) | $11,250 | $45,000 |
| Policy Fees (100 × $50) | $1,250 | $5,000 |
| **Total Other Revenue** | **~$139K** | **~$557K** |

Plus UW Revenue (NEP + fronting fees) and Investment Income from existing tabs.

---

## VALIDATION GATES (Revenue)

20. All new RowIDs resolve correctly — no #REF! errors
21. UWEX_PROG_COUNT shows correct count of active programs
22. Formula-driven revenues scale when GWP changes (change a program's premium → carrier fee changes)
23. Per-program revenues scale when program count changes (add a program → platform subscription increases)
24. ORD_TOTAL = ORD_SW_TOTAL + ORD_FEE_TOTAL + ORD_CON_TOTAL
25. Revenue Summary picks up new totals (same RowID references)
26. Revenue Summary label updated: "Software Revenue" → "Technology Revenue"
27. IS Total Revenue reflects the expanded revenue
28. BS still balances, CFS still reconciles
29. Seeded defaults produce reasonable quarterly Other Revenue (~$139K/quarter with 3 programs, $5M GWP)
