# Model Requirements Template

**Version:** 1.0
**Date:** April 2026
**Purpose:** Standardized template for documenting business requirements that lead to model building on the RDK platform. Every new model follows this template.

---

## 1. Model Identity

| Field | Value |
|-------|-------|
| Model Name | |
| Config Directory | config_{name}/ |
| Workbook Name | (branding_config WorkbookName setting) |
| Domain Module | {Name}DomainEngine.bas |
| Author | |
| Version | |

---

## 2. Business Context

**What decision does this model support?**

(Describe the business question. Who uses the output? What do they do with it?)

**What inputs does the user provide?**

(List input tabs, key assumptions, number of entities/programs)

**What outputs does the user consume?**

(List output tabs, key metrics, presentation format)

---

## 3. Formula Model (DS-01: Required)

The formula model is the primary product. It must function with no VBA execution.

### 3.1 Input Tabs

| Tab | Purpose | Key Inputs |
|-----|---------|------------|
| | | |

### 3.2 Formula Tabs (Output)

| Tab | Computes | References | Key RowIDs |
|-----|----------|------------|------------|
| | | | |

### 3.3 Simplified Assumptions

For each computation that the full model handles with VBA granularity, document the simplified formula assumption:

| Metric | Full Model (VBA) | Formula Model (Simplified) | Expected Tolerance |
|--------|-------------------|---------------------------|-------------------|
| | | | |

Example:
| Gross Earned Premium | Monthly earning pattern via days-on-risk | GEP = GWP (earned = written at quarterly grain) | < 5% |
| Loss Reserves | CDF curve development, monthly IBNR | Reserves = incurred * (1 - paid%) cumulative | < 10% |

---

## 4. Detail Model (DS-01: Optional)

The VBA detail model adds actuarial/analytical granularity. It runs on-demand via a button.

### 4.1 What VBA Computes

| Output | Granularity | Method |
|--------|-------------|--------|
| | | |

### 4.2 Detail Tab Schema

| Column | FieldClass | Description |
|--------|------------|-------------|
| | | |

### 4.3 Reconciliation Points

For each metric that exists in both the formula model and the VBA detail model, define the reconciliation check:

| Metric | Formula Source (Tab!RowID) | VBA Source (QS RowID or Detail aggregate) | Tolerance | Action on Fail |
|--------|---------------------------|------------------------------------------|-----------|---------------|
| | | | | |

---

## 5. Tab Registry

| Tab | Type | Category | QuarterlyColumns | Formula-Only? | VBA-Required? |
|-----|------|----------|-----------------|---------------|---------------|
| | | | | | |

Tabs marked "VBA-Required" are excluded from config_{name}_lite/.

---

## 6. Config Checklist

Every model must ship with both config variants:

- [ ] config_{name}/ — full model (VBA + formulas)
- [ ] config_{name}_lite/ — formula-only model
- [ ] branding_config.csv includes WorkbookName (different for each variant)
- [ ] regression_config.csv lists all formula output tabs
- [ ] reconciliation_config.csv (if detail model exists) defines tolerance checks
- [ ] Setup.bat updated with new option
- [ ] Both configs have all 31+ CSV files

---

## 7. Validation Gates

| Gate | Check | Pass Criteria |
|------|-------|---------------|
| 1 | Formula model bootstraps and populates all tabs | No #REF!, no blank output tabs |
| 2 | Input changes propagate to all output tabs | Change one input, verify downstream recalculation |
| 3 | Detail model runs without error | Run Model completes, Detail tab populated |
| 4 | Reconciliation passes | All metrics within tolerance |
| 5 | BS balances, CFS reconciles | BS_CHECK = 0, CFS_CHECK = 0 |
| 6 | Fingerprint stamped | VeryHidden _fp sheet present, copyright headers on all .bas |

---

## 8. Decision Log

| ID | Decision | Date | Rationale |
|----|----------|------|-----------|
| | | | |
