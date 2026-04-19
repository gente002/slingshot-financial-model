"""Static audit scanner for the RDK formula_tab_config.csv.

Scans for the four bug classes surfaced during the RBC-tab bug hunt of
2026-04-18/19 and reports every occurrence with file location + severity.
Intended to be run pre-commit and pre-packaging.

Bug classes detected:
  1. ROWID-to-static-cell in quarterly formula
     Col="C" (quarterly fan) formula that contains {ROWID:X} where X's
     target cell is a single-cell static input (Col numeric). Produces
     empty values at Q2+ because ROWID resolves per-current-column.
  2. REF-to-non-existent RowID
     {REF:Tab!RowID} where RowID does not exist on Tab. Produces #REF!.
  3. Duplicate-row collision
     Two different RowIDs occupying the same (row, col) cell, or the
     same RowID at conflicting rows. Typically breaks ROWID resolution.
  4. Sign-convention heuristic
     Mirror formula of the form `{REF:X!Y}` (no ABS) where Y's name
     matches a known signed-negative pattern (CRSV, CEDED_*, _CEDED,
     _CRSV). Flagged as potential sign bug; human review required.

This is intentionally a STATIC scanner. It does not execute formulas
or invoke Excel. It works by parsing the CSV config.
"""
import csv
import re
import sys
from collections import defaultdict

CONFIG_PATH = r"c:/Users/gente/Downloads/RDK_v1.3.0_20260403_latest/config/formula_tab_config.csv"
TAB_REGISTRY_PATH = r"c:/Users/gente/Downloads/RDK_v1.3.0_20260403_latest/config/tab_registry.csv"
NAMED_RANGE_PATH = r"c:/Users/gente/Downloads/RDK_v1.3.0_20260403_latest/config/named_range_registry.csv"

# Tabs whose RowIDs are populated at runtime by VBA (not present in
# formula_tab_config.csv). For these, {REF:Tab!X} references should be
# validated against named_range_registry.csv or accepted if the RowID
# matches a documented naming convention.
DYNAMIC_TABS = {
    "Quarterly Summary",   # Ins_QuarterlyAgg populates per-program QS_*_N
    "Detail",              # Domain engine populates the pipeline output
    "Run Metadata",        # KernelBootstrap/SaveRunState populates ENTITY_N etc.
    "Loss Triangles",      # Ins_Triangles populates
    "Count Triangles",     # Ins_Triangles populates
}

# Naming conventions recognized as dynamically-generated (e.g. per-program)
DYNAMIC_ROWID_PATTERNS = {
    "Quarterly Summary": [
        re.compile(r"^QS_[A-Z_]+_\d+$"),        # per-program
        re.compile(r"^QS_[A-Z_]+_TOTAL$"),      # aggregate
    ],
    "Run Metadata": [
        re.compile(r"^ENTITY_\d+$"),
        re.compile(r"^[A-Z_]+$"),
    ],
    "Detail": [
        re.compile(r"^.*$"),  # anything goes on Detail
    ],
    "Loss Triangles": [re.compile(r"^.*$")],
    "Count Triangles": [re.compile(r"^.*$")],
}

ROWID_TOKEN = re.compile(r"\{ROWID:([^}]+)\}")
REF_TOKEN = re.compile(r"\{REF:([^!}]+)!([^}]+)\}")
NAMED_TOKEN = re.compile(r"\{NAMED:([^}]+)\}")

# Heuristic name patterns for signed-negative mirrors
SIGNED_PATTERNS = [
    re.compile(r"CRSV$"),            # Ceded reserves
    re.compile(r"^.*CEDED.*$", re.I),
    re.compile(r"_CESSION$", re.I),
]


def load_named_range_tabs():
    """Return set of tab names that have RowIDs in the named-range registry."""
    tabs_with_named = set()
    try:
        with open(NAMED_RANGE_PATH, "r", newline="") as f:
            r = csv.reader(f)
            header = next(r)
            # columns: RangeName, TabName, RowID, CellAddress, RangeType, Description
            for row in r:
                if len(row) > 1 and row[1]:
                    tabs_with_named.add(row[1])
    except FileNotFoundError:
        pass
    return tabs_with_named


def load_tab_registry():
    """Return dict: tab_name -> {'quarterly': bool, 'category': str, 'type': str}."""
    out = {}
    with open(TAB_REGISTRY_PATH, "r", newline="") as f:
        r = csv.reader(f)
        header = next(r)
        # Column indices from header
        idx_name = header.index("TabName")
        idx_cat = header.index("Category")
        idx_type = header.index("Type")
        idx_q = header.index("QuarterlyColumns")
        for row in r:
            if len(row) <= idx_q:
                continue
            out[row[idx_name]] = {
                "quarterly": row[idx_q].upper() == "TRUE",
                "category": row[idx_cat],
                "type": row[idx_type],
            }
    return out


def load_config():
    """Return list of dicts for each formula_tab_config row."""
    rows = []
    with open(CONFIG_PATH, "r", newline="") as f:
        r = csv.reader(f)
        header = next(r)
        for row in r:
            if not row or not row[0]:
                continue
            rows.append({
                "tab": row[0],
                "rowid": row[1] if len(row) > 1 else "",
                "row": row[2] if len(row) > 2 else "",
                "col": row[3] if len(row) > 3 else "",
                "celltype": row[4] if len(row) > 4 else "",
                "content": row[5] if len(row) > 5 else "",
                "comment": row[14] if len(row) > 14 else "",
            })
    return rows


def build_rowid_maps(rows):
    """
    Returns:
      tab_rowid_cells: tab -> rowid -> list of (row, col, celltype)
      rowid_static_only: set of (tab, rowid) where ALL cells are Col numeric (single-cell)
    """
    tab_rowid_cells = defaultdict(lambda: defaultdict(list))
    for r in rows:
        if not r["rowid"]:
            continue
        tab_rowid_cells[r["tab"]][r["rowid"]].append((r["row"], r["col"], r["celltype"]))
    # A rowid is "static-only" if every cell has numeric col (single-cell write)
    rowid_static_only = set()
    for tab, rids in tab_rowid_cells.items():
        for rid, cells in rids.items():
            # Only consider Formula/Input/Label cells (not Spacer/Section); static if numeric col
            data_cells = [c for c in cells if c[2] in ("Formula", "Input", "Label")]
            if not data_cells:
                continue
            all_numeric = all(c[1].isdigit() for c in data_cells)
            if all_numeric:
                rowid_static_only.add((tab, rid))
    return tab_rowid_cells, rowid_static_only


def check_rowid_to_static(rows, tab_registry, rowid_static_only):
    """Bug class 1: quarterly Col="C" formula with {ROWID:X} where X is static-only."""
    findings = []
    for r in rows:
        if r["celltype"] != "Formula":
            continue
        # Only flag if formula lives in a cell that fans across quarters.
        # Fan trigger: tab has QuarterlyColumns=TRUE AND col is alpha (not numeric).
        tab_info = tab_registry.get(r["tab"], {})
        if not tab_info.get("quarterly"):
            continue
        if r["col"].isdigit():
            continue  # single-cell static formula, not fanned
        # Extract ROWID tokens
        for rid in ROWID_TOKEN.findall(r["content"]):
            if (r["tab"], rid) in rowid_static_only:
                findings.append({
                    "class": "ROWID-to-static-cell in quarterly formula",
                    "tab": r["tab"],
                    "rowid": r["rowid"],
                    "row": r["row"],
                    "col": r["col"],
                    "target": rid,
                    "severity": "CRITICAL",
                    "note": f"Resolves to empty cell at Q2+ columns",
                })
    return findings


def check_ref_resolves(rows, tab_rowid_cells, tab_registry):
    """Bug class 2: {REF:Tab!RowID} where RowID is not present on Tab.
    Tabs listed in DYNAMIC_TABS are exempt because their RowIDs are
    VBA-populated at runtime (e.g. Quarterly Summary per-program cells).
    For those, only the RowID naming convention is checked."""
    findings = []
    for r in rows:
        if r["celltype"] != "Formula":
            continue
        for (target_tab, target_rid) in REF_TOKEN.findall(r["content"]):
            # Tab must exist in tab_registry
            if target_tab not in tab_registry:
                findings.append({
                    "class": "REF-to-non-existent-tab",
                    "tab": r["tab"],
                    "rowid": r["rowid"],
                    "row": r["row"],
                    "col": r["col"],
                    "target": f"{target_tab}!{target_rid}",
                    "severity": "CRITICAL",
                    "note": f"Tab '{target_tab}' not in tab_registry",
                })
                continue
            # Dynamic tab: check naming convention, don't require CSV presence
            if target_tab in DYNAMIC_TABS:
                patterns = DYNAMIC_ROWID_PATTERNS.get(target_tab, [])
                if patterns and not any(p.match(target_rid) for p in patterns):
                    findings.append({
                        "class": "REF-to-dynamic-tab-unrecognized-pattern",
                        "tab": r["tab"],
                        "rowid": r["rowid"],
                        "row": r["row"],
                        "col": r["col"],
                        "target": f"{target_tab}!{target_rid}",
                        "severity": "MAJOR",
                        "note": f"RowID '{target_rid}' does not match any known pattern for dynamic tab '{target_tab}'",
                    })
                continue
            # Non-dynamic tab: RowID must be in formula_tab_config
            if target_rid not in tab_rowid_cells.get(target_tab, {}):
                findings.append({
                    "class": "REF-to-non-existent-RowID",
                    "tab": r["tab"],
                    "rowid": r["rowid"],
                    "row": r["row"],
                    "col": r["col"],
                    "target": f"{target_tab}!{target_rid}",
                    "severity": "CRITICAL",
                    "note": f"RowID '{target_rid}' not defined on tab '{target_tab}'",
                })
    return findings


def check_row_collisions(rows):
    """Bug class 3: two RowIDs at same (row, col) OR one RowID split across rows."""
    findings = []
    # Group by (tab, row, col) -> set of RowIDs
    cell_to_rowids = defaultdict(set)
    # And by (tab, rowid) -> set of rows
    rowid_to_rows = defaultdict(set)
    for r in rows:
        if not r["rowid"] or not r["row"].isdigit():
            continue
        cell_to_rowids[(r["tab"], r["row"], r["col"])].add(r["rowid"])
        rowid_to_rows[(r["tab"], r["rowid"])].add(r["row"])
    # Cell-level collisions
    for (tab, row, col), rids in cell_to_rowids.items():
        if len(rids) > 1:
            findings.append({
                "class": "Duplicate-cell collision",
                "tab": tab,
                "rowid": ", ".join(sorted(rids)),
                "row": row,
                "col": col,
                "target": "N/A",
                "severity": "CRITICAL",
                "note": f"{len(rids)} RowIDs at same (row, col). RowID cache will only index one.",
            })
    # RowID spread across multiple rows (should be single-row unless intentional)
    for (tab, rid), row_set in rowid_to_rows.items():
        if len(row_set) > 1:
            findings.append({
                "class": "RowID split across rows",
                "tab": tab,
                "rowid": rid,
                "row": ", ".join(sorted(row_set)),
                "col": "N/A",
                "target": "N/A",
                "severity": "MAJOR",
                "note": f"RowID '{rid}' appears at rows {sorted(row_set)}. ROWID resolution is ambiguous.",
            })
    return findings


def check_sign_convention(rows):
    """Bug class 4: {REF:X!Y} without ABS where Y matches signed-negative pattern."""
    findings = []
    for r in rows:
        if r["celltype"] != "Formula":
            continue
        # Only check mirror-style formulas (no algebra beyond the REF)
        for (target_tab, target_rid) in REF_TOKEN.findall(r["content"]):
            for pat in SIGNED_PATTERNS:
                if pat.search(target_rid):
                    # Check if this REF is inside an ABS() wrapper OR preceded
                    # by a leading minus sign (both flip the sign correctly).
                    ref_literal = f"{{REF:{target_tab}!{target_rid}}}"
                    if "ABS(" + ref_literal in r["content"]:
                        continue  # already wrapped
                    if "=-" + ref_literal in r["content"] or "-" + ref_literal in r["content"]:
                        continue  # explicit sign flip
                    findings.append({
                        "class": "Signed-convention mirror without ABS",
                        "tab": r["tab"],
                        "rowid": r["rowid"],
                        "row": r["row"],
                        "col": r["col"],
                        "target": f"{target_tab}!{target_rid}",
                        "severity": "MAJOR",
                        "note": f"Target RowID name '{target_rid}' matches signed-negative pattern. Review: should it be wrapped in ABS()?",
                    })
                    break
    return findings


def main():
    rows = load_config()
    tab_registry = load_tab_registry()
    tab_rowid_cells, rowid_static_only = build_rowid_maps(rows)

    all_findings = []
    all_findings += check_rowid_to_static(rows, tab_registry, rowid_static_only)
    all_findings += check_ref_resolves(rows, tab_rowid_cells, tab_registry)
    all_findings += check_row_collisions(rows)
    all_findings += check_sign_convention(rows)

    # Summary
    by_class = defaultdict(list)
    for f in all_findings:
        by_class[f["class"]].append(f)

    print("=" * 80)
    print("RDK STATIC AUDIT SCAN")
    print(f"Scanned {len(rows):,} config rows across {len(tab_registry)} tabs")
    print("=" * 80)
    print()
    print(f"TOTAL FINDINGS: {len(all_findings)}")
    for cls, fs in sorted(by_class.items()):
        sev = fs[0]["severity"]
        print(f"  [{sev:<8}] {cls:<45} {len(fs):>4}")
    print()

    # Detailed by class
    for cls, fs in sorted(by_class.items()):
        if not fs: continue
        print(f"\n--- {cls} ({len(fs)}) ---")
        # Group by tab
        by_tab = defaultdict(list)
        for f in fs:
            by_tab[f["tab"]].append(f)
        for tab in sorted(by_tab):
            print(f"  Tab: {tab}  ({len(by_tab[tab])} findings)")
            for f in by_tab[tab][:5]:  # cap per tab to avoid flood
                print(f"    row {f['row']:<4} col {f['col']:<3} {f['rowid']:<28} -> {f['target']}")
                print(f"          {f['note']}")
            if len(by_tab[tab]) > 5:
                print(f"    ... and {len(by_tab[tab])-5} more")

    # Exit code = number of CRITICAL findings
    crit = sum(1 for f in all_findings if f["severity"] == "CRITICAL")
    return crit


if __name__ == "__main__":
    sys.exit(main())
