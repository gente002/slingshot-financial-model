"""
Regenerate config/assumptions_schema.csv (and config_insurance/ copy) from:
  - config/input_schema.csv         (Assumptions tab globals)
  - config/formula_tab_config.csv   (all Input-typed rows)
  - config/granularity_config.csv   (TimeHorizon for quarterly expansion)

Every schema row maps to exactly one cell. Quarterly replicator inputs
(Col="C" in formula_tab_config) are expanded to one row per quarterly
and annual-total column across the horizon.

Run this whenever:
  - TimeHorizon changes in granularity_config.csv
  - Input rows are added, removed, or relocated in formula_tab_config.csv
  - input_schema.csv is modified

Usage:
    python scripts/regen_assumptions_schema.py
"""
import csv
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# Layout constants -- match engine/KernelConstants.bas
DATA_START_COL = 3   # col C
COLS_PER_YEAR = 5    # Q1, Q2, Q3, Q4, YearTotal


def col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def infer_datatype(fmt: str) -> str:
    f = (fmt or "").lower()
    if "%" in f: return "Pct"
    if "$" in f or "currency" in f: return "Currency"
    if "#" in f: return "Number"
    if "yyyy" in f or "mm" in f or "date" in f: return "Date"
    return "Text"


def read_horizon_years(granularity_path: Path) -> int:
    with granularity_path.open(newline="") as f:
        for row in csv.DictReader(f):
            if row["Setting"] == "TimeHorizon":
                return int(row["Value"]) // 12
    return 5


def generate(config_dir: Path) -> list:
    input_schema = config_dir / "input_schema.csv"
    formula_tab  = config_dir / "formula_tab_config.csv"
    granularity  = config_dir / "granularity_config.csv"

    num_years = read_horizon_years(granularity)
    print(f"[{config_dir.name}] Horizon: {num_years} years -> "
          f"{num_years * COLS_PER_YEAR} replicator cells per row")

    rows = []

    # 1) Assumptions tab globals
    with input_schema.open(newline="") as f:
        for row in csv.DictReader(f):
            rows.append({
                "TabName": "Assumptions",
                "AssumptionID": row["ParamName"],
                "Address": f"$C${row['Row']}",
                "Section": row["Section"],
                "DataType": row["DataType"],
                "DefaultValue": row["Default"],
                "Description": row["Tooltip"],
            })

    # 2) Input-typed rows from formula_tab_config
    static_cnt = 0
    replicator_cnt = 0
    with formula_tab.open(newline="") as f:
        for row in csv.DictReader(f):
            if row["CellType"] != "Input":
                continue
            col_raw = row["Col"]
            dt = infer_datatype(row.get("Format", ""))
            try:
                col_num = int(col_raw)
            except ValueError:
                col_num = None

            if col_num is not None:
                rows.append({
                    "TabName": row["TabName"],
                    "AssumptionID": row["RowID"],
                    "Address": f"${col_letter(col_num)}${row['Row']}",
                    "Section": row["TabName"],
                    "DataType": dt,
                    "DefaultValue": row["Content"],
                    "Description": row.get("Comment", ""),
                })
                static_cnt += 1
            elif col_raw == "C":
                for yr in range(num_years):
                    for q in range(COLS_PER_YEAR):
                        c = DATA_START_COL + yr * COLS_PER_YEAR + q
                        rows.append({
                            "TabName": row["TabName"],
                            "AssumptionID": row["RowID"],
                            "Address": f"${col_letter(c)}${row['Row']}",
                            "Section": row["TabName"],
                            "DataType": dt,
                            "DefaultValue": row["Content"],
                            "Description": row.get("Comment", ""),
                        })
                replicator_cnt += 1
            else:
                print(f"  WARN: unhandled Col value '{col_raw}' "
                      f"for {row['TabName']}/{row['RowID']}", file=sys.stderr)

    # Dedup by (TabName, AssumptionID, Address)
    seen, uniq = set(), []
    for r in rows:
        key = (r["TabName"], r["AssumptionID"], r["Address"])
        if key in seen:
            continue
        seen.add(key)
        uniq.append(r)

    print(f"  Static: {static_cnt}  Replicators: {replicator_cnt} "
          f"(-> {replicator_cnt * num_years * COLS_PER_YEAR} cells)")
    print(f"  Total schema rows: {len(uniq)}")
    return uniq


def write_schema(rows: list, out_path: Path) -> None:
    cols = ["TabName", "AssumptionID", "Address", "Section", "DataType",
            "DefaultValue", "Description"]
    with out_path.open("w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols, quoting=csv.QUOTE_ALL)
        w.writeheader()
        for r in rows:
            w.writerow(r)


def write_meta(cfg_dir: Path, rows: list, num_years: int) -> None:
    """Write assumptions_schema.meta.csv with a fingerprint of the inputs.

    Drift detection: at import time, VBA reloads the current values of the
    source inputs and compares against the stored fingerprint. Mismatch
    means the schema is stale relative to the current config -- the user
    should re-run this script before importing.
    """
    import datetime

    # Fingerprint components: count of Input rows in formula_tab_config,
    # count of rows in input_schema, and numYears from granularity.
    input_count = 0
    replicator_count = 0
    with (cfg_dir / "formula_tab_config.csv").open(newline="") as f:
        for row in csv.DictReader(f):
            if row["CellType"] == "Input":
                input_count += 1
                if row["Col"] == "C":
                    replicator_count += 1

    global_count = 0
    with (cfg_dir / "input_schema.csv").open(newline="") as f:
        for _ in csv.DictReader(f):
            global_count += 1

    meta = [
        {"Key": "GeneratedAt",         "Value": datetime.datetime.now().isoformat(timespec="seconds")},
        {"Key": "NumYears",            "Value": str(num_years)},
        {"Key": "InputRowCount",       "Value": str(input_count)},
        {"Key": "ReplicatorRowCount",  "Value": str(replicator_count)},
        {"Key": "GlobalRowCount",      "Value": str(global_count)},
        {"Key": "SchemaRowCount",      "Value": str(len(rows))},
    ]

    out_path = cfg_dir / "assumptions_schema.meta.csv"
    with out_path.open("w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["Key", "Value"], quoting=csv.QUOTE_ALL)
        w.writeheader()
        for r in meta:
            w.writerow(r)
    print(f"  meta: {out_path}")


def main() -> int:
    ok = True
    for cfg_dir_name in ("config", "config_insurance"):
        cfg_dir = ROOT / cfg_dir_name
        if not cfg_dir.exists():
            print(f"SKIP: {cfg_dir_name} not found")
            continue
        try:
            rows = generate(cfg_dir)
        except Exception as e:
            print(f"ERROR in {cfg_dir_name}: {e}", file=sys.stderr)
            ok = False
            continue

        out = cfg_dir / "assumptions_schema.csv"
        write_schema(rows, out)
        print(f"  -> {out}")

        num_years = read_horizon_years(cfg_dir / "granularity_config.csv")
        write_meta(cfg_dir, rows, num_years)

        tabs = Counter(r["TabName"] for r in rows)
        for tab, n in tabs.most_common():
            print(f"      {tab}: {n}")
        print()

    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
