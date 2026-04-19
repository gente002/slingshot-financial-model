"""Fix Provision for Reinsurance #REF! by deriving it from CRSV x PROV_PCT.

Balance Sheet tab has no `BS_PROV_REINS` RowID (flagged as integration
risk in SESSION_NOTES 2026-04-18 entry #4). Instead of adding a new row
to the Balance Sheet (which we should not edit without wider consideration),
we derive Provision on-tab as CRSV x PROV_PCT, matching the reference
model's `ASSM_ProvReinsPC = 0.02` (NAIC 2% of recoverables for authorized
reinsurers).

Changes:
  1. Add new NAIC Charge `RBC_CHG_PROV_PCT` at row 62 (default 0.02).
  2. Shift rows >= 62 by +1 to make room.
  3. Replace Provision mirror formula from `{REF:Balance Sheet!BS_PROV_REINS}`
     to `{ROWID:RBC_MIR_CRSV}*$C$62`.
"""
import csv, io


def serialize(fields):
    buf = io.StringIO()
    w = csv.writer(buf, quoting=csv.QUOTE_ALL, lineterminator="")
    w.writerow(fields)
    return buf.getvalue()


def make_cell(rowid, row, col, ctype, content, **kw):
    return ["RBC Capital Model", rowid, str(row), str(col), ctype, content,
            kw.get("fmt",""), kw.get("font_style",""), kw.get("fill",""),
            kw.get("font_color",""), kw.get("col_span",""),
            kw.get("border_bot",""), kw.get("border_top",""),
            kw.get("indent",""), kw.get("comment",""), kw.get("halign","")]


def process(path):
    with open(path, "r", newline="") as f:
        raw = f.read()
    norm = raw.replace("\r\n", "\n").replace("\r", "\n")
    lines = norm.split("\n")
    trailing = lines and lines[-1] == ""
    if trailing: lines = lines[:-1]

    out = []
    rbc = []
    for ln in lines:
        if ln.startswith('"RBC Capital Model"'):
            rbc.append(next(csv.reader([ln])))
            out.append(("rbc", None))
        else:
            out.append(("keep", ln))

    # Step 1: replace Provision mirror formula
    for f in rbc:
        if f[1] == "RBC_MIR_PROVISION" and f[3] == "C" and f[4] == "Formula":
            f[5] = "={ROWID:RBC_MIR_CRSV}*$C$62"
            f[14] = ("Derived: Ceded Reserves x NAIC Provision % (default 2% "
                     "for authorized reinsurers). Type a dollar value over this "
                     "formula to override with a specific Schedule F provision.")

    # Also update the label comment to reflect new derivation
    for f in rbc:
        if f[1] == "RBC_MIR_PROVISION" and f[3] == "2" and f[4] == "Label":
            f[14] = "Computed: CRSV x PROV_PCT (row 62). Default NAIC 2%."

    # Step 2: shift rows >= 62 by +1
    for f in rbc:
        if not f[2].isdigit(): continue
        r = int(f[2])
        if r >= 62:
            f[2] = str(r + 1)

    # Step 3: insert new NAIC Charge at row 62
    rbc.append(make_cell("RBC_CHG_PROV_PCT", 62, 2, "Label",
        "Provision for Reinsurance % (of Recov)",
        indent="1",
        comment="NAIC default 2% for authorized reinsurers. Higher for unauthorized."))
    rbc.append(make_cell("RBC_CHG_PROV_PCT", 62, 3, "Input",
        "0.020", fmt="0.0%", font_color="0000FF"))

    # Sort by (row, col)
    def keyfn(f):
        try: r = int(f[2])
        except: r = 99999
        c = f[3]
        try: ci = int(c)
        except: ci = 500 + (ord(c[0].upper()) - ord("A"))
        return (r, ci, f[1])
    rbc.sort(key=keyfn)

    # Emit
    result = []
    emitted = False
    for kind, ln in out:
        if kind == "keep":
            result.append(ln)
        else:
            if not emitted:
                result.extend(serialize(f) for f in rbc)
                emitted = True
    final = "\r\n".join(result) + "\r\n"
    with open(path, "w", newline="") as f:
        f.write(final)
    print(f"{path[-60:]}: Provision fix applied")


if __name__ == "__main__":
    for p in [r"c:/Users/gente/Downloads/RDK_v1.3.0_20260403_latest/config/formula_tab_config.csv",
              r"c:/Users/gente/Downloads/RDK_v1.3.0_20260403_latest/config_insurance/formula_tab_config.csv"]:
        process(p)
