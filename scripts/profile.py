"""Read-only format profiler for the genealogist's horizontal-tree spreadsheets.

Each new tree file may lay its columns out differently from the reference file
(`BlsGrnLivMcCl H.Tr. #305`). Before trusting the reader's hardwired column map on a
new file we profile it: characterise every column by content type, guess which field
each column holds, and diff that guess against the reference layout in reader.py.

Usage:
    PYTHONPATH=. uv run python scripts/profile.py "data/input/Stiff:Taylor H.Tree #275"
    PYTHONPATH=. uv run python scripts/profile.py            # profiles all data/input/*

Read-only: opens workbooks, prints a report, writes nothing.
"""

from __future__ import annotations

import glob
import os
import re
import sys

import xlrd

from family_tree_converter import reader as R

# The reference layout, pulled straight from reader.py so this never drifts.
REF_COLS = {
    R._C_GENERATION: "generation",
    R._C_CODE: "code",
    R._C_FATHER: "father",
    R._C_MOTHER: "mother",
    R._C_SURNAME: "surname",
    R._C_GIVEN: "given",
    R._C_DATE1: "birth/marr-date",
    R._C_FLAG: "flag(C/M)",
    R._C_TOWN: "town",
    R._C_COUNTY: "county",
    R._C_DEATH_DATE: "death-date",
    R._C_BURIED: "buried",
    R._C_LONGEVITY: "longevity",
    R._C_MARRIAGE: "marr-date",
    R._C_MARRIED_PLACE: "marr-place",
    R._C_OCCUPATION: "occupation",
    R._C_NOTES: "notes",
}
for _c in range(R._C_LINE_FIRST, R._C_LINE_LAST + 1):
    REF_COLS[_c] = "lineage"

_CODE_RE = re.compile(r"^[A-Za-z]{2,}([/|\-][A-Za-z0-9()?]+)*$")  # HntJm/Jn, JnsB-C
_DATESERIAL_RE = re.compile(r"^(1[5-9]\d{2}|20\d{2})([01]\d)([0-3]\d)$")  # YYYYMMDD
_NAME_RE = re.compile(r"^[A-Z][a-z]+(\s+[A-Z][a-z'.]+)+")  # John Smith, Mary Ann


def _classify(val) -> str:
    """One-word content class for a single cell value."""
    if val == "" or val is None:
        return ""
    if isinstance(val, (int, float)):
        f = float(val)
        if f == int(f):
            n = int(f)
            if 1500 <= n <= 2100:
                return "year"
            if 15000000 <= n <= 21000000:
                return "dateserial"
            if 1 <= n <= 120:
                return "smallnum"
            if 20000 <= n <= 60000:
                return "excelserial"
        return "num"
    s = str(val).strip()
    if not s:
        return ""
    if s in ("C", "M", "X", "D", "Ø"):
        return f"flag:{s}"
    if _DATESERIAL_RE.match(s):
        return "dateserial-str"
    if _CODE_RE.match(s) and any(c in s for c in "/|-") and len(s) <= 14:
        return "code"
    if _NAME_RE.match(s):
        return "name"
    if s.isupper() and s.isalpha() and len(s) <= 18:
        return "SURNAME"
    if len(s) > 40:
        return "longtext"
    return "text"


def _data_start(ws) -> int:
    """First row that looks like real data: a generation number in col 4, else
    the first row with >=4 non-empty cells after the legend block."""
    for r in range(min(ws.nrows, 40)):
        try:
            val = ws.cell(r, R._C_GENERATION).value
        except IndexError:
            continue
        if isinstance(val, (int, float)) and val != "" and 0 <= float(val) <= 60:
            return r
    # fallback: first dense row
    for r in range(min(ws.nrows, 40)):
        filled = sum(1 for c in range(ws.ncols) if str(ws.cell(r, c).value).strip())
        if filled >= 4:
            return r
    return 17


def profile(path: str) -> None:
    name = os.path.basename(path)
    wb = xlrd.open_workbook(path)
    ws = wb.sheet_by_index(0)
    ds = _data_start(ws)
    print(f"\n{'='*78}\n### {name}")
    print(f"  dims={ws.nrows}x{ws.ncols}  notes={len(ws.cell_note_map)}  data_start_row={ds}"
          f"  (ref={R.DATA_START_ROW})")

    rows = range(ds, ws.nrows)
    print(f"  {'col':>3} {'ref-label':<15} {'top-classes':<34} sample")
    for c in range(ws.ncols):
        from collections import Counter
        classes = Counter()
        samples = []
        for r in rows:
            val = ws.cell(r, c).value
            cls = _classify(val)
            if cls:
                classes[cls] += 1
                if len(samples) < 3 and str(val).strip():
                    sv = str(val).strip()
                    samples.append(sv[:22])
        if not classes:
            continue
        top = ", ".join(f"{k}:{v}" for k, v in classes.most_common(3))
        ref = REF_COLS.get(c, "")
        flag = "" if (ref or not classes) else "  <-- not in ref map"
        print(f"  {c:>3} {ref:<15} {top:<34} {' | '.join(samples)}{flag}")


def main() -> None:
    args = sys.argv[1:]
    if args:
        targets = args
    else:
        targets = sorted(
            f for f in glob.glob("data/input/*")
            if not f.endswith(".DS_Store")
        )
    for t in targets:
        profile(t)


if __name__ == "__main__":
    main()
