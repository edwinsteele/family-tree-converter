"""Read individuals and relationships from the source Excel spreadsheet."""

from __future__ import annotations

import calendar
import datetime
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import xlrd

# Column indices
_C_GENERATION = 4
_C_CODE = 15
_C_FATHER = 16
_C_MOTHER = 17
_C_SURNAME = 18
_C_GIVEN = 19
_C_DATE1 = 20       # birth date (individual rows) or marriage date (marriage rows)
_C_FLAG = 21        # 'C'=christening, 'M'=marriage row, ''=birth
_C_TOWN = 22
_C_COUNTY = 23
_C_DEATH_DATE = 24
_C_BURIED = 25
_C_LONGEVITY = 29      # recorded age at death (not emitted; cross-check only)
_C_MARRIAGE = 31       # marriage date on an individual (spouse) row
_C_MARRIED_PLACE = 32  # where married, on an individual (spouse) row
_C_OCCUPATION = 33
_C_NOTES = 34
_C_LINE_FIRST = 38     # first of the principal-lineage membership columns
_C_LINE_LAST = 42      # last of the principal-lineage membership columns

DATA_START_ROW = 17

_MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN",
           "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]

# Month names (full and common abbreviations) → GEDCOM 3-letter code.
_MONTH_NAMES = {
    "january": "JAN", "jan": "JAN", "february": "FEB", "feb": "FEB",
    "march": "MAR", "mar": "MAR", "april": "APR", "apr": "APR",
    "may": "MAY", "june": "JUN", "jun": "JUN", "july": "JUL", "jul": "JUL",
    "august": "AUG", "aug": "AUG", "september": "SEP", "sept": "SEP",
    "sep": "SEP", "october": "OCT", "oct": "OCT", "november": "NOV",
    "nov": "NOV", "december": "DEC", "dec": "DEC",
}


@dataclass
class Individual:
    id: str
    given_name: str
    surname: str
    birth_date: str | None = None
    birth_place: str | None = None
    birth_is_christening: bool = False
    death_date: str | None = None
    death_place: str | None = None
    sex: str | None = None
    occupation: str | None = None
    notes: str | None = None
    # Free-text annotations harvested from Excel cell comments.
    note_list: list[str] = field(default_factory=list)
    # Principal lineage charts this person belongs to (e.g. {"Belshaw"}).
    lineage_lines: set[str] = field(default_factory=set)
    # Successive married surnames for a woman recorded as "X then Y".
    married_surnames: list[str] = field(default_factory=list)


@dataclass
class Family:
    id: str
    husband_id: str | None = None
    wife_id: str | None = None
    marriage_date: str | None = None
    marriage_place: str | None = None
    child_ids: list[str] = field(default_factory=list)
    # Free-text annotations harvested from Excel cell comments on marriage rows.
    note_list: list[str] = field(default_factory=list)
    # Internal: base code of the husband row, used during construction
    _husband_base: str = field(default="", repr=False, compare=False)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_DECADE_RE = re.compile(
    r"^(?P<qualifier>early|mid|late|v\.approx\.|approx\.)?\s*(?P<year>\d{4})s?$",
    re.IGNORECASE,
)


def _parse_approx_string(s: str) -> str | None:
    """Convert informal approximate date strings to GEDCOM date phrases."""
    if s in ("?", "??", "unknown"):
        return None

    # Normalise separators so "mid.1950s" and "mid 1950s" both match
    normalised = s.replace(".", " ").strip().lower()

    # "v approx YYYY" → "ABT YYYY"
    m = re.match(r"^v\s+approx\s+(\d{4})$", normalised)
    if m:
        return f"ABT {m.group(1)}"

    # "approx YYYY" → "ABT YYYY"
    m = re.match(r"^approx\s+(\d{4})$", normalised)
    if m:
        return f"ABT {m.group(1)}"

    # Decade patterns: "1900s", "mid 1950s", "late 1980s", "early 1900s"
    m = re.match(r"^(?P<qual>early|mid|late)?\s*(?P<decade>\d{3})0s?$", normalised)
    if m:
        base = int(m.group("decade")) * 10
        qual = m.group("qual")
        if qual == "early":
            return f"BET {base} AND {base + 4}"
        if qual == "mid":
            return f"BET {base + 3} AND {base + 7}"
        if qual == "late":
            return f"BET {base + 6} AND {base + 9}"
        # bare "1900s"
        return f"BET {base} AND {base + 9}"

    # "1930/1", "1831/2" → uncertain year, convert to BET range
    m = re.match(r"^(\d{4})/\d{1,2}$", s.strip())
    if m:
        year = int(m.group(1))
        return f"BET {year} AND {year + 1}"

    # "1828 ??" → uncertain single year → ABT
    m = re.match(r"^(\d{4})\s*\?+$", s.strip())
    if m:
        return f"ABT {m.group(1)}"

    # ISO "1908-09-05" → "5 SEP 1908"; "1908-09" → "SEP 1908"
    m = re.match(r"^(\d{4})-(\d{2})(?:-(\d{2}))?$", s.strip())
    if m:
        year, month = int(m.group(1)), int(m.group(2))
        if 1 <= month <= 12:
            if m.group(3):
                return f"{int(m.group(3))} {_MONTHS[month - 1]} {year}"
            return f"{_MONTHS[month - 1]} {year}"

    # Uncertain-decade ISO "194?-02-28" → "BET 1940 AND 1949"
    m = re.match(r"^(\d{3})\?-\d{2}(?:-\d{2})?$", s.strip())
    if m:
        base = int(m.group(1)) * 10
        return f"BET {base} AND {base + 9}"

    # "pre 1911" → "BEF 1911"; "post/after 1911" → "AFT 1911"
    m = re.match(r"^(pre|before)\s+(\d{4})$", normalised)
    if m:
        return f"BEF {m.group(2)}"
    m = re.match(r"^(post|after)\s+(\d{4})$", normalised)
    if m:
        return f"AFT {m.group(2)}"

    # "c 1920", "ca 1920", "circa 1920" → "ABT 1920"
    m = re.match(r"^(c|ca|circa)\s+(\d{4})$", normalised)
    if m:
        return f"ABT {m.group(2)}"

    # Month-name forms: "April 1888", "Feb. 1948", "5 June 1789"
    m = re.match(r"^(?:(\d{1,2})\s+)?([a-z]+)\s+(\d{4})$", normalised)
    if m and m.group(2) in _MONTH_NAMES:
        mon = _MONTH_NAMES[m.group(2)]
        if m.group(1):
            return f"{int(m.group(1))} {mon} {m.group(3)}"
        return f"{mon} {m.group(3)}"

    return s  # pass through unchanged; writer will emit as-is


def _parse_date(val: Any) -> str | None:
    """Convert YYYYMMDD float or approximate string to a GEDCOM-friendly date."""
    if val == "" or val is None:
        return None
    if isinstance(val, float):
        date_int = int(val)
        year = date_int // 10000
        remainder = date_int % 10000
        month = remainder // 100
        day = remainder % 100
        if year < 1000:
            if date_int > 2200:
                # Excel date serial (e.g. 24166 → 28 FEB 1966)
                d = datetime.date(1899, 12, 30) + datetime.timedelta(days=date_int)
                return f"{d.day} {_MONTHS[d.month - 1]} {d.year}"
            # Year-only stored as plain integer (e.g. 1760.0)
            return str(date_int)
        if 1 <= month <= 12:
            if day > 0:
                # Clamp impossible days (e.g. 29 FEB 1978, a data error) to
                # the last valid day of the month rather than emitting junk.
                last_day = calendar.monthrange(year, month)[1]
                day = min(day, last_day)
                return f"{day} {_MONTHS[month - 1]} {year}"
            return f"{_MONTHS[month - 1]} {year}"
        return str(year)
    s = str(val).strip()
    if not s:
        return None
    return _parse_approx_string(s)


def _date_precision_note(val: Any, label: str) -> str | None:
    """Return a note preserving precision lost when a date is degraded.

    e.g. "194?-02-28" has a known day and month but a decade-uncertain year,
    so it is emitted as "BET 1940 AND 1949". The headline range hides that the
    event was a 28 Feb; this note records it so nothing is silently lost.
    """
    if not isinstance(val, str):
        return None
    m = re.match(r"^(\d{3})\?-(\d{2})(?:-(\d{2}))?$", val.strip())
    if not m:
        return None
    month = int(m.group(2))
    if not 1 <= month <= 12:
        return None
    known = _MONTHS[month - 1]
    if m.group(3):
        known = f"{int(m.group(3))} {known}"
    decade = int(m.group(1)) * 10
    return (
        f"{label} date recorded as \"{val.strip()}\": {known} is known, but "
        f"the year is uncertain within the {decade}s."
    )


def _longevity_discrepancy_note(
    birth: str | None, death: str | None, longevity: Any
) -> str | None:
    """Flag a contradiction between the recorded age at death (col 29) and the
    age implied by the emitted birth and death dates.

    Surfaces a likely transcription error for other researchers without
    asserting *which* fact is wrong (and without inventing a corrected date).
    Only fires for exact years on both dates and a numeric longevity, and only
    when they disagree by more than a year — so approximate dates and
    infant-death 'Months/Days' longevities never produce noise.
    """
    if not isinstance(longevity, (int, float)) or not birth or not death:
        return None
    if any(q in birth or q in death for q in ("ABT", "BET", "BEF", "AFT", "EST")):
        return None
    bm = re.search(r"\b(\d{4})\b", birth)
    dm = re.search(r"\b(\d{4})\b", death)
    if not (bm and dm):
        return None
    computed = int(dm.group(1)) - int(bm.group(1))
    age = int(longevity)
    if abs(computed - age) <= 1:
        return None
    return (
        f"Recorded age at death ({age}) does not match the {computed} years "
        f"between the recorded birth and death dates — one of these is "
        f"likely a transcription error."
    )


def _parse_parent_name(raw: str) -> tuple[str | None, str | None] | None:
    """Parse a parent-column value into (given, surname).

    Returns None when the entry is entirely unknown.
    Partial names are returned as (None, surname) or (given, None).

    Examples:
        'James Steele'  → ('James', 'Steele')
        '?  Hunter'     → (None, 'Hunter')
        'Joanna  ?'     → ('Joanna', None)
        '?'             → None
    """
    s = " ".join(raw.split())
    if not s or s == "?":
        return None
    parts = s.split()
    if parts[0] == "?":
        rest = " ".join(p for p in parts[1:] if p != "?").strip()
        return (None, rest or None)
    if parts[-1] == "?":
        given = " ".join(p for p in parts[:-1] if p != "?").strip()
        return (given or None, None)
    return (" ".join(parts[:-1]) or None, parts[-1])


def _build_place(*parts: str) -> str | None:
    cleaned = [p.strip() for p in parts if str(p).strip() and str(p).strip() != "?"]
    return ", ".join(cleaned) or None


def _strip_nee(surname: str) -> str:
    """'SMITH [née JONES]' → 'SMITH'."""
    return re.sub(r"\s*\[.*?\]", "", surname).strip()


def _clean_given(given: str) -> tuple[str, str | None]:
    """Split an editorial bracket annotation out of a given name.

    Square-bracket tags like '[Infant death]', '[Child death]' or '[MISSIONARY]'
    are the genealogist's notes, not part of the name, yet they leak into the
    GEDCOM NAME field ('Mary Ann [Infant death]'). Strip them and return the
    annotation so it can be preserved as a NOTE instead. Parenthetical nicknames
    and disambiguators ('(Harry)', '(No.2)', '(1)') ARE kept inline — they are
    standard genealogical name notation.

    Returns (clean_given, annotation_or_None).
    """
    m = re.search(r"\[(.+?)\]", given)
    if not m:
        return given, None
    clean = re.sub(r"\s*\[.+?\]", "", given).strip()
    return clean, m.group(1).strip()


def _maiden_name(surname: str) -> str | None:
    """Extract the maiden surname from 'MARRIED [née Maiden]'.

    Returns None when no née clause is present or the maiden name is unknown
    ('[née  ? ]').
    """
    m = re.search(r"\[\s*n[ée]e\s+(.*?)\s*\]", surname, re.IGNORECASE)
    if not m:
        return None
    val = " ".join(m.group(1).split())
    if not val or val == "?":
        return None
    return val


def _married_surnames(surname_base: str) -> list[str]:
    """Split a 'PONTING then PETTY' progression into ['PONTING', 'PETTY'].

    A single surname returns a one-element list.
    """
    parts = [p.strip() for p in re.split(r"\s+then\s+", surname_base, flags=re.IGNORECASE)]
    return [p for p in parts if p]


def _infer_sex(surname: str) -> str | None:
    if "née" in surname:
        return "F"
    return None


def _code_role(code: str) -> str:
    """Classify a col-15 code as 'husband', 'wife', 'child', or 'ex_spouse'.

    Codes are slash-separated paths into the tree; the role is decided by the
    *final* path segment so nesting depth doesn't matter:
      'BelAl/Cl/Ar'     → child     (a child of family 'BelAl/Cl')
      'BelAl/Cl/Ar-Dp'  → wife      (spouse who married into family 'BelAl/Cl/Ar')
      'BelAl'           → husband   (a family head)
      'GreJeAds-Ada-EmGe' → ex_spouse (Ada's *prior* husband before Jesse Green)
    A '-suffix' on the last segment marks a married-in spouse and wins over the
    '/' that merely signals depth. A *second* trailing '-suffix' (two hyphens in
    the last segment) marks a prior-marriage chain: the deepest token is the
    earlier spouse of the person named by the rest of the segment. The hyphen
    suffix can follow a digit ('P2-J') or be unknown ('L-?'), so split on the
    hyphen rather than requiring letters on both sides.
    """
    if not code:
        return "unknown"
    last = code.rsplit("/", 1)[-1]
    hyphens = last.count("-")
    if hyphens >= 2:
        return "ex_spouse"
    if hyphens == 1:
        return "wife"
    if "/" in code:
        return "child"
    return "husband"


def _partner_code(code: str) -> str:
    """For a prior-spouse chain, the code of the linking spouse they married.

    'GreJeAds-Ada-EmGe' (George, Ada's prior husband) → 'GreJeAds-Ada' (Ada).
    """
    path, _, last = code.rpartition("/")
    last = last.rsplit("-", 1)[0]
    return f"{path}/{last}" if path else last


def _code_base(code: str) -> str:
    """Return the base code of the family this person *belongs to*.

    A child of 'A/B/C' belongs to the family headed by its immediate parent
    'A/B' (not the top ancestor 'A'). A spouse 'A/B-X' married into family
    'A/B'. Examples:
      'HntJm/Jn'     → 'HntJm'      (child)
      'BelAl/Cl/Ar'  → 'BelAl/Cl'   (child, immediate parent)
      'HntJm-Ca'     → 'HntJm'      (wife)
      'BelAl/Cl/Ar-Dp' → 'BelAl/Cl/Ar' (wife)
    """
    if _code_role(code) == "wife":
        # Strip the spouse suffix from the final path segment.
        path, _, last = code.rpartition("/")
        last_base = last.split("-", 1)[0]
        return f"{path}/{last_base}" if path else last_base
    if "/" in code:
        # Child: belongs to the family headed by its immediate parent.
        return code.rsplit("/", 1)[0]
    return code


def _code_self(code: str) -> str:
    """Return the base code of the family this person *heads*.

    Husbands and children head a family identified by their own code path
    ('BelAl/Cl/Ar' for the child Arthur, whose own children are 'BelAl/Cl/Ar/*').
    Spouses don't head a family of their own; they join their partner's, so
    their self-base is the family they married into.
    """
    if _code_role(code) == "wife":
        return _code_base(code)
    return code


def _dedup_key(given: str, surname_base: str, father_raw: str, birth_raw: Any) -> tuple:
    """Stable key identifying the same person across multiple rows."""
    given_norm = given.strip().lower().split("(")[0].strip()
    surname_norm = surname_base.strip().upper()

    # Father's first name (or '?' if unknown)
    father_first = str(father_raw).strip().split()[0].lower() if str(father_raw).strip() else "?"
    if father_first in ("?", ""):
        father_first = "?"

    # Birth year
    birth_year: int | None = None
    if isinstance(birth_raw, float):
        full = int(birth_raw)
        if full > 9999:
            birth_year = full // 10000
        else:
            birth_year = full
    elif isinstance(birth_raw, str):
        m = re.search(r"\b(\d{4})", birth_raw)
        if m:
            birth_year = int(m.group(1))

    return (surname_norm, given_norm, father_first, birth_year)


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def read_spreadsheet(path: Path) -> tuple[list[Individual], list[Family]]:
    wb = xlrd.open_workbook(str(path))
    ws = wb.sheet_by_index(0)

    def v(r: int, col: int) -> Any:
        return ws.cell(r, col).value

    # Excel cell comments hold the genealogist's research notes (alternate
    # spellings, birth-vs-christening clarifications, sources). They are not
    # part of any cell's value, so gather them per row, ordered by column.
    row_notes: dict[int, list[str]] = {}
    for (nr, nc), note in sorted(ws.cell_note_map.items()):
        text = (note.text or "").strip()
        if text:
            row_notes.setdefault(nr, []).append(text)

    # The header legend (rows above the data) maps the single-letter lineage
    # codes in cols 38-42 to the principal family-line surnames, e.g.
    # "B=Belshaw", "L=Livingstone". Build that lookup so per-person membership
    # codes can be rendered as readable names.
    line_map: dict[str, str] = {}
    for r in range(0, DATA_START_ROW):
        for c in range(_C_LINE_FIRST, ws.ncols):
            m = re.match(r"^([A-Za-z]{1,2})\s*=\s*(.+)$", str(v(r, c)).strip())
            if m:
                line_map.setdefault(m.group(1), m.group(2).strip())

    def _lines_for(r: int) -> set[str]:
        names: set[str] = set()
        for c in range(_C_LINE_FIRST, ws.ncols):
            code = str(v(r, c)).strip()
            if code:
                names.add(line_map.get(code, code))
        return names

    # ------------------------------------------------------------------
    # Pass 1 – collect raw person and marriage rows
    # ------------------------------------------------------------------
    person_rows: list[dict] = []
    marriage_rows: list[dict] = []
    # Given-name-only rows (no generation/code/surname/flag) list extra children
    # under the family above them — e.g. Alexander Livingstone's daughters
    # Stella and Maud. Captured with the base of the preceding coded row.
    loose_child_rows: list[tuple[int, str, str | None]] = []
    last_child_base: str | None = None

    for r in range(DATA_START_ROW, ws.nrows):
        flag = str(v(r, _C_FLAG)).strip()
        generation = v(r, _C_GENERATION)

        if flag == "M":
            combined = str(v(r, _C_GIVEN)).strip()
            marr_notes = list(row_notes.get(r, []))
            mn = _date_precision_note(v(r, _C_DATE1), "Marriage")
            if mn:
                marr_notes.append(mn)
            marriage_rows.append({
                "row": r,
                "combined": combined,
                "date": _parse_date(v(r, _C_DATE1)),
                "place": _build_place(str(v(r, _C_TOWN)), str(v(r, _C_COUNTY))),
                "cell_notes": marr_notes,
            })
        elif generation != "":
            surname = str(v(r, _C_SURNAME)).strip()
            given, given_annotation = _clean_given(str(v(r, _C_GIVEN)).strip())
            father = str(v(r, _C_FATHER)).strip()
            # '|' is an occasional typo for the '/' child separator
            code = str(v(r, _C_CODE)).strip().replace("|", "/")
            surname_base = _strip_nee(surname)

            person_notes = list(row_notes.get(r, []))
            if given_annotation:
                person_notes.append(f"Name annotation: [{given_annotation}].")
            for cell, lbl in ((_C_DATE1, "Birth"), (_C_DEATH_DATE, "Death"),
                              (_C_MARRIAGE, "Marriage")):
                pn = _date_precision_note(v(r, cell), lbl)
                if pn:
                    person_notes.append(pn)
            ld = _longevity_discrepancy_note(
                _parse_date(v(r, _C_DATE1)), _parse_date(v(r, _C_DEATH_DATE)),
                v(r, _C_LONGEVITY))
            if ld:
                person_notes.append(ld)

            person_rows.append({
                "row": r,
                "code": code,
                "role": _code_role(code),
                "base": _code_base(code),
                "father_raw": father,
                "mother_raw": str(v(r, _C_MOTHER)).strip(),
                "surname": surname,
                "surname_base": surname_base,
                "maiden": _maiden_name(surname),
                "given": given,
                "birth_date": _parse_date(v(r, _C_DATE1)),
                "birth_is_chr": flag == "C",
                "birth_place": _build_place(str(v(r, _C_TOWN)), str(v(r, _C_COUNTY))),
                "death_date": _parse_date(v(r, _C_DEATH_DATE)),
                "death_place": str(v(r, _C_BURIED)).strip() or None,
                # Marriage date/place are recorded on each spouse's own row
                # (cols 31/32), not only on the rarer 'M'-flag rows.
                "marr_date": _parse_date(v(r, _C_MARRIAGE)),
                "marr_place": str(v(r, _C_MARRIED_PLACE)).strip() or None,
                "sex": _infer_sex(surname),
                "occupation": str(v(r, _C_OCCUPATION)).strip() or None,
                "notes": str(v(r, _C_NOTES)).strip() or None,
                "cell_notes": person_notes,
                "lines": _lines_for(r),
                "dedup_key": _dedup_key(given, surname_base, father, v(r, _C_DATE1)),
            })
            last_child_base = _code_base(code)
        else:
            given = str(v(r, _C_GIVEN)).strip()
            if (
                given
                and not str(v(r, _C_SURNAME)).strip()
                and not str(v(r, _C_CODE)).strip()
            ):
                loose_child_rows.append((r, given, last_child_base))

    # ------------------------------------------------------------------
    # Pass 2 – deduplicate individuals
    # ------------------------------------------------------------------
    dedup_map: dict[tuple, Individual] = {}
    _id_counter = 0

    def _get_or_create(p: dict) -> Individual:
        nonlocal _id_counter
        dk = p["dedup_key"]
        if dk in dedup_map:
            existing = dedup_map[dk]
            # Merge: fill gaps with data from the second appearance
            if not existing.birth_date and p["birth_date"]:
                existing.birth_date = p["birth_date"]
            if not existing.birth_place and p["birth_place"]:
                existing.birth_place = p["birth_place"]
            if not existing.death_date and p["death_date"]:
                existing.death_date = p["death_date"]
            if not existing.death_place and p["death_place"]:
                existing.death_place = p["death_place"]
            if not existing.occupation and p["occupation"]:
                existing.occupation = p["occupation"]
            if existing.sex is None and p["sex"]:
                existing.sex = p["sex"]
            for n in p["cell_notes"]:
                if n not in existing.note_list:
                    existing.note_list.append(n)
            existing.lineage_lines |= p["lines"]
            return existing
        _id_counter += 1
        # A woman recorded as "PONTING then PETTY [née Richey]" carries two
        # successive married surnames in one cell. Genealogically she is filed
        # under her maiden (birth) surname — which also matches her position in
        # her own family's chart — with the married names preserved separately.
        married = _married_surnames(p["surname_base"])
        if len(married) > 1 and p["maiden"]:
            display_surname = p["maiden"]
            married_surnames = married
        else:
            display_surname = p["surname_base"]
            married_surnames = []
        ind = Individual(
            id=f"I{_id_counter}",
            given_name=p["given"],
            surname=display_surname,
            birth_date=p["birth_date"],
            birth_place=p["birth_place"],
            birth_is_christening=p["birth_is_chr"],
            death_date=p["death_date"],
            death_place=p["death_place"],
            sex=p["sex"],
            occupation=p["occupation"],
            notes=p["notes"],
            note_list=list(p["cell_notes"]),
            lineage_lines=set(p["lines"]),
            married_surnames=married_surnames,
        )
        dedup_map[dk] = ind
        return ind

    for p in person_rows:
        _get_or_create(p)

    # ------------------------------------------------------------------
    # Pass 2b – reconcile the same person recorded in multiple lineage
    # charts. Someone who bridges several family lines (e.g. appears as a
    # child in two different ancestral charts) gets distinct dedup keys
    # because their parent names differ between charts. When the name AND
    # an exact birth date (day + month + year) match, they are certainly
    # the same individual: merge them onto one record and repoint the
    # dedup keys so every chart contributes a parent family. This only
    # ever merges — a shared birth *year* alone is not enough, so same-
    # named cousins are never collapsed.
    # ------------------------------------------------------------------
    def _has_full_date(d: str | None) -> bool:
        return bool(d) and any(mon in d for mon in _MONTHS)

    def _merge_into(dst: Individual, src: Individual) -> None:
        if not dst.birth_place and src.birth_place:
            dst.birth_place = src.birth_place
        if not dst.death_date and src.death_date:
            dst.death_date = src.death_date
        if not dst.death_place and src.death_place:
            dst.death_place = src.death_place
        if not dst.occupation and src.occupation:
            dst.occupation = src.occupation
        if dst.sex is None and src.sex:
            dst.sex = src.sex
        if not dst.notes and src.notes:
            dst.notes = src.notes
        for n in src.note_list:
            if n not in dst.note_list:
                dst.note_list.append(n)

    strong_canon: dict[tuple, Individual] = {}
    for dk, ind in list(dedup_map.items()):
        if not _has_full_date(ind.birth_date):
            continue
        given_norm = (ind.given_name or "").lower().split("(")[0].strip()
        sk = ((ind.surname or "").upper(), given_norm, ind.birth_date)
        canon = strong_canon.get(sk)
        if canon is None:
            strong_canon[sk] = ind
        elif canon is not ind:
            _merge_into(canon, ind)
            dedup_map[dk] = canon

    # ------------------------------------------------------------------
    # Pass 3 – build families using code-based role classification
    #
    # The col-15 code encodes family structure:
    #   'HntJm'    → husband (family head)
    #   'HntJm-Ca' → wife of HntJm family
    #   'HntJm/Jn' → child of HntJm family
    #
    # We scan rows in spreadsheet order. A new 'husband' code always
    # starts a new family. Wives and children are matched by base code
    # to the most recently opened family with that base.
    # ------------------------------------------------------------------
    families: list[Family] = []
    _fam_counter = 0
    # Maps base code → the most recently opened Family with that base
    active_family: dict[str, Family] = {}

    # Maps a code to the individual who heads the family of that code, so an
    # intermediate child (e.g. 'BelAl/Cl/Ar') can anchor a family that its own
    # children ('BelAl/Cl/Ar/Dv') and spouse ('BelAl/Cl/Ar-Dp') attach to.
    code_to_head: dict[str, Individual] = {}
    for p in person_rows:
        if p["role"] in ("husband", "child"):
            code_to_head.setdefault(_code_self(p["code"]), dedup_map[p["dedup_key"]])

    # Literal code → individual, so a prior-spouse chain ('GreJeAds-Ada-EmGe')
    # can find the linking spouse it married ('GreJeAds-Ada').
    full_code_to_ind: dict[str, Individual] = {}
    for p in person_rows:
        full_code_to_ind.setdefault(p["code"], dedup_map[p["dedup_key"]])

    def _open_family(base: str, husband_id: str | None = None) -> Family:
        nonlocal _fam_counter
        if husband_id is None:
            head = code_to_head.get(base)
            husband_id = head.id if head else None
        _fam_counter += 1
        fam = Family(id=f"F{_fam_counter}", husband_id=husband_id, _husband_base=base)
        families.append(fam)
        active_family[base] = fam
        return fam

    def _find_head_code(base: str) -> str | None:
        """Map a family base to a known head code, tolerating abbreviation.

        Codes are sometimes abbreviated inconsistently — Barbara's own code is
        'BelAl/Rg/Do/Ba' but her spouses reference 'BelAl/Rg/Do/B'. Match within
        the *same parent path* and accept a final segment that is a prefix of
        (or prefixed by) the head's, preferring the longest (most specific) one.
        """
        if base in code_to_head:
            return base
        bpath, _, blast = base.rpartition("/")
        best: str | None = None
        for hc in code_to_head:
            hpath, _, hlast = hc.rpartition("/")
            if hpath != bpath:
                continue
            if hlast.startswith(blast) or blast.startswith(hlast):
                if best is None or len(hc) > len(best):
                    best = hc
        return best

    def _family_for(base: str, create: bool = True) -> Family | None:
        """Resolve (and optionally lazily anchor) the family for a base code."""
        fam = active_family.get(base)
        if fam is not None:
            return fam
        hc = _find_head_code(base)
        if hc is None:
            return None
        if hc in active_family:
            active_family[base] = active_family[hc]  # alias the abbreviation
            return active_family[hc]
        if not create:
            return None
        head = code_to_head.get(hc)
        fam = _open_family(hc, husband_id=head.id if head else None)
        if hc != base:
            active_family[base] = fam  # alias so sibling codes resolve here too
        return fam

    # Spouse-id → Individual, for matching marriage rows to families by name.
    id_to_ind = {ind.id: ind for ind in dedup_map.values()}

    def _marriage_tokens(combined: str) -> set[str]:
        return set(re.findall(r"[a-z]+", combined.lower()))

    # Interleave person and marriage rows in row order
    all_events: list[tuple[int, str, dict]] = (
        [(p["row"], "person", p) for p in person_rows]
        + [(m["row"], "marriage", m) for m in marriage_rows]
    )
    all_events.sort(key=lambda x: x[0])

    for _, etype, data in all_events:
        if etype == "marriage":
            # Attach marriage info to the most recently opened family whose
            # husband or wife surname appears in the combined "X m. Y" string.
            # The surname guard prevents a stray marriage row (e.g. from a
            # note block whose people were never read as individuals) from
            # corrupting an unrelated family that merely happens to lack a
            # marriage date.
            tokens = _marriage_tokens(data["combined"])
            for fam in reversed(families):
                if fam.marriage_date is not None:
                    continue
                h = id_to_ind.get(fam.husband_id)
                w = id_to_ind.get(fam.wife_id)
                h_sn = (h.surname or "").lower().strip() if h else ""
                w_sn = (w.surname or "").lower().strip() if w else ""
                if (h_sn and h_sn in tokens) or (w_sn and w_sn in tokens):
                    fam.marriage_date = data["date"]
                    fam.marriage_place = data["place"]
                    fam.note_list.extend(data["cell_notes"])
                    data["attached"] = True
                    break
            continue

        # Person row
        ind = dedup_map[data["dedup_key"]]
        role = data["role"]
        base = data["base"]

        if role == "husband":
            existing = active_family.get(base)
            if existing is not None and existing.husband_id in (None, ind.id):
                # Family already opened — either a spouse appeared first
                # (female-headed entry, husband_id still None) or it was
                # lazily anchored with this same person as head. Reuse it
                # rather than creating a duplicate.
                existing.husband_id = ind.id
            else:
                _open_family(base, husband_id=ind.id)

        elif role == "wife":
            # The family she married into may not have been opened yet — an
            # intermediate head (e.g. child 'BelAl/Cl/Ar') only anchors a family
            # lazily, when its first spouse or child appears. Fall back to a
            # headless family only if no head can be resolved at all.
            fam = _family_for(base) or _open_family(base)
            if fam.wife_id is not None and fam.wife_id != ind.id:
                # A second spouse on the same base is a remarriage — open a
                # parallel family sharing the same head so it isn't dropped.
                fam = _open_family(base, husband_id=fam.husband_id)
            if fam.wife_id is None:
                fam.wife_id = ind.id

        elif role == "child":
            # Anchor the immediate-parent family lazily; abbreviated child codes
            # (e.g. 'PettL' vs family 'PettLuGe') resolve via _find_head_code.
            fam = _family_for(base)
            if fam is not None:
                fam.child_ids.append(ind.id)
            # If no matching family, the child is an unresolved reference — skip.

        elif role == "ex_spouse":
            # A prior-marriage chain: this person married the linking spouse
            # ('GreJeAds-Ada') in an earlier marriage, so they head their own
            # family with that spouse — NOT a member of the later family. Place
            # the partner opposite by sex (a née-female partner ⇒ this is the
            # husband) and let the final sweep fill the remaining sex.
            partner = full_code_to_ind.get(_partner_code(data["code"]))
            if partner is not None:
                if partner.sex == "M":
                    h_id, w_id = ind.id, partner.id
                else:
                    h_id, w_id = partner.id, ind.id
                _fam_counter += 1
                families.append(Family(id=f"F{_fam_counter}",
                                       husband_id=h_id, wife_id=w_id))

    # ------------------------------------------------------------------
    # Post-process: swap husband/wife when the assigned "husband" is
    # demonstrably female (has née in surname). This corrects for families
    # where the female descendant is the tree's connecting person and
    # therefore gets the base code rather than the -suffix code.
    # ------------------------------------------------------------------
    id_to_ind = {ind.id: ind for ind in dedup_map.values()}
    for fam in families:
        if fam.husband_id and fam.wife_id:
            h = id_to_ind[fam.husband_id]
            if h.sex == "F":
                fam.husband_id, fam.wife_id = fam.wife_id, fam.husband_id

    # Attach marriage date/place recorded on the spouses' own rows (cols 31/32).
    # The 'M'-flag marriage rows cover only a fraction of couples; the rest
    # carry their marriage details on each individual's row. Index by spouse id
    # and fill any family still lacking a marriage. Keying on the husband's row
    # keeps remarriages distinct (each husband row belongs to one family); the
    # wife's row is a fallback when the husband has no recorded date.
    row_marr: dict[str, dict] = {}
    for p in person_rows:
        if p["role"] not in ("husband", "wife", "ex_spouse"):
            continue
        if not (p["marr_date"] or p["marr_place"]):
            continue
        row_marr.setdefault(dedup_map[p["dedup_key"]].id, p)
    # A person who married more than once carries a single col-31 value that
    # describes only one of their marriages, so it must not be copied onto every
    # family they belong to (e.g. Henry Ponting's 1884 date leaking onto his
    # first wife Maud's family). Prefer the spouse who appears in the *fewest*
    # families — their col-31 unambiguously refers to this marriage — and break
    # ties husband-first.
    spouse_fam_count: dict[str, int] = {}
    for fam in families:
        for sid in (fam.husband_id, fam.wife_id):
            if sid:
                spouse_fam_count[sid] = spouse_fam_count.get(sid, 0) + 1
    for fam in families:
        if fam.marriage_date or fam.marriage_place:
            continue
        cands = [s for s in (fam.husband_id, fam.wife_id) if s and s in row_marr]
        cands.sort(key=lambda s: (spouse_fam_count.get(s, 0),
                                  0 if s == fam.husband_id else 1))
        if cands:
            src = row_marr[cands[0]]
            fam.marriage_date = src["marr_date"]
            fam.marriage_place = src["marr_place"]

    # ------------------------------------------------------------------
    # Pass 3.5 – loose given-name-only children
    #
    # Attach each bare given-name row to the family it sits under, taking the
    # father's surname (their maiden/paternal name). Skip names that merely
    # repeat an already-coded child (e.g. "Mabel" == "Mabel Annie", LivAx/Mb)
    # to avoid duplicating that individual.
    # ------------------------------------------------------------------
    for row, given, ctx_base in loose_child_rows:
        if not ctx_base:
            continue
        fam = active_family.get(ctx_base)
        if fam is None:
            continue
        existing_givens = {
            (id_to_ind[c].given_name or "").lower().split()[0]
            for c in fam.child_ids
            if c in id_to_ind and id_to_ind[c].given_name
        }
        if given.lower().split()[0] in existing_givens:
            continue
        father = id_to_ind.get(fam.husband_id)
        surname = (father.surname if father else "") or ""
        _id_counter += 1
        child = Individual(id=f"I{_id_counter}", given_name=given, surname=surname)
        dedup_map[("__loose__", row)] = child
        id_to_ind[child.id] = child
        fam.child_ids.append(child.id)

    # ------------------------------------------------------------------
    # Pass 4 – Synthetic parents for family-head rows
    #
    # Husband and wife rows (no '/' in code) carry their parents' names
    # in columns Q/R but are never assigned FAMC links by the code-based
    # Pass 3. Here we parse those names, match to existing individuals
    # where possible, create synthetic Individual records where not, and
    # wire up parent Family records accordingly.
    # ------------------------------------------------------------------
    def _ind_name_key(given: str | None, surname: str | None) -> tuple[str, str]:
        g_str = (given or "").lower().split("(")[0].strip()
        g = g_str.split()[0] if g_str else ""
        return ((surname or "").upper(), g)

    # Name-keyed lookup across all real individuals (first match wins)
    existing_by_name: dict[tuple[str, str], Individual] = {}
    for ind in dedup_map.values():
        k = _ind_name_key(ind.given_name or None, ind.surname or None)
        if k != ("", ""):
            existing_by_name.setdefault(k, ind)

    synth_by_name: dict[tuple[str, str], Individual] = {}
    parent_fams: dict[tuple[str | None, str | None], Family] = {}
    already_child: set[str] = {cid for fam in families for cid in fam.child_ids}

    def _descendants(root_id: str | None) -> set[str]:
        """Ids reachable as descendants of root_id through current families."""
        if root_id is None:
            return set()
        seen: set[str] = set()
        stack = [root_id]
        while stack:
            cur = stack.pop()
            for fam in families:
                if cur in (fam.husband_id, fam.wife_id):
                    for c in fam.child_ids:
                        if c not in seen:
                            seen.add(c)
                            stack.append(c)
        return seen

    def _resolve_parent(
        parsed: tuple[str | None, str | None] | None,
        sex: str,
        avoid: frozenset[str] = frozenset(),
    ) -> Individual | None:
        nonlocal _id_counter
        if parsed is None or (parsed[0] is None and parsed[1] is None):
            return None
        key = _ind_name_key(parsed[0], parsed[1])
        parent = existing_by_name.get(key) or synth_by_name.get(key)
        # A loose name match (surname + first given word) can collide with a
        # descendant of the very person we are giving parents to — e.g. Bruce
        # Dallas's father "William H. Dallas" matching his grandson "William
        # John Peter Dallas". Such a match would invert the tree, so reject it
        # and mint a distinct ancestor instead.
        if parent is not None and parent.id in avoid:
            parent = None
        if parent is None:
            _id_counter += 1
            parent = Individual(
                id=f"I{_id_counter}",
                given_name=parsed[0] or "",
                surname=parsed[1] or "",
                sex=sex,
            )
            synth_by_name[key] = parent
        return parent

    for _, etype, data in all_events:
        if etype != "person" or data["role"] not in ("husband", "wife"):
            continue
        ind = dedup_map[data["dedup_key"]]
        if ind.id in already_child:
            continue

        fp = _parse_parent_name(data["father_raw"])
        mp = _parse_parent_name(data["mother_raw"])
        if fp is None and mp is None:
            continue

        avoid = frozenset(_descendants(ind.id) | {ind.id})
        father_ind = _resolve_parent(fp, "M", avoid)
        mother_ind = _resolve_parent(mp, "F", avoid)

        if father_ind is None and mother_ind is None:
            continue

        pkey = (
            father_ind.id if father_ind else None,
            mother_ind.id if mother_ind else None,
        )
        if pkey not in parent_fams:
            _fam_counter += 1
            pf = Family(id=f"F{_fam_counter}", husband_id=pkey[0], wife_id=pkey[1])
            families.append(pf)
            parent_fams[pkey] = pf
        parent_fams[pkey].child_ids.append(ind.id)
        already_child.add(ind.id)

    # ------------------------------------------------------------------
    # Pass 5 – flag-based event blocks
    #
    # Some ancestry is recorded not with generation numbers and codes but as
    # free-standing B (birth) / D (death) / M (marriage) event rows — e.g. the
    # Forster line behind Maud Ponting (née Forster), who is already in the
    # main tree. Merge each person's B and D rows, recover birth/death/notes,
    # and link them via the parent-name columns, reusing Pass-4 resolution so
    # block people unify with the synthetic parents the main tree already
    # implied (e.g. Maud's father "John Forster") instead of duplicating them.
    # ------------------------------------------------------------------
    def _parent_family(fid: str | None, mid: str | None) -> Family:
        nonlocal _fam_counter
        # Reuse a half-known couple (father-only or mother-only) rather than
        # splitting a sibling group across two families.
        for cand in ((fid, mid), (fid, None), (None, mid)):
            if cand == (None, None):
                continue
            fam = parent_fams.get(cand)
            if fam is not None:
                if fid and fam.husband_id is None:
                    fam.husband_id = fid
                if mid and fam.wife_id is None:
                    fam.wife_id = mid
                parent_fams.pop(cand, None)
                parent_fams[(fam.husband_id, fam.wife_id)] = fam
                return fam
        _fam_counter += 1
        fam = Family(id=f"F{_fam_counter}", husband_id=fid, wife_id=mid)
        families.append(fam)
        parent_fams[(fid, mid)] = fam
        return fam

    block_people: dict[tuple, dict] = {}
    block_order: list[tuple] = []
    for r in range(DATA_START_ROW, ws.nrows):
        if str(v(r, _C_FLAG)).strip() not in ("B", "D"):
            continue
        if v(r, _C_GENERATION) != "":
            continue
        surname = str(v(r, _C_SURNAME)).strip()
        if not surname:
            continue
        flag = str(v(r, _C_FLAG)).strip()
        # Strip clarifying annotations like "John {Maud's father}".
        given = re.sub(r"\s*\{.*?\}", "", str(v(r, _C_GIVEN)).strip()).strip()
        given, given_annotation = _clean_given(given)
        surname_base = _strip_nee(surname)
        father_raw = str(v(r, _C_FATHER)).strip()
        mother_raw = str(v(r, _C_MOTHER)).strip()
        f_first = father_raw.split()[0].lower() if father_raw and father_raw != "?" else "?"
        m_first = mother_raw.split()[0].lower() if mother_raw and mother_raw != "?" else "?"
        # Parents distinguish same-named people (e.g. two different "John"s).
        key = (surname_base.upper(), given.lower(), f_first, m_first)
        rec = block_people.get(key)
        if rec is None:
            rec = {
                "given": given, "surname_base": surname_base,
                "father_raw": father_raw, "mother_raw": mother_raw,
                "birth": None, "death": None, "death_place": None,
                "sex": _infer_sex(surname), "notes": [],
            }
            block_people[key] = rec
            block_order.append(key)
        block_notes = list(row_notes.get(r, []))
        if given_annotation:
            block_notes.append(f"Name annotation: [{given_annotation}].")
        pn = _date_precision_note(v(r, _C_DATE1), "Birth" if flag == "B" else "Death")
        if pn:
            block_notes.append(pn)
        for n in block_notes:
            if n not in rec["notes"]:
                rec["notes"].append(n)
        date = _parse_date(v(r, _C_DATE1))
        if flag == "B":
            rec["birth"] = rec["birth"] or date
        else:
            rec["death"] = rec["death"] or date
            place = str(v(r, _C_TOWN)).strip()
            if place and not rec["death_place"]:
                rec["death_place"] = place

    # Full-given-name index over every individual built so far, used to unify
    # a block person with the record the main tree already implied (e.g. Maud,
    # or her father "John Forster"). Keyed on the *whole* given name so
    # "John T." stays distinct from "John".
    def _full_key(given: str | None, surname: str | None) -> tuple[str, str]:
        return ((surname or "").upper(), (given or "").lower().split("(")[0].strip())

    full_index: dict[tuple[str, str], Individual] = {}
    for ind in list(dedup_map.values()) + list(synth_by_name.values()):
        full_index.setdefault(_full_key(ind.given_name, ind.surname), ind)

    # Materialise each block person, unifying with existing/synthetic records.
    # Same full name + different parents (e.g. John Forster the father vs John
    # Forster the infant brother) must stay distinct: only the first occurrence
    # of a given name unifies with a pre-existing record; later ones are new.
    block_created: list[Individual] = []
    block_fullname_seen: set[tuple[str, str]] = set()
    for key in block_order:
        rec = block_people[key]
        nk = _ind_name_key(rec["given"], rec["surname_base"])
        fk = _full_key(rec["given"], rec["surname_base"])
        ind = None if fk in block_fullname_seen else full_index.get(fk)
        block_fullname_seen.add(fk)
        if ind is None:
            _id_counter += 1
            ind = Individual(
                id=f"I{_id_counter}", given_name=rec["given"],
                surname=rec["surname_base"], sex=rec["sex"],
            )
            block_created.append(ind)
        # Pin the loose name keys to the first (primary) holder so a later
        # same-named block person can't redirect parent lookups to itself.
        synth_by_name.setdefault(nk, ind)
        existing_by_name.setdefault(nk, ind)
        full_index.setdefault(fk, ind)
        if not ind.birth_date and rec["birth"]:
            ind.birth_date = rec["birth"]
        if not ind.death_date and rec["death"]:
            ind.death_date = rec["death"]
        if not ind.death_place and rec["death_place"]:
            ind.death_place = rec["death_place"]
        if ind.sex is None and rec["sex"]:
            ind.sex = rec["sex"]
        for n in rec["notes"]:
            if n not in ind.note_list:
                ind.note_list.append(n)
        rec["ind"] = ind
        # Married women are referenced as a mother by given name only; register
        # a given-name alias so those parent lookups unify with this record.
        if rec["sex"] == "F" and nk[1]:
            synth_by_name.setdefault(("", nk[1]), ind)

    # Link block people to their parents.
    for key in block_order:
        rec = block_people[key]
        ind = rec["ind"]
        if ind.id in already_child:
            continue
        fp = _parse_parent_name(rec["father_raw"])
        mp = _parse_parent_name(rec["mother_raw"])
        if fp is None and mp is None:
            continue
        avoid = frozenset(_descendants(ind.id) | {ind.id})
        father_ind = _resolve_parent(fp, "M", avoid)
        mother_ind = _resolve_parent(mp, "F", avoid)
        if father_ind is None and mother_ind is None:
            continue
        fam = _parent_family(
            father_ind.id if father_ind else None,
            mother_ind.id if mother_ind else None,
        )
        fam.child_ids.append(ind.id)
        already_child.add(ind.id)

    # Attach the block's marriage rows (skipped by Pass 3 as no family existed
    # for them then) to the now-built couples, matching both spouses by name.
    all_ind_by_id = {
        i.id: i for i in list(dedup_map.values()) + list(synth_by_name.values())
    }

    def _spouse_matches(ind: Individual | None, tokens: set[str]) -> bool:
        if ind is None:
            return False
        g = (ind.given_name or "").lower().split("(")[0].split()
        s = (ind.surname or "").lower().strip()
        first = g[0] if g else ""
        if s and s not in tokens:
            return False
        if first and first not in tokens:
            return False
        return bool(s or first)

    for m in marriage_rows:
        if m.get("attached") or (m["date"] is None and m["place"] is None):
            continue
        tokens = set(re.findall(r"[a-z]+", m["combined"].lower()))
        for fam in families:
            if fam.marriage_date is not None:
                continue
            h = all_ind_by_id.get(fam.husband_id)
            w = all_ind_by_id.get(fam.wife_id)
            if h and w and _spouse_matches(h, tokens) and _spouse_matches(w, tokens):
                fam.marriage_date = m["date"]
                fam.marriage_place = m["place"]
                fam.note_list.extend(m["cell_notes"])
                m["attached"] = True
                break

    # Pass 2b can point several dedup keys at one merged individual, so
    # de-duplicate the final list by identity while preserving order.
    seen_ids: set[str] = set()
    ordered: list[Individual] = []
    for ind in list(dedup_map.values()) + list(synth_by_name.values()) + block_created:
        if ind.id not in seen_ids:
            seen_ids.add(ind.id)
            ordered.append(ind)

    # A person's role in a family establishes their sex: the husband is male,
    # the wife female. Previously sex was only inferred from a "née" surname,
    # leaving every family-head husband with no SEX. Fill the gap across all
    # families (Pass 3 plus the synthetic/block parents from Passes 4-5),
    # without overriding an explicit inference.
    final_by_id = {ind.id: ind for ind in ordered}
    for fam in families:
        h = final_by_id.get(fam.husband_id)
        if h is not None and h.sex is None:
            h.sex = "M"
        w = final_by_id.get(fam.wife_id)
        if w is not None and w.sex is None:
            w.sex = "F"

    # Emit the married-surname progression as a note once per person (after all
    # merges have unioned the data). Lineage-chart membership stays on
    # Individual.lineage_lines and is written as custom _GROUP tags by the
    # writer, rather than as a repetitive freeform note.
    for ind in ordered:
        if ind.married_surnames:
            ind.note_list.insert(
                0, "Married surname" + ("s" if len(ind.married_surnames) > 1 else "")
                + ": " + ", then ".join(ind.married_surnames) + ".")
    return ordered, families
