"""Read individuals and relationships from the source Excel spreadsheet."""

from __future__ import annotations

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
_C_OCCUPATION = 33
_C_NOTES = 34

DATA_START_ROW = 17

_MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN",
           "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]


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


@dataclass
class Family:
    id: str
    husband_id: str | None = None
    wife_id: str | None = None
    marriage_date: str | None = None
    marriage_place: str | None = None
    child_ids: list[str] = field(default_factory=list)
    # Internal: base code of the husband row, used during construction
    _husband_base: str = field(default="", repr=False, compare=False)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

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
            # Year-only stored as plain integer (e.g. 1760.0)
            return str(date_int)
        if 1 <= month <= 12:
            if day > 0:
                return f"{day} {_MONTHS[month - 1]} {year}"
            return f"{_MONTHS[month - 1]} {year}"
        return str(year)
    s = str(val).strip()
    return s or None


def _build_place(*parts: str) -> str | None:
    cleaned = [p.strip() for p in parts if str(p).strip() and str(p).strip() != "?"]
    return ", ".join(cleaned) or None


def _strip_nee(surname: str) -> str:
    """'SMITH [née JONES]' → 'SMITH'."""
    return re.sub(r"\s*\[.*?\]", "", surname).strip()


def _infer_sex(surname: str) -> str | None:
    if "née" in surname:
        return "F"
    return None


def _code_role(code: str) -> str:
    """Classify a col-15 code as 'husband', 'wife', or 'child'."""
    if not code:
        return "unknown"
    if "/" in code:
        return "child"
    if re.search(r"[A-Za-z]-[A-Za-z]", code):
        return "wife"
    return "husband"


def _code_base(code: str) -> str:
    """Extract family base from a code: 'HntJm/Jn' → 'HntJm', 'HntJm-Ca' → 'HntJm'."""
    if "/" in code:
        return code.split("/")[0]
    m = re.match(r"^([A-Za-z]+)-[A-Za-z]", code)
    if m:
        return m.group(1)
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

    # ------------------------------------------------------------------
    # Pass 1 – collect raw person and marriage rows
    # ------------------------------------------------------------------
    person_rows: list[dict] = []
    marriage_rows: list[dict] = []

    for r in range(DATA_START_ROW, ws.nrows):
        flag = str(v(r, _C_FLAG)).strip()
        generation = v(r, _C_GENERATION)

        if flag == "M":
            combined = str(v(r, _C_GIVEN)).strip()
            marriage_rows.append({
                "row": r,
                "combined": combined,
                "date": _parse_date(v(r, _C_DATE1)),
                "place": _build_place(str(v(r, _C_TOWN)), str(v(r, _C_COUNTY))),
            })
        elif generation != "" and generation != 0:
            surname = str(v(r, _C_SURNAME)).strip()
            given = str(v(r, _C_GIVEN)).strip()
            father = str(v(r, _C_FATHER)).strip()
            code = str(v(r, _C_CODE)).strip()
            surname_base = _strip_nee(surname)

            person_rows.append({
                "row": r,
                "code": code,
                "role": _code_role(code),
                "base": _code_base(code),
                "father_raw": father,
                "mother_raw": str(v(r, _C_MOTHER)).strip(),
                "surname": surname,
                "surname_base": surname_base,
                "given": given,
                "birth_date": _parse_date(v(r, _C_DATE1)),
                "birth_is_chr": flag == "C",
                "birth_place": _build_place(str(v(r, _C_TOWN)), str(v(r, _C_COUNTY))),
                "death_date": _parse_date(v(r, _C_DEATH_DATE)),
                "death_place": str(v(r, _C_BURIED)).strip() or None,
                "sex": _infer_sex(surname),
                "occupation": str(v(r, _C_OCCUPATION)).strip() or None,
                "notes": str(v(r, _C_NOTES)).strip() or None,
                "dedup_key": _dedup_key(given, surname_base, father, v(r, _C_DATE1)),
            })

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
            return existing
        _id_counter += 1
        ind = Individual(
            id=f"I{_id_counter}",
            given_name=p["given"],
            surname=p["surname_base"],
            birth_date=p["birth_date"],
            birth_place=p["birth_place"],
            birth_is_christening=p["birth_is_chr"],
            death_date=p["death_date"],
            death_place=p["death_place"],
            sex=p["sex"],
            occupation=p["occupation"],
            notes=p["notes"],
        )
        dedup_map[dk] = ind
        return ind

    for p in person_rows:
        _get_or_create(p)

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

    # Interleave person and marriage rows in row order
    all_events: list[tuple[int, str, dict]] = (
        [(p["row"], "person", p) for p in person_rows]
        + [(m["row"], "marriage", m) for m in marriage_rows]
    )
    all_events.sort(key=lambda x: x[0])

    for _, etype, data in all_events:
        if etype == "marriage":
            # Attach marriage info to the most recently opened family whose
            # husband name appears in the combined string.
            combined = data["combined"]
            # Try to match by scanning active families (most recent first)
            for fam in reversed(families):
                if fam.marriage_date is None:
                    fam.marriage_date = data["date"]
                    fam.marriage_place = data["place"]
                    break
            continue

        # Person row
        ind = dedup_map[data["dedup_key"]]
        role = data["role"]
        base = data["base"]

        if role == "husband":
            existing = active_family.get(base)
            if existing is not None and existing.husband_id is None:
                # Spouse appeared before family-head (female-headed entry) — reuse
                existing.husband_id = ind.id
                fam = existing
            else:
                _fam_counter += 1
                fam = Family(id=f"F{_fam_counter}", husband_id=ind.id, _husband_base=base)
                families.append(fam)
                active_family[base] = fam

        elif role == "wife":
            fam = active_family.get(base)
            if fam is None:
                # Wife with no preceding husband row — create a family
                _fam_counter += 1
                fam = Family(id=f"F{_fam_counter}", _husband_base=base)
                families.append(fam)
                active_family[base] = fam
            if fam.wife_id is None:
                fam.wife_id = ind.id

        elif role == "child":
            fam = active_family.get(base)
            if fam is None:
                # Try prefix match: child codes are sometimes abbreviated vs parent base
                # e.g. child base 'PettL' matches family base 'PettLuGe'
                for fb, f in reversed(list(active_family.items())):
                    if fb.startswith(base) or base.startswith(fb):
                        fam = f
                        break
            if fam is not None:
                fam.child_ids.append(ind.id)
            # If no matching family, the child is an unresolved reference — skip.

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

    return list(dedup_map.values()), families
