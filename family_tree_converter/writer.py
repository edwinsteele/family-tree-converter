"""Write GEDCOM 5.5.1 output from parsed individuals and families."""

import re
from collections import defaultdict
from pathlib import Path

from .reader import Family, Individual

# GEDCOM 5.5.1 limits a physical line to 255 chars. Wrap note payloads well
# under that (accounting for the level/tag prefix and any multi-byte chars).
_NOTE_WRAP = 200


def _note_lines(text: str) -> list[str]:
    """Serialise one note as GEDCOM NOTE/CONC/CONT lines.

    Embedded line breaks become CONT; overlong logical lines are split with
    CONC so no physical line exceeds the GEDCOM length limit.
    """
    out: list[str] = []
    logical_lines = re.split(r"\r\n|\r|\n", text)
    for li, logical in enumerate(logical_lines):
        tag = "NOTE" if li == 0 else "CONT"
        level = "1" if li == 0 else "2"
        if logical == "":
            out.append(f"{level} {tag}")
            continue
        chunk, rest = logical[:_NOTE_WRAP], logical[_NOTE_WRAP:]
        out.append(f"{level} {tag} {chunk}")
        while rest:
            out.append(f"2 CONC {rest[:_NOTE_WRAP]}")
            rest = rest[_NOTE_WRAP:]
    return out

_HEADER = """\
0 HEAD
1 SOUR family-tree-converter
1 GEDC
2 VERS 5.5.1
2 FORM LINEAGE-LINKED
1 CHAR UTF-8"""

_TRAILER = "0 TRLR"


def write_gedcom(
    individuals: list[Individual],
    families: list[Family],
    output_path: Path,
) -> None:
    # Build back-reference maps required by GEDCOM 5.5.1
    fams: dict[str, list[str]] = defaultdict(list)  # ind_id → [fam_ids as spouse]
    famc: dict[str, list[str]] = defaultdict(list)  # ind_id → [fam_ids as child]

    for fam in families:
        for spouse_id in (fam.husband_id, fam.wife_id):
            if spouse_id:
                fams[spouse_id].append(fam.id)
        for child_id in fam.child_ids:
            famc[child_id].append(fam.id)

    lines: list[str] = [_HEADER]

    for ind in individuals:
        lines.append(f"0 @{ind.id}@ INDI")
        lines.append(f"1 NAME {ind.given_name} /{ind.surname}/")
        if ind.sex:
            lines.append(f"1 SEX {ind.sex}")

        # Principal-lineage chart membership as a custom (underscore) tag, the
        # GEDCOM convention for non-standard data: filterable, preserved by
        # tolerant importers, safely ignored by strict ones.
        for line_name in sorted(ind.lineage_lines):
            lines.append(f"1 _GROUP {line_name}")

        if ind.birth_date or ind.birth_place:
            tag = "CHR" if ind.birth_is_christening else "BIRT"
            lines.append(f"1 {tag}")
            if ind.birth_date:
                lines.append(f"2 DATE {ind.birth_date}")
            if ind.birth_place:
                lines.append(f"2 PLAC {ind.birth_place}")

        if ind.death_date or ind.death_place:
            lines.append("1 DEAT")
            if ind.death_date:
                lines.append(f"2 DATE {ind.death_date}")
        if ind.death_place:
            lines.append("1 BURI")
            lines.append(f"2 PLAC {ind.death_place}")

        if ind.occupation:
            lines.append(f"1 OCCU {ind.occupation}")
        if ind.notes:
            lines.extend(_note_lines(ind.notes))
        for note in ind.note_list:
            lines.extend(_note_lines(note))

        for fam_id in fams.get(ind.id, []):
            lines.append(f"1 FAMS @{fam_id}@")
        for fam_id in famc.get(ind.id, []):
            lines.append(f"1 FAMC @{fam_id}@")

    for fam in families:
        lines.append(f"0 @{fam.id}@ FAM")
        if fam.husband_id:
            lines.append(f"1 HUSB @{fam.husband_id}@")
        if fam.wife_id:
            lines.append(f"1 WIFE @{fam.wife_id}@")
        if fam.marriage_date or fam.marriage_place:
            lines.append("1 MARR")
            if fam.marriage_date:
                lines.append(f"2 DATE {fam.marriage_date}")
            if fam.marriage_place:
                lines.append(f"2 PLAC {fam.marriage_place}")
        for child_id in fam.child_ids:
            lines.append(f"1 CHIL @{child_id}@")
        for note in fam.note_list:
            lines.extend(_note_lines(note))

    lines.append(_TRAILER)
    output_path.write_text("\n".join(lines), encoding="utf-8")
