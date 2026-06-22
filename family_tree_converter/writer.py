"""Write GEDCOM output from parsed individuals and families."""

from pathlib import Path

from .reader import Family, Individual

GEDCOM_HEADER = """\
0 HEAD
1 SOUR family-tree-converter
1 GEDC
2 VERS 5.5.1
2 FORM LINEAGE-LINKED
1 CHAR UTF-8
"""

GEDCOM_TRAILER = "0 TRLR\n"


def write_gedcom(
    individuals: list[Individual],
    families: list[Family],
    output_path: Path,
) -> None:
    lines: list[str] = [GEDCOM_HEADER]

    for ind in individuals:
        lines.append(f"0 @{ind.id}@ INDI")
        lines.append(f"1 NAME {ind.given_name} /{ind.surname}/")
        if ind.sex:
            lines.append(f"1 SEX {ind.sex}")
        if ind.birth_date or ind.birth_place:
            lines.append("1 BIRT")
            if ind.birth_date:
                lines.append(f"2 DATE {ind.birth_date}")
            if ind.birth_place:
                lines.append(f"2 PLAC {ind.birth_place}")
        if ind.death_date or ind.death_place:
            lines.append("1 DEAT")
            if ind.death_date:
                lines.append(f"2 DATE {ind.death_date}")
            if ind.death_place:
                lines.append(f"2 PLAC {ind.death_place}")

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

    lines.append(GEDCOM_TRAILER)
    output_path.write_text("\n".join(lines), encoding="utf-8")
