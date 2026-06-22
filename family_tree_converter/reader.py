"""Read individuals and relationships from the source Excel spreadsheet."""

from dataclasses import dataclass, field
from pathlib import Path

import openpyxl


@dataclass
class Individual:
    id: str
    given_name: str
    surname: str
    birth_date: str | None = None
    birth_place: str | None = None
    death_date: str | None = None
    death_place: str | None = None
    sex: str | None = None  # "M", "F", or None


@dataclass
class Family:
    id: str
    husband_id: str | None = None
    wife_id: str | None = None
    marriage_date: str | None = None
    marriage_place: str | None = None
    child_ids: list[str] = field(default_factory=list)


def read_spreadsheet(path: Path) -> tuple[list[Individual], list[Family]]:
    """Parse the source Excel file and return individuals and families.

    The exact column layout will be determined once the spreadsheet format
    is understood. This is a placeholder implementation.
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    individuals: list[Individual] = []
    families: list[Family] = []

    # TODO: implement once spreadsheet format is known
    _ = ws  # suppress unused warning until format is defined

    return individuals, families
