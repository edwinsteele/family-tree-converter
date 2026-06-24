"""Integration regression tests for the "C & A Stl" tree (multi-tree merge).

This file uses the same alpha-code convention as the reference spreadsheet but a
different, compact column layout: code in col 6, a single birthplace column, no
flag column, and a blank generation column (so person rows are recognised by
their code). These tests guard that the CASTL_PROFILE keeps reading it correctly.

The spreadsheet is local-only (gitignored), so the tests skip when absent.
"""

from pathlib import Path

import pytest

from family_tree_converter.reader import profile_for, read_spreadsheet
from family_tree_converter.validate import validate

_SOURCE = Path(__file__).resolve().parent.parent / "data" / "input" / "C & A Stl H.Tree #81"


@pytest.fixture(scope="module")
def parsed():
    if not _SOURCE.exists():
        pytest.skip(f"source spreadsheet not present: {_SOURCE}")
    return read_spreadsheet(_SOURCE, profile_for(_SOURCE))


def _find(individuals, given_sub, surname):
    return [
        i for i in individuals
        if surname.lower() in (i.surname or "").lower()
        and given_sub.lower() in (i.given_name or "").lower()
    ]


def test_profile_selected_by_name():
    assert profile_for(_SOURCE).name == "C & A Stl"


def test_counts_and_integrity(parsed):
    individuals, families = parsed
    # Fully-coded file: 178 coded persons + synthetic parents.
    assert len(individuals) > 190
    assert len(families) > 60
    result = validate(individuals, families)
    assert result["errors"] == []


def test_coded_person_read_despite_blank_generation(parsed):
    # The generation column is blank for every row; people must still be read
    # via their path code. Clive Norman Steele is one such row.
    individuals, _ = parsed
    clive = _find(individuals, "Clive Norman", "STEELE")
    assert clive, "Clive Norman Steele not parsed"
    assert clive[0].birth_date == "26 APR 1916"
    assert clive[0].birth_place and "St.Peters" in clive[0].birth_place


def test_single_place_column_and_sex(parsed):
    individuals, _ = parsed
    edna = _find(individuals, "Edna May", "STEELE")
    assert edna and edna[0].sex == "F"
    # Single combined birthplace column read straight through.
    assert edna[0].birth_place and "Richmond" in edna[0].birth_place


def test_child_parent_chain_linked(parsed):
    individuals, families = parsed
    by_id = {i.id: i for i in individuals}
    clive = _find(individuals, "Clive Norman", "STEELE")[0]
    parents = next((f for f in families if clive.id in f.child_ids), None)
    assert parents is not None
    father = by_id.get(parents.husband_id)
    assert father is not None and (father.surname or "").upper() == "STEELE"
