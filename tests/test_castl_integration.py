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


def test_nickname_extracted_to_field(parsed):
    individuals, _ = parsed
    edith = _find(individuals, "Edith Rosetta", "STEELE")
    assert edith and edith[0].nickname == "Edie or Cissy"
    assert "(" not in (edith[0].given_name or "")


def test_maiden_name_annotation_stripped_from_surname(parsed):
    individuals, _ = parsed
    jan = [i for i in individuals if i.given_name == "Janette"
           and (i.surname or "").upper() == "STEELE"]
    assert jan, "Janette Steele not found / surname not cleaned"
    assert "(" not in jan[0].surname


def test_divorce_marker_emits_div_on_family(parsed):
    individuals, families = parsed
    # The 8 col-5 'Dv' markers become divorce events on their one family each.
    assert sum(1 for f in families if f.divorced) == 8
    # Elizabeth Ann Morrison (Sp #1, Dv) — her marriage to Noel is divorced,
    # his second (to Lee) is not.
    by_id = {i.id: i for i in individuals}
    noel = _find(individuals, "Noel", "MORRISON")[0]
    noel_fams = [f for f in families if noel.id in (f.husband_id, f.wife_id)]
    divorced = [f for f in noel_fams if f.divorced]
    assert len(divorced) == 1
    wife = by_id.get(divorced[0].wife_id)
    assert wife and "Elizabeth" in (wife.given_name or "")


def test_twin_and_prev_divorced_notes(parsed):
    individuals, _ = parsed
    notes = "\n".join(n for i in individuals for n in i.note_list)
    assert notes.count("Recorded as a twin.") == 4
    assert notes.count("Recorded as previously divorced.") == 1


def test_child_parent_chain_linked(parsed):
    individuals, families = parsed
    by_id = {i.id: i for i in individuals}
    clive = _find(individuals, "Clive Norman", "STEELE")[0]
    parents = next((f for f in families if clive.id in f.child_ids), None)
    assert parents is not None
    father = by_id.get(parents.husband_id)
    assert father is not None and (father.surname or "").upper() == "STEELE"
