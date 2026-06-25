"""Tests for the hardening reports (family_tree_converter.checks).

The pure functions (name collisions) are tested on hand-built individuals; the
diagnostics-driven checks (row coverage, synthetic/placeholder manifest,
generation consistency) are tested both on synthetic diagnostics and against the
real Hcks tree, which exercises every name-linked code path.
"""

from pathlib import Path

import pytest

from family_tree_converter import checks
from family_tree_converter.reader import Family, Individual, profile_for, read_spreadsheet

_HCKS = Path(__file__).resolve().parent.parent / "data" / "input" / "Hcks:Thos:Krsl H.Tr.#120"
_BLSGRN = Path(__file__).resolve().parent.parent / "data" / "input" / "BlsGrnLivMcCl H.Tr. #305.xls"


# --- name collisions (pure) ---
def test_name_collision_groups_by_surname_and_first_token():
    inds = [
        Individual(id="I1", given_name="Ernest", surname="Luke", birth_date="1880"),
        Individual(id="I2", given_name="Ernest George", surname="Luke", birth_date="1909"),
        Individual(id="I3", given_name="Mary", surname="Luke"),
    ]
    lines = checks.name_collisions(inds)
    assert len(lines) == 1
    assert "ernest /LUKE/ ×2" in lines[0]
    assert "I1" in lines[0] and "I2" in lines[0]


def test_name_collision_ignores_unique_names():
    inds = [
        Individual(id="I1", given_name="Ada", surname="Green"),
        Individual(id="I2", given_name="Bob", surname="Green"),
    ]
    assert checks.name_collisions(inds) == []


# --- row coverage (synthetic diagnostics) ---
def test_row_coverage_reports_unconsumed_named_rows():
    diag = {
        "row_class": {17: "gen-person", 18: "no-gen", 19: "marriage", 20: "blank/layout"},
        "row_text": {17: ("LUKE", "Alfred"), 18: ("LUKE", "Jack"), 19: ("", "A m. B")},
        "consumed_rows": {17},
    }
    tally, unaccounted = checks.row_coverage(diag)
    assert tally["gen-person"] == 1
    # row 18 bears a name and no pass consumed it -> flagged; 19/20 excluded.
    assert len(unaccounted) == 1
    assert "row 18" in unaccounted[0] and "no-gen" in unaccounted[0]


# --- generation consistency (synthetic diagnostics) ---
def test_generation_consistency_flags_off_by_more_than_one():
    inds = [
        Individual(id="P", given_name="Ernest George", surname="Luke"),
        Individual(id="C", given_name="Lillian", surname="Morgan"),
    ]
    fams = [Family(id="F1", husband_id="P", child_ids=["C"])]
    diag = {
        "generation_by_id": {"P": 9.0, "C": 9.0},  # parent should be child+1
        "name_linked_family_ids": {"F1"},
    }
    issues = checks.generation_consistency(diag, inds, fams)
    assert len(issues) == 1
    assert "F1" in issues[0]

    diag["generation_by_id"]["P"] = 8.0  # now consistent (8 == 7+1? no: child 9)
    diag["generation_by_id"]["C"] = 7.0
    assert checks.generation_consistency(diag, inds, fams) == []


def test_generation_consistency_skips_coded_families():
    inds = [Individual(id="P", given_name="A", surname="X"),
            Individual(id="C", given_name="B", surname="X")]
    fams = [Family(id="F1", husband_id="P", child_ids=["C"])]
    diag = {"generation_by_id": {"P": 5.0, "C": 5.0}, "name_linked_family_ids": set()}
    # F1 is not name-linked, so its (deliberately wrong) gens are not checked.
    assert checks.generation_consistency(diag, inds, fams) == []


# --- synthetic / placeholder manifest (synthetic diagnostics) ---
def test_synthetic_manifest_lists_context():
    inds = [
        Individual(id="I1", given_name="James", surname="Steele"),
        Individual(id="I2", given_name="Charles", surname="Steele"),
        Individual(id="I3", given_name="[Unnamed]", surname="Luke"),
    ]
    fams = [
        Family(id="F1", husband_id="I1", child_ids=["I2"]),
        Family(id="F2", husband_id="I2", child_ids=["I3"]),
    ]
    diag = {"synthetic_ids": {"I1"}, "placeholder_ids": {"I3"}}
    lines = checks.synthetic_manifest(diag, inds, fams)
    assert any("SYNTHETIC PARENT" in line and "James" in line for line in lines)
    assert any("PLACEHOLDER" in line and "Charles" in line for line in lines)


# --- against the real Hcks tree ---
@pytest.fixture(scope="module")
def hcks():
    if not _HCKS.exists():
        pytest.skip(f"source spreadsheet not present: {_HCKS}")
    diag: dict = {}
    individuals, families = read_spreadsheet(_HCKS, profile_for(_HCKS), diag)
    return individuals, families, diag


def test_hcks_row_coverage_is_complete(hcks):
    _, _, diag = hcks
    tally, unaccounted = checks.row_coverage(diag)
    # Every name-bearing row is consumed: the 18 no-gen rows in particular.
    assert unaccounted == []
    assert tally["coded-person"] == 79
    assert tally["gen-person"] == 67
    assert tally["no-gen"] == 18


def test_hcks_has_eleven_placeholders(hcks):
    _, _, diag = hcks
    assert len(diag["placeholder_ids"]) == 11


def test_hcks_generation_consistency_clean(hcks):
    individuals, families, diag = hcks
    # The Ernest-Luke disambiguation keeps every name-linked link gen-consistent.
    assert checks.generation_consistency(diag, individuals, families) == []


def test_diagnostics_are_optional_and_do_not_change_output():
    # read_spreadsheet without diagnostics still works; with diagnostics it only
    # populates the dict (see test_golden for byte-identity).
    if not _BLSGRN.exists():
        pytest.skip(f"source spreadsheet not present: {_BLSGRN}")
    diag: dict = {}
    individuals, families = read_spreadsheet(_BLSGRN, profile_for(_BLSGRN), diag)
    assert diag["consumed_rows"]
    tally, unaccounted = checks.row_coverage(diag)
    assert unaccounted == []  # no whole rows dropped in the reference file
