"""Tests for the post-conversion integrity checker."""

from pathlib import Path

import pytest

from family_tree_converter.reader import Family, Individual, read_spreadsheet
from family_tree_converter.validate import validate

_SOURCE = Path(__file__).resolve().parent.parent / "data" / "input" / "BlsGrnLivMcCl H.Tr. #305.xls"


def test_real_conversion_has_no_integrity_errors():
    if not _SOURCE.exists():
        pytest.skip(f"source spreadsheet not present: {_SOURCE}")
    individuals, families = read_spreadsheet(_SOURCE)
    result = validate(individuals, families)
    assert result["errors"] == []
    # The only warnings are the two plausible older-father Ponting births.
    assert all("PONTING" in w for w in result["warnings"])
    assert len(result["warnings"]) == 2


def test_validate_catches_birth_after_death():
    inds = [Individual(id="I1", given_name="A", surname="X",
                       birth_date="1900", death_date="1850")]
    assert any("after death" in e for e in validate(inds, [])["errors"])


def test_validate_catches_dangling_pointer_and_cycle():
    inds = [Individual(id="I1", given_name="A", surname="X"),
            Individual(id="I2", given_name="B", surname="X")]
    # Dangling child pointer.
    fams = [Family(id="F1", husband_id="I1", child_ids=["I99"])]
    assert any("missing individual I99" in e for e in validate(inds, fams)["errors"])
    # I1 is its own grandparent: I1 -> child I2, and I2 -> child I1.
    cyc = [Family(id="F1", husband_id="I1", child_ids=["I2"]),
           Family(id="F2", husband_id="I2", child_ids=["I1"])]
    assert any("cycle" in e for e in validate(inds, cyc)["errors"])


def test_validate_catches_sex_role_mismatch():
    inds = [Individual(id="I1", given_name="A", surname="X", sex="F"),
            Individual(id="I2", given_name="B", surname="Y", sex="F")]
    fams = [Family(id="F1", husband_id="I1", wife_id="I2")]
    assert any("husband" in e and "female" in e for e in validate(inds, fams)["errors"])


def test_validate_flags_implausible_maternal_age():
    # A mother bearing a child at 55 is flagged (the Mary Muldoon / 22-child
    # block pattern); the father at the same age is not.
    inds = [Individual(id="I1", given_name="Dad", surname="P", sex="M",
                       birth_date="1831"),
            Individual(id="I2", given_name="Mum", surname="P", sex="F",
                       birth_date="1833"),
            Individual(id="I3", given_name="Late", surname="P", birth_date="1888")]
    fams = [Family(id="F1", husband_id="I1", wife_id="I2", child_ids=["I3"])]
    warns = validate(inds, fams)["warnings"]
    assert any("maternal age" in w for w in warns)
    assert not any("father" in w and "maternal" in w for w in warns)


def test_validate_maternal_age_charitable_to_ranges():
    # A range birth ("BET 1800 AND 1809") is judged at its latest year so a
    # plausibly-young mother does not false-fire (Sarah Clow).
    inds = [Individual(id="I2", given_name="Mum", surname="C", sex="F",
                       birth_date="BET 1800 AND 1809"),
            Individual(id="I3", given_name="Kid", surname="C", birth_date="1853")]
    fams = [Family(id="F1", wife_id="I2", child_ids=["I3"])]
    assert not any("maternal age" in w for w in validate(inds, fams)["warnings"])
