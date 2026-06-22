"""Integration regression tests against the real source spreadsheet.

These guard three reader bugs found while linking the Steele line:

  * generation-0 rows must not be dropped (they are the reference generation)
  * a person recorded in several lineage charts (same name + exact birth
    date, different parents) must collapse to one individual with multiple
    FAMC links
  * '|' must be treated as the '/' child separator

The spreadsheet is local-only (gitignored), so the tests skip when absent.
"""

from pathlib import Path

import pytest

from family_tree_converter.reader import read_spreadsheet

_SOURCE = Path(__file__).resolve().parent.parent / "data" / "input" / "BlsGrnLivMcCl H.Tr. #305.xls"


@pytest.fixture(scope="module")
def parsed():
    if not _SOURCE.exists():
        pytest.skip(f"source spreadsheet not present: {_SOURCE}")
    individuals, families = read_spreadsheet(_SOURCE)
    return individuals, families


def _find(individuals, given_sub, surname="STEELE"):
    return [
        i
        for i in individuals
        if surname.lower() in (i.surname or "").lower()
        and given_sub.lower() in (i.given_name or "").lower()
    ]


def test_generation_zero_rows_are_kept(parsed):
    individuals, _ = parsed
    # Allan and Adrienne are both generation 0 and must survive.
    assert _find(individuals, "Allan Richard Paul")
    assert _find(individuals, "Adrienne Lois")


def test_child_parent_chain_is_linked(parsed):
    individuals, families = parsed
    by_id = {i.id: i for i in individuals}
    edwin = _find(individuals, "Edwin Richard Paul")[0]
    allan = _find(individuals, "Allan Richard Paul")[0]

    edwin_parents = next(f for f in families if edwin.id in f.child_ids)
    assert allan.id in (edwin_parents.husband_id, edwin_parents.wife_id)

    allan_parents = next(f for f in families if allan.id in f.child_ids)
    father = by_id.get(allan_parents.husband_id)
    assert father is not None and father.given_name.startswith("Clive")


def test_bridging_person_is_deduplicated(parsed):
    individuals, families = parsed
    adriennes = _find(individuals, "Adrienne Lois")
    # Exactly one Adrienne Lois Steele, not one per lineage chart.
    assert len(adriennes) == 1
    adrienne = adriennes[0]

    # She is a child in two charts (Belshaw and Green) -> two FAMC links.
    parent_fams = [f for f in families if adrienne.id in f.child_ids]
    assert len(parent_fams) == 2
    by_id = {i.id: i for i in individuals}
    parent_surnames = {
        (by_id[f.husband_id].surname if f.husband_id else "")
        for f in parent_fams
    }
    assert "BELSHAW" in parent_surnames
    assert "GREEN" in parent_surnames
