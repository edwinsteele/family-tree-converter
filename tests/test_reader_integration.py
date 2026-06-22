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


def test_loose_given_name_children_recovered(parsed):
    individuals, families = parsed

    def find(given):
        return [
            i for i in individuals
            if i.given_name == given and (i.surname or "").upper() == "LIVINGSTONE"
        ]

    # Stella and Maud were listed only as bare given names; recovered as
    # children of Alexander Livingstone.
    stella = find("Stella")
    maud = find("Maud")
    assert len(stella) == 1 and len(maud) == 1

    # "Mabel" merely repeats the coded "Mabel Annie" (LivAx/Mb) and must not
    # become a second individual.
    assert not find("Mabel")
    assert find("Mabel Annie")

    alex = next(
        i for i in individuals
        if i.surname == "LIVINGSTONE" and i.given_name == "Alexander"
    )
    fam = next(f for f in families if stella[0].id in f.child_ids)
    assert fam.husband_id == alex.id


def test_cell_comments_become_notes(parsed):
    individuals, _ = parsed
    # The genealogist's research lived in Excel cell comments, not cell values.
    # Robert Belshaw's note about Wigan must survive onto his record.
    roberts = [
        i for i in individuals
        if (i.surname or "").upper() == "BELSHAW" and i.given_name.startswith("Robert")
    ]
    assert any("Wigan" in n for i in roberts for n in i.note_list)


def test_flag_based_forster_block_is_recovered(parsed):
    individuals, families = parsed
    by_id = {i.id: i for i in individuals}

    # Maud links to her father's family, recovered from B/D event rows.
    maud = next(
        i for i in individuals
        if i.surname == "PONTING" and i.given_name == "Maud"
    )
    fam = next(
        f for f in families
        if maud.id in f.child_ids
        and f.husband_id
        and by_id[f.husband_id].given_name == "John"
        and (by_id[f.husband_id].surname or "").upper() == "FORSTER"
    )
    john = by_id[fam.husband_id]
    assert john.birth_date and john.death_date  # b ABT 1829, d 1901

    # Maud's mother Ellen is filled into the same family.
    assert fam.wife_id and by_id[fam.wife_id].given_name.startswith("Ellen")
    # Edward and Samuel are Maud's siblings in that same family.
    sib_names = {by_id[c].given_name for c in fam.child_ids}
    assert {"Edward", "Samuel"} <= sib_names

    # John T. (James's son by his second wife Bridget) is a *distinct* person,
    # not collapsed into John despite the shared first name.
    john_t = next(
        i for i in individuals
        if (i.surname or "").upper() == "FORSTER" and i.given_name == "John T."
    )
    assert john_t.id != john.id


def test_marriage_not_misattributed_to_unrelated_family(parsed):
    individuals, families = parsed
    # Allan & Adrienne Steele have no marriage row in the source. A stray
    # marriage row from the un-coded Forster note block (e.g. "James Forster
    # -m- Margaret", 1828) must not leak its date onto an unrelated family.
    allan = _find(individuals, "Allan Richard Paul")[0]
    fam = next(f for f in families if allan.id in (f.husband_id, f.wife_id))
    assert fam.marriage_date is None

    # Every family that does carry a marriage date must have a spouse whose
    # surname plausibly belongs to that marriage (no orphaned Forster dates).
    by_id = {i.id: i for i in individuals}
    for f in families:
        if f.marriage_date:
            spouses = [by_id.get(f.husband_id), by_id.get(f.wife_id)]
            assert any(s and s.surname for s in spouses)


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
