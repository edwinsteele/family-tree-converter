"""Integration regression tests for the "Hcks:Thos:Krsl" tree (multi-tree merge).

This file is *half* coded: ~79 rows use the same alpha path-code convention as
the reference spreadsheet, while the rest are generation-numbered people who name
their parents in the Father/Mother columns (linked by name, not code). It also
carries cross-tree "bridge" codes, private "not for publication" note columns,
and informal parent references (nickname / initials / middle name / maiden).

These tests guard the coded path (Stage 1) and the uncoded name-linking
(Stage 2). The no-generation "compact descendant" rows (Stage 3) are not yet
captured, so counts are asserted with thresholds rather than exact values.

The spreadsheet is local-only (gitignored), so the tests skip when absent.
"""

from pathlib import Path

import pytest

from family_tree_converter.reader import profile_for, read_spreadsheet
from family_tree_converter.validate import validate

_SOURCE = Path(__file__).resolve().parent.parent / "data" / "input" / "Hcks:Thos:Krsl H.Tr.#120"


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


def _parents(individuals, families, ind):
    by_id = {i.id: i for i in individuals}
    fam = next((f for f in families if ind.id in f.child_ids), None)
    if fam is None:
        return None, None
    return by_id.get(fam.husband_id), by_id.get(fam.wife_id)


def test_profile_selected_by_name():
    assert profile_for(_SOURCE).name == "Hcks:Thos:Krsl"


def test_counts_and_integrity(parsed):
    individuals, families = parsed
    # 79 coded + 67 uncoded people, plus synthetic ancestors.
    assert len(individuals) > 160
    assert len(families) > 55
    result = validate(individuals, families)
    assert result["errors"] == []


def test_coded_person_read(parsed):
    individuals, _ = parsed
    pharaoh = _find(individuals, "Pharaoh", "THOMAS")
    assert pharaoh and pharaoh[0].birth_date == "25 JUN 1841"


def test_bridge_spouse_is_wife_not_child(parsed):
    # Esther (code 'TmJo/P-HckC/Es') is Pharaoh Thomas's wife *and* Esther Hicks,
    # daughter of Charles Hicks. She must be Pharaoh's WIFE, not his child
    # (she was born 1830, eleven years before him), while remaining Charles
    # Hicks's daughter — the cross-tree bridge.
    individuals, families = parsed
    by_id = {i.id: i for i in individuals}
    esther = [i for i in individuals
              if i.given_name == "Esther" and i.surname == "THOMAS"][0]
    father, _ = _parents(individuals, families, esther)
    assert father and father.given_name == "Charles" and father.surname == "HICKS"
    spouse_fams = [f for f in families if esther.id in (f.husband_id, f.wife_id)]
    husbands = {by_id[f.husband_id].given_name for f in spouse_fams if f.husband_id}
    assert "Pharaoh" in husbands


def test_uncoded_linked_by_middle_name(parsed):
    # "Alfred Ernest Luke" names his father "Leslie Luke" — the middle name of
    # Alfred Leslie Luke. The unique-token index must resolve it to the real
    # person rather than minting a duplicate.
    individuals, families = parsed
    ernest = _find(individuals, "Alfred Ernest", "LUKE")[0]
    father, mother = _parents(individuals, families, ernest)
    assert father and father.given_name == "Alfred Leslie"
    assert mother and mother.given_name == "Clara"
    # no duplicate bare "Leslie Luke"
    assert not [i for i in individuals if i.given_name == "Leslie"
                and (i.surname or "") == "LUKE"]


def test_uncoded_linked_by_nickname(parsed):
    # "Jack Luke" (father) is the nickname of Alfred John Luke.
    individuals, families = parsed
    ann = _find(individuals, "Ann Patricia", "RICHARDS")[0]
    father, _ = _parents(individuals, families, ann)
    assert father and father.given_name == "Alfred John"


def test_uncoded_mother_matched_by_maiden(parsed):
    # Gregory James's mother is named "Jennifer Phillips" but she is filed under
    # her married surname JAMES; the maiden-name index must still resolve her.
    individuals, families = parsed
    greg = _find(individuals, "Gregory", "JAMES")[0]
    _, mother = _parents(individuals, families, greg)
    assert mother and mother.given_name == "Jennifer"


def test_chronology_guard_rejects_younger_namesake(parsed):
    # Edith Joan (b 1915) names her mother "Dorothy Julian". A younger Dorothy
    # Williams née Julian (b 1925) exists; the parent must NOT be her.
    individuals, families = parsed
    edith = _find(individuals, "Edith Joan", "JACKSON")[0]
    _, mother = _parents(individuals, families, edith)
    assert mother is not None
    assert mother.birth_date != "1925"


def test_private_notes_excluded(parsed):
    # Cell comments in the "Not for publication" columns (35/36) must not leak.
    individuals, _ = parsed
    all_notes = "\n".join(n for i in individuals for n in i.note_list)
    assert "Many details of Pharaoh" not in all_notes
