"""Integration regression tests for the "Brc:Stl" tree (multi-tree merge).

This file is the Steele line's Scottish (Bruce/Steel) ancestry, drawn as a
*descending* horizontal tree with NO path codes at all: every person is
generation-numbered (counting down 11→0 from the oldest ancestor) and names
their parents in the Father/Mother columns, so the whole file links by name
(Pass 4b) rather than by code.

These tests guard the distinctive behaviours: descending-generation name
linking, the STEEL→STEELE spelling bridge between the Scottish ancestors and
their Australian descendants, cross-block deduplication of a person who appears
in two family charts, the chronologically-impossible-death guard, and the
exclusion of marginal annotation rows, an unflagged marriage row, and the
trailing speculative "alternatives" appendix.

The spreadsheet is local-only (gitignored), so the tests skip when absent.
"""

from pathlib import Path

import pytest

from family_tree_converter.reader import profile_for, read_spreadsheet
from family_tree_converter.validate import validate

_SOURCE = Path(__file__).resolve().parent.parent / "data" / "input" / "Brc:Stl H.Tree #46"


@pytest.fixture(scope="module")
def parsed():
    if not _SOURCE.exists():
        pytest.skip(f"source spreadsheet not present: {_SOURCE}")
    return read_spreadsheet(_SOURCE, profile_for(_SOURCE))


def _find(individuals, given_sub, surname):
    return [
        i for i in individuals
        if surname.lower() == (i.surname or "").lower()
        and given_sub.lower() in (i.given_name or "").lower()
    ]


def _parents(individuals, families, ind):
    by_id = {i.id: i for i in individuals}
    fam = next((f for f in families if ind.id in f.child_ids), None)
    if fam is None:
        return None, None
    return by_id.get(fam.husband_id), by_id.get(fam.wife_id)


def test_profile_selected_by_name():
    assert profile_for(_SOURCE).name == "Brc:Stl"


def test_counts_and_integrity(parsed):
    individuals, families = parsed
    assert len(individuals) == 141
    assert len(families) == 54
    result = validate(individuals, families)
    assert result["errors"] == []
    assert result["warnings"] == []
    # Every individual belongs to at least one family (no orphans).
    in_family = {
        s for f in families
        for s in (f.husband_id, f.wife_id, *f.child_ids) if s
    }
    assert [i.given_name + " " + i.surname
            for i in individuals if i.id not in in_family] == []


def test_descending_generation_links_repeated_namesakes(parsed):
    # The Bruce line repeats "David Bruce" across generations (b. ~1640, 1663,
    # 1712). Descending generation numbers must still link each David to the
    # David one generation above — not collapse them or invert the chain.
    individuals, families = parsed
    davids = sorted(_find(individuals, "David", "BRUCE"), key=lambda i: i.birth_date or "")
    assert len(davids) == 3
    # The 1663 David is a child of the ~1640 David; the 1712 David descends from
    # the 1663 David (one generation apart, never the same person).
    d1663 = next(i for i in davids if "1663" in (i.birth_date or ""))
    father, _ = _parents(individuals, families, d1663)
    assert father and father.surname == "BRUCE" and "164" in (father.birth_date or "")


def test_steel_steele_spelling_bridge(parsed):
    # The Scottish ancestors are charted "STEEL", their Australian descendants
    # "STEELE"; the children of James Steel (b.1829) name their father "James
    # Steele". The spelling-tolerant generation match must resolve them to the
    # real gen-4 father — NOT mint synthetics and NOT attach them to his own
    # same-named son James (b.1861), who is himself one of the children.
    individuals, families = parsed
    by_id = {i.id: i for i in individuals}
    father = next(i for i in _find(individuals, "James", "STEEL")
                  if "1829" in (i.birth_date or ""))
    fam = next(f for f in families if f.husband_id == father.id)
    kids = [by_id[c] for c in fam.child_ids]
    given_surnames = {(k.given_name, k.surname) for k in kids}
    # The son James STEELE (b.1861) is a CHILD here, not a parent.
    son = next(i for i in _find(individuals, "James", "STEELE")
               if "1861" in (i.birth_date or ""))
    assert son.id in fam.child_ids
    assert not any(f.husband_id == son.id for f in families)
    # The whole gen-3 sibling set lands in this one family (no split / synthetic).
    assert ("Charles", "STEELE") in given_surnames
    assert ("Mary", "STEELE") in given_surnames


def test_cross_block_dedup_married_woman(parsed):
    # Margaret (née Buchan) appears both in the Bruce block as David Bruce's
    # wife and in the Buchan block as George Buchan's daughter. She must be ONE
    # individual: a child of the Buchans and a spouse in a Bruce family.
    individuals, families = parsed
    margarets = [i for i in _find(individuals, "Margaret", "BRUCE")
                 if "1714" in (i.birth_date or "")]
    assert len(margarets) == 1
    m = margarets[0]
    father, _ = _parents(individuals, families, m)
    assert father and father.surname == "BUCHAN"
    assert any(m.id in (f.husband_id, f.wife_id) for f in families)


def test_impossible_death_demoted_to_note(parsed):
    # Mary Ann Costigan's recorded death (1904) precedes her birth (1917) and her
    # parents' 1911 marriage. The impossible death event is dropped; the figure
    # survives verbatim as a note, with the birth (the corroborated value) kept.
    individuals, _ = parsed
    annie = _find(individuals, "Mary Ann", "COSTIGAN")[0]
    assert annie.birth_date == "1917"
    assert annie.death_date is None
    assert any("1904" in n and "precedes the recorded birth" in n
               for n in annie.note_list)


def test_per_person_marriage_attaches_to_name_linked_family(parsed):
    # Most couples record their marriage on each spouse's own row (col 35), not on
    # a separate 'M'-flag row. Those dates must attach to the name-linked family —
    # e.g. Henry Joseph Costigan married Jane (née Steel) in 1871, recorded only
    # per-person. (Regression: the per-person attach used to skip role-"unknown"
    # rows, so every no-code couple without an 'M' row lost its marriage.)
    individuals, families = parsed
    henry = next(i for i in _find(individuals, "Henry Joseph", "COSTIGAN"))
    fam = next(f for f in families if f.husband_id == henry.id)
    assert fam.marriage_date == "20 OCT 1871"


def test_marginal_annotation_row_excluded(parsed):
    # A generation-numbered "{2nd wife - ??}" placeholder is marginalia, not a
    # person — it must not become an individual.
    individuals, _ = parsed
    assert not [i for i in individuals if "{" in (i.given_name or "")]
    assert not [i for i in individuals if "wife" in (i.given_name or "").lower()]


def test_unflagged_marriage_row_not_a_person(parsed):
    # "Henry Steel married Agnes Anderson" is a marriage row lacking its 'M'
    # flag; it must neither become a person nor a loose child.
    individuals, _ = parsed
    assert not [i for i in individuals if "married" in (i.given_name or "").lower()]


def test_alternatives_appendix_excluded(parsed):
    # The sheet ends with a speculative "BELOW ARE ALTERNATIVES…" / "IS THIS A
    # GOER?" appendix (rows 181+); none of it should appear as individuals.
    individuals, _ = parsed
    blob = " ".join((i.given_name or "") + " " + (i.surname or "")
                    for i in individuals).upper()
    assert "ALTERNATIVE" not in blob
    assert "GOER" not in blob
