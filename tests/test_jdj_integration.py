"""Integration regression tests for the "J & D J Steele" tree (multi-tree merge).

Another Steele branch, drawn like Brc:Stl with NO path codes: the main tree
(rows 14-158) is generation-numbered and names parents in the Father/Mother
columns, so it links by name (Pass 4b). The distinctive features this file
exercises:

  * col 15 is the *Mother* column (not an empty code spare like Brc), so the
    profile points the unused code column at the genuinely-empty col 5;
  * a mid-file disavowed appendix — the genealogist's note "The following 13
    entries are now known to have nothing to do with ... family" and the 13
    Tasmanian PRICE rows after it — excluded via a skip range, with the real
    PRICE in-married family (its own embedded header, marriage row, 22 children)
    beginning afterwards and kept;
  * the PRICE block's parentless children adopt their block marriage couple, and
    John/Mary/Dorothy unify with the records the main tree already implied;
  * initials-aware parent matching ("J. A. Bruce Steele" = James Alexander Bruce
    Steele; "Fred." = Frederick/Fredrick), and generation disambiguation of the
    five different "James Steele"s;
  * an "[Adopted]" parentage note, a split Sp/Dv marker (divorce), and a
    Y/M/D numeric date.

The spreadsheet is local-only (gitignored), so the tests skip when absent.
"""

from pathlib import Path

import pytest

from family_tree_converter.reader import profile_for, read_spreadsheet
from family_tree_converter.validate import validate

_SOURCE = (Path(__file__).resolve().parent.parent
           / "data" / "input" / "J & D J Steele H.Tree #30")


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
    assert profile_for(_SOURCE).name == "J & D J Steele"


def test_counts_and_integrity(parsed):
    individuals, families = parsed
    assert len(individuals) == 169
    assert len(families) == 51
    result = validate(individuals, families)
    assert result["errors"] == []
    assert result["warnings"] == []


def test_only_unknowable_parentage_is_orphaned(parsed):
    # Everyone links into a family except the two whose parentage the source
    # leaves unknowable: the adopted Queenie and "Amy" (recorded with no surname).
    individuals, families = parsed
    in_family = {
        s for f in families
        for s in (f.husband_id, f.wife_id, *f.child_ids) if s
    }
    orphan_givens = sorted(
        (i.given_name or "") for i in individuals if i.id not in in_family)
    assert orphan_givens == ["Amy", "Queenie"]


def test_disavowed_price_block_excluded(parsed):
    # Rows 165-178 are a PRICE family (Tasmania) the genealogist flagged as
    # "nothing to do with" the tree; none of its distinctive members/places
    # (Launceston, Pt. Sorell, the "(Female)" stillbirths) should appear.
    individuals, _ = parsed
    blob = " ".join((i.given_name or "") + " " + (i.surname or "")
                    + " " + (i.birth_place or "") for i in individuals)
    assert "following 13 entries" not in blob
    assert "Launceston" not in blob
    assert "Sorell" not in blob
    assert "(Female)" not in blob


def test_real_price_family_kept_and_unified(parsed):
    # The real in-married PRICE family is John Price (b.1831) + Mary Muldoon,
    # married 1852, with 22 children. John/Mary unify with the parents implied by
    # Dorothy Jane (née Price)'s row, and Dorothy is one of the 22 children — not
    # a duplicate.
    individuals, families = parsed
    johns = [i for i in _find(individuals, "John", "Price")
             if i.birth_date and "1831" in i.birth_date]
    assert len(johns) == 1
    fam = next(f for f in families if f.husband_id == johns[0].id)
    assert fam.marriage_date == "1852"
    assert len(fam.child_ids) == 22
    # Dorothy Jane (née Price) appears exactly once and is among those children.
    dorothys = [i for i in individuals
                if i.given_name and "Dorothy Jane" in i.given_name]
    assert len(dorothys) == 1
    assert dorothys[0].id in fam.child_ids


def test_james_steele_initials_and_generation(parsed):
    # "J. A. Bruce Steele" (referenced by his children) is James Alexander Bruce
    # Steele (b.1904), NOT the same-token "Hazel J." nor the b.1861 root. And the
    # children of "James Steele" married to Olive Castle resolve to James (b.1928)
    # — the generation-correct James, not the b.1861 root (which would make him a
    # ~97-year-old father).
    individuals, families = parsed
    by_id = {i.id: i for i in individuals}
    jab = next(i for i in _find(individuals, "James Alexander Bruce", "STEELE"))
    fam = next(f for f in families if f.husband_id == jab.id)
    assert {by_id[c].given_name for c in fam.child_ids} >= {
        "Matthew James Bruce", "Coral", "Raymond David"}
    louise = next(i for i in individuals
                  if i.given_name and "Louise Dorian" in i.given_name)
    father, _ = _parents(individuals, families, louise)
    assert father and father.given_name == "James"
    assert father.birth_date and "1928" in father.birth_date


def test_fred_abbreviation_links_not_duplicated(parsed):
    # The Abberton children name their father "Fred. T." — an abbreviation of
    # Frederick (the source also spells it Fredrick). It must resolve to the real
    # Frederick T. Abberton row, not mint a duplicate that orphans him.
    individuals, families = parsed
    freds = _find(individuals, "Frederick T.", "ABBERTON")
    assert len(freds) == 1
    assert any(f.husband_id == freds[0].id and f.child_ids for f in families)


def test_adopted_parentage_note(parsed):
    # Queenie's parents are recorded only as "[Adopted]"; that builds no bogus
    # self-married family, and the annotation is preserved as a note.
    individuals, _ = parsed
    queenie = _find(individuals, "Queenie", "WILLIAMS")[0]
    assert any("Adopted" in n for n in queenie.note_list)


def test_split_divorce_marker(parsed):
    # Warren Butler is married-in (Sp, col 13) and divorced (Dv, col 11 — a
    # separate status column); his marriage is flagged as ended in divorce.
    individuals, families = parsed
    warren = _find(individuals, "Warren", "BUTLER")[0]
    fam = next(f for f in families
               if warren.id in (f.husband_id, f.wife_id))
    assert fam.divorced is True


def test_year_first_numeric_date(parsed):
    # An infant death recorded "1931/2/3" (Y/M/D) becomes a valid GEDCOM date.
    individuals, _ = parsed
    john = next(i for i in _find(individuals, "John", "STEELE")
                if i.death_date == "3 FEB 1931")
    assert john is not None
