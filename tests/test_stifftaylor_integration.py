"""Integration regression tests for the "Stiff:Taylor" tree (multi-tree merge).

The most entwined tree (16 people shared with the other files) but structurally
the cleanest of the no-code files: every row is generation-numbered and names its
parents in the Father/Mother columns, so the whole file links by name (Pass 4b)
like Brc:Stl. There are NO path codes, NO 'M'-flag rows (marriage is per-person),
no embedded headers and no disavowed appendix. The distinctive features this file
exercises:

  * the role markers are split across three columns — Sp (spouse, col 16), the
    'X' re-appears marker plus 1/2 first/second-spouse (col 15), and Dv (col 13);
  * a née woman recorded once under her birth name and once under a married name
    (Isabella Taylor — de-facto of Henry Joseph Bowden, later wife of Charles
    Mullins) is merged into ONE person by maiden + parents + exact birth date,
    keeping BOTH marriages;
  * a partner tagged '[De-facto]' on the surname: the relationship is preserved
    as a note and the family carries no (legal) marriage date, even when the
    other spouse's own legal-marriage date would otherwise leak onto it;
  * a spouse whose name lands only in the secondary "SURNAME"/"Other Names"
    display columns (Frederick James Giles), recovered so he is not nameless;
  * new approximate-date forms: a two-digit-end year range ("1832-37"),
    "late Feb 1994", and a non-date marriage cell ("...never married.").

The spreadsheet is local-only (gitignored), so the tests skip when absent.
"""

from pathlib import Path

import pytest

from family_tree_converter.reader import profile_for, read_spreadsheet
from family_tree_converter.validate import validate

_SOURCE = (Path(__file__).resolve().parent.parent
           / "data" / "input" / "Stiff:Taylor H.Tree #275")


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


def test_profile_selected_by_name():
    assert profile_for(_SOURCE).name == "Stiff:Taylor"


def test_counts_and_integrity(parsed):
    individuals, families = parsed
    assert len(individuals) == 306
    assert len(families) == 95
    result = validate(individuals, families)
    assert result["errors"] == []
    assert result["warnings"] == []


def test_no_orphans(parsed):
    individuals, families = parsed
    in_family = {
        s for f in families
        for s in (f.husband_id, f.wife_id, *f.child_ids) if s
    }
    orphan_givens = sorted(
        (i.given_name or "") for i in individuals if i.id not in in_family)
    assert orphan_givens == []


def test_isabella_merged_across_maiden_and_married_surname(parsed):
    # Isabella Taylor (de-facto of Henry Joseph Bowden) and Isabella Mullins
    # (née Taylor) are the same person — same parents (John Taylor + Eliza
    # Conway) and the same 19 Feb 1851 birth — flagged by the genealogist's 'X'
    # re-appears marker. She appears exactly once, filed under her birth surname
    # with MULLINS recorded as a married surname, and is wife in BOTH unions.
    individuals, families = parsed
    by_id = {i.id: i for i in individuals}
    isabellas = [i for i in individuals
                 if "Isabella" in (i.given_name or "")
                 and i.birth_date and "1851" in i.birth_date]
    assert len(isabellas) == 1
    isabella = isabellas[0]
    assert isabella.surname == "TAYLOR"
    assert "MULLINS" in isabella.married_surnames
    husbands = {by_id[f.husband_id].surname
                for f in families if f.wife_id == isabella.id and f.husband_id}
    assert {"BOWDEN", "MULLINS"} <= husbands


def test_defacto_union_has_no_marriage_date_and_note(parsed):
    # Henry Joseph Bowden & Isabella Taylor never married (the marriage cell reads
    # "Joseph & Isabella never married."): their family carries no marriage date
    # and a de-facto note. The SAME applies to Joseph C.D. Bowden & Helen Wiseman.
    individuals, families = parsed
    by_id = {i.id: i for i in individuals}
    for husband_given, wife_surname in (
            ("Henry Joseph", "TAYLOR"), ("Joseph Charles Dandy", "WISEMAN")):
        fam = next(
            f for f in families
            if f.husband_id and husband_given in (by_id[f.husband_id].given_name or "")
            and f.wife_id and by_id[f.wife_id].surname == wife_surname)
        assert fam.marriage_date is None
        assert any("de-facto" in n.lower() for n in fam.note_list)


def test_legal_marriage_date_not_lost_to_defacto_family(parsed):
    # Henry Joseph Bowden's only recorded date (7 Feb 1885) is his legal marriage
    # to Elizabeth Goscomb, not the de-facto union — it must land on the Goscomb
    # family, which it does (and the de-facto family stays dateless, above).
    individuals, families = parsed
    by_id = {i.id: i for i in individuals}
    eliz = _find(individuals, "Elizabeth", "BOWDEN")
    fam = next(f for f in families
               if f.wife_id in {i.id for i in eliz} and f.husband_id
               and "Henry Joseph" in (by_id[f.husband_id].given_name or ""))
    assert fam.marriage_date == "7 FEB 1885"


def test_defacto_surname_tag_preserved_as_note(parsed):
    # The '[De-facto]' surname tag is stripped from the surname but preserved as a
    # relationship note on the partner.
    individuals, _ = parsed
    isabella = [i for i in individuals
                if i.surname == "TAYLOR" and "Isabella" in (i.given_name or "")][0]
    assert any("de-facto" in n.lower() for n in isabella.note_list)


def test_secondary_name_columns_recover_nameless_spouse(parsed):
    # Frederick James Giles's name is recorded only in the dup "SURNAME"/"Other
    # Names" display columns (the primary cells are blank); he must be named, not
    # left as an empty individual.
    individuals, _ = parsed
    giles = _find(individuals, "Frederick James", "GILES")
    assert len(giles) == 1
    assert giles[0].given_name == "Frederick James"


def test_children_homed_on_correct_wife_not_namesake(parsed):
    # Clive Steele's children name their mother "Edna Stiff" — which must resolve
    # to his wife Edna May Steele née Stiff (b.1917, dau. of George Stiff + Violet
    # Taylor), NOT the namesake Edna Mary Stiff née Tracey (b.1919, "Aunty Mary"
    # who married Albert Stiff). The consolidation re-homes the children onto the
    # wife whose surname matches the husband's and drops the spurious family.
    individuals, families = parsed
    by_id = {i.id: i for i in individuals}
    allan = _find(individuals, "Allan Richard Paul", "STEELE")[0]
    fam = next(f for f in families if allan.id in f.child_ids)
    mother = by_id[fam.wife_id]
    assert mother.given_name == "Edna May"
    assert mother.birth_date == "22 SEP 1917"
    # her parents are George Stiff + Violet Taylor, not Tracey + Jones
    mfam = next(f for f in families if mother.id in f.child_ids)
    assert by_id[mfam.husband_id].surname == "STIFF"
    assert "Violet" in (by_id[mfam.wife_id].given_name or "")
    # Aunty Mary is only ever the wife of Albert Stiff, never Clive's wife
    mary = next(i for i in individuals
                if i.given_name == "Edna Mary" and i.surname == "STIFF")
    mary_husbands = {by_id[f.husband_id].given_name
                     for f in families if f.wife_id == mary.id and f.husband_id}
    assert mary_husbands == {"Albert Arthur"}


def test_duplicate_synthetic_wives_consolidated(parsed):
    # Seven mothers named given-only by their children (Jean, Doreen, Debbie,
    # Susan, Madge, Annie, Ann) each minted a synthetic that duplicated the
    # wife's own row, producing a spurious childless family. After consolidation
    # each husband has a single family with his real, fully-named wife.
    individuals, families = parsed
    by_id = {i.id: i for i in individuals}
    for husband_given, husband_surname in (
            ("Herbert Campbell", "EVANS"), ("Allan E.", "LOCK"),
            ("Albert Edward", "TAYLOR")):
        h = _find(individuals, husband_given, husband_surname)[0]
        wife_fams = [f for f in families if f.husband_id == h.id]
        assert len(wife_fams) == 1
        wife = by_id[wife_fams[0].wife_id]
        assert wife.surname == husband_surname  # real wife, took his name


def test_approximate_year_marker_wraps_abt(parsed):
    # The "Approximate (A)" flag (col 32) marks a year as uncertain: a year-only
    # date on such a row is emitted ABT (John David Evans b.1867, Robert Henry
    # Feil b.1890), while full day/month/year dates stay exact.
    individuals, _ = parsed
    robert = _find(individuals, "Robert Henry", "FEIL")[0]
    assert robert.birth_date == "ABT 1890"


def test_unknown_given_kept_faithfully(parsed):
    # Children whose given name the genealogist left as "?" keep that marker
    # rather than being dropped (e.g. the unnamed PINFIELD children).
    individuals, _ = parsed
    assert any(i.given_name == "?" and i.surname == "PINFIELD"
               for i in individuals)
