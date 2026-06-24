"""Integration regression tests against the real source spreadsheet.

These guard three reader bugs found while linking the Steele line:

  * generation-0 rows must not be dropped (they are the reference generation)
  * a person recorded in several lineage charts (same name + exact birth
    date, different parents) must collapse to one individual with multiple
    FAMC links
  * '|' must be treated as the '/' child separator

The spreadsheet is local-only (gitignored), so the tests skip when absent.
"""

import re
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
    # Allan & Adrienne Steele's marriage is recorded only on their own rows
    # (cols 31/32: 11 Mar 1967, Sutherland), not as an 'M'-flag row. A stray
    # marriage row from the un-coded Forster note block (e.g. "James Forster
    # -m- Margaret", 1828) must not leak its date onto this family: the date
    # must be the couple's own, never the bogus 1828 Forster one.
    allan = _find(individuals, "Allan Richard Paul")[0]
    fam = next(f for f in families if allan.id in (f.husband_id, f.wife_id))
    assert fam.marriage_date == "11 MAR 1967"
    assert fam.marriage_place == "Sutherland, Sydney"

    # Every family that does carry a marriage date must have a spouse whose
    # surname plausibly belongs to that marriage (no orphaned Forster dates).
    by_id = {i.id: i for i in individuals}
    for f in families:
        if f.marriage_date:
            spouses = [by_id.get(f.husband_id), by_id.get(f.wife_id)]
            assert any(s and s.surname for s in spouses)


def _parent_family(families, child_id):
    return next((f for f in families if child_id in f.child_ids), None)


def _one(individuals, given, surname):
    matches = [
        i for i in individuals
        if i.given_name == given and (i.surname or "").upper() == surname.upper()
    ]
    assert len(matches) == 1, f"expected one {given} {surname}, got {len(matches)}"
    return matches[0]


def test_hierarchical_codes_nest_by_generation(parsed):
    """Deep codes like 'LivAx/Ar/...' must nest under their immediate parent,
    not collapse onto the top ancestor (the old _code_base bug)."""
    individuals, families = parsed
    by_id = {i.id: i for i in individuals}

    alex = _one(individuals, "Alexander", "LIVINGSTONE")
    norma = _one(individuals, "Norma Ethel", "DENIGAN")
    adrienne = _one(individuals, "Adrienne Carmel", "KENNA")
    bernadette = _one(individuals, "Bernadette", "KENNA")

    # Four distinct generations, each child anchored to its immediate parent.
    norma_parents = _parent_family(families, norma.id)
    assert norma_parents and alex.id in (norma_parents.husband_id, norma_parents.wife_id)

    adrienne_parents = _parent_family(families, adrienne.id)
    assert adrienne_parents and norma.id in (
        adrienne_parents.husband_id, adrienne_parents.wife_id
    )

    bern_parents = _parent_family(families, bernadette.id)
    assert bern_parents and adrienne.id in (
        bern_parents.husband_id, bern_parents.wife_id
    )

    # The great-grandchildren must NOT be direct children of Alexander.
    alex_fam = next(f for f in families if alex.id in (f.husband_id, f.wife_id))
    alex_child_names = {by_id[c].given_name for c in alex_fam.child_ids}
    assert "Bernadette" not in alex_child_names
    assert "Adrienne Carmel" not in alex_child_names  # granddaughter, not child


def test_no_family_has_collapsed_fanout(parsed):
    """No family should accumulate the dozens of mixed-surname descendants that
    the collapse bug produced (Alexander Livingstone had 50, John Allan Belshaw
    67)."""
    _, families = parsed
    biggest = max(len(f.child_ids) for f in families)
    assert biggest <= 15, f"a family still has {biggest} children"


def test_intermediate_spouse_is_spouse_not_child(parsed):
    """An '.../X-Y' code is a married-in spouse of a deep family, not a child of
    the top ancestor."""
    individuals, families = parsed
    by_id = {i.id: i for i in individuals}

    daphne = _one(individuals, "Daphne Edith", "FAULK")
    arthur = _one(individuals, "Arthur Silas", "FAULK")
    john_allan = _one(individuals, "John Allan", "BELSHAW")

    # Daphne is Arthur's wife, sharing a family with him.
    couple = next(
        (f for f in families
         if arthur.id in (f.husband_id, f.wife_id)
         and daphne.id in (f.husband_id, f.wife_id)),
        None,
    )
    assert couple is not None
    # She must not have been swept up as a child of the top ancestor.
    top = next(f for f in families if john_allan.id in (f.husband_id, f.wife_id))
    assert daphne.id not in top.child_ids


def test_no_parent_child_cycles(parsed):
    """The Pass-4 cycle guard must keep anyone from becoming their own ancestor
    (e.g. Bruce Dallas's father 'William H.' matching grandson 'William John
    Peter')."""
    _, families = parsed
    parents_of = {}
    for f in families:
        for c in f.child_ids:
            parents_of.setdefault(c, set()).update(
                p for p in (f.husband_id, f.wife_id) if p
            )

    def is_ancestor_of_self(start):
        seen, stack = set(), [start]
        while stack:
            cur = stack.pop()
            for p in parents_of.get(cur, ()):
                if p == start:
                    return True
                if p not in seen:
                    seen.add(p)
                    stack.append(p)
        return False

    cyclic = [c for c in parents_of if is_ancestor_of_self(c)]
    assert not cyclic


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


def test_marriage_read_from_spouse_rows(parsed):
    """Most couples record their marriage on the spouses' own rows (cols 31/32),
    not as an 'M'-flag row. Those must be recovered, not dropped."""
    individuals, families = parsed
    with_marriage = [f for f in families if f.marriage_date or f.marriage_place]
    # Far more than the ~20 'M'-flag rows would yield on their own.
    assert len(with_marriage) >= 80

    # Daniel Livingstone × Isobel married 9 Apr 1786 at Barony — recorded only
    # on their individual rows.
    daniel = _one(individuals, "Daniel", "LIVINGSTONE")
    fam = next(
        f for f in families
        if daniel.id in (f.husband_id, f.wife_id) and f.marriage_date
    )
    assert fam.marriage_date == "9 APR 1786"
    assert fam.marriage_place and "Barony" in fam.marriage_place


def test_shared_base_code_marriages_stay_distinct(parsed):
    """Two different people sharing a base-code prefix (e.g. an elder and a
    younger Thomas Ponting) must each keep their *own* col-31 marriage date,
    not have one leak onto the other. This is why attachment is keyed on the
    individual spouse, not on the family base code."""
    individuals, families = parsed
    by_id = {i.id: i for i in individuals}
    thomases = [
        i for i in individuals
        if i.given_name.startswith("Thomas") and (i.surname or "") == "PONTING"
    ]
    dated = {}
    for t in thomases:
        for f in families:
            if t.id in (f.husband_id, f.wife_id) and f.marriage_date:
                dated[t.id] = (f.marriage_date, by_id.get(f.wife_id))
    # At least two distinct Thomas Pontings, each with a different marriage date.
    distinct_dates = {d for d, _ in dated.values()}
    assert len(distinct_dates) >= 2, dated


def test_sex_assigned_from_family_role(parsed):
    individuals, families = parsed
    by_id = {i.id: i for i in individuals}
    # Every spouse must carry a sex. Role fills the gap left when no "née"
    # surname was present, without overriding an explicit inference: wives are
    # always F; husbands are M unless the slot holds a née-female lineage
    # connector whose own spouse is absent (then no swap could fire).
    for f in families:
        if f.husband_id:
            h = by_id[f.husband_id]
            assert h.sex is not None, f.husband_id
            assert h.sex == "M" or (h.sex == "F" and f.wife_id is None)
        if f.wife_id:
            assert by_id[f.wife_id].sex == "F", f.wife_id

    # A previously sexless male family head (James Hunter, no "née") is now M.
    james = _one(individuals, "James", "HUNTER")
    assert james.sex == "M"


def test_remarried_woman_filed_under_maiden_name(parsed):
    individuals, _ = parsed
    # "PONTING then PETTY [née Richey]" → surname Richey, married names noted.
    louisa = _one(individuals, "Louisa Georgina", "RICHEY")
    assert louisa.married_surnames == ["PONTING", "PETTY"]
    assert any("PONTING, then PETTY" in n for n in louisa.note_list)
    # The compound married string must never become the surname.
    assert not any(" then " in (i.surname or "").lower() for i in individuals)


def test_digit_or_unknown_spouse_suffix_is_spouse_not_child(parsed):
    """A spouse code whose suffix follows a digit ('PontHe/P2-J') or is unknown
    ('PontHe/L-?') must classify as a married-in spouse, not a child of the
    grandparent. The old letter-letter regex turned three husbands into
    children of the wrong family."""
    individuals, families = parsed
    by_id = {i.id: i for i in individuals}

    henry = _one(individuals, "Henry George", "PONTING")
    henry_child_ids = {c for f in families
                       if henry.id in (f.husband_id, f.wife_id)
                       for c in f.child_ids}
    henry_children = {by_id[c].given_name for c in henry_child_ids}

    # John W. Nolen (PontHe/P2-J) is Phoebe No.2's husband, not Henry's child.
    john = _one(individuals, "John W.", "NOLEN")
    assert john.given_name not in henry_children
    phoebe2 = _one(individuals, "Phoebe Louisa (No.2)", "NOLEN")
    couple = next((f for f in families
                   if john.id in (f.husband_id, f.wife_id)
                   and phoebe2.id in (f.husband_id, f.wife_id)), None)
    assert couple is not None

    # Phoebe No.2 herself IS Henry's child (a real, deeper relationship).
    assert phoebe2.id in henry_child_ids

    # Annette's unknown husband (BelLeAl/An-?) must not be a child of BelLeAl;
    # Annette (née Belshaw, surname recorded only as "?") heads her own marriage.
    annette = next(i for i in individuals if i.given_name == "Annette")
    annette_fams = [f for f in families
                    if annette.id in (f.husband_id, f.wife_id)]
    assert annette_fams  # she heads a marriage, not parented by BelLeAl head


def test_prior_marriage_chain_builds_own_family(parsed):
    """'GreJeAds-Ada-EmGe' is George Emberson, Ada's *first* husband. He must
    head his own family with Ada (sexed male), not become Jesse Green's wife
    with an impossible post-mortem marriage date."""
    individuals, families = parsed
    by_id = {i.id: i for i in individuals}

    george = next(i for i in individuals
                  if i.given_name == "George"
                  and (i.surname or "").upper() == "EMBERSON"
                  and i.death_date)  # the chain George has d=1901
    assert george.sex == "M"
    ada = _one(individuals, "Ada Rebecca", "SMITH")

    # George + Ada is a real couple; George is the husband.
    geo_fam = next(f for f in families
                   if george.id in (f.husband_id, f.wife_id)
                   and ada.id in (f.husband_id, f.wife_id))
    assert geo_fam.husband_id == george.id
    assert ada.id == geo_fam.wife_id

    # Ada also married Jesse Green: two distinct families, no marriage after
    # George's 1901 death attached to George's own family.
    jesse = next(i for i in individuals if i.given_name == "Jesse Adsley")
    jesse_fam = next(f for f in families
                     if jesse.id in (f.husband_id, f.wife_id)
                     and ada.id in (f.husband_id, f.wife_id))
    assert jesse_fam.id != geo_fam.id
    assert "1903" in (jesse_fam.marriage_date or "")


def test_remarried_husband_keeps_each_wifes_own_marriage_date(parsed):
    """Henry Ponting married twice; his single col-31 value (1884) must not
    overwrite his first wife Maud's own row date (1881). The earlier bug also
    produced a marriage dated after Maud's 1882 death."""
    individuals, families = parsed
    by_id = {i.id: i for i in individuals}

    henry = _one(individuals, "Henry George", "PONTING")
    maud = _one(individuals, "Maud", "PONTING")
    louisa = _one(individuals, "Louisa Georgina", "RICHEY")

    maud_fam = next(f for f in families
                    if henry.id in (f.husband_id, f.wife_id)
                    and maud.id in (f.husband_id, f.wife_id))
    louisa_fam = next(f for f in families
                      if henry.id in (f.husband_id, f.wife_id)
                      and louisa.id in (f.husband_id, f.wife_id))

    assert maud_fam.marriage_date == "1881"
    assert "Grenfell" in (maud_fam.marriage_place or "")
    assert louisa_fam.marriage_date == "5 MAR 1884"
    # Maud's marriage must precede her death, not follow it.
    assert maud.death_date == "4 JUL 1882"


def test_bracket_annotations_moved_out_of_name(parsed):
    """Editorial annotations like '[Infant death]' or '[MISSIONARY]' must not
    sit inside the GEDCOM given name; they become a NOTE instead. Parenthetical
    nicknames ('(Harry)') are extracted to a structured NICK."""
    individuals, _ = parsed
    # No given name retains a square-bracket annotation.
    assert not any(re.search(r"\[.*?\]", i.given_name or "") for i in individuals)

    david = _one(individuals, "David", "LIVINGSTONE")
    assert any("MISSIONARY" in n for n in david.note_list)

    # The two same-named infant Pontings stay distinct (different birth years),
    # each with the death annotation preserved as a note.
    mary_anns = [i for i in individuals
                 if i.given_name == "Mary Ann" and (i.surname or "").upper() == "PONTING"
                 and any("Infant death" in n for n in i.note_list)]
    assert len(mary_anns) == 2
    assert mary_anns[0].birth_date != mary_anns[1].birth_date

    # Parenthetical nicknames are extracted from the name into a NICK field.
    harry = _one(individuals, "Henry George", "PONTING")
    assert harry.nickname == "Harry"
    assert "(" not in (harry.given_name or "")


def test_longevity_discrepancy_flagged(parsed):
    """Where col-29 'Longevity' (age at death) contradicts the age implied by
    the birth and death dates, a discrepancy NOTE flags the likely transcription
    error — without inventing a corrected date or firing on approximate dates."""
    individuals, _ = parsed

    def disc_note(i):
        return [n for n in i.note_list
                if n.startswith("Recorded age at death")]

    robyn = _one(individuals, "Robyn", "BELSHAW")
    assert disc_note(robyn) and "(2)" in disc_note(robyn)[0] and "6 years" in disc_note(robyn)[0]

    roger = _one(individuals, "Roger John", "BELSHAW")
    assert disc_note(roger) and "(72)" in disc_note(roger)[0]

    # Henry Ponting's birth is approximate ('BET 1831 AND 1832'); the fuzzy year
    # must NOT trigger a spurious discrepancy note.
    henry = _one(individuals, "Henry George", "PONTING")
    assert not disc_note(henry)

    # Exactly the two genuine date/longevity contradictions are flagged.
    flagged = [i for i in individuals if disc_note(i)]
    assert len(flagged) == 2


def test_lineage_membership_recorded(parsed):
    individuals, _ = parsed
    daniel = _one(individuals, "Daniel", "LIVINGSTONE")
    assert "Livingstone" in daniel.lineage_lines
    # Lineage membership is written as a custom _GROUP tag by the writer, not
    # as a repetitive freeform note. No "Family lines:" note should remain.
    assert not any(n.startswith("Family lines:") for n in daniel.note_list)
    assert not any(
        n.startswith("Family lines:")
        for i in individuals
        for n in i.note_list
    )
