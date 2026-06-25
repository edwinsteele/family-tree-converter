"""Integration regression tests for the "Hcks:Thos:Krsl" tree (multi-tree merge).

This file is *half* coded: ~79 rows use the same alpha path-code convention as
the reference spreadsheet, while the rest are generation-numbered people who name
their parents in the Father/Mother columns (linked by name, not code). It also
carries cross-tree "bridge" codes, private "not for publication" note columns,
and informal parent references (nickname / initials / middle name / maiden).

These tests guard the coded path (Stage 1), the uncoded name-linking (Stage 2),
the no-generation "compact descendant" adjacency pass (Stage 3), and the audit
reconciliations (Stage 4: name-vs-code mother, generation disambiguation, and
duplicate-couple merging).

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
    # 79 coded + 67 uncoded + 18 no-generation people, plus synthetic ancestors.
    assert len(individuals) == 196
    assert len(families) == 72
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


# --- Stage 3: no-generation "compact descendant" adjacency ---

def _children(individuals, families, husband_sub, hsur, wife_sub, wsur):
    by_id = {i.id: i for i in individuals}
    for f in families:
        h, w = by_id.get(f.husband_id), by_id.get(f.wife_id)
        if (h and w and husband_sub in (h.given_name or "")
                and hsur in (h.surname or "") and wife_sub in (w.given_name or "")
                and wsur in (w.surname or "")):
            return f, [by_id[c] for c in f.child_ids]
    return None, []


def test_nogen_spouse_above_pairs_husband(parsed):
    # Peter Richards has no generation/code; he sits directly above the née-female
    # Ann Patricia (née Luke) Richards, so he is her husband, and the plain no-gen
    # rows below (Mark, Sarah) are their children.
    individuals, families = parsed
    fam, kids = _children(individuals, families, "Peter", "RICHARDS",
                          "Ann Patricia", "RICHARDS")
    assert fam is not None
    given = sorted(k.given_name for k in kids)
    assert given == ["Mark", "Sarah"]


def test_nogen_nameless_children_get_placeholder(parsed):
    # John Buchanan + Janet Mary (née Luke) have two children recorded with a
    # surname only; each is captured under a placeholder name with a note.
    individuals, families = parsed
    fam, kids = _children(individuals, families, "John", "BUCHANAN",
                          "Janet Mary", "BUCHANAN")
    assert fam is not None and len(kids) == 2
    for k in kids:
        assert k.given_name == "[Unnamed]"
        assert any("unnamed child" in n for n in k.note_list)


def test_childless_spouse_orphans_paired(parsed):
    # The three childless married-in spouses must each land in a family: Edward
    # Greenslade Pearce (above his née-female wife), George Hartely (Lorna
    # Kerslake's second husband, matched by her "NAGLE then HARTLEY" surnames
    # despite the HARTELY/HARTLEY spelling wobble), and Alfred Ernest Luke's
    # unnamed wife.
    individuals, families = parsed
    spouses = {s for f in families for s in (f.husband_id, f.wife_id) if s}
    for given, sur in (("Edward Greenslade", "PEARCE"), ("George", "HARTELY")):
        person = _find(individuals, given, sur)[0]
        assert person.id in spouses
    # George Hartely is paired specifically with Lorna Kerslake.
    by_id = {i.id: i for i in individuals}
    george = _find(individuals, "George", "HARTELY")[0]
    wives = {by_id[f.wife_id].given_name for f in families
             if f.husband_id == george.id and f.wife_id}
    assert "Lorna Ruth" in wives


def test_nogen_married_daughter_filed_under_parents(parsed):
    # Jeanette (née Luke, married surname unknown) is a no-gen married daughter
    # of Jack (Alfred John) Luke — filed under his family via her maiden name.
    individuals, families = parsed
    jeanette = _find(individuals, "Jeanette", "LUKE")[0]
    father, _ = _parents(individuals, families, jeanette)
    assert father and father.given_name == "Alfred John"


# --- Stage 4: audit reconciliations ---

def test_nicholls_children_reconciled_to_named_mother(parsed):
    # Charles Nicholls's children are coded under his second marriage (to Lydia
    # Ann Thomas) but every child names its mother as "Mary Thomas" — his first
    # wife Mary Louisa (née Thomas). They must be filed under her.
    individuals, families = parsed
    by_id = {i.id: i for i in individuals}
    for given, sur in (("Ethel Marian", "PARKER"), ("Ada Selina", "NICHOLLS"),
                       ("William C.", "NICHOLLS"), ("Eva L.", "NICHOLLS")):
        child = _find(individuals, given, sur)[0]
        _, mother = _parents(individuals, families, child)
        assert mother is not None and mother.given_name == "Mary Louisa"
    # Charles + Lydia remains as a (now childless) real marriage.
    charles = _find(individuals, "Charles", "NICHOLLS")[0]
    lydia_fams = [f for f in families if f.husband_id == charles.id
                  and by_id.get(f.wife_id) and by_id[f.wife_id].given_name == "Lydia Ann"]
    assert lydia_fams and lydia_fams[0].child_ids == []


def test_ernest_collision_resolved_by_generation(parsed):
    # "Ernest Luke" is ambiguous: Alfred Ernest (nicknamed Ernest, b 1880) and
    # his great-nephew Ernest George (b 1909). Lillian Morgan's father must be
    # the generation-correct Alfred Ernest, not the younger Ernest George.
    individuals, families = parsed
    lillian = _find(individuals, "Lillian", "MORGAN")[0]
    father, _ = _parents(individuals, families, lillian)
    assert father and father.given_name == "Alfred Ernest"
    assert father.birth_date == "1880"


def test_no_duplicate_couple_families(parsed):
    # A path-code family and a name-linked child must not split one couple into
    # two families (e.g. Arthur Nagle + Lorna Kerslake, Walter Watson + Myrtle).
    individuals, families = parsed
    couples = [(f.husband_id, f.wife_id) for f in families
               if f.husband_id and f.wife_id]
    assert len(couples) == len(set(couples))
