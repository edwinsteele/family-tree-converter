"""Tests for the cross-file merge layer.

The synthetic unit tests build small Individual/Family graphs by hand and so run
anywhere. The integration tests need the (gitignored) source spreadsheets and
skip when they are absent.
"""

import hashlib

import pytest

from family_tree_converter.golden import (
    CONVERTED_FILES,
    INPUT_DIR,
    load_golden,
)
from family_tree_converter.merge import (
    Tree,
    _renumber,
    merge_files,
    merge_trees,
)
from family_tree_converter.reader import (
    Family,
    Individual,
    profile_for,
    read_spreadsheet,
)
from family_tree_converter.validate import cross_file_audit, validate
from family_tree_converter.writer import render_gedcom

# --------------------------------------------------------------------------
# Synthetic helpers
# --------------------------------------------------------------------------

def _ind(iid, given, surname, **kw):
    return Individual(id=iid, given_name=given, surname=surname, **kw)


def _fam(fid, husband=None, wife=None, children=(), **kw):
    return Family(id=fid, husband_id=husband, wife_id=wife,
                  child_ids=list(children), **kw)


def _tree(tag, order, individuals, families):
    return Tree(name=tag, tag=tag, order=order,
                individuals=individuals, families=families)


# --------------------------------------------------------------------------
# Namespacing
# --------------------------------------------------------------------------

def test_renumber_prefixes_ids_and_references():
    inds = [_ind("I1", "Ann", "SMITH", adopted_famc={"F1"})]
    fams = [_fam("F1", husband="I2", wife="I3", children=["I1"])]
    _renumber(inds, fams, "BG")
    assert inds[0].id == "BG_I1"
    assert inds[0].adopted_famc == {"BG_F1"}
    assert fams[0].id == "BG_F1"
    assert fams[0].husband_id == "BG_I2"
    assert fams[0].child_ids == ["BG_I1"]


def test_single_tree_merge_is_passthrough():
    inds = [_ind("I1", "Ann", "SMITH")]
    fams = [_fam("F1", children=["I1"])]
    res = merge_trees([_tree("A", 0, inds, fams)])
    assert res.individuals == inds
    assert res.families == fams
    assert res.clean_merges == 0


# --------------------------------------------------------------------------
# Cross-file person merge
# --------------------------------------------------------------------------

def test_clean_match_unifies_and_repoints():
    # Same James Steele b.1861 in two trees, each naming father "James Steele".
    a = _tree("A", 0,
              [_ind("A_I1", "James", "STEELE", birth_date="1861"),
               _ind("A_I2", "James", "STEELE"),
               _ind("A_I3", "Margaret", "STEELE")],
              [_fam("A_F1", husband="A_I2", wife="A_I3", children=["A_I1"])])
    b = _tree("B", 1,
              [_ind("B_I1", "James", "STEELE", birth_date="1861"),
               _ind("B_I2", "James", "STEELE"),
               _ind("B_I3", "Margaret", "STEELE")],
              [_fam("B_F1", husband="B_I2", wife="B_I3", children=["B_I1"])])
    res = merge_trees([a, b])
    assert res.clean_merges == 1
    jameses = [i for i in res.individuals
               if i.given_name == "James" and i.birth_date == "1861"]
    assert len(jameses) == 1                      # unified to one record
    canon = jameses[0]
    assert canon.id == "A_I1"                      # earliest tree is canonical
    # Both child links now point at the canonical record.
    childfams = [f for f in res.families if canon.id in f.child_ids]
    assert childfams
    # The dateless fathers unify via the shared child (up-propagation).
    assert res.ancestors_unified == 1
    fathers = [i for i in res.individuals
               if i.given_name == "James" and i.birth_date is None]
    assert len(fathers) == 1


def test_different_surname_does_not_merge():
    a = _tree("A", 0, [_ind("A_I1", "Charles", "HICKS", birth_date="1837")], [])
    b = _tree("B", 1, [_ind("B_I1", "Charles", "MULLINS", birth_date="1837")], [])
    res = merge_trees([a, b])
    assert res.clean_merges == 0
    assert len(res.individuals) == 2


def test_conflict_keeps_both_and_flags():
    # One person, agreeing father, but a different recorded mother → keep both.
    a = _tree("A", 0,
              [_ind("A_D", "David Lionel", "LONG", birth_date="9 MAY 1961"),
               _ind("A_F", "Geoffrey", "LONG"),
               _ind("A_M", "Jocelyn", "TILLEY")],
              [_fam("A_F1", husband="A_F", wife="A_M", children=["A_D"])])
    b = _tree("B", 1,
              [_ind("B_D", "David Lionel", "LONG", birth_date="9 MAY 1961"),
               _ind("B_F", "Geoffrey", "LONG"),
               _ind("B_M", "Joy", "")],
              [_fam("B_F1", husband="B_F", wife="B_M", children=["B_D"])])
    res = merge_trees([a, b])
    assert res.conflicts == 1
    assert res.clean_merges == 0
    davids = [i for i in res.individuals if i.given_name == "David Lionel"]
    assert len(davids) == 2                        # both kept
    assert all(any("kept separate" in n.lower() for n in d.note_list)
               for d in davids)


def test_shared_couple_families_collapse():
    couple_a = [_ind("A_H", "John", "SMITH", birth_date="1900"),
                _ind("A_W", "Mary", "SMITH", birth_date="1902"),
                _ind("A_C", "Anne", "SMITH", birth_date="1925")]
    couple_b = [_ind("B_H", "John", "SMITH", birth_date="1900"),
                _ind("B_W", "Mary", "SMITH", birth_date="1902"),
                _ind("B_C", "Beth", "SMITH", birth_date="1927")]
    a = _tree("A", 0, couple_a,
              [_fam("A_F1", husband="A_H", wife="A_W", children=["A_C"])])
    b = _tree("B", 1, couple_b,
              [_fam("B_F1", husband="B_H", wife="B_W", children=["B_C"])])
    res = merge_trees([a, b])
    # John and Mary unify; the duplicate family collapses onto one with both kids.
    families = [f for f in res.families if f.husband_id and f.wife_id]
    assert len(families) == 1
    assert set(families[0].child_ids) == {"A_C", "B_C"}
    assert res.families_collapsed == 1


def test_maiden_mother_unifies_via_shared_child():
    # A wife charted under her maiden name in one tree and married name in the
    # other still unifies when a shared child and matching husband confirm it.
    a = _tree("A", 0,
              [_ind("A_C", "James", "STEEL", birth_date="1861"),
               _ind("A_H", "James", "STEEL"),
               _ind("A_W", "Margaret", "STEEL")],
              [_fam("A_F1", husband="A_H", wife="A_W", children=["A_C"])])
    b = _tree("B", 1,
              [_ind("B_C", "James", "STEELE", birth_date="1861"),
               _ind("B_H", "James", "STEELE"),
               _ind("B_W", "Margaret", "BRUCE")],
              [_fam("B_F1", husband="B_H", wife="B_W", children=["B_C"])])
    res = merge_trees([a, b])
    margarets = [i for i in res.individuals if i.given_name == "Margaret"]
    assert len(margarets) == 1                     # one mother, not two
    assert "BRUCE" in margarets[0].married_surnames or \
        margarets[0].surname.upper() == "STEEL"


# --------------------------------------------------------------------------
# Cross-file audit
# --------------------------------------------------------------------------

def test_cross_file_audit_reports_under_and_over_merge():
    inds = [
        _ind("I1", "David", "LONG", birth_date="1961"),
        _ind("I2", "David", "LONG", birth_date="1961"),  # under-merge pair
        _ind("C", "Kid", "LONG"),
        _ind("H1", "A", "LONG"), _ind("W1", "B", "LONG"),
        _ind("H2", "C", "LONG"), _ind("W2", "D", "LONG"),
    ]
    fams = [
        _fam("F1", husband="H1", wife="W1", children=["C"]),
        _fam("F2", husband="H2", wife="W2", children=["C"]),  # over-merge
    ]
    lines = cross_file_audit(inds, fams)
    text = "\n".join(lines)
    assert "OVER-MERGE" in text and "Kid LONG" in text
    assert "UNDER-MERGE" in text and "David LONG" in text


# --------------------------------------------------------------------------
# Integration (need the private source spreadsheets)
# --------------------------------------------------------------------------

def _have_all():
    return all((INPUT_DIR / n).exists() for n in CONVERTED_FILES)


requires_sources = pytest.mark.skipif(
    not _have_all(), reason="source spreadsheets not present")


@requires_sources
@pytest.mark.parametrize("name", CONVERTED_FILES)
def test_single_file_merge_is_byte_identical(name):
    """merge.py on one file reproduces that file's per-file output exactly."""
    path = INPUT_DIR / name
    inds, fams = read_spreadsheet(path, profile_for(path))
    direct = render_gedcom(inds, fams)
    res = merge_files([path])
    assert render_gedcom(res.individuals, res.families) == direct
    golden = load_golden()
    if name in golden:
        h = hashlib.sha256(direct.encode("utf-8")).hexdigest()
        assert h == golden[name]


@requires_sources
def test_full_merge_honours_join_points_and_is_clean():
    paths = [INPUT_DIR / n for n in CONVERTED_FILES]
    res = merge_files(paths)

    # James Steele b.1861 is the J&DJ <-> Brc join point: one unified record.
    james = [i for i in res.individuals
             if i.given_name == "James" and i.surname.upper() == "STEELE"
             and i.birth_date == "1861"]
    assert len(james) == 1

    # Known shared people unify rather than duplicate.
    for given in ("Clive Norman", "Edna May", "Allan Richard Paul"):
        matches = [i for i in res.individuals if i.given_name == given
                   and i.surname.upper() == "STEELE"]
        assert len(matches) == 1, given

    # The one genuine conflict (David Long's two recorded mothers) is preserved.
    davids = [i for i in res.individuals
              if i.given_name == "David Lionel" and i.surname.upper() == "LONG"]
    assert len(davids) == 2
    assert res.conflicts == 1

    # The combined graph is structurally sound.
    result = validate(res.individuals, res.families)
    assert result["errors"] == []
    ids = {i.id for i in res.individuals}
    for f in res.families:
        for ptr in (f.husband_id, f.wife_id, *f.child_ids):
            assert ptr is None or ptr in ids
