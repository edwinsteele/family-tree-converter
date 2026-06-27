"""Cross-file merge layer — combine the per-file conversions into one GEDCOM.

The genealogist drew several overlapping "horizontal trees", each converted to
its own ``Individual``/``Family`` graph by :mod:`reader`. The same person often
appears in two or three of them (e.g. Clive Steele is charted in the main file,
in "C & A Stl" and in "Stiff:Taylor"), and MacFamilyTree's content auto-merge
leaves manual duplicates. This module folds the per-file graphs into ONE graph
with cross-file duplicates unified.

Design (user-agreed, see project memory):

* **Conflict-first.** When the same person appears in two trees with *conflicting*
  parents / dates / sex, the records are KEPT SEPARATE and flagged — a NOTE on
  each record plus a line in the merge report — never silently reconciled. A
  merge that quietly picks a winner destroys the very signal a human needs to
  spot a mislink. Only records that are *consistent* are unified.
* **Conservative identity.** Two records are candidates for the same person only
  when they share the within-file dedup key idea — ``(surname, given,
  father's-first-name, birth-year)`` — with the surname allowed to differ by a
  single transcription typo (STEEL/STEELE). A shared birth year alone, or a
  shared given name alone, is never enough.
* **Positional ids, namespaced per file.** Each tree keeps its positional
  ``I*``/``F*`` ids, prefixed with a short per-file tag so the combined graph has
  no collisions. A single-file merge therefore reproduces that file's output
  byte-for-byte (the deterministic name-slug ``_UID`` scheme stays parked).

The merge report is the user's targeted spot-check list: every clean cross-file
unification and every flagged conflict is listed, so the ~16 known shared people
can be verified in MacFamilyTree afterwards.
"""

from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

from .reader import (
    Family,
    Individual,
    _similar_surname,
    profile_for,
    read_spreadsheet,
)
from .validate import cross_file_audit, validate
from .writer import render_gedcom

_YEAR_RE = re.compile(r"\b(1[5-9]\d\d|20\d\d)\b")


def _birth_year(date: str | None) -> int | None:
    if not date:
        return None
    m = _YEAR_RE.search(date)
    return int(m.group(1)) if m else None


_MONTHS = ("JAN", "FEB", "MAR", "APR", "MAY", "JUN",
           "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")


def _is_full_date(date: str | None) -> bool:
    """A specific day/month/year date (as opposed to a bare year or a range)."""
    return bool(date) and any(mon in date for mon in _MONTHS) and "BET" not in date


def _given_tokens(given: str | None) -> set[str]:
    return set((given or "").lower().split("(")[0].replace(".", " ").split())


def _first_given(given: str | None) -> str:
    toks = (given or "").lower().split("(")[0].split()
    return toks[0] if toks else ""


def _first_word(name: str | None) -> str:
    name = (name or "").strip()
    if not name or name == "?":
        return "?"
    return name.split()[0].lower()


# ---------------------------------------------------------------------------
# Tree loading + id namespacing
# ---------------------------------------------------------------------------

# Short, stable per-file tags so combined ids stay readable (@BG_I1@ etc.) and
# never collide. Keyed by the FormatProfile.name each file resolves to.
_FILE_TAGS = {
    "BlsGrnLivMcCl": "BG",
    "C & A Stl": "CA",
    "Hcks:Thos:Krsl": "HK",
    "Brc:Stl": "BR",
    "J & D J Steele": "JD",
    "Stiff:Taylor": "ST",
}


@dataclass
class Tree:
    """One per-file conversion, ready to merge."""
    name: str           # FormatProfile.name (human-readable tree name)
    tag: str            # short id prefix for the combined graph
    order: int          # load order, used to pick the canonical record
    individuals: list[Individual]
    families: list[Family]


def _tag_for(name: str, used: set[str]) -> str:
    tag = _FILE_TAGS.get(name)
    if tag is None:
        # Derive a 2-letter tag from the name's alphanumerics; disambiguate.
        letters = re.sub(r"[^A-Za-z0-9]", "", name).upper() or "T"
        tag = letters[:2]
    base, n = tag, 2
    while tag in used:
        tag = f"{base}{n}"
        n += 1
    used.add(tag)
    return tag


def load_trees(paths: list[Path]) -> list[Tree]:
    """Convert each source file and namespace its ids for the combined graph."""
    trees: list[Tree] = []
    used: set[str] = set()
    single = len(paths) == 1
    for order, path in enumerate(paths):
        prof = profile_for(path)
        inds, fams = read_spreadsheet(path, prof)
        tag = _tag_for(prof.name, used)
        if not single:
            _renumber(inds, fams, tag)
        trees.append(Tree(name=prof.name, tag=tag, order=order,
                          individuals=inds, families=fams))
    return trees


def _renumber(individuals: list[Individual], families: list[Family],
              tag: str) -> None:
    """Prefix every id and every reference in place with ``tag``."""
    def t(i: str | None) -> str | None:
        return f"{tag}_{i}" if i else i

    for ind in individuals:
        ind.id = t(ind.id)
        ind.adopted_famc = {t(x) for x in ind.adopted_famc}
    for fam in families:
        fam.id = t(fam.id)
        fam.husband_id = t(fam.husband_id)
        fam.wife_id = t(fam.wife_id)
        fam.child_ids = [t(c) for c in fam.child_ids]


# ---------------------------------------------------------------------------
# Cross-file person merge
# ---------------------------------------------------------------------------

def _parent_names(individuals: list[Individual],
                  families: list[Family]) -> dict[str, tuple[str, str]]:
    """Map each individual id → (father display name, mother display name).

    Read from the FAMC family, so two records of the same person in different
    trees can be compared on the *names* of their parents (whose own records may
    not have merged).
    """
    by_id = {i.id: i for i in individuals}
    child_fam: dict[str, Family] = {}
    for fam in families:
        for c in fam.child_ids:
            child_fam.setdefault(c, fam)

    def _disp(ind: Individual | None) -> str:
        if ind is None:
            return ""
        return f"{ind.given_name} {ind.surname}".strip()

    out: dict[str, tuple[str, str]] = {}
    for ind in individuals:
        fam = child_fam.get(ind.id)
        if fam is None:
            out[ind.id] = ("", "")
            continue
        out[ind.id] = (_disp(by_id.get(fam.husband_id)),
                       _disp(by_id.get(fam.wife_id)))
    return out


def _surname_match(a: str | None, b: str | None) -> bool:
    au, bu = (a or "").upper(), (b or "").upper()
    if not au or not bu:
        return False
    return au == bu or _similar_surname(au, bu)


def _given_compatible(a: str | None, b: str | None) -> bool:
    ta, tb = _given_tokens(a), _given_tokens(b)
    if not ta or not tb:
        return False
    return ta <= tb or tb <= ta


def _father_compatible(fa: str, fb: str) -> bool:
    wa, wb = _first_word(fa), _first_word(fb)
    return wa == "?" or wb == "?" or wa == wb


def _same_person(a: Individual, b: Individual,
                 parents: dict[str, tuple[str, str]]) -> bool:
    """Conservative cross-file identity test (already share given+year bucket)."""
    if not _surname_match(a.surname, b.surname):
        return False
    if not _given_compatible(a.given_name, b.given_name):
        return False
    return _father_compatible(parents[a.id][0], parents[b.id][0])


def _conflicts(group: list[Individual],
               parents: dict[str, tuple[str, str]]) -> list[str]:
    """Hard contradictions that mean a group must be KEPT SEPARATE."""
    reasons: list[str] = []

    births = {i.birth_date for i in group if _is_full_date(i.birth_date)}
    if len(births) > 1:
        reasons.append("different birth dates (" + " vs ".join(sorted(births)) + ")")

    deaths = {i.death_date for i in group if _is_full_date(i.death_date)}
    if len(deaths) > 1:
        reasons.append("different death dates (" + " vs ".join(sorted(deaths)) + ")")

    sexes = {i.sex for i in group if i.sex}
    if len(sexes) > 1:
        reasons.append("different recorded sex")

    # Mother names: a contradiction only when both name a mother and the two
    # share no given-name token and the surnames are not a typo apart.
    mothers = [parents[i.id][1] for i in group if parents[i.id][1]]
    for x in range(len(mothers)):
        for y in range(x + 1, len(mothers)):
            mx, my = mothers[x], mothers[y]
            tx, ty = _given_tokens(mx), _given_tokens(my)
            sx = mx.split()[-1] if mx.split() else ""
            sy = my.split()[-1] if my.split() else ""
            if not (tx & ty) and not _similar_surname(sx.upper(), sy.upper()):
                reasons.append(f"different mothers ({mx} vs {my})")
                break
        else:
            continue
        break

    return reasons


@dataclass
class MergeResult:
    individuals: list[Individual]
    families: list[Family]
    report: list[str]
    clean_merges: int = 0
    conflicts: int = 0
    families_collapsed: int = 0
    ancestors_unified: int = 0


def _fill_blanks(dst: Individual, src: Individual) -> None:
    """Union ``src`` into the canonical ``dst`` without overwriting set data."""
    for attr in ("nickname", "birth_date", "birth_place", "death_date",
                 "death_place", "sex", "occupation", "notes"):
        if getattr(dst, attr) in (None, "") and getattr(src, attr):
            setattr(dst, attr, getattr(src, attr))
    if src.birth_is_christening and not dst.birth_date:
        dst.birth_is_christening = True
    for n in src.note_list:
        if n not in dst.note_list:
            dst.note_list.append(n)
    dst.lineage_lines |= src.lineage_lines
    for ms in src.married_surnames:
        if ms not in dst.married_surnames:
            dst.married_surnames.append(ms)


def _repoint(families: list[Family], remap: dict[str, str]) -> None:
    """Replace every individual reference per ``remap`` (old id → canonical)."""
    def r(i: str | None) -> str | None:
        return remap.get(i, i) if i else i

    for fam in families:
        fam.husband_id = r(fam.husband_id)
        fam.wife_id = r(fam.wife_id)
        seen: set[str] = set()
        fam.child_ids = [c for c in (r(c) for c in fam.child_ids)
                         if not (c in seen or seen.add(c))]


def merge_trees(trees: list[Tree]) -> MergeResult:
    """Fold per-file trees into one graph, unifying consistent duplicates."""
    individuals: list[Individual] = [i for t in trees for i in t.individuals]
    families: list[Family] = [f for t in trees for f in t.families]
    tree_of: dict[str, Tree] = {}
    for t in trees:
        for i in t.individuals:
            tree_of[i.id] = t

    report: list[str] = []
    if len(trees) == 1:
        return MergeResult(individuals, families, report)

    parents = _parent_names(individuals, families)

    # Bucket by (first given word, birth year); only dated, named records can
    # match. Within a bucket, group transitively by the conservative identity
    # test.
    buckets: dict[tuple[str, int], list[Individual]] = defaultdict(list)
    for ind in individuals:
        year = _birth_year(ind.birth_date)
        fw = _first_given(ind.given_name)
        if year is not None and fw:
            buckets[(fw, year)].append(ind)

    remap: dict[str, str] = {}          # merged-away id → canonical id
    removed: set[str] = set()
    conflict_ids: set[str] = set()      # ids deliberately kept separate
    clean_lines: list[str] = []
    conflict_lines: list[str] = []
    clean = conflicts = 0

    for members in buckets.values():
        if len(members) < 2:
            continue
        # Union-find over the bucket members.
        parent_ix = list(range(len(members)))

        def find(x: int) -> int:
            while parent_ix[x] != x:
                parent_ix[x] = parent_ix[parent_ix[x]]
                x = parent_ix[x]
            return x

        for a in range(len(members)):
            for b in range(a + 1, len(members)):
                if _same_person(members[a], members[b], parents):
                    parent_ix[find(a)] = find(b)

        groups: dict[int, list[Individual]] = defaultdict(list)
        for ix, ind in enumerate(members):
            groups[find(ix)].append(ind)

        for group in groups.values():
            trees_in = {tree_of[i.id].tag for i in group}
            if len(group) < 2 or len(trees_in) < 2:
                continue  # same person within one file is already deduped
            group.sort(key=lambda i: tree_of[i.id].order)
            label = _label(group[0])
            tags = ", ".join(tree_of[i.id].tag for i in group)
            reasons = _conflicts(group, parents)
            if reasons:
                conflicts += 1
                conflict_ids.update(i.id for i in group)
                conflict_lines.append(f"  - {label}: in {tags} — "
                                      + "; ".join(reasons) + " — KEPT SEPARATE")
                _flag_conflict(group, tree_of, parents, reasons)
            else:
                clean += 1
                canon = group[0]
                others = group[1:]
                clean_lines.append(f"  - {label}: {tags} → @{canon.id}@")
                for o in others:
                    _fill_blanks(canon, o)
                    remap[o.id] = canon.id
                    removed.add(o.id)
                _provenance_note(canon, [tree_of[i.id].name for i in group])

    _repoint(families, remap)
    individuals = [i for i in individuals if i.id not in removed]

    # Propagate merges UP the tree: once a child is unified across trees, the
    # two parent couples charting that child are the same couple (often dateless
    # married-in ancestors that no birth-year bucket could match) — unify them
    # too. This honours the documented join points: J&DJ's synthetic
    # "James Steele + Margaret Bruce" == Brc's James Steel b.1829 + Margaret;
    # the dateless "Edith Kerslake"/"George Stiff + Violet Taylor" parents == the
    # real dated records.
    before_f = len(families)
    anc_clean, anc_conflict, ancestors, up_conflicts = _propagate_up(
        individuals, families, conflict_ids, tree_of)
    conflicts += up_conflicts

    _collapse_families(families)
    collapsed = before_f - len(families)

    report += _format_report(trees, individuals, families,
                             clean, conflicts, collapsed, ancestors,
                             clean_lines, conflict_lines,
                             anc_clean, anc_conflict)
    return MergeResult(individuals, families, report,
                       clean_merges=clean, conflicts=conflicts,
                       families_collapsed=collapsed,
                       ancestors_unified=ancestors)


def _label(ind: Individual) -> str:
    y = _birth_year(ind.birth_date)
    yr = f" b.{y}" if y else ""
    return f"{ind.given_name} {ind.surname}{yr}"


def _provenance_note(canon: Individual, tree_names: list[str]) -> None:
    uniq = sorted(set(tree_names))
    note = ("Unified across genealogist trees: " + ", ".join(uniq) + ".")
    if note not in canon.note_list:
        canon.note_list.append(note)


def _flag_conflict(group: list[Individual], tree_of: dict[str, Tree],
                   parents: dict[str, tuple[str, str]],
                   reasons: list[str]) -> None:
    for ind in group:
        others = [o for o in group if o is not ind]
        where = ", ".join(sorted({tree_of[o.id].name for o in others}))
        note = (f"Possibly the same person as {_label(others[0])} recorded in "
                f"the {where} tree; kept separate for review because the records "
                f"differ: {'; '.join(reasons)}.")
        if note not in ind.note_list:
            ind.note_list.append(note)


def _side(x: Individual | None, y: Individual | None,
          conflict_ids: set[str]) -> str:
    """Classify a husband-vs-husband or wife-vs-wife comparison.

    Returns one of: ``same`` (identical record), ``merge`` (same given + same/
    typo surname), ``maiden`` (same given, different surname — a maiden/married
    pair, valid only for a wife), ``conflict`` (disjoint given names — different
    people), or ``incompatible`` (a record missing, or one is a flagged
    conflict, so do not act).
    """
    if x is None or y is None:
        return "incompatible"
    if x.id == y.id:
        return "same"
    if x.id in conflict_ids or y.id in conflict_ids:
        return "incompatible"
    gc = _given_compatible(x.given_name, y.given_name)
    if gc and _surname_match(x.surname, y.surname):
        return "merge"
    if gc:
        return "maiden"
    return "conflict"


def _propagate_up(individuals: list[Individual], families: list[Family],
                  conflict_ids: set[str],
                  tree_of: dict[str, Tree]) -> tuple[list[str], list[str], int, int]:
    """Unify duplicate parent couples revealed by a shared cross-file child.

    Iterates to a fixpoint: each round finds parent families that name a common
    child and have name-compatible husbands and wives, merges those spouses, and
    lets :func:`_collapse_families` fold the now-identical couples. Conservative
    — a husband must match by surname (men keep theirs); a wife may differ only
    as maiden vs married; disjoint wife names are flagged, not merged.
    """
    order_of = {i.id: tree_of[i.id].order for i in individuals if i.id in tree_of}
    clean_lines: list[str] = []
    conflict_lines: list[str] = []
    clean = conflicts = 0
    flagged: set[tuple[str, str]] = set()

    while True:
        by_id = {i.id: i for i in individuals}
        famc: dict[str, list[Family]] = defaultdict(list)
        for fam in families:
            for c in fam.child_ids:
                famc[c].append(fam)

        remap: dict[str, str] = {}
        removed: set[str] = set()
        touched: set[str] = set()

        def schedule(a: Individual, b: Individual, maiden: bool) -> Individual:
            nonlocal clean
            canon, other = (a, b) if order_of.get(a.id, 0) <= order_of.get(
                b.id, 0) else (b, a)
            _fill_blanks(canon, other)
            if maiden and other.surname and (
                    other.surname.upper() != (canon.surname or "").upper()
                    and other.surname not in canon.married_surnames):
                canon.married_surnames.append(other.surname)
            remap[other.id] = canon.id
            removed.add(other.id)
            touched.update({canon.id, other.id})
            _provenance_note(canon, [tree_of[canon.id].name, tree_of[other.id].name])
            return canon

        for fl in famc.values():
            if len(fl) < 2:
                continue
            for a in range(len(fl)):
                for b in range(a + 1, len(fl)):
                    f1, f2 = fl[a], fl[b]
                    if f1.id in removed or f2.id in removed:
                        continue
                    h1, h2 = by_id.get(f1.husband_id), by_id.get(f2.husband_id)
                    w1, w2 = by_id.get(f1.wife_id), by_id.get(f2.wife_id)
                    if not (h1 and h2 and w1 and w2):
                        continue
                    ids = {h1.id, h2.id, w1.id, w2.id}
                    if ids & touched:
                        continue  # serialise overlapping merges across rounds
                    hs = _side(h1, h2, conflict_ids)
                    if hs not in ("same", "merge"):
                        continue
                    ws = _side(w1, w2, conflict_ids)
                    if ws == "conflict":
                        key = tuple(sorted((w1.id, w2.id)))
                        if key not in flagged:
                            flagged.add(key)
                            conflicts += 1
                            conflict_lines.append(
                                f"  - {_label(w1)} vs {_label(w2)} (mothers of "
                                f"{_label(by_id.get(_canon_child(f1, remap)))}) "
                                "— KEPT SEPARATE")
                            _flag_conflict([w1, w2], tree_of,
                                           {w1.id: ("", ""), w2.id: ("", "")},
                                           ["different recorded names for the "
                                            "same person's mother"])
                        continue
                    if ws not in ("same", "merge", "maiden"):
                        continue
                    if hs == "merge":
                        schedule(h1, h2, maiden=False)
                    if ws in ("merge", "maiden"):
                        schedule(w1, w2, maiden=(ws == "maiden"))
                    if hs == "merge" or ws in ("merge", "maiden"):
                        clean += 1
                        clean_lines.append(
                            f"  - {_label(h1)} + {_label(w1)} "
                            f"({tree_of[h1.id].tag}+{tree_of[h2.id].tag})")

        if not remap:
            break
        _repoint(families, remap)
        individuals[:] = [i for i in individuals if i.id not in removed]
        _collapse_families(families)

    return clean_lines, conflict_lines, clean, conflicts


def _canon_child(fam: Family, remap: dict[str, str]) -> str:
    for c in fam.child_ids:
        return remap.get(c, c)
    return ""


def _collapse_families(families: list[Family]) -> int:
    """Merge families that name the same (husband, wife) pair after person merge.

    The same couple charted in two trees (e.g. Allan & Adrienne Steele appear in
    three) yields duplicate families once both spouses are unified — collapse
    them onto the first, unioning children, marriage facts and notes.
    """
    canon_by_pair: dict[tuple[str, str], Family] = {}
    removed: set[str] = set()
    collapsed = 0
    for fam in families:
        if not fam.husband_id or not fam.wife_id:
            continue
        key = (fam.husband_id, fam.wife_id)
        first = canon_by_pair.get(key)
        if first is None:
            canon_by_pair[key] = fam
            continue
        # Fold this duplicate into the first family for the couple.
        for c in fam.child_ids:
            if c not in first.child_ids:
                first.child_ids.append(c)
        if not first.marriage_date and fam.marriage_date:
            first.marriage_date = fam.marriage_date
        if not first.marriage_place and fam.marriage_place:
            first.marriage_place = fam.marriage_place
        first.divorced = first.divorced or fam.divorced
        for n in fam.note_list:
            if n not in first.note_list:
                first.note_list.append(n)
        removed.add(fam.id)
        collapsed += 1
    if removed:
        families[:] = [f for f in families if f.id not in removed]
    return collapsed


def _format_report(trees: list[Tree], individuals: list[Individual],
                   families: list[Family], clean: int, conflicts: int,
                   collapsed: int, ancestors: int, clean_lines: list[str],
                   conflict_lines: list[str], anc_clean: list[str],
                   anc_conflict: list[str]) -> list[str]:
    before_i = sum(len(t.individuals) for t in trees)
    before_f = sum(len(t.families) for t in trees)
    lines = [
        "CROSS-FILE MERGE REPORT",
        "=" * 60,
        "Trees merged (tag = id prefix):",
    ]
    for t in trees:
        lines.append(f"  {t.tag} = {t.name} "
                     f"({len(t.individuals)} individuals, {len(t.families)} families)")
    lines += [
        "",
        f"Individuals: {before_i} → {len(individuals)} "
        f"({before_i - len(individuals)} unified)",
        f"Families:    {before_f} → {len(families)} "
        f"({collapsed} duplicate couples collapsed)",
        "",
        "=" * 60,
        f"CLEAN CROSS-FILE MATCHES ({clean}) — unified into one record",
        "=" * 60,
    ]
    lines += clean_lines or ["  (none)"]
    lines += [
        "",
        "=" * 60,
        f"ANCESTOR COUPLES UNIFIED VIA A SHARED CHILD ({ancestors})",
        "=" * 60,
        "Married-in parents (often dateless) merged because a unified child "
        "names them.",
    ]
    lines += anc_clean or ["  (none)"]
    lines += [
        "",
        "=" * 60,
        f"CONFLICTS ({conflicts}) — same identity, differing data; KEPT SEPARATE",
        "=" * 60,
        "Review each in MacFamilyTree; the records carry a cross-reference NOTE.",
    ]
    lines += (conflict_lines + anc_conflict) or ["  (none)"]

    lines += [
        "",
        "=" * 60,
        "CROSS-FILE AUDIT (review — not necessarily errors)",
        "=" * 60,
    ]
    lines += cross_file_audit(individuals, families)

    result = validate(individuals, families)
    lines += [
        "",
        "=" * 60,
        "INTEGRITY VALIDATION OF THE COMBINED GRAPH",
        "=" * 60,
        f"ERRORS (impossible — should be 0): {len(result['errors'])}",
    ]
    lines += [f"  - {e}" for e in result["errors"]]
    lines.append(f"WARNINGS (suspicious but possible): {len(result['warnings'])}")
    lines += [f"  - {w}" for w in result["warnings"]]
    return lines


# ---------------------------------------------------------------------------
# Public entry points
# ---------------------------------------------------------------------------

def merge_files(paths: list[Path]) -> MergeResult:
    return merge_trees(load_trees(paths))


def main() -> None:
    import argparse
    import sys

    parser = argparse.ArgumentParser(
        description="Merge several converted family-tree spreadsheets into one "
                    "combined GEDCOM, unifying cross-file duplicates.")
    parser.add_argument("inputs", type=Path, nargs="+",
                        help="Source spreadsheet files to merge")
    parser.add_argument("-o", "--output", type=Path,
                        default=Path("data/output/combined.ged"),
                        help="Combined GEDCOM output path "
                             "(default: data/output/combined.ged, gitignored)")
    args = parser.parse_args()

    for p in args.inputs:
        if not p.exists():
            print(f"Error: input file not found: {p}", file=sys.stderr)
            sys.exit(1)

    result = merge_files(args.inputs)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        render_gedcom(result.individuals, result.families), encoding="utf-8")
    report_path = args.output.with_suffix(".merge-report.txt")
    report_path.write_text("\n".join(result.report) + "\n", encoding="utf-8")

    print(f"Merged {len(args.inputs)} trees → {len(result.individuals)} "
          f"individuals, {len(result.families)} families.")
    print(f"  {result.clean_merges} clean cross-file matches, "
          f"{result.conflicts} conflicts kept separate, "
          f"{result.families_collapsed} duplicate couples collapsed.")
    print(f"Output: {args.output}")
    print(f"Report: {report_path}")


if __name__ == "__main__":
    main()
