"""Hardening reports run alongside conversion.

Unlike :mod:`validate` (which asserts structural/chronological integrity), these
checks surface *latent* risks for a human to review — silent data loss, name
collisions that could become mislinks, minted placeholder records, and
generation inconsistencies in name-linked families. Every function here only
*reports*; none mutates the individuals or families. They consume the
``diagnostics`` dict that :func:`reader.read_spreadsheet` optionally populates.

The CLI appends their output to the conversion's ``*.report.txt``.
"""

from __future__ import annotations

import re
from collections import Counter

from .reader import Family, Individual

# Buckets a row can fall into; the person-bearing ones must be consumed by a
# pass, the others are expected to leave no individual behind.
_PERSON_BUCKETS = ("coded-person", "gen-person", "no-gen", "block", "loose")
_NON_PERSON_BUCKETS = ("marriage", "blank/layout")


def _year(date: str | None) -> str:
    m = re.search(r"\b(\d{4})\b", date or "")
    return m.group(1) if m else "?"


def _first_token(given: str | None) -> str:
    g = (given or "").lower().split("(")[0].split()
    return g[0] if g else ""


def _describe(ind: Individual | None) -> str:
    if ind is None:
        return "<missing>"
    return f"{ind.id} {ind.given_name} /{ind.surname}/ (b.{_year(ind.birth_date)})"


# ---------------------------------------------------------------------------
# 1. Row-coverage accounting
# ---------------------------------------------------------------------------
def row_coverage(diag: dict) -> tuple[Counter, list[str]]:
    """Tally every data row by bucket and list name-bearing rows no pass
    consumed. Unaccounted rows are silent data-loss suspects (warnings)."""
    row_class: dict[int, str] = diag.get("row_class", {})
    row_text: dict[int, tuple[str, str]] = diag.get("row_text", {})
    consumed: set[int] = diag.get("consumed_rows", set())

    tally = Counter(row_class.values())
    unaccounted: list[str] = []
    for r in sorted(row_class):
        bucket = row_class[r]
        if bucket in _NON_PERSON_BUCKETS:
            continue
        if r in consumed:
            continue
        # A row bearing a surname or given name that no pass turned into output.
        if r in row_text:
            surname, given = row_text[r]
            unaccounted.append(
                f"row {r} [{bucket}]: surname={surname!r} given={given!r}")
    return tally, unaccounted


# ---------------------------------------------------------------------------
# 2. Name-collision report
# ---------------------------------------------------------------------------
def name_collisions(individuals: list[Individual]) -> list[str]:
    """Every (surname, first-given-token) shared by >=2 individuals — the
    hotspots where a name-based link could attach to the wrong person."""
    groups: dict[tuple[str, str], list[Individual]] = {}
    for i in individuals:
        key = ((i.surname or "").upper(), _first_token(i.given_name))
        if key == ("", ""):
            continue
        groups.setdefault(key, []).append(i)

    lines: list[str] = []
    for (surname, token), members in sorted(groups.items()):
        if len(members) < 2:
            continue
        who = ", ".join(_describe(m) for m in members)
        lines.append(f"{token or '?'} /{surname or '?'}/ ×{len(members)}: {who}")
    return lines


# ---------------------------------------------------------------------------
# 3. Synthetic / placeholder manifest
# ---------------------------------------------------------------------------
def synthetic_manifest(
    diag: dict, individuals: list[Individual], families: list[Family]
) -> list[str]:
    """List every synthetic parent and every ``[Unnamed]`` person with their
    family context, so a human can spot a missed real person or a duplicate."""
    by_id = {i.id: i for i in individuals}
    spouse_fams: dict[str, list[Family]] = {}
    child_fams: dict[str, list[Family]] = {}
    for f in families:
        for sid in (f.husband_id, f.wife_id):
            if sid:
                spouse_fams.setdefault(sid, []).append(f)
        for cid in f.child_ids:
            child_fams.setdefault(cid, []).append(f)

    def _context(iid: str) -> str:
        bits: list[str] = []
        for f in spouse_fams.get(iid, []):
            other = f.wife_id if f.husband_id == iid else f.husband_id
            kids = [by_id.get(c) for c in f.child_ids]
            kid_str = ", ".join(
                f"{k.given_name} {k.surname}".strip() for k in kids if k
            )
            bits.append(
                f"spouse={_describe(by_id.get(other)) if other else 'none'}"
                f"; children=[{kid_str}]")
        for f in child_fams.get(iid, []):
            bits.append(
                f"parents={_describe(by_id.get(f.husband_id))} & "
                f"{_describe(by_id.get(f.wife_id))}")
        return " | ".join(bits) if bits else "no family links"

    lines: list[str] = []
    for label, ids in (("SYNTHETIC PARENT", diag.get("synthetic_ids", set())),
                       ("PLACEHOLDER", diag.get("placeholder_ids", set()))):
        for iid in sorted(ids, key=lambda x: (len(x), x)):
            ind = by_id.get(iid)
            if ind is None:
                continue  # merged away before output; nothing to review
            lines.append(f"{label}: {_describe(ind)} — {_context(iid)}")
    return lines


# ---------------------------------------------------------------------------
# 4. Generation-consistency on name-linked families
# ---------------------------------------------------------------------------
def generation_consistency(
    diag: dict, individuals: list[Individual], families: list[Family]
) -> list[str]:
    """For families built by name/adjacency (Pass 4b / Pass 6), verify a parent
    sits exactly one generation above each child whose generation is known."""
    gen: dict[str, float] = diag.get("generation_by_id", {})
    name_linked: set[str] = diag.get("name_linked_family_ids", set())
    by_id = {i.id: i for i in individuals}

    lines: list[str] = []
    for f in families:
        if f.id not in name_linked:
            continue
        for pid, role in ((f.husband_id, "father"), (f.wife_id, "mother")):
            pg = gen.get(pid)
            if pg is None:
                continue
            for cid in f.child_ids:
                cg = gen.get(cid)
                if cg is None:
                    continue
                if pg != cg + 1:
                    lines.append(
                        f"{f.id}: {role} {_describe(by_id.get(pid))} gen {pg:g} "
                        f"vs child {_describe(by_id.get(cid))} gen {cg:g} "
                        f"(expected parent gen {cg + 1:g})")
    return lines


# ---------------------------------------------------------------------------
# Report assembly
# ---------------------------------------------------------------------------
def run_checks(
    individuals: list[Individual], families: list[Family], diag: dict
) -> list[str]:
    """Build the hardening-report section lines for a conversion."""
    out: list[str] = []

    tally, unaccounted = row_coverage(diag)
    out.append("ROW COVERAGE")
    for bucket in (*_PERSON_BUCKETS, *_NON_PERSON_BUCKETS):
        if tally.get(bucket):
            out.append(f"  {bucket}: {tally[bucket]}")
    out.append(f"  unaccounted name-bearing rows: {len(unaccounted)}")
    out += [f"    - {u}" for u in unaccounted]
    out.append("")

    collisions = name_collisions(individuals)
    out.append(f"NAME COLLISIONS (surname + first given token, >=2): "
               f"{len(collisions)}")
    out += [f"  - {c}" for c in collisions]
    out.append("")

    manifest = synthetic_manifest(diag, individuals, families)
    out.append(f"SYNTHETIC / PLACEHOLDER MANIFEST: {len(manifest)}")
    out += [f"  - {m}" for m in manifest]
    out.append("")

    gen_issues = generation_consistency(diag, individuals, families)
    out.append(f"GENERATION CONSISTENCY (name-linked families): "
               f"{len(gen_issues)} issue(s)")
    out += [f"  - {g}" for g in gen_issues]

    return out
