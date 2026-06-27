"""Post-conversion integrity checks.

These run automatically on every conversion. Their primary job is to catch
*conversion* regressions — silently dropped links, cycles, impossible dates —
especially when the heuristics (tuned to one source file) meet a new file.

`errors` are structurally or chronologically impossible and should always be
zero. `warnings` are suspicious-but-possible (e.g. a parent older than 60) and
are surfaced for a human to glance at, not necessarily fixed.
"""

from __future__ import annotations

import re
from collections import defaultdict

from .reader import Family, Individual, _similar_surname


def _year(date: str | None) -> int | None:
    if not date:
        return None
    m = re.search(r"\b(\d{4})\b", date)
    return int(m.group(1)) if m else None


def _year_latest(date: str | None) -> int | None:
    """The latest 4-digit year in a date phrase — the most charitable reading of
    a range like ``BET 1800 AND 1809`` (1809). Used so an age check only fires
    when implausible even at the youngest possible birth year."""
    if not date:
        return None
    years = re.findall(r"\b(\d{4})\b", date)
    return int(years[-1]) if years else None


def validate(
    individuals: list[Individual], families: list[Family]
) -> dict[str, list[str]]:
    by_id = {i.id: i for i in individuals}
    errors: list[str] = []
    warnings: list[str] = []

    def name(iid: str | None) -> str:
        i = by_id.get(iid)
        return f"{i.given_name} {i.surname}".strip() if i else f"<{iid}>"

    # --- structural integrity ---
    ids = [i.id for i in individuals]
    if len(ids) != len(set(ids)):
        errors.append("duplicate individual ids")
    fam_ids = [f.id for f in families]
    if len(fam_ids) != len(set(fam_ids)):
        errors.append("duplicate family ids")
    id_set = set(ids)
    for f in families:
        for ptr in (f.husband_id, f.wife_id, *f.child_ids):
            if ptr and ptr not in id_set:
                errors.append(f"{f.id} references missing individual {ptr}")
        if not f.husband_id and not f.wife_id and not f.child_ids:
            errors.append(f"{f.id} is empty (no spouses, no children)")
        if len(f.child_ids) != len(set(f.child_ids)):
            errors.append(f"{f.id} lists a child more than once")

    # --- parent/child maps ---
    parents_of: dict[str, set[str]] = {}
    for f in families:
        for c in f.child_ids:
            parents_of.setdefault(c, set()).update(
                p for p in (f.husband_id, f.wife_id) if p
            )

    # cycles (someone their own ancestor)
    def ancestor_of_self(start: str) -> bool:
        seen, stack = set(), [start]
        while stack:
            for p in parents_of.get(stack.pop(), ()):
                if p == start:
                    return True
                if p not in seen:
                    seen.add(p)
                    stack.append(p)
        return False

    for c in parents_of:
        if ancestor_of_self(c):
            errors.append(f"parent-child cycle involving {name(c)}")

    # --- chronology ---
    for i in individuals:
        b, d = _year(i.birth_date), _year(i.death_date)
        if b and d:
            if b > d:
                errors.append(f"{name(i.id)}: born {i.birth_date} after death {i.death_date}")
            elif d - b > 105:
                warnings.append(f"{name(i.id)}: lifespan {d - b}y ({i.birth_date}–{i.death_date})")

    for f in families:
        my = _year(f.marriage_date)
        for sid in (f.husband_id, f.wife_id):
            s = by_id.get(sid)
            if not (s and my):
                continue
            sb, sd = _year(s.birth_date), _year(s.death_date)
            if sb and my < sb:
                errors.append(
                    f"{f.id}: marriage {f.marriage_date} before "
                    f"{name(sid)} born {s.birth_date}")
            if sd and my > sd:
                errors.append(
                    f"{f.id}: marriage {f.marriage_date} after "
                    f"{name(sid)} died {s.death_date}")
        for sid, role in ((f.husband_id, "father"), (f.wife_id, "mother")):
            s = by_id.get(sid)
            pb = _year(s.birth_date) if s else None
            if not pb:
                continue
            for c in f.child_ids:
                cb = _year(by_id[c].birth_date) if c in by_id else None
                if not cb:
                    continue
                age = cb - pb
                if age < 0:
                    errors.append(
                        f"{name(c)} born {by_id[c].birth_date} before "
                        f"{role} {name(sid)} ({s.birth_date})")
                elif age < 13 or age > 60:
                    warnings.append(f"{name(sid)} was {age} when {role} of {name(c)}")
                elif role == "mother" and age > 50:
                    # A father may credibly be over 50; a mother bearing a child
                    # past ~50 is biologically implausible and usually signals a
                    # mis-grouped generation (grandchildren listed flat among
                    # children) or an approximate maternal birth year — surfaced
                    # for review, not silently restructured. Use the *latest*
                    # possible maternal birth year so a range like
                    # "BET 1800 AND 1809" is judged charitably and does not
                    # false-fire (Sarah Clow, plausibly born 1809).
                    late = _year_latest(s.birth_date)
                    if late is not None and cb - late > 50:
                        warnings.append(
                            f"{name(sid)} was {cb - late} when mother of "
                            f"{name(c)} (implausible maternal age)")

    # --- sex/role consistency ---
    for f in families:
        h, w = by_id.get(f.husband_id), by_id.get(f.wife_id)
        if h and w and h.sex == "F":
            errors.append(f"{f.id}: husband {name(f.husband_id)} is female")
        if w and w.sex == "M":
            errors.append(f"{f.id}: wife {name(f.wife_id)} is male")

    return {"errors": errors, "warnings": warnings}


def _given_tokens(given: str | None) -> set[str]:
    return set((given or "").lower().split("(")[0].replace(".", " ").split())


def cross_file_audit(
    individuals: list[Individual], families: list[Family]
) -> list[str]:
    """Report potential cross-file over- and under-merges for human review.

    Run by :mod:`merge` after the per-file trees are combined. Neither category
    is necessarily an error:

    * **Over-merge** — an individual with more than one *distinct* parent couple
      (FAMC > 1). Usually the genealogist's own cross-chart bridging (a person
      charted in two lineage charts) or the downstream effect of a deliberately
      kept conflict, but a genuine mistaken unification would also show here.
    * **Under-merge** — two records that share surname (allowing a single
      transcription typo), a compatible given name, the same birth year and a
      compatible father, yet were *not* unified. Usually a deliberately kept
      conflict, occasionally a missed match worth a look.

    Both are surfaced as a spot-check list, not auto-corrected
    (preserve-don't-assert).
    """
    by_id = {i.id: i for i in individuals}

    def name(iid: str | None) -> str:
        i = by_id.get(iid)
        return f"{i.given_name} {i.surname}".strip() if i else f"<{iid}>"

    lines: list[str] = []

    # --- over-merge: a child with two or more distinct parent couples ---
    famc: dict[str, list[Family]] = defaultdict(list)
    for f in families:
        for c in f.child_ids:
            famc[c].append(f)
    over: list[str] = []
    for cid, fl in famc.items():
        couples = {(f.husband_id, f.wife_id) for f in fl}
        if len(couples) > 1:
            desc = "; ".join(
                f"{name(h)} + {name(w)}" for h, w in sorted(couples, key=str))
            over.append(f"  - {name(cid)}: {len(couples)} parent couples — {desc}")
    lines.append(f"OVER-MERGE — individuals with >1 parent couple ({len(over)}):")
    lines += over or ["  (none)"]

    # --- under-merge: same-identity records left distinct ---
    father_first: dict[str, str] = {}
    for f in families:
        h = by_id.get(f.husband_id)
        if h:
            fw = (h.given_name or "").split()
            for c in f.child_ids:
                father_first.setdefault(c, fw[0].lower() if fw else "?")

    def byear(d: str | None) -> int | None:
        m = re.search(r"\b(1[5-9]\d\d|20\d\d)\b", d or "")
        return int(m.group(1)) if m else None

    buckets: dict[tuple[str, int], list[Individual]] = defaultdict(list)
    for i in individuals:
        y = byear(i.birth_date)
        toks = (i.given_name or "").lower().split("(")[0].split()
        if y is not None and toks:
            buckets[(toks[0], y)].append(i)

    under: list[str] = []
    for members in buckets.values():
        for a in range(len(members)):
            for b in range(a + 1, len(members)):
                x, y = members[a], members[b]
                xs, ys = (x.surname or "").upper(), (y.surname or "").upper()
                if not xs or not ys or not (xs == ys or _similar_surname(xs, ys)):
                    continue
                tx, ty = _given_tokens(x.given_name), _given_tokens(y.given_name)
                if not (tx <= ty or ty <= tx):
                    continue
                fx = father_first.get(x.id, "?")
                fy = father_first.get(y.id, "?")
                if fx != "?" and fy != "?" and fx != fy:
                    continue
                under.append(
                    f"  - {name(x.id)} (b.{x.birth_date}) vs "
                    f"{name(y.id)} (b.{y.birth_date})")
    lines += ["", f"UNDER-MERGE — same-identity records left separate "
              f"({len(under)}):"]
    lines += under or ["  (none)"]
    return lines
