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

from .reader import Family, Individual


def _year(date: str | None) -> int | None:
    if not date:
        return None
    m = re.search(r"\b(\d{4})\b", date)
    return int(m.group(1)) if m else None


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
                errors.append(f"{f.id}: marriage {f.marriage_date} before {name(sid)} born {s.birth_date}")
            if sd and my > sd:
                errors.append(f"{f.id}: marriage {f.marriage_date} after {name(sid)} died {s.death_date}")
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
                    errors.append(f"{name(c)} born {by_id[c].birth_date} before {role} {name(sid)} ({s.birth_date})")
                elif age < 13 or age > 60:
                    warnings.append(f"{name(sid)} was {age} when {role} of {name(c)}")

    # --- sex/role consistency ---
    for f in families:
        h, w = by_id.get(f.husband_id), by_id.get(f.wife_id)
        if h and w and h.sex == "F":
            errors.append(f"{f.id}: husband {name(f.husband_id)} is female")
        if w and w.sex == "M":
            errors.append(f"{f.id}: wife {name(f.wife_id)} is male")

    return {"errors": errors, "warnings": warnings}
