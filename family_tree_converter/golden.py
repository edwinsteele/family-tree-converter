"""Local golden-hash snapshots, to flag output drift across refactors.

The family tree is not kept in a public repo, so we cannot commit a golden
``.ged`` file. Instead we store a sha256 of each converted file's output under a
gitignored directory and compare against it. Output is deterministic for a given
input (positional ids are assigned in a stable read order), so a changed hash
means the conversion changed — intended after a real improvement, a regression
otherwise.

Run ``scripts/update_golden.py`` to (re)record the current hashes after a change
you have verified is correct.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

from .reader import profile_for, read_spreadsheet
from .writer import render_gedcom

# The source spreadsheets that have been converted and must stay stable.
_ROOT = Path(__file__).resolve().parent.parent
INPUT_DIR = _ROOT / "data" / "input"
GOLDEN_PATH = _ROOT / "tests" / "golden" / "hashes.json"

CONVERTED_FILES = [
    "BlsGrnLivMcCl H.Tr. #305.xls",
    "C & A Stl H.Tree #81",
    "Hcks:Thos:Krsl H.Tr.#120",
    "Brc:Stl H.Tree #46",
    "J & D J Steele H.Tree #30",
    "Stiff:Taylor H.Tree #275",
]


def output_hash(path: Path) -> str:
    """sha256 of the GEDCOM text produced for ``path``."""
    individuals, families = read_spreadsheet(path, profile_for(path))
    return hashlib.sha256(render_gedcom(individuals, families).encode("utf-8")).hexdigest()


def load_golden() -> dict[str, str]:
    if GOLDEN_PATH.exists():
        return json.loads(GOLDEN_PATH.read_text(encoding="utf-8"))
    return {}


def save_golden(hashes: dict[str, str]) -> None:
    GOLDEN_PATH.parent.mkdir(parents=True, exist_ok=True)
    GOLDEN_PATH.write_text(json.dumps(hashes, indent=2, sort_keys=True) + "\n",
                           encoding="utf-8")


def current_hashes() -> dict[str, str]:
    """Hashes for every converted source file present locally."""
    out: dict[str, str] = {}
    for name in CONVERTED_FILES:
        src = INPUT_DIR / name
        if src.exists():
            out[name] = output_hash(src)
    return out
