"""Local golden-hash drift test.

Compares each converted file's current output hash against the gitignored
snapshot in tests/golden/hashes.json. A mismatch means the conversion changed —
intended after a verified improvement (re-run scripts/update_golden.py), a
regression otherwise. Files with no recorded golden (or no local source) skip.
"""

import pytest

from family_tree_converter.golden import (
    CONVERTED_FILES,
    INPUT_DIR,
    load_golden,
    output_hash,
)


@pytest.mark.parametrize("name", CONVERTED_FILES)
def test_output_matches_golden(name):
    src = INPUT_DIR / name
    if not src.exists():
        pytest.skip(f"source spreadsheet not present: {src}")
    golden = load_golden()
    if name not in golden:
        pytest.skip(
            f"no golden hash recorded for {name}; "
            f"run scripts/update_golden.py to create one")
    assert output_hash(src) == golden[name], (
        f"GEDCOM output for {name} drifted from the golden snapshot. "
        f"If intended, re-run scripts/update_golden.py.")
