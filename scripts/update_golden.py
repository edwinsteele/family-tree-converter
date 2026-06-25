"""Record the current GEDCOM output hashes as the local golden snapshot.

Run after a change you have verified produces correct output:

    PYTHONPATH=. uv run python scripts/update_golden.py

The snapshot lives at tests/golden/hashes.json (gitignored). The golden-hash
test then fails if any later refactor changes a file's output unexpectedly.
"""

from family_tree_converter.golden import GOLDEN_PATH, current_hashes, load_golden, save_golden


def main() -> None:
    old = load_golden()
    new = current_hashes()
    if not new:
        print("No source spreadsheets present locally; nothing to record.")
        return
    for name, h in sorted(new.items()):
        prev = old.get(name)
        status = "unchanged" if prev == h else ("NEW" if prev is None else "CHANGED")
        print(f"  {status:9} {name}: {h[:16]}")
    save_golden(new)
    print(f"Wrote {len(new)} hash(es) to {GOLDEN_PATH}")


if __name__ == "__main__":
    main()
