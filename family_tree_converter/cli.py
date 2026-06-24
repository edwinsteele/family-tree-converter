"""Command-line entry point for the family tree converter."""

import argparse
import sys
from pathlib import Path

from .reader import profile_for, read_spreadsheet
from .validate import validate
from .writer import write_gedcom


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert a custom Excel family tree spreadsheet to GEDCOM format."
    )
    parser.add_argument("input", type=Path, help="Path to the source Excel file (.xlsx)")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Output GEDCOM file path (default: same name as input with .ged extension)",
    )
    args = parser.parse_args()

    if not args.input.exists():
        print(f"Error: input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    output_path: Path = args.output or args.input.with_suffix(".ged")

    individuals, families = read_spreadsheet(args.input, profile_for(args.input))
    write_gedcom(individuals, families, output_path)

    print(f"Converted {len(individuals)} individuals and {len(families)} families.")
    print(f"Output written to: {output_path}")

    # Integrity check — catches conversion regressions (dropped links, cycles,
    # impossible dates) and flags suspicious data for a human to review.
    result = validate(individuals, families)
    errors, warnings = result["errors"], result["warnings"]
    if errors or warnings:
        report_path = output_path.with_suffix(".report.txt")
        lines = [f"Validation report for {output_path.name}", ""]
        lines.append(f"ERRORS (impossible — should be 0): {len(errors)}")
        lines += [f"  - {e}" for e in errors]
        lines.append("")
        lines.append(f"WARNINGS (suspicious but possible): {len(warnings)}")
        lines += [f"  - {w}" for w in warnings]
        report_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        print(f"Validation: {len(errors)} error(s), {len(warnings)} warning(s) "
              f"— see {report_path.name}")
        if errors:
            print("  WARNING: integrity errors detected; review the report.",
                  file=sys.stderr)
    else:
        print("Validation: clean (0 errors, 0 warnings).")


if __name__ == "__main__":
    main()
