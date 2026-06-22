"""Command-line entry point for the family tree converter."""

import argparse
import sys
from pathlib import Path

from .reader import read_spreadsheet
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

    individuals, families = read_spreadsheet(args.input)
    write_gedcom(individuals, families, output_path)

    print(f"Converted {len(individuals)} individuals and {len(families)} families.")
    print(f"Output written to: {output_path}")


if __name__ == "__main__":
    main()
