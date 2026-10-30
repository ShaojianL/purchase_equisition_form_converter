"""Command-line entry point for the requisition converter."""
from __future__ import annotations

import argparse
import csv
from pathlib import Path
import zipfile

from docx.opc.exceptions import PackageNotFoundError
from lxml.etree import XMLSyntaxError
from openpyxl.utils.exceptions import IllegalCharacterError

from .csv_export import write_csv
from .excel_export import csv_to_excel, write_excel
from .word_parser import parse_docx

def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("documents", nargs="*", type=Path, help="Completed .docx requisitions")
    parser.add_argument("-o", "--output", type=Path, help="Destination .csv or .xlsx (required with inputs)")
    parser.add_argument("--from-csv", type=Path, help="Convert an existing UTF-8 CSV to .xlsx")
    args = parser.parse_args()
    if args.from_csv and args.documents:
        parser.error("Use either Word documents or --from-csv")
    if not args.documents and not args.from_csv:
        if args.output:
            parser.error("Provide at least one document with --output")
        from .gui import launch_gui

        launch_gui()
        return
    if not args.output:
        parser.error("--output is required when converting files")
    if args.output.suffix.lower() not in (".csv", ".xlsx"):
        parser.error("Output must end in .csv or .xlsx")
    try:
        if args.from_csv:
            csv_to_excel(args.from_csv, args.output)
        else:
            rows = [row for path in args.documents for row in parse_docx(path)]
            writer = write_excel if args.output.suffix.lower() == ".xlsx" else write_csv
            writer(rows, args.output)
    except (OSError, ValueError, KeyError, csv.Error, IllegalCharacterError,
            zipfile.BadZipFile, PackageNotFoundError, XMLSyntaxError) as exc:
        parser.exit(1, f"Conversion failed: {exc}\n")
    print(f"Exported to {args.output}")
