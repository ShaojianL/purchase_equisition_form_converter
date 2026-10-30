"""Write formatted Excel workbooks from requisition rows or existing CSV files."""
from __future__ import annotations

import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .schema import COLUMNS

def write_workbook(table: list[list[str]], path: str | Path) -> None:
    if not table or not table[0]:
        raise ValueError("The CSV has no header row.")
    if len(table) > 1_048_576 or max(map(len, table)) > 16_384:
        raise ValueError("The data exceeds Excel's worksheet size limits.")
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Requisitions"
    for row_number, values in enumerate(table, 1):
        for column, value in enumerate(values, 1):
            if len(value) > 32_767:
                raise ValueError(f"Cell in row {row_number}, column {column} exceeds Excel's text limit.")
            cell = sheet.cell(row_number, column, value)
            # Explicit strings preserve account IDs, dates and formula-like text.
            cell.data_type = "s"
            cell.number_format = "@"
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="002145")
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for column in range(1, sheet.max_column + 1):
        sheet.column_dimensions[get_column_letter(column)].width = 28
    workbook.save(path)


def write_excel(rows: list[dict[str, str]], path: str | Path) -> None:
    write_workbook([COLUMNS, *[[row.get(key, "") for key in COLUMNS] for row in rows]], path)


def csv_to_excel(source: str | Path, destination: str | Path) -> None:
    if Path(destination).suffix.lower() != ".xlsx":
        raise ValueError("Excel output must use the .xlsx extension.")
    with Path(source).open(encoding="utf-8-sig", newline="") as stream:
        table = list(csv.reader(stream, strict=True))
    write_workbook(table, destination)
