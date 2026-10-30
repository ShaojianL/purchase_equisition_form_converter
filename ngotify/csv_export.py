"""Format requisition rows and write spreadsheet-safe UTF-8 CSV files."""
from __future__ import annotations

import csv
import io
from pathlib import Path

from .schema import COLUMNS

def safe_cell(value: str) -> str:
    # Prevent spreadsheet applications from executing formulas in user-entered text.
    return "'" + value if value.lstrip().startswith(("=", "+", "-", "@")) else value


def csv_text(rows: list[dict[str, str]]) -> str:
    output = io.StringIO(newline="")
    writer = csv.DictWriter(output, fieldnames=COLUMNS)
    writer.writeheader()
    writer.writerows({key: safe_cell(row.get(key, "")) for key in COLUMNS} for row in rows)
    return output.getvalue()


def write_csv(rows: list[dict[str, str]], path: str | Path) -> None:
    Path(path).write_text(csv_text(rows), encoding="utf-8-sig", newline="")
