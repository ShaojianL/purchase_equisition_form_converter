"""Extract requisition fields and requested items from Word documents."""
from __future__ import annotations

from pathlib import Path
import re

from docx import Document
from docx.text.paragraph import Paragraph

from .schema import COLUMNS, FIELDS, ITEM_FIELDS

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{" + NS["w"] + "}"

def clean(value: str) -> str:
    value = re.sub(r"_{3,}|…+", "", value).strip()
    return "" if not value.strip("$ .\t") else value


def paragraph_text(paragraph: Paragraph) -> str:
    # Normal paragraphs, including hyperlink display text, use the public API.
    if not paragraph._p.xpath(".//w:checkBox"):
        return paragraph.text.strip()
    # Legacy form checkboxes have no high-level API; preserve their positions.
    parts = []
    for node in paragraph._p.iter():
        if node.tag == W + "t":
            parts.append(node.text or "")
        elif node.tag in (W + "br", W + "tab"):
            parts.append(" ")
        elif node.tag == W + "checkBox":  # Legacy Word checkbox form fields.
            checked = node.find(W + "checked")
            enabled = checked is not None and checked.get(W + "val", "1") not in ("0", "false", "off")
            parts.append("☒ " if enabled else "☐ ")
    return "".join(parts).strip()


def selected(text: str) -> list[str]:
    return [label.strip() for marker, label in
            re.findall(r"(☐|☑|☒|\[[ xX✓]\])\s*([^☐☑☒\[]*)", text)
            if marker in ("☑", "☒", "[x]", "[X]", "[✓]") and label.strip()]


def parse_docx(path: str | Path) -> list[dict[str, str]]:
    path = Path(path)
    if path.suffix.lower() != ".docx":
        raise ValueError("Please select a .docx file; legacy .doc and scanned files are not supported.")
    document = Document(path)
    paragraphs = [paragraph_text(p) for p in document.paragraphs]
    if not any("Purchase Requisition Form" in p for p in paragraphs):
        raise ValueError("This document does not match the purchase requisition template.")
    record = dict.fromkeys(COLUMNS, "")
    record["source_file"] = path.name
    block = None
    special = []
    for text in paragraphs:
        if re.match(r"^\d+\.\s", text) or text.startswith(("Interim Procedure:", "An approval email")):
            block = None
            continue
        label, separator, value = text.partition(":")
        label = re.sub(r"\s+", " ", label).strip()
        if separator and label in FIELDS:
            record[FIELDS[label]] = clean(value)
            block = None
        elif label == "Special requirements (check all that apply)":
            block = "special"
        elif label in ("Approval", "Order status") and separator:
            block = "approval" if label == "Approval" else "order_status"
            record[block] = "; ".join(selected(value))
        elif label in ("Additional instructions", "Notes / discrepancies") and separator:
            block = "additional_instructions" if label == "Additional instructions" else "notes_discrepancies"
            record[block] = clean(value)
        elif block == "special":
            special.extend(selected(text))
        elif block in ("approval", "order_status"):
            values = selected(text)
            if values:
                record[block] = "; ".join(filter(None, [record[block], *values]))
        elif block and clean(text):
            record[block] = "\n".join(filter(None, [record[block], clean(text)]))
    record["special_requirements"] = "; ".join(special)
    items = []
    found_table = False
    for table in document.tables:
        rows = table.rows
        if not rows:
            continue
        headers = [cell.text for cell in rows[0].cells]
        if len(headers) != 4 or "description" not in headers[0].lower() or "qty" not in headers[2].lower():
            continue
        found_table = True
        for index, row in enumerate(rows[1:], 1):
            cells = [clean("\n".join(paragraph_text(p) for p in cell.paragraphs))
                     for cell in row.cells]
            if len(cells) != 4 or any(cell.grid_span != 1 for cell in row.cells):
                raise ValueError(f"Item row {index} must have four columns. Check merged cells.")
            match = re.match(r"^\s*(\d+)\.\s*(.*)$", cells[0], re.S)
            number, description = (match[1], clean(match[2])) if match else (str(index), cells[0])
            if not any([description, *cells[1:]]):
                continue
            items.append(dict(zip(ITEM_FIELDS, [number, description, *cells[1:]])))
    if not found_table:
        raise ValueError("Could not find the four-column Items Requested table.")
    return [{**record, **item} for item in items] if items else [record]
