import csv
import io
from pathlib import Path
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from openpyxl import load_workbook

from ngotify.schema import COLUMNS
from ngotify.word_parser import NS, W, parse_docx
from ngotify.csv_export import csv_text
from ngotify.excel_export import write_excel, csv_to_excel

TEMPLATE = Path(__file__).parent / "word_doc/Purchase Requisition Form - Draft.docx"


class ConverterTests(unittest.TestCase):
    def test_excel_export_preserves_text(self):
        rows = [{"funding_account": "00123", "item_description": '=HYPERLINK("bad")',
                 "additional_instructions": "First line\nSecond line", "approval": "#N/A"}]
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "export.xlsx"
            write_excel(rows, path)
            workbook = load_workbook(path)
            sheet = workbook.active
            self.assertEqual([cell.value for cell in sheet[1]], COLUMNS)
            for key, value in rows[0].items():
                cell = sheet.cell(2, COLUMNS.index(key) + 1)
                self.assertEqual(cell.value, value)
                self.assertEqual(cell.data_type, "s")
            self.assertEqual(sheet.freeze_panes, "A2")
            self.assertTrue(sheet.auto_filter.ref)
            workbook.close()

    def test_csv_to_excel(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.csv"
            destination = Path(directory) / "output.xlsx"
            with source.open("w", encoding="utf-8-sig", newline="") as stream:
                csv.writer(stream).writerows([
                    ["name", "account", "notes"],
                    ['Jane, "Lab"', "00123", "Line 1\nLine 2"],
                    ["=1+1", "", "☒ Approved"],
                ])
            csv_to_excel(source, destination)
            workbook = load_workbook(destination)
            sheet = workbook.active
            self.assertEqual(sheet["A2"].value, 'Jane, "Lab"')
            self.assertEqual(sheet["B2"].value, "00123")
            self.assertEqual(sheet["C2"].value, "Line 1\nLine 2")
            self.assertEqual(sheet["A3"].value, "=1+1")
            self.assertEqual(sheet["A3"].data_type, "s")
            self.assertEqual(sheet["C3"].value, "☒ Approved")
            workbook.close()

    def test_blank_template(self):
        rows = parse_docx(TEMPLATE)
        self.assertEqual(len(rows), 1)
        self.assertTrue(all(not value for key, value in rows[0].items() if key != "source_file"))

    def test_completed_form_and_csv_roundtrip(self):
        with zipfile.ZipFile(TEMPLATE) as archive:
            entries = {name: archive.read(name) for name in archive.namelist()}
        root = ET.fromstring(entries["word/document.xml"])
        replacements = {
            "Requester name:": 'Requester name: Jane_Doe, "Lab"',
            "Funding worktag / account:": "Funding worktag / account: ABC_123",
            "Approval:": "Approval: ☐ Approved ☒ Not Approved",
            "☐ Heavy": "☒ Heavy or oversized item",
            "Order status:": "Order status:☐ Pending ☒ Ordered ☐ Partially Received ☐ Completed ☐ On Hold",
            "Additional instructions:": "Additional instructions: Deliver to lab",
            "Notes / discrepancies:": "Notes / discrepancies: Check packaging",
        }
        for p in root.findall("w:body/w:p", NS):
            text = "".join(t.text or "" for t in p.findall(".//w:t", NS))
            for prefix, replacement in replacements.items():
                if text.startswith(prefix):
                    for child in list(p):
                        p.remove(child)
                    ET.SubElement(ET.SubElement(p, W + "r"), W + "t").text = replacement
                    break
        table = root.find("w:body/w:tbl", NS)
        for row, values in zip(table.findall("w:tr", NS)[1:], [
            ["1. Sensor, pressure", "https://example.com/sensor", "2", "$25.50"],
            ["2. =HYPERLINK(\"bad\")", "Supplier B", "1", "10"],
        ]):
            for cell, value in zip(row.findall("w:tc", NS), values):
                for child in list(cell):
                    cell.remove(child)
                ET.SubElement(ET.SubElement(ET.SubElement(cell, W + "p"), W + "r"), W + "t").text = value
        entries["word/document.xml"] = ET.tostring(root)
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "filled.docx"
            with zipfile.ZipFile(path, "w") as archive:
                for name, data in entries.items():
                    archive.writestr(name, data)
            rows = parse_docx(path)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["requester_name"], 'Jane_Doe, "Lab"')
        self.assertEqual(rows[1]["funding_account"], "ABC_123")
        self.assertEqual(rows[0]["approval"], "Not Approved")
        self.assertEqual(rows[0]["order_status"], "Ordered")
        self.assertEqual(rows[0]["special_requirements"], "Heavy or oversized item")
        self.assertEqual(rows[0]["additional_instructions"], "Deliver to lab")
        self.assertEqual(rows[0]["notes_discrepancies"], "Check packaging")
        exported = list(csv.DictReader(io.StringIO(csv_text(rows))))
        self.assertEqual(exported[0]["item_description"], "Sensor, pressure")
        self.assertTrue(exported[1]["item_description"].startswith("'="))

    def test_wrong_file_type(self):
        with self.assertRaises(ValueError):
            parse_docx("form.doc")

    def test_legacy_checkbox_and_multiline_notes(self):
        document = Document(TEMPLATE)
        for paragraph in document.paragraphs:
            if paragraph.text.startswith("Approval:"):
                paragraph.text = "Approval: "
                field = OxmlElement("w:ffData")
                checkbox = OxmlElement("w:checkBox")
                checked = OxmlElement("w:checked")
                checked.set(qn("w:val"), "1")
                checkbox.append(checked)
                field.append(checkbox)
                paragraph.add_run()._r.append(field)
                paragraph.add_run("Approved")
            elif paragraph.text.startswith("Additional instructions:"):
                paragraph.text = "Additional instructions: First line\nSecond line"
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "legacy.docx"
            document.save(path)
            row = parse_docx(path)[0]
        self.assertEqual(row["approval"], "Approved")
        self.assertEqual(row["additional_instructions"], "First line\nSecond line")

    def test_merged_item_cells_rejected(self):
        document = Document(TEMPLATE)
        document.tables[0].cell(1, 0).merge(document.tables[0].cell(1, 1))
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "merged.docx"
            document.save(path)
            with self.assertRaisesRegex(ValueError, "merged cells"):
                parse_docx(path)


if __name__ == "__main__":
    unittest.main()
