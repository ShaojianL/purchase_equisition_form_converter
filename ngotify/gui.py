"""Desktop interface for loading Word forms and exporting Excel and CSV together."""
from __future__ import annotations

import csv
from pathlib import Path

from openpyxl.utils.exceptions import IllegalCharacterError

from .csv_export import csv_text, write_csv
from .excel_export import write_excel
from .word_parser import parse_docx

def launch_gui() -> None:
    try:
        from PySide6.QtGui import QFontDatabase
        from PySide6.QtWidgets import (
            QApplication, QFileDialog, QHBoxLayout, QLabel, QMessageBox,
            QPlainTextEdit, QPushButton, QVBoxLayout, QWidget,
        )
    except ImportError as exc:
        raise SystemExit(f"Desktop interface unavailable: {exc}\nActivate .ngotify and run: python -m pip install -r requirements.txt") from exc
    app = QApplication.instance() or QApplication([])
    window = QWidget()
    window.setWindowTitle("UBC Purchase Requisition → CSV / Excel")
    window.resize(1050, 650)
    layout = QVBoxLayout(window)
    layout.setContentsMargins(20, 20, 20, 20)
    heading = QLabel("Purchase Requisition → CSV / Excel")
    font = heading.font()
    font.setPointSize(20)
    font.setBold(True)
    heading.setFont(font)
    layout.addWidget(heading)
    description = QLabel("Open Word forms, review the preview, then export both Excel and CSV files. Files stay on this computer.")
    description.setWordWrap(True)
    layout.addWidget(description)
    controls = QHBoxLayout()
    layout.addLayout(controls)
    status = QLabel("No documents loaded.")
    status.setWordWrap(True)
    rows: list[dict[str, str]] = []
    preview = QPlainTextEdit()
    preview.setReadOnly(True)
    preview.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
    preview.setFont(QFontDatabase.systemFont(QFontDatabase.SystemFont.FixedFont))

    def load() -> None:
        paths, _ = QFileDialog.getOpenFileNames(window, "Select requisition forms", "", "Word documents (*.docx)")
        if not paths:
            return
        parsed = []
        try:
            for path in paths:
                parsed.extend(parse_docx(path))
        except Exception as exc:
            QMessageBox.critical(window, "Could not load forms", f"{Path(path).name}: {exc}\n\nNo new forms were loaded.")
            return
        rows[:] = parsed
        preview.setPlainText(csv_text(rows))
        blank = sum(not row["item_number"] for row in rows)
        status.setText(f"{len(paths)} document(s) • {len(rows)} row(s) • {blank} form(s) with no items. Review values before use.")
        export.setEnabled(True)

    def save() -> None:
        dialog = QFileDialog(window, "Export Excel and CSV")
        dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptSave)
        dialog.setNameFilter("Excel workbook (*.xlsx)")
        dialog.setDefaultSuffix("xlsx")
        dialog.setOption(QFileDialog.Option.DontConfirmOverwrite, True)
        dialog.selectFile("purchase_requisitions.xlsx")
        if not dialog.exec():
            return
        excel_path = Path(dialog.selectedFiles()[0])
        if excel_path.suffix.lower() != ".xlsx":
            QMessageBox.critical(window, "Export failed", "Please use the .xlsx extension. A matching .csv file will also be created.")
            return
        csv_path = excel_path.with_suffix(".csv")
        existing = [str(path) for path in (excel_path, csv_path) if path.exists()]
        if existing:
            answer = QMessageBox.question(window, "Replace existing files?",
                                         "Replace these files?\n\n" + "\n".join(existing),
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                         QMessageBox.StandardButton.No)
            if answer != QMessageBox.StandardButton.Yes:
                return
        completed = []
        try:
            write_excel(rows, excel_path)
            completed.append(str(excel_path))
            write_csv(rows, csv_path)
            completed.append(str(csv_path))
            status.setText("Exported both files:\n" + "\n".join(completed))
        except (OSError, ValueError, csv.Error, IllegalCharacterError) as exc:
            detail = "\n\nSuccessfully wrote:\n" + "\n".join(completed) if completed else ""
            status.setText("Export incomplete. " + ("Created: " + ", ".join(completed) if completed else "No export completed."))
            QMessageBox.critical(window, "Export failed", str(exc) + detail)

    open_button = QPushButton("Open Word forms…")
    open_button.clicked.connect(load)
    controls.addWidget(open_button)
    export = QPushButton("Export Excel + CSV…")
    export.clicked.connect(save)
    export.setEnabled(False)
    controls.addWidget(export)
    controls.addStretch()
    layout.addWidget(status)
    layout.addWidget(preview, 1)
    footer = QLabel("One row per requested item. Blank fields stay empty; checked boxes are exported as labels.")
    footer.setWordWrap(True)
    layout.addWidget(footer)
    window.show()
    app.exec()
