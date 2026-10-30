# Purchase requisition to Excel and CSV

A local Python application for the UBC Mining Engineering purchase requisition Word template in `word_doc/`. It uses `python-docx` to read completed `.docx` forms and exports Excel workbooks and UTF-8 CSV files. No files are uploaded.

## Setup

Requires Python 3.10 or newer. Run setup commands from the project folder.

### Linux / macOS (Bash or Zsh)

Create and activate a virtual environment, then install the dependencies:

```sh
python3 -m venv .ngotify
source .ngotify/bin/activate
python -m pip install -r requirements.txt
```

### Windows (PowerShell)

Create a Windows virtual environment and install the dependencies using its Python directly:

```powershell
cd C:\path\to\ngotification
python -m venv .ngotify
.\.ngotify\Scripts\python.exe -m pip install -r requirements.txt
.\.ngotify\Scripts\python.exe app.py
```

Replace the example project path with your local folder. For later launches, run `.\.ngotify\Scripts\python.exe app.py` from that folder. Activation is not required. Create a separate environment on each operating system; a Linux `.ngotify` environment cannot be reused on Windows.

### Linux (PowerShell)

To launch using the existing Linux environment without activating it:

```powershell
cd /home/liangs/ngotification
./.ngotify/bin/python app.py
```

## Desktop application

The desktop interface uses PySide6, installed inside `.ngotify` through `requirements.txt`. It does not require Tkinter or a separate Qt installation. A graphical desktop session is required. With `.ngotify` activated, run from this folder:

```sh
python3 app.py
```

Choose **Open Word forms…**, select one or more completed forms, review the CSV preview, and choose **Export Excel + CSV…**. Choose an Excel filename such as `purchase_requisitions.xlsx`; the app writes that workbook and `purchase_requisitions.csv` in the same folder, directly from the loaded Word data. Multiple forms are combined in both files. The app asks before replacing either existing file and shows both output paths when done. Selecting new forms replaces the current selection. Correct values in Word and reopen the forms if needed. Existing CSV conversion remains available through the command line.

## Command line

```sh
python3 app.py "word_doc/Purchase Requisition Form - Draft.docx" -o requisitions.csv
python3 app.py first.docx second.docx -o combined.csv
python3 app.py first.docx -o requisitions.xlsx
python3 app.py --from-csv requisitions.csv -o requisitions.xlsx
```

The output path is overwritten if it already exists. All inputs must parse successfully before output is written.

Excel export uses `openpyxl`, installed within `.ngotify` by `requirements.txt`. Workbooks include a styled header, filters, a frozen header row and wrapped text. Values remain text to preserve account numbers, dates, and entered prices; Excel does not evaluate formula-like entries. Existing UTF-8 comma-separated CSV files (with or without a BOM) can be converted with their column order intact. Any apostrophes already added to CSV values are preserved. Numeric calculations require converting the relevant columns to numbers in Excel.

## Export format

Each requested item becomes one CSV row. Requester, funding, delivery, approval, and technical-team fields repeat on each row. `source_file` identifies the input. An empty item table produces one row with blank item fields so request details are retained. Empty placeholders stay empty; no approval, quantity, price, date, or status is inferred. Dates and prices retain their entered format.

Selected Unicode checkboxes (`☒`, `☑`), text checkboxes (`[x]`), and legacy Word checkbox fields are supported. Multiple checked labels are joined with semicolons. To fill a checkbox in this template, replace `☐` with `☒`. Supplier links must appear as visible text; hidden hyperlink destinations are not extracted. Multiline notes are preserved in quoted CSV cells. Potential spreadsheet formulas are prefixed with an apostrophe for safe spreadsheet viewing.

This converter expects the supplied template's labels and four-column item table. It supports typed document content, including more item rows added to that table. It does not read scanned images, handwriting, signatures embedded as images, legacy `.doc` files, floating text boxes, or attached approval emails. Exported approval text is a transcription, not verification. Review exports before using them in a purchasing workflow.

## Code organization

- `app.py`: launcher for the existing command.
- `ngotify/schema.py`: form labels and output columns.
- `ngotify/word_parser.py`: Word field, checkbox and item extraction.
- `ngotify/csv_export.py`: CSV formatting and writing.
- `ngotify/excel_export.py`: Excel writing and CSV-to-Excel conversion.
- `ngotify/gui.py`: desktop interface and paired export flow.
- `ngotify/cli.py`: command-line arguments and conversion routing.
- `ngotify/__main__.py`: alternate launcher, `python -m ngotify`.

## Verification

```sh
python3 -m unittest -v
```

Tests cover the actual blank template, a generated completed form with multiple items and checked options, CSV quoting, and spreadsheet formula protection.
