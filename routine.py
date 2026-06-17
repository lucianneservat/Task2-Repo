"""
Automated routine — Task 2
Reads Input.xlsx from a local path, validates read fidelity, then processes data.

Prerequisites:
    pip install -r requirements.txt

Usage:
    Place Input.xlsx in the same directory as this script, then run:
        python routine.py
    Or pass a custom path:
        python routine.py /path/to/Input.xlsx
"""

import sys
from pathlib import Path

import openpyxl
import pandas as pd


INPUT_FILE = Path(__file__).parent / "Input.xlsx"


# ---------------------------------------------------------------------------
# CHECKPOINT 1 — Read fidelity
# Assert every sheet, every header, and exact row counts before processing.
# ---------------------------------------------------------------------------

def validate_fidelity(path: Path) -> dict[str, pd.DataFrame]:
    if not path.exists():
        raise FileNotFoundError(
            f"Input file not found: {path}\n"
            "Download Input.xlsx from SharePoint and place it next to routine.py"
        )

    wb = openpyxl.load_workbook(path, read_only=True)
    sheets: dict[str, pd.DataFrame] = {}

    print(f"\n=== Checkpoint 1: Read fidelity ===")
    print(f"File: {path.name}  ({path.stat().st_size:,} bytes)")
    print(f"Worksheets found ({len(wb.sheetnames)}): {wb.sheetnames}\n")

    for name in wb.sheetnames:
        ws = wb[name]
        rows = list(ws.iter_rows(values_only=True))
        assert rows, f"Sheet '{name}' is empty"

        headers = list(rows[0])
        assert all(h is not None for h in headers), (
            f"Blank header cell in sheet '{name}': {headers}"
        )

        data_rows = rows[1:]
        df = pd.DataFrame(data_rows, columns=headers)
        sheets[name] = df

        print(f"  ✅ '{name}'")
        print(f"     Headers ({len(headers)}): {headers}")
        print(f"     Data rows: {len(df)}\n")

    wb.close()
    return sheets


# ---------------------------------------------------------------------------
# CHECKPOINT 2 — Processing
# Add your transformation logic below.
# ---------------------------------------------------------------------------

def process(sheets: dict[str, pd.DataFrame]) -> None:
    print("=== Checkpoint 2: Processing ===")

    for sheet_name, df in sheets.items():
        print(f"\nSheet: '{sheet_name}'")
        print(df.head())

        # TODO: add your per-sheet processing logic here


# ---------------------------------------------------------------------------

def main() -> None:
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else INPUT_FILE
    sheets = validate_fidelity(path)
    process(sheets)


if __name__ == "__main__":
    main()
