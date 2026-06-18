"""
Automated routine — Task 2
1. Reads template files to observe their column structure.
2. Reads every sheet in Input.xlsx (read fidelity checkpoint).
3. For each sheet, normalizes phone numbers, de-duplicates, maps columns,
   and builds two output Excel files. Ambiguous/invalid phones are quarantined
   to separate .review files.

Prerequisites:
    pip install openpyxl pandas

Usage:
    python routine.py
"""

import re
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl import Workbook

REPO               = Path(__file__).parent
INPUT_FILE         = REPO / "Input.xlsx"
CAMPAIGN_TEMPLATE  = REPO / "campaign_Otros_Proyectos.xlsx"
CUSTOMERS_TEMPLATE = REPO / "customers_Otros_Proyectos.xlsx"
OUTPUT_DIR         = REPO / "output"


# ---------------------------------------------------------------------------
# CHECKPOINT 0 — Read template structures
# ---------------------------------------------------------------------------

def read_template_headers(path: Path) -> list[str]:
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.worksheets[0]
    headers = [cell.value for cell in next(ws.iter_rows(max_row=1))]
    wb.close()
    print(f"  Template '{path.name}': {headers}")
    return headers


# ---------------------------------------------------------------------------
# CHECKPOINT 1 — Read fidelity on Input.xlsx
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

        df = pd.DataFrame(rows[1:], columns=headers)
        sheets[name] = df

        print(f"  ✅ '{name}'")
        print(f"     Headers ({len(headers)}): {headers}")
        print(f"     Data rows: {len(df)}\n")

    wb.close()
    return sheets


# ---------------------------------------------------------------------------
# Phone normalization
# ---------------------------------------------------------------------------

def normalize_phone(raw) -> tuple[str | None, bool]:
    """
    Returns (normalized_digits, valid).
    Rules:
    - Strip all non-digit characters.
    - 0057 prefix → strip the leading 00.
    - 12 digits starting with 573 → valid, return as-is.
    - 10 digits starting with 3   → prepend 57, return.
    - Everything else             → invalid, quarantine to .review.
    """
    if pd.isna(raw) or str(raw).strip() in ("", "None"):
        return None, False

    digits = re.sub(r"\D", "", str(raw))

    if digits.startswith("0057"):
        digits = digits[2:]

    if len(digits) == 12 and digits.startswith("573"):
        return digits, True

    if len(digits) == 10 and digits.startswith("3"):
        return "57" + digits, True

    return None, False


# ---------------------------------------------------------------------------
# CHECKPOINT 2 — Map columns, normalize phones, build output files
# ---------------------------------------------------------------------------

def slugify(name: str) -> str:
    return name.replace(" ", "_")


def build_campaign(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    records, review = [], []
    seen: set[str] = set()

    for _, row in df.iterrows():
        phone, valid = normalize_phone(row["Números"])

        if not valid:
            review.append({"review": "review"})
            continue

        if phone in seen:
            continue

        seen.add(phone)
        records.append({
            "number":          phone,
            "nombre_cliente":  f"{row['Nombre']} {row['Apellidos']}",
            "hubspot_deal_id": str(int(row["Negocio ID"])) if pd.notna(row["Negocio ID"]) else "",
        })

    return pd.DataFrame(records), pd.DataFrame(review)


def build_customers(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    records, review = [], []
    seen: set[str] = set()

    for _, row in df.iterrows():
        phone, valid = normalize_phone(row["Números"])

        if not valid:
            review.append({"review": "review"})
            continue

        if phone in seen:
            continue

        seen.add(phone)
        records.append({
            "phone":                 phone,
            "firstname":             row["Nombre"],
            "lastname":              row["Apellidos"],
            "email":                 None,
            "voice_model_selection": row["Nombre del proyecto"],
        })

    return pd.DataFrame(records), pd.DataFrame(review)


def save_excel(df: pd.DataFrame, path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(list(df.columns))
    for row in df.itertuples(index=False, name=None):
        ws.append(list(row))
    wb.save(path)
    print(f"  Created: {path.name}  ({len(df)} rows)")


def create_outputs(sheets: dict[str, pd.DataFrame]) -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    print(f"=== Checkpoint 2: Building output files in '{OUTPUT_DIR}' ===\n")

    for sheet_name, df in sheets.items():
        slug = slugify(sheet_name)

        campaign_df,  campaign_review  = build_campaign(df)
        customers_df, customers_review = build_customers(df)

        save_excel(campaign_df,  OUTPUT_DIR / f"campaign_{slug}.xlsx")
        save_excel(customers_df, OUTPUT_DIR / f"customers_{slug}.xlsx")

        if not campaign_review.empty:
            save_excel(campaign_review,  OUTPUT_DIR / f"campaign_{slug}.review.xlsx")
        if not customers_review.empty:
            save_excel(customers_review, OUTPUT_DIR / f"customers_{slug}.review.xlsx")

    print(f"\nDone. Upload the contents of output/ to SharePoint manually.")


# ---------------------------------------------------------------------------

def main() -> None:
    print("=== Checkpoint 0: Reading template structures ===")
    read_template_headers(CAMPAIGN_TEMPLATE)
    read_template_headers(CUSTOMERS_TEMPLATE)

    sheets = validate_fidelity(INPUT_FILE)
    create_outputs(sheets)


if __name__ == "__main__":
    main()
