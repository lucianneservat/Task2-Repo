"""
Automated routine — Task 2
Reads Input.xlsx from SharePoint via Microsoft Graph API, validates read fidelity,
then processes contact and project data.

Prerequisites:
    pip install msal httpx openpyxl pandas

Required environment variables (set in .env or shell):
    AZURE_TENANT_ID
    AZURE_CLIENT_ID
    AZURE_CLIENT_SECRET
"""

import os
import tempfile

import httpx
import msal
import openpyxl
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

TENANT_ID     = os.environ["AZURE_TENANT_ID"]
CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]

DRIVE_ID = "b!KR22wzNh7EGhNkG9pT8SDmQj3_cA0QpLmYQ10EYyrppYZ0cfEu0SSZYVTa0P-cwb"
ITEM_ID  = "01GHYZT3XNATTYIS7LPFEIL4WIKYLQOSRU"
GRAPH    = "https://graph.microsoft.com/v1.0"


def get_token() -> str:
    app = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    )
    result = app.acquire_token_for_client(["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description')}")
    return result["access_token"]


def download_workbook(token: str) -> str:
    """Download Input.xlsx to a temp file and return its path."""
    headers = {"Authorization": f"Bearer {token}"}
    r = httpx.get(
        f"{GRAPH}/drives/{DRIVE_ID}/items/{ITEM_ID}/content",
        headers=headers,
        follow_redirects=True,
        timeout=60,
    )
    r.raise_for_status()
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.write(r.content)
    tmp.close()
    return tmp.name


# ---------------------------------------------------------------------------
# CHECKPOINT 1 — Read fidelity
# Assert we can access every sheet, every header, and exact row counts.
# ---------------------------------------------------------------------------

def validate_fidelity(path: str) -> dict[str, pd.DataFrame]:
    """Return a dict of {sheet_name: DataFrame} after asserting structural integrity."""
    wb = openpyxl.load_workbook(path, read_only=True)
    sheets: dict[str, pd.DataFrame] = {}

    print(f"\n=== Checkpoint 1: Read fidelity ===")
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
    token = get_token()
    print("Token acquired.")

    tmp_path = download_workbook(token)
    print(f"Workbook downloaded to: {tmp_path}")

    sheets = validate_fidelity(tmp_path)
    process(sheets)


if __name__ == "__main__":
    main()
