import csv
import json
import logging
import os
import smtplib
import sys
import traceback
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from gspread.exceptions import WorksheetNotFound

# ============================================================================
# CONFIG — edit these values for your setup
#
# Everything below has an environment-variable override. When the matching
# env var is NOT set, the original hardcoded local (Windows) values are used
# unchanged, so this script still runs exactly as before on your machine.
# In GitHub Actions, the workflow sets these env vars (from Secrets) instead.
# ============================================================================

# ---- Google service account / source spreadsheet ----
SERVICE_ACCOUNT_FILE = os.environ.get(
    "SERVICE_ACCOUNT_FILE", r"D:\Fabric Inventory\Google service account\service_account.json"
)
# In GitHub Actions we pass the key's JSON content directly instead of a file path.
SERVICE_ACCOUNT_JSON = os.environ.get("SERVICE_ACCOUNT_JSON", "")

SOURCE_SPREADSHEET_ID_OR_URL = os.environ.get(
    "SOURCE_SPREADSHEET_ID_OR_URL", "1rLcgwVua5XXyY7lZx6G16xqmNvbwyp_BWS_y61TBssA"
) or "1rLcgwVua5XXyY7lZx6G16xqmNvbwyp_BWS_y61TBssA"
# ^ the `or` fallback matters: if the env var/secret exists but is set to an
# empty string (common when a GitHub secret is missing/blank), os.environ.get
# would otherwise silently return "" instead of falling back to the default.

# ---- Local working folders / files (mirrors the original notebook layout) ----
# Set PIPELINE_BASE_DIR to relocate all working folders under one root
# (the GitHub Action sets this to "output" since Windows drive paths
# like D:\ don't exist on the Linux runner).
_BASE_DIR = os.environ.get("PIPELINE_BASE_DIR", "")

if _BASE_DIR:
    _base = Path(_BASE_DIR)
    SKU_MAPPING_DIR = _base / "SKU x FAB mapping"
    CONSUMPTION_DIR = _base / "SKU x FAB Consumption"
    CURRENT_INV_DIR = _base / "Fabric current Inventory"
    CURRENT_FEAS_DIR = _base / "Partial Possible Current"
    OPEN_PO_DIR = _base / "Open PO Inventory"
    OPEN_PO_FEAS_DIR = _base / "Partial Possible Open PO"
    WIP_DIR = _base / "WIP"
    WIP_FEAS_DIR = _base / "Partial Possible WIP"
else:
    SKU_MAPPING_DIR = Path(r"D:\Fabric Inventory\SKU\SKU x FAB mapping")
    CONSUMPTION_DIR = Path(r"D:\Fabric Inventory\SKU\SKU x FAB Consumption")
    CURRENT_INV_DIR = Path(r"D:\Fabric Inventory\Current Possible\Fabric current Inventory")
    CURRENT_FEAS_DIR = Path(r"D:\Fabric Inventory\Current Possible\Partial Possible Current")
    OPEN_PO_DIR = Path(r"D:\Fabric Inventory\Open PO\Open PO Inventory")
    OPEN_PO_FEAS_DIR = Path(r"D:\Fabric Inventory\Open PO\Partial Possible Open PO")
    WIP_DIR = Path(r"D:\Fabric Inventory\WIP\WIP")
    WIP_FEAS_DIR = Path(r"D:\Fabric Inventory\WIP\Partial Possible WIP")

SKU_MAPPING_FILE = SKU_MAPPING_DIR / "SKU x FAB mapping.csv"
CONSUMPTION_FILE = CONSUMPTION_DIR / "SKU x Fab consumption.csv"
CURRENT_INV_FILE = CURRENT_INV_DIR / "Fabric Current Inventory.csv"
CURRENT_FEAS_FILE = CURRENT_FEAS_DIR / "partial_possible_Current.csv"
OPEN_PO_FILE = OPEN_PO_DIR / "Open PO.csv"
OPEN_PO_FEAS_FILE = OPEN_PO_FEAS_DIR / "Partial_possible_OPEN_PO.csv"
WIP_FILE = WIP_DIR / "WIP.csv"
WIP_FEAS_FILE = WIP_FEAS_DIR / "Partial_possible_WIP.csv"

# ---- Optional: push feasibility results back into a Google Sheet ----
ENABLE_SYNC_TO_SHEETS = os.environ.get("ENABLE_SYNC_TO_SHEETS", "true").strip().lower() not in ("false", "0", "no")
SYNC_SPREADSHEET_ID = os.environ.get("SYNC_SPREADSHEET_ID", "1oKefmjFRJI0N_CniKjqLNzJ0C3qqbS0XBT1BP0nA31g")
SYNC_JOBS = [
    (CURRENT_FEAS_FILE, "Partial Possible Current"),
    (OPEN_PO_FEAS_FILE, "Open PO"),
    (WIP_FEAS_FILE, "WIP"),
]

# ---- Email notification (Gmail SMTP) ----
GMAIL_CREDENTIALS_CSV = os.environ.get("GMAIL_CREDENTIALS_CSV", r"D:\Credentials\Gmail Credentials.csv")
# If these two env vars are set (as they are in GitHub Actions, from Secrets),
# they're used directly and GMAIL_CREDENTIALS_CSV is skipped entirely.
GMAIL_SENDER_EMAIL = os.environ.get("GMAIL_SENDER_EMAIL", "").strip().replace("\xa0", "")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "").strip().replace("\xa0", "")
EMAIL_TO = (os.environ.get("EMAIL_TO", "") or "").strip().replace("\xa0", "") or None  # None -> sends the report to the same Gmail address used to send it
EMAIL_SUBJECT_SUCCESS = "Fabric KPI Pipeline — Success"
EMAIL_SUBJECT_FAILURE = "Fabric KPI Pipeline — FAILED"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465

# ============================================================================
# Logging
# ============================================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-7s | %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("fabric_kpi")

SCOPES_READONLY = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]
SCOPES_READWRITE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ============================================================================
# Google Sheets helpers
# ============================================================================
def get_client(scopes):
    if SERVICE_ACCOUNT_JSON:
        try:
            info = json.loads(SERVICE_ACCOUNT_JSON)
        except json.JSONDecodeError as e:
            raise RuntimeError(
                "SERVICE_ACCOUNT_JSON is set but is not valid JSON. If this comes from a "
                "GitHub secret, make sure the ENTIRE key file contents were pasted in, "
                "with no extra quoting or truncation."
            ) from e
        log.info(f"    Using service account (from JSON env var): {info.get('client_email', '<unknown>')}")
        creds = Credentials.from_service_account_info(info, scopes=scopes)
        return gspread.authorize(creds)

    creds_path = Path(SERVICE_ACCOUNT_FILE)
    if not creds_path.exists():
        raise FileNotFoundError(f"Service account file not found at: {SERVICE_ACCOUNT_FILE}")
    with open(creds_path) as f:
        info = json.load(f)
    log.info(f"    Using service account (from file {SERVICE_ACCOUNT_FILE}): {info.get('client_email', '<unknown>')}")
    creds = Credentials.from_service_account_file(str(creds_path), scopes=scopes)
    return gspread.authorize(creds)


def open_spreadsheet(client, id_or_url):
    # Defensive cleanup: env vars / GitHub secrets can easily pick up stray
    # quotes, surrounding whitespace, or a trailing newline, all of which
    # will make Google return its generic "unable to open file" HTML page
    # instead of a clean auth/not-found error.
    cleaned = id_or_url.strip().strip('"').strip("'")
    if not cleaned:
        raise RuntimeError(
            "SOURCE_SPREADSHEET_ID_OR_URL is empty. Check that the GitHub secret/variable "
            "of that name either isn't referenced in the workflow, or is set to a real "
            "spreadsheet ID/URL (not blank)."
        )
    if cleaned != id_or_url:
        log.warning(
            "SOURCE_SPREADSHEET_ID_OR_URL had leading/trailing whitespace or quotes; "
            "using cleaned value."
        )

    try:
        if cleaned.startswith("http"):
            return client.open_by_url(cleaned)
        return client.open_by_key(cleaned)
    except gspread.exceptions.APIError as e:
        raise RuntimeError(
            "Could not open the Google Sheet with "
            f"SOURCE_SPREADSHEET_ID_OR_URL={cleaned!r}. This usually means one of:\n"
            "  1) The ID/URL is wrong (check for a copy-paste error or stray characters)\n"
            "  2) The sheet has NOT been shared with the service account's client_email\n"
            "  3) The Google Sheets API and/or Drive API is not enabled on that GCP project\n"
            "  4) The service account key is invalid or revoked\n"
            f"Original gspread error: {e}"
        ) from e


def col_letter_to_index(letter):
    letter = letter.strip().upper()
    index = 0
    for char in letter:
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index - 1


def extract_columns(worksheet, columns, header_row):
    all_values = worksheet.get_all_values()
    if len(all_values) < header_row:
        raise ValueError(f"Worksheet only has {len(all_values)} rows; header_row={header_row} is out of range.")
    headers_row_values = all_values[header_row - 1]
    data_rows = all_values[header_row:]
    col_indices = [col_letter_to_index(c) for c in columns]
    headers = []
    for idx, letter in zip(col_indices, columns):
        header_value = headers_row_values[idx].strip() if idx < len(headers_row_values) else ""
        headers.append(header_value if header_value else letter)
    records = []
    for row in data_rows:
        record = [row[idx] if idx < len(row) else "" for idx in col_indices]
        if any(cell.strip() for cell in record):
            records.append(record)
    return pd.DataFrame(records, columns=headers)


def filter_by_status(df, status_column_name, statuses_to_keep, keep_blank):
    if status_column_name not in df.columns:
        raise ValueError(f"Column '{status_column_name}' not found in: {list(df.columns)}")
    status_series = df[status_column_name].astype(str).str.strip()
    statuses_norm = [s.strip() for s in statuses_to_keep]
    if keep_blank:
        is_blank = status_series == ""
        is_kept = status_series.isin(statuses_norm)
        return df[is_blank | is_kept].reset_index(drop=True)
    status_lower = status_series.str.lower()
    keep_lower = [s.lower() for s in statuses_norm]
    return df[status_lower.isin(keep_lower)].reset_index(drop=True)


def sum_duplicates(df, group_by_column, sum_column):
    if group_by_column not in df.columns:
        raise ValueError(f"Column '{group_by_column}' not found in: {list(df.columns)}")
    if sum_column not in df.columns:
        raise ValueError(f"Column '{sum_column}' not found in: {list(df.columns)}")
    work_df = df[[group_by_column, sum_column]].copy()
    work_df[sum_column] = work_df[sum_column].astype(str).str.replace(",", "", regex=False).str.strip()
    work_df[sum_column] = pd.to_numeric(work_df[sum_column], errors="coerce").fillna(0).abs()
    return work_df.groupby(group_by_column, as_index=False, sort=False)[sum_column].sum()


# ============================================================================
# Stage 1-4: Google Sheet extracts
# ============================================================================
def stage1_extract_mapping():
    client = get_client(SCOPES_READONLY)
    spreadsheet = open_spreadsheet(client, SOURCE_SPREADSHEET_ID_OR_URL)
    worksheet = spreadsheet.worksheet("SKU x FAB mapping")
    columns = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"]
    df = extract_columns(worksheet, columns, header_row=1)
    SKU_MAPPING_DIR.mkdir(parents=True, exist_ok=True)
    df.to_csv(SKU_MAPPING_FILE, index=False, encoding="utf-8-sig")
    log.info(f"    Extracted {len(df)} rows -> {SKU_MAPPING_FILE}")


def stage2_extract_current_inventory():
    client = get_client(SCOPES_READONLY)
    spreadsheet = open_spreadsheet(client, SOURCE_SPREADSHEET_ID_OR_URL)
    worksheet = spreadsheet.worksheet("Fabric Current Inventory")
    df = extract_columns(worksheet, ["A", "B"], header_row=1)
    CURRENT_INV_DIR.mkdir(parents=True, exist_ok=True)
    df.to_csv(CURRENT_INV_FILE, index=False, encoding="utf-8-sig")
    log.info(f"    Extracted {len(df)} rows -> {CURRENT_INV_FILE}")


def stage3_extract_open_po():
    client = get_client(SCOPES_READONLY)
    spreadsheet = open_spreadsheet(client, SOURCE_SPREADSHEET_ID_OR_URL)
    worksheet = spreadsheet.worksheet("Open PO")
    df = extract_columns(worksheet, ["A", "B", "C"], header_row=1)
    df = filter_by_status(df, "SL", ["Part-Close"], keep_blank=True)
    df = sum_duplicates(df, "FAB", "Balance Qty")
    OPEN_PO_DIR.mkdir(parents=True, exist_ok=True)
    df.to_csv(OPEN_PO_FILE, index=False, encoding="utf-8-sig")
    log.info(f"    Extracted {len(df)} rows -> {OPEN_PO_FILE}")


def stage4_extract_wip():
    client = get_client(SCOPES_READONLY)
    spreadsheet = open_spreadsheet(client, SOURCE_SPREADSHEET_ID_OR_URL)
    worksheet = spreadsheet.worksheet("WIP")
    df = extract_columns(worksheet, ["A", "B", "C"], header_row=1)
    df = filter_by_status(df, "Status", ["Open"], keep_blank=False)
    df = sum_duplicates(df, "FAB ID", "Pending")
    WIP_DIR.mkdir(parents=True, exist_ok=True)
    df.to_csv(WIP_FILE, index=False, encoding="utf-8-sig")
    log.info(f"    Extracted {len(df)} rows -> {WIP_FILE}")


# ============================================================================
# Stage 5: Unpivot SKU x FAB mapping into SKU x Fab Consumption
# ============================================================================
def stage5_build_consumption():
    df = pd.read_csv(SKU_MAPPING_FILE)
    fab_pairs = [
        ("FAB ID", "Consp_FabID_1"),
        ("FabID_2", "Consp_FabID_2"),
        ("FabID_3", "Consp_FabID_3"),
        ("FabID_4", "Consp_FabID_4"),
        ("FabID_5", "FabID_5 Consump"),
    ]
    chunks = []
    for fab_col, consp_col in fab_pairs:
        chunk = df[["SKU CODE", fab_col, consp_col]].copy()
        chunk.columns = ["SKU", "Fabric", "Consumption"]
        chunk = chunk.dropna(subset=["SKU", "Fabric"])
        chunks.append(chunk)

    result = pd.concat(chunks, ignore_index=True)
    result["Consumption"] = pd.to_numeric(result["Consumption"], errors="coerce")
    result = result.dropna(subset=["SKU", "Fabric"])
    result = result.sort_values(["SKU", "Fabric"]).reset_index(drop=True)

    CONSUMPTION_DIR.mkdir(parents=True, exist_ok=True)
    result.to_csv(CONSUMPTION_FILE, index=False)
    log.info(f"    {len(result)} rows written -> {CONSUMPTION_FILE}")


# ============================================================================
# Stage 6-8: Feasibility calculations
# ============================================================================
def build_feasibility(inventory_file, output_file):
    inv = pd.read_csv(inventory_file)
    sku = pd.read_csv(CONSUMPTION_FILE)

    inv.columns = ["Fabric", "Current Inventory"]
    merged = sku.merge(inv, on="Fabric", how="left")
    merged["Current Inventory"] = merged["Current Inventory"].fillna(0)
    merged["Consumption"] = merged["Consumption"].fillna(0)

    merged["Units Made out of it"] = merged.apply(
        lambda r: int(r["Current Inventory"] / r["Consumption"]) if r["Consumption"] > 0 else 0,
        axis=1,
    )
    merged["Min unit"] = merged.groupby("SKU")["Units Made out of it"].transform("min")

    result = merged[["SKU", "Fabric", "Consumption", "Current Inventory", "Units Made out of it", "Min unit"]]

    Path(output_file).parent.mkdir(parents=True, exist_ok=True)
    result.to_csv(output_file, index=False)

    total_skus = result["SKU"].nunique()
    producible = result[result["Min unit"] > 0]["SKU"].nunique()
    log.info(f"    {len(result)} rows | {total_skus} SKUs | {producible} producible -> {output_file}")


def stage6_feasibility_current():
    build_feasibility(CURRENT_INV_FILE, CURRENT_FEAS_FILE)


def stage7_feasibility_open_po():
    build_feasibility(OPEN_PO_FILE, OPEN_PO_FEAS_FILE)


def stage8_feasibility_wip():
    build_feasibility(WIP_FILE, WIP_FEAS_FILE)


# ============================================================================
# Stage 9 (optional): Sync feasibility results back to a Google Sheet
# ============================================================================
def stage9_sync_to_sheets():
    client = get_client(SCOPES_READWRITE)
    sh = client.open_by_key(SYNC_SPREADSHEET_ID)

    for csv_path, tab_name in SYNC_JOBS:
        df = pd.read_csv(csv_path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        try:
            ws = sh.worksheet(tab_name)
        except WorksheetNotFound:
            ws = sh.add_worksheet(title=tab_name, rows=max(len(df) + 1, 100), cols=max(len(df.columns), 26))

        ws.clear()
        values = [df.columns.tolist()] + df.values.tolist()
        needed_rows, needed_cols = len(values), len(df.columns)
        if ws.row_count < needed_rows or ws.col_count < needed_cols:
            ws.resize(rows=max(needed_rows, ws.row_count), cols=max(needed_cols, ws.col_count))
        ws.update(values, value_input_option="USER_ENTERED")
        log.info(f"    Synced {len(df)} rows -> tab '{tab_name}'")


# ============================================================================
# Email notification
# ============================================================================
def _read_csv_rows_any_encoding(path):
    """Read a CSV's rows, tolerating whatever encoding Excel/Notepad saved it in.

    Tries UTF-8 first (most common), then falls back to Windows-1252/Latin-1
    (common when the file was saved from Excel on Windows), and finally
    UTF-8 with invalid bytes replaced so this never hard-crashes on a stray
    character like a non-breaking space (0xA0).
    """
    encodings_to_try = ["utf-8-sig", "cp1252", "latin-1"]
    last_error = None
    for enc in encodings_to_try:
        try:
            with open(path, newline="", encoding=enc) as f:
                return [r for r in csv.reader(f) if any(cell.strip() for cell in r)]
        except UnicodeDecodeError as e:
            last_error = e
            continue
    # Last resort: decode UTF-8, replacing any bad bytes instead of failing
    with open(path, newline="", encoding="utf-8", errors="replace") as f:
        return [r for r in csv.reader(f) if any(cell.strip() for cell in r)]


def read_gmail_credentials(csv_path):
    path = Path(csv_path)
    if not path.exists():
        raise FileNotFoundError(f"Gmail credentials CSV not found at: {csv_path}")

    rows = _read_csv_rows_any_encoding(path)

    if not rows:
        raise ValueError(f"Gmail credentials CSV is empty: {csv_path}")

    row = rows[0]
    # Skip an optional header row like "email,password"
    if row[0].strip().lower() in ("mail", "email", "e-mail", "username"):
        if len(rows) < 2:
            raise ValueError("Gmail credentials CSV only has a header row, no actual credentials.")
        row = rows[1]

    if len(row) < 2:
        raise ValueError("Gmail credentials CSV must have two columns: email, password.")

    # Strip whitespace AND stray non-breaking spaces (\xa0) that Excel sometimes inserts
    email_addr = row[0].strip().replace("\xa0", "")
    password = row[1].strip().replace("\xa0", "")

    if not email_addr or not password:
        raise ValueError("Gmail credentials CSV has a blank email or password.")

    return email_addr, password


def send_email(subject, body_text, body_html):
    if GMAIL_SENDER_EMAIL and GMAIL_APP_PASSWORD:
        sender_email, sender_password = GMAIL_SENDER_EMAIL, GMAIL_APP_PASSWORD
    else:
        sender_email, sender_password = read_gmail_credentials(GMAIL_CREDENTIALS_CSV)
    to_addr = EMAIL_TO or sender_email

    # Fail fast with a clear message if any address still has non-ASCII
    # characters (e.g. a stray \xa0 or other invisible unicode character),
    # instead of letting smtplib raise an opaque UnicodeEncodeError later.
    for field_name, value in (("sender_email", sender_email), ("to_addr", to_addr)):
        try:
            value.encode("ascii")
        except UnicodeEncodeError as e:
            bad_char = value[e.start:e.end]
            raise RuntimeError(
                f"{field_name}={value!r} contains a non-ASCII character "
                f"({bad_char!r}, U+{ord(bad_char):04X}) at position {e.start}. "
                "This is usually a non-breaking space or other invisible character "
                "picked up from copy-pasting into a GitHub secret or spreadsheet cell. "
                "Re-type the value manually to fix it."
            ) from e

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = sender_email
    msg["To"] = to_addr
    msg.attach(MIMEText(body_text, "plain"))
    msg.attach(MIMEText(body_html, "html"))

    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, to_addr, msg.as_string())

    log.info(f"Notification email sent to {to_addr}")


def build_success_body(completed_stages):
    timestamp = datetime.now().strftime("%a, %b %d, %Y %I:%M %p")
    lines = "\n".join(f" - {s}" for s in completed_stages)
    text = f"The Fabric KPI pipeline ran SUCCESSFULLY. All stages completed:\n\n{lines}\n\nCompleted at: {timestamp}"

    html_items = "".join(f"<li>{s}</li>" for s in completed_stages)
    html = f"""\
<html><body style="font-family:Arial,sans-serif;color:#202124;">
  <div style="display:flex;align-items:center;gap:8px;margin-bottom:16px;">
    <span style="font-size:22px;">✅</span>
    <span style="font-size:20px;font-weight:600;">Fabric KPI Pipeline — Success</span>
  </div>
  <p>The Fabric KPI pipeline ran <b>SUCCESSFULLY</b>. All stages completed:</p>
  <ul>{html_items}</ul>
  <p style="color:#5f6368;font-size:12px;">Completed at: {timestamp}</p>
</body></html>
"""
    return text, html


def build_failure_body(completed_stages, failed_stage, error, tb_text):
    timestamp = datetime.now().strftime("%a, %b %d, %Y %I:%M %p")
    completed_lines = "\n".join(f" - {s}" for s in completed_stages) or " (none)"
    text = (
        f"The Fabric KPI pipeline FAILED at stage: {failed_stage}\n\n"
        f"Error: {error}\n\n"
        f"Stages completed before failure:\n{completed_lines}\n\n"
        f"Traceback:\n{tb_text}\n\n"
        f"Failed at: {timestamp}"
    )
    completed_items = "".join(f"<li>{s}</li>" for s in completed_stages) or "<li>(none)</li>"
    html = f"""\
<html><body style="font-family:Arial,sans-serif;color:#202124;">
  <div style="display:flex;align-items:center;gap:8px;margin-bottom:16px;">
    <span style="font-size:22px;">❌</span>
    <span style="font-size:20px;font-weight:600;color:#c5221f;">Fabric KPI Pipeline — FAILED</span>
  </div>
  <p>The pipeline failed at stage: <b>{failed_stage}</b></p>
  <p style="color:#c5221f;"><b>Error:</b> {error}</p>
  <p>Stages completed before failure:</p>
  <ul>{completed_items}</ul>
  <pre style="background:#f1f3f4;padding:12px;border-radius:6px;font-size:12px;white-space:pre-wrap;">{tb_text}</pre>
  <p style="color:#5f6368;font-size:12px;">Failed at: {timestamp}</p>
</body></html>
"""
    return text, html


# ============================================================================
# Main
# ============================================================================
def main():
    stages = [
        ("SKU x FAB mapping (Google Sheet extract)", stage1_extract_mapping),
        ("Fabric Current Inventory (Google Sheet extract)", stage2_extract_current_inventory),
        ("Open PO (Google Sheet extract)", stage3_extract_open_po),
        ("WIP (Google Sheet extract)", stage4_extract_wip),
        ("SKU x Fab Consumption (unpivot)", stage5_build_consumption),
        ("Feasibility vs Current Inventory", stage6_feasibility_current),
        ("Feasibility vs Open PO", stage7_feasibility_open_po),
        ("Feasibility vs WIP", stage8_feasibility_wip),
    ]
    if ENABLE_SYNC_TO_SHEETS:
        stages.append(("Sync results to Google Sheet", stage9_sync_to_sheets))

    completed = []

    log.info("=" * 70)
    log.info("Starting Fabric KPI Pipeline")
    log.info("=" * 70)

    for name, func in stages:
        log.info(f"--> {name}")
        try:
            func()
            completed.append(name)
        except Exception as e:
            tb_text = traceback.format_exc()
            log.error(f"Stage FAILED: {name}\n{tb_text}")
            text, html = build_failure_body(completed, name, e, tb_text)
            try:
                send_email(EMAIL_SUBJECT_FAILURE, text, html)
            except Exception as mail_err:
                log.error(f"Additionally failed to send failure notification email: {mail_err}")
            sys.exit(1)

    log.info("=" * 70)
    log.info("All stages completed successfully.")
    log.info("=" * 70)

    text, html = build_success_body(completed)
    try:
        send_email(EMAIL_SUBJECT_SUCCESS, text, html)
    except Exception as mail_err:
        log.error(f"Pipeline succeeded but failed to send success email: {mail_err}")
        sys.exit(2)


if __name__ == "__main__":
    main()
