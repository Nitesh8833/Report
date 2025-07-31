#!/usr/bin/env python
# -*- coding: utf-8 -*-
# ==== Roster Mapper — All-in-One (SOURCE -> OUTPUT mapping) + KPIs + Email =====
# This script:
#   • Reads a source file (.xlsx or .csv). Optional: specific sheet for Excel
#   • Uses a mapping where LEFT side is the SOURCE header in your file
#     and RIGHT side is the OUTPUT header you want in the result
#   • Maps & extracts required columns using robust header matching
#   • Computes KPI columns (see "Derived KPI columns" below)
#   • Writes to .xlsx (auto-sized columns + frozen header) or .csv
#   • (Optional) Emails the generated report as an attachment
# ===============================================================================

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional, Dict, List, Iterable, Union, Tuple
import pandas as pd

# ------------------- Email imports -------------------
import logging
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# ===================== USER PARAMETERS =====================
# 1) Input & Output
SRC_PATH = r"D:\Python_Task\Data2\roster_full_table.xlsx"         # your source .xlsx or .csv
OUT_PATH = r"D:\Python_Task\Data2\roster_full_table_output.xlsx"  # .xlsx (recommended) or .csv
SHEET: Optional[str] = None                                       # Excel sheet name, e.g. "Sheet1"; None = first sheet

# 2) MAPPING (SOURCE -> OUTPUT)
#    LEFT side = source header in your file (robust match)
#    RIGHT side = desired output column name
MAPPING: Dict[str, str] = {
    "Business Team": "business_owner",
    "Group Team": "group_type",
    "Roster ID": "roster_id",
    "Provider Entity": "roster_name",
    "Parent Transaction Type": "parent_transaction_type",
    "Transaction Type": "transaction_type",
}

# 3) Email settings (optional)
SEND_EMAIL: bool = False  # set True to send email
EMAIL_SENDER: str = "rudra.annangi@aetna.com"
EMAIL_TO: Union[str, Iterable[str]] = "recipient@example.com"  # or ["a@x.com", "b@y.com"]
EMAIL_SUBJECT: str = "Roster Report Ready"
EMAIL_TEXT: str = "Hello,\n\nPlease find the latest roster report attached.\n\nRegards,"
EMAIL_HTML: Optional[str] = "<p>Hello,<br><br>Please find the latest <b>roster report</b> attached.<br><br>Regards,</p>"
SMTP_SERVER: str = "extmail.aetna.com"
SMTP_PORT: int = 25
USE_TLS: bool = False
SMTP_USERNAME: Optional[str] = None  # e.g., "user"
SMTP_PASSWORD: Optional[str] = None  # e.g., "pass"

LOG_LEVEL = "INFO"  # DEBUG, INFO, WARNING, ERROR
# ===========================================================


# ------------------------------ Helpers --------------------------------------
def _normalize(s: str) -> str:
    """Normalize a header for robust matching: lower, strip, remove spaces/_/-/dots."""
    return re.sub(r"[ \t\-_\.]+", "", str(s).strip().lower())


def read_source(path: str, sheet: Optional[str]) -> pd.DataFrame:
    """Read .xlsx (optionally with sheet) or .csv based on file extension."""
    ext = Path(path).suffix.lower()
    if ext in [".xlsx", ".xlsm", ".xltx", ".xltm"]:
        print(f"[INFO] Reading Excel: {path} (sheet={sheet})")
        return pd.read_excel(path, sheet_name=sheet) if sheet else pd.read_excel(path)
    elif ext == ".csv":
        print(f"[INFO] Reading CSV: {path}")
        return pd.read_csv(path)
    else:
        raise ValueError(f"Unsupported source extension '{ext}'. Use .xlsx or .csv")


def build_renamer(df: pd.DataFrame, mapping_src_to_out: Dict[str, str]) -> Dict[str, str]:
    """
    Build a renamer dict {source_col_name_in_df: output_col_name} using robust matching.
    *** IMPORTANT: mapping is SOURCE -> OUTPUT ***
    """
    norm_to_real = {_normalize(col): col for col in df.columns}
    renamer: Dict[str, str] = {}
    matched, missing = 0, 0
    for src_label, out_col in mapping_src_to_out.items():
        norm_key = _normalize(src_label)
        if norm_key in norm_to_real:
            src_real = norm_to_real[norm_key]
            renamer[src_real] = out_col
            matched += 1
            print(f"[DEBUG] Matched source '{src_label}' -> '{src_real}'  as output '{out_col}'")
        else:
            missing += 1
            print(f"[WARN] Source column '{src_label}' not found. Output '{out_col}' will be empty.")
    print(f"[INFO] Mapping summary: matched={matched}, missing={missing}")
    return renamer


def extract_and_rename(df: pd.DataFrame, mapping_src_to_out: Dict[str, str],
                       output_order: List[str]) -> pd.DataFrame:
    """
    Select & rename per mapping (SOURCE -> OUTPUT), create missing outputs as empty,
    reorder to output_order, and lightly clean strings.
    """
    renamer = build_renamer(df, mapping_src_to_out)
    selected = df[list(renamer.keys())].rename(columns=renamer) if renamer else pd.DataFrame()

    # Ensure all desired output columns exist (create empty where missing)
    for out_col in output_order:
        if out_col not in selected.columns:
            selected[out_col] = pd.NA

    # Reorder columns to OUTPUT order (derived KPI columns will be appended later)
    selected = selected[output_order]

    # Light string cleanup
    for c in selected.columns:
        if pd.api.types.is_string_dtype(selected[c]):
            selected[c] = selected[c].astype("string").str.strip()

    return selected.reset_index(drop=True)


# ---------- Column lookup & common transforms ----------
def _find(df: pd.DataFrame, *candidates: str) -> Optional[str]:
    """Find a column by normalized name."""
    norm_to_real = {_normalize(c): c for c in df.columns}
    for c in candidates:
        k = _normalize(c)
        if k in norm_to_real:
            return norm_to_real[k]
    return None


def _version_status_series(df_src: pd.DataFrame, df_out: pd.DataFrame) -> Optional[pd.Series]:
    """Return normalized version_status (NEW_FILE/NEW_VERSION/EXISTING_VERSION) aligned with output rows."""
    col = _find(df_src, "version_status", "Version Status")
    if not col or len(df_src) != len(df_out):
        return None
    return (df_src[col].astype("string").str.strip()
            .str.replace(r"[\s\-]+", "_", regex=True).str.upper())


def _to_timedelta_str(td: pd.Timedelta) -> str | pd._libs.NaTType:
    """Format a Timedelta as DD/HH:MM:SS or NA."""
    if pd.isna(td):
        return pd.NA
    total_sec = int(td.total_seconds())
    sign = "-" if total_sec < 0 else ""
    total_sec = abs(total_sec)
    days, rem = divmod(total_sec, 86400)
    hrs, rem = divmod(rem, 3600)
    mins, secs = divmod(rem, 60)
    return f"{sign}{days:02d}/{hrs:02d}:{mins:02d}:{secs:02d}"


# --------------------------- Derived KPI columns -----------------------------
def add_new_roster_formats(
    df_src: pd.DataFrame,
    df_out: pd.DataFrame,
    roster_id_out_col: str = "roster_id",
    include_statuses: Tuple[str, ...] = ("NEW_FILE", "NEW_VERSION"),
) -> pd.DataFrame:
    """
    NEW LOGIC:
      For each roster_id, count rows where version_status in include_statuses.
      Put that SAME COUNT on every row for that roster_id.
      Example: roster_id=76 has three NEW_VERSION rows -> all rows with roster_id=76 get 3.
    """
    out = df_out.copy().reset_index(drop=True)

    if roster_id_out_col not in out.columns:
        out["# New Roster Formats"] = 0
        return out

    vs = _version_status_series(df_src, out)
    if vs is None:
        out["# New Roster Formats"] = 0
        return out

    statuses = {s.upper().replace(" ", "_").replace("-", "_") for s in include_statuses}
    is_new = vs.isin(statuses)  # boolean (may contain NA)
    counts_per_id = is_new.fillna(False).groupby(out[roster_id_out_col]).transform("sum")
    out["# New Roster Formats"] = counts_per_id.fillna(0).astype(int)
    return out


def add_changed_roster_formats(
    df_src: pd.DataFrame,
    df_out: pd.DataFrame,
    roster_id_out_col: str = "roster_id",
) -> pd.DataFrame:
    """
    # Changed Roster Formats (COUNT per roster_id):
      Count "change events" per roster_id where:
        change_event := (version_status == NEW_VERSION) AND (previous status for the same roster_id != NEW_VERSION)
      Broadcast that count to all rows for that roster_id.
    """
    out = df_out.copy().reset_index(drop=True)
    if roster_id_out_col not in out.columns:
        out["# Changed Roster Formats"] = 0
        return out

    vs = _version_status_series(df_src, out)
    if vs is None:
        out["# Changed Roster Formats"] = 0
        return out

    prev_vs = vs.groupby(out[roster_id_out_col]).shift(1)
    change_event = vs.eq("NEW_VERSION") & prev_vs.notna() & (prev_vs != "NEW_VERSION")
    counts_per_id = change_event.fillna(False).groupby(out[roster_id_out_col]).transform("sum")
    out["# Changed Roster Formats"] = counts_per_id.fillna(0).astype(int)
    return out


def add_no_setup_or_format_change(
    df_src: pd.DataFrame,
    df_out: pd.DataFrame,
    roster_id_out_col: str = "roster_id",
    new_col_name: str = "# of Rosters with no Set up or Format Change",
) -> pd.DataFrame:
    """
    New logic (counts all-NA as 'no change' too):
      For each roster_id, if the number of unique non-NA version_status values is
      0 (all NA) or 1 (all the same), mark 1 for all rows of that roster_id; else 0.
    """
    out = df_out.copy().reset_index(drop=True)

    if roster_id_out_col not in out.columns:
        out[new_col_name] = 0
        return out

    vs = _version_status_series(df_src, out)
    if vs is None or len(vs) != len(out):
        out[new_col_name] = 0
        return out

    ids = out[roster_id_out_col]

    # Count unique non-NA statuses per roster_id; nunique ignores NA
    unique_counts = vs.groupby(ids).transform(lambda s: s.dropna().nunique())

    # "No change" if 0 (all NA) or 1 (all same defined status)
    no_change_flag = (unique_counts <= 1).fillna(False)

    out[new_col_name] = no_change_flag.astype(int)
    return out


def add_complex_rosters(
    df_src: pd.DataFrame,
    df_out: pd.DataFrame,
    roster_id_out_col: str = "roster_id",
    complexity_candidates: Tuple[str, ...] = ("complexity", "Complexity"),
    include_values: Tuple[str, ...] = ("COMPLEX",),
    new_col_name: str = "# Complex Rosters",
) -> pd.DataFrame:
    """
    # Complex Rosters (COUNT per roster_id):
      For each roster_id, count rows where complexity equals any of include_values (default: COMPLEX).
      Put that same count on every row for that roster_id.
    """
    out = df_out.copy().reset_index(drop=True)

    if roster_id_out_col not in out.columns:
        out[new_col_name] = 0
        return out

    cx_col = _find(df_src, *complexity_candidates)
    if not cx_col or len(df_src) != len(out):
        out[new_col_name] = 0
        return out

    cx_norm = (
        df_src[cx_col]
        .astype("string").str.strip()
        .str.replace(r"[\s\-]+", "_", regex=True).str.upper()
    )
    target = {v.upper().replace(" ", "_").replace("-", "_") for v in include_values}

    is_complex = cx_norm.isin(target)
    counts_per_id = is_complex.fillna(False).groupby(out[roster_id_out_col]).transform("sum")
    out[new_col_name] = counts_per_id.fillna(0).astype(int)
    return out


def add_all_rosters(
    df_out: pd.DataFrame,
    roster_id_out_col: str = "roster_id",
    new_col_name: str = "All Rosters",
) -> pd.DataFrame:
    """
    All Rosters (COUNT per roster_id):
      For each roster_id, count how many rows exist with that roster_id and
      write the same count to every row for that id.
    """
    out = df_out.copy().reset_index(drop=True)

    if roster_id_out_col not in out.columns:
        out[new_col_name] = 1  # fallback: one per row
        return out

    # Count rows per roster_id (including NaN as a group)
    counts = out[roster_id_out_col].value_counts(dropna=False)
    out[new_col_name] = out[roster_id_out_col].map(counts).astype(int)
    return out


def add_conformance_tat(df_src: pd.DataFrame, df_out: pd.DataFrame) -> pd.DataFrame:
    """
    Conformance TAT = prms_posted_timestamp (or update_timestamp) - file_ingestion_timestamp,
    formatted as DD/HH:MM:SS.
    """
    out = df_out.copy().reset_index(drop=True)
    start_col = _find(df_src, "file_ingestion_timestamp", "ingestion_timestamp")
    end_col = _find(df_src, "prms_posted_timestamp", "update_timestamp", "processed_timestamp")
    if start_col and end_col and len(df_src) == len(out):
        start = pd.to_datetime(df_src[start_col], errors="coerce", utc=True)
        end = pd.to_datetime(df_src[end_col], errors="coerce", utc=True)
        tat = end - start
        out["Conformance TAT"] = tat.map(lambda x: _to_timedelta_str(x) if pd.notna(x) else pd.NA)
    else:
        out["Conformance TAT"] = pd.NA
    return out


def add_rows_counts(df_src: pd.DataFrame, df_out: pd.DataFrame) -> pd.DataFrame:
    """Add # of rows in/out using input_rec_count and conformed_rec_count (if present)."""
    out = df_out.copy().reset_index(drop=True)
    col_in = _find(df_src, "input_rec_count", "input records count")
    col_out = _find(df_src, "conformed_rec_count", "output_rec_count", "conformed records count")
    out["# of rows in"] = df_src[col_in] if col_in and len(df_src) == len(out) else pd.NA
    out["# of rows out"] = df_src[col_out] if col_out and len(df_src) == len(out) else pd.NA
    return out


def add_unique_npi_counts(df_src: pd.DataFrame, df_out: pd.DataFrame) -> pd.DataFrame:
    """Add unique NPI counts from input_unique_npi_count / conformed_unique_npi_count (if present)."""
    out = df_out.copy().reset_index(drop=True)
    col_in = _find(df_src, "input_unique_npi_count", "unique npi input")
    col_out = _find(df_src, "conformed_unique_npi_count", "unique npi output")
    out["# of unique NPI's in Input"] = df_src[col_in] if col_in and len(df_src) == len(out) else pd.NA
    out["# of unique NPI's in Output"] = df_src[col_out] if col_out and len(df_src) == len(out) else pd.NA
    return out


# --------------------------- Excel formatting --------------------------------
def autosize_and_freeze(writer: pd.ExcelWriter, df: pd.DataFrame, sheet_name: str) -> None:
    """Auto-size columns and freeze header (safe: no iterable unpacking)."""
    ws = writer.sheets[sheet_name]
    from openpyxl.utils import get_column_letter

    for idx, col in enumerate(df.columns, start=1):
        series = df[col].astype(str)
        lengths = series.map(len)
        max_cell = int(lengths.max()) if not lengths.empty else 0
        max_len = max(len(str(col)), max_cell)
        ws.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 60)

    ws.freeze_panes = "A2"  # freeze the header row


def write_output(df: pd.DataFrame, out_path: str) -> None:
    """Write DataFrame to .xlsx (with autosize + freeze) or .csv based on OUT_PATH extension."""
    ext = Path(out_path).suffix.lower()
    if ext == ".xlsx":
        try:
            import openpyxl  # noqa: F401
        except ImportError as e:
            raise ImportError("openpyxl is required to write .xlsx files. Install with: pip install openpyxl") from e

        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="output")
            autosize_and_freeze(writer, df, "output")
        print(f"[INFO] Wrote Excel: {out_path}")
    elif ext == ".csv":
        df.to_csv(out_path, index=False)
        print(f"[INFO] Wrote CSV: {out_path}")
    else:
        raise ValueError(f"Unsupported output extension '{ext}'. Use .xlsx or .csv")


# --------------------------- Email helper ------------------------------------
def send_email_alert(
    recipient: Union[str, Iterable[str]],
    subject: str,
    message_body: str,
    html_content: Optional[str] = None,
    attachments: Optional[Iterable[Path]] = None,
    smtp_server: str = SMTP_SERVER,
    smtp_port: int = SMTP_PORT,
    sender: str = EMAIL_SENDER,
    use_tls: bool = USE_TLS,
    username: Optional[str] = SMTP_USERNAME,
    password: Optional[str] = SMTP_PASSWORD,
) -> bool:
    """Send an email alert with optional HTML and attachments."""
    logging.basicConfig(level=getattr(logging, LOG_LEVEL.upper(), logging.INFO),
                        format="%(asctime)s | %(levelname)s | %(message)s")

    # Normalize recipients
    if isinstance(recipient, (list, tuple, set)):
        recipients = list(recipient)
        to_header = ", ".join(recipients)
    else:
        recipients = [str(recipient)]
        to_header = recipients[0]

    msg = MIMEMultipart("alternative")
    msg["From"] = sender
    msg["To"] = to_header
    msg["Subject"] = subject

    # Plain
    msg.attach(MIMEText(message_body, "plain"))
    # HTML optional
    if html_content:
        msg.attach(MIMEText(html_content, "html"))

    # Attachments
    if attachments:
        for p in attachments:
            p = Path(p)
            if not p.exists():
                logging.warning("Attachment not found, skipping: %s", p)
                continue
            with p.open("rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f'attachment; filename="{p.name}"')
            msg.attach(part)

    try:
        logging.info("Connecting SMTP: %s:%s", smtp_server, smtp_port)
        server = smtplib.SMTP(smtp_server, smtp_port, timeout=30)
        if use_tls:
            server.starttls()
        if username and password:
            server.login(username, password)
        logging.info("Sending email to: %s", to_header)
        server.send_message(msg)
        server.quit()
        logging.info("Email sent.")
        return True
    except Exception as e:
        logging.error("Email send failed: %s", e)
        return False


# --------------------------------- RUN ---------------------------------------
if __name__ == "__main__":
    try:
        # The base OUTPUT order is per the RIGHT-hand side names in MAPPING
        OUTPUT_ORDER: List[str] = list(MAPPING.values())

        # 1) Read source
        df_src = read_source(SRC_PATH, SHEET)
        print(f"[INFO] Loaded source: {SRC_PATH}  shape={df_src.shape}")

        # 2) Map & extract (SOURCE -> OUTPUT)
        out_df = extract_and_rename(df_src, MAPPING, OUTPUT_ORDER)
        print(f"[INFO] Prepared output shape: {out_df.shape}")

        # 3) Add derived KPI columns (five modified functions + other KPIs)
        out_df = add_new_roster_formats(df_src, out_df)              # per-roster_id count of NEW_* statuses
        out_df = add_changed_roster_formats(df_src, out_df)          # per-roster_id count of change events
        out_df = add_no_setup_or_format_change(df_src, out_df)       # 1 if all statuses same or all NA for that id
        out_df = add_complex_rosters(df_src, out_df)                 # per-roster_id count of COMPLEX
        out_df = add_all_rosters(out_df)                             # per-roster_id total rows
        out_df = add_conformance_tat(df_src, out_df)
        out_df = add_rows_counts(df_src, out_df)
        out_df = add_unique_npi_counts(df_src, out_df)

        # 4) Write output
        Path(OUT_PATH).parent.mkdir(parents=True, exist_ok=True)
        write_output(out_df, OUT_PATH)

        # 5) Optional: Send email with the generated file
        if SEND_EMAIL:
            row_count = len(out_df)
            col_count = len(out_df.columns)
            text = (EMAIL_TEXT or "") + f"\n\nRows: {row_count} | Columns: {col_count}\nFile: {OUT_PATH}"
            html = EMAIL_HTML or None
            send_email_alert(
                recipient=EMAIL_TO,
                subject=EMAIL_SUBJECT,
                message_body=text,
                html_content=html,
                attachments=[OUT_PATH],
            )
        print("[INFO] Done.")
    except Exception as e:
        print("[ERROR]", type(e).__name__, "-", e)
        raise
