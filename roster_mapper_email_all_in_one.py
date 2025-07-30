#!/usr/bin/env python
# -*- coding: utf-8 -*-
# ==== Roster Mapper — All-in-One (SOURCE -> OUTPUT mapping) + Email Sender ======
# This script:
#   • Reads a source file (.xlsx or .csv). Optional: specific sheet for Excel
#   • Uses a mapping where LEFT side is the **SOURCE header** in your file
#     and RIGHT side is the **OUTPUT header** you want in the result
#   • Maps & extracts required columns using robust header matching
#   • Writes to .xlsx (auto-sized columns + frozen header) or .csv
#   • (Optional) Emails the generated report as an attachment
# ===============================================================================

import re
from pathlib import Path
from typing import Optional, Dict, List, Iterable, Union
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
# Example (human-friendly headers):
# MAPPING = {
#     "Business Team": "business_owner",
#     "Group Team": "group_type",
#     "Roster ID": "roster_id",
#     "Provider Entity": "roster_name",
#     "Parent Transaction Type": "parent_transaction_type",
#     "Transaction Type": "transaction_type",
#     "Total Number of Rows With Error": "total_rows_with_errors",
#     "Critical Error Codes": "critical_error_codes",
#     "Error Description": "error_details",
# }
# If your source file uses underscore headers, put those on the LEFT.
MAPPING: Dict[str, str] = {
    "Business Team": "business_owner",
    "Group Team": "group_type",
    "Roster ID": "roster_id",
    "Provider Entity": "roster_name",
    "Parent Transaction Type": "parent_transaction_type",
    "Transaction Type": "transaction_type",
    "Total Number of Rows With Error": "total_rows_with_errors",
    "Critical Error Codes": "critical_error_codes",
    "Error Description": "error_details",
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
    # Create a lookup from normalized source headers in the DataFrame to their real names
    norm_to_real = {_normalize(col): col for col in df.columns}

    renamer: Dict[str, str] = {}
    matched, missing = 0, 0
    for src_label, out_col in mapping_src_to_out.items():
        norm_key = _normalize(src_label)
        if norm_key in norm_to_real:
            src_real = norm_to_real[norm_key]         # actual column name in df
            renamer[src_real] = out_col               # rename to desired output column
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

    # Reorder columns to OUTPUT order
    selected = selected[output_order]

    # Light string cleanup
    for c in selected.columns:
        if pd.api.types.is_string_dtype(selected[c]):
            selected[c] = selected[c].astype("string").str.strip()

    return selected


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
        # Ensure openpyxl is available
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
        # The final OUTPUT order is the set of desired output names (RIGHT-hand side).
        OUTPUT_ORDER: List[str] = list(MAPPING.values())

        # 1) Read source
        df_src = read_source(SRC_PATH, SHEET)
        print(f"[INFO] Loaded source: {SRC_PATH}  shape={df_src.shape}")

        # 2) Map & extract (SOURCE -> OUTPUT)
        out_df = extract_and_rename(df_src, MAPPING, OUTPUT_ORDER)
        print(f"[INFO] Prepared output shape: {out_df.shape}")

        # 3) Write output
        Path(OUT_PATH).parent.mkdir(parents=True, exist_ok=True)
        write_output(out_df, OUT_PATH)

        # 4) Optional: Send email with the generated file
        if SEND_EMAIL:
            # Enrich email body with quick stats
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
