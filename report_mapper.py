# ==== Roster Mapper — All-in-One (Single Cell, SOURCE -> OUTPUT mapping) ==========
# This single cell:
#   • Reads a source file (.xlsx or .csv). Optional: specific sheet for Excel
#   • Uses a mapping where the LEFT side is the **SOURCE header** in your file
#     and the RIGHT side is the **OUTPUT header** you want in the result
#   • Maps & extracts required columns using robust header matching
#   • Writes to .xlsx (auto-sized columns + frozen header) or .csv
#
# HOW TO USE:
#   1) Set SRC_PATH and OUT_PATH below (and SHEET if needed).
#   2) Update MAPPING so that LEFT = source header in your file, RIGHT = desired output name.
#   3) Run this cell.
#
# NOTES:
#   • Header matching is robust: case-insensitive; ignores spaces/_/-/.
#     e.g., "Business Team", "business_team", "Business-Team", "Business  Team" all match.
#   • Missing source columns are created as empty in the output (stable schema).
#   • If writing .xlsx, you need openpyxl installed (pip install openpyxl).
# ================================================================================

import re
from pathlib import Path
from typing import Optional, Dict, List
import pandas as pd

# ===================== USER PARAMETERS =====================
# 1) Input & Output
SRC_PATH = r"D:\Python_Task\Data2\roster_full_table.xlsx"         # your source .xlsx or .csv
OUT_PATH = r"D:\Python_Task\Data2\roster_full_table_output.xlsx"  # .xlsx (recommended) or .csv
SHEET: Optional[str] = None                                       # Excel sheet name, e.g. "Sheet1"; None = first sheet

# 2) MAPPING (SOURCE -> OUTPUT)
#    LEFT side = exact/approx source header in your file (robust matching)
#    RIGHT side = desired output column name
#    --- Example for human-friendly headers in source: ---
# MAPPING: Dict[str, str] = {
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
#    --- If your source uses underscore headers, use something like: ---
MAPPING = {
    "business_owner": "Business Team",
    "group_type": "Group Team",
    "roster_id": "Roster_ID",
    "roster_name": "Provider Entity",
    "parent_transaction_type": "Parent_Transaction_Type",
    "transaction_type": "Transaction_Type",
    "total_rows_with_errors": "Total Number of Rows With Errors",
    "critical_error_codes": "Critical Error Codes",
    "error_details": "Error Description",
}
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


# --------------------------------- RUN ---------------------------------------
try:
    # The final OUTPUT order is based on the RIGHT-hand side (desired output names)
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

    # 4) Preview first rows (comment out if running outside notebooks)
    try:
        display(out_df.head(10))
    except Exception:
        # display() may not exist outside Jupyter; fallback to print
        print(out_df.head(10))

except Exception as e:
    print("[ERROR]", type(e).__name__, "-", e)
    raise
