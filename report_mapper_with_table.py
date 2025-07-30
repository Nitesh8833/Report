# ==== Roster Mapper — Cloud SQL Source (Single Cell, SOURCE -> OUTPUT mapping) ==========
# This cell replaces the file-based source with a Cloud SQL table.
#   • Reads from a Cloud SQL (MySQL/PostgreSQL) table or SQL query via SQLAlchemy
#   • Uses a mapping where LEFT = SOURCE header (column in DB) and RIGHT = OUTPUT header
#   • Applies robust header matching (case-insensitive; ignores spaces/_/-/.)
#   • Writes to .xlsx (auto-sized columns + frozen header) or .csv
#
# REQUIREMENTS (install as needed):
#   pip install pandas SQLAlchemy cloud-sql-python-connector pymysql pg8000 openpyxl
#
# HOW TO USE:
#   1) Fill the CLOUD SQL connection settings below.
#   2) Choose DB_TYPE = "mysql" or "postgres".
#   3) Set USE_CONNECTOR=True for the Cloud SQL Python Connector (recommended when no public IP),
#      otherwise set USE_CONNECTOR=False and provide HOST/PORT for a direct TCP connection.
#   4) Either set SQL_QUERY (explicit SELECT ...) or set TABLE_NAME to read all rows.
#   5) Set OUT_PATH (.xlsx recommended) and adjust MAPPING so LEFT=source col name, RIGHT=desired output col name.
#   6) Run the cell.

import re
import os
from pathlib import Path
from typing import Optional, Dict, List, Tuple, Callable
import pandas as pd
import sqlalchemy

# ===================== OUTPUT TARGET =====================
OUT_PATH = r"D:\Python_Task\Data2\roster_from_cloudsql_output.xlsx"  # .xlsx or .csv

# ===================== CLOUD SQL SETTINGS =====================
DB_TYPE = "mysql"        # "mysql" or "postgres"
USE_CONNECTOR = True     # True -> Cloud SQL Python Connector; False -> TCP (HOST/PORT)

# --- Common (both MySQL & Postgres) ---
DB_USER = os.getenv("DB_USER", "myuser")
DB_PASSWORD = os.getenv("DB_PASSWORD", "mypassword")
DB_NAME = os.getenv("DB_NAME", "mydatabase")

# --- Connector settings (no HOST/PORT needed) ---
# Format: "<PROJECT_ID>:<REGION>:<INSTANCE_NAME>"
INSTANCE_CONNECTION_NAME = os.getenv("INSTANCE_CONNECTION_NAME", "my-project:us-central1:my-instance")
USE_PRIVATE_IP = False  # set True if your instance has a private IP and you're on the same VPC

# --- TCP settings (public IP / direct connection) ---
HOST = os.getenv("DB_HOST", "127.0.0.1")
PORT = int(os.getenv("DB_PORT", "3306"))  # for Postgres this is usually 5432

# ===================== SOURCE SELECTION =====================
# Provide ONE of the following:
SQL_QUERY: Optional[str] = None
# Example:
# SQL_QUERY = """
#   SELECT
#     business_owner,
#     group_type,
#     roster_id AS `Roster_ID`,
#     `Provider Entity`,
#     parent_transaction_type AS Parent_Transaction_Type,
#     transaction_type AS Transaction_Type,
#     `Total Number of Rows With Errors`,
#     `Critical Error Codes`,
#     `Error Description`
#   FROM my_table
# """

TABLE_NAME: Optional[str] = "my_table"  # set None if you provided SQL_QUERY above

# ===================== MAPPING (SOURCE -> OUTPUT) =====================
# LEFT side = source column in DB (robustly matched), RIGHT side = desired output
MAPPING: Dict[str, str] = {
    # If your DB uses underscores, keep them on the LEFT (source).
    # If your DB has spaces/case, put them as-is; matching is robust.
    "business_owner": "business_owner",
    "group_type": "group_type",
    "roster_id": "roster_id",                              # e.g., DB col: roster_id   -> output: roster_id
    "Provider Entity": "roster_name",                      # e.g., DB col: Provider Entity -> output: roster_name
    "Parent_Transaction_Type": "parent_transaction_type",  # e.g., DB col: Parent_Transaction_Type -> output: ...
    "Transaction_Type": "transaction_type",
    "Total Number of Rows With Errors": "total_rows_with_errors",
    "Critical Error Codes": "critical_error_codes",
    "Error Description": "error_details",
}

# ===================== HELPERS =====================
def _normalize(s: str) -> str:
    """Normalize a header for robust matching: lower, strip, remove spaces/_/-/dots."""
    return re.sub(r"[ \t\-_\.]+", "", str(s).strip().lower())

def make_engine_cloudsql_via_connector() -> Tuple[sqlalchemy.Engine, Optional[Callable[[], None]]]:
    """Create SQLAlchemy engine using Cloud SQL Python Connector. Returns (engine, closer)."""
    from google.cloud.sql.connector import Connector, IPTypes
    connector = Connector()

    if DB_TYPE == "mysql":
        def getconn():
            conn = connector.connect(
                INSTANCE_CONNECTION_NAME,
                "pymysql",
                user=DB_USER,
                password=DB_PASSWORD,
                db=DB_NAME,
                ip_type=IPTypes.PRIVATE if USE_PRIVATE_IP else IPTypes.PUBLIC,
            )
            return conn
        engine = sqlalchemy.create_engine("mysql+pymysql://", creator=getconn)
    elif DB_TYPE == "postgres":
        def getconn():
            conn = connector.connect(
                INSTANCE_CONNECTION_NAME,
                "pg8000",
                user=DB_USER,
                password=DB_PASSWORD,
                db=DB_NAME,
                ip_type=IPTypes.PRIVATE if USE_PRIVATE_IP else IPTypes.PUBLIC,
            )
            return conn
        engine = sqlalchemy.create_engine("postgresql+pg8000://", creator=getconn)
    else:
        raise ValueError("DB_TYPE must be 'mysql' or 'postgres'")

    def closer():
        try:
            connector.close()
        except Exception:
            pass

    return engine, closer

def make_engine_tcp() -> Tuple[sqlalchemy.Engine, Optional[Callable[[], None]]]:
    """Create SQLAlchemy engine for direct TCP connections (public IP)."""
    if DB_TYPE == "mysql":
        url = sqlalchemy.engine.URL.create(
            drivername="mysql+pymysql",
            username=DB_USER,
            password=DB_PASSWORD,
            host=HOST,
            port=PORT,
            database=DB_NAME,
        )
    elif DB_TYPE == "postgres":
        url = sqlalchemy.engine.URL.create(
            drivername="postgresql+pg8000",
            username=DB_USER,
            password=DB_PASSWORD,
            host=HOST,
            port=PORT if PORT else 5432,
            database=DB_NAME,
        )
    else:
        raise ValueError("DB_TYPE must be 'mysql' or 'postgres'")
    engine = sqlalchemy.create_engine(url)
    return engine, None

def read_from_cloudsql() -> pd.DataFrame:
    """Read a DataFrame from Cloud SQL via SQLAlchemy."""
    if USE_CONNECTOR:
        engine, closer = make_engine_cloudsql_via_connector()
    else:
        engine, closer = make_engine_tcp()

    try:
        if SQL_QUERY:
            sql = SQL_QUERY
        else:
            if not TABLE_NAME:
                raise ValueError("Either SQL_QUERY must be provided or TABLE_NAME must be set.")
            # Build a simple SELECT * for the table name.
            # For Postgres with schema, pass something like TABLE_NAME = 'public.my_table'
            sql = f"SELECT * FROM {TABLE_NAME}"
        df = pd.read_sql(sql, engine)
        return df
    finally:
        try:
            engine.dispose()
        except Exception:
            pass
        if closer:
            try:
                closer()
            except Exception:
                pass

def build_renamer(df: pd.DataFrame, mapping_src_to_out: Dict[str, str]) -> Dict[str, str]:
    """
    Build a renamer dict {source_col_name_in_df: output_col_name} using robust matching.
    *** mapping is SOURCE -> OUTPUT ***
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
    """Select & rename (SOURCE -> OUTPUT); create missing outputs as empty; reorder; lightly clean strings."""
    renamer = build_renamer(df, mapping_src_to_out)
    selected = df[list(renamer.keys())].rename(columns=renamer) if renamer else pd.DataFrame()

    # Ensure all desired output columns exist
    for out_col in output_order:
        if out_col not in selected.columns:
            selected[out_col] = pd.NA

    # Reorder columns to final output order
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
    ws.freeze_panes = "A2"

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

# --------------------------------- RUN ---------------------------------------
try:
    # 1) Read source from Cloud SQL
    df_src = read_from_cloudsql()
    print(f"[INFO] Loaded from Cloud SQL. shape={df_src.shape}")

    # 2) Determine OUTPUT order (right-hand side of mapping)
    OUTPUT_ORDER: List[str] = list(MAPPING.values())

    # 3) Map & extract (SOURCE -> OUTPUT)
    out_df = extract_and_rename(df_src, MAPPING, OUTPUT_ORDER)
    print(f"[INFO] Prepared output shape: {out_df.shape}")

    # 4) Write output
    Path(OUT_PATH).parent.mkdir(parents=True, exist_ok=True)
    write_output(out_df, OUT_PATH)

    # 5) Preview
    try:
        display(out_df.head(10))
    except Exception:
        print(out_df.head(10))

except Exception as e:
    print("[ERROR]", type(e).__name__, "-", e)
    raise
