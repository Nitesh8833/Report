import re
import pandas as pd

def remove_trailing_notes(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # 1) Keep only rows that have data in at least one key column
    key_cols = [c for c in ["Medicaid Provider #", "Medicare Provider #", "NPI #", "Location Name"] if c in df.columns]
    if key_cols:
        df = df[df[key_cols].notna().any(axis=1)]

    # 2) Drop rows that look like notes/paths in Facility
    if "Facility" in df.columns:
        note_pat = r"(?i)^(note\b|[A-Z]:\\|resource\s+directory|matrix)"
        df = df[~df["Facility"].astype(str).str.strip().str.match(note_pat, na=False)]

        # 3) If a Facility cell has multiple lines, keep only the first line
        df["Facility"] = df["Facility"].astype(str).str.splitlines().str[0].str.strip()
        df.loc[df["Facility"].eq("nan"), "Facility"] = pd.NA  # clean artifact

    # 4) Drop rows that are completely empty after cleanup
    df = df.dropna(how="all").reset_index(drop=True)
    return df
