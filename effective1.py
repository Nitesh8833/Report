import re
import pandas as pd

def drop_rows_only_effective(df: pd.DataFrame) -> pd.DataFrame:
    # Find the Effective date column (case/spacing-insensitive)
    def norm(s): return re.sub(r"\W+", " ", str(s)).strip().lower()
    eff_col = next((c for c in df.columns if norm(c) == "effective date"), None)
    if eff_col is None:
        return df  # nothing to do

    # Treat pure-blank strings in non-effective columns as NA
    other_cols = [c for c in df.columns if c != eff_col]
    obj_cols = [c for c in other_cols if df[c].dtype == "object"]
    for col in obj_cols:
        s = df[col]
        blank = s.notna() & s.astype(str).str.strip().eq("")
        df.loc[blank, col] = pd.NA

    # Drop rows where all non-effective columns are NA but Effective date has a value
    only_eff_mask = df[other_cols].isna().all(axis=1) & df[eff_col].notna()
    df = df.loc[~only_eff_mask].reset_index(drop=True)
    return df
