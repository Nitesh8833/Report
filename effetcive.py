def extract_effective_date(ws, search_rows: int = 50, search_cols: int = 40) -> str:
    """
    Robustly find 'Effective as of <date>' (or variants) in the banner area.
    - Scans the first `search_rows` x `search_cols` cells.
    - Works with dates like 12/12/2024, 12-12-2024, 12.12.2024, or 'Dec 12, 2024'.
    - If the cell next to 'Effective...' holds the date (often due to merges), it
      looks rightward in the same row.
    Returns a 'MM/DD/YYYY' string, or '' if not found.
    """
    import re
    import pandas as pd

    # Accept several date formats
    date_tokens = [
        r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b",               # 12/12/2024 or 12-12-2024
        r"\b\d{1,2}\.\d{1,2}\.\d{2,4}\b",                   # 12.12.2024
        r"\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+\d{1,2},\s+\d{4}\b",
    ]
    re_date = re.compile("|".join(date_tokens), re.I)
    re_eff  = re.compile(r"effective\s*(?:as\s*of|date)?", re.I)

    max_r = min(search_rows, ws.max_row)
    max_c = min(search_cols, ws.max_column)

    def _to_mmddyyyy(s: str) -> str:
        dt = pd.to_datetime(s, errors="coerce")
        return dt.strftime("%m/%d/%Y") if pd.notna(dt) else ""

    # Pass 1: same cell contains 'effective' and the date
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            v = ws.cell(r, c).value
            if v is None:
                continue
            s = str(v)
            if re_eff.search(s):
                m = re_date.search(s)
                if m:
                    out = _to_mmddyyyy(m.group(0))
                    if out:
                        return out
                # Often the date is in the next cell(s) on the same row
                for cc in range(c + 1, min(c + 5, max_c) + 1):
                    u = ws.cell(r, cc).value
                    if u is None:
                        continue
                    if hasattr(u, "year"):  # real Excel date
                        return u.strftime("%m/%d/%Y")
                    m2 = re_date.search(str(u))
                    if m2:
                        out = _to_mmddyyyy(m2.group(0))
                        if out:
                            return out

    # Pass 2: any early standalone date value (Excel date or string)
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            v = ws.cell(r, c).value
            if v is None:
                continue
            if hasattr(v, "year"):  # datetime/date cell
                return v.strftime("%m/%d/%Y")
            m = re_date.search(str(v))
            if m:
                out = _to_mmddyyyy(m.group(0))
                if out:
                    return out

    return ""
