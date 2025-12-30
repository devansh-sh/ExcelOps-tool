# utils.py
import pandas as pd

def clean_numeric_series(s: pd.Series) -> pd.Series:
    """Strip percent signs and commas, coerce to numeric (float)."""
    s2 = s.astype(str).str.strip()
    s2 = s2.str.replace("%", "", regex=False).str.replace(",", "", regex=False)
    s2 = s2.replace({"": None, "nan": None})
    return pd.to_numeric(s2, errors="coerce")
