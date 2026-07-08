# vlookup_helper.py
"""
VLOOKUP helper for ExcelOps.

Usage:
- Called from main.py. It will prompt the user to select a lookup file (xlsx/csv)
  then ask for key columns and value column(s) and perform a left-join.
- Returns the merged DataFrame (or None on cancel or error).
"""

import pandas as pd
from tkinter import filedialog, simpledialog, messagebox


def _dedupe_keep_order(items):
    seen = set()
    out = []
    for item in items:
        if item in seen:
            continue
        seen.add(item)
        out.append(item)
    return out


def _find_duplicate_columns(df: pd.DataFrame):
    cols = [str(c) for c in df.columns]
    seen = set()
    dupes = []
    for col in cols:
        if col in seen and col not in dupes:
            dupes.append(col)
        seen.add(col)
    return dupes



def _normalize_key_series(s: pd.Series) -> pd.Series:
    """Normalize VLOOKUP keys so Excel/CSV type differences still match."""
    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s, errors="coerce").dt.strftime("%Y-%m-%d").fillna("")

    text = (
        s.astype(str)
        .str.replace("\u00a0", " ", regex=False)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )
    text = text.replace({"nan": "", "nat": "", "none": ""})

    non_empty = text[text != ""]
    if not non_empty.empty and non_empty.str.contains(r"[/-]", regex=True).mean() >= 0.6:
        parsed_dates = pd.to_datetime(text, errors="coerce", dayfirst=True)
        if parsed_dates.notna().mean() >= 0.6:
            return parsed_dates.dt.strftime("%Y-%m-%d").fillna("")

    return text.str.casefold()


def _read_lookup_file(app, lookup_path: str) -> pd.DataFrame:
    if lookup_path.lower().endswith(".csv") and hasattr(app, "_read_csv_safely"):
        return app._read_csv_safely(lookup_path)
    if lookup_path.lower().endswith(".csv"):
        return pd.read_csv(lookup_path)
    return pd.read_excel(lookup_path)


def _ask_choice(prompt: str, options: list[str], parent=None) -> str | None:
    """Simple helper that asks user for a string; validates it is in options."""
    if not options:
        messagebox.showerror("VLOOKUP", "No columns available for selection.")
        return None
    text = f"{prompt}\nAvailable: {', '.join(options)}"
    val = simpledialog.askstring("VLOOKUP", text, parent=parent)
    if val is None:
        return None
    val = val.strip()
    if val not in options:
        messagebox.showerror("VLOOKUP", f"'{val}' is not a valid option.")
        return None
    return val


def perform_vlookup(
    app,
    sheet,
    preset: dict | None = None,
    lookup_df: pd.DataFrame | None = None,
    lookup_path: str | None = None,
    interactive: bool = True,
):
    """
    Perform left-join (VLOOKUP-like) merging the current sheet's filtered DataFrame
    with a user-selected lookup file.

    Parameters:
      - app: reference to main app (for parent windows, updating)
      - sheet: the sheet dictionary (as used in your main app)

    Returns:
      - merged DataFrame on success, or None on cancel/error.
    """

    # derive the DataFrame that will be merged into (the sheet's current result)
    try:
        if hasattr(app, "_generate_base_df"):
            main_df = app._generate_base_df(sheet)
        else:
            main_df = app._generate_filtered_df(sheet)
    except Exception as e:
        messagebox.showerror("VLOOKUP", f"Failed to get main sheet data: {e}")
        return None

    if main_df is None or main_df.empty:
        messagebox.showwarning("VLOOKUP", "No data available in the current sheet to perform VLOOKUP on.")
        return None

    # Ask user to pick lookup file if one is not already selected from the VLOOKUP tab.
    if lookup_df is None:
        if not interactive:
            messagebox.showwarning("VLOOKUP", "Lookup file is required in preset to auto-run VLOOKUP.")
            return None
        lookup_path = filedialog.askopenfilename(
            title="Select lookup file (Excel or CSV)",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")]
        )
        if not lookup_path:
            return None

        # Load lookup DataFrame
        try:
            lookup_df = _read_lookup_file(app, lookup_path)
        except Exception as e:
            messagebox.showerror("VLOOKUP", f"Failed to load lookup file:\n{e}")
            return None

    if lookup_df is None or lookup_df.empty:
        messagebox.showwarning("VLOOKUP", "Lookup file is empty.")
        return None

    # Columns lists
    main_cols = list(main_df.columns.astype(str))
    lookup_cols = list(lookup_df.columns.astype(str))
    main_col_map = {c.strip().lower(): c for c in main_cols}
    lookup_col_map = {c.strip().lower(): c for c in lookup_cols}

    preset = preset or {}

    if preset.get("main_keys"):
        keys_main = [c.strip() for c in preset.get("main_keys", "").split(",") if c.strip()]
        raw_lookup = preset.get("lookup_keys", "").strip()
        if raw_lookup:
            keys_lookup = [c.strip() for c in raw_lookup.split(",") if c.strip()]
        else:
            keys_lookup = list(keys_main)
    else:
        keys_main = []
        keys_lookup = []

    if not keys_main or not keys_lookup:
        key_main = _ask_choice("Enter key column in main sheet (exact name):", main_cols, parent=app)
        if not key_main:
            return None
        key_lookup = _ask_choice("Enter key column in lookup file (exact name):", lookup_cols, parent=app)
        if not key_lookup:
            return None

        keys_main = [key_main]
        keys_lookup = [key_lookup]

    if preset.get("values"):
        val_cols = [c.strip() for c in preset.get("values", "").split(",") if c.strip()]
    else:
        val_cols = []

    if not val_cols:
        if not interactive:
            messagebox.showwarning("VLOOKUP", "Lookup value columns are required to auto-run VLOOKUP.")
            return None
        # Ask which lookup value columns to bring (allow multiple, comma separated)
        raw_vals = simpledialog.askstring(
            "VLOOKUP",
            f"Enter lookup column(s) to fetch (comma-separated):\nAvailable: {', '.join(lookup_cols)}",
            parent=app
        )
        if not raw_vals:
            return None
        val_cols = [c.strip() for c in raw_vals.split(",") if c.strip()]
        missing_vals = [v for v in val_cols if v.strip().lower() not in lookup_col_map]
        if missing_vals:
            messagebox.showerror("VLOOKUP", f"Lookup value column(s) not found: {', '.join(missing_vals)}")
            return None

    normalized_main_keys = []
    for key in keys_main:
        normalized = main_col_map.get(key.strip().lower())
        if not normalized:
            messagebox.showerror("VLOOKUP", f"Main key not found: {key}")
            return None
        normalized_main_keys.append(normalized)

    normalized_lookup_keys = []
    for key in keys_lookup:
        normalized = lookup_col_map.get(key.strip().lower())
        if not normalized:
            messagebox.showerror(
                "VLOOKUP",
                f"Lookup key not found: {key}\nAvailable lookup columns: {', '.join(lookup_cols)}"
            )
            return None
        normalized_lookup_keys.append(normalized)

    normalized_values = []
    for col in val_cols:
        normalized = lookup_col_map.get(col.strip().lower())
        if not normalized:
            messagebox.showerror("VLOOKUP", f"Lookup value column not found: {col}")
            return None
        normalized_values.append(normalized)

    if len(normalized_main_keys) != len(normalized_lookup_keys):
        messagebox.showerror("VLOOKUP", "Number of keys on both sides must match.")
        return None

    main_dupes = _find_duplicate_columns(main_df)
    if main_dupes:
        messagebox.showerror("VLOOKUP", f"Main file has duplicate column names: {', '.join(main_dupes)}")
        return None

    lookup_dupes = _find_duplicate_columns(lookup_df)
    if lookup_dupes:
        messagebox.showerror("VLOOKUP", f"Lookup file has duplicate column names: {', '.join(lookup_dupes)}")
        return None

    prefix = preset.get("prefix", "")
    default_fill = preset.get("default_fill", "")
    prefix = prefix.strip() if isinstance(prefix, str) else ""
    default_fill = default_fill if default_fill != "" else None

    if not preset.get("prefix") and not preset.get("default_fill") and interactive:
        # New column prefix / name handling: ask for a prefix to avoid collisions
        prefix = simpledialog.askstring("VLOOKUP", "Enter prefix for added columns (leave blank for none):", parent=app)
        if prefix is None:
            return None
        prefix = prefix.strip()

        # Ask fill value for missing matches
        default_fill = simpledialog.askstring("VLOOKUP", "Enter default value for missing matches (leave blank for None):", parent=app)
        if default_fill is None:
            return None
        default_fill = default_fill if default_fill != "" else None

    # Prepare for merge: normalize join keys and exclude key columns from value fetch.
    try:
        left_on = _dedupe_keep_order(normalized_main_keys)
        right_on = _dedupe_keep_order(normalized_lookup_keys)

        key_set = set(right_on)
        lookup_value_cols = _dedupe_keep_order([c for c in normalized_values if c not in key_set])

        main_work = main_df.copy()
        lookup_sub = lookup_df.loc[:, _dedupe_keep_order(right_on + lookup_value_cols)].copy()

        left_tmp_keys = []
        right_tmp_keys = []
        for i, (left_col, right_col) in enumerate(zip(left_on, right_on)):
            left_tmp = f"__excelops_vlookup_left_key_{i}__"
            right_tmp = f"__excelops_vlookup_right_key_{i}__"
            main_work[left_tmp] = _normalize_key_series(main_work[left_col])
            lookup_sub[right_tmp] = _normalize_key_series(lookup_sub[right_col])
            left_tmp_keys.append(left_tmp)
            right_tmp_keys.append(right_tmp)

        # Excel VLOOKUP uses the first lookup hit. Prevent duplicate lookup keys from multiplying rows.
        lookup_sub = lookup_sub.drop_duplicates(subset=right_tmp_keys, keep="first")

        added_cols = []
        rename_map = {}
        reserved = set(str(c) for c in main_work.columns)
        for col in lookup_value_cols:
            base_name = f"{prefix}{col}" if prefix else col
            new_name = base_name
            suffix_i = 1
            while new_name in reserved:
                new_name = f"{base_name}_lk{suffix_i}"
                suffix_i += 1
            reserved.add(new_name)
            rename_map[col] = new_name
            added_cols.append(new_name)

        if rename_map:
            lookup_sub.rename(columns=rename_map, inplace=True)

        merge_cols = right_tmp_keys + added_cols
        merged = main_work.merge(
            lookup_sub.loc[:, merge_cols],
            how="left",
            left_on=left_tmp_keys,
            right_on=right_tmp_keys,
            suffixes=("", "_lk"),
        )

        merged.drop(columns=left_tmp_keys + right_tmp_keys, inplace=True, errors="ignore")

        if default_fill is not None:
            for col in added_cols:
                if col in merged.columns:
                    merged[col] = merged[col].fillna(default_fill)

        if added_cols:
            messagebox.showinfo("VLOOKUP", f"VLOOKUP merge complete — added: {', '.join(added_cols)}")
        else:
            messagebox.showinfo("VLOOKUP", "VLOOKUP match complete — no value columns added (keys-only match).")
        return merged

    except Exception as e:
        messagebox.showerror("VLOOKUP", f"Merge failed: {e}")
        return None
