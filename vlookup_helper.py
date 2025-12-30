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


def perform_vlookup(app, sheet, multi_key: bool = False):
    """
    Perform left-join (VLOOKUP-like) merging the current sheet's filtered DataFrame
    with a user-selected lookup file.

    Parameters:
      - app: reference to main app (for parent windows, updating)
      - sheet: the sheet dictionary (as used in your main app)
      - multi_key (bool): if True, user can specify multiple key columns separated by commas.

    Returns:
      - merged DataFrame on success, or None on cancel/error.
    """

    # derive the DataFrame that will be merged into (the sheet's current result)
    try:
        main_df = app._generate_filtered_df(sheet)
    except Exception as e:
        messagebox.showerror("VLOOKUP", f"Failed to get main sheet data: {e}")
        return None

    if main_df is None or main_df.empty:
        messagebox.showwarning("VLOOKUP", "No data available in the current sheet to perform VLOOKUP on.")
        return None

    # Ask user to pick lookup file
    lookup_path = filedialog.askopenfilename(
        title="Select lookup file (Excel or CSV)",
        filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")]
    )
    if not lookup_path:
        return None

    # Load lookup DataFrame
    try:
        if lookup_path.lower().endswith(".csv"):
            lookup_df = pd.read_csv(lookup_path)
        else:
            lookup_df = pd.read_excel(lookup_path)
    except Exception as e:
        messagebox.showerror("VLOOKUP", f"Failed to load lookup file:\n{e}")
        return None

    if lookup_df is None or lookup_df.empty:
        messagebox.showwarning("VLOOKUP", "Lookup file is empty.")
        return None

    # Columns lists
    main_cols = list(main_df.columns.astype(str))
    lookup_cols = list(lookup_df.columns.astype(str))

    # Ask for keys (support multi-key if user chooses)
    if not multi_key:
        key_main = _ask_choice("Enter key column in main sheet (exact name):", main_cols, parent=app)
        if not key_main:
            return None
        key_lookup = _ask_choice("Enter key column in lookup file (exact name):", lookup_cols, parent=app)
        if not key_lookup:
            return None

        keys_main = [key_main]
        keys_lookup = [key_lookup]
    else:
        # ask for comma-separated list, validate each
        prompt_main = "Enter comma-separated key columns in main sheet (in order):"
        prompt_lookup = "Enter comma-separated key columns in lookup file (in same order):"
        raw_main = simpledialog.askstring("VLOOKUP", f"{prompt_main}\nAvailable: {', '.join(main_cols)}", parent=app)
        if not raw_main:
            return None
        keys_main = [c.strip() for c in raw_main.split(",") if c.strip()]
        if any(k not in main_cols for k in keys_main):
            messagebox.showerror("VLOOKUP", "One or more main sheet keys are invalid.")
            return None

        raw_lookup = simpledialog.askstring("VLOOKUP", f"{prompt_lookup}\nAvailable: {', '.join(lookup_cols)}", parent=app)
        if not raw_lookup:
            return None
        keys_lookup = [c.strip() for c in raw_lookup.split(",") if c.strip()]
        if any(k not in lookup_cols for k in keys_lookup):
            messagebox.showerror("VLOOKUP", "One or more lookup file keys are invalid.")
            return None

        if len(keys_main) != len(keys_lookup):
            messagebox.showerror("VLOOKUP", "Number of keys on both sides must match.")
            return None

    # Ask which lookup value columns to bring (allow multiple, comma separated)
    raw_vals = simpledialog.askstring("VLOOKUP",
                                      f"Enter lookup column(s) to fetch (comma-separated):\nAvailable: {', '.join(lookup_cols)}",
                                      parent=app)
    if not raw_vals:
        return None
    val_cols = [c.strip() for c in raw_vals.split(",") if c.strip()]
    if any(v not in lookup_cols for v in val_cols):
        messagebox.showerror("VLOOKUP", "One or more selected value columns not found.")
        return None

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

    # Prepare for merge: rename lookup keys to match main keys if necessary
    try:
        # If key names identical length >1, create list-of-tuples merges
        left_on = keys_main
        right_on = keys_lookup

        # Select subset of lookup df with right keys + value columns
        sel = right_on + val_cols
        lookup_sub = lookup_df.loc[:, sel].copy()

        # Perform the merge
        merged = main_df.merge(lookup_sub, how="left", left_on=left_on, right_on=right_on, suffixes=("", "_lk"))

        # If lookup used different key names, drop duplicate right-key columns if present
        for rk, lk in zip(right_on, left_on):
            if rk != lk and rk in merged.columns:
                merged.drop(columns=[rk], inplace=True)

        # Rename added value columns to have prefix if provided (avoid overwriting existing)
        for v in val_cols:
            if v in merged.columns:
                new_name = f"{prefix}{v}" if prefix else v
                # If name collides, attempt to make unique
                if new_name in main_df.columns:
                    # append "_lk" or numeric suffix
                    base = new_name
                    i = 1
                    while new_name in merged.columns:
                        new_name = f"{base}_lk{i}"
                        i += 1
                merged.rename(columns={v: new_name}, inplace=True)

                # fill missing
                if default_fill is not None:
                    merged[new_name].fillna(default_fill, inplace=True)

        messagebox.showinfo("VLOOKUP", f"VLOOKUP merge complete — added: {', '.join([ (prefix + v if prefix else v) for v in val_cols]) }")
        return merged

    except Exception as e:
        messagebox.showerror("VLOOKUP", f"Merge failed: {e}")
        return None
