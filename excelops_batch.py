#!/usr/bin/env python3
"""
excelops_batch.py

Batch-process an Excel file to produce per-user sheets filtered by Activity and
apply shared edits loaded from (or saved to) a JSON config file.

Usage:
    python3 excelops_batch.py
"""

import json
import os
import sys
import pandas as pd
import numpy as np
from pathlib import Path

DEFAULT_UPLOAD = "/mnt/data/Activity Log_30.09.2025.xlsx"

# Allowed operators
OPS = {
    "==": lambda a, b: a == b,
    "!=": lambda a, b: a != b,
    ">": lambda a, b: a > b,
    "<": lambda a, b: a < b,
    ">=": lambda a, b: a >= b,
    "<=": lambda a, b: a <= b,
    "contains": lambda a, b: a.astype(str).str.contains(str(b), case=False, na=False),
}

ACTIVITY_MATCH_SET = {"trade cnfm", "cxnc cnfm", "modf confirmation"}  # case-insensitive matching by substring


def clean_numeric_series(series: pd.Series) -> pd.Series:
    """Strip % and commas and coerce to numeric. Non-convertible become NaN."""
    s = series.astype(str).str.strip()
    s = s.str.replace("%", "", regex=False).str.replace(",", "", regex=False)
    s = s.replace({"": None})
    return pd.to_numeric(s, errors="coerce")


def choose(prompt: str, default: str | None = None) -> str:
    txt = input(f"{prompt}" + (f" [{default}]" if default else "") + ": ")
    return txt.strip() if txt.strip() else (default or "")


def load_excel(path: str) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path)
    else:
        df = pd.read_excel(path)
    return df


def list_unique_values(df: pd.DataFrame, col: str):
    if col not in df.columns:
        return []
    vals = df[col].dropna().astype(str).unique().tolist()
    vals_sorted = sorted(vals, key=lambda x: str(x).lower())
    return vals_sorted


def prompt_choose_users(df: pd.DataFrame, user_col: str = "USER"):
    users = list_unique_values(df, user_col)
    if not users:
        print(f"No values found in column '{user_col}'. Please check the column name.")
        return []

    print(f"\nFound {len(users)} users in column '{user_col}':")
    for i, u in enumerate(users):
        print(f"  {i+1}. {u}")

    sel = input("\nEnter comma-separated indices or user names to process (or 'all'): ").strip()
    if not sel:
        return []
    if sel.lower() == "all":
        return users
    chosen = []
    for part in [p.strip() for p in sel.split(",") if p.strip()]:
        if part.isdigit():
            idx = int(part) - 1
            if 0 <= idx < len(users):
                chosen.append(users[idx])
        else:
            # match user name case-insensitively
            matches = [u for u in users if u.lower() == part.lower()]
            if matches:
                chosen.append(matches[0])
            else:
                # try substring match
                subs = [u for u in users if part.lower() in u.lower()]
                if len(subs) == 1:
                    chosen.append(subs[0])
                elif len(subs) > 1:
                    print(f"Ambiguous selection '{part}'; matches: {subs}. Skipping.")
                else:
                    print(f"No match for '{part}'. Skipping.")
    # dedupe preserving order
    seen = set()
    out = []
    for x in chosen:
        if x not in seen:
            out.append(x)
            seen.add(x)
    return out


def prompt_json_config() -> (dict, str):
    """
    Prompt user to pick or create a JSON config file.
    Returns (config_dict, config_path)
    Schema:
      {
        "filters": [ {"col":"AUM (Cr)","op":">","value":"1000","value_type":"value"} , ... ],
        "filters_logic": "AND" | "OR",
        "sorts": [ {"col":"AUM (Cr)","ascending":false}, ... ],
        "columns": { "keep": ["Plan","Category Name","AuM (Cr)"], "reorder": ["Plan","AuM (Cr)","Category Name"] }
      }
    """
    while True:
        path = input("\nEnter path to JSON config file (existing) or a new name to create: ").strip()
        if not path:
            print("Please enter a path.")
            continue
        p = Path(path)
        if p.exists():
            try:
                data = json.loads(p.read_text(encoding="utf-8"))
                print(f"Loaded config from {p}")
                return data, str(p)
            except Exception as e:
                print("Failed to load JSON:", e)
                continue
        else:
            create = input(f"No file at {p}. Create new config here? (y/n) [y]: ").strip().lower() or "y"
            if create.startswith("y"):
                cfg = create_new_config_interactive()
                # save
                try:
                    p.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
                    print(f"Saved new config to {p}")
                except Exception as e:
                    print("Failed to save:", e)
                return cfg, str(p)
            else:
                continue


def create_new_config_interactive() -> dict:
    print("\n--- Create new JSON config interactively ---")
    cfg = {"filters": [], "filters_logic": "AND", "sorts": [], "columns": {"keep": [], "reorder": []}}

    logic = input("Combine filters with AND or OR? [AND]: ").strip().upper() or "AND"
    cfg["filters_logic"] = "AND" if logic not in ("AND", "OR") else logic

    # add filters
    while True:
        add = input("Add a filter? (y/n) [n]: ").strip().lower() or "n"
        if not add.startswith("y"):
            break
        col = input("  Column name for filter: ").strip()
        op = input("  Operator (==, !=, >, <, >=, <=, contains) [==]: ").strip() or "=="
        val_type = input("  Value type: 'value' or 'column_average' [value]: ").strip().lower() or "value"
        if val_type == "column_average":
            value = ""  # not used
        else:
            value = input("  Value (for contains, use a substring): ").strip()
        cfg["filters"].append({"col": col, "op": op, "value_type": val_type, "value": value})

    # sorts
    while True:
        add = input("Add a sort (multi-level)? (y/n) [n]: ").strip().lower() or "n"
        if not add.startswith("y"):
            break
        col = input("  Column name to sort by: ").strip()
        asc = input("  Ascending? (y/n) [y]: ").strip().lower() or "y"
        cfg["sorts"].append({"col": col, "ascending": True if asc.startswith("y") else False})

    # columns keep/order
    print("Optionally specify columns to KEEP (leave empty to keep all).")
    cols = input("Enter comma-separated column names to keep (or leave blank): ").strip()
    if cols:
        keep = [c.strip() for c in cols.split(",") if c.strip()]
        cfg["columns"]["keep"] = keep
        # reorder optional
        ro = input("Optional: specify final column order as comma-separated list (or leave blank): ").strip()
        if ro:
            cfg["columns"]["reorder"] = [c.strip() for c in ro.split(",") if c.strip()]

    print("Created config:", json.dumps(cfg, indent=2))
    return cfg


def apply_config_filters(df: pd.DataFrame, filters: list, logic: str = "AND", base_df: pd.DataFrame | None = None) -> pd.DataFrame:
    """
    filters: list of dicts: {"col":..., "op":..., "value_type": "value"|"column_average", "value":...}
    logic: "AND" or "OR"
    """
    if df is None or df.empty or not filters:
        return df
    if base_df is None:
        base_df = df

    masks = []
    for f in filters:
        col = f.get("col")
        op = f.get("op", "==")
        val_type = f.get("value_type", "value")
        val = f.get("value", "")

        if (col not in df.columns) or (op not in OPS and op != "contains"):
            # skip invalid
            continue
        try:
            s_raw = df[col]
            if val_type == "column_average":
                # compute avg from base_df column (numeric)
                if col not in base_df.columns:
                    masks.append(pd.Series([False] * len(df), index=df.index))
                    continue
                avg = clean_numeric_series(base_df[col]).mean(skipna=True)
                if pd.isna(avg):
                    masks.append(pd.Series([False] * len(df), index=df.index))
                else:
                    s_num = clean_numeric_series(s_raw)
                    mask = OPS[op](s_num, avg)
                    masks.append(mask.fillna(False))
                continue

            # value type
            if op == "contains":
                mask = OPS[op](s_raw.astype(str), val)
                masks.append(mask.fillna(False))
                continue

            # try numeric compare
            vnum = None
            try:
                vnum = float(str(val).replace("%", "").replace(",", ""))
            except Exception:
                vnum = None

            if vnum is not None:
                s_num = clean_numeric_series(s_raw)
                mask = OPS[op](s_num, vnum)
                masks.append(mask.fillna(False))
            else:
                # string equality/inequality
                if op in ("==", "!="):
                    mask = OPS[op](s_raw.astype(str), str(val))
                    masks.append(mask.fillna(False))
                else:
                    # incompatible op on strings -> skip
                    masks.append(pd.Series([False] * len(df), index=df.index))
        except Exception as e:
            masks.append(pd.Series([False] * len(df), index=df.index))

    if not masks:
        return df
    if logic.upper() == "AND":
        final = masks[0]
        for m in masks[1:]:
            final &= m
    else:
        final = masks[0]
        for m in masks[1:]:
            final |= m
    return df[final]


def apply_config_sorts(df: pd.DataFrame, sorts: list) -> pd.DataFrame:
    if df is None or df.empty or not sorts:
        return df
    cols = []
    asc = []
    for s in sorts:
        c = s.get("col")
        if c not in df.columns:
            continue
        cols.append(c)
        asc.append(bool(s.get("ascending", True)))
    if not cols:
        return df
    try:
        return df.sort_values(by=cols, ascending=asc, na_position="last")
    except Exception:
        return df


def apply_config_columns(df: pd.DataFrame, columns_cfg: dict) -> pd.DataFrame:
    """
    columns_cfg: {"keep": [cols] , "reorder": [cols]}.
    keep: if empty -> keep all columns
    reorder: optional final order (only columns present will be kept)
    """
    if df is None or df.empty:
        return df
    if not columns_cfg:
        return df
    keep = columns_cfg.get("keep", [])
    reorder = columns_cfg.get("reorder", [])
    if keep:
        keep_preserve = [c for c in df.columns if c in keep]
        df = df.loc[:, keep_preserve]
    if reorder:
        # only keep columns that exist now
        newcols = [c for c in reorder if c in df.columns]
        if newcols:
            df = df.loc[:, newcols]
    return df


def sanitize_sheet_name(name: str) -> str:
    # Excel sheet name max length 31. Also remove forbidden chars: : \ / ? * [ ]
    bad = set(r'[]:*?/\\')
    s = "".join(ch for ch in name if ch not in bad)
    return s[:31]


def process_for_users(df: pd.DataFrame, users: list, config: dict, user_col: str = "USER", activity_col: str = "Activity"):
    """
    Returns dict: {user_name: dataframe}
    """
    out = {}
    for user in users:
        sub = df[df[user_col].astype(str) == str(user)].copy()
        if sub.empty:
            print(f"[WARN] No rows for user {user}")
            out[user] = pd.DataFrame(columns=df.columns)
            continue

        # Activity filter: keep rows where activity contains any of the target strings (case-insensitive)
        act_series = sub[activity_col].astype(str).str.strip().str.lower() if activity_col in sub.columns else pd.Series([""] * len(sub), index=sub.index)
        mask_act = act_series.apply(lambda v: any(tok in v for tok in ACTIVITY_MATCH_SET))
        sub = sub[mask_act]

        # Apply config filters (common edits)
        sub = apply_config_filters(sub, config.get("filters", []), logic=config.get("filters_logic", "AND"), base_df=df)

        # Apply sorts
        sub = apply_config_sorts(sub, config.get("sorts", []))

        # Apply columns
        sub = apply_config_columns(sub, config.get("columns", {}))

        out[user] = sub
        print(f"[INFO] Processed user '{user}': {len(sub)} rows after filtering.")
    return out


def write_workbook(user_dfs: dict, dest_path: str):
    dest = Path(dest_path)
    try:
        with pd.ExcelWriter(dest, engine="openpyxl") as writer:
            for user, df in user_dfs.items():
                sheet = sanitize_sheet_name(str(user))
                df.to_excel(writer, sheet_name=sheet or "sheet", index=False)
        print(f"Saved workbook to {dest}")
    except Exception as e:
        print("Failed to save workbook:", e)


def main():
    print("=== ExcelOps: Batch filter by USER -> Activity -> Apply shared edits ===")

    # 1) get excel path
    default_path = DEFAULT_UPLOAD if Path(DEFAULT_UPLOAD).exists() else ""
    path = input(f"Excel file path [{default_path or 'enter path'}]: ").strip() or default_path
    if not path:
        print("No file provided. Exiting.")
        return

    try:
        df = load_excel(path)
    except Exception as e:
        print("Failed to load file:", e)
        return

    print(f"Loaded file with {len(df)} rows and {len(df.columns)} columns.")
    # show some columns for convenience
    print("Columns:", ", ".join(list(df.columns)[:50]))

    # 2) select users
    users = prompt_choose_users(df, user_col="USER")
    if not users:
        print("No users selected. Exiting.")
        return
    print("Selected users:", users)

    # 3) pick or create JSON config
    config, cfg_path = prompt_json_config()
    print("Using config:", cfg_path)

    # 4) Process each user
    user_dfs = process_for_users(df, users, config, user_col="USER", activity_col="Activity")

    # 5) Ask output path and save workbook
    out_default = str(Path.cwd() / "excelops_users_output.xlsx")
    out_path = input(f"Output workbook path [{out_default}]: ").strip() or out_default
    write_workbook(user_dfs, out_path)

    print("Done.")


if __name__ == "__main__":
    main()
