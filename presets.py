# presets.py
import os
import json
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd

PRESET_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "presets")


class PresetManager:
    """
    Handles saving, loading, managing presets for:
    - Filters
    - Sorts
    - Columns
    - Pivot
    """

    # ---------- filesystem helpers ----------

    @staticmethod
    def _ensure_dir():
        os.makedirs(PRESET_DIR, exist_ok=True)

    @staticmethod
    def _preset_path(name: str) -> str:
        return os.path.join(PRESET_DIR, f"{name}.json")

    @staticmethod
    def list_presets():
        PresetManager._ensure_dir()
        return sorted(
            f[:-5] for f in os.listdir(PRESET_DIR) if f.endswith(".json")
        )

    @staticmethod
    def _load_raw_preset(name: str) -> dict:
        path = PresetManager._preset_path(name)
        if not os.path.exists(path):
            raise FileNotFoundError(f"Preset not found: {name}")
        with open(path, "r") as f:
            return json.load(f)

    # ---------- UI actions ----------

    @staticmethod
    def save(app):
        """
        Save preset from CURRENT sheets
        """
        if not app.sheets:
            messagebox.showwarning("No Sheets", "Nothing to save.")
            return

        name = simpledialog.askstring("Save Preset", "Preset name:")
        if not name:
            return

        PresetManager._ensure_dir()
        path = PresetManager._preset_path(name)

        data = {
            "sheets": []
        }

        for s in app.sheets:
            sheet_cfg = {
                "name": s["name"],
                "filters": s["filters"].get_config(),
                "sorts": s["sorts"].get_config(),
                "columns": s["columns"].get_config(),
                "pivot": s["pivot"].get_config(),
                "vlookup": s["vlookup"].get_config() if "vlookup" in s else {},
            }
            data["sheets"].append(sheet_cfg)

        with open(path, "w") as f:
            json.dump(data, f, indent=2)

        messagebox.showinfo("Preset Saved", f"Preset '{name}' saved successfully.")

    @staticmethod
    def load(app):
        """
        Load preset into UI (creates sheets)
        """
        presets = PresetManager.list_presets()
        if not presets:
            messagebox.showinfo("No Presets", "No presets available.")
            return

        name = simpledialog.askstring(
            "Load Preset",
            "Available presets:\n\n" + "\n".join(presets) + "\n\nEnter preset name:"
        )
        if not name:
            return

        try:
            cfg = PresetManager._load_raw_preset(name)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        # Clear existing sheets
        for s in list(app.sheets):
            try:
                app.nb.forget(s["tab"])
            except Exception:
                pass
        app.sheets.clear()

        # Load preset sheets
        for sheet_cfg in cfg.get("sheets", []):
            app.add_sheet(sheet_cfg.get("name", "Sheet"))

            s = app.sheets[-1]
            s["filters"].load_config(sheet_cfg.get("filters", {}))
            s["sorts"].load_config(sheet_cfg.get("sorts", {}))
            s["columns"].load_config(sheet_cfg.get("columns", {}))
            s["pivot"].load_config(sheet_cfg.get("pivot", {}))
            if "vlookup" in s:
                s["vlookup"].load_config(sheet_cfg.get("vlookup", {}))

            # ensure df is wired so dropdowns populate
            try:
                s["filters"].refresh_source_df(app.df)
            except Exception:
                pass
            try:
                s["columns"].refresh_source_df(app.df)
            except Exception:
                pass
            try:
                s["pivot"].refresh_source_df(app.df)
            except Exception:
                pass

        app.update_preview()
        messagebox.showinfo("Preset Loaded", f"Preset '{name}' loaded.")

    @staticmethod
    def manage(app):
        """
        Simple preset manager (delete only)
        """
        presets = PresetManager.list_presets()
        if not presets:
            messagebox.showinfo("No Presets", "No presets available.")
            return

        win = tk.Toplevel(app)
        win.title("Manage Presets")
        win.geometry("300x300")

        lb = tk.Listbox(win)
        lb.pack(fill="both", expand=True, padx=10, pady=10)

        for p in presets:
            lb.insert("end", p)

        def delete():
            sel = lb.curselection()
            if not sel:
                return
            name = lb.get(sel[0])
            path = PresetManager._preset_path(name)
            if messagebox.askyesno("Delete", f"Delete preset '{name}'?"):
                os.remove(path)
                lb.delete(sel[0])

        ttk.Button(win, text="Delete Selected", command=delete).pack(pady=6)

    # ==========================================================
    # 🔽 AUTOMATION SUPPORT (used by watcher / CLI)
    # ==========================================================

    @staticmethod
    def prompt_select_preset(parent):
        presets = PresetManager.list_presets()
        if not presets:
            messagebox.showerror("No Presets", "No presets available.")
            return None

        return simpledialog.askstring(
            "Select Preset",
            "Available presets:\n\n" + "\n".join(presets) + "\n\nEnter preset name:",
            parent=parent
        )

    @staticmethod
    def load_preset_data(name: str) -> dict:
        return PresetManager._load_raw_preset(name)

    @staticmethod
    def apply_preset_to_df(df: pd.DataFrame, preset_cfg: dict, extra_filters: dict | None = None):
        """
        Applies a SINGLE-SHEET preset to a dataframe.
        Used by automation / watcher.
        """

        if not preset_cfg.get("sheets"):
            return df

        sheet_cfg = preset_cfg["sheets"][0]

        # Filters
        if "filters" in sheet_cfg:
            from filters import FiltersFrame
            f = FiltersFrame(None, None, df)
            f.load_config(sheet_cfg["filters"])
            df = f.apply_filters(df)

        # Extra filters (e.g. User)
        if extra_filters:
            for col, val in extra_filters.items():
                if col in df.columns:
                    df = df[df[col].astype(str) == str(val)]

        # Sorts
        if "sorts" in sheet_cfg:
            from sorts import SortsFrame
            s = SortsFrame(None, None, df)
            s.load_config(sheet_cfg["sorts"])
            df = s.apply_sorts(df)

        # Columns
        if "columns" in sheet_cfg:
            from columns_manager import ColumnsManagerFrame
            c = ColumnsManagerFrame(None, None, df)
            c.load_config(sheet_cfg["columns"])
            df = c.apply_columns(df)

        return df
