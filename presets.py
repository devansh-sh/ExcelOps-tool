# presets.py
import os
import json
import shutil
import sys
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd

APP_NAME = "ExcelOps"


def _user_preset_dir() -> str:
    """Return a writable workflow directory in source and packaged modes."""
    override = os.environ.get("EXCELOPS_PRESET_DIR")
    if override:
        return os.path.abspath(os.path.expanduser(override))
    if not getattr(sys, "frozen", False):
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "presets")
    if sys.platform == "darwin":
        base = os.path.expanduser("~/Library/Application Support")
    elif os.name == "nt":
        base = os.environ.get("APPDATA", os.path.expanduser("~"))
    else:
        base = os.environ.get("XDG_DATA_HOME", os.path.expanduser("~/.local/share"))
    return os.path.join(base, APP_NAME, "workflows")


PRESET_DIR = _user_preset_dir()


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
        if getattr(sys, "frozen", False):
            bundled = os.path.join(getattr(sys, "_MEIPASS", ""), "presets")
            if os.path.isdir(bundled):
                for filename in os.listdir(bundled):
                    if not filename.endswith(".json"):
                        continue
                    target = os.path.join(PRESET_DIR, filename)
                    if not os.path.exists(target):
                        shutil.copy2(os.path.join(bundled, filename), target)

    @staticmethod
    def _preset_path(name: str) -> str:
        safe_name = str(name).strip().replace("/", "_").replace("\\", "_")
        if not safe_name or safe_name in {".", ".."}:
            raise ValueError("Workflow name cannot be empty.")
        return os.path.join(PRESET_DIR, f"{safe_name}.json")

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
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)

    # ---------- UI actions ----------

    @staticmethod
    def _choose_preset_name(parent, title: str, presets: list[str]) -> str | None:
        win = tk.Toplevel(parent)
        win.title(title)
        win.geometry("500x500")
        win.minsize(420, 380)
        win.transient(parent)
        win.grab_set()

        heading = ttk.Frame(win)
        heading.pack(fill="x", padx=18, pady=(18, 10))
        ttk.Label(heading, text="Select a workflow", style="Brand.TLabel").pack(anchor="w")
        ttk.Label(
            heading,
            text="Choose a saved process, then ExcelOps will ask for the files it needs.",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        search_var = tk.StringVar()
        search = ttk.Entry(win, textvariable=search_var)
        search.pack(fill="x", padx=18, pady=(0, 10), ipady=5)

        lb = tk.Listbox(win, exportselection=False)
        lb.pack(fill="both", expand=True, padx=18, pady=(0, 12))

        def refresh_list(*_args):
            query = search_var.get().strip().casefold()
            matches = [name for name in presets if query in name.casefold()]
            lb.delete(0, "end")
            for workflow_name in matches:
                lb.insert("end", workflow_name)
            if matches:
                lb.selection_set(0)

        search_var.trace_add("write", refresh_list)
        refresh_list()

        selected = {"value": None}

        def on_ok():
            sel = lb.curselection()
            if not sel:
                return
            selected["value"] = lb.get(sel[0])
            win.destroy()

        def on_cancel():
            win.destroy()

        def focus_list(_event=None):
            lb.focus_set()
            return "break"

        btns = ttk.Frame(win)
        btns.pack(fill="x", padx=18, pady=(0, 18))
        ttk.Button(btns, text="Cancel", command=on_cancel).pack(side="right", padx=4)
        ttk.Button(btns, text="Continue", command=on_ok, style="Accent.TButton").pack(side="right", padx=4)

        lb.bind("<Double-1>", lambda e: on_ok())
        lb.bind("<Return>", lambda e: on_ok())
        search.bind("<Down>", focus_list)
        search.focus_set()
        win.wait_window()
        return selected["value"]


    @staticmethod
    def save(app):
        """
        Save preset from CURRENT sheets
        """
        if not app.sheets:
            messagebox.showwarning("No Sheets", "Nothing to save.")
            return

        name = simpledialog.askstring("Save Workflow", "Workflow name:")
        if not name:
            return

        PresetManager._ensure_dir()
        path = PresetManager._preset_path(name)

        data = {
            "sheets": []
        }

        for s in app.sheets:
            pivot_cfg = s["pivot"].get_config()
            vlookup_cfg = s["vlookup"].get_config() if "vlookup" in s else {}
            vlookup_cfg = PresetManager._normalize_vlookup_config_for_sheet(s, pivot_cfg, vlookup_cfg)
            columns_cfg = s["columns"].get_config()
            sheet_cfg = {
                "name": s["name"],
                "filters": s["filters"].get_config(),
                "sorts": s["sorts"].get_config(),
                "columns": columns_cfg,
                "pivot": pivot_cfg,
                "vlookup": vlookup_cfg,
                "workflow": PresetManager._build_workflow_for_sheet(pivot_cfg, vlookup_cfg, columns_cfg),
            }
            data["sheets"].append(sheet_cfg)

        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)

        messagebox.showinfo("Workflow Saved", f"Workflow '{name}' saved successfully.")

    @staticmethod
    def _build_workflow_for_sheet(pivot_cfg: dict, vlookup_cfg: dict, columns_cfg: dict) -> list[dict]:
        """Save an explicit ordered workflow for each sheet.

        The current UI's supported ordered workflow is:
        pivot (when generated) -> saved VLOOKUP steps in run order ->
        calculations/column editing. This makes preset replay deterministic
        instead of reconstructing order from independent feature configs.
        """
        workflow = []
        if (pivot_cfg or {}).get("generated"):
            workflow.append({"type": "pivot"})

        runs = list((vlookup_cfg or {}).get("runs", []))
        if not runs:
            has_keys = bool(((vlookup_cfg or {}).get("main_keys", "") or "").strip())
            has_values = bool(((vlookup_cfg or {}).get("values", "") or "").strip())
            if has_keys and has_values:
                runs = [vlookup_cfg]
        for run in runs:
            workflow.append({"type": "vlookup", "config": dict(run or {})})

        if columns_cfg:
            workflow.append({"type": "columns"})
        return workflow

    @staticmethod
    def _normalize_vlookup_config_for_sheet(sheet, pivot_cfg: dict, vlookup_cfg: dict) -> dict:
        """Keep saved VLOOKUP steps aligned with the sheet workflow.

        If the current sheet already has a final pivot-result VLOOKUP output, the
        preset must replay the pivot before VLOOKUP. Older/manual runs can still
        carry ``input_mode=base`` from before this workflow option existed, so we
        correct those saved runs at preset-save time.
        """
        cfg = dict(vlookup_cfg or {})
        runs = [dict(run or {}) for run in cfg.get("runs", [])]
        has_final_pivot_vlookup = sheet.get("final_output_df") is not None and bool(pivot_cfg.get("generated"))
        if has_final_pivot_vlookup:
            cfg["input_mode"] = "pivot_result"
            for run in runs:
                run["input_mode"] = "pivot_result"
        else:
            cfg.setdefault("input_mode", "base")
            for run in runs:
                run.setdefault("input_mode", cfg["input_mode"])
        cfg["runs"] = runs
        return cfg

    @staticmethod
    def load(app):
        """
        Load preset into UI (creates sheets)
        """
        presets = PresetManager.list_presets()
        if not presets:
            messagebox.showinfo("No Workflows", "No saved workflows available.")
            return

        name = PresetManager._choose_preset_name(app, "Open Workflow", presets)
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
                if hasattr(app, "_generate_base_df"):
                    s["pivot"].refresh_source_df(app._generate_base_df(s))
                else:
                    s["pivot"].refresh_source_df(app.df)
            except Exception:
                pass

            # Re-run saved VLOOKUP steps from the preset. Always ask for the lookup
            # files so presets remain portable across machines and folders.
            try:
                vlookup_cfg = sheet_cfg.get("vlookup", {})
                runs = list(vlookup_cfg.get("runs", []))
                if not runs:
                    has_keys = bool((vlookup_cfg.get("main_keys", "") or "").strip())
                    has_values = bool((vlookup_cfg.get("values", "") or "").strip())
                    if has_keys and has_values:
                        runs = [vlookup_cfg]
                if runs and hasattr(app, "_run_vlookup_for_sheet"):
                    should_run = messagebox.askyesno(
                        "Run workflow VLOOKUP steps?",
                        (
                            f"Workflow sheet '{s['name']}' contains {len(runs)} saved VLOOKUP step(s).\n\n"
                            "Do you want to select the lookup file(s) and run them now?"
                        ),
                    )
                    if should_run:
                        if hasattr(app, "_run_ordered_sheet_workflow"):
                            app._run_ordered_sheet_workflow(s, sheet_cfg, prompt_for_files=True)
                        else:
                            for run_cfg in runs:
                                if not app._run_vlookup_for_sheet(
                                    s,
                                    interactive=False,
                                    preset_override=run_cfg,
                                    prompt_for_file=True,
                                    record_history=False,
                                ):
                                    break
            except Exception:
                pass

        app.update_preview()
        messagebox.showinfo("Workflow Loaded", f"Workflow '{name}' loaded.")

    @staticmethod
    def manage(app):
        """
        Simple preset manager (delete only)
        """
        presets = PresetManager.list_presets()
        if not presets:
            messagebox.showinfo("No Workflows", "No saved workflows available.")
            return

        win = tk.Toplevel(app)
        win.title("Manage Workflows")
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
            if messagebox.askyesno("Delete", f"Delete workflow '{name}'?"):
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
            messagebox.showerror("No Workflows", "No saved workflows available.")
            return None

        return PresetManager._choose_preset_name(parent, "Select Workflow", presets)

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
