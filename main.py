# main.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import os
import sys

from vlookup_helper import perform_vlookup
from vlookup_frame import VlookupFrame

def is_batch_mode() -> bool:
    """
    Determines whether ExcelOps is launched in batch mode.
    Usage:
        python main.py --batch
    """
    return "--batch" in sys.argv


# Existing helper frames (must exist already)
from filters import FiltersFrame
from sorts import SortsFrame
from presets import PresetManager

# New merged columns + calculations manager and pivot
from columns_manager import ColumnsManagerFrame
from pivot import PivotFrame


class ExcelOpsApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ExcelOps")
        self.geometry("1280x820")

        self.df: pd.DataFrame | None = None
        self.preview_row_index_map: dict[str, int] = {}
        self.preview_deletable = False
        self._last_csv_sep: str | None = None
        self._last_csv_encoding: str | None = None
        # sheets: list of dicts {name, tab, inner_nb, filters, sorts, columns, pivot}
        self.sheets = []
        self.plus_tab = None  # identifier for '+' tab
        self.preview_tab_id = None  # stores preview tab id if open

        self._build_ui()

    def _build_ui(self):
        self._build_menu()

        top = ttk.Frame(self)
        top.pack(side="top", fill="x", padx=8, pady=(6, 0))
        ttk.Label(top, text="Preview:").pack(side="left")
        self.preview_selector = ttk.Combobox(top, state="readonly", width=40, values=[])
        self.preview_selector.pack(side="left", padx=6)
        self.preview_selector.bind("<<ComboboxSelected>>", lambda e: self.update_preview())

        ttk.Button(top, text="Delete Selected Rows", command=self.delete_selected_rows).pack(side="left", padx=6)

        self.show_all_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="Show full", variable=self.show_all_var, command=self.update_preview).pack(side="right")

        # vertical paned window: top preview tree removed (preview is in separate tab now)
        bottom = ttk.Frame(self)
        bottom.pack(fill="both", expand=True, padx=6, pady=6)

        # Notebook: sheet tabs (with '+' tab)
        self.nb = ttk.Notebook(bottom)
        self.nb.pack(fill="both", expand=True)
        self.nb.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        # right-click & double-click handlers
        self.nb.bind("<Button-3>", self._on_tab_right_click)
        self.nb.bind("<Double-1>", self._on_tab_double_click)
        self._make_tab_menu()
        self._ensure_plus_tab()

        # Preview Tree (kept hidden; created when preview tab is added)
        self._create_preview_tree_holder()

    def _build_menu(self):
        m = tk.Menu(self)
        self.config(menu=m)

        file_m = tk.Menu(m, tearoff=False)
        m.add_cascade(label="File", menu=file_m)
        file_m.add_command(label="Load File…", command=self.load_file)
        file_m.add_separator()
        file_m.add_command(label="Show Preview Tab", command=self.open_preview_tab)
        file_m.add_command(label="Export Current Sheet…", command=self.export_current_sheet)
        file_m.add_command(label="Export Workbook…", command=self.export_workbook)

        presets_m = tk.Menu(m, tearoff=False)
        m.add_cascade(label="Presets", menu=presets_m)
        presets_m.add_command(label="Save Preset…", command=lambda: PresetManager.save(self))
        presets_m.add_command(label="Load Preset…", command=lambda: PresetManager.load(self))
        presets_m.add_command(label="Manage Presets…", command=lambda: PresetManager.manage(self))

        tools_m = tk.Menu(m, tearoff=False)
        m.add_cascade(label="Tools", menu=tools_m)
        tools_m.add_command(label="VLOOKUP…", command=self.apply_vlookup)
        tools_m.add_command(label="VLOOKUP (Multi-Key)…", command=lambda: self.apply_vlookup(multi_key=True))

        edit_m = tk.Menu(m, tearoff=False)
        m.add_cascade(label="Edit", menu=edit_m)
        edit_m.add_command(label="Add Sheet", command=lambda: self.add_sheet(self._next_sheet_name()))
        edit_m.add_command(label="Remove Sheet", command=self.close_sheet)
        edit_m.add_separator()
        edit_m.add_command(label="Delete Selected Rows", command=self.delete_selected_rows)

    def _make_tab_menu(self):
        self.tab_menu = tk.Menu(self, tearoff=False)
        self.tab_menu.add_command(label="Rename…", command=self.rename_sheet)
        self.tab_menu.add_command(label="Duplicate", command=self.duplicate_sheet)
        self.tab_menu.add_command(label="Close", command=self.close_sheet)

    def _on_tab_right_click(self, event):
        try:
            idx = self._tab_index_at(event.x, event.y)
            if idx is None:
                return
            tab_id = self.nb.tabs()[idx]
            # block context menu if '+' tab or preview tab
            if self.plus_tab and tab_id == self.plus_tab:
                return
            if self.preview_tab_id and tab_id == self.preview_tab_id:
                return
            self.nb.select(idx)
            self.tab_menu.tk_popup(event.x_root, event.y_root)
        finally:
            try:
                self.tab_menu.grab_release()
            except Exception:
                pass

    def _on_tab_double_click(self, event):
        idx = self._tab_index_at(event.x, event.y)
        if idx is None:
            return
        tab_id = self.nb.tabs()[idx]
        if self.plus_tab and tab_id == self.plus_tab:
            return
        if self.preview_tab_id and tab_id == self.preview_tab_id:
            return
        self.nb.select(idx)
        self.rename_sheet(tab_index=idx)

    def _tab_index_at(self, x, y):
        try:
            return self.nb.index(f"@{x},{y}")
        except Exception:
            return None

    def _ensure_plus_tab(self):
        # remove existing plus if present
        if self.plus_tab and self.plus_tab in self.nb.tabs():
            try:
                self.nb.forget(self.plus_tab)
            except Exception:
                pass
            self.plus_tab = None
        # add at end
        plus_frame = ttk.Frame(self.nb)
        self.nb.add(plus_frame, text="  +  ")
        # store its tab identifier (string)
        self.plus_tab = self.nb.tabs()[-1]

    def _on_tab_changed(self, event=None):
        sel = self.nb.select()
        # if plus tab selected -> add a new sheet
        if self.plus_tab and sel == self.plus_tab:
            if self.df is None:
                messagebox.showwarning("No data", "Load a file first.")
                # move selection back (if any real tabs)
                tabs = [t for t in self.nb.tabs() if t != self.plus_tab]
                if tabs:
                    self.nb.select(tabs[0])
                return
            self.add_sheet(self._next_sheet_name())
            return
        # if preview tab selected -> do nothing (preview already rendered)
        # else -> update preview to currently selected sheet if preview tab open & set in selector
        self.update_preview()

    def _create_preview_tree_holder(self):
        # Create a reusable treeview for the preview tab (created hidden initially)
        self.preview_tree = ttk.Treeview(self, show="headings", selectmode="extended")
        self.preview_vs = ttk.Scrollbar(self, orient="vertical", command=self.preview_tree.yview)
        self.preview_hs = ttk.Scrollbar(self, orient="horizontal", command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=self.preview_vs.set, xscrollcommand=self.preview_hs.set)

    # ---------------- File ops ----------------
    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")])
        if not path:
            return
        try:
            if path.lower().endswith(".csv"):
                self.df = self._read_csv_safely(path)
            else:
                self.df = pd.read_excel(path)
            self.df = self.df.reset_index(drop=True)
        except Exception as e:
            messagebox.showerror("Load error", str(e))
            return

        # remove all tabs
        for t in list(self.nb.tabs()):
            self.nb.forget(t)
        self.sheets.clear()
        self.plus_tab = None
        self.preview_tab_id = None

        # add one starter sheet and a '+' tab
        self.add_sheet("Sheet1")

        # Ensure filters frames have the source DF for their dropdowns
        for s in self.sheets:
            try:
                if hasattr(s["filters"], "refresh_source_df"):
                    s["filters"].refresh_source_df(self.df)
            except Exception:
                pass

        self._refresh_filters_after_data_change()

        self._ensure_plus_tab()
        self._refresh_preview_selector()
        messagebox.showinfo("Loaded", f"Loaded {os.path.basename(path)} with {len(self.df)} rows, {len(self.df.columns)} columns.")

    def _read_csv_safely(self, path: str) -> pd.DataFrame:
        sample_lines = []
        encodings = ("utf-8-sig", "utf-8", "latin-1")
        for encoding in encodings:
            try:
                with open(path, "r", encoding=encoding, errors="replace") as handle:
                    for _ in range(30):
                        line = handle.readline()
                        if not line:
                            break
                        if line.strip():
                            sample_lines.append(line)
                if sample_lines:
                    break
            except Exception:
                continue
        if not sample_lines:
            return pd.read_csv(path)
        sep = self._detect_csv_delimiter(sample_lines)
        self._last_csv_sep = sep
        for encoding in encodings:
            try:
                df = pd.read_csv(
                    path,
                    sep=sep,
                    engine="python",
                    encoding=encoding,
                    index_col=False
                )
                self._last_csv_encoding = encoding
                break
            except Exception:
                df = None
        if df is None:
            df = pd.read_csv(path, index_col=False)
            self._last_csv_encoding = None
        if len(df.columns) == 1:
            df = self._retry_common_delimiters(path, encodings, df)
        return df

    def _retry_common_delimiters(
        self,
        path: str,
        encodings: tuple[str, ...],
        fallback_df: pd.DataFrame,
    ) -> pd.DataFrame:
        best_df = fallback_df
        best_cols = len(fallback_df.columns)
        for delim in (",", ";", "\t", "|"):
            for encoding in encodings:
                try:
                    candidate = pd.read_csv(
                        path,
                        sep=delim,
                        engine="python",
                        encoding=encoding,
                        index_col=False
                    )
                except Exception:
                    continue
                if len(candidate.columns) > best_cols:
                    best_df = candidate
                    best_cols = len(candidate.columns)
                    self._last_csv_sep = delim
                    self._last_csv_encoding = encoding
        return best_df

    def _detect_csv_delimiter(self, sample_lines: list[str]) -> str | None:
        import csv
        candidates = [",", ";", "\t", "|"]
        header_line = sample_lines[0] if sample_lines else ""
        header_counts = {d: header_line.count(d) for d in candidates}
        if header_counts:
            best_header = max(header_counts, key=header_counts.get)
            if header_counts[best_header] > 0:
                return best_header
        best = (None, 1, float("inf"))
        for delim in candidates:
            try:
                reader = csv.reader(sample_lines, delimiter=delim)
                counts = [len(row) for row in reader if row]
            except Exception:
                continue
            if not counts:
                continue
            counts.sort()
            median = counts[len(counts) // 2]
            variance = sum((c - median) ** 2 for c in counts) / len(counts)
            if median > best[1] or (median == best[1] and variance < best[2]):
                best = (delim, median, variance)
        return None if best[0] is None or best[1] <= 1 else best[0]


    def _read_csv_with_prompt(self, path: str, df: pd.DataFrame) -> pd.DataFrame:
        hint = str(df.columns[0])
        delimiter = simpledialog.askstring(
            "CSV Delimiter",
            "Auto-detected a single-column CSV.\n"
            "Enter the delimiter to use (examples: , ; \\t |):",
            initialvalue="," if "," in hint else ";" if ";" in hint else "\\t" if "\t" in hint else "|",
            parent=self
        )
        if delimiter is None:
            return df
        delimiter = delimiter.strip()
        if delimiter == "\\t":
            delimiter = "\t"
        if not delimiter:
            return df
        try:
            return pd.read_csv(path, sep=delimiter, engine="python")
        except Exception:
            return df

    # ---------------- Sheets ----------------
    def _next_sheet_name(self):
        base = "Sheet"
        existing = {s["name"] for s in self.sheets}
        i = 1
        while True:
            cand = f"{base}{i}"
            if cand not in existing:
                return cand
            i += 1

    def add_sheet(self, name="Sheet"):
        if self.df is None:
            messagebox.showwarning("No data", "Load a file first.")
            return
        tab = ttk.Frame(self.nb)
        # insert before plus tab if exists
        if self.plus_tab and self.plus_tab in self.nb.tabs():
            self.nb.insert(self.plus_tab, tab, text=name)
        else:
            self.nb.add(tab, text=name)

        # inner notebook for filter/sort/columns/pivot
        inner_nb = ttk.Notebook(tab)
        inner_nb.pack(fill="both", expand=True, padx=6, pady=6)

        filters_frame = FiltersFrame(inner_nb, self.on_sheet_change, self.df)
        # <-- important: immediately ensure filters_frame knows the current df so comboboxes populate
        try:
            filters_frame.refresh_source_df(self.df)
        except Exception:
            pass

        sorts_frame = SortsFrame(inner_nb, self.on_sheet_change, self.df)
        columns_frame = ColumnsManagerFrame(inner_nb, self.on_sheet_change, self.df)  # merged manager
        pivot_frame = PivotFrame(inner_nb, self.on_pivot_preview, self.df, data_provider=lambda: self._generate_base_df(sheet))
        vlookup_frame = VlookupFrame(inner_nb, self.apply_vlookup, lambda: self.apply_vlookup(multi_key=True))

        inner_nb.add(filters_frame, text="Filter")
        inner_nb.add(sorts_frame, text="Sort")
        inner_nb.add(columns_frame, text="Columns")
        inner_nb.add(pivot_frame, text="Pivot")
        inner_nb.add(vlookup_frame, text="VLOOKUP")

        sheet = {
            "name": name,
            "tab": tab,
            "inner_nb": inner_nb,
            "filters": filters_frame,
            "sorts": sorts_frame,
            "columns": columns_frame,
            "pivot": pivot_frame,
            "vlookup": vlookup_frame,
        }
        self.sheets.append(sheet)

        # recreate plus tab to remain at end
        self._ensure_plus_tab()
        # select the new tab
        self.nb.select(tab)
        self._refresh_preview_selector()
        # live update preview if preview tab open and selected in selector
        self.update_preview()

    def rename_sheet(self, tab_index: int | None = None):
        if tab_index is None:
            idx = self._active_sheet_index()
        else:
            # map tab_index -> sheet index (skip plus & preview)
            tab_list = [t for t in self.nb.tabs() if t != self.plus_tab and t != self.preview_tab_id]
            if tab_index < 0 or tab_index >= len(tab_list):
                return
            tab_id = tab_list[tab_index]
            idx = next((i for i, s in enumerate(self.sheets) if str(s["tab"]) == tab_id), None)
        if idx is None:
            return
        cur = self.sheets[idx]["name"]
        new = simpledialog.askstring("Rename Sheet", "New name:", initialvalue=cur, parent=self)
        if not new:
            return
        self.sheets[idx]["name"] = new
        try:
            self.nb.tab(self.sheets[idx]["tab"], text=new)
        except Exception:
            pass
        self._refresh_preview_selector()
        self.update_preview()

    def duplicate_sheet(self):
        idx = self._active_sheet_index()
        if idx is None:
            return
        src = self.sheets[idx]
        new_name = self._next_sheet_name()
        self.add_sheet(new_name)
        dst = self.sheets[-1]
        # copy configs if frames expose config functions
        try:
            dst["filters"].load_config(src["filters"].get_config())
            # after load, ensure the df is set on the dst filters so values populate
            if hasattr(dst["filters"], "refresh_source_df"):
                dst["filters"].refresh_source_df(self.df)
            dst["sorts"].load_config(src["sorts"].get_config())
            dst["columns"].load_config(src["columns"].get_config())
            dst["pivot"].load_config(src["pivot"].get_config())
        except Exception:
            pass
        self.update_preview()

    def close_sheet(self):
        idx = self._active_sheet_index()
        if idx is None:
            return
        try:
            self.nb.forget(self.sheets[idx]["tab"])
        except Exception:
            pass
        self.sheets.pop(idx)
        # ensure plus tab exists
        self._ensure_plus_tab()
        self._refresh_preview_selector()
        self.update_preview()

    def reset_active(self):
        idx = self._active_sheet_index()
        if idx is None:
            return
        s = self.sheets[idx]
        try:
            s["filters"].reset()
            s["sorts"].reset()
            s["columns"].reset()
            s["pivot"].reset()
        except Exception:
            pass
        self.update_preview()

    def _active_sheet_index(self):
        if not self.sheets:
            return None
        try:
            sel = self.nb.select()
            if self.plus_tab and sel == self.plus_tab:
                return None
            if self.preview_tab_id and sel == self.preview_tab_id:
                return None
            for i, s in enumerate(self.sheets):
                if str(s["tab"]) == sel:
                    return i
        except Exception:
            pass
        return None

    def apply_vlookup(self, multi_key: bool = False):
        if self.df is None:
            messagebox.showwarning("No data", "Load a file first.")
            return
        idx = self._active_sheet_index()
        if idx is None:
            messagebox.showwarning("No sheet", "Select a sheet to run VLOOKUP.")
            return
        sheet = self.sheets[idx]
        merged = perform_vlookup(self, sheet, multi_key=multi_key)
        if merged is None:
            return
        self.df = merged.reset_index(drop=True)
        self._refresh_filters_after_data_change()
        self.update_preview()

    # ---------------- Preview tab handling ----------------
    def open_preview_tab(self):
        # create preview tab next to '+' (at end). If exists, select it.
        if self.preview_tab_id and self.preview_tab_id in self.nb.tabs():
            self.nb.select(self.preview_tab_id)
            return
        # create frame and place before plus tab
        pf = ttk.Frame(self.nb)
        # place before plus tab if exists
        if self.plus_tab and self.plus_tab in self.nb.tabs():
            self.nb.insert(self.plus_tab, pf, text="Preview")
        else:
            self.nb.add(pf, text="Preview")
        self.preview_tab_id = self.nb.tabs()[-1]  # last added id

        # attach the reusable preview_tree and scrollbars into this frame
        tree_holder = ttk.Frame(pf)
        tree_holder.pack(fill="both", expand=True)
        self.preview_tree = ttk.Treeview(tree_holder, show="headings", selectmode="extended")
        self.preview_tree.pack(side="left", fill="both", expand=True)
        vs = ttk.Scrollbar(tree_holder, orient="vertical", command=self.preview_tree.yview)
        hs = ttk.Scrollbar(tree_holder, orient="horizontal", command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        vs.pack(side="right", fill="y")
        hs.pack(side="bottom", fill="x")

        # select this tab
        self.nb.select(pf)
        # update preview to current selection
        self.update_preview()

    def _refresh_preview_selector(self):
        items = ["Raw Data"]
        items += [s["name"] for s in self.sheets]
        cur = self.preview_selector.get()
        self.preview_selector["values"] = items
        if cur in items:
            self.preview_selector.set(cur)
        else:
            self.preview_selector.set(items[0] if items else "")

    def update_preview(self, raw: bool = False):
        # live updates: render into preview_tree if preview tab exists and selected; otherwise update preview_selector
        if self.df is None:
            # clear if nothing loaded
            if hasattr(self, "preview_tree"):
                try:
                    self.preview_tree.delete(*self.preview_tree.get_children())
                    self.preview_tree["columns"] = ()
                except Exception:
                    pass
            return

        # set preview selector values
        self._refresh_preview_selector()

        # if no preview tab open -> nothing to render
        if not (self.preview_tab_id and self.preview_tab_id in self.nb.tabs()):
            return

        # determine target dataset
        target = self.preview_selector.get()
        if raw or target == "Raw Data" or not self.sheets:
            df = self.df
        else:
            sheet = next((s for s in self.sheets if s["name"] == target), None)
            if not sheet:
                df = self.df
            else:
                df = self._generate_filtered_df(sheet)

        # render df into preview_tree (live)
        self._render_df_into_tree(df, self.preview_tree)

    def _render_df_into_tree(self, df, tree_widget):
        try:
            tree_widget.delete(*tree_widget.get_children())
        except Exception:
            pass
        self.preview_row_index_map = {}
        if df is None or df.empty:
            tree_widget["columns"] = ()
            self.preview_deletable = False
            return
        cols = list(df.columns)
        tree_widget["columns"] = cols
        tree_widget["show"] = "headings"
        for c in cols:
            tree_widget.heading(c, text=str(c))
            tree_widget.column(c, width=max(120, min(360, 10 * len(str(c)))), anchor="w")
        n = len(df) if self.show_all_var.get() else min(1000, len(df))
        for i, (_, row) in enumerate(df.head(n).iterrows()):
            vals = [("" if pd.isna(row[c]) else row[c]) for c in cols]
            iid = str(row.name)
            if iid in self.preview_row_index_map:
                iid = f"{iid}-{i}"
            self.preview_row_index_map[iid] = row.name
            tree_widget.insert("", "end", iid=iid, values=vals)
        self.preview_deletable = True

    # ---------------- Core df generation ----------------
    def _generate_filtered_df(self, sheet) -> pd.DataFrame:
        df = self.df.copy()
        try:
            df = sheet["filters"].apply_filters(df)
        except Exception:
            pass
        try:
            df = sheet["sorts"].apply_sorts(df)
        except Exception:
            pass
        try:
            df = sheet["columns"].apply_columns(df)
        except Exception:
            pass
        try:
            # pivot is preview-only unless user generates a pivot sheet explicitly
            df = sheet["pivot"].apply_pivot_if_requested(df)
        except Exception:
            pass
        return df

    def _generate_base_df(self, sheet) -> pd.DataFrame:
        df = self.df.copy()
        try:
            df = sheet["filters"].apply_filters(df)
        except Exception:
            pass
        try:
            df = sheet["sorts"].apply_sorts(df)
        except Exception:
            pass
        try:
            df = sheet["columns"].apply_columns(df)
        except Exception:
            pass
        return df

    def on_sheet_change(self):
        # called by inner frames to request live preview update
        self.update_preview()

    def on_pivot_preview(self, pivot_df):
        # called by pivot frame to show pivot result in preview
        if pivot_df is None or pivot_df.empty:
            messagebox.showinfo("Pivot", "No pivot result with current selections.")
            return
        # ensure preview tab exists and select it
        self.open_preview_tab()
        self._render_df_into_tree(pivot_df, self.preview_tree)
        self.preview_deletable = False

    # ---------------- Export ----------------
    def export_current_sheet(self):
        if self.df is None:
            messagebox.showwarning("Nothing to export", "Load a file first.")
            return
        name = self.preview_selector.get()
        if name == "Raw Data":
            data = self.df
            sheet_name = "Raw Data"
        else:
            sheet = next((s for s in self.sheets if s["name"] == name), None)
            if not sheet:
                messagebox.showwarning("No sheet", "Select a valid sheet to export.")
                return
            data = self._generate_filtered_df(sheet)
            sheet_name = sheet["name"]
        dest = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not dest:
            return
        try:
            # SINGLE sheet export (no Raw Data default unless requested)
            with pd.ExcelWriter(dest, engine="openpyxl") as w:
                data.to_excel(w, sheet_name=sheet_name[:31], index=False)
            messagebox.showinfo("Exported", f"Saved to {dest}")
        except Exception as e:
            messagebox.showerror("Export error", str(e))

    def export_workbook(self):
        if self.df is None or not self.sheets:
            messagebox.showwarning("Nothing to export", "Load a file and add at least one sheet.")
            return
        dest = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not dest:
            return
        try:
            with pd.ExcelWriter(dest, engine="openpyxl") as w:
                # don't export raw by default (you asked)
                for s in self.sheets:
                    df = self._generate_filtered_df(s)
                    df.to_excel(w, sheet_name=s["name"][:31], index=False)
            messagebox.showinfo("Exported", f"Workbook saved to {dest}")
        except Exception as e:
            messagebox.showerror("Export error", str(e))

    # ---------------- Row deletion ----------------
    def delete_selected_rows(self):
        if self.df is None:
            messagebox.showwarning("No data", "Load a file first.")
            return
        if not (self.preview_tab_id and self.preview_tab_id in self.nb.tabs()):
            messagebox.showwarning("Preview not open", "Open the Preview tab to select rows for deletion.")
            return
        if not self.preview_deletable:
            messagebox.showinfo("Unavailable", "Row deletion is only available in data previews.")
            return
        selected = self.preview_tree.selection()
        if not selected:
            messagebox.showinfo("No selection", "Select one or more rows in the preview table.")
            return
        indices = [self.preview_row_index_map.get(iid) for iid in selected]
        indices = [idx for idx in indices if idx is not None]
        if not indices:
            messagebox.showwarning("No rows", "Selected rows could not be resolved.")
            return
        if not messagebox.askyesno("Delete Rows", f"Delete {len(indices)} selected row(s) from the dataset?"):
            return
        self.df = self.df.drop(index=indices, errors="ignore")
        self.df = self.df.reset_index(drop=True)
        self._refresh_filters_after_data_change()
        self.update_preview()

    def _refresh_filters_after_data_change(self):
        for s in self.sheets:
            try:
                if hasattr(s["filters"], "refresh_source_df"):
                    s["filters"].refresh_source_df(self.df)
            except Exception:
                pass
            try:
                if hasattr(s["sorts"], "refresh_columns"):
                    s["sorts"].refresh_columns(self.df)
            except Exception:
                pass
            try:
                if hasattr(s["columns"], "refresh_source_df"):
                    s["columns"].refresh_source_df(self.df)
            except Exception:
                pass
            try:
                if hasattr(s["pivot"], "refresh_source_df"):
                    s["pivot"].refresh_source_df(self.df)
            except Exception:
                pass

def run_batch_mode():
    """
    Entry point for Activity Log automation.
    This will be expanded step-by-step.
    """
    print("🔁 ExcelOps running in BATCH MODE")

    # Placeholder – next steps will fill this
    print("✔ Batch mode initialized successfully")


if __name__ == "__main__":
    if is_batch_mode():
        run_batch_mode()
    else:
        app = ExcelOpsApp()
        app.mainloop()
