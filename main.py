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
        self.configure(bg="#f4f6f8")

        self.df: pd.DataFrame | None = None
        self.datasets: dict[str, pd.DataFrame] = {}
        self.active_dataset_name: str | None = None
        self.preview_row_index_map: dict[str, int] = {}
        self.preview_deletable = False
        self._last_csv_sep: str | None = None
        self._last_csv_encoding: str | None = None
        self.lookup_path: str | None = None
        self.lookup_df: pd.DataFrame | None = None
        # sheets: list of dicts {name, tab, inner_nb, filters, sorts, columns, pivot}
        self.sheets = []
        self.plus_tab = None  # identifier for '+' tab
        self.preview_tab_id = None  # stores preview tab id if open

        self._apply_modern_theme()
        self._build_ui()

    def _apply_modern_theme(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        self.ui_bg = "#eef2f7"
        self.card_bg = "#ffffff"
        self.accent = "#6b7f2a"
        self.accent_dark = "#26311f"
        self.muted = "#64748b"
        self.configure(bg=self.ui_bg)

        style.configure("TFrame", background=self.ui_bg)
        style.configure("Card.TFrame", background=self.card_bg)
        style.configure("Toolbar.TFrame", background=self.card_bg)
        style.configure("Footer.TFrame", background=self.ui_bg)
        style.configure("TLabelframe", background=self.card_bg, borderwidth=1, relief="solid")
        style.configure(
            "TLabelframe.Label",
            background=self.card_bg,
            foreground="#1f2937",
            font=("Arial", 10, "bold"),
        )
        style.configure("TLabel", background=self.ui_bg, foreground="#111827", font=("Arial", 10))
        style.configure("Card.TLabel", background=self.card_bg, foreground="#111827", font=("Arial", 10))
        style.configure("Muted.TLabel", background=self.ui_bg, foreground=self.muted, font=("Arial", 9))
        style.configure(
            "Hero.TLabel",
            background=self.ui_bg,
            foreground=self.accent_dark,
            font=("Arial", 28, "bold"),
        )
        style.configure(
            "Subtitle.TLabel",
            background=self.ui_bg,
            foreground=self.muted,
            font=("Arial", 12),
        )
        style.configure(
            "Brand.TLabel",
            background=self.ui_bg,
            foreground=self.accent_dark,
            font=("Arial", 14, "bold"),
        )
        style.configure(
            "BrandCard.TLabel",
            background=self.card_bg,
            foreground=self.accent_dark,
            font=("Arial", 15, "bold"),
        )
        style.configure("TButton", padding=(12, 7), font=("Arial", 10), borderwidth=0)
        style.map("TButton", background=[("active", "#dbe7c2")])
        style.configure("Accent.TButton", padding=(16, 9), font=("Arial", 10, "bold"))
        style.configure("Hero.TButton", padding=(24, 16), font=("Arial", 13, "bold"))
        style.configure("Quiet.TButton", padding=(12, 7), font=("Arial", 10))
        style.configure("TCheckbutton", background=self.ui_bg, foreground="#111827")
        style.configure("Card.TCheckbutton", background=self.card_bg, foreground="#111827")
        style.configure("TCombobox", padding=5)
        style.configure("TNotebook", background=self.ui_bg, borderwidth=0)
        style.configure("TNotebook.Tab", padding=(14, 8), font=("Arial", 10, "bold"))
        style.map(
            "TNotebook.Tab",
            background=[("selected", self.card_bg), ("active", "#e6eddb")],
            foreground=[("selected", self.accent_dark), ("active", self.accent_dark)],
        )
        style.configure(
            "Treeview",
            rowheight=26,
            fieldbackground=self.card_bg,
            background=self.card_bg,
            foreground="#111827",
        )
        style.configure(
            "Treeview.Heading",
            font=("Arial", 10, "bold"),
            background="#e8eed9",
            foreground=self.accent_dark,
        )

    def _build_branding_footer(self, parent=None):
        parent = parent or self
        footer = ttk.Frame(parent, style="Footer.TFrame")
        footer.pack(side="bottom", fill="x", padx=14, pady=(0, 10))
        self.footer_frame = footer

        ttk.Label(
            footer,
            textvariable=getattr(self, "status_var", tk.StringVar(value="Ready")),
            style="Muted.TLabel",
        ).pack(side="left", anchor="w")

        right = ttk.Frame(footer, style="Footer.TFrame")
        right.pack(side="right", anchor="e")

        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "excelops_logo.png")
        self.logo_image = None
        if os.path.exists(logo_path):
            try:
                self.logo_image = tk.PhotoImage(file=logo_path)
                ttk.Label(right, image=self.logo_image).pack(side="right", padx=(6, 0))
            except Exception:
                self.logo_image = None

        if self.logo_image is None:
            self._draw_footer_logo(right).pack(side="right", padx=(10, 0))

        text_wrap = ttk.Frame(right, style="Footer.TFrame")
        text_wrap.pack(side="right", anchor="e")
        ttk.Label(text_wrap, text="ExcelOps", style="Brand.TLabel").pack(anchor="e")
        ttk.Label(
            text_wrap,
            text="powered by BDS",
            style="Muted.TLabel",
        ).pack(anchor="e")

    def _draw_footer_logo(self, parent):
        logo = tk.Canvas(
            parent,
            width=82,
            height=54,
            bg=self.ui_bg,
            highlightthickness=0,
            bd=0,
        )

        bar_color = "#b7c56b"
        leaf_color = self.accent_dark
        logo.create_rectangle(8, 31, 20, 44, fill=bar_color, outline=bar_color)
        logo.create_rectangle(28, 21, 40, 44, fill=bar_color, outline=bar_color)
        logo.create_rectangle(48, 9, 60, 44, fill=bar_color, outline=bar_color)

        leaf_shapes = [
            (54, 2, 48, 17, 54, 24, 60, 17),
            (42, 8, 34, 17, 42, 24, 48, 17),
            (66, 8, 60, 17, 66, 24, 76, 17),
            (30, 22, 18, 27, 30, 32, 42, 27),
            (68, 22, 56, 27, 68, 32, 80, 27),
        ]
        for points in leaf_shapes:
            logo.create_polygon(points, fill=leaf_color, outline=leaf_color, smooth=True)

        logo.create_text(42, 50, text="BDS", fill=leaf_color, font=("Arial", 13, "bold"))
        return logo

    def _hide_main_ui(self):
        for widget in getattr(self, "main_widgets", []):
            try:
                widget.pack_forget()
            except Exception:
                pass
        if getattr(self, "preview_only_frame", None) is not None:
            try:
                self.preview_only_frame.pack_forget()
            except Exception:
                pass

    def _show_main_ui(self):
        if getattr(self, "home_frame", None) is not None:
            self.home_frame.pack_forget()
        if getattr(self, "preview_only_frame", None) is not None:
            self.preview_only_frame.pack_forget()
        self.toolbar_frame.pack(side="top", fill="x", padx=12, pady=(10, 6), ipady=6)
        self.workspace_frame.pack(fill="both", expand=True, padx=12, pady=8)
        self.footer_frame.pack(side="bottom", fill="x", padx=14, pady=(0, 10))

    def _show_home_screen(self):
        self._hide_main_ui()
        if self.home_frame is None:
            self.home_frame = ttk.Frame(self, style="Footer.TFrame")

            hero = ttk.Frame(self.home_frame, style="Footer.TFrame")
            hero.pack(fill="both", expand=True, padx=42, pady=34)

            header = ttk.Frame(hero, style="Footer.TFrame")
            header.pack(fill="x", pady=(10, 28))
            ttk.Label(header, text="ExcelOps", style="Hero.TLabel").pack(anchor="center")
            ttk.Label(
                header,
                text="Build powerful spreadsheet workflows once. Run them in a few clicks.",
                style="Subtitle.TLabel",
            ).pack(anchor="center", pady=(8, 0))

            cards = ttk.Frame(hero, style="Footer.TFrame")
            cards.pack(anchor="center", fill="x", expand=True)
            cards.columnconfigure(0, weight=1)
            cards.columnconfigure(1, weight=1)

            create_card = ttk.LabelFrame(cards, text="Create Workflow")
            create_card.grid(row=0, column=0, sticky="nsew", padx=(0, 14), pady=10)
            ttk.Label(
                create_card,
                text=(
                    "Use the full ExcelOps editor to load files, create filters, "
                    "build pivots, add VLOOKUP steps, create calculations, and save the workflow."
                ),
                wraplength=430,
                justify="left",
            ).pack(anchor="w", padx=18, pady=(18, 12))
            ttk.Button(
                create_card,
                text="✨ Create or Edit Workflow",
                command=self._open_editor_mode,
                style="Hero.TButton",
            ).pack(anchor="w", padx=18, pady=(0, 18))

            run_card = ttk.LabelFrame(cards, text="Run Workflow")
            run_card.grid(row=0, column=1, sticky="nsew", padx=(14, 0), pady=10)
            ttk.Label(
                run_card,
                text=(
                    "Choose a saved workflow, select the required files, let ExcelOps "
                    "process everything, then preview and export the final workbook."
                ),
                wraplength=430,
                justify="left",
            ).pack(anchor="w", padx=18, pady=(18, 12))
            ttk.Button(
                run_card,
                text="▶ Run Saved Workflow",
                command=self._run_workflow_from_home,
                style="Hero.TButton",
            ).pack(anchor="w", padx=18, pady=(0, 18))

            tips = ttk.LabelFrame(hero, text="Simple workflow")
            tips.pack(fill="x", pady=(24, 0))
            ttk.Label(
                tips,
                text=(
                    "Create Workflow is for setup. Run Workflow is for day-to-day users. "
                    "After a workflow runs, open Preview to review the generated sheets before exporting."
                ),
                wraplength=980,
                justify="left",
            ).pack(anchor="w", padx=16, pady=14)

        self.home_frame.pack(fill="both", expand=True)
        self.status_var.set("Ready — select Create Workflow or Run Workflow")

    def _open_editor_mode(self):
        self._show_main_ui()
        self.status_var.set("Editor mode — build or update your workflow")

    def _run_workflow_from_home(self):
        self.status_var.set("Workflow runner — select a workflow and required files")
        self.after(100, lambda: self.run_preset_workflow(runner_mode=True))

    def _restore_home_after_runner_cancel(self):
        self.deiconify()
        self._show_home_screen()

    def _show_preview_only_screen(self):
        self.deiconify()
        self._hide_main_ui()
        if getattr(self, "home_frame", None) is not None:
            self.home_frame.pack_forget()

        if getattr(self, "preview_only_frame", None) is None:
            self.preview_only_frame = ttk.Frame(self, style="Footer.TFrame")

            header = ttk.Frame(self.preview_only_frame, style="Footer.TFrame")
            header.pack(fill="x", padx=24, pady=(18, 10))
            ttk.Label(header, text="Workflow Results", style="Hero.TLabel").pack(side="left")
            ttk.Button(header, text="🏠 Back Home", command=self._show_home_screen).pack(side="right", padx=6)
            ttk.Button(header, text="📤 Export Workbook", command=self.export_workbook, style="Accent.TButton").pack(side="right", padx=6)

            control = ttk.Frame(self.preview_only_frame, style="Footer.TFrame")
            control.pack(fill="x", padx=24, pady=(0, 10))
            ttk.Label(control, text="Preview output sheet:", style="Subtitle.TLabel").pack(side="left")
            self.runner_preview_selector = ttk.Combobox(control, state="readonly", width=42)
            self.runner_preview_selector.pack(side="left", padx=(8, 0))
            self.runner_preview_selector.bind("<<ComboboxSelected>>", lambda _event: self._update_runner_preview())

            tree_holder = ttk.Frame(self.preview_only_frame)
            tree_holder.pack(fill="both", expand=True, padx=24, pady=(0, 18))
            self.runner_preview_tree = ttk.Treeview(tree_holder, show="headings")
            self.runner_preview_tree.pack(side="left", fill="both", expand=True)
            vs = ttk.Scrollbar(tree_holder, orient="vertical", command=self.runner_preview_tree.yview)
            hs = ttk.Scrollbar(tree_holder, orient="horizontal", command=self.runner_preview_tree.xview)
            self.runner_preview_tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
            vs.pack(side="right", fill="y")
            hs.pack(side="bottom", fill="x")

        self.preview_only_frame.pack(fill="both", expand=True)
        self._refresh_runner_preview_selector()
        self.status_var.set("Workflow complete — previewing generated output")

    def _refresh_runner_preview_selector(self):
        if not hasattr(self, "runner_preview_selector"):
            return
        items = [s["name"] for s in self.sheets] or ["Raw Data"]
        self.runner_preview_selector["values"] = items
        cur = self.runner_preview_selector.get()
        self.runner_preview_selector.set(cur if cur in items else items[0])
        self._update_runner_preview()

    def _update_runner_preview(self):
        if not hasattr(self, "runner_preview_tree"):
            return
        target = self.runner_preview_selector.get()
        sheet = next((s for s in self.sheets if s["name"] == target), None)
        df = self._generate_filtered_df(sheet) if sheet else self.df
        self._render_df_into_tree(df, self.runner_preview_tree)

    def _build_ui(self):
        self.status_var = tk.StringVar(value="Ready — choose Create Workflow or Run Workflow")
        self._build_menu()
        self.home_frame = None
        self.main_widgets = []

        top = ttk.Frame(self, style="Toolbar.TFrame")
        top.pack(side="top", fill="x", padx=12, pady=(10, 6), ipady=6)
        self.toolbar_frame = top
        self.main_widgets.append(top)
        ttk.Label(top, text="ExcelOps", style="BrandCard.TLabel").pack(side="left", padx=(12, 16))
        ttk.Button(top, text="📂 Load", command=self.load_file, style="Accent.TButton").pack(side="left", padx=(0, 6))
        ttk.Button(top, text="👁 Preview", command=self.open_preview_tab).pack(side="left", padx=6)
        ttk.Button(top, text="💾 Save Workflow", command=lambda: PresetManager.save(self)).pack(side="left", padx=6)
        ttk.Button(top, text="▶ Run Workflow", command=self.run_preset_workflow, style="Accent.TButton").pack(side="left", padx=6)
        ttk.Button(top, text="📤 Export", command=self.export_workbook).pack(side="left", padx=(6, 14))

        ttk.Label(top, text="File:", style="Card.TLabel").pack(side="left", padx=(10, 0))
        self.dataset_selector = ttk.Combobox(top, state="readonly", width=30, values=[])
        self.dataset_selector.pack(side="left", padx=(6, 12))
        self.dataset_selector.bind("<<ComboboxSelected>>", self._on_dataset_selected)

        ttk.Label(top, text="Preview:", style="Card.TLabel").pack(side="left")
        self.preview_selector = ttk.Combobox(top, state="readonly", width=40, values=[])
        self.preview_selector.pack(side="left", padx=6)
        self.preview_selector.bind("<<ComboboxSelected>>", lambda e: self.update_preview())

        ttk.Button(top, text="🗑 Delete Rows", command=self.delete_selected_rows).pack(side="left", padx=6)

        self.show_all_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            top,
            text="Show full",
            variable=self.show_all_var,
            command=self.update_preview,
            style="Card.TCheckbutton",
        ).pack(side="right", padx=(0, 10))

        # vertical paned window: top preview tree removed (preview is in separate tab now)
        bottom = ttk.Frame(self, style="Card.TFrame")
        bottom.pack(fill="both", expand=True, padx=12, pady=8)
        self.workspace_frame = bottom
        self.main_widgets.append(bottom)

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
        self._build_branding_footer()
        self.main_widgets.append(self.footer_frame)
        self._show_home_screen()

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
        m.add_cascade(label="Workflows", menu=presets_m)
        presets_m.add_command(label="Save Workflow…", command=lambda: PresetManager.save(self))
        presets_m.add_command(label="Open Workflow…", command=lambda: PresetManager.load(self))
        presets_m.add_command(label="Run Workflow…", command=self.run_preset_workflow)
        presets_m.add_command(label="Manage Workflows…", command=lambda: PresetManager.manage(self))

        tools_m = tk.Menu(m, tearoff=False)
        m.add_cascade(label="Tools", menu=tools_m)
        tools_m.add_command(label="VLOOKUP…", command=self.apply_vlookup)

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
        load_many = messagebox.askyesno(
            "Load files",
            "Do you want to upload more than one file?\n\n"
            "Choose Yes to select multiple Excel/CSV files, or No to select one file.",
        )
        filetypes = [("Excel/CSV files", "*.xlsx *.xls *.csv"), ("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]
        if load_many:
            paths = filedialog.askopenfilenames(filetypes=filetypes)
        else:
            path = filedialog.askopenfilename(filetypes=filetypes)
            paths = (path,) if path else ()
        if not paths:
            return

        loaded: dict[str, pd.DataFrame] = {}
        errors = []
        for path in paths:
            try:
                df = self._read_data_file(path).reset_index(drop=True)
                loaded[self._unique_dataset_name(os.path.basename(path), loaded)] = df
            except Exception as e:
                errors.append(f"{os.path.basename(path)}: {e}")

        if not loaded:
            messagebox.showerror("Load error", "No files could be loaded.\n\n" + "\n".join(errors))
            return

        self.datasets = loaded
        self.active_dataset_name = next(iter(self.datasets))
        self.df = self.datasets[self.active_dataset_name]
        self._reset_workspace_for_dataset()
        self._refresh_dataset_selector()
        self.status_var.set(f"Loaded {len(self.datasets)} file(s) — ready to build workflow")

        loaded_msg = "\n".join(f"• {name}: {len(df)} rows, {len(df.columns)} columns" for name, df in self.datasets.items())
        if errors:
            loaded_msg += "\n\nSome files were skipped:\n" + "\n".join(errors)
        messagebox.showinfo("Loaded", f"Loaded {len(self.datasets)} file(s):\n\n{loaded_msg}")

    def _read_data_file(self, path: str) -> pd.DataFrame:
        if path.lower().endswith(".csv"):
            return self._read_csv_safely(path)
        return pd.read_excel(path)

    def _unique_dataset_name(self, name: str, loaded: dict[str, pd.DataFrame]) -> str:
        existing = set(self.datasets) | set(loaded)
        if name not in existing:
            return name
        stem, ext = os.path.splitext(name)
        i = 2
        while True:
            candidate = f"{stem} ({i}){ext}"
            if candidate not in existing:
                return candidate
            i += 1

    def _reset_workspace_for_dataset(self):
        for t in list(self.nb.tabs()):
            self.nb.forget(t)
        self.sheets.clear()
        self.plus_tab = None
        self.preview_tab_id = None
        self.preview_row_index_map = {}
        self.add_sheet("Sheet1")
        self._refresh_filters_after_data_change()
        self._ensure_plus_tab()
        self._refresh_preview_selector()

    def _refresh_dataset_selector(self):
        names = list(self.datasets)
        self.dataset_selector["values"] = names
        self.dataset_selector.set(self.active_dataset_name or "")

    def _on_dataset_selected(self, event=None):
        name = self.dataset_selector.get()
        if not name or name == self.active_dataset_name or name not in self.datasets:
            return
        if self.active_dataset_name and self.df is not None:
            self.datasets[self.active_dataset_name] = self.df.reset_index(drop=True)
        self.active_dataset_name = name
        self.df = self.datasets[name].reset_index(drop=True)
        self.datasets[name] = self.df
        self._refresh_filters_after_data_change()
        self._refresh_preview_selector()
        self.update_preview()

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
        vlookup_frame = self._build_vlookup_frame(inner_nb)
        vlookup_frame.set_columns(list(self.df.columns))
        if self.lookup_df is not None:
            vlookup_frame.set_lookup_source(self.lookup_path or "", list(self.lookup_df.columns))

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
            "vlookup_base_df": None,
            "final_output_df": None,
            "workflow_output_finalized": False,
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
            if "vlookup" in src and "vlookup" in dst:
                dst["vlookup"].load_config(src["vlookup"].get_config())
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

    def _build_vlookup_frame(self, parent):
        return VlookupFrame(parent, self.apply_vlookup, self.choose_lookup_file)

    def _prompt_lookup_file(self, title="Select lookup file (Excel or CSV)"):
        path = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")],
        )
        if not path:
            return None, None
        try:
            if path.lower().endswith(".csv"):
                lookup_df = self._read_csv_safely(path)
            else:
                lookup_df = pd.read_excel(path)
        except Exception as e:
            messagebox.showerror("Lookup load error", str(e))
            return None, None
        return path, lookup_df

    def choose_lookup_file(self):
        path, lookup_df = self._prompt_lookup_file()
        if lookup_df is None:
            return

        self.lookup_path = path
        self.lookup_df = lookup_df

        for s in self.sheets:
            vf = s.get("vlookup")
            if vf and hasattr(vf, "set_lookup_source"):
                vf.set_lookup_source(path, list(lookup_df.columns))

    def _load_lookup_file_for_path(self, path: str):
        if not path:
            return None
        try:
            if path.lower().endswith(".csv"):
                return self._read_csv_safely(path)
            return pd.read_excel(path)
        except Exception:
            return None

    def _run_vlookup_for_sheet(
        self,
        sheet,
        interactive=True,
        preset_override=None,
        prompt_for_file=False,
        record_history=True,
        main_df_override=None,
        update_sheet_state=True,
    ):
        preset_cfg = preset_override or {}
        if not preset_override and "vlookup" in sheet and hasattr(sheet["vlookup"], "get_config"):
            preset_cfg = sheet["vlookup"].get_config()
        preset_cfg = dict(preset_cfg or {})
        if "input_mode" not in preset_cfg and getattr(sheet.get("pivot"), "generated", False):
            # Backward compatibility for older saved VLOOKUP steps created after
            # generating a pivot but before input_mode existed.
            preset_cfg["input_mode"] = "pivot_result"

        lookup_path = self.lookup_path
        lookup_df = self.lookup_df
        preset_lookup_path = (preset_cfg or {}).get("lookup_file", "").strip()
        if prompt_for_file:
            label = os.path.basename(preset_lookup_path) if preset_lookup_path else "lookup file"
            lookup_path, lookup_df = self._prompt_lookup_file(f"Select lookup file for saved VLOOKUP ({label})")
            if lookup_df is None:
                return False
        elif preset_lookup_path:
            if lookup_path != preset_lookup_path:
                loaded = self._load_lookup_file_for_path(preset_lookup_path)
                if loaded is not None:
                    lookup_path = preset_lookup_path
                    lookup_df = loaded
            elif lookup_df is None:
                loaded = self._load_lookup_file_for_path(preset_lookup_path)
                if loaded is not None:
                    lookup_df = loaded

        if interactive and lookup_df is None:
            self.choose_lookup_file()
            lookup_path = self.lookup_path
            lookup_df = self.lookup_df
            if lookup_df is None:
                return False
            if "vlookup" in sheet and hasattr(sheet["vlookup"], "get_config"):
                preset_cfg = sheet["vlookup"].get_config()

        if lookup_df is not None and "vlookup" in sheet and hasattr(sheet["vlookup"], "set_lookup_source"):
            sheet["vlookup"].set_lookup_source(lookup_path or "", list(lookup_df.columns))
            if not preset_override:
                preset_cfg = sheet["vlookup"].get_config()
            else:
                preset_cfg = dict(preset_cfg)
                preset_cfg["lookup_file"] = lookup_path or preset_cfg.get("lookup_file", "")

        merged = perform_vlookup(
            self,
            sheet,
            preset=preset_cfg,
            lookup_df=lookup_df,
            lookup_path=lookup_path,
            interactive=interactive,
            main_df_override=main_df_override,
        )
        if merged is None:
            return False

        merged = merged.reset_index(drop=True)
        self.lookup_path = lookup_path
        self.lookup_df = lookup_df
        if not update_sheet_state:
            return merged
        if (preset_cfg or {}).get("input_mode") == "pivot_result":
            # Keep pivot -> VLOOKUP results at sheet level. The raw loaded file
            # (self.df / datasets) remains unchanged.
            sheet["final_output_df"] = merged
            sheet["workflow_output_finalized"] = False
        else:
            # Normal VLOOKUP augments this sheet's processed base data only; it
            # does not overwrite the raw loaded file.
            sheet["vlookup_base_df"] = merged
            sheet["final_output_df"] = None
            sheet["workflow_output_finalized"] = False
        if "columns" in sheet and hasattr(sheet["columns"], "refresh_source_df"):
            sheet["columns"].refresh_source_df(merged)
        if (preset_cfg or {}).get("input_mode") != "pivot_result" and "pivot" in sheet and hasattr(sheet["pivot"], "refresh_source_df"):
            sheet["pivot"].refresh_source_df(merged)
        if "vlookup" in sheet and hasattr(sheet["vlookup"], "set_columns"):
            sheet["vlookup"].set_columns(list(merged.columns))
        if record_history and "vlookup" in sheet and hasattr(sheet["vlookup"], "add_run_config"):
            sheet["vlookup"].add_run_config(preset_cfg)
        self.update_preview()
        return True

    def _generate_pre_workflow_base_df(self, sheet) -> pd.DataFrame:
        """Generate the starting table for an ordered sheet workflow.

        This intentionally applies filters and sorts only. Column/calculation
        edits are handled by the explicit ``columns`` workflow step so presets
        can represent workflows such as: pivot -> VLOOKUP1 -> VLOOKUP2 ->
        calculations/column editing.
        """
        df = self.df.copy()
        try:
            df = sheet["filters"].apply_filters(df)
        except Exception:
            pass
        try:
            df = sheet["sorts"].apply_sorts(df)
        except Exception:
            pass
        return df

    def _build_default_workflow_steps(self, sheet_cfg: dict) -> list[dict]:
        steps: list[dict] = []
        pivot_cfg = sheet_cfg.get("pivot", {}) or {}
        vlookup_cfg = sheet_cfg.get("vlookup", {}) or {}
        runs = list(vlookup_cfg.get("runs", []))
        if not runs and (vlookup_cfg.get("main_keys") or "").strip() and (vlookup_cfg.get("values") or "").strip():
            runs = [vlookup_cfg]

        if pivot_cfg.get("generated"):
            steps.append({"type": "pivot"})
        for run in runs:
            steps.append({"type": "vlookup", "config": dict(run or {})})
        if sheet_cfg.get("columns"):
            steps.append({"type": "columns"})
        return steps

    def _run_ordered_sheet_workflow(self, sheet, sheet_cfg: dict, prompt_for_files: bool = False) -> bool:
        """Replay a saved per-sheet workflow in the exact stored order."""
        workflow = list(sheet_cfg.get("workflow", [])) or self._build_default_workflow_steps(sheet_cfg)
        sheet["vlookup_base_df"] = None
        sheet["final_output_df"] = None
        sheet["workflow_output_finalized"] = False

        df = self._generate_pre_workflow_base_df(sheet)
        for step in workflow:
            step_type = step.get("type")
            if step_type == "pivot":
                try:
                    sheet["pivot"].load_config(sheet_cfg.get("pivot", {}))
                    sheet["pivot"].refresh_source_df(df)
                    pivot_df = sheet["pivot"]._build_pivot(df)
                    if pivot_df is not None:
                        df = pivot_df.reset_index(drop=True)
                        sheet["pivot"].generated = True
                except Exception as e:
                    messagebox.showerror("Workflow error", f"Pivot step failed on '{sheet['name']}':\n{e}")
                    return False
            elif step_type == "vlookup":
                run_cfg = dict(step.get("config") or {})
                merged = self._run_vlookup_for_sheet(
                    sheet,
                    interactive=False,
                    preset_override=run_cfg,
                    prompt_for_file=prompt_for_files,
                    record_history=False,
                    main_df_override=df,
                    update_sheet_state=False,
                )
                if merged is False or merged is None:
                    return False
                df = merged.reset_index(drop=True)
            elif step_type == "columns":
                try:
                    sheet["columns"].refresh_source_df(df)
                    df = sheet["columns"].apply_columns(df).reset_index(drop=True)
                except Exception as e:
                    messagebox.showerror("Workflow error", f"Columns/calculations step failed on '{sheet['name']}':\n{e}")
                    return False

        sheet["final_output_df"] = df.reset_index(drop=True)
        sheet["workflow_output_finalized"] = True
        if "columns" in sheet and hasattr(sheet["columns"], "refresh_source_df"):
            sheet["columns"].refresh_source_df(sheet["final_output_df"])
        if "vlookup" in sheet and hasattr(sheet["vlookup"], "set_columns"):
            sheet["vlookup"].set_columns(list(sheet["final_output_df"].columns))
        return True

    def apply_vlookup(self):
        if self.df is None:
            messagebox.showwarning("No data", "Load a file first.")
            return
        idx = self._active_sheet_index()
        if idx is None:
            messagebox.showwarning("No sheet", "Select a sheet to run VLOOKUP.")
            return
        sheet = self.sheets[idx]
        self._run_vlookup_for_sheet(sheet, interactive=True)

    def _apply_preset_config_to_workspace(self, cfg: dict, run_vlookups: bool = False):
        for tab_id in list(self.nb.tabs()):
            try:
                self.nb.forget(tab_id)
            except Exception:
                pass
        self.sheets.clear()
        self.plus_tab = None
        self.preview_tab_id = None

        for sheet_cfg in cfg.get("sheets", []):
            self.add_sheet(sheet_cfg.get("name", "Sheet"))
            sheet = self.sheets[-1]
            sheet["filters"].load_config(sheet_cfg.get("filters", {}))
            sheet["sorts"].load_config(sheet_cfg.get("sorts", {}))
            sheet["columns"].load_config(sheet_cfg.get("columns", {}))
            sheet["pivot"].load_config(sheet_cfg.get("pivot", {}))
            if "vlookup" in sheet:
                sheet["vlookup"].load_config(sheet_cfg.get("vlookup", {}))

            for key, refresh_name in (
                ("filters", "refresh_source_df"),
                ("columns", "refresh_source_df"),
            ):
                try:
                    frame = sheet[key]
                    if hasattr(frame, refresh_name):
                        getattr(frame, refresh_name)(self.df)
                except Exception:
                    pass
            try:
                sheet["pivot"].refresh_source_df(self._generate_base_df(sheet))
            except Exception:
                pass

            if run_vlookups:
                self._run_ordered_sheet_workflow(sheet, sheet_cfg, prompt_for_files=True)

        self._ensure_plus_tab()
        self._refresh_preview_selector()
        self.update_preview()

    def _safe_excel_sheet_name(self, name: str, used: set[str]) -> str:
        invalid = "[]:*?/\\"
        base = "".join("_" if ch in invalid else ch for ch in str(name or "Sheet")).strip() or "Sheet"
        base = base[:31]
        candidate = base
        i = 2
        while candidate in used:
            suffix = f"_{i}"
            candidate = f"{base[:31 - len(suffix)]}{suffix}"
            i += 1
        used.add(candidate)
        return candidate

    def _write_workbook(self, dest: str):
        with pd.ExcelWriter(dest, engine="openpyxl") as writer:
            used = set()
            for sheet in self.sheets:
                df = self._generate_filtered_df(sheet)
                sheet_name = self._safe_excel_sheet_name(sheet["name"], used)
                df.to_excel(writer, sheet_name=sheet_name, index=False)

    def run_preset_workflow(self, runner_mode: bool = False):
        if runner_mode:
            # Keep the full ExcelOps editor out of sight while the user selects
            # the workflow, main file, lookup files, and export location.
            self.withdraw()
        workflow_name = PresetManager.prompt_select_preset(self)
        if not workflow_name:
            if runner_mode:
                self._restore_home_after_runner_cancel()
            return
        try:
            preset_cfg = PresetManager.load_preset_data(workflow_name)
        except Exception as e:
            if runner_mode:
                self.deiconify()
            messagebox.showerror("Workflow error", str(e))
            if runner_mode:
                self._restore_home_after_runner_cancel()
            return

        main_path = filedialog.askopenfilename(
            title=f"Select the main file for workflow: {workflow_name}",
            filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv"), ("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")],
        )
        if not main_path:
            if runner_mode:
                self._restore_home_after_runner_cancel()
            return
        try:
            self.df = self._read_data_file(main_path).reset_index(drop=True)
        except Exception as e:
            if runner_mode:
                self.deiconify()
            messagebox.showerror("Load error", f"Could not load main file:\n{e}")
            if runner_mode:
                self._restore_home_after_runner_cancel()
            return

        self.datasets = {os.path.basename(main_path): self.df}
        self.active_dataset_name = os.path.basename(main_path)
        self._refresh_dataset_selector()

        self._apply_preset_config_to_workspace(preset_cfg, run_vlookups=True)
        self.status_var.set("Workflow processed — choose where to save the workbook")

        dest = filedialog.asksaveasfilename(
            title="Save processed workbook",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not dest:
            if runner_mode:
                self._restore_home_after_runner_cancel()
            return
        try:
            self._write_workbook(dest)
            if runner_mode:
                self._show_preview_only_screen()
            else:
                self.open_preview_tab()
                self.status_var.set("Workflow complete — preview is available and workbook was exported")
                messagebox.showinfo(
                    "Workflow complete",
                    f"Processed workflow '{workflow_name}' and saved workbook to:\n{dest}",
                )
        except Exception as e:
            if runner_mode:
                self.deiconify()
            messagebox.showerror("Workflow export error", str(e))
            if runner_mode:
                self._restore_home_after_runner_cancel()

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
        self.preview_tab_id = str(pf)

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
    def _apply_result_columns(self, sheet, df: pd.DataFrame) -> pd.DataFrame:
        try:
            return sheet["columns"].apply_columns(df)
        except Exception:
            return df

    def _generate_filtered_df(self, sheet) -> pd.DataFrame:
        if sheet.get("final_output_df") is not None:
            if sheet.get("workflow_output_finalized"):
                return sheet["final_output_df"].copy()
            return self._apply_result_columns(sheet, sheet["final_output_df"].copy())
        if sheet.get("vlookup_base_df") is not None:
            df = sheet["vlookup_base_df"].copy()
            try:
                df = sheet["pivot"].apply_pivot_if_requested(df)
            except Exception:
                pass
            return self._apply_result_columns(sheet, df)

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
            before_pivot = df
            df = sheet["pivot"].apply_pivot_if_requested(df)
            if df is not before_pivot:
                df = self._apply_result_columns(sheet, df)
        except Exception:
            pass
        return df

    def _generate_base_df(self, sheet) -> pd.DataFrame:
        if sheet.get("vlookup_base_df") is not None:
            return sheet["vlookup_base_df"].copy()
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
        idx = self._active_sheet_index()
        if idx is not None and idx < len(self.sheets):
            sheet = self.sheets[idx]
            if "vlookup" in sheet and hasattr(sheet["vlookup"], "set_columns"):
                try:
                    if "columns" in sheet and hasattr(sheet["columns"], "refresh_source_df"):
                        sheet["columns"].refresh_source_df(pivot_df)
                    sheet["vlookup"].set_columns(list(pivot_df.columns))
                    if hasattr(sheet["vlookup"], "use_pivot_result_input"):
                        sheet["vlookup"].use_pivot_result_input()
                except Exception:
                    pass
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
            self._write_workbook(dest)
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
        if self.active_dataset_name:
            self.datasets[self.active_dataset_name] = self.df
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
            try:
                if "vlookup" in s and hasattr(s["vlookup"], "set_columns"):
                    s["vlookup"].set_columns(list(self.df.columns))
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
