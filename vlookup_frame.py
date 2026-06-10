import tkinter as tk
from tkinter import ttk


class VlookupFrame(ttk.Frame):
    def __init__(self, parent, on_vlookup, on_pick_lookup_file=None, columns=None):
        super().__init__(parent)
        self.on_vlookup = on_vlookup
        self.on_pick_lookup_file = on_pick_lookup_file
        self.columns = columns or []
        self.lookup_columns = []
        self.main_keys_var = tk.StringVar()
        self.lookup_keys_var = tk.StringVar()
        self.values_var = tk.StringVar()
        self.prefix_var = tk.StringVar()
        self.default_fill_var = tk.StringVar()
        self.lookup_file_var = tk.StringVar()
        self.same_keys_var = tk.BooleanVar(value=True)
        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 6, "pady": 6}
        ttk.Label(
            self,
            text=(
                "How to use VLOOKUP:\n"
                "1) Choose your main sheet and apply filters if needed.\n"
                "2) Choose a lookup file. Its columns will appear below.\n"
                "3) Pick the main key, lookup key, and lookup value column(s).\n"
                "4) Click Run VLOOKUP. Preview updates automatically."
            ),
            wraplength=820,
            justify="left"
        ).pack(anchor="w", **pad)

        form = ttk.LabelFrame(self, text="VLOOKUP Settings")
        form.pack(fill="x", **pad)
        for col in range(3):
            form.columnconfigure(col, weight=1)

        ttk.Label(form, text="1. Main key(s) from current file").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        self.main_keys_lb = tk.Listbox(form, selectmode="multiple", height=7, exportselection=False)
        self.main_keys_lb.grid(row=1, column=0, sticky="nsew", padx=6, pady=4)
        self.main_keys_lb.bind("<<ListboxSelect>>", lambda _event: self._on_main_keys_changed())

        ttk.Label(form, text="2. Lookup key(s) from lookup file").grid(row=0, column=1, sticky="w", padx=6, pady=4)
        lookup_key_panel = ttk.Frame(form)
        lookup_key_panel.grid(row=1, column=1, sticky="nsew", padx=6, pady=4)
        lookup_key_panel.columnconfigure(0, weight=1)

        ttk.Checkbutton(
            lookup_key_panel,
            text="Lookup keys have the same names as main keys",
            variable=self.same_keys_var,
            command=self._toggle_lookup_keys
        ).grid(row=0, column=0, sticky="w", pady=(0, 4))

        self.lookup_keys_lb = tk.Listbox(lookup_key_panel, selectmode="multiple", height=4, exportselection=False)
        self.lookup_keys_lb.grid(row=1, column=0, sticky="nsew")
        self.lookup_keys_lb.bind("<<ListboxSelect>>", lambda _event: self._sync_lookup_key_var())

        ttk.Label(lookup_key_panel, text="Manual lookup key(s), comma-separated").grid(row=2, column=0, sticky="w", pady=(6, 0))
        self.lookup_keys_entry = ttk.Entry(lookup_key_panel, textvariable=self.lookup_keys_var, width=34)
        self.lookup_keys_entry.grid(row=3, column=0, sticky="ew", pady=(2, 0))

        ttk.Label(form, text="3. Column(s) to bring from lookup file").grid(row=0, column=2, sticky="w", padx=6, pady=4)
        self.values_lb = tk.Listbox(form, selectmode="multiple", height=7, exportselection=False)
        self.values_lb.grid(row=1, column=2, sticky="nsew", padx=6, pady=4)

        options = ttk.Frame(form)
        options.grid(row=2, column=0, columnspan=3, sticky="ew", padx=6, pady=(6, 4))
        options.columnconfigure(1, weight=1)
        options.columnconfigure(3, weight=1)

        ttk.Label(options, text="Prefix for added columns (optional)").grid(row=0, column=0, sticky="w", padx=(0, 6))
        ttk.Entry(options, textvariable=self.prefix_var, width=28).grid(row=0, column=1, sticky="ew", padx=(0, 16))

        ttk.Label(options, text="Default fill for missing matches (optional)").grid(row=0, column=2, sticky="w", padx=(0, 6))
        ttk.Entry(options, textvariable=self.default_fill_var, width=28).grid(row=0, column=3, sticky="ew")

        lookup_picker = ttk.Frame(self)
        lookup_picker.pack(fill="x", **pad)
        ttk.Label(lookup_picker, text="Lookup file:").pack(side="left")
        ttk.Entry(lookup_picker, textvariable=self.lookup_file_var, state="readonly", width=70).pack(side="left", padx=(6, 6))
        ttk.Button(
            lookup_picker,
            text="Choose Lookup File",
            command=self.on_pick_lookup_file if callable(self.on_pick_lookup_file) else None,
        ).pack(side="left")

        actions = ttk.Frame(self)
        actions.pack(fill="x", **pad)

        ttk.Button(actions, text="Run VLOOKUP", command=self._run_vlookup_clicked).pack(side="left", padx=4)
        self._toggle_lookup_keys()

    def _run_vlookup_clicked(self):
        if not self.lookup_file_var.get() and callable(self.on_pick_lookup_file):
            self.on_pick_lookup_file()
            if not self.lookup_file_var.get():
                return
        if callable(self.on_vlookup):
            self.on_vlookup()

    def get_config(self):
        main_keys = self._selected(self.main_keys_lb)
        values = self._selected(self.values_lb)
        selected_lookup_keys = self._selected(self.lookup_keys_lb)
        manual_lookup_keys = self.lookup_keys_var.get().strip()
        lookup_keys = ""
        if not self.same_keys_var.get():
            lookup_keys = ", ".join(selected_lookup_keys) if selected_lookup_keys else manual_lookup_keys
        return {
            "main_keys": ", ".join(main_keys),
            "lookup_keys": lookup_keys,
            "values": ", ".join(values),
            "prefix": self.prefix_var.get(),
            "default_fill": self.default_fill_var.get(),
            "lookup_file": self.lookup_file_var.get(),
        }

    def load_config(self, cfg: dict):
        self.main_keys_var.set(cfg.get("main_keys", ""))
        self.lookup_keys_var.set(cfg.get("lookup_keys", ""))
        self.values_var.set(cfg.get("values", ""))
        self.prefix_var.set(cfg.get("prefix", ""))
        self.default_fill_var.set(cfg.get("default_fill", ""))
        self.lookup_file_var.set(cfg.get("lookup_file", ""))
        self.same_keys_var.set(cfg.get("lookup_keys", "") == "")
        self._apply_listbox_selections()
        self._toggle_lookup_keys()

    def set_columns(self, columns):
        self.columns = columns or []
        self._refresh_listbox(self.main_keys_lb, self.columns)
        self._apply_listbox_selections()

    def set_lookup_source(self, path: str, columns):
        self.lookup_file_var.set(path or "")
        self.lookup_columns = columns or []
        self._refresh_listbox(self.lookup_keys_lb, self.lookup_columns)
        self._refresh_listbox(self.values_lb, self.lookup_columns)
        self._auto_match_lookup_key()
        self._apply_listbox_selections()
        self._auto_select_lookup_value()
        self._toggle_lookup_keys()

    def _refresh_listbox(self, listbox, values):
        listbox.delete(0, "end")
        for item in values:
            listbox.insert("end", item)

    def _apply_listbox_selections(self):
        main_keys = [c.strip() for c in self.main_keys_var.get().split(",") if c.strip()]
        lookup_keys = [c.strip() for c in self.lookup_keys_var.get().split(",") if c.strip()]
        values = [c.strip() for c in self.values_var.get().split(",") if c.strip()]
        self._select_values(self.main_keys_lb, main_keys)
        self._select_values(self.lookup_keys_lb, lookup_keys)
        self._select_values(self.values_lb, values)

    def _selected(self, listbox):
        return [listbox.get(i) for i in listbox.curselection()]

    def _select_values(self, listbox, values):
        listbox.selection_clear(0, "end")
        value_lookup = {v.strip().lower() for v in values}
        for i in range(listbox.size()):
            if str(listbox.get(i)).strip().lower() in value_lookup:
                listbox.selection_set(i)

    def _on_main_keys_changed(self):
        self.main_keys_var.set(", ".join(self._selected(self.main_keys_lb)))
        self._auto_match_lookup_key()
        self._apply_listbox_selections()
        self._auto_select_lookup_value()
        self._toggle_lookup_keys()

    def _sync_lookup_key_var(self):
        selected = self._selected(self.lookup_keys_lb)
        if selected:
            self.lookup_keys_var.set(", ".join(selected))

    def _auto_match_lookup_key(self):
        if not self.lookup_columns:
            return
        selected_main_keys = self._selected(self.main_keys_lb)
        if not selected_main_keys:
            selected_main_keys = [c.strip() for c in self.main_keys_var.get().split(",") if c.strip()]
        if not selected_main_keys:
            return

        lookup_names = {str(c).strip().lower() for c in self.lookup_columns}
        all_main_keys_exist_in_lookup = all(str(k).strip().lower() in lookup_names for k in selected_main_keys)
        if all_main_keys_exist_in_lookup:
            if not self.lookup_keys_var.get().strip():
                self.same_keys_var.set(True)
            return

        # Common case: current file has Clientcode, lookup pivot uses Row Labels.
        # Switch to explicit lookup keys and pick the first lookup column so users can run without typing.
        if self.same_keys_var.get() or not self.lookup_keys_var.get().strip():
            self.same_keys_var.set(False)
            self.lookup_keys_var.set(str(self.lookup_columns[0]))

    def _auto_select_lookup_value(self):
        if self.values_lb.curselection() or self.values_var.get().strip():
            return
        selected_lookup_keys = set(self._selected(self.lookup_keys_lb))
        if not selected_lookup_keys and self.lookup_keys_var.get().strip():
            selected_lookup_keys = {c.strip() for c in self.lookup_keys_var.get().split(",") if c.strip()}
        for i in range(self.values_lb.size()):
            if self.values_lb.get(i) not in selected_lookup_keys:
                self.values_lb.selection_set(i)
                self.values_var.set(self.values_lb.get(i))
                break

    def _toggle_lookup_keys(self):
        state = "disabled" if self.same_keys_var.get() else "normal"
        self.lookup_keys_entry.configure(state=state)
        self.lookup_keys_lb.configure(state=state)
