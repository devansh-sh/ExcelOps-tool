import tkinter as tk
from tkinter import ttk


class VlookupFrame(ttk.Frame):
    def __init__(self, parent, on_vlookup, on_vlookup_multi):
        super().__init__(parent)
        self.on_vlookup = on_vlookup
        self.on_vlookup_multi = on_vlookup_multi
        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 6, "pady": 6}
        ttk.Label(
            self,
            text=(
                "How to use VLOOKUP:\n"
                "1) Choose your main sheet and apply filters if needed.\n"
                "2) Click 'Run VLOOKUP' (single key) or 'Run Multi-Key VLOOKUP'.\n"
                "3) Select the lookup file when prompted.\n"
                "4) Pick key column(s) and the lookup value columns.\n"
                "5) Preview updates automatically in the Preview tab."
            ),
            wraplength=760,
            justify="left"
        ).pack(anchor="w", **pad)

        actions = ttk.Frame(self)
        actions.pack(fill="x", **pad)

        ttk.Button(actions, text="Run VLOOKUP", command=self.on_vlookup).pack(side="left", padx=4)
        ttk.Button(actions, text="Run Multi-Key VLOOKUP", command=self.on_vlookup_multi).pack(side="left", padx=4)
