import tkinter as tk
from tkinter import simpledialog, messagebox
import pandas as pd
import os

from presets import PresetManager


def run_automation(input_file: str):
    """
    Called by watcher when a new Excel file appears
    """

    root = tk.Tk()
    root.withdraw()  # hide main window

    # 1️⃣ Ask for user IDs
    user_input = simpledialog.askstring(
        "User Selection",
        "Enter User IDs (comma separated):\nExample: 2009-T44, 2010-A11"
    )

    if not user_input:
        messagebox.showinfo("Cancelled", "No users entered. Skipping file.")
        return

    users = [u.strip() for u in user_input.split(",") if u.strip()]
    if not users:
        messagebox.showerror("Error", "No valid User IDs entered.")
        return

    # 2️⃣ Ask for preset
    preset_name = PresetManager.prompt_select_preset(root)
    if not preset_name:
        messagebox.showinfo("Cancelled", "No preset selected.")
        return

    # 3️⃣ Load data
    df = pd.read_excel(input_file)

    # 4️⃣ Load preset config
    preset_cfg = PresetManager.load_preset_data(preset_name)

    output_path = os.path.splitext(input_file)[0] + "_OUTPUT.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for user in users:
            df_user = df.copy()

            # Apply preset filters
            df_user = PresetManager.apply_preset_to_df(
                df_user,
                preset_cfg,
                extra_filters={
                    "User": user
                }
            )

            if df_user.empty:
                continue

            sheet_name = user[:31]
            df_user.to_excel(writer, sheet_name=sheet_name, index=False)

    messagebox.showinfo(
        "Done",
        f"Output file created:\n{os.path.basename(output_path)}"
    )
