from pathlib import Path

import tkinter as tk
from tkinter import messagebox

is_help_mode = False

def help_files(category, index):
    helpdocs_path = Path(__file__).with_name("helpdocs.txt")
    help_key = f"{category}.{index}"

    with helpdocs_path.open(encoding="utf-8") as helpdocs:
        for line in helpdocs:
            key, separator, message = line.partition(":")
            if separator and key.strip() == help_key:
                messagebox.showinfo("Help", message.strip())
                return

    messagebox.showerror(
        "Help unavailable",
        f"No help message found for {help_key}.",
    )

def activate_help():
    global is_help_mode
    is_help_mode = not is_help_mode
    return

def get_help(category, index):
    if is_help_mode:
        help_files(category, index)