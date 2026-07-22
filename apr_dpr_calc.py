import tkinter as tk
from tkinter import messagebox

import openpyxl
from openpyxl import Workbook

class AppreciateDepreciateCalculator:
    def __init__(self, file_exists = False):
        self.file_exists = file_exists
        self.initial_value = 0.0
        self.is_growing = False
        self.apr_dpr_rate = 0.0
    
    def parse_float(self, raw_value):
        if raw_value is None:
            raise ValueError("Missing value")
        if isinstance(raw_value, (int, float)):
            return float(raw_value)
        raw_value = raw_value.strip().replace(",", "")
        if raw_value.endswith("%"):
            raw_value = raw_value[:-1]
        if raw_value.startswith("$"):
            raw_value = raw_value[1:]
        return float(raw_value)
    
    def create_gui(self):
        window = tk.Toplevel()
        self.window = window
        window.title("Appreciation/Depreciation Calculator")
        window.geometry("460x520")
        window.resizable(False, False)

        header = tk.Label(
            window,
            text="Net Worth Calculator",
            font=("Segoe UI", 16, "bold"),
            wraplength=420,
            justify="center",
            pady=12,
        )
        header.pack()

        instructions = tk.Label(
            window,
            text="Type in an item name, the current value, whether or not it is growing in value, and the yearly rate of growth.",
            font=("Segoe UI", 10),
            wraplength=420,
            justify="center",
        )
        instructions.pack(pady=(0, 10))

        apr_frame = tk.Frame(window)
        apr_frame.pack(padx=20, pady=8, fill="x")

        button_frame = tk.Frame(window)
        button_frame.pack(pady=16)

        window.grab_set()
        window.mainloop()
    


def increase_decrease(file_exists):
    return file_exists