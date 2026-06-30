import tkinter as tk
from tkinter import messagebox

import openpyxl
from openpyxl import Workbook

class InheritanceCalculator:
    def __init__(self, file_exists):
        self.total_inheritance = 0.0
        self.estate_tax = -1.0
        self.inherit_percent = -1.0
        self.total_debt = -1.0
        self.personal_spend_percent = -1.0
        self.file_exists = file_exists

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
        window.title("Inheritance Calculator")
        window.geometry("460x460")
        window.resizable(False, False)

        header = tk.Label(
            window,
            text="Inheritance Calculator",
            font=("Segoe UI", 16, "bold"),
            wraplength=420,
            justify="center",
            pady=12,
        )
        header.pack()

        instructions = tk.Label(
            window,
            text="Determine the Amount of Money that you will Receive from Inheritance.",
            font=("Segoe UI", 10),
            wraplength=420,
            justify="center",
        )
        instructions.pack(pady=(0, 10))

        window.grab_set()
        window.mainloop()

def inherit_money(file_exists):
    calculator = InheritanceCalculator(file_exists)
    calculator.create_gui()
    return file_exists