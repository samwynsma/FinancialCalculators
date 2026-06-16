import tkinter as tk
from tkinter import messagebox

import openpyxl
from openpyxl import Workbook

class RetirementDurationCalculator:
    def __init__(self, file_exists = False):
        self.current_savings = -1.0
        self.growth_rate = -1.0
        self.social_security = -1.0
        self.pension = -1.0
        self.monthly_needs = -1.0
        self.expected_years_left = -1
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
    
    def parse_int(self, raw_value):
        if raw_value is None:
            raise ValueError("Missing value")
        if isinstance(raw_value, (int)):
            return int(raw_value)
        raw_value = raw_value.strip().replace(",", "")
        if raw_value.endswith("%"):
            raw_value = raw_value[:-1]
        if raw_value.startswith("$"):
            raw_value = raw_value[1:]
        return int(raw_value)
    
    def create_gui(self):
        window = tk.Toplevel()
        window.title("Retirement Duration Calculator")
        window.geometry("460x460")
        window.resizable(False, False)

        header = tk.Label(
            window,
            text="Retirement Duration Calculator",
            font=("Segoe UI", 16, "bold"),
            wraplength=420,
            justify="center",
            pady=12,
        )
        header.pack()

        instructions = tk.Label(
            window,
            text="Enter your current savings and information below to make a retirement plan.",
            font=("Segoe UI", 10),
            wraplength=420,
            justify="center",
        )
        instructions.pack(pady=(0, 10))

        inv_frame = tk.Frame(window)
        inv_frame.pack(padx=20, pady=8, fill = "x")

        tk.Label(inv_frame, text="Current Savings ($):", anchor="w").grid(row=0, column=0, sticky="w", pady=6)
        self.savings_entry = tk.Entry(inv_frame, width=28)
        self.savings_entry.grid(row=0, column=1, pady=6)

        tk.Label(inv_frame, text="Annual interest rate (%):", anchor="w").grid(row=1, column=0, sticky="w", pady=6)
        self.interest_rate_entry = tk.Entry(inv_frame, width=28)
        self.interest_rate_entry.grid(row=1, column=1, pady=6)

        tk.Label(inv_frame, text="Monthly Needs ($):", anchor="w").grid(row=2, column=0, sticky="w", pady=6)
        self.needs_entry = tk.Entry(inv_frame, width=28)
        self.needs_entry.grid(row=2, column=1, pady=6)

        tk.Label(inv_frame, text="Socialy Security ($):", anchor="w").grid(row=3, column=0, sticky="w", pady=6)
        self.age_entry = tk.Entry(inv_frame, width=28)
        self.age_entry.grid(row=3, column=1, pady=6)

        tk.Label(inv_frame, text="Pension ($):", anchor="w").grid(row=4, column=0, sticky="w", pady=6)
        self.retirement_entry = tk.Entry(inv_frame, width=28)
        self.retirement_entry.grid(row=4, column=1, pady=6)

        tk.Label(inv_frame, text="Expected Remaining Life Expectancy:", anchor="w").grid(row=5, column=0, sticky="w", pady=6)
        self.end_entry = tk.Entry(inv_frame, width=28)
        self.end_entry.grid(row=5, column=1, pady=6)

        button_frame = tk.Frame(window)
        button_frame.pack(pady=16)

        tk.Button(button_frame, text="Calculate", width=16, command=self.on_calculate).grid(row=0, column=0, padx=6)
        tk.Button(button_frame, text="Quit", width=16, command=window.destroy).grid(row=0, column=1, padx=6)

        window.grab_set()
        window.mainloop()
    
    def on_calculate(self):
        return



def retirement_dur(file_exists):
    calculator = RetirementDurationCalculator(file_exists)
    calculator.create_gui()
    return calculator.file_exists