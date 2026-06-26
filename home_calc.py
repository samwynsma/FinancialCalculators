import tkinter as tk
from tkinter import messagebox

import openpyxl
from openpyxl import Workbook

class HomeAffordabilityCalculator:
    def __init__(self, file_exists=False):
        self.down_payment = -1.0
        self.monthly_salary = -1.0
        self.interest_rate = -1.0
        self.additional_debt = -1.0
        self.mortgage_type = "30 year"
        self.file_exists = False

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
        window.title("Home Affordability Calculator")
        window.geometry("580x400")
        window.resizable(False, False)

        header = tk.Label(
            window,
            text="Home Affordability Calculator",
            font=("Segoe UI", 16, "bold"),
            wraplength=420,
            justify="center",
            pady=12,
        )
        header.pack()

        instructions = tk.Label(
            window,
            text="Enter your savings, down payment, interest, and monthly post-tax salary below, and see what kind of house you can afford.",
            font=("Segoe UI", 10),
            wraplength=420,
            justify="center",
        )
        instructions.pack(pady=(0, 10))

        home_frame = tk.Frame(window)
        home_frame.pack(padx=20, pady=8, fill="x")

        tk.Label(home_frame, text="Down Payment ($):", anchor="w").grid(row=0, column=0, sticky="w", pady=6)
        self.start_entry = tk.Entry(home_frame, width=28)
        self.start_entry.grid(row=0, column=1, pady=6)

        tk.Label(home_frame, text="Monthly Post-Tax Salary ($):", anchor="w").grid(row=1, column=0, sticky="w", pady=6)
        self.start_entry = tk.Entry(home_frame, width=28)
        self.start_entry.grid(row=1, column=1, pady=6)

        tk.Label(home_frame, text="Additional Debt ($):", anchor="w").grid(row=2, column=0, sticky="w", pady=6)
        self.start_entry = tk.Entry(home_frame, width=28)
        self.start_entry.grid(row=2, column=1, pady=6)

        tk.Label(home_frame, text="Interest Rate (%):", anchor="w").grid(row=3, column=0, sticky="w", pady=6)
        self.start_entry = tk.Entry(home_frame, width=28)
        self.start_entry.grid(row=3, column=1, pady=6)

        self.loan_length_var = tk.StringVar(value = "30 year")
        tk.Label(home_frame, text="Loan Length:", anchor="w").grid(row=4, column=0, sticky="w", pady=6)
        loan_length_frame = tk.Frame(home_frame)
        loan_length_frame.grid(row=4, column=1, columnspan=5, sticky="w", pady=6)
        tk.Radiobutton(loan_length_frame, text="30 year", variable=self.loan_length_var, value="30 year").pack(side="left", padx=(0, 12))
        tk.Radiobutton(loan_length_frame, text="25 year", variable=self.loan_length_var, value="25 year").pack(side="left", padx=(0, 12))
        tk.Radiobutton(loan_length_frame, text="20 year", variable=self.loan_length_var, value="20 year").pack(side="left", padx=(0, 12))
        tk.Radiobutton(loan_length_frame, text="15 year", variable=self.loan_length_var, value="15 year").pack(side="left", padx=(0, 12))
        tk.Radiobutton(loan_length_frame, text="10 year", variable=self.loan_length_var, value="10 year").pack(side="left")

        self.result_label = tk.Label(window, text="", font=("Segoe UI", 10), fg="green", wraplength=420, justify="center")
        self.result_label.pack(pady=(8, 0))
        
        button_frame = tk.Frame(window)
        button_frame.pack(pady=16)

        tk.Button(button_frame, text="Quit", width=16, command=window.destroy).grid(row=0, column=0, padx=6)

        window.grab_set()
        window.mainloop()


def house_affordability(file_exists):
    calculator = HomeAffordabilityCalculator(file_exists)
    calculator.create_gui()
    return calculator.file_exists