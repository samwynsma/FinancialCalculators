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

        inh_frame = tk.Frame(window)
        inh_frame.pack(padx=20, pady=8, fill = "x")

        tk.Label(inh_frame, text="Total Inheritance Amount ($):", anchor="w").grid(row=0, column=0, sticky="w", pady=6)
        self.total_inh_entry = tk.Entry(inh_frame, width=28)
        self.total_inh_entry.grid(row=0, column=1, pady=6)

        tk.Label(inh_frame, text="Estate Tax (%):", anchor="w").grid(row=1, column=0, sticky="w", pady=6)
        self.tax_entry = tk.Entry(inh_frame, width=28)
        self.tax_entry.grid(row=1, column=1, pady=6)

        tk.Label(inh_frame, text="Percent Inherited (%):", anchor="w").grid(row=2, column=0, sticky="w", pady=6)
        self.percent_inherit_entry = tk.Entry(inh_frame, width=28)
        self.percent_inherit_entry.grid(row=2, column=1, pady=6)

        tk.Label(inh_frame, text="Total Debt ($):", anchor="w").grid(row=3, column=0, sticky="w", pady=6)
        self.debt_entry = tk.Entry(inh_frame, width=28)
        self.debt_entry.grid(row=3, column=1, pady=6)

        tk.Label(inh_frame, text="Personal Spend (%):", anchor="w").grid(row=4, column=0, sticky="w", pady=6)
        self.personal_entry = tk.Entry(inh_frame, width=28)
        self.personal_entry.grid(row=4, column=1, pady=6)

        self.result_label = tk.Label(window, text="", font=("Segoe UI", 10), fg="green", wraplength=420, justify="center")
        self.result_label.pack(pady=(8, 0))

        button_frame = tk.Frame(window)
        button_frame.pack(pady=16)

        tk.Button(button_frame, text="Calculate", width=16, command=self.on_calculate).grid(row=0, column=0, padx=6)
        tk.Button(button_frame, text="Quit", width=16, command=window.destroy).grid(row=0, column=1, padx=6)

        window.grab_set()
        window.mainloop()
    
    def on_calculate(self):
        try:
            inherit = self.parse_float(self.total_inh_entry.get())
            tax = self.parse_float(self.tax_entry.get())
            percent = self.parse_float(self.percent_inherit_entry.get())
            debt = self.parse_float(self.debt_entry.get())
            personal = self.parse_float(self.personal_entry.get())
        except ValueError:
            messagebox.showerror("Input error", "Please enter valid numeric values for all fields.")
            return
        
        if(inherit <= 0.0):
            messagebox.showerror("Input error", "Inheritance must be positive")
            return
        
        if(tax < 0.0):
            messagebox.showerror("Input error", "Tax cannot be negative")
            return

        if(percent <= 0.0):
            messagebox.showerror("Input error", "Must receive a portion of the inheritance to be relevant.")
            return
        
        if(debt < 0.0):
            messagebox.showerror("Input error", "Debt cannot be negative")
            return
        
        if(personal < 0.0):
            messagebox.showerror("Input error", "Personal fun spend cannot be negative")
            return
        
        self.total_inheritance = inherit
        self.estate_tax = tax
        self.inherit_percent = percent
        self.total_debt = debt
        self.personal_spend_percent = personal

        return

def inherit_money(file_exists):
    calculator = InheritanceCalculator(file_exists)
    calculator.create_gui()
    return file_exists