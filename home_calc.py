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
        window.geometry("400x400")
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


        button_frame = tk.Frame(window)
        button_frame.pack(pady=16)

        tk.Button(button_frame, text="Quit", width=16, command=window.destroy).grid(row=0, column=0, padx=6)

        window.grab_set()
        window.mainloop()


def house_affordability(file_exists):
    calculator = HomeAffordabilityCalculator(file_exists)
    calculator.create_gui()
    return calculator.file_exists