import tkinter as tk
from tkinter import messagebox

import openpyxl
from openpyxl import Workbook

class TakeHomeCalculator:
    def __init__(self, file_exists = False):
        self.pay_rate = -1.0
        self.percent_tax_rate = -1.0
        self.set_aside = -1.0
        self.pay_freq = "Hourly"
        self.period = "Month"
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
        window.title("Take Home Pay Calculator")
        window.geometry("460x460")
        window.resizable(False, False)

        header = tk.Label(
            window,
            text="Take Home Pay Calculator",
            font=("Segoe UI", 16, "bold"),
            wraplength=420,
            justify="center",
            pady=12,
        )
        header.pack()

        instructions = tk.Label(
            window,
            text="Enter your pay and tax information below to see what your take home will be.",
            font=("Segoe UI", 10),
            wraplength=420,
            justify="center",
        )
        instructions.pack(pady=(0, 10))

        inv_frame = tk.Frame(window)
        inv_frame.pack(padx=20, pady=8, fill = "x")

        tk.Label(inv_frame, text="Pay Rate ($):", anchor="w").grid(row=0, column=0, sticky="w", pady=6)
        self.pay_entry = tk.Entry(inv_frame, width=28)
        self.pay_entry.grid(row=0, column=1, pady=6)

        self.pay_freq_var = tk.StringVar(value="hourly")
        tk.Label(inv_frame, text="Pay Frequency:", anchor="w").grid(row=1, column=0, sticky="w", pady=6)
        pay_freq_frame = tk.Frame(inv_frame)
        pay_freq_frame.grid(row=1, column=1, columnspan=5, sticky="w", pady=6)
        tk.Radiobutton(pay_freq_frame, text="Year", variable=self.pay_freq_var, value="yearly").pack(side="left", padx=(0, 12))
        tk.Radiobutton(pay_freq_frame, text="Month", variable=self.pay_freq_var, value="monthly").pack(side="left", padx=(0, 12))
        tk.Radiobutton(pay_freq_frame, text="Week", variable=self.pay_freq_var, value="weekly").pack(side="left", padx=(0, 12))
        tk.Radiobutton(pay_freq_frame, text="Day", variable=self.pay_freq_var, value="daily").pack(side="left", padx=(0, 12))
        tk.Radiobutton(pay_freq_frame, text="Hour", variable=self.pay_freq_var, value="hourly").pack(side="left")

        tk.Label(inv_frame, text="Tax Rate (%):", anchor="w").grid(row=2, column=0, sticky="w", pady=6)
        self.tax_entry = tk.Entry(inv_frame, width=28)
        self.tax_entry.grid(row=2, column=1, pady=6)

        tk.Label(inv_frame, text="Set Aside (%):", anchor="w").grid(row=3, column=0, sticky="w", pady=6)
        self.aside_entry = tk.Entry(inv_frame, width=28)
        self.aside_entry.grid(row=3, column=1, pady=6)

        self.take_home_freq_var = tk.StringVar(value = "weekly")
        tk.Label(inv_frame, text="Take Home Frequency:", anchor="w").grid(row=4, column=0, sticky="w", pady=6)
        take_home_freq_frame = tk.Frame(inv_frame)
        take_home_freq_frame.grid(row=4, column=1, columnspan=5, sticky="w", pady=6)
        tk.Radiobutton(take_home_freq_frame, text="Year", variable=self.take_home_freq_var, value="yearly").pack(side="left", padx=(0, 12))
        tk.Radiobutton(take_home_freq_frame, text="Month", variable=self.take_home_freq_var, value="monthly").pack(side="left", padx=(0, 12))
        tk.Radiobutton(take_home_freq_frame, text="Week", variable=self.take_home_freq_var, value="weekly").pack(side="left", padx=(0, 12))
        tk.Radiobutton(take_home_freq_frame, text="Day", variable=self.take_home_freq_var, value="daily").pack(side="left", padx=(0, 12))
        tk.Radiobutton(take_home_freq_frame, text="Hour", variable=self.take_home_freq_var, value="hourly").pack(side="left")

        self.result_label = tk.Label(window, text="", font=("Segoe UI", 10), fg="green", wraplength=420, justify="center")
        self.result_label.pack(pady=(8, 0))

        button_frame = tk.Frame(window)
        button_frame.pack(pady=16)

        tk.Button(button_frame, text="Calculate", width=16, command=self.on_calculate).grid(row=0, column=0, padx=6)
        tk.Button(button_frame, text="Quit", width=16, command=window.destroy).grid(row=0, column=1, padx=6)

        window.grab_set()
        window.mainloop()
    
    def on_calculate(self):
        
        self.result_label.config(text="Results saved to InterestCalculation.xlsx")
        messagebox.showinfo("Take home document", "Take home calculation saved to InterestCalculation.xlsx.")

    

def take_home(file_exists):
    calculator = TakeHomeCalculator(file_exists)
    calculator.create_gui()
    return calculator.file_exists