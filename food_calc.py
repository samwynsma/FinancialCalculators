import tkinter as tk
from tkinter import messagebox

import openpyxl
from openpyxl import Workbook

class FoodCostCalculator:
    def __init__(self, file_exists):
        self.family_size = -1
        self.meals_per_day = -1
        self.snacks_per_day = -1
        self.current_food_spend = -1.0
        self.cost_per_meal = -1.0
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
        window.title("Food Cost Calculator")
        window.geometry("460x460")
        window.resizable(False, False)

        header = tk.Label(
            window,
            text="Food Cost Calculator",
            font=("Segoe UI", 16, "bold"),
            wraplength=420,
            justify="center",
            pady=12,
        )
        header.pack()

        instructions = tk.Label(
            window,
            text="Using the size of your family, the number of meals that you eat per person plus snacks, and the current food spend, I will determine how much you can save on a budget plan.",
            font=("Segoe UI", 10),
            wraplength=420,
            justify="center",
        )
        instructions.pack(pady=(0, 10))

        food_frame = tk.Frame(window)
        food_frame.pack(padx=20, pady=8, fill="x")

        tk.Label(food_frame, text="Number of people in your family:", anchor="w").grid(row=0, column=0, sticky="w", pady=6)
        self.family_entry = tk.Entry(food_frame, width=28)
        self.family_entry.grid(row=0, column=1, pady=6)

        tk.Label(food_frame, text="Average meals per day per person:", anchor="w").grid(row=1, column=0, sticky="w", pady=6)
        self.day_meals_entry = tk.Entry(food_frame, width=28)
        self.day_meals_entry.grid(row=1, column=1, pady=6)

        window.grab_set()
        window.mainloop()
    


def food_coster(file_exists):
    calculator = FoodCostCalculator(file_exists)
    calculator.create_gui()
    return file_exists