import tkinter as tk
from tkinter import messagebox

import openpyxl
from openpyxl import Workbook

class FoodCostCalculator:
    def __init__(self, file_exists):
        self.file_exists = file_exists
        self.family_size = -1
        self.meals_per_day = -1
        self.snacks_per_day = -1
        self.current_food_spend = -1.0
        self.cost_per_meal = -1.0
    
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
        window.title("Inheritance Calculator")
        window.geometry("460x460")
        window.resizable(False, False)
    


def food_coster(file_exists):
    return file_exists