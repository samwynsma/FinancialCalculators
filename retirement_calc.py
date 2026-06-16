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
        return



def retirement_dur(file_exists):
    calculator = RetirementDurationCalculator(file_exists)
    calculator.create_gui()
    return calculator.file_exists