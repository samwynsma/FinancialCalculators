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
    


def increase_decrease(file_exists):
    return file_exists