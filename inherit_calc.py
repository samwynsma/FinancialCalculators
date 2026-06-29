import tkinter as tk
from tkinter import messagebox

import openpyxl
from openpyxl import Workbook

class InheritanceCalculator:
    def __init__(self, file_exists):
        self.total_inheritance = 0.0
        self.inherit_percent = -1.0
        self.total_debt = -1.0
        self.personal_spend_percent = -1.0
        self.file_exists = file_exists

def inherit_money(file_exists):
    return file_exists