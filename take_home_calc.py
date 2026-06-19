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

    

def take_home(file_exists):
    return file_exists