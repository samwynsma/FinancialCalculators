import tkinter as tk
from tkinter import messagebox

is_help_mode = False

def help_files(category, index):
    messagebox.showinfo(
        "Help",
        "Placeholder Message: " + category + " and " + str(index)
    )
    return

def activate_help():
    is_help_mode = not is_help_mode
    return

def get_help(category, index):
    if is_help_mode:
        help_files(category, index)