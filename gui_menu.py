import tkinter as tk
from tkinter import messagebox

from budget_calc import budget_maker
from interest_calc import generate_interest
from loans_calc import loan_payoff
from retirement_goal_calc import retirement_goal
from retirement_calc import retirement_dur


def launch_calculator(calc_func, excelFileExists, title):
    messagebox.showinfo(
        title,
        "This calculator will open a new window, while you will then use to input further information."
    )
    result = calc_func(excelFileExists)
    messagebox.showinfo(
        title,
        "Finished creating the results. Check the console output and the generated Excel document."
    )
    return result


def create_gui():
    excelFileExists = False

    root = tk.Tk()
    root.title("Sam's Financial Planning")
    root.geometry("420x400")
    root.resizable(False, False)

    header = tk.Label(
        root,
        text="Welcome to Sam's financial planning application",
        font=("Segoe UI", 14, "bold"),
        wraplength=380,
        justify="center",
        pady=12,
    )
    header.pack()

    instructions = tk.Label(
        root,
        text="Select a calculator below. Detailed questions will appear in the console.",
        font=("Segoe UI", 10),
        wraplength=380,
        justify="center",
    )
    instructions.pack(pady=(0, 14))

    button_frame = tk.Frame(root)
    button_frame.pack(padx=20, fill="x")

    def handle_choice(choice):
        nonlocal excelFileExists
        if choice == "investments":
            excelFileExists = launch_calculator(generate_interest, excelFileExists, "Investments")
        elif choice == "loan":
            excelFileExists = launch_calculator(loan_payoff, excelFileExists, "Loan Payoff Time")
        elif choice == "retirement":
            excelFileExists = launch_calculator(retirement_goal, excelFileExists, "Retirement Goals")
        elif choice == "budget":
            excelFileExists = launch_calculator(budget_maker, excelFileExists, "Budget Calculator")
        elif choice == "duration":
            excelFileExists = launch_calculator(retirement_dur, excelFileExists, "Retirement Duration")

    tk.Button(button_frame, text="1. Investments", width=34, command=lambda: handle_choice("investments")).pack(pady=4)
    tk.Button(button_frame, text="2. Loan Payoff Time", width=34, command=lambda: handle_choice("loan")).pack(pady=4)
    tk.Button(button_frame, text="3. Retirement Goals Calculator", width=34, command=lambda: handle_choice("retirement")).pack(pady=4)
    tk.Button(button_frame, text="4. Budget Calculator", width=34, command=lambda: handle_choice("budget")).pack(pady=4)
    tk.Button(button_frame, text="5. Retirement Duration Calculator", width=34, command=lambda : handle_choice("duration")).pack(pady=4)

    tk.Button(root, text="Quit", width=16, command=root.destroy).pack(pady=16)

    root.mainloop()


def main():
    create_gui()


if __name__ == "__main__":
    main()
