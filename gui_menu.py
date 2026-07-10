import tkinter as tk
from tkinter import messagebox

from budget_calc import budget_maker
from college_calc import college_save
from food_calc import food_coster
from home_calc import house_affordability
from inherit_calc import inherit_money
from interest_calc import generate_interest
from loans_calc import loan_payoff
from net_worth_calc import net_worth
from retirement_goal_calc import retirement_goal
from retirement_calc import retirement_dur
from take_home_calc import take_home


def launch_calculator(calc_func, excelFileExists, title):
    messagebox.showinfo(
        title,
        "This calculator will open a new window, while you will then use to input further information."
    )
    result = calc_func(excelFileExists)
    return result


def create_gui():
    excelFileExists = False

    root = tk.Tk()
    root.title("Sam's Financial Planning")
    root.geometry("600x410")
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
        elif choice == "take_home":
            excelFileExists = launch_calculator(take_home, excelFileExists, "Take Home Pay Calculator")
        elif choice == "college":
            excelFileExists = launch_calculator(college_save, excelFileExists, "College Savings Calculator")
        elif choice == "new house":
            excelFileExists = launch_calculator(house_affordability, excelFileExists, "House Affordability Calculator")
        elif choice == "inheritance":
            excelFileExists = launch_calculator(inherit_money, excelFileExists, "Inheritance Calculator")
        elif choice == "food_cost":
            excelFileExists = launch_calculator(food_coster, excelFileExists, "Food Cost Calculator")
        elif choice == "net_worth":
            excelFileExists = launch_calculator(net_worth, excelFileExists, "Net Worth Calculator")

    tk.Button(button_frame, text="1. Investments", width=34, command=lambda: handle_choice("investments")).grid(row=0, column=0, sticky="w", pady=4)
    tk.Button(button_frame, text="2. Loan Payoff Time", width=34, command=lambda: handle_choice("loan")).grid(row=1, column=0, sticky="w", pady=4)
    tk.Button(button_frame, text="3. Retirement Goals Calculator", width=34, command=lambda: handle_choice("retirement")).grid(row=2, column=0, sticky="w", pady=4)
    tk.Button(button_frame, text="4. Budget Calculator", width=34, command=lambda: handle_choice("budget")).grid(row=3, column=0, sticky="w", pady=4)
    tk.Button(button_frame, text="5. Retirement Duration Calculator", width=34, command=lambda : handle_choice("duration")).grid(row=4, column=0, sticky="w", pady=4)
    tk.Button(button_frame, text="6. Take Home Pay Calculator", width=34, command=lambda : handle_choice("take_home")).grid(row=5, column=0, sticky="w", pady=4)
    tk.Button(button_frame, text="7. Child Savings Calculator", width=34, command=lambda : handle_choice("college")).grid(row=0, column=1, sticky="w", pady=4)
    tk.Button(button_frame, text="8. Home Affordability Calculator", width=34, command=lambda : handle_choice("new house")).grid(row=1, column=1, sticky="w", pady=4)
    tk.Button(button_frame, text="9. Inheritance Calculator", width=34, command=lambda : handle_choice("inheritance")).grid(row=2, column=1, sticky="w", pady=4)
    tk.Button(button_frame, text="10. Food Cost Calculator", width=34, command=lambda : handle_choice("food_cost")).grid(row=3, column=1, sticky="w", pady=4)
    tk.Button(button_frame, text="11. Net Worth Calculator", width=34, command=lambda : handle_choice("net_worth")).grid(row=4, column=1, sticky="w", pady=4)

    tk.Button(root, text="Quit", width=16, command=root.destroy).pack(pady=16)

    root.mainloop()


def main():
    create_gui()


if __name__ == "__main__":
    main()
