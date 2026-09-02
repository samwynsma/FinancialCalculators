import tkinter as tk
from tkinter import messagebox

from apr_dpr_calc import increase_decrease
from budget_calc import budget_maker
from college_calc import college_save
from display_info import display_info
from excel_doc import ExcelDocument
from food_calc import food_coster
from home_calc import house_affordability
from inherit_calc import inherit_money
from interest_calc import generate_interest
from loans_calc import loan_payoff
from net_worth_calc import net_worth
from retirement_goal_calc import retirement_goal
from retirement_calc import retirement_dur
from take_home_calc import take_home

document = ExcelDocument()

def launch_calculator(calc_func, document_obj, title):
    messagebox.showinfo(
        title,
        "This calculator will open a new window, while you will then use to input further information."
    )
    return calc_func(document_obj)


def create_gui():
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

    utility_frame = tk.Frame(root)
    utility_frame.pack(padx=20, anchor="center")

    def handle_choice(choice):
        if choice == "investments":
            launch_calculator(generate_interest, document, "Investments")
        elif choice == "loan":
            launch_calculator(loan_payoff, document, "Loan Payoff Time")
        elif choice == "retirement":
            launch_calculator(retirement_goal, document, "Retirement Goals")
        elif choice == "budget":
            launch_calculator(budget_maker, document, "Budget Calculator")
        elif choice == "duration":
            launch_calculator(retirement_dur, document, "Retirement Duration")
        elif choice == "take_home":
            launch_calculator(take_home, document, "Take Home Pay Calculator")
        elif choice == "college":
            launch_calculator(college_save, document, "College Savings Calculator")
        elif choice == "new house":
            launch_calculator(house_affordability, document, "House Affordability Calculator")
        elif choice == "inheritance":
            launch_calculator(inherit_money, document, "Inheritance Calculator")
        elif choice == "food_cost":
            launch_calculator(food_coster, document, "Food Cost Calculator")
        elif choice == "net_worth":
            launch_calculator(net_worth, document, "Net Worth Calculator")
        elif choice == "apr_dep":
            launch_calculator(increase_decrease, document, "Appreciation/Depreciation Calculator")
    
    def open_file():
        display_info(document)

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
    tk.Button(button_frame, text="12. Appreciation/Depreciation Calculator", width=34, command=lambda : handle_choice("apr_dep")).grid(row=5, column=1, sticky='w', pady=4)

    tk.Button(utility_frame, text="File", width=16, command=lambda : open_file()).grid(row=0, column=0, pady=4)
    tk.Button(utility_frame, text="Quit", width=16, command=root.destroy).grid(row=0, column=1, pady=4)

    root.mainloop()


def main():
    create_gui()


if __name__ == "__main__":
    main()
