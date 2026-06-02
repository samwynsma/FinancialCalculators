from budget_calc import budget_maker
from interest_calc import generate_interest
from loans_calc import loan_payoff
from retirement_goal_calc import retirement_goal

def main():
    menu_choice = ""
    excelFileExists = False
    while menu_choice != "quit" and menu_choice != "q":
        menu_choice = input("Welcome to Sam's financial planning application. Choose one of the options below. To close the application, type \"quit\" \n 1. Investments \n 2. Loan Payoff Time \n 3. Retirement Goals Calculator \n 4. Budget Calculator ").lower()
        if(menu_choice == "1" or menu_choice == "investments"):
            excelFileExists = generate_interest(excelFileExists)
        elif(menu_choice == "2" or menu_choice == "loan payoff time"):
            excelFileExists = loan_payoff(excelFileExists)
        elif(menu_choice == "3" or menu_choice == "retirement goals"):
            excelFileExists = retirement_goal(excelFileExists)
        elif(menu_choice == "4" or menu_choice == "budget maker"):
            excelFileExists = budget_maker(excelFileExists)
        

if __name__ == "__main__":
    main()