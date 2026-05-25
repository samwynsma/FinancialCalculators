import xlsxwriter

from interest_calc import generate_interest
from loans_calc import loan_payoff

def main():
    menu_choice = ""
    excelFileExists = False
    while menu_choice != "quit" and menu_choice != "q":
        menu_choice = input("Welcome to Sam's financial planning application. Choose one of the options below. To close the application, type \"quit\" \n 1. Investments \n 2. Loan Payoff Time: ").lower()
        if(menu_choice == "1" or menu_choice == "investments"):
            excelFileExists = generate_interest(excelFileExists)
        elif(menu_choice == "2" or menu_choice == "loan payoff time"):
            excelFileExists = loan_payoff(excelFileExists)
        

if __name__ == "__main__":
    main()