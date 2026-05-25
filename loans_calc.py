import xlsxwriter
import openpyxl
from openpyxl import Workbook

def loan_payoff(fileExists):
    print("Welcome to the loan payoff calculator! Today, I will take a starting loan amount, the interest of the loan, and the amount being paid each month, and I will calculate how long it will take for the loan to be paid off, showing how much loan will remain each month.")
    starting_loan = -1.0
    interest_rate = -1.0
    amount_paid = -1.0

    while starting_loan <= 0.0:
        num_input = input("How much is your initial loan amount? ")
        try:
            starting_loan = float(num_input)
            if(starting_loan <= 0.0):
                print("Invalid input: You cannot have a negative loan. If you have a negative loan, that's called an investment. Check the investment calculator for how to deal with those.")
        except:
            print("Invalid input: Loan must be a numeric value")
    
    while interest_rate <= 0.0:
        num_input = input("What is the annual interest rate? ")
        try:
            interest_rate = float(num_input)
            if(interest_rate <= 0.0):
                print("Invalid input: an interest rate cannot be negative for a loan. Any bank that would do that would go bankrupt.")
        except:
            print("Invalid input: try putting a number there")

    while amount_paid <= 0.0:
        num_input = input("How much do you plan to pay each month in loan payments? ")
        try:
            amount_paid = float(num_input)
            if(amount_paid <= 0.0):
                print("Invalid input: you need to spend money to pay off your loan")
            if(amount_paid < (starting_loan * (interest_rate / 1200.0))):
               print("If you pay that amount, you will never pay off the loan. Sorry, but you'll have to pay a larger amount.")
               amount_paid = -1.0
        except:
            print("Invalid input: try putting a number there")
    
    file = generate_loan_documents(starting_loan, interest_rate, amount_paid, fileExists)
    return file

def generate_loan_documents(loan, interest, payment, fileExists):
    loan_remaining = []
    loan_remaining.append(loan)

    interest = interest / 12.0
    current_money = loan

    while(current_money > 0.0 and len(loan_remaining) <= 600):
        current_money = round(current_money * (1 + interest / 100.0), 2) - payment
        if(current_money > 0.0):
            loan_remaining.append(current_money)
        else:
            loan_remaining.append(0.0)
    
    for i in range(len(loan_remaining)):
        print("Money at the beginning of month %d: %d" % (i, loan_remaining[i]))

    return fileExists