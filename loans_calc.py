def loan_payoff():
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

    return