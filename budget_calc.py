def budget_maker(fileExists):
    print("Welcome to the budget calculator. I'm going to ask you some questions about your life to determine your budget.")
    print("Furthermore, I'm going to ask what your current spending is, so we can do a side by side comparison.")
    print("At the end, we will see how much leftover money there is, and I will give some recommendations.")
    people = 0
    pets = 0
    cars = 0
    mortgage_rent = 0.0
    auto_payments = 0.0
    debt_payments = 0.0
    utilities = 0.0
    insurance = 0.0
    food = 0.0
    extra_expenditures = 0.0
    medical = 0.0
    giving = 0.0
    investments = 0.0

    while(people <= 0):
        num_input = input("How many people are in your family? ")
        try:
            people = int(num_input)
            if people <= 0:
                print("Invalid input: you count as one of those people.")
        except:
            print("Invalid input: try putting a number there")
    
    while(pets < 0):
        num_input = input("How many pets do you have? ")
        try:
            pets = int(num_input)
            if pets < 0:
                print("Invalid input: you cannot have a negative number of pets.")
        except:
            print("Invalid input: try putting a number there")
    
    while(cars < 0):
        num_input = input("How many cars do you own? ")
        try:
            cars = int(num_input)
            if cars < 0:
                print("Invalid input: you cannot have a negative number of cars.")
        except:
            print("Invalid input: try putting a number there")

    while mortgage_rent < 0.0:
        num_input = input("What is your current mortgage or rent per month? ")
        try:
            mortgage_rent = float(num_input)
            if mortgage_rent < 0.0:
                print("Invalid input: mortgage or rent cannot be negative.")
        except:
            print("Invalid input: try putting a number there")
    
    while auto_payments < 0.0:
        num_input = input("How much do you pay for your car loans per month? ")
        try:
            auto_payments = float(num_input)
            if auto_payments < 0.0:
                print("Invalid input: car loan payment cannot be negative.")
        except:
            print("Invalid input: try putting a number there")
    
    while debt_payments < 0.0:
        num_input = input("What are your current debt payments besides mortgage and car loan? ")
        try:
            debt_payments = float(num_input)
            if debt_payments < 0.0:
                print("Invalid input: debt payments cannot be negative.")
        except:
            print("Invalid input: try putting a number there")

    while utilities < 0.0:
        num_input = input("How much do you pay on utilities every month? ")
        try:
            debt_payments = float(num_input)
            if debt_payments < 0.0:
                print("Invalid input: utilities cannot be negative.")
        except:
            print("Invalid input: try putting a number there")
    
    while insurance < 0.0:
        num_input = input("How much do you pay on insurance (medical, pet, auto, home, life) every month? ")
        try:
            debt_payments = float(num_input)
            if debt_payments < 0.0:
                print("Invalid input: insurance cannot be negative.")
        except:
            print("Invalid input: try putting a number there")
    
    while food <= 0.0:
        num_input = input("How much do you spend on food each month. Include doordash, grocery store, and restaurants.")
        try:
            food = float(num_input)
            if food <= 0.0:
                print("Invalid input: You must spend money on food.")
        except:
            print("Invalid input: try putting a number there")

    
    return fileExists