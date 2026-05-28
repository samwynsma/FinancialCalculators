def retirement_goal(fileExists):
    print("Welcome to the retirement goals calculator. The goal here is to see how much you need to save for retirement to survive.")
    print("First, I will ask you questions to gather all the information that I need.")
    monthly_needs = -1.0
    current_savings = -1.0
    growth_rate = -1.0
    current_age = -1
    expected_retirement_age = -1
    expected_death_age = -1

    while monthly_needs <= 0.0:
        num_input = input("How much money will you need to take out each month to live?")
        try:
            monthly_needs = float(num_input)
            if monthly_needs <= 0.0:
                print("Invalid input: you will need to make a certain amount of money to live.")
        except:
            print("Invalid input: try putting a number there")
    
    while current_savings < 0.0:
        num_input = input("How much money do you curretly have saved? Do not include debts or non-liquid assets here.")
        try:
            monthly_needs = float(num_input)
            if monthly_needs < 0.0:
                print("Invalid input: savings cannot be negative.")
        except:
            print("Invalid input: try putting a number there")
    
    while growth_rate < 0.0:
        num_input = input("What is the percentage growth rate that you expect your retirement savings to grow at. Note: this growth continues after retirement.")
        try:
            monthly_needs = float(num_input)
            if monthly_needs < 0.0:
                print("Invalid input: interest rate must be positive.")
        except:
            print("Invalid input: try putting a number there")
    
    while current_age < 18 or current_age > 70:
        num_input = input("How old are you right now?")
        try:
            current_age = int(num_input)
            if(current_age < 18):
                print("If you are younger than eighteen, retirement is nice to focus on, but you might not be legally able to invest in a good account. Wait until you're eighteen.")
            elif(current_age > 70):
                print("You are over the maximum age of social security in the USA. Use the retirement option or the supplemental income option from the menu.")
        except:
            print("Invalid input. Age must be a number between eighteen and seventy. For older users, use the retirement option to calculate your current retirement account, or the supplemental income option to determine how much more money you need to make.")
    

    return fileExists