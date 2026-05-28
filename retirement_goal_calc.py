def retirement_goal(fileExists):
    print("Welcome to the retirement goals calculator. The goal here is to see how much you need to save for retirement to survive.")
    print("First, I will ask you questions to gather all the information that I need.")
    monthly_needs = -1.0
    current_savings = -1.0
    growth_rate = -1.0
    current_age = -1.0
    expected_retirement_age = -1.0
    expected_death_age = -1.0

    while(monthly_needs <= 0.0):
        num_input = input("How much money will you need to take out each month to live?")
        try:
            monthly_needs = float(num_input)
            if(monthly_needs <= 0.0):
                print("Invalid input: you will need to make a certain amount of money to live.")
        except:
            print("Invalid input: try putting a number there")

    return fileExists