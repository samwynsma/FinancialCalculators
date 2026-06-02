import openpyxl
from openpyxl import Workbook

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
        num_input = input("How much money will you need to take out each month to live? ")
        try:
            monthly_needs = float(num_input)
            if monthly_needs <= 0.0:
                print("Invalid input: you will need to make a certain amount of money to live.")
        except:
            print("Invalid input: try putting a number there")
    
    while current_savings < 0.0:
        num_input = input("How much money do you currently have saved? Do not include debts or non-liquid assets here. ")
        try:
            current_savings = float(num_input)
            if current_savings < 0.0:
                print("Invalid input: savings cannot be negative.")
        except:
            print("Invalid input: try putting a number there")
    
    while growth_rate < 0.0:
        num_input = input("What is the percentage growth rate that you expect your retirement savings to grow at. Note: this growth continues after retirement. ")
        try:
            growth_rate = float(num_input)
            if growth_rate < 0.0:
                print("Invalid input: interest rate must be positive.")
        except:
            print("Invalid input: try putting a number there")
    
    while current_age < 18 or current_age >= 70:
        num_input = input("How old are you right now? ")
        try:
            current_age = int(num_input)
            if(current_age < 18):
                print("If you are younger than eighteen, retirement is nice to focus on, but you might not be legally able to invest in a good account. Wait until you're eighteen.")
            elif(current_age >= 70):
                print("You are over the maximum age of social security in the USA. Use the retirement option or the supplemental income option from the menu.")
        except:
            print("Invalid input. Age must be a number between eighteen and seventy. For older users, use the retirement option to calculate your current retirement account, or the supplemental income option to determine how much more money you need to make.")
    
    while expected_retirement_age <= current_age or expected_retirement_age > 70:
        num_input = input("How old do you expect to be when you retire? ")
        try:
            expected_retirement_age = int(num_input)
            if(expected_retirement_age <= current_age):
                print("If you are already retired and you're looking to see how far you'll make it, try the retirement option instead. ")
            elif(expected_retirement_age > 70):
                print("In the USA, retiring past 70 carries a social security penalty. For the sake of this, please set your retirement age earlier. ")
        except:
            print("Invalid input. Please select a number greater than your current age and up to seventy. ")
    
    while expected_death_age <= expected_retirement_age or expected_death_age > 105:
        num_input = input("When do you expect to expire, or pass away? ")
        try:
            expected_death_age = int(num_input)
            if(expected_death_age <= expected_retirement_age):
                print("If you expect to die before you retire, then this app is pointless. Let's assume that you won't, okay? ")
            elif(expected_death_age > 105):
                print("You have a less than 0.01 percent chance to make it that far. Let's be a little more reasonable. ")
        except:
            print("Invalid input. Please select a number greater than your retirement age and up to one hundred and five. ")

    fileExists = generate_retirement_document(monthly_needs, current_savings, growth_rate, current_age, expected_retirement_age, expected_death_age, fileExists)
    return fileExists

def generate_retirement_document(monthly_needs, current_savings, growth_rate, age, retirement, death, fileExists = False):
    month_age = age * 12
    month_ret = retirement * 12
    month_death = death * 12
    month_growth = growth_rate / 12.0

    age_to_ret_gap = month_ret - month_age
    ret_to_death_gap = month_death - month_ret

    scenario_one = current_savings # Scenario one: money lasts until death plus five years
    scenario_two = current_savings # Scenario two: account maintains itself: growth rate = take out rate.
    scenario_three = current_savings # Scenario three: account user takes out half the growth at the starting year.

    scenario_one_goal = monthly_needs * (1 - pow(1 + (month_growth / 100.0), -(ret_to_death_gap + 60))) / (month_growth / 100.0)
    scenario_two_goal = monthly_needs / (month_growth / 100.0)
    scenario_three_goal = monthly_needs / (month_growth / 200.0)

    print(round(scenario_one_goal, 2))
    print(scenario_two_goal)
    print(scenario_three_goal)
    print(" ")

    scenario_one_monthly = (scenario_one_goal - current_savings) * (month_growth / 100.0) / (pow(1.0 + month_growth / 100.0, age_to_ret_gap) - 1.0)
    scenario_two_monthly = (scenario_two_goal - current_savings) * (month_growth / 100.0) / (pow(1.0 + month_growth / 100.0, age_to_ret_gap) - 1.0)
    scenario_three_monthly = (scenario_three_goal - current_savings) * (month_growth / 100.0) / (pow(1.0 + month_growth / 100.0, age_to_ret_gap) - 1.0)

    print(scenario_one_monthly)
    print(scenario_two_monthly)
    print(scenario_three_monthly)

    money_s1 = []
    money_s2 = []
    money_s3 = []

    money_s1.append(scenario_one)
    money_s2.append(scenario_two)
    money_s3.append(scenario_three)

    for i in range(age_to_ret_gap):
        scenario_one = round(scenario_one * (1 + month_growth / 100.0), 2) + scenario_one_monthly
        scenario_two = round(scenario_two * (1 + month_growth / 100.0), 2) + scenario_two_monthly
        scenario_three = round(scenario_three * (1 + month_growth / 100.0), 2) + scenario_three_monthly
        money_s1.append(scenario_one)
        money_s2.append(scenario_two)
        money_s3.append(scenario_three)
    
    workbook = None
    worksheet = None

    if(not fileExists):
        workbook = Workbook()
        worksheet = workbook.active
        fileExists = True
    else:
        workbook = openpyxl.load_workbook("InterestCalculation.xlsx")
        worksheet = workbook.create_sheet()
    

    worksheet.title = "Retirement %d." % int(monthly_needs)

    worksheet["A1"] = "Retirement Calculation"
    worksheet["A2"] = "Current savings: $%.2f" % current_savings
    worksheet["A3"] = "Growth rate: %.4f percent per month" % growth_rate
    worksheet["A4"] = "For money to last until death: $%.2f per month." % scenario_one_monthly
    worksheet["A5"] = "For money to be sustainable: $%.2f per month." % scenario_two_monthly
    worksheet["A6"] = "For money to be economy-resilient: $%.2f per month." % scenario_three_monthly


    workbook.save("InterestCalculation.xlsx")
    print("Finished creating financial documents.")
    return fileExists