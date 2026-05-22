import xlsxwriter

def generate_interest():
    print("Welcome to the investment calculator. Here, we will take a starting amount, interest rate, and time and give you a final value")
    starting_value = -1.0
    interest_rate = -1.0
    time = -1
    while starting_value < 0.0:
        num_input = input("How much money are we starting with: ")
        try:
            starting_value = float(num_input)
        except:
            print("Invalid input: try putting a number there")
    while interest_rate < 0.0:
        num_input = input("What is the annual interest rate: ")
        try:
            interest_rate = float(num_input)
        except:
            print("Invalid input: try putting a number there")
    while time < 0:
        num_input = input("How many years is the money going to sit there: ")
        try:
            time = int(num_input)
        except:
            print("Invalid input: try putting a number there")

    generate_interest_document(starting_value, interest_rate, time)
    return

def generate_interest_document(start, interest, periods, duration_of_period="Year", additional_money=0.0):
    money_by_period = []
    money_by_period.append(start)

    if(duration_of_period == "month"):
        interest = interest / 12
        duration_of_period = "Month"
    else:
        duration_of_period = "Year"

    current_money = start
    for i in range(periods):
        current_money = round(current_money * (1 + interest / 100), 2) + additional_money
        money_by_period.append(current_money)
    

    workbook = xlsxwriter.Workbook("InterestCalculation.xlsx", {'strings_to_numbers' : True})
    worksheet = workbook.add_worksheet("Money with Interest")

    worksheet.write("A1", "Budget calculation at %d percent interest" % (interest))
    worksheet.write("A2", duration_of_period)
    worksheet.write("B2", "Money at %s" % (duration_of_period))
    money_format = workbook.add_format({'num_format': '$#,##0.00'})


    for i in range(periods + 1):
        print("Money at the beginning of year %d: %d" % (i, money_by_period[i]))
        worksheet.write(i+2, 0, i)
        worksheet.write(i+2, 1, money_by_period[i], money_format)

    workbook.close()
    print("Finished creating financial documents.")
    return