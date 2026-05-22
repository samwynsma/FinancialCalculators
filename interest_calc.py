import xlsxwriter

def generate_interest():
    print("Welcome to the investment calculator. Here, we will take a starting amount, interest rate, and time and give you a final value")
    starting_value = -1.0
    interest_rate = -1.0
    time = -1
    money_by_year = []
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

    money_by_year.append(starting_value)
    current_money = starting_value
    for i in range(time):
        current_money = round(current_money * (1 + interest_rate / 100), 2)
        money_by_year.append(current_money)
    

    workbook = xlsxwriter.Workbook("InterestCalculation.xlsx")
    worksheet = workbook.add_worksheet()

    worksheet.write("A1", "Budget calculation at %d percent interest" % (interest_rate))
    worksheet.write("A2", "Year")
    worksheet.write("B2", "Money at Year")


    for i in range(time + 1):
        print("Money at the beginning of year %d: %d" % (i, money_by_year[i]))
        worksheet.write(i+2, 0, i)
        worksheet.write(i+2, 1, money_by_year[i])

    workbook.close()
    print("Finished creating financial documents.")
    return

def generate_interest_document():
    return