import xlsxwriter
import pandas

def generate_interest():
    print("Welcome to the investment calculator. Here, we will take a starting amount, interest rate, and time and give you a final value")
    starting_value = -1.0
    interest_rate = -1.0
    time = -1
    period_type = ""
    added_investment = 0.0
    will_add = ""

    while starting_value <=0.0:
        num_input = input("How much money are we starting with? ")
        try:
            starting_value = float(num_input)
            if(starting_value <= 0.0):
                print("Invalid input: positive amounts only. If you want to discuss negative amounts, go to the loan calculator")
        except:
            print("Invalid input: try putting a number there")

    while interest_rate <= 0.0:
        num_input = input("What is the annual interest rate? ")
        try:
            interest_rate = float(num_input)
            if(interest_rate <= 0.0):
                print("If your interest rate is negative, don't invest in that source. Put your money into savings or a low-yield safe investment.")
        except:
            print("Invalid input: try putting a number there")



    while period_type != "month" and period_type != "year":
        period_type = input("How often are is the interest and additional investments calculated? Each month, or each year? ").lower()

    while time < 0:
        num_input = input("How many %ss is the money going to sit there? " % (period_type))
        try:
            time = int(num_input)
            if(time < 0):
                print("Invalid input: time must be a positive number. We can't take our money back in time, only forward.")
        except:
            print("Invalid input: try putting a number there")
    
    while will_add != 'yes' and will_add != 'no':
        will_add = input("Will you add additional investments (yes or no)? ")
    
    if(will_add == "yes"):
        while added_investment <= 0.0:
            num_input = input("How much money will you add each %s? " % (period_type))
            try:
                added_investment = float(num_input)
            except:
                print("Please enter a number greater than $0.00. If you have a retirement account and are pulling out money, try the retirement calculator.")
    


    generate_interest_document(starting_value, interest_rate, time, period_type, added_investment)
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

    worksheet.write("A1", "Budget calculation")
    worksheet.write("A2", "Starting amount: %.2f" % (start))
    worksheet.write("A3", "Interest rate: %.4f percent per %s" % (interest, duration_of_period))
    worksheet.write("A4", "Additional investment of %.2f per %s" % (additional_money, duration_of_period))
    worksheet.write("A5", "Total time investigated: %d %ss" % (periods, duration_of_period))
    worksheet.write("B2", duration_of_period)
    worksheet.write("C2", "Money at %s" % (duration_of_period))
    money_format = workbook.add_format({'num_format': '$#,##0.00'})


    for i in range(periods + 1):
        print("Money at the beginning of %s %d: %d" % (duration_of_period, i, money_by_period[i]))
        worksheet.write(i+2, 1, i)
        worksheet.write(i+2, 2, money_by_period[i], money_format)

    worksheet.autofit()

    workbook.close()
    print("Finished creating financial documents.")
    return