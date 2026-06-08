from openpyxl import Workbook
import openpyxl


def budget_maker(fileExists):
    print("Welcome to the budget calculator. I'm going to ask you some questions about your life to determine your budget.")
    print("Furthermore, I'm going to ask what your current spending is, so we can do a side by side comparison.")
    print("At the end, we will see how much leftover money there is, and I will give some recommendations.")
    budget_types = ["monthly pay", "people supported", "pets supported", "cars owned", "mortgage or rent", "auto payments", "debt payments", "utilities", "insurance", "food", "medical", "giving", "investments", "extra"]
    budget_information = [-1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1]


    while budget_information[0] <= 0.0:
        num_input = input("How much money do you take home each month? This is after taxes, 401k, etc. ")
        try:
            budget_information[0] = int(num_input)
            if(budget_information[0] <= 0):
                print("I'm not doing this if you don't take home any monthly pay.")
        except:
            print("Invalid input: try putting a number there")

    while budget_information[1] <= 0:
        num_input = input("How many people are in your family? ")
        try:
            budget_information[1] = int(num_input)
            if budget_information[1] <= 0:
                print("Invalid input: you count as one of those people.")
        except:
            print("Invalid input: try putting a number there")
    
    while budget_information[2] < 0:
        num_input = input("How many pets do you have? ")
        try:
            budget_information[2] = int(num_input)
            if budget_information[2] < 0:
                print("Invalid input: you cannot have a negative number of pets.")
        except:
            print("Invalid input: try putting a number there")
    
    while budget_information[3] < 0:
        num_input = input("How many cars do you own? ")
        try:
            budget_information[3] = int(num_input)
            if budget_information[3] < 0:
                print("Invalid input: you cannot have a negative number of cars.")
        except:
            print("Invalid input: try putting a number there")

    while budget_information[4] < 0.0:
        num_input = input("What is your current mortgage or rent per month? ")
        try:
            budget_information[4] = float(num_input)
            if budget_information[4] < 0.0:
                print("Invalid input: mortgage or rent cannot be negative.")
        except:
            print("Invalid input: try putting a number there")
    
    while budget_information[5] < 0.0:
        num_input = input("How much do you pay for your car loans per month? ")
        try:
            budget_information[5] = float(num_input)
            if budget_information[5] < 0.0:
                print("Invalid input: car loan payment cannot be negative.")
        except:
            print("Invalid input: try putting a number there")
    
    while budget_information[6] < 0.0:
        num_input = input("What are your current debt payments besides mortgage and car loan? ")
        try:
            budget_information[6] = float(num_input)
            if budget_information[6] < 0.0:
                print("Invalid input: debt payments cannot be negative.")
        except:
            print("Invalid input: try putting a number there")

    while budget_information[7] < 0.0:
        num_input = input("How much do you pay on utilities every month? ")
        try:
            budget_information[7] = float(num_input)
            if budget_information[7] < 0.0:
                print("Invalid input: utilities cannot be negative.")
        except:
            print("Invalid input: try putting a number there")
    
    while budget_information[8] < 0.0:
        num_input = input("How much do you pay on insurance (medical, pet, auto, home, life) every month? ")
        try:
            budget_information[8] = float(num_input)
            if budget_information[8] < 0.0:
                print("Invalid input: insurance cannot be negative.")
        except:
            print("Invalid input: try putting a number there")
    
    while budget_information[9] <= 0.0:
        num_input = input("How much do you spend on food each month. Include doordash, grocery store, and restaurants.")
        try:
            budget_information[9] = float(num_input)
            if budget_information[9] <= 0.0:
                print("Invalid input: You must spend money on food.")
        except:
            print("Invalid input: try putting a number there")
    
    while budget_information[10] < 0.0:
        num_input = input("Do you have any medical co-pays?")
        try:
            budget_information[10] = float(num_input)
            if budget_information[10] < 0.0:
                print("Invalid input: medical copays cannot be negative.")
        except:
            print("Invalid input: try putting a number there")
    
    while budget_information[11] < 0.0:
        num_input = input("How much money do you give every month? ")
        try:
            budget_information[11] = float(num_input)
            if budget_information[11] < 0.0:
                print("Invalid input: giving cannot be negative.")
        except:
            print("Invalid input: try putting a number there")
    
    while budget_information[12] < 0.0:
        num_input = input("How much money do you invest every month? ")
        try:
            budget_information[12] = float(num_input)
            if budget_information[12] < 0.0:
                print("Invalid input: investments cannot be negative.")
        except:
            print("Invalid input: try putting a number there")
    
    while budget_information[13] < 0.0:
        num_input = input("How much money do you spend on extra expenditures beyond those listed here? ")
        try:
            budget_information[13] = float(num_input)
            if budget_information[13] < 0.0:
                print("Invalid input: extra expenditures cannot be negative.")
        except:
            print("Invalid input: try putting a number there")

    generate_budget_document(budget_types, budget_information, fileExists)
    return fileExists

def generate_budget_document(types, info, fileExists = False):

    workbook = None
    worksheet = None

    giving_ten_percent = info[0] / 10
    food_cost_min = info[1] * 250
    extras_cost_min = info[1] * 20 + 60
    pet_cost_min = info[2] * 20
    pet_insurance_min = info[2] * 20
    insurance_min = info[3] * 50 + 150
    
    if(not fileExists):
        workbook = Workbook()
        worksheet = workbook.active
        fileExists = True
    else:
        workbook = openpyxl.load_workbook("InterestCalculation.xlsx")
        worksheet = workbook.create_sheet()
    
    worksheet.title = "Budget creator $%.2f" % info[0]

    worksheet["A1"] = "Budget calculation"
    worksheet["A2"] = "Monthly pay: $%.2f" % info[0]
    worksheet["C1"] = "Current Spending"
    worksheet["D1"] = "Frugal Chart"
    worksheet["E1"] = "Generous Chart"
    money_format = "$#,##0.00"

    for i in range(len(types)):
        worksheet["B%d" % (i + 2)] = types[i]
    
    worksheet["B16"] = "Total expenditures"
    worksheet["B17"] = "Remaining Funds"

    for c in "CDE":
        worksheet[c + "2"] = info[0]
        worksheet[c + "2"].number_format = money_format
        worksheet[c + "3"] = info[1]
        worksheet[c + "4"] = info[2]
        worksheet[c + "5"] = info[3]
        worksheet[c + "6"] = info[4]
        worksheet[c + "7"] = info[5]
        worksheet[c + "8"] = info[6]
        worksheet[c + "9"] = info[7]
        worksheet[c + "12"] = info[10]
    
    worksheet["C10"] = info[8]
    worksheet["C11"] = info[9]
    worksheet["C13"] = info[11]
    worksheet["C14"] = info[12]
    worksheet["C15"] = info[13]
    worksheet["C16"] = f"=SUM(C6:C15)"
    worksheet["C17"] = f"=C2-C16"
    worksheet["D10"] = insurance_min + pet_insurance_min
    worksheet["D11"] = food_cost_min
    worksheet["D13"] = min(100, giving_ten_percent)
    worksheet["D14"] = 250
    worksheet["D15"] = extras_cost_min + pet_cost_min
    worksheet["D16"] = f"=SUM(D6:D15)"
    worksheet["D17"] = f"=D2-D16"
    worksheet["E10"] = (insurance_min + pet_insurance_min) * 1.25
    worksheet["E11"] = food_cost_min * 1.5
    worksheet["E13"] = giving_ten_percent
    worksheet["E14"] = 1000
    worksheet["E15"] = extras_cost_min + pet_cost_min + 250
    worksheet["E16"] = f"=SUM(E6:E15)"
    worksheet["E17"] = f"=E2-E16"

    for c in "CDE":
        for i in range(6, 18):
            worksheet["%c%d" % (c, i)].number_format = money_format

    for col in worksheet.columns:
        length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if(len(str(cell.value)) > length):
                    length = len(str(cell.value))
            except:
                pass
        worksheet.column_dimensions[column].width = length + 2
    
    workbook.save("InterestCalculation.xlsx")
    print("Finished creating budget documents.")

        
    return fileExists