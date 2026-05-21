def main():
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

    final_money = round(starting_value * pow(1 + interest_rate/100, time), 2)
    print("Your final money is %d", final_money)
    

if __name__ == "__main__":
    main()