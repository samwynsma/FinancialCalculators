def main():
    print("Python test is working")
    print("Invesment calculator is the first thing I will add here")
    starting_value = -1.0
    while starting_value < 0.0:
        num_input = input("How much money are we starting with: ")
        try:
            starting_value = float(num_input)
        except:
            print("invalid input: try putting a number there")
    

if __name__ == "__main__":
    main()