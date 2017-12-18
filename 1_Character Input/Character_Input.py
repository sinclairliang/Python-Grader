def ask_for_age():
    age = int(input("what is your age?\n"))
    year = 2016 - age + 100
    print("You will turn 100 years old in year " + str(year) + "!")
    return 0

ask_for_age()
