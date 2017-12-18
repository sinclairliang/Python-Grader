def odd_or_even():
    number = input("what is your number?\n")
    try:
        number = int(number)
        if number % 2 == 0:
            print("It is an even number!")
        else:
            print("It is an odd number!")
    except:
        print("It must be an int")


odd_or_even()
