user_input = int(input("please input an int\n"))
print(user_input)
#
# if(user_input == 1):
#     exit()
# elif(user_input%2 != 2):
#     user_input = 3*(user_input)+1
# else:
#     user_input = (user_input)/2
global cycle

cycle = 0
def threeNPlusOne(n):

    if(n == 1):
        # print(n)
        cycle += 1
        print("cycle = "+ str(cycle))
        exit()
    elif(n%2 != 0):
        n = 3*n+1
        print(n)
        cycle += 1
        threeNPlusOne(n)
    else:
        n = n//2
        print(n)
        cycle += 1
        threeNPlusOne(n)


threeNPlusOne(user_input)


