
bigger = int(input("Enter a positive integer: "))
smaller = int(input("Enter another positive integer: "))
r = bigger%smaller
while r != 0:
    bigger = smaller
    smaller = r
    #result = bigger // smaller
    r = bigger%smaller
    #print(str(bigger) + "/" +str(smaller) + "=" + str(result) + "..." + str(r))
    #print(str(bigger) + "=" +str(smaller) + "*" + str(result) + "+" + str(r))
print(smaller)
