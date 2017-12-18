def poly(c, y):
    #result = (1*y)**(3)+(2*y)**(2)+(4*y)**(1)
    rounds = len(c)-1
    i = 0
    result = 0
    while rounds > 0:
        result += (c[i]*(y**(rounds)))
        rounds = rounds - 1
        i += 1
    result += c[len(c)-1]
    return (result)

def diff(c):
    rounds = len(c)-1
    d = []
    i = 0
    while rounds > 0:
        d.append(rounds*c[i])
        i += 1
        rounds = rounds - 1
    return d

def findroot(C, a, b, tolerance):
    width = b - a
    mid = (a+b)/2
    while width > tolerance:
        mid = (a+b)/2
        if ((poly(C, a) * poly(C, mid)) <= 0):
            b = mid
        else:
            a = mid
        width = b - a
    return mid

def display():
    #array = [int(x) for x in input().split()]
    array = [2,0,0]
    resultion = 10**(-2)
    tolerance = 10**(-7)
    threshold = 10**(-3)
    L = -10
    R = 10
    l = L
    r = l + resultion
    while r <= R:
        # if (poly(array,l)*poly(array,r)) < 0:
        #     re1 = findroot(array, l, r, tolerance)
        #     print(re1)
        if (poly(diff(array),l)*poly(diff(array),r)) < 0:
            re = findroot((diff(array)),l,r,tolerance)
            if abs(poly(array,re)) < threshold:
                print(re)
        l = r
        r = r + resultion
    return 0



#array = [int(x) for x in input().split()]
#array = [-6, 11, -6, 1]
array = [2,0,0]
#print(poly(array,2))
#re = findroot(array, -1, 1, 0.001)
#print(re)
#display()
#print(findroot(diff(array),-1, 1, 0.00000001))
re = findroot((diff(array)),-1,1,0.01)
print(poly(array,re))
