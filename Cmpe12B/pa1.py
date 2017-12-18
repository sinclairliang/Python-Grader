def reverse1(List, a):
    if a > 0:
        print(List[a-1])
        reverse1(List, a-1)
    return 0

Array = [-1, 2, 6, 12, 9, 2, -5, -2, 8, 5, 7]
Array2 = [-1, 2, 6, 12, 9, 2, -5, -2, 8, 5, 7]

def reverse3(List,b,c):
    if b < ((len(List)/2)):
        reverse3(List,b+1,c-1)
    temp = List[c]
    List[c] = List[b]
    List[b] = temp

    # temp = List[c-2]
    # List[c-2] = List[b+1]
    # List[b+1] = temp
def maxIndex(List,a,b):
    mid = 0
    maxindex = 0
    if (a<b):
        mid = int((a+b)/2)
        maxLeft = maxIndex(List,a,mid)
        maxRight = maxIndex(List,mid+1,b)
        # print(maxLeft,maxRight)
        if List[maxLeft] > List[maxRight]:
            maxindex = maxLeft
        else:
            maxindex = maxRight
    return maxindex

def maxie(List,a,b):
    maxindex = 0
    if List[a] > List[b]:
        maxindex = a
    else:
        maxindex = b
    return maxindex


# reverse3(Array,0,(len(Array)+1))

# print(Array)

mins = maxIndex(Array2,0,(len(Array2)-1))
print(mins)
