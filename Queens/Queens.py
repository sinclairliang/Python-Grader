


def reverse2(li):
    n = len(li)
    R = []
    i = 0
    while i < n:
        R.append(li[n-i-1])
        i += 1
    return R

def swap(L,a,b):
    temp = L[a]
    L[a] = L[b]
    L[b] = temp


def nextPermutation1(nums):
    pivot = 0
    for i in range(len(nums)-2,-1,-1):
        if nums[i] < nums[i+1]:
            pivot = i
            break
    else:
        nums.reverse()
        return 0
    successor = 0
    for j in range(len(nums)-1,-1,-1):
        if nums[j] > nums[pivot]:
            successor = j
            swap(nums,pivot,successor)
            break
    nums[pivot+1:] = reversed(nums[pivot+1:])
    return 0
# for i in range(24):
#     nextPermutation1(array)
array = [0,1, 3, 5, 2, 4]
array1 = [3,1]

def isSolution(List):
    k = 0
    n = len(List)
    for i in range(1,n,1):
        for j in range(i+1,n,1):
            vertical = abs(List[i] - List[j])
            horizontal = abs(i - j)
            k += 1
            print("comparing "+ str(i)+"th element " + str(j)+"th element "+str(List[i])+","+str(List[j]))
            print("i = "+str(List[i])+" j= "+str(List[j]))
            print("horizontal distance = "+ str(horizontal))
            print("Vertical distance = "+str(vertical))
            print("status = "+ str(not(vertical==horizontal)))
            if (vertical == horizontal) == (True):
                 return False
            elif k == ((n-1)*(n-2)/2):
                return True
            else:
                continue

    #return 0


# array2 = [1,2,3,4,5]
# for i in range(120):
#     nextPermutation1(array2)
#     if isSolution(array2) == True:
#         print(array2)
#     else:
#         continue

def factorio(n):
    if n == 1:
        return 1
    else:
        return factorio(n-1)*n


print(isSolution(array))
