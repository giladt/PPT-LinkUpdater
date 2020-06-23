import datetime
import random

def Quicksort(array):

    if len(array)<=1:
        return array

    less=[]
    greater=[]

    pivot = array[int(len(array)/2)]
    array.remove(pivot)
    for item in array:
        if item <= pivot:
            less.append(item)
        else:
            greater.append(item)
    
    return Quicksort(less)+[pivot]+Quicksort(greater)


start = datetime.datetime.now()
vArray = random.sample(range(1, 1000001), 1000000) #
timeStop1 = datetime.datetime.now()
print('Original: Set', vArray[-1])
print('Creating array:', timeStop1-start)

timeStop2 = datetime.datetime.now()
vArray = Quicksort(vArray)
timeStop3 = datetime.datetime.now()
print('Result: Done')
print('Sorting:', timeStop3-timeStop2)

end = datetime.datetime.now()
print('Total runtime:', end-start)