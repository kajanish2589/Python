https://www.w3schools.com/python/python_tuples.asp

thistuple = ("apple",)
print(type(thistuple))

#NOT a tuple
thistuple = ("apple")
print(type(thistuple))


tuple1 = ("apple", "banana", "cherry")
tuple2 = (1, 5, 7, 9, 3)
tuple3 = (True, False, False)

tuple1 = ("abc", 34, True, 40, "male") 

"""
https://www.geeksforgeeks.org/python-difference-between-list-and-tuple/ """ 

# Creating a List with
# the use of Numbers
# code to test that tuples are mutable
List = [1, 2, 4, 4, 3, 3, 3, 6, 5]
print("Original list ", List)

List[3] = 77
print("Example to show mutability ", List)

# code to test that tuples are immutable

tuple1 = (0, 1, 2, 3)
tuple1[0] = 4
print(tuple1)
