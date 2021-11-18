#
# Example file for working with functions
#

# define a basic function
def func1():
    print("I am an Autobot")

# function that takes arguments
def func2(arg1, arg2):
    print('Argument 1: ', arg1)
    print('Argument 2: ', arg2)

# function that returns a value
def cube(x):
    return x*x*x

# function with default value for an argument
def power(num, x=1):
    result = 1
    for i in range(x):
        result*=num
    return result
        

#function with variable number of arguments
def multi_add(*args):
    result = 0
    for x in args:
        result = result + x
    return result


func1()
print(func1())
print(func1)

func2("autobot", "deceptacon")

print(cube(7))
print(power(2))
print(power(2,3))
print(power(x=3,num=2))
print(multi_add(10, 4, 5, 6, 8))