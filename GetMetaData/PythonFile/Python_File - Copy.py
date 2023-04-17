import sys
def add_numbers(x,y):
   sum = x + y
   return sum

num1 = int(sys.argv[1])
num2 = int(sys.argv[2])
print(add_numbers(num1, num2))
