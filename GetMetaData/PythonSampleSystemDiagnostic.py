import sys

class calculator:
    def add(self, x, y):
        return x + y

    def increment(self, x):
        x += 1;
        return x;

#creating object of class
calculatorObj = calculator()
#capturing input from command line and casting to integer
x = int(sys.argv[1])
y = int(sys.argv[2])
z = calculatorObj.add(x, y)
#printing result on console
print(z)