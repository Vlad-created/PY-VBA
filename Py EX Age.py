#'Types  of data 
int age = 31
float price = 9.99
str name = "Vladisalv"
bool is_active = True 
list colors = ["red", "green"]
tuple coords = (42.5,23,3)
dict user = {"name": "Vladislav", "age": 35}
set unique = {"apple", "banana"}

if age >32
print ("You're cool")
else:
print ("You're the best")

#
age = 30 
if age <18:
    print ("Minor")
elif age < 65:
    print ("Adult")
else: 
    print ("Senior")

#
colors = ["red","blue","green"]

for color in colors:
    print ("Color:",color)

counter = 0 
while counter <5:
    print ("Counter", counter)
    counter += 1
for number in range (10):
    if number ==5:
    break #Stop the loop 
print (number)
#
cool_colors = ["blue", "green"]
colors = ["red", "yellow", "blue", "green"]
for color in colors:
if color in cool_colors:
print (color, "is cool!")
else: 
print (color, "is good")
