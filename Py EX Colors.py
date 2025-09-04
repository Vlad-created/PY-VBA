colors = ["green", "yellow", "red"]
unique = ["apple", "banana", "Kiwi"]

if "apple" in unique:
    print("red")
elif "banana" in unique:
    print("yellow")
elif "Kiwi" in unique:
    print("green")
else:
    print("No same color")

#
fruit_colors = {
    "apple": "red",
    "banana": "yellow",
    "Kiwi": "green"
}

unique = ["apple", "banana", "Kiwi"]

for fruit in unique:
    if fruit in fruit_colors:
        print(fruit_colors[fruit])
        break
else:
    print("No same color")

