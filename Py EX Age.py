name_age = {
    "Vlad": 31, 
    "Vladic": 3, 
    "Kristina":30,
    "Alina": 0.6
}
unique = ["Vlad","Kristina","Alina","Vladic"]

for name in unique:
    if name in name_age:
        print (name_age[name])
        break
else:
    print("No match name")

