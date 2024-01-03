import os

directory = [f for f in os.listdir() if '.' not in f and f != "__pycache__" and f != "languages"]

for d in directory:
    for person in os.listdir(d):
        print(person)