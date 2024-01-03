from Autosaveexcel import Autosave
import os

while True:
    select = input("Person index (None to extract all) : ")
    mouth = input("Input mouth (M/YYYY): ")
    branch = None
    exel = None
    if len(select) == 0:
        select = None
    if select is not None:
        branch = input("branch : ")
        
    for filename in os.listdir(os.getcwd()):
            if filename.endswith('.xlsx'):
                exel = filename

    if exel == None:
        print("Excel file not found!!")
    else:
        save = Autosave(exel,select,branch,mouth)
        save.main()


