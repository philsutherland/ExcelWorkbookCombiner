from openpyxl import Workbook
import time
from os import listdir
from os.path import isfile, join

book = Workbook()
sheet = book.active

sheet['A1'] = 56
sheet['A2'] = 43

now = time.strftime("%x")
sheet['A3'] = now

files = listdir("Data")

filtered_files = []

for file_ in files:
    if "Model_confusion_matrix" in file_:
        filtered_files.append(file_)

print(filtered_files)
print(len(filtered_files))


# book.save("sample.xlsx")
