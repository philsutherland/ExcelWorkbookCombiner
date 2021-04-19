from openpyxl import Workbook
import openpyxl
import time
from os import listdir
from os.path import isfile, join
import csv

book = Workbook()
sheet = book.active

sheet['A1'] = 56
sheet['A2'] = 43

now = time.strftime("%x")
sheet['A3'] = now

files = listdir("renovated data")

filtered_files = []

for file_ in files:
    if "Model_confusion_matrix" in file_:
        filtered_files.append(file_)

print(filtered_files)
print(len(filtered_files))


aphanitic = [0, 0, 0, 0, 0, 0]
breccia = [0, 0, 0, 0, 0, 0]
phaneritic = [0, 0, 0, 0, 0, 0]
phorphyry = [0, 0, 0, 0, 0, 0]
stockwork = [0, 0, 0, 0, 0, 0]
vein = [0, 0, 0, 0, 0, 0]

for file_ in filtered_files:
    # print(rf'{file_.split(".")[0]}')
    wb = openpyxl.load_workbook(rf'renovated data/{file_}')
    sheet = wb.worksheets[0]

    for _ in range(2, 8):
        aphanitic[_ - 2] += float(sheet[f'B{_}'].value)

    for _ in range(2, 8):
        breccia[_ - 2] += float(sheet[f'C{_}'].value)

    for _ in range(2, 8):
        phaneritic[_ - 2] += float(sheet[f'D{_}'].value)

    for _ in range(2, 8):
        phorphyry[_ - 2] += float(sheet[f'E{_}'].value)

    for _ in range(2, 8):
        stockwork[_ - 2] += float(sheet[f'F{_}'].value)

    for _ in range(2, 8):
        vein[_ - 2] += float(sheet[f'G{_}'].value)


print(aphanitic)
print(breccia)
print(phaneritic)
print(phorphyry)
print(stockwork)
print(vein)

# book.save("sample.xlsx")
