import csv
from os import listdir
from os.path import isfile, join
import openpyxl


files = listdir("data")

for file_ in files:
    wb = openpyxl.Workbook()
    ws = wb.active

    print('data/' + file_)

    with open('data/' + file_, 'r') as f:
        for row in csv.reader(f):
            ws.append(row)

    wb.save(f'renovated data/{file_.split(".")[0]}.xlsx')
