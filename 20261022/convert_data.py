import os
import glob
import openpyxl
import csv

data_dir = os.path.abspath(os.path.join(os.getcwd(), '..', 'data'))
print("Data dir:", data_dir)
files = os.listdir(data_dir)
target_file = None
for f in files:
    if '20260331' in f:
        target_file = os.path.join(data_dir, f)
        break

print("Target file found:", target_file)

wb = openpyxl.load_workbook(target_file, data_only=True)
sheet = wb.active
with open('data.csv', 'w', newline='', encoding='utf-8-sig') as fp:
    writer = csv.writer(fp)
    for row in sheet.iter_rows(values_only=True):
        writer.writerow(row)

print("data.csv created successfully!")
