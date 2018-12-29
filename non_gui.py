import openpyxl
import os

filepath = "YOUR/FILE/PATH/example.xlsx"
wb=openpyxl.load_workbook(filepath)

# select sheet
sheet=wb['Sheet']

id = sheet.max_row

student = input("Enter student name:")
mark = int(input("Enter marks:"))

sheet.append((id,student,mark))

# save file
wb.save(filepath)
os.startfile(filepath)
