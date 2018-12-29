import openpyxl
import os
from tkinter import *
from tkinter import simpledialog

mainwindow = Tk()
mainwindow.withdraw()
filepath = "YOUR/FILE/PATH/example.xlsx"
wb=openpyxl.load_workbook(filepath)

# select sheet
sheet=wb['Sheet']

id = sheet.max_row

student = simpledialog.askstring("Name","Enter student name")
mark = simpledialog.askinteger("Test Mark","Enter marks")

sheet.append((id,student,mark))

# save file
wb.save(filepath)
os.startfile(filepath)
