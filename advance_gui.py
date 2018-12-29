from tkinter import *
import openpyxl
import os

class App(Tk):

    #command function
    def btnclickfunc(self):
        wb=openpyxl.load_workbook(self.filepath)

        # select sheet
        sheet=wb['Sheet']

        id = sheet.max_row

        student_name = self.student.get()
        student_mark = int(self.marks.get())

        sheet.append((id,student_name,student_mark))

        # save file
        wb.save(self.filepath)
        os.startfile(self.filepath)
        
    #initial function    
    def __init__(self):
        Tk.__init__(self)
        self.title("New Window")
        self.geometry("300x100")

        #variable
        self.filepath = "YOUR/FILE/PATH/example.xlsx"
        self.student = StringVar()
        self.marks = StringVar()

        #label
        self.lbl = Label (self, text="Student Name:" )
        self.lbl.grid(row = 0 , column = 0)
        self.lbl = Label (self, text="Marks:" )
        self.lbl.grid(row = 1 , column = 0)

        #text box
        self.txt_input = Entry(self, textvariable=self.student)
        self.txt_input.grid(row = 0, column = 1)
        self.txt_input2 = Entry(self, textvariable=self.marks)
        self.txt_input2.grid(row = 1, column = 1)

        #button
        self.btn = Button(self, text="SAVE", command = self.btnclickfunc)
        self.btn.grid(row = 2, column = 1)


if __name__ == '__main__':
    App().mainloop()
