import xlrd #.cell(rowx, colx)
import docx
from tkinter import filedialog

class Event():
    def __init__(self, title, time, session):
        self.title = title
        self.session = session
        self.time = time


def main(file = None):
    if file is None:
        book = xlrd.open_workbook(filedialog.askopenfilename(title = "Select File",filetypes = (("xlsx files","*.xlsx"),("xls files","*.xls"),("all files","*.*"))))
    else:
        book = xlrd.open_workbook(file)

    sheet = book.sheet_by_index(0)

    for row in range(sheet.nrows):
        if row is 0 or row is 1:
            continue
        
        sched = []
        room = ''
        number = ''
        name = ''
        for col in range(sheet.ncols):
            if col is 0:
                room = sheet.cell_value(row,col)
                continue
            if col is 1:
                number = sheet.cell_value(row, col)
                continue
            if col is 2:
                name = sheet.cell_value(row,col)
            
            if sheet.cell_type(row, col) is not xlrd.empty_cell:
                sched.append(Event(sheet.cell_value(row, col), sheet.cell_value(1, col), sheet.cell(0, col)))
