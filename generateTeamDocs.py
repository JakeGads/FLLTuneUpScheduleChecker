import xlrd #.cell(rowx, colx)
# import docx
from tkinter import filedialog
from datetime import time
class Event():
    def __init__(self, title, time, session):
        self.title = title
        self.session = session
        self.time = time
    def __str__(self):
        return '{time}:{session}\t{title}'.format(time=self.time, session=self.session, title=self.title)

def time_convert(x):
    x = int(x * 24 * 3600) # convert to number of seconds
    hour = time(x//3600, (x%3600)//60).hour
    if hour > 12:
        hour -= 12
    minute = time(x//3600, (x%3600)//60).minute + 1
    if minute < 10:
        minute = '0' + str(minute)
    return str(hour) + ':' + str(minute) 


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
                room = sheet.cell(rowx=row,colx=col).value.replace('-ii','').replace('-i','')
                continue
            if col is 1:
                number = sheet.cell(rowx=row,colx=col).value
                continue
            if col is 2:
                name = sheet.cell(rowx=row,colx=col).value
                continue
            if col is 3:
                continue

            if sheet.cell_type(row, col) != xlrd.empty_cell and sheet.cell_value(row,col) != '':
                sched.append(Event(sheet.cell(rowx=row,colx=col).value, time_convert(sheet.cell(rowx=1,colx=col).value), sheet.cell(rowx=0,colx=col).value))

        print('\n\n{room}\t{number}\t{name}'.format(room=room, number=number, name=name))
        
        for i in sched:
            print('\t\t\t'+str(i))

