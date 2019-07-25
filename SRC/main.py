from tkinter import filedialog
import os
from datetime import time
# User defined
import xlrd

def reversColConvert(x):
    alpha = ['a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z']
    length = len(alpha)
    if length >= 2:
        #creates the dual letter colums
        for i in range (0, length):
            for h in range (0, length):
                alpha.append((alpha[i] + alpha[h]))
        if length >= 3:
            # creates the tri colums
            for i in range (0, length): # for(int i = 0; i < length; i++)
                for h in range (0, length):
                    for j in range (0, length):
                        alpha.append((alpha[i] + alpha[h] + alpha[j]))            

    return alpha[x]

def verticalCheck(file = None):

    log = open('verticalCheck.txt', 'w+')
    log.close()

    if file is None:
        book = xlrd.open_workbook(filedialog.askopenfilename(title = "Select File",filetypes = (("xlsx files","*.xlsx"),("xls files","*.xls"),("all files","*.*"))))
    else:
        book = xlrd.open_workbook(file)

    sheet = book.sheet_by_index(0)

    nrows = sheet.nrows - 1
    ncols = sheet.ncols - 1

    check = True

    for col in range(ncols):
        if col in [0,1,2,3]:
            continue

        data = []
        for row in range(nrows):
            value = sheet.cell(rowx=row, colx=col).value

            if value == '' or value == 'Judge’s Break' or value == 'Coach Meeting' or value == 'Opening Ceremony' or value == 'Line Dancing':
                continue

            if value in data:
                log = open('verticalCheck.txt', 'a+')
                log.write('col:{col}\trow:{row}\trepeat_value:{value}\n'.format(col= reversColConvert(col), row=row + 1, value=value))
                log.close()
                check = False
            else:
                data.append(value)

    if check:
        print('Vertical Checks passed')
    else:
        print('1 or more Vertical Checks failed see more info @ verticalCheck.txt')
    
    return check

def roomCheck(file=None):
    log = open('roomCheck.txt', 'w+')
    log.close()
    if file is None:
        book = xlrd.open_workbook(filedialog.askopenfilename(title = "Select File",filetypes = (("xlsx files","*.xlsx"),("xls files","*.xls"),("all files","*.*"))))
    else:
        book = xlrd.open_workbook(file)

    sheet = book.sheet_by_index(0)

    check = True

    roomCol = 0

    try:
        for row in range(sheet.nrows):
            firstTeam = sheet.cell(rowx=row, colx=roomCol).value.replace('-ii','').replace('-i','')
            secondTeam = sheet.cell(rowx=row + 1, colx=roomCol).value.replace('-ii','').replace('-i','')

            if firstTeam != secondTeam:
                continue

            for col in range(sheet.ncols):
                first_cell = sheet.cell(rowx=row, colx=col).value
                second_cell = sheet.cell(rowx=row + 1, colx=col).value

                if first_cell == 'Judge’s Break' or first_cell == 'Coach Meeting' or first_cell == 'Opening Ceremony' or first_cell == 'Line Dancing':
                    continue

                if first_cell == '' and second_cell == '':
                    continue
                if (first_cell == '' and 'field' in second_cell) or ('field' in first_cell and second_cell == ''):
                    continue

                if 'judge' in first_cell.lower() or 'judge' in second_cell.lower():
                    if 'field' in first_cell.lower() or 'field' in second_cell.lower():
                        None
                    else:
                        check = False
                        log = open('roomCheck.txt', 'a+')
                        log.write('col:{col}\trow:{row}\n'.format(col=reversColConvert(col), row=row + 1))
                        log.close()
                
    
                    
    except:
        pass

    if check:
        print('RoomCheck has been cleared')
    else:
        print('1 or more room failed please check @ roomCheck.txt for more details')

    return check

def generateTeamDocs(file = None):
    class TeamEvent():
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
                    sched.append(TeamEvent(sheet.cell(rowx=row,colx=col).value, time_convert(sheet.cell(rowx=1,colx=col).value), sheet.cell(rowx=0,colx=col).value))

            print('\n\n{room}\t{number}\t{name}'.format(room=room, number=number, name=name))
            
            for i in sched:
                print('\t\t\t'+str(i))

    main(file)






if __name__ == "__main__":
    # file = filedialog.askopenfilename(title = "Select File",filetypes = (("xlsx files","*.xlsx"),("xls files","*.xls"),("all files","*.*")))
    file = 'FLL Schedule.xlsx'


    vert = verticalCheck(file=file) 
    room = roomCheck(file=file)
    
    # generateTeamDocs(file=file) 