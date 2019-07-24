import xlrd
from tkinter import filedialog

def main(file = None):

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
    try:
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
                    log.write('col:{col}\trow:{row}\trepeat_value:{value}\n'.format(col=col, row=row, value=value))
                    log.close()
                    check = False
                else:
                    data.append(value)
    except:
        pass

    if check:
        print('Vertical Checks passed')
    else:
        print('1 or more Vertical Checks failed see more info @ verticalCheck.txt')
    
    return check
