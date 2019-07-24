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
            if col is 0 or col is 1 or col is 2 or col is 3:
                continue
            hits = []
            for row in range(nrows):
                alreadyInList = False
                val = sheet.cell_value(rowx=row, colx=col)
                for i in hits:
                    if val == i:
                        alreadyInList = True

                if not alreadyInList:
                    hits.append(val)
                else:
                    check = False
                    log = open('verticalCheck.txt', 'a+')
                    log.write('col:{col}\trow:{row}\trepeat_value:{value}\n'.format(col=col, row=row, value=sheet.cell_value(col,row)))
                    log.close()
    except:
        pass

    if check:
        print('Vertical Checks passed')
    else:
        print('1 or more Vertical Checks failed see more info @ verticalCheck.txt')
    
    return check
