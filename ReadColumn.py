import openpyxl

def readColumn(column,startRow,endRow,workbook,sheetName):
    list = []
    wb = openpyxl.load_workbook(workbook)
    sheet = wb[sheetName]
    
    for i in range(startRow,endRow):
        cell = sheet.cell(row = i, column = column)
        list.append(cell.value)
        #print("Value:{}  Row:{}  Coordinates:{}".format(cell.value,cell.row,cell.coordinate))
        #print("")

    return list
    

