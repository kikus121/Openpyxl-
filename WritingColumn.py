#Takes a list of values and updated a specific woorkbook

def writeColumn(list,column,startRow,endRow,workbook,sheetName):
    index = 0
    wb = openpyxl.load_workbook(workbook)
    sheet = wb[sheetName]

    for i in range(startRow,endRow):
        cell = sheet.cell(row = i, column = column)
        if list[index]:
            cell.value = list[index]
        else:
            cell.value = 'None'
        index += 1

    saveWorkbook(wb,workbook)
