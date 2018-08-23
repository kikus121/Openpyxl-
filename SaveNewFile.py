import openpyxl
"""
Takes in workbook and saves it to new file called "workbookUpdated.xlsx"
"""

def saveWorkbook(workbook,name):
    if name.endswith('.xlsx'):
        name = name[:-5]
    name = name + 'Updated' + '.xlsx'

    workbook.save(name)

