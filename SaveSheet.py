import openpyxl

wb = openpyxl.load_workbook('TestSheet.xlsx')
sheet = wb.active
sheet.title = 'Spam Spam Spam'
wb.save('example_copy.xlsx')
