import Functions as fl
import openpyxl
import DataStructs
from openpyxl.utils import column_index_from_string, get_column_letter


def letterToNum(letter):
    return column_index_from_string(letter)

def findSizeOfBox(lib):
    startCell = lib["excelSheetLocation"]["CellStart"]
    endCell = lib["excelSheetLocation"]["CellEnd"]
    startRow = ""
    endRow = ""
    startColumn = ""
    endColumn = ""
    list = [1,2,3,4]

    #Find start/end column
    for i in  range(0,2):
        startColumn = startColumn + startCell[i] 
        endColumn = endColumn + endCell[i]

    
    startColumn = letterToNum(startColumn)
    endColumn = letterToNum(endColumn)
    list[0] = startColumn
    list[2] = endColumn

    #find start row
    for i in range(-(len(startCell)-2),0):
        startRow = startRow + startCell[i] 

    startRow = int(startRow)
    list[1] = startRow
     
    for i in range(-(len(endCell)-2),0):
        endRow = endRow + endCell[i]  

    endRow = int(endRow)
    list[3] = endRow

    return list

