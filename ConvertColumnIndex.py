import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

def letterToNum(letter):
    return column_index_from_string(letter)
    

def numToLetter(num):
    return get_column_letter(num)
    


print(letterToNum('AAA'))


print(numToLetter(234))
