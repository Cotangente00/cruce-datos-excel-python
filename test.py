import openpyxl
from manipular_Hoja1 import *

wb = openpyxl.load_workbook('test.xlsx')
ws = wb['Hoja1']

find_table_and_move_to_A5(ws)

wb.save('result.xlsx')