import openpyxl
from openpyxl import *
from manipular_Hoja1 import *
from manipular_INFORME_SOLICITUDES import *


wb = openpyxl.load_workbook('test.xlsx')
ws = wb['Hoja1']

find_table_and_move_to_A5(ws)

wb.save('result.xlsx')
abrir_excel('result.xlsx')