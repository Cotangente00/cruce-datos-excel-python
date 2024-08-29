import openpyxl
from manipular_INFORME_SOLICITUDES import *
from manipular_Hoja1 import *
from openpyxl import *
from funciones_weekend import *

# Cargar el archivo de Excel y seleccionar la hoja

wb = openpyxl.load_workbook('test.xlsx')
ws = wb[wb.sheetnames[1]]

find_table_and_move_to_A5_xlsx(ws)
a5 = ws['A5']
if a5.value < 1000:
    ws.delete_rows(a5.row, 1)
    ws.delete_cols(a5.column, 1)

wb.save('result.xlsx')
abrir_excel('result.xlsx')