import openpyxl
from manipular_INFORME_SOLICITUDES import *
from manipular_Hoja1 import *
from openpyxl import *
from funciones_weekend import *

# Cargar el archivo de Excel y seleccionar la hoja

wb = openpyxl.load_workbook('test.xlsx')
ws = wb[wb.sheetnames[1]]

inicio_fila = None
inicio_columna = None
for fila in ws.iter_rows(min_row=1, max_col=ws.max_column):
    for cell in fila:
        if cell.value is not None:
            inicio_fila = cell.row
            inicio_columna = cell.column
            if inicio_fila < 1000:
                ws.delete_rows(inicio_fila, 1)  
                ws.delete_cols(inicio_columna, 1)  
            break
    if inicio_fila:        
        break   

wb.save('result.xlsx')
abrir_excel('result.xlsx')