import openpyxl
from manipular_INFORME_SOLICITUDES import *
from manipular_Hoja1 import *
from openpyxl import *
from funciones_weekend import *

# Cargar el archivo de Excel y seleccionar la hoja

wb = openpyxl.load_workbook('test.xlsx')
#ws = wb[wb.sheetnames[0]]

#ejecucion_funciones(ws)


#ws = wb[wb.sheetnames[1]]
ws2 = wb[wb.sheetnames[0]]

#ejecucion_funciones2_viernes_sabado(ws,ws2)

for cell in column:
    if cell.value:
      cell.font = cell.font.copy(bold=True)
      cell.style = 'Currency'  # Aplicar formato de moneda para subrayado


wb.save('result.xlsx')
abrir_excel('result.xlsx')