import openpyxl
from manipular_INFORME_SOLICITUDES import *
# Cargar el archivo de Excel y seleccionar la hoja

ordenar_tabla_por_columna_N(ws)
# Guardar el archivo de Excel
wb.save('result.xlsx')          
abrir_excel('result.xlsx')
