import openpyxl
from manipular_INFORME_SOLICITUDES import *
from manipular_Hoja1 import *
from openpyxl import *
# Cargar el archivo de Excel y seleccionar la hoja

wb = openpyxl.load_workbook('test.xlsx')
ws = wb[wb.sheetnames[0]]

delete_filas(ws)
delete_ciudades_columnas(ws)
date_format(ws)
styles_columnSize(ws)
int_format(ws)

ws = wb[wb.sheetnames[1]]
ws2 = wb[wb.sheetnames[0]]

concatenar_nombres_apellidos(ws)
delete_columns(ws)
move_data_to_D5(ws)
encontrar_y_mover_coincidencias_cedulas_y_nombres(ws,ws2)
organizar_tabla_alfabeticamente(ws2)
encontrar_y_mover_coincidencias_nombres(ws,ws2)
no_service_copypaste(ws2,ws)
novedades_expertas(ws2)


wb.save('result.xlsx')
abrir_excel('result.xlsx')