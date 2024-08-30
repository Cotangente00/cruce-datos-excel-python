import openpyxl
from manipular_INFORME_SOLICITUDES import *
from manipular_Hoja1 import *
from openpyxl import *
from funciones_weekend import *

# Cargar el archivo de Excel y seleccionar la hoja

wb = openpyxl.load_workbook('test.xlsx')
ws = wb[wb.sheetnames[0]]
max_width = 200

for column in ws.columns:
        max_length = 0
        column = column[0:]  
        for cell in column:
            try:  # Manejar posibles errores de tipo
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass

        adjusted_width = min(max_length + 4, max_width) if max_width else max_length + 4
        # Obtener la letra de la columna a partir del objeto columna
        column_letter = column[0].column_letter
        # Ajustar el ancho de la columna usando el nombre de la columna
        ws.column_dimensions[column_letter].width = adjusted_width

wb.save('result.xlsx')
abrir_excel('result.xlsx')