import openpyxl
from openpyxl import *
from manipular_Hoja1 import *
from manipular_INFORME_SOLICITUDES import *

wb = openpyxl.load_workbook('test.xlsx')
ws = wb['INFORME SOLICITUDES']

ejecucion_funciones(ws)

ws=wb['Hoja1']
ws2=wb['INFORME SOLICITUDES']

find_table_and_move_to_A5(ws)
concatenar_nombres_apellidos(ws)
delete_columns(ws)
move_data_to_D5(ws)
encontrar_y_mover_coincidencias_cedulas_y_nombres_copy(ws,ws2)
encontrar_y_mover_coincidencias_nombres_copy(ws,ws2)
no_service_copypaste(ws2,ws)
                

wb.save('result.xlsx')
abrir_excel('result.xlsx')


def encontrar_y_mover_coincidencias_nombres_copy(ws,ws2):

    # Iterar sobre las filas de la hoja Hoja1, comenzando desde la fila 5 
    for fila_hoja1 in ws.iter_rows(min_row=5, min_col=1, max_col=10):
        numero_documento_hoja1 = fila_hoja1[3].value  # Columna D (índice 3)

        # Iterar sobre las filas de la hoja INFORME SOLCITUDES, comenzando desde la fila 2 
        for fila_informe in ws2.iter_rows(min_row=2, max_col=ws2.max_column):
            numero_documento_informe = fila_informe[9].value  # Columna J (índice 9)

            # Si se encuentra una coincidencia, copiar los datos a la hoja de informe
            if numero_documento_hoja1 == numero_documento_informe:
                fila_hoja1[7].value = fila_informe[10].value  
                

def encontrar_y_mover_coincidencias_cedulas_y_nombres_copy(ws,ws2):
    # Iterar sobre las filas de la hoja de informe, comenzando desde la fila 2 (excluyendo el encabezado)
    for fila_informe in ws2.iter_rows(min_row=2, max_col=ws2.max_column):
        numero_documento_informe = fila_informe[9].value  # Columna J (índice 9)

        # Iterar sobre las filas de la hoja Hoja1, comenzando desde la fila 5 (excluyendo el encabezado)
        for fila_hoja1 in ws.iter_rows(min_row=5, max_col=ws.max_column):
            numero_documento_hoja1 = fila_hoja1[3].value  # Columna D (índice 3)

            # Si se encuentra una coincidencia, copiar los datos a la hoja de informe
            if numero_documento_informe == numero_documento_hoja1:
                fila_informe[12].value = fila_hoja1[3].value  # Columna M (índice 12)
                fila_informe[13].value = fila_hoja1[4].value  # Columna N (índice 13)