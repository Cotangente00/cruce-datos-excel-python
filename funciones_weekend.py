#En este archivo están todas las funciones que se ejecutan para los avances de viernes a sabado
from openpyxl.utils.cell import get_column_letter
import xlrd
import xlwt
from xlutils.copy import copy
from manipular_Hoja1 import *


'''------Eliminar fecha y el resto de columnas innecesarias------'''
def delete_columns_viernes_sabado(ws):
    #Columnas a eliminar: (C, F, G, H y I) Fecha, Sexo, localidad, número de celular y TCVA
    columnas_eliminar = [3, 5, 6, 7, 8, 9, 10]
    # Eliminar las columnas
    for col_idx in sorted(columnas_eliminar, reverse=True):
        # Iterar en orden inverso para evitar problemas con los índices cambiando
        col_letra = get_column_letter(col_idx)
        ws.delete_cols(col_idx)

    # Columna a eliminar: C, nuevamente para no dejar espacios vacíos
    columna_eliminar = 3

    # Elimina la columna C (vacía)
    ws.delete_cols(columna_eliminar, 1)  



'''------Mover datos a D5------'''
def move_data_to_D5_viernes_sabado(ws):
    ws.column_dimensions['E'].width = 45
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['F'].width = 18
    ws.column_dimensions['H'].width = 45

    cedula_copiar = []
    profesional_copiar = []
    hogares_pymes_cuidadoras_copiar = []
    horas_copiar = []

    # Iterar sobre las filas de la Hoja1
    for fila in ws.iter_rows(min_row=5, values_only=True): # Se empieza desde la celda A5
        cedula = fila[1 - 1] # columna A
        profesional = fila[2 - 1] # columna B
        hogares_pymes_cuidadoras = fila[3 - 1] # columna C
        horas = fila[4 - 1] # columna D


        cedula_copiar.append(cedula)
        profesional_copiar.append(profesional)
        hogares_pymes_cuidadoras_copiar.append(hogares_pymes_cuidadoras) 
        horas_copiar.append(horas)

        if cedula is None:  #para el ciclo cuando detexte un campo vacío
            break

    # Pegar los datos en las celdas D, E, F y G
    for i, (cedula, profesional, hogares_pymes_cuidadoras, horas) in enumerate(zip(cedula_copiar, profesional_copiar, hogares_pymes_cuidadoras_copiar, horas_copiar), start=5):
        ws[f'D{i}'] = cedula
        ws[f'E{i}'] = profesional
        ws[f'F{i}'] = hogares_pymes_cuidadoras
        ws[f'G{i}'] = horas


    # Eliminar contenido de la columna A
    for fila in ws.iter_rows():
        celda = fila[0]  # Celda en la columna A
        celda.value = None  # Eliminar valor de la celda

    # Eliminar contenido de la columna B 
    for fila in ws.iter_rows():
        celda = fila[1]
        celda.value = None

    # Eliminar contenido de la columna C
    for fila in ws.iter_rows():
        celda = fila[2]
        celda.value = None


'''------Sección para copiar  y pegar todas las expertas que NO tienen servicio------'''
def no_service_copypaste_viernes_sabado(ws,ws2):
    # Listas para almacenar cédulas y nombres completos
    cedulas_sin_servicio = []
    nombres_sin_servicio = []
    tipo_sin_servicio = []

    cedulas_sin_servicio_horas = []
    nombres_sin_servicio_horas = []
    tipo_sin_servicio_horas = []

    # Iterar sobre las filas de la Hoja1
    for fila in ws2.iter_rows(min_row=2, values_only=True):  
        cedula = fila[4 - 1]  # Columna D, Python usa base cero
        nombre = fila[5 - 1]  # Columna E
        tipo = fila[6 - 1]  #Columna F 
        horas = fila[7 - 1]  #Columna G
        servicio = fila[8 - 1]  # Columna H


        #Si el servicio (columna H) está vacío se almacenan los datos en las tuplas anteriormente definidas 
        if servicio is None or servicio.strip() == "":
            cedulas_sin_servicio.append(cedula)
            nombres_sin_servicio.append(nombre)
            tipo_sin_servicio.append(tipo)
            if horas == 120 or horas == 240 or horas == 235:
                cedulas_sin_servicio_horas.append(cedula)
                nombres_sin_servicio_horas.append(nombre)
                tipo_sin_servicio_horas.append(tipo)


    # Pegar los datos en la Hoja1, columnas Q, R y S
    for i, (cedula, nombre, tipo) in enumerate(zip(cedulas_sin_servicio_horas, nombres_sin_servicio_horas, tipo_sin_servicio_horas), start=4):
        ws[f'R{i}'] = cedula
        ws[f'S{i}'] = nombre
        ws[f'T{i}'] = tipo


'''------Función que globaliza todas las funciones anteriormenete definidas (VIERNES-SÁBADO)-------'''
def ejecucion_funciones2_viernes_sabado(ws,ws2):
    concatenar_nombres_apellidos(ws)
    delete_columns_viernes_sabado(ws)
    move_data_to_D5_viernes_sabado(ws)
    encontrar_y_mover_coincidencias_cedulas_y_nombres(ws,ws2)
    encontrar_y_mover_coincidencias_nombres(ws,ws2)
    no_service_copypaste_viernes_sabado(ws2,ws)  #argumentos de hojas invertidos para mayor comodidad (originalmente ws es INFORME SOLICITUDES y ws2 es Hoja1)