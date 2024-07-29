from openpyxl.utils.cell import get_column_letter
import openpyxl

'''------Concatenar nombres y apellidos------'''
def concatenar_nombres_apellidos(ws):
    # Iterar sobre las filas
    for row in ws.iter_rows(min_row=5, min_col=2, max_col=3):  # empieza desde la quinta fila (min_row=5) A5
        # Obtener los valores de las columnas B y C
        valor_b = row[0].value
        valor_c = row[1].value
        
        # Concatenar los valores de B y C y almacenar el resultado en la columna B
        concatenated_value = f"{valor_b} {valor_c}"
        row[0].value = concatenated_value

    # Columna a eliminar: C, una vez concatenados los nombres con los apellidos
    columna_eliminar = 3

    # Elimina la columna C (Apellidos) 
    ws.delete_cols(columna_eliminar, 1)  


'''------Eliminar fecha y el resto de columnas innecesarias------'''
def delete_columns(ws):
    #Columnas a eliminar: (C, F, G, H y I) Fecha, Sexo, localidad, número de celular y TCVA
    columnas_eliminar = [3, 5, 6, 7, 8, 9]
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
def move_data_to_D5(ws):
    ws.column_dimensions['E'].width = 45
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['F'].width = 18
    ws.column_dimensions['H'].width = 45

    cedula_copiar = []
    profesional_copiar = []
    hogares_pymes_cuidadoras_copiar = []

    # Iterar sobre las filas de la Hoja1
    for fila in ws.iter_rows(min_row=5, values_only=True): # Se empieza desde la celda A5
        cedula = fila[1 - 1] # columna A
        profesional = fila[2 - 1] # columna B
        hogares_pymes_cuidadoras = fila[3 - 1] # columna C

        cedula_copiar.append(cedula)
        profesional_copiar.append(profesional)
        hogares_pymes_cuidadoras_copiar.append(hogares_pymes_cuidadoras) 

        if cedula is None:
            break

    # Pegar los datos en las celdas D, E y F, columnas Q y R
    for i, (cedula, profesional, hogares_pymes_cuidadoras) in enumerate(zip(cedula_copiar, profesional_copiar, hogares_pymes_cuidadoras_copiar), start=5):
        ws[f'D{i}'] = cedula
        ws[f'E{i}'] = profesional
        ws[f'F{i}'] = hogares_pymes_cuidadoras

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




'''------Sección para buscar coincidencias de la hoja INFORME SOLICITUDES------'''
def encontrar_y_mover_coincidencias_cedulas_y_nombres(ws,ws2):
    # Obtener los números de cédula y los nombres de la hoja Hoja1
    cedulas_hoja1 = {}
    for i in range(2, ws2.max_row + 1):
        cedula = str(ws2[f'D{i}'].value)
        nombre_apellido = ws2[f'E{i}'].value
        cedulas_hoja1[cedula] = nombre_apellido

    # Obtener los números de cédula de la columna J de INFORME SOLICITUDES
    cedulas_informe = {str(ws[f'J{i}'].value) for i in range(2, ws.max_row + 1)}

    # Insertar las coincidencias en la columna N y O de INFORME SOLICITUDES
    fila = 2
    for celda in ws['J']:
        cedula = str(celda.value)
        if cedula in cedulas_hoja1:
            ws[f'M{celda.row}'] = float(cedula)
            ws[f'N{celda.row}'] = cedulas_hoja1[cedula]

    #wb.save('Avance.xlsx')



'''------Sección para buscar coincidencias de la hoja Hoja1------'''
def encontrar_y_mover_coincidencias_nombres(ws2,ws):
    # Obtener los números de cédula y los nombres de la hoja INFORME SOLICITUDES
    cedulas_INFORME_SOLICITUDES = {}
    for i in range(2, ws.max_row + 1):
        cedula = str(ws[f'J{i}'].value)
        nombre_apellido = ws[f'K{i}'].value
        cedulas_INFORME_SOLICITUDES[cedula] = nombre_apellido

    # Obtener los números de cédula de la columna D de Hoja1
    cedulas_Hoja1 = {str(ws2[f'D{i}'].value) for i in range(2, ws.max_row + 1)}

    # Insertar las coincidencias en la columna H de Hoja1
    fila = 5
    for celda in ws2['D']:
        cedula = str(celda.value)
        if cedula in cedulas_INFORME_SOLICITUDES:
            ws2[f'H{celda.row}'] = cedulas_INFORME_SOLICITUDES[cedula]



'''------Sección para copiar  y pegar todas las expertas que NO tienen servicio------'''
def no_service_copypaste(ws,ws2):
    # Listas para almacenar cédulas y nombres completos
    cedulas_sin_servicio = []
    nombres_sin_servicio = []

    # Iterar sobre las filas de la Hoja1
    for fila in ws2.iter_rows(min_row=2, values_only=True):  
        cedula = fila[4 - 1]  # Columna D, Python usa base cero
        nombre = fila[5 - 1]  # Columna E, Python usa base cero
        servicio = fila[8 - 1]  # Columna H, Python usa base cero

        if servicio is None or servicio.strip() == "":
            cedulas_sin_servicio.append(cedula)
            nombres_sin_servicio.append(nombre)

    # Pegar los datos en la Hoja1, columnas Q y R
    for i, (cedula, nombre) in enumerate(zip(cedulas_sin_servicio, nombres_sin_servicio), start=1):
        ws[f'Q{i}'] = cedula
        ws[f'R{i}'] = nombre



'''
# Cargar el libro de Excel
wb = openpyxl.load_workbook('Avance.xlsx')

# Seleccionar las hojas
ws = wb["INFORME SOLICITUDES"]
ws2 = wb["Hoja1"]

encontrar_y_mover_coincidencias_cedulas(ws,ws2)
encontrar_y_mover_coincidencias_nombres(ws,ws2)

'''