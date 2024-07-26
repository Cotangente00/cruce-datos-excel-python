from openpyxl.utils.cell import get_column_letter

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







'''------Guardar datos de las columnas antes de eliminarlos------'''
def guardar_datos_columnas(ws, columnas_eliminar):
    datos_columnas = {}
    for col_idx in columnas_eliminar:
        columna_datos = []
        for row in ws.iter_rows(min_row=5, min_col=col_idx, max_col=col_idx):
            columna_datos.append(row[0].value)
        datos_columnas[col_idx] = columna_datos
    return datos_columnas


'''------Eliminar columnas innecesarias------'''
def delete_columns(ws):
    #Columnas a eliminar: (C, F, G, H y I) Fecha, Sexo, localidad, número de celular y TCVA
    columnas_eliminar = [3, 5, 6, 7, 8, 9]

    # Guardar los datos antes de eliminar las columnas
    datos_columnas = guardar_datos_columnas(ws, columnas_eliminar)

    # Eliminar las columnas
    for col_idx in sorted(columnas_eliminar, reverse=True):
        ws.delete_cols(col_idx)

    # Columna a eliminar: C, nuevamente para no dejar espacios vacíos
    columna_eliminar = 3

    # Elimina la columna C (vacía)
    ws.delete_cols(columna_eliminar, 1)  

    return datos_columnas


'''------Restaurar columnas eliminadas (En caso de que el usuario quiera revertir los cambios)------'''
def restaurar_columnas(ws, datos_columnas):
    columnas_a_restaurar = sorted(datos_columnas.keys(), reverse=True)

    for col_idx in columnas_a_restaurar:
        # Insertar columna vacía en la posición correcta
        ws.insert_cols(col_idx)
        
        # Recuperar los datos de la columna
        datos = datos_columnas[col_idx]
        
        # Rellenar la columna con los datos recuperados
        for row_idx, valor in enumerate(datos, start=5):
            ws.cell(row=row_idx, column=col_idx, value=valor)








'''------Guadar datos antes de moverlos a la celda D5, (en caso de que el usuario quiera revertir los cambios)------'''
def guardar_datos_originales(ws):
    datos_originales = {
        'A': [],
        'B': [],
        'C': []
    }
    for fila in ws.iter_rows(min_row=5, max_col=3, values_only=True):
        datos_originales['A'].append(fila[0])
        datos_originales['B'].append(fila[1])
        datos_originales['C'].append(fila[2])
    return datos_originales


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
    for fila in ws.iter_rows(min_row=5, values_only=True): # Se empieza a partir de la fila 5
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


'''------Restaurar datos a la celda original A5 (en caso de que el usuario haya tomado la desición de revertir los cambios)------'''
def restaurar_datos(ws, datos_originales):
    for i, (cedula, profesional, hogares_pymes_cuidadoras) in enumerate(zip(
            datos_originales['A'],
            datos_originales['B'],
            datos_originales['C']
        ), start=5):
        ws[f'A{i}'] = cedula
        ws[f'B{i}'] = profesional
        ws[f'C{i}'] = hogares_pymes_cuidadoras

    for col in ['D', 'E', 'F']:
        for fila in ws.iter_rows(min_row=5, max_col=ws.max_column):
            ws[f'{col}{fila[0].row}'].value = None

    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 45
    ws.column_dimensions['F'].width = 18
    ws.column_dimensions['H'].width = 45