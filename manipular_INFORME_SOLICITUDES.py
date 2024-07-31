import openpyxl
from openpyxl.styles import NamedStyle
from openpyxl.styles import Font
from openpyxl.styles import PatternFill
import os


'''-------Sección para elminar filas innecesarias-------'''
def delete_filas(ws):
    #filas innecesarias a eliminar
    filas_eliminar = [1,2,3,4]

    # Iterar sobre las filas a eliminar
    for row_idx in sorted(filas_eliminar, reverse=True):
        # Eliminar la fila
        ws.delete_rows(row_idx) 


'''------Sección para eliminar columnas y ciudades innecesarias------'''
def delete_ciudades_columnas(ws):
    #columnas innecesarias a eliminar
    columnas_eliminar = ['C', 'D', 'J', 'K', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

    # Eliminar las columnas
    for col in reversed(columnas_eliminar):
        # Convertir la letra de la columna a índice numérico
        col_idx = openpyxl.utils.cell.column_index_from_string(col)
        # Eliminar la columna utilizando el índice
        ws.delete_cols(col_idx)

    ciudades_eliminar = ['Medellin', 'Cali', 'BUCARAMANGA', 'Medellín', 'PEREIRA', 'Barranquilla', 'BARRANQUILLA']
    # Iterar sobre las filas y eliminar las que contienen las ciudades a eliminar
    filas_a_eliminar = []
    for row in ws.iter_rows(min_row=2):  # Empezamos desde la segunda fila (ya que en la primeera fila están los encabezados)
        ciudad = row[12].value  #ubicamos la columna de las ciudades M (índice 12)
        if ciudad in ciudades_eliminar:
            filas_a_eliminar.append(row)

    # Eliminar las filas identificadas
    for row in filas_a_eliminar:
        ws.delete_rows(row[0].row)  # row[0] es la celda en la primera columna (A) que contiene el número de fila

    # Columna a eliminar: ciudades una vez descartadas las que NO se necesitan
    columna_eliminar = 13

    # Elimina la columna M (ciudades) 
    ws.delete_cols(columna_eliminar, 1)  
    


'''------Sección para conservar el formato de fecha corta------'''
def date_format(ws):
    # Configurar el estilo de formato de fecha corta
    date_style = NamedStyle(name='date_style', number_format='DD/MM/YYYY')

    # Aplicar el estilo a la columna de fechas (en este caso, columna D)
    for cell in ws['D']:  # Ajustar 'D' a la letra de la columna de fechas
        cell.style = date_style



'''------Sección para modificar el tamaño de las columnas y el aspecto------'''
def styles_columnSize(ws):
    ws.column_dimensions['A'].width = 13
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 13
    ws.column_dimensions['D'].width = 13
    ws.column_dimensions['E'].width = 55
    ws.column_dimensions['H'].width = 15
    ws.column_dimensions['F'].width = 50
    ws.column_dimensions['I'].width = 55
    ws.column_dimensions['J'].width = 22
    ws.column_dimensions['K'].width = 45
    ws.column_dimensions['M'].width = 18
    ws.column_dimensions['L'].width = 12
    ws.column_dimensions['N'].width = 45
    ws.row_dimensions[1].height= 20

    fila = 1
    fila_excel = ws[f'A{fila}:Z{fila}']  # Establecemos un rango de columnas A a Z para la fila específica

    # Establecer el estilo de la fuente para negrita y subrayado
    estilo_negrita_subrayado = Font(bold=True, underline='single')

    # Aplicar el estilo a cada celda en la fila
    for fila_celdas in fila_excel:
        for celda in fila_celdas:
            celda.font = estilo_negrita_subrayado



'''------Sección para pasar todos lo números que están almacenados como texto a formato número------'''
def int_format(ws):
    columnas_a_convertir = ['A', 'B', 'J']

    # Iterar sobre las columnas a convertir
    for col in columnas_a_convertir:
        for cell in ws[col]:  # Iterar sobre todas las celdas de la columna
            # Convertir el valor de texto a número si es posible
            try:
                valor = float(cell.value)
                cell.value = valor  # Actualizar el valor en la celda con el número convertido
            except (ValueError, TypeError):
                # Manejar el caso donde no se puede convertir a número
                continue 



'''------Sección para resaltar a todas las que tienen novedades (nombre y cédula) y crear una hoja nueva------'''
def novedades_expertas(ws):
    # Color de fondo amarillo
    relleno_amarillo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # Recorrer las filas (empezando desde la segunda fila por los encabezados)
    for fila in ws.iter_rows(min_row=2, min_col=10, max_col=11):  # Columnas J (10) y K (11)
        novedad = ws.cell(row=fila[0].row, column=13).value  # Columna M (13)

        # Verificar si la columna "novedad" es "Si"
        if novedad == "Si":
            # Aplicar formato de color amarillo a las celdas de nombre (J) y cedula (K)
            for celda in fila:
                celda.fill = relleno_amarillo

    # Columna a eliminar: Tiene_novedad una vez resaltadas las que SI tienen novedades
    columna_eliminar = 13

    # Elimina la columna M (Tiene novedad) 
    ws.delete_cols(columna_eliminar, 1)  




'''------Ordenar una tabla alfabéticamente, usando una columna como índice------'''
def a_z(ws):
    data = [] #Tupla para almacernar los datos 
    for row in ws.iter_rows(min_row=2, values_only=True): #se empieza desde la fila dos por los encabezados
        data.append((row[10], *row)) # row[10] columna K

    #Se ordenan la lista de tuplas por el elemento o la columna K 
    data.sort()

    #Se escriben los datos ya ordenados alfabéticamente
    for row_index, row_data in enumerate(data, start=2):
        for column_index, cell_value in enumerate(row_data[1:], start=1):
            ws.cell(row=row_index, column=column_index, value=cell_value)




'''------Abrir archivo Excel automáticamente una vez hechos los cambios------'''
def abrir_excel(filepath):
    try:
        os.startfile(filepath)
    except OSError as e:
        print(f"No se pudo abrir el archivo '{filepath}': {e}")