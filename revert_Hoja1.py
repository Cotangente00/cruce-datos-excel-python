'''------Deshacer concatenación de nombres y apellidos------'''

def deshacer_concatenacion(ws):
    # Iterar sobre las filas
    for row in ws.iter_rows(min_row=5, min_col=2, max_col=2):  # Solo la columna B (concatenado)
        # Obtener el valor concatenado de la columna B
        valor_b = row[0].value
        
        if valor_b:  # Verificar que no esté vacío
            # Separar el valor concatenado en nombre y apellido
            partes = valor_b.split(' ', 1)  # Divide en dos partes: nombre y apellido
            
            if len(partes) == 2:  # Asegurarse de que se obtuvo nombre y apellido
                nombre, apellido = partes
            else:
                nombre = partes[0]
                apellido = ""  # Si solo hay una parte, el apellido se deja vacío
            
            # Colocar el nombre y apellido en las columnas correspondientes
            ws.cell(row=row[0].row, column=2, value=nombre)  # Columna B
            ws.cell(row=row[0].row, column=3, value=apellido)  # Columna C

    # Eliminar la columna B (concatenado)
    columna_eliminar = 2
    ws.delete_cols(columna_eliminar, 1)

#-------------------------------------------------------------------------------------------------------
#La función de restaurar columnas eliminadas, ya está definida en el archivo "manipular_Hoja1.py"
#-------------------------------------------------------------------------------------------------------


#----------------------------------------------------------------------------------------------------------------------------
#La función para deshacer los cambios de la función: "move_data_to_D5" del archivo "manipular_Hoja1.py", ya fue definida
#----------------------------------------------------------------------------------------------------------------------------