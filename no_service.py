'''------Sección para copiar  y pegar todas las expertas que NO tienen servicio------'''
def no_service_copypaste(ws,ws2):
    # Listas para almacenar cédulas y nombres completos
    cedulas_sin_servicio = []
    nombres_sin_servicio = []

    # Iterar sobre las filas de la Hoja2
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