from tkinter import messagebox

def button_1st_part(wb):
    # Seleccionar la hoja de trabajo
    ws = wb.active


    # Mensaje de confirmación para que el usuario sea consciente que de el listado de expertas fue copiado en la celda A5
    confirmacion = messagebox.askyesno('Confirmar modificación', 'Asegurese de que el listado de expertas haya sido copiado en la celda A5 de la hoja "Hoja1", para evitar dañar el contenido del archivo. ¿Desea continuar?')
    if not confirmacion:
      return
            
    #buscar el marcador de modificación 
    marcador = ws['AZ1'] #Marcador en la celda AZ1
    modificado_antes = marcador.value == 'MODIFICADO' 

    if modificado_antes:
      # Mensaje de confirmación si el archivo ha sido modificado por la aplicación previamente  
      resultado = messagebox.askyesno('Confirmar modificación', 'Este archivo ya ha sido modificado previamente, si continúa, el contenido del archivo será distorsionado. ¿Desea continuar?')
      if not resultado:
        return


def button_2nd_part(wb,ws2):
    ws2['R2'] = 'Expertas que NO tienen servicio' 

    #Agregar marcador de que el archivo ha sido modificado por la aplicación en la celdd AZ1        
    ws2['AZ1'] = 'MODIFICADO' 

    # Pedir al usuario la ruta y nombre para guardar el nuevo archivo
    filepath_save = filedialog.asksaveasfilename(title="Guardar archivo Excel modificado como", defaultextension=".xlsx", filetypes=[("Archivos Excel", "*.xlsx")])
    if not filepath_save:
      return  # Salir si el usuario cancela el diálogo de guardado

    # Guardar el archivo modificado
    wb.save(filepath_save)
    messagebox.showinfo("Proceso completado", """
    • Cambios en la hoja "INFORME SOLICITUDES":
    - Se eliminaron las filas 1, 2, 3, y 4.
    - Total servicios, Tipo, Turno partido, Jornada fija, 
      Concepto: novedad y ausencias, Concepto: novedad 
      control empleados, CC experta cambios, Experta cambio,
      Notificación SMS, Notificación SMS cliente, SMS 
      enviado cliente. Columnas eliminadas.
    - Formato de fecha conservado DD/MM/YYYY.
    - Tamaño horizontal de las columnas adaptado.
    - Solicitud, Referencia Externa y Cédula modificada a 
      formato numérico.
    - Expertas que SI tienen novedad, resaltadas con color 
      amarillo.
    • Cambios en la hoja "Hoja1":
    - Nombres y apellidos concatendados.
    - Fecha, Sexo, localidad, número de celular y 
      TCVA eliminadas.
    - Datos trasladados a la celda D5.
    - BUSCARV desde Hoja1 a INFORME SOLICITUDES 
      número de documento y nombre completo en 
      las columnas M y N.
    - BUSCARV desde INFORME SOLICITUDES a Hoja1 
      nombre completo en la columna H.
    - Listado de expertas sin servicio copiado en las 
      columnas Q, R y S.
    """)
    abrir_excel(filepath_save)