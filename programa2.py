import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os
import openpyxl
from manipular_INFORME_SOLICITUDES import *
from manipular_Hoja1 import *
import xlrd
import logging
import sys

def procesar_INFORME_SOLICITUDES():
    filepath = filedialog.askopenfilename(title="Selecciona el archivo Excel a modificar", filetypes=[("Archivos Excel", "*.xlsx;*.xls")])
    if not filepath:
        return

    try:
        if filepath.endswith('.xlsx'):
            # Cargar el libro de Excel (xlsx)
            wb = openpyxl.load_workbook(filepath)
        elif filepath.endswith('.xls'):
            # Cargar el libro de Excel (xls)
            xls_workbook = xlrd.open_workbook(filepath)
            wb = openpyxl.Workbook()

            # Copiar hojas de xls a xlsx
            for sheet in xls_workbook.sheets():
                ws_xls = xls_workbook.sheet_by_name(sheet.name)
                ws_new = wb.create_sheet(sheet.name)

                for row in range(ws_xls.nrows):
                    for col in range(ws_xls.ncols):
                        cell_value = ws_xls.cell_value(row, col)
                        ws_new.cell(row=row+1, column=col+1, value=cell_value)

            # Eliminar la hoja predeterminada de Workbook
            wb.remove(wb.active)

        # Seleccionar la hoja de trabajo
        ws = wb.active

        # Función compuesta de funciones que ejecutan los cambios necesarios en la hoja INFORME SOLICITUDES
        ejecucion_funciones(ws)

        #Volver a cargar las hojas del archivo después de ejecurtar los cambios en la hoja INFORME SOLICITUDES
        ws=wb['Hoja1']
        ws2=wb['INFORME SOLICITUDES']

        # Funciones que hacen los cambios en la Hoja1
        concatenar_nombres_apellidos(ws)
        delete_columns(ws)   
        move_data_to_D5(ws)
        encontrar_y_mover_coincidencias_cedulas_y_nombres(ws2,ws) #argumentos de hojas invertidos para mayor comodidad (originalmente ws es INFORME SOLICITUDES y ws2 es Hoja1)
        encontrar_y_mover_coincidencias_nombres(ws,ws2) #argumentos de hojas invertidos para mayor comodidad (originalmente ws es INFORME SOLICITUDES y ws2 es Hoja1)
        no_service_copypaste(ws2,ws) #argumentos de hojas invertidos para mayor comodidad (originalmente ws es INFORME SOLICITUDES y ws2 es Hoja1)


        ws2['Q2'] = 'Expertas que NO tienen servicio'  

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
    except Exception as e:
        messagebox.showerror("Error", f"Ha ocurrido un error al procesar el archivo:\n{str(e)}")


# Configurar la interfaz gráfica
root = tk.Tk()
root.wm_title("Informe Solicitudes y Expertas Disponibles")
root.geometry('420x80')
root.resizable(width=False, height=False)

def recurso_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


icon_path = recurso_path('icon.ico')
root.iconbitmap(icon_path)

btn_procesar_informe_solicitudes = tk.Button(root, text="1. Procesar Archivo Excel", command=procesar_INFORME_SOLICITUDES)
btn_procesar_informe_solicitudes.pack(pady=20)

root.mainloop() 