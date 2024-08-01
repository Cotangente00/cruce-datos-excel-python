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

        # Aplicar las modificaciones utilizando Openpyxl
        delete_filas(ws)
        delete_ciudades_columnas(ws)
        date_format(ws)
        styles_columnSize(ws)
        int_format(ws)
        novedades_expertas(ws)
        a_z(ws)

         # Crea una nueva hoja llamada "Hoja1"
        hoja_nueva = wb.create_sheet("Hoja1")

        hoja_nueva['A1'] = 'Pegar datos de expertas en la celda A5'

        # Pedir al usuario la ruta y nombre para guardar el nuevo archivo
        filepath_save = filedialog.asksaveasfilename(title="Guardar archivo Excel modificado como", defaultextension=".xlsx", filetypes=[("Archivos Excel", "*.xlsx")])
        if not filepath_save:
            return  # Salir si el usuario cancela el diálogo de guardado

        # Guardar el archivo modificado
        wb.save(filepath_save)
        messagebox.showinfo("Proceso completado", """
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
        - Tabla ordenada alfabéticamente (con la colummna K 
          como base).
        """)
        abrir_excel(filepath_save)
    except Exception as e:
        messagebox.showerror("Error", f"Ha ocurrido un error al procesar el archivo:\n{str(e)}")

# Configurar el logging para errores
logging.basicConfig(filename='app.log', level=logging.ERROR)


def procesar_Hoja1():
    # Abrir el archivo Excel seleccionado
    filepath = filedialog.askopenfilename(title="Selecciona el archivo Excel a modificar", filetypes=[("Archivos Excel", "*.xlsx")])
    if not filepath:
        return

    try:
        # Cargar el libro de Excel
        wb = openpyxl.load_workbook(filepath)
        # Seleccionar la hoja de trabajo
        ws = wb['Hoja1']
        ws2 = wb['INFORME SOLICITUDES']

        #buscar el marcador de modificación 
        marcador = ws['AZ1'] #Marcador en la celda AD1
        modificado_antes = marcador.value == 'MODIFICADO' 

        if modificado_antes:
            # Mensaje de confirmación si el archivo ha sido modificado por la aplicación previamente  
            resultado = messagebox.askyesno('Confirmar modificación', 'Este archivo ya ha sido modificado previamente, si continúa, el contenido del archivo será distorsionado. ¿Desea continuar?')
            if not resultado:
                return
        

        # Aplicar las modificaciones utilizando Openpyxl
        # Aquí se coloca la lógica para modificar los datos del archivo Excel
        concatenar_nombres_apellidos(ws)
        delete_columns(ws)   
        move_data_to_D5(ws)
        encontrar_y_mover_coincidencias_cedulas_y_nombres(ws2,ws) #argumentos de hojas invertidos para mayor comodidad (originalmente ws es INFORME SOLICITUDES y ws2 es Hoja1)
        encontrar_y_mover_coincidencias_nombres(ws,ws2) #argumentos de hojas invertidos para mayor comodidad (originalmente ws es INFORME SOLICITUDES y ws2 es Hoja1)
        no_service_copypaste(ws2,ws) #argumentos de hojas invertidos para mayor comodidad (originalmente ws es INFORME SOLICITUDES y ws2 es Hoja1)

        ws2['Q2'] = 'Expertas que NO tienen servicio'  

        #Agregar marcador de que el archivo ha sido modificado por la aplicación
        ws['AZ1'] = 'MODIFICADO'

        # Reescribir o guardar los cambios en el mismo archivo modificado 
        wb.save(filepath)
        messagebox.showinfo("Proceso completado", """
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
          columnas Q y R.
        """)
        abrir_excel(filepath)
    except PermissionError:
        messagebox.showerror("Error de Permiso", f"No se puede modificar el archivo '{filepath}'. Asegúrate de que el archivo esté cerrado y que no esté siendo utilizado por otro programa.")
        logging.error(f"Error al abrir el archivo {filepath}: Permission denied")
        return 
    except Exception as e:
        messagebox.showerror("Error", f"Ha ocurrido un error al procesar el archivo:\n{str(e)}")
        return


# Configurar la interfaz gráfica
root = tk.Tk()
root.wm_title("Informe Solicitudes y Expertas Disponibles")
root.geometry('420x180')
root.resizable(width=False, height=False)

def recurso_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


icon_path = recurso_path('icon.ico')
root.iconbitmap(icon_path)

btn_procesar_informe_solicitudes = tk.Button(root, text="1. Procesar INFORME SOLICITUDES", command=procesar_INFORME_SOLICITUDES)
btn_procesar_informe_solicitudes.pack(pady=20)

btn_procesar_hoja1 = tk.Button(root, text="2. Procesar Hoja1", command=procesar_Hoja1)
btn_procesar_hoja1.pack(pady=20)

root.mainloop() 