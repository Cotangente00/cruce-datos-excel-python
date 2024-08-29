import tkinter as tk
from tkinter import filedialog, Label
from tkinter import messagebox
import os
import openpyxl
from manipular_INFORME_SOLICITUDES import *
from manipular_Hoja1 import *
import xlrd
import logging
import sys
import os
import xlwt
from convert_xlsx_to_xls import *
from funciones_weekend import *
from funciones_botones import *   
import datetime
from intro_botones import *

'''------Función para Lunes a Jueves------'''
def procesar_archivo_excel():

  filepath = filedialog.askopenfilename(title="Selecciona el archivo Excel a modificar", filetypes=[("Archivos Excel", "*.xlsx;*.xls")])
  if not filepath:
    return

  # Obtener la fecha de creación del archivo
  creation_time = os.path.getctime(filepath)
  creation_date = datetime.datetime.fromtimestamp(creation_time)

  # Verificar si el día de la semana es lunes, martes, miercoles y/o jueves
  if creation_date.weekday() in [0,1,2,3]: # Lunes, martes, miercoles y jueves
    intro_function_lunes_jueves(filepath) #ejecutar toda la lógica del botón
  
  # Verificar si el día de la semana es viernes y/o sábado
  elif creation_date.weekday() in [4,5]: # Viernes y sábado
    intro_function_viernes_sabado(filepath) #ejecutar toda la lógica del botón

  else:
    error = messagebox.showerror("ERROR", "Ha seleccionado un archivo que pertenece una fecha equivocada. No es posible modificar este archivo")
    if error:
      return
  
# Configurar la interfaz gráfica
root = tk.Tk()
root.wm_title("Informe Solicitudes y Expertas Disponibles")
root.geometry('420x120')
root.resizable(width=False, height=False)

#Función para que el ícono de la ventana funcione correctamente en cojunto con el comando 
def recurso_path(relative_path):
  try:
    base_path = sys._MEIPASS
  except Exception:
    base_path = os.path.abspath(".")
  return os.path.join(base_path, relative_path)


icon_path = recurso_path('icon.ico')
root.iconbitmap(icon_path)
btn_procesar_archivo_excel = tk.Button(root, text="Procesar Archivo Excel", command=procesar_archivo_excel).pack(pady=20)
Label(root, text="ASEGURESE DE COPIAR LA TABLA DE EXPERTAS SIN ENCABEZADOS").pack()
root.mainloop() 