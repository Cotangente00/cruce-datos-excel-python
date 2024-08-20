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

'''------Función para Lunes a Jueves------'''
def procesar_archivo_excel():
  filepath = filedialog.askopenfilename(title="Selecciona el archivo Excel a modificar", filetypes=[("Archivos Excel", "*.xlsx;*.xls")])
  if not filepath:
    return

  try:

    if filepath.endswith('.xlsx'):
      #convertir de xlsx a xls para convertirlo de vuelta
      convert_xlsx_to_xls(filepath, 'temp.xls')

      # Cargar el libro de Excel (xls)
      xls_workbook = xlrd.open_workbook('temp.xls')
      wb = openpyxl.Workbook()

      # Copiar hojas de xls a xlsx
      for sheet in xls_workbook.sheets():
        ws_xls = xls_workbook.sheet_by_name(sheet.name)
        ws_new = wb.create_sheet(sheet.name)

        for row in range(ws_xls.nrows):
          for col in range(ws_xls.ncols):
            cell_value = ws_xls.cell_value(row, col)
            ws_new.cell(row=row+1, column=col+1, value=cell_value)
            
      wb.remove(wb.active)

      os.remove('temp.xls')

    elif filepath.endswith('.xls'):

      find_table_and_move_to_A5_xls(filepath, 'temp.xls')
      # Cargar el libro de Excel (xls)
      xls_workbook = xlrd.open_workbook('temp.xls')
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

      os.remove('temp.xls')
      

    button_lunes_jueves(wb)

  except Exception as e:
    messagebox.showerror("Error", f"Ha ocurrido un error al procesar el archivo:\n{str(e)}")



'''------Función para Viernes a Sábado------'''
def procesar_archivo_excel_viernes_sabado():
  filepath = filedialog.askopenfilename(title="Selecciona el archivo Excel a modificar", filetypes=[("Archivos Excel", "*.xlsx;*.xls")])
  if not filepath:
    return

  try:

    if filepath.endswith('.xlsx'):

      #convertir de xlsx a xls para convertirlo de vuelta
      convert_xlsx_to_xls_viernes_sabado(filepath, 'temp.xls')

      # Cargar el libro de Excel (xls)
      xls_workbook = xlrd.open_workbook('temp.xls')
      wb = openpyxl.Workbook()

      # Copiar hojas de xls a xlsx
      for sheet in xls_workbook.sheets():
        ws_xls = xls_workbook.sheet_by_name(sheet.name)
        ws_new = wb.create_sheet(sheet.name)

        for row in range(ws_xls.nrows):
          for col in range(ws_xls.ncols):
            cell_value = ws_xls.cell_value(row, col)
            ws_new.cell(row=row+1, column=col+1, value=cell_value)
            
      wb.remove(wb.active)

      os.remove('temp.xls')
        
    
    elif filepath.endswith('.xls'):
      
      find_table_and_move_to_A5_xls_viernes_sabado(filepath, 'temp.xls')
      # Cargar el libro de Excel (xls)
      xls_workbook = xlrd.open_workbook('temp.xls')
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
      
      os.remove('temp.xls')

    # Seleccionar la hoja de trabajo
    button_viernes_sabado(wb)
  except Exception as e:
    messagebox.showerror("Error", f"Ha ocurrido un error al procesar el archivo:\n{str(e)}")


# Configurar la interfaz gráfica
root = tk.Tk()
root.wm_title("Informe Solicitudes y Expertas Disponibles")
root.geometry('420x160')
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
btn_procesar_archivo_excel = tk.Button(root, text="Procesar Archivo Excel (Lunes-Jueves)", command=procesar_archivo_excel).pack(pady=20)
btn_procesar_archivo_excel_viernes_sabado = tk.Button(root, text="Procesar Archivo Excel (Viernes-Sábado)", command=procesar_archivo_excel_viernes_sabado).pack(pady=20)
root.mainloop() 