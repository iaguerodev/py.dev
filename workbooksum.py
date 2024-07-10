import os
import pandas as pd
from openpyxl import Workbook, load_workbook
from copy import copy
import sys

# Establecer el directorio del ejecutable
if getattr(sys, 'frozen', False):
    script_directory = os.path.dirname(sys.executable)
else:
    script_directory = os.path.abspath(os.path.dirname(__file__))

print("Directorio del script:", script_directory)

# Nombre del archivo de salida
output_file_name = 'Suma_Total_Excel.xlsx'
output_file_path = os.path.join(script_directory, output_file_name)

# Lista de archivos Excel para procesar
excel_files = [f for f in os.listdir(script_directory) if f.endswith(('.xls', '.xlsx', '.xlsm')) and f != output_file_name]

# Crear un nuevo workbook para las sumas
sum_wb = Workbook()
sum_ws = sum_wb.active
sum_ws.title = "Suma Total"

# Variable para almacenar el workbook original si es necesario copiar formatos
original_wb = None

# Procesar cada archivo
for file in excel_files:
    file_path = os.path.join(script_directory, file)
    print(f"Procesando archivo: {file}")  # Imprimir el nombre del archivo actual
    df = pd.read_excel(file_path, sheet_name=None)  # Leer todas las hojas del archivo con Pandas

    # Cargar el workbook con openpyxl si es necesario para copiar formatos
    if file.endswith('.xlsx') or file.endswith('.xlsm'):
        original_wb = load_workbook(file_path, data_only=True)

    # Sumar los valores numéricos de cada hoja
    for sheet_name, data in df.items():
        for idx, row in data.iterrows():
            for col_idx, value in enumerate(row):
                if pd.notna(value) and isinstance(value, (int, float)):
                    cell_key = (idx + 1, col_idx + 1)
                    if cell_key in sum_ws._cells:
                        sum_ws.cell(row=idx + 1, column=col_idx + 1).value += value
                    else:
                        sum_ws.cell(row=idx + 1, column=col_idx + 1).value = value

                    # Copiar formatos del primer archivo si es necesario
                    if original_wb and idx == 0:  # Solo copiar formatos de la primera fila como ejemplo
                        original_cell = original_wb.active.cell(row=idx + 1, column=col_idx + 1)
                        target_cell = sum_ws.cell(row=idx + 1, column=col_idx + 1)
                        target_cell.font = copy(original_cell.font)
                        target_cell.border = copy(original_cell.border)
                        target_cell.fill = copy(original_cell.fill)
                        target_cell.number_format = original_cell.number_format
                        target_cell.alignment = copy(original_cell.alignment)

# Guardar el archivo de sumas
sum_wb.save(output_file_path)
print("El archivo se ha guardado en:", output_file_path)
