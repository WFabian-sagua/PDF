# PDF to Excel Conversion Script

Este script convierte un archivo PDF en una hoja de cálculo de Excel, aplicando formato y estilos a los datos extraídos.

## Autor
- Nombre: Fabian Sagua
- Correo electrónico: datacountai@proton.me

## Librerías
Para ejecutar el script es necesario contar con las siguientes librerías en python:

- pdfplumber  -> importa con el comando: pip install pdfplumber
- pandas      -> importa con el comando: pip install pandas
- openpyxl    -> importa con el comando: pip install openpyxl


## Código

```python
import pdfplumber  # Importa la biblioteca pdfplumber para trabajar con archivos PDF
import pandas as pd  # Importa la biblioteca pandas para trabajar con DataFrames
from openpyxl import Workbook  # Importa la clase Workbook de la biblioteca openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side  # Importa clases para formato en Excel
import os  # Importa el módulo os para trabajar con rutas de archivos

# Ubicación y nombre del PDF
pdf_filename = 'report.pdf'  # Nombre de tu archivo PDF
pdf_path = os.path.join(os.path.dirname(__file__), pdf_filename)

# Abrir PDF con pdfplumber
with pdfplumber.open(pdf_path) as pdf:
    primera_pagina = pdf.pages[0]  # Obtén la primera página
    tablas = primera_pagina.extract_tables()  # Extrae tablas de la página

# Extraer los nombres de las columnas desde la primera fila de la tabla
nombres_columnas = tablas[0][0]

# Convertir la tabla en un DataFrame en pandas
data_frame = pd.DataFrame(tablas[0][1:], columns=nombres_columnas)

# Crear un libro de trabajo en Excel
libro_excel = Workbook()
hoja_excel = libro_excel.active

# Dar formato en negrita y bordes a los encabezados
encabezado_fill = PatternFill(start_color='000000', end_color='000000', fill_type='solid')
encabezado_font = Font(color='FFFFFF', bold=True)
encabezado_border = Border(bottom=Side(border_style='thin'))

for col, encabezado in enumerate(nombres_columnas, start=1):
    celda = hoja_excel.cell(row=1, column=col, value=encabezado)
    celda.font = encabezado_font
    celda.fill = encabezado_fill
    celda.border = encabezado_border

# Agregar los datos del DataFrame a la hoja de Excel
for fila_datos in data_frame.itertuples(index=False):
    hoja_excel.append(fila_datos)

# Dar formato a todas las celdas con bordes
bordes = Border(left=Side(border_style='thin'), right=Side(border_style='thin'),
                top=Side(border_style='thin'), bottom=Side(border_style='thin'))
for fila in hoja_excel.iter_rows(min_row=1, max_row=hoja_excel.max_row,
                                 min_col=1, max_col=hoja_excel.max_column):
    for celda in fila:
        celda.border = bordes

# Guardar el archivo Excel
archivo_excel = 'datos_extraidos.xlsx'
libro_excel.save(archivo_excel)

print(f'Datos exportados con éxito a {archivo_excel}')
