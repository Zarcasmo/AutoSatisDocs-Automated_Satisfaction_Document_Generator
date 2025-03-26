# -*- coding: utf-8 -*-
"""
Created on Tue Feb 25 10:44:13 2025

@author: alopagui

Script para la conversión de archivos Word a PDF de manera automatizada.

Este script toma un conjunto de archivos .docx y los convierte en archivos .pdf utilizando la automatización de Microsoft Word a través de win32com.client. 
Incluye una barra de progreso para monitorear el avance del proceso y muestra el tiempo transcurrido.

Requisitos:
- Python 3
- Bibliotecas: os, time, colorama, win32com.client

Uso:
1. Asegurar que Microsoft Word está instalado en el sistema.
2. Ubicar los archivos .docx en la carpeta de origen.
3. Ejecutar el script para convertir los archivos en formato PDF.
"""

import pandas as pd
from docx import Document
from docx.shared import Inches
import os
import comtypes.client  
import time
import sys
from colorama import Fore, Style, init

# Inicializar colorama para colores en la consola
init(autoreset=True)

#Tamaño que tendran las firmas
tam_imagen=1

# Cargar el archivo Excel con los datos de entrada
# Este archivo debe contener las columnas necesarias como "Nombre", "Cedula", "Calidad", "Municipio", etc.
df = pd.read_excel("datos.xlsx")

# Ruta de la plantilla Word que se usará como base para generar los documentos
template_path = "formato_socializa.docx"
output_folder = "output_pdfs" # Carpeta donde se guardarán los archivos generados
firmas_folder = "Firmas"  # Carpeta donde están las imágenes de firmas
os.makedirs(output_folder, exist_ok=True) # Crear la carpeta de salida si no existe

# Archivo Excel de salida con la consolidación de resultados
output_excel_path = os.path.join(output_folder, "Resumen_Resultados.xlsx")

# Inicializar lista para almacenar los resultados
resultados = []

# Función para imprimir una barra de progreso durante la conversión a PDF
def print_progress_bar(iteration, total, length=50, start_time=None):
    """
    Imprime una barra de progreso en la consola y muestra el tiempo transcurrido al completar.

    :param iteration: El número actual de iteraciones.
    :param total: El número total de iteraciones.
    :param length: La longitud de la barra de progreso.
    :param start_time: El tiempo de inicio de la ejecución.
    """
    progress = (iteration / total)
    arrow = '█' * int(round(progress * length))
    spaces = '░' * (length - len(arrow))
    percent = int(round(progress * 100))
    
    # Construcción de la barra con colores
    bar = f'[{Fore.GREEN}{arrow}{spaces}{Style.RESET_ALL}]'
    percent_text = f'{Fore.CYAN}{percent}%{Style.RESET_ALL}'
    
    if start_time:
        elapsed_time = time.time() - start_time
        elapsed_time_formatted = time.strftime("%H:%M:%S", time.gmtime(elapsed_time))
        sys.stdout.write(f'\r{bar} {percent_text} Complete - Time elapsed: {elapsed_time_formatted}')
    else:
        sys.stdout.write(f'\r{bar} {percent_text} Complete')
    
    # Si ha llegado al 100%, imprime una nueva línea
    if iteration == total:
        sys.stdout.write('\n')
    
    sys.stdout.flush()

# Función para reemplazar texto manteniendo formato
def replace_text_keep_format(doc, old_text, new_text):
    # Buscar en párrafos normales
    for para in doc.paragraphs:
        if old_text in para.text:
            for run in para.runs:
                run.text = run.text.replace(old_text, new_text)
    
    # Buscar en tablas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if old_text in cell.text:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.text = run.text.replace(old_text, new_text)

# Función para insertar imágenes en el documento Word reemplazando texto específico
def replace_text_with_image(doc, placeholder, image_path, width=tam_imagen):
    for para in doc.paragraphs:
        if placeholder in para.text:
            para.clear()
            run = para.add_run()
            run.add_picture(image_path, width=Inches(width))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if placeholder in cell.text:
                    cell.text = ""  
                    run = cell.paragraphs[0].add_run()
                    run.add_picture(image_path, width=Inches(width))

# Listas para almacenar rutas de archivos generados
word_files = []
pdf_files = []

# Procesar cada fila del Excel
for index, row in df.iterrows():
    word_output_path = os.path.join(output_folder, f"documento_{index}.docx")
    pdf_output_path = os.path.join(output_folder, f"documento_{index}.pdf")
    status = "Éxito"
    error_msg = ""

    try:
        # Crear un nuevo documento basado en la plantilla
        doc = Document(template_path)

        #Datos usuarios
        replace_text_keep_format(doc, "@NOMBRE@", row["NOMBRE"])
        replace_text_keep_format(doc, "@CEDULA@", str(row["CEDULA"]))
        replace_text_keep_format(doc, "@CALIDAD@", row["CALIDAD"])
        #Ubicacion
        replace_text_keep_format(doc, "@MUNICIPIO@", row["MUNICIPIO"])
        replace_text_keep_format(doc, "@VEREDA@", row["VEREDA"])
        replace_text_keep_format(doc, "@DIRECCION@", row["DIRECCION"])
        replace_text_keep_format(doc, "@LATITUD@", str(row["LATITUD"]))
        replace_text_keep_format(doc, "@LONGITUD@", str(row["LONGITUD"]))
        #Comentarios
        replace_text_keep_format(doc, "@COMENTARIO_EDEQ@", row["COMENTARIO_EDEQ"])
        replace_text_keep_format(doc, "@COMENTARIO_USUARIO@", row["COMENTARIO_USUARIO"])
        #Fecha y OT
        replace_text_keep_format(doc, "@OT@", row["OT"])
            

        # Cargar e insertar imágenes de firmas si existen
        autorizacion_path = os.path.join(firmas_folder, row["Autorizacion"])
        satisfaccion_path = os.path.join(firmas_folder, row["Satisfaccion"])

        if os.path.exists(autorizacion_path):
            replace_text_with_image(doc, "@FIRMA_AUTORIZA@", autorizacion_path)
        else:
            status = "Fallo"
            error_msg += f"Falta imagen de autorización: {autorizacion_path}. "

        if os.path.exists(satisfaccion_path):
            replace_text_with_image(doc, "@FIRMA_SATISFACCION@", satisfaccion_path)
        else:
            status = "Fallo"
            error_msg += f"Falta imagen de satisfacción: {satisfaccion_path}. "

        # Guardar documento Word
        doc.save(word_output_path)
        word_files.append(word_output_path)
        pdf_files.append(pdf_output_path)

    except Exception as e:
        status = "Fallo"
        error_msg += f"Error en generación Word: {str(e)}. "

    # Guardar información en la lista de resultados
    row_data = row.to_dict()
    row_data["Estado"] = status
    row_data["Errores"] = error_msg
    resultados.append(row_data)

# Crear DataFrame de resultados y guardarlo en Excel
df_resultados = pd.DataFrame(resultados)
df_resultados.to_excel(output_excel_path, index=False)

# Conversión a PDF con barra de progreso
total = len(word_files)-1
start_time = time.time()  # Captura el tiempo de inicio
cuenta_pdf = 0

try:
    word_app = comtypes.client.CreateObject('Word.Application')
    word_app.Visible = False  

    for word_path, pdf_path in zip(word_files, pdf_files):
        try:
            doc = word_app.Documents.Open(os.path.abspath(word_path))
            doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)  # Guardar como PDF
            doc.Close()
        except Exception as e:
            resultados[cuenta_pdf]["Estado"] = "Fallo"
            resultados[cuenta_pdf]["Errores"] += f"Error en conversión a PDF: {str(e)}. "
        
        os.remove(word_path)
        print_progress_bar(cuenta_pdf, total, start_time=start_time)
        cuenta_pdf += 1

    word_app.Quit()
except Exception as e:
    print(f"\n❌ Error en la inicialización de Word: {str(e)}")

# Guardar nuevamente el Excel actualizado con errores de PDF
df_resultados = pd.DataFrame(resultados)
df_resultados.to_excel(output_excel_path, index=False)

print(f"\n✅ Proceso completado. Se generaron los documentos y el reporte en Excel de {cuenta_pdf} actas.")

