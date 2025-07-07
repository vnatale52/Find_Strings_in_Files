import os
import io
import fitz  # PyMuPDF
import docx  # python-docx
import openpyxl
from openpyxl.utils import get_column_letter
import xlrd

def procesar_pdf(ruta_archivo, lista_strings):
    """
    Procesa un único archivo PDF, devolviendo hallazgos y problemas.
    """
    hallazgos, problemas = [], []
    nombre_base = os.path.basename(ruta_archivo)
    try:
        documento = fitz.open(ruta_archivo)
        if not documento.page_count:
            problemas.append(f"Archivo: '{nombre_base}' -> ERROR: El PDF está vacío o corrupto (0 páginas).")
            return hallazgos, problemas

        documento_sin_texto = all(not pagina.get_text("text").strip() for pagina in documento)
        if documento_sin_texto:
            problemas.append(f"Archivo: '{nombre_base}' -> ADVERTENCIA: El documento PDF parece contener solo imágenes y no tiene texto extraíble.")
            documento.close()
            return hallazgos, problemas

        for num_pagina, pagina in enumerate(documento, start=1):
            for string_a_buscar in lista_strings:
                # La sintaxis ':=' (operador morsa) asigna y comprueba en un solo paso
                if ocurrencias := pagina.search_for(string_a_buscar):
                    hallazgos.append(f"Archivo: '{nombre_base}', Página: {num_pagina} -> Encontrado: '{string_a_buscar}' ({len(ocurrencias)} ocurrencia(s)).")
        documento.close()
    except Exception as e:
        problemas.append(f"Archivo: '{nombre_base}' -> ERROR: No se pudo procesar. Razón: {e}")
    return hallazgos, problemas

def procesar_docx(ruta_archivo, lista_strings):
    """
    Procesa un archivo .docx. NOTA: No es posible obtener el nº de página de forma fiable,
    por lo que se informa el nº de párrafo, que es la mejor localización posible.
    """
    hallazgos, problemas = [], []
    nombre_base = os.path.basename(ruta_archivo)
    try:
        documento = docx.Document(ruta_archivo)
        for num_parrafo, parrafo in enumerate(documento.paragraphs, start=1):
            for string_a_buscar in lista_strings:
                if string_a_buscar.lower() in parrafo.text.lower():
                    conteo = parrafo.text.lower().count(string_a_buscar.lower())
                    hallazgos.append(f"Archivo: '{nombre_base}', Párrafo: {num_parrafo} -> Encontrado: '{string_a_buscar}' ({conteo} ocurrencia(s)).")
    except Exception as e:
        problemas.append(f"Archivo: '{nombre_base}' -> ERROR: No se pudo procesar. Razón: {e}")
    return hallazgos, problemas

def procesar_excel(ruta_archivo, lista_strings):
    """
    Procesa un archivo Excel (.xlsx o .xls), buscando en todas las celdas de texto.
    """
    hallazgos, problemas = [], []
    nombre_base = os.path.basename(ruta_archivo)
    try:
        # Procesar archivos .xlsx (formato moderno)
        if ruta_archivo.lower().endswith('.xlsx'):
            wb = openpyxl.load_workbook(ruta_archivo, data_only=True) # data_only=True para obtener valores de fórmulas
            for nombre_hoja in wb.sheetnames:
                hoja = wb[nombre_hoja]
                for fila_idx, fila in enumerate(hoja.iter_rows(), start=1):
                    for col_idx, celda in enumerate(fila, start=1):
                        # Solo buscar en celdas con valor de tipo string
                        if celda.value and isinstance(celda.value, str):
                            for string_a_buscar in lista_strings:
                                if string_a_buscar.lower() in celda.value.lower():
                                    conteo = celda.value.lower().count(string_a_buscar.lower())
                                    celda_ref = f"{get_column_letter(col_idx)}{fila_idx}"
                                    mensaje = f"Archivo: '{nombre_base}', Hoja: '{nombre_hoja}', Celda: {celda_ref} -> Encontrado: '{string_a_buscar}' ({conteo} ocurrencia(s))."
                                    hallazgos.append(mensaje)
        # Procesar archivos .xls (formato antiguo)
        elif ruta_archivo.lower().endswith('.xls'):
            wb = xlrd.open_workbook(ruta_archivo)
            for nombre_hoja in wb.sheet_names():
                hoja = wb.sheet_by_name(nombre_hoja)
                for fila_idx in range(hoja.nrows):
                    for col_idx in range(hoja.ncols):
                        # Convertir siempre a string para una búsqueda segura
                        valor_celda = str(hoja.cell_value(fila_idx, col_idx))
                        for string_a_buscar in lista_strings:
                            if string_a_buscar.lower() in valor_celda.lower():
                                conteo = valor_celda.lower().count(string_a_buscar.lower())
                                # Aproximación de la referencia de celda
                                celda_ref = f"{chr(65+col_idx)}{fila_idx+1}"
                                mensaje = f"Archivo: '{nombre_base}', Hoja: '{nombre_hoja}', Celda: ~{celda_ref} -> Encontrado: '{string_a_buscar}' ({conteo} ocurrencia(s))."
                                hallazgos.append(mensaje)
    except Exception as e:
        problemas.append(f"Archivo: '{nombre_base}' -> ERROR: No se pudo procesar. Razón: {e}")
    return hallazgos, problemas

def procesar_txt(ruta_archivo, lista_strings):
    """
    Procesa un archivo de texto plano (.txt).
    """
    hallazgos, problemas = [], []
    nombre_base = os.path.basename(ruta_archivo)
    try:
        # Abrir con encoding utf-8 y manejar errores para máxima compatibilidad
        with open(ruta_archivo, 'r', encoding='utf-8', errors='ignore') as f:
            for num_linea, linea in enumerate(f, start=1):
                for string_a_buscar in lista_strings:
                    if string_a_buscar.lower() in linea.lower():
                        conteo = linea.lower().count(string_a_buscar.lower())
                        mensaje = f"Archivo: '{nombre_base}', Línea: {num_linea} -> Encontrado: '{string_a_buscar}' ({conteo} ocurrencia(s))."
                        hallazgos.append(mensaje)
    except Exception as e:
        problemas.append(f"Archivo: '{nombre_base}' -> ERROR: No se pudo procesar. Razón: {e}")
    return hallazgos, problemas


def generar_informe(carpeta_entrada, lista_strings):
    """
    Función principal que busca en los archivos y DEVUELVE el informe como string.
    """
    hallazgos_totales, archivos_problematicos, archivos_ignorados = [], [], []
    archivos_procesados = 0
    extensiones_soportadas = ('.pdf', '.docx', '.xlsx', '.xls', '.txt')

    for nombre_archivo in os.listdir(carpeta_entrada):
        ruta_completa = os.path.join(carpeta_entrada, nombre_archivo)
        if os.path.isfile(ruta_completa):
            extension = os.path.splitext(nombre_archivo)[1].lower()
            if extension in extensiones_soportadas:
                archivos_procesados += 1
                if extension == '.pdf':
                    hallazgos, problemas = procesar_pdf(ruta_completa, lista_strings)
                elif extension == '.docx':
                    hallazgos, problemas = procesar_docx(ruta_completa, lista_strings)
                elif extension in ['.xlsx', '.xls']:
                    hallazgos, problemas = procesar_excel(ruta_completa, lista_strings)
                elif extension == '.txt':
                    hallazgos, problemas = procesar_txt(ruta_completa, lista_strings)
                
                hallazgos_totales.extend(hallazgos)
                archivos_problematicos.extend(problemas)
            else:
                archivos_ignorados.append(nombre_archivo)

    # Escribimos a un objeto de string en memoria para devolverlo al final
    output = io.StringIO()
    output.write("="*30 + " INFORME DE BÚSQUEDA " + "="*30 + "\n")
    output.write(f"Textos Buscados: {lista_strings}\n")
    output.write(f"Extensiones Soportadas: {', '.join(extensiones_soportadas)}\n")
    output.write("="*79 + "\n\n")
    
    output.write("--- OCURRENCIAS HALLADAS ---\n\n")
    if hallazgos_totales:
        for hallazgo in hallazgos_totales:
            output.write(hallazgo + "\n")
    else:
        output.write("No se encontraron ocurrencias de los textos buscados.\n")

    output.write("\n\n--- ARCHIVOS PROCESADOS CON PROBLEMAS O ADVERTENCIAS ---\n\n")
    if archivos_problematicos:
        for problema in archivos_problematicos:
            output.write(problema + "\n")
    else:
        output.write("Todos los archivos soportados fueron analizados sin errores.\n")

    output.write("\n\n--- ARCHIVOS NO SOPORTADOS E IGNORADOS ---\n\n")
    total_ignorados = len(archivos_ignorados)
    output.write(f"Total de archivos ignorados: {total_ignorados}\n\n")
    if archivos_ignorados:
        for archivo in sorted(archivos_ignorados):
            output.write(f"- {archivo}\n")
    else:
        output.write("No se encontraron archivos con formatos no soportados.\n")
    
    # Devolvemos todo el contenido del string
    return output.getvalue()