import os
import io
import fitz  # PyMuPDF
import docx  # python-docx
import openpyxl
from openpyxl.utils import get_column_letter
import xlrd

# --- Las funciones procesar_* (pdf, docx, excel, txt) no cambian ---
# (Se incluyen completas para que el archivo sea autocontenido)

def procesar_pdf(ruta_archivo, lista_strings):
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
                if ocurrencias := pagina.search_for(string_a_buscar):
                    hallazgos.append(f"Archivo: '{nombre_base}', Página: {num_pagina} -> Encontrado: '{string_a_buscar}' ({len(ocurrencias)} ocurrencia(s)).")
        documento.close()
    except Exception as e:
        problemas.append(f"Archivo: '{nombre_base}' -> ERROR: No se pudo procesar. Razón: {e}")
    return hallazgos, problemas

def procesar_docx(ruta_archivo, lista_strings):
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
    hallazgos, problemas = [], []
    nombre_base = os.path.basename(ruta_archivo)
    try:
        if ruta_archivo.lower().endswith('.xlsx'):
            wb = openpyxl.load_workbook(ruta_archivo, data_only=True)
            for nombre_hoja in wb.sheetnames:
                hoja = wb[nombre_hoja]
                for fila_idx, fila in enumerate(hoja.iter_rows(), start=1):
                    for col_idx, celda in enumerate(fila, start=1):
                        if celda.value and isinstance(celda.value, str):
                            for string_a_buscar in lista_strings:
                                if string_a_buscar.lower() in celda.value.lower():
                                    conteo = celda.value.lower().count(string_a_buscar.lower())
                                    celda_ref = f"{get_column_letter(col_idx)}{fila_idx}"
                                    mensaje = f"Archivo: '{nombre_base}', Hoja: '{nombre_hoja}', Celda: {celda_ref} -> Encontrado: '{string_a_buscar}' ({conteo} ocurrencia(s))."
                                    hallazgos.append(mensaje)
        elif ruta_archivo.lower().endswith('.xls'):
            wb = xlrd.open_workbook(ruta_archivo)
            for nombre_hoja in wb.sheet_names():
                hoja = wb.sheet_by_name(nombre_hoja)
                for fila_idx in range(hoja.nrows):
                    for col_idx in range(hoja.ncols):
                        valor_celda = str(hoja.cell_value(fila_idx, col_idx))
                        for string_a_buscar in lista_strings:
                            if string_a_buscar.lower() in valor_celda.lower():
                                conteo = valor_celda.lower().count(string_a_buscar.lower())
                                celda_ref = f"{chr(65+col_idx)}{fila_idx+1}"
                                mensaje = f"Archivo: '{nombre_base}', Hoja: '{nombre_hoja}', Celda: ~{celda_ref} -> Encontrado: '{string_a_buscar}' ({conteo} ocurrencia(s))."
                                hallazgos.append(mensaje)
    except Exception as e:
        problemas.append(f"Archivo: '{nombre_base}' -> ERROR: No se pudo procesar. Razón: {e}")
    return hallazgos, problemas

def procesar_txt(ruta_archivo, lista_strings):
    hallazgos, problemas = [], []
    nombre_base = os.path.basename(ruta_archivo)
    try:
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
    Función principal que busca en los archivos y DEVUELVE el informe como string,
    incluyendo una sección final de resumen con totales.
    """
    hallazgos_totales, archivos_problematicos, archivos_ignorados = [], [], []
    
    # --- NUEVO: Contadores para el resumen ---
    archivos_procesados = 0
    # Usamos un 'set' para contar archivos únicos con problemas,
    # incluso si un archivo genera múltiples errores.
    set_archivos_con_problemas = set()
    
    extensiones_soportadas = ('.pdf', '.docx', '.xlsx', '.xls', '.txt')

    # --- Lógica de procesamiento (sin cambios) ---
    for nombre_archivo in os.listdir(carpeta_entrada):
        ruta_completa = os.path.join(carpeta_entrada, nombre_archivo)
        if os.path.isfile(ruta_completa):
            extension = os.path.splitext(nombre_archivo)[1].lower()
            if extension in extensiones_soportadas:
                archivos_procesados += 1
                if extension == '.pdf': hallazgos, problemas = procesar_pdf(ruta_completa, lista_strings)
                elif extension == '.docx': hallazgos, problemas = procesar_docx(ruta_completa, lista_strings)
                elif extension in ['.xlsx', '.xls']: hallazgos, problemas = procesar_excel(ruta_completa, lista_strings)
                elif extension == '.txt': hallazgos, problemas = procesar_txt(ruta_completa, lista_strings)
                
                hallazgos_totales.extend(hallazgos)
                archivos_problematicos.extend(problemas)
                
                # --- NUEVO: Registrar si este archivo tuvo problemas ---
                if problemas:
                    set_archivos_con_problemas.add(nombre_archivo)
            else:
                archivos_ignorados.append(nombre_archivo)

    # --- Cálculo de totales para el resumen ---
    total_con_problemas = len(set_archivos_con_problemas)
    total_ignorados = len(archivos_ignorados)
    total_sin_problemas = archivos_procesados - total_con_problemas
    total_seleccionados = archivos_procesados + total_ignorados

    # --- Generación del informe ---
    output = io.StringIO()
    output.write("="*30 + " INFORME DE BÚSQUEDA " + "="*30 + "\n")
    output.write(f"Textos Buscados: {lista_strings}\n")
    output.write(f"Extensiones Soportadas: {', '.join(extensiones_soportadas)}\n")
    output.write("="*79 + "\n\n")
    
    # Secciones de detalles (sin cambios)
    output.write("--- OCURRENCIAS HALLADAS ---\n\n")
    if hallazgos_totales:
        for hallazgo in hallazgos_totales: output.write(hallazgo + "\n")
    else:
        output.write("No se encontraron ocurrencias de los textos buscados.\n")

    output.write("\n\n--- ARCHIVOS PROCESADOS CON PROBLEMAS O ADVERTENCIAS ---\n\n")
    if archivos_problematicos:
        for problema in archivos_problematicos: output.write(problema + "\n")
    else:
        output.write("Todos los archivos soportados fueron analizados sin errores.\n")

    output.write("\n\n--- ARCHIVOS NO SOPORTADOS E IGNORADOS ---\n\n")
    output.write(f"Total: {total_ignorados}\n\n")
    if archivos_ignorados:
        for archivo in sorted(archivos_ignorados): output.write(f"- {archivo}\n")
    else:
        output.write("No se encontraron archivos con formatos no soportados.\n")
    
    # --- NUEVA SECCIÓN DE RESUMEN FINAL ---
    output.write("\n\n" + "="*33 + " RESUMEN FINAL " + "="*33 + "\n")
    output.write(f"TOTAL DE ARCHIVOS SELECCIONADOS: {total_seleccionados}\n")
    output.write(f"  - TOTAL DE ARCHIVOS PROCESADOS SIN PROBLEMAS: {total_sin_problemas}\n")
    output.write(f"  - TOTAL DE ARCHIVOS PROCESADOS CON PROBLEMAS O ADVERTENCIAS: {total_con_problemas}\n")
    output.write(f"  - TOTAL DE ARCHIVOS NO SOPORTADOS E IGNORADOS: {total_ignorados}\n")
    
    return output.getvalue()