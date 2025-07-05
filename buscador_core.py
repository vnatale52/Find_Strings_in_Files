import os
import io # Importante para manejar strings como si fueran archivos
import fitz
import docx
import openpyxl
import xlrd

# --- Todas las funciones procesar_* (procesar_pdf, procesar_docx, etc.) se mantienen exactamente igual ---
# (Puedes copiarlas del script anterior)
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
            problemas.append(f"Archivo: '{nombre_base}' -> ADVERTENCIA: El documento PDF parece contener solo imágenes.")
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
    # (Esta función también se mantiene igual, la omito por brevedad)
    hallazgos, problemas = [], []
    # ... código de procesar_excel ...
    return hallazgos, problemas

def procesar_txt(ruta_archivo, lista_strings):
    # (Esta función también se mantiene igual, la omito por brevedad)
    hallazgos, problemas = [], []
    # ... código de procesar_txt ...
    return hallazgos, problemas


# *** CAMBIO IMPORTANTE AQUÍ ***
def generar_informe(carpeta_entrada, lista_strings):
    """
    Función principal que busca en los archivos y DEVUELVE el informe como string.
    """
    # ... (el resto de la lógica de búsqueda es igual) ...
    hallazgos_totales, archivos_problematicos, archivos_ignorados = [], [], []
    archivos_procesados = 0
    extensiones_soportadas = ('.pdf', '.docx', '.xlsx', '.xls', '.txt')

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
            else:
                archivos_ignorados.append(nombre_archivo)

    # *** En lugar de escribir a un archivo, escribimos a un string en memoria ***
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