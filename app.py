import os
import uuid
import shutil
from flask import Flask, render_template, request, redirect, url_for, flash, session, Response
from werkzeug.utils import secure_filename

# Importar nuestra lógica de búsqueda refactorizada
from buscador_core import generar_informe

app = Flask(__name__)
app.config['SECRET_KEY'] = 'una-clave-secreta-muy-dificil-de-adivinar'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # Límite de 32 MB para el total de archivos

# Asegurarse de que la carpeta a subir existe
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

@app.route('/')
def index():
    """Muestra el formulario principal."""
    return render_template('index.html')

@app.route('/buscar', methods=['POST'])
def buscar():
    """
    Recibe los archivos y los textos, procesa la búsqueda y muestra los resultados.
    """
    if 'files' not in request.files:
        flash('No se seleccionó ningún archivo.', 'danger')
        return redirect(url_for('index'))

    files = request.files.getlist('files')
    search_strings_raw = request.form.get('search_strings')

    if not search_strings_raw or not search_strings_raw.strip():
        flash('Debes introducir al menos un texto para buscar.', 'danger')
        return redirect(url_for('index'))
    
    if not any(f.filename for f in files):
        flash('Debes seleccionar al menos un archivo.', 'danger')
        return redirect(url_for('index'))

    lista_strings = [s.strip() for s in search_strings_raw.split(';') if s.strip()]

    # Crear una carpeta temporal única para esta búsqueda
    temp_dir = os.path.join(app.config['UPLOAD_FOLDER'], str(uuid.uuid4()))
    os.makedirs(temp_dir)

    try:
        # Guardar los archivos en la carpeta temporal
        for file in files:
            if file.filename:
                filename = secure_filename(file.filename)
                file.save(os.path.join(temp_dir, filename))

        # Llamar a nuestra función de búsqueda
        reporte_str = generar_informe(temp_dir, lista_strings)
        
        # Guardar el reporte en la sesión del usuario para poder descargarlo después
        session['reporte'] = reporte_str
        
        return render_template('resultados.html', report_content=reporte_str)

    finally:
        # Limpiar: borrar la carpeta temporal y su contenido
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

@app.route('/descargar_reporte')
def descargar_reporte():
    """Permite al usuario descargar el último informe generado."""
    reporte = session.get('reporte', 'No hay ningún informe para descargar.')
    
    return Response(
        reporte,
        mimetype="text/plain",
        headers={"Content-disposition": "attachment; filename=informe_busqueda.txt"}
    )

if __name__ == '__main__':
    # Esto es para ejecutar en modo de desarrollo local
    app.run(debug=True)