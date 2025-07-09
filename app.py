import os
import uuid
import shutil
from flask import Flask, render_template, request, redirect, url_for, flash, session, Response
from werkzeug.utils import secure_filename
from buscador_core import generar_informe

app = Flask(__name__)
app.config['SECRET_KEY'] = 'una-clave-secreta-muy-dificil-de-adivinar'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 128 * 1024 * 1024  # Límite de 128 MB

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/buscar', methods=['POST'])
def buscar():
    if 'files' not in request.files:
        flash('No se seleccionó ningún archivo.', 'danger')
        return redirect(url_for('index'))

    files = request.files.getlist('files')
    search_strings_raw = request.form.get('search_strings')
    
    # === CAMBIO AQUÍ: La validación ahora permite el valor 0 ===
    try:
        context_chars = int(request.form.get('context_chars', 240))
        # Validar el rango de 0 a 1000
        if not (0 <= context_chars <= 1000):
            context_chars = 240
    except (ValueError, TypeError):
        context_chars = 240

    if not search_strings_raw or not search_strings_raw.strip():
        flash('Debes introducir al menos un texto para buscar.', 'danger')
        return redirect(url_for('index'))
    
    if not any(f.filename for f in files):
        flash('Debes seleccionar al menos un archivo.', 'danger')
        return redirect(url_for('index'))

    lista_strings = [s.strip() for s in search_strings_raw.split(';') if s.strip()]
    temp_dir = os.path.join(app.config['UPLOAD_FOLDER'], str(uuid.uuid4()))
    os.makedirs(temp_dir)

    try:
        for file in files:
            if file.filename:
                filename = secure_filename(file.filename)
                file.save(os.path.join(temp_dir, filename))

        reporte_str = generar_informe(temp_dir, lista_strings, context_chars)
        
        session['reporte'] = reporte_str
        
        return render_template('resultados.html', report_content=reporte_str)
    finally:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

@app.route('/descargar_reporte')
def descargar_reporte():
    reporte = session.get('reporte', 'No hay ningún informe para descargar.')
    return Response(
        reporte,
        mimetype="text/plain",
        headers={"Content-disposition": "attachment; filename=informe_busqueda_contexto.txt"}
    )

if __name__ == '__main__':
    app.run(debug=True)