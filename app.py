from flask import Flask, render_template, request, send_file, jsonify
from docx import Document
from docx.shared import Mm
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image
import os
import tempfile
import shutil
from datetime import datetime
import uuid

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500MB max total

# Extensiones válidas de imagen
VALID_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp'}

def allowed_file(filename):
    return os.path.splitext(filename.lower())[1] in VALID_EXTENSIONS

def images_to_word(image_paths, output_file, mode='standard'):
    """Convierte una lista de imágenes a un documento Word"""
    document = Document()
    
    # Configurar márgenes según el modo
    section = document.sections[0]
    
    if mode == 'standard':
        # Modo estándar con márgenes normales
        section.left_margin = Mm(10)
        section.right_margin = Mm(10)
        section.top_margin = Mm(10)
        section.bottom_margin = Mm(10)
    else:
        # Modo recibos: SIN MÁRGENES (100% de la hoja)
        section.left_margin = Mm(0)
        section.right_margin = Mm(0)
        section.top_margin = Mm(0)
        section.bottom_margin = Mm(0)

    # Dimensiones de página
    page_width = section.page_width
    page_height = section.page_height
    
    # Calcular ancho y alto disponibles
    available_width = page_width - section.left_margin - section.right_margin
    available_height = page_height - section.top_margin - section.bottom_margin

    processed = 0
    errors = []

    if mode == 'standard':
        # MODO ESTÁNDAR: Priorizar visibilidad completa (generalmente 1 por página si son grandes)
        for i, filepath in enumerate(image_paths):
            try:
                with Image.open(filepath) as img:
                    img_width, img_height = img.size
                    aspect_ratio = img_width / img_height
                    
                    target_width = available_width
                    target_height = int(target_width / aspect_ratio)
                    
                    if target_height > available_height:
                        target_height = available_height
                        target_width = int(target_height * aspect_ratio)
                    
                    document.add_picture(filepath, width=target_width, height=target_height)
                    last_paragraph = document.paragraphs[-1]
                    last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    if i > 0:
                        last_paragraph.paragraph_format.page_break_before = True
                    
                    processed += 1
            except Exception as e:
                errors.append(f"{os.path.basename(filepath)}: {str(e)}")

    else:
        # MODO RECIBOS: Grid 2 columnas x 2 filas (4 imágenes por hoja)
        # Crear tabla
        table = document.add_table(rows=0, cols=2)
        table.autofit = False 
        
        # Calcular dimensiones de celda para 2 columnas y 2 filas por página
        col_width = available_width / 2
        row_height = available_height / 2
        
        # Sin márgenes internos para ocupar toda la superficie
        cell_margin = Mm(0)
        
        # Dimensiones máximas de imagen dentro de la celda (toda la celda)
        max_img_width = col_width
        max_img_height = row_height

        current_row = None

        for i, filepath in enumerate(image_paths):
            try:
                # Determinar columna (0 o 1)
                col_idx = i % 2
                
                # Si es columna 0, crear nueva fila
                if col_idx == 0:
                    current_row = table.add_row()
                    current_row.height = int(row_height)
                    current_row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
                
                cell = current_row.cells[col_idx]
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                cell.width = int(col_width)
                
                # Procesar imagen
                with Image.open(filepath) as img:
                    img_width, img_height = img.size
                    aspect_ratio = img_width / img_height
                    
                    # Calcular tamaño para llenar completamente la celda
                    target_width = max_img_width
                    target_height = int(target_width / aspect_ratio)
                    
                    # Si es muy alta, ajustar por alto
                    if target_height > max_img_height:
                        target_height = max_img_height
                        target_width = int(target_height * aspect_ratio)
                    
                    # Limpiar párrafo existente en la celda
                    paragraph = cell.paragraphs[0]
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    run = paragraph.add_run()
                    run.add_picture(filepath, width=int(target_width), height=int(target_height))
                    
                processed += 1

            except Exception as e:
                errors.append(f"{os.path.basename(filepath)}: {str(e)}")

    document.save(output_file)
    return processed, errors

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'images' not in request.files:
        return jsonify({'error': 'No se encontraron imágenes'}), 400
    
    files = request.files.getlist('images')
    mode = request.form.get('mode', 'standard')
    
    if not files or all(f.filename == '' for f in files):
        return jsonify({'error': 'No se seleccionaron archivos'}), 400
    
    # Crear directorio temporal
    temp_dir = tempfile.mkdtemp()
    image_paths = []
    
    try:
        # Guardar archivos temporalmente y ordenarlos por nombre
        saved_files = []
        for file in files:
            if file and file.filename and allowed_file(file.filename):
                # Mantener nombre original para ordenar
                filename = file.filename
                filepath = os.path.join(temp_dir, filename)
                file.save(filepath)
                saved_files.append((filename, filepath))
        
        if not saved_files:
            return jsonify({'error': 'No se encontraron imágenes válidas'}), 400
        
        # Ordenar por nombre de archivo
        saved_files.sort(key=lambda x: x[0])
        image_paths = [f[1] for f in saved_files]
        
        # Generar documento Word
        output_filename = f"documento_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        output_path = os.path.join(temp_dir, output_filename)
        
        processed, errors = images_to_word(image_paths, output_path, mode)
        
        if processed == 0:
            return jsonify({'error': 'No se pudo procesar ninguna imagen'}), 400
        
        # Enviar archivo
        response = send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        
        # Limpiar después de enviar (usando callback)
        @response.call_on_close
        def cleanup():
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
            except:
                pass
        
        return response
        
    except Exception as e:
        # Limpiar en caso de error
        shutil.rmtree(temp_dir, ignore_errors=True)
        return jsonify({'error': f'Error al procesar: {str(e)}'}), 500

if __name__ == '__main__':
    # Crear carpeta templates si no existe
    os.makedirs('templates', exist_ok=True)
    print("🚀 Servidor iniciado en http://localhost:5001")
    app.run(debug=True, host='0.0.0.0', port=5001)
