from flask import Flask, request, send_file, Response
from flask_cors import CORS
import sys
import os
import requests

# Asegúrate de que el path de tu app esté en sys.path
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from app.service import fill_excel_template, generate_invoice_pdf

app = Flask(__name__)
# Habilitar CORS para todas las rutas
CORS(app)

@app.route('/api/fill-invoice', methods=['POST'])
def fill_invoice():
    data = request.json
    excel_buffer = fill_excel_template(data)
    
    # Enviamos el archivo como un stream de bytes
    return Response(
        excel_buffer.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": "inline",
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }
    )

@app.route('/api/generate-invoice-pdf', methods=['POST'])
def generate_pdf():
    try:
        data = request.json
        if not data or 'image_urls' not in data:
            return Response(
                '{"error": "No se encontraron URLs de imágenes"}',
                status=400,
                mimetype='application/json'
            )
        
        image_urls = data['image_urls']
        if not image_urls:
            return Response(
                '{"error": "La lista de URLs está vacía"}',
                status=400,
                mimetype='application/json'
            )

        # Descargar las imágenes desde las URLs
        images_data = []
        for url in image_urls:
            try:
                response = requests.get(url)
                response.raise_for_status()
                images_data.append(response.content)
            except requests.RequestException as e:
                return Response(
                    f'{{"error": "Error al descargar imagen desde {url}: {str(e)}"}}',
                    status=400,
                    mimetype='application/json'
                )
        
        # Generar el PDF
        pdf_buffer = generate_invoice_pdf(images_data)
        
        return Response(
            pdf_buffer.getvalue(),
            mimetype="application/pdf",
            headers={
                "Content-Disposition": "inline",
                "Content-Type": "application/pdf"
            }
        )
    except Exception as e:
        return Response(
            f'{{"error": "{str(e)}"}}',
            status=400,
            mimetype='application/json'
        )