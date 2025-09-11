from flask import Blueprint, request, send_file, jsonify
from .service import fill_excel_template, generate_invoice_pdf, generate_exportable_excel
import requests
from io import BytesIO

main = Blueprint('main', __name__)

@main.route('/fill-invoice', methods=['POST'])
def fill_invoice():
    data = request.json
    try:
        excel_buffer = fill_excel_template(data)
        return send_file(
            excel_buffer,
            as_attachment=True,
            download_name="gastos_viaje_filled.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 400

@main.route('/generate-invoice-pdf', methods=['POST'])
def generate_pdf():
    try:
        data = request.json
        if not data or 'image_urls' not in data:
            return jsonify({"error": "No se encontraron URLs de imágenes"}), 400
        
        image_urls = data['image_urls']
        if not image_urls:
            return jsonify({"error": "La lista de URLs está vacía"}), 400

        # Descargar las imágenes desde las URLs
        images_data = []
        for url in image_urls:
            try:
                response = requests.get(url)
                response.raise_for_status()  # Verifica si hay errores en la respuesta
                images_data.append(response.content)
            except requests.RequestException as e:
                return jsonify({"error": f"Error al descargar imagen desde {url}: {str(e)}"}), 400
        
        # Generar el PDF
        pdf_buffer = generate_invoice_pdf(images_data)
        
        return send_file(
            pdf_buffer,
            as_attachment=True,
            download_name="facturas.pdf",
            mimetype="application/pdf"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 400

@main.route('/generate-exportable-excel', methods=['POST'])
def generate_exportable_excel_route():
    """
    Endpoint para generar Excel contable desde datos del modelo Exportable
    """
    try:
        data = request.json
        if not data:
            return jsonify({"error": "No se encontraron datos"}), 400
        
        # Validar que tenga la estructura esperada
        required_fields = ['data']
        if not all(field in data for field in required_fields):
            return jsonify({"error": "Estructura de datos inválida. Se requiere: data"}), 400
        
        # Validar que data sea una lista
        if not isinstance(data.get('data'), list):
            return jsonify({"error": "El campo 'data' debe ser una lista"}), 400
        
        excel_buffer = generate_exportable_excel(data)
        
        filename = data.get('filename', 'exportable_contable')
        if not filename.endswith('.xlsx'):
            filename += '.xlsx'
        
        return send_file(
            excel_buffer,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500