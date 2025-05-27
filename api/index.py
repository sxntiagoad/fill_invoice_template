from flask import Flask, request, send_file, Response
from flask_cors import CORS
import sys
import os

# Asegúrate de que el path de tu app esté en sys.path
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from app.service import fill_excel_template

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