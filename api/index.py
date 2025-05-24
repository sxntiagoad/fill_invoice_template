from flask import Flask, request, send_file
import sys
import os

# Asegúrate de que el path de tu app esté en sys.path
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from app.service import fill_excel_template

app = Flask(__name__)

@app.route('/api/fill-invoice', methods=['POST'])
def fill_invoice():
    data = request.json
    excel_buffer = fill_excel_template(data)
    return send_file(
        excel_buffer,
        as_attachment=True,
        download_name="gastos_viaje_filled.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Vercel busca la variable "app"