from flask import Blueprint, request, send_file, jsonify
from .service import fill_excel_template

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