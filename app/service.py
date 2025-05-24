import io
import os
from openpyxl import load_workbook

def get_template_path():
    base_dir = os.path.dirname(os.path.dirname(__file__))
    return os.path.join(base_dir, 'template', 'GASTOS_VIAJE.xlsx')

def obtener_celda_principal(hoja, celda):
    """Obtiene la celda principal si está en un rango fusionado"""
    for merged_range in hoja.merged_cells.ranges:
        if celda.coordinate in merged_range:
            return hoja.cell(merged_range.min_row, merged_range.min_col)
    return celda

def fill_excel_template(data):
    wb = load_workbook(get_template_path())
    ws = wb.active

    def obtener_celda_principal(hoja, celda):
        for merged_range in hoja.merged_cells.ranges:
            if celda.coordinate in merged_range:
                return hoja.cell(merged_range.min_row, merged_range.min_col)
        return celda

    # Llenar campos básicos
    campos = {
        'E4': data.get('empresa', ''),
        'E6': data.get('nit', ''),
        'B7': data.get('placa', ''),
        'H7': data.get('conductor', ''),
        'E10': data.get('desde', ''),
        'I10': data.get('hasta', ''),
        'B13': data.get('fecha', ''),
        'G14': data.get('anticipo', ''),
        'J14': data.get('flete', '')
    }
    for celda, valor in campos.items():
        cell = ws[celda]
        main_cell = obtener_celda_principal(ws, cell)
        main_cell.value = valor

    # Llenar gastos
    gastos = data.get('gastos', {})
    mapping = {
        'acpm': 'G16',
        'cargue': 'G18',
        'descargue': 'G20',
        'peajes': 'G22',
        'comision_empresa': 'G24',
        'llantas': 'G26',
        'engrase': 'G28',
        'lavada': 'G30',
        'parqueadero': 'G32',
        'carrosada': 'G34',
        'descarrrosada': 'G36',
        'otros': 'G38',
        'bonificacion': 'G40'
    }
    for key, celda in mapping.items():
        cell = ws[celda]
        main_cell = obtener_celda_principal(ws, cell)
        main_cell.value = gastos.get(key, 0)

    # Calcular totales
    flete = float(data.get('flete', 0) or 0)
    anticipo = float(data.get('anticipo', 0) or 0)
    bonificacion = float(gastos.get('bonificacion', 0) or 0)
    total_gastos = sum(float(gastos.get(k, 0) or 0) for k in mapping.keys())

    valor_viaje = flete + bonificacion
    menos_anticipo = anticipo

    saldo_a_favor = max((flete + bonificacion) - total_gastos + anticipo, 0)
    saldo_en_contra = max(total_gastos - (flete + bonificacion) - anticipo, 0)

    resumen = {
        'I41': valor_viaje,
        'I42': total_gastos,
        'I43': menos_anticipo,
        'I44': saldo_a_favor,
        'I45': saldo_en_contra
    }
    for celda, valor in resumen.items():
        cell = ws[celda]
        main_cell = obtener_celda_principal(ws, cell)
        main_cell.value = valor

    # Guardar en buffer
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer