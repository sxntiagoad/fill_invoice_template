import io
import os
import cv2
import numpy as np
from openpyxl import load_workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from PIL import Image

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

def generate_invoice_pdf(images_data):
    """
    Genera un PDF con las imágenes de facturas organizadas dinámicamente manteniendo su proporción.
    images_data: Lista de bytes de las imágenes
    """
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    # Configuración de márgenes y espaciado
    margin = 30
    spacing = 15
    max_width = width - (2 * margin)
    max_height = height - (2 * margin)
    
    current_x = margin
    current_y = height - margin
    current_row_height = 0
    page_number = 1

    for img_data in images_data:
        # NUEVO: Detectar y recortar automáticamente el documento
        cropped_img_data = detect_and_crop_document(img_data)
        
        # Procesar la imagen para obtener sus dimensiones originales
        img = Image.open(io.BytesIO(cropped_img_data))
        original_width, original_height = img.size
        aspect_ratio = original_width / original_height

        # Determinar si la imagen es muy pequeña (necesita agrandarse más)
        is_small_image = min(original_width, original_height) < 800

        # Calcular el tamaño final manteniendo la proporción
        if aspect_ratio > 1.5:  # Imagen muy horizontal
            base_width = max_width * (0.7 if is_small_image else 0.55)
            final_width = min(base_width, max_width * 0.8)
            final_height = final_width / aspect_ratio
        elif aspect_ratio < 0.7:  # Imagen muy vertical
            base_height = max_height * (0.6 if is_small_image else 0.4)
            final_height = min(base_height, max_height * 0.7)
            final_width = final_height * aspect_ratio
        else:  # Imagen cuadrada o proporción normal
            if is_small_image:
                # Para imágenes pequeñas, usar más espacio
                base_size = min(max_width * 0.6, max_height * 0.5)
            else:
                base_size = min(max_width * 0.45, max_height * 0.35)
            
            if aspect_ratio > 1:
                final_width = base_size
                final_height = final_width / aspect_ratio
            else:
                final_height = base_size
                final_width = final_height * aspect_ratio

        # Verificar si la imagen cabe en la fila actual
        if current_x + final_width > width - margin:
            # Pasar a la siguiente fila
            current_x = margin
            current_y -= (current_row_height + spacing)
            current_row_height = 0

        # Verificar si la imagen cabe en la página actual
        if current_y - final_height < margin:
            # Crear nueva página
            c.showPage()
            current_x = margin
            current_y = height - margin
            current_row_height = 0
            page_number += 1

        # Dibujar la imagen (ahora usando la imagen recortada)
        y_position = current_y - final_height
        img_reader = ImageReader(img)
        c.drawImage(img_reader, current_x, y_position, width=final_width, height=final_height)

        # Actualizar posiciones
        current_x += final_width + spacing
        current_row_height = max(current_row_height, final_height)

    c.save()
    buffer.seek(0)
    return buffer

def detect_and_crop_document(image_data):
    """
    Detecta automáticamente el contorno de un documento/factura en una imagen y lo recorta.
    Elimina el fondo (mesa, manos, etc.) y devuelve solo el papel de la factura.
    """
    try:
        # Convertir bytes a imagen
        nparr = np.frombuffer(image_data, np.uint8)
        img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
        
        if img is None:
            # Si no se puede procesar, devolver la imagen original
            return image_data
        
        original_img = img.copy()
        height, width = img.shape[:2]
        
        # Convertir a escala de grises
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # Aplicar filtro Gaussiano para reducir ruido
        blurred = cv2.GaussianBlur(gray, (5, 5), 0)
        
        # Detectar bordes usando Canny
        edged = cv2.Canny(blurred, 75, 200)
        
        # Encontrar contornos
        contours, _ = cv2.findContours(edged, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        # Ordenar contornos por área (el más grande primero)
        contours = sorted(contours, key=cv2.contourArea, reverse=True)
        
        document_contour = None
        
        # Buscar el contorno rectangular más grande (probablemente la factura)
        for contour in contours:
            # Calcular el perímetro del contorno
            peri = cv2.arcLength(contour, True)
            # Aproximar el contorno
            approx = cv2.approxPolyDP(contour, 0.02 * peri, True)
            
            # Si encontramos un contorno con 4 puntos y área suficiente
            if len(approx) == 4 and cv2.contourArea(contour) > (width * height * 0.1):
                document_contour = approx
                break
        
        # Si no encontramos un contorno rectangular, intentar con el contorno más grande
        if document_contour is None and contours:
            largest_contour = contours[0]
            if cv2.contourArea(largest_contour) > (width * height * 0.2):
                peri = cv2.arcLength(largest_contour, True)
                document_contour = cv2.approxPolyDP(largest_contour, 0.05 * peri, True)
        
        # Si encontramos un contorno válido, recortar la imagen
        if document_contour is not None and len(document_contour) >= 4:
            # Ordenar los puntos del contorno
            pts = document_contour.reshape(4, 2)
            
            # Ordenar puntos: top-left, top-right, bottom-right, bottom-left
            rect = np.zeros((4, 2), dtype="float32")
            
            s = pts.sum(axis=1)
            rect[0] = pts[np.argmin(s)]  # top-left
            rect[2] = pts[np.argmax(s)]  # bottom-right
            
            diff = np.diff(pts, axis=1)
            rect[1] = pts[np.argmin(diff)]  # top-right
            rect[3] = pts[np.argmax(diff)]  # bottom-left
            
            # Calcular las dimensiones del documento
            (tl, tr, br, bl) = rect
            widthA = np.sqrt(((br[0] - bl[0]) ** 2) + ((br[1] - bl[1]) ** 2))
            widthB = np.sqrt(((tr[0] - tl[0]) ** 2) + ((tr[1] - tl[1]) ** 2))
            maxWidth = max(int(widthA), int(widthB))
            
            heightA = np.sqrt(((tr[0] - br[0]) ** 2) + ((tr[1] - br[1]) ** 2))
            heightB = np.sqrt(((tl[0] - bl[0]) ** 2) + ((tl[1] - bl[1]) ** 2))
            maxHeight = max(int(heightA), int(heightB))
            
            # Definir los puntos de destino para la transformación de perspectiva
            dst = np.array([
                [0, 0],
                [maxWidth - 1, 0],
                [maxWidth - 1, maxHeight - 1],
                [0, maxHeight - 1]
            ], dtype="float32")
            
            # Aplicar transformación de perspectiva
            matrix = cv2.getPerspectiveTransform(rect, dst)
            warped = cv2.warpPerspective(original_img, matrix, (maxWidth, maxHeight))
            
            # Convertir de vuelta a bytes
            _, buffer = cv2.imencode('.jpg', warped, [cv2.IMWRITE_JPEG_QUALITY, 95])
            return buffer.tobytes()
        
        # Si no se detectó documento, devolver imagen original
        return image_data
        
    except Exception as e:
        # En caso de error, devolver imagen original
        print(f"Error procesando imagen: {e}")
        return image_data