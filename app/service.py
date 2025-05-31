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
    Detecta automáticamente el contorno de un documento/factura y aplica transformación de perspectiva.
    Similar a las librerías de document scanning como Dynamsoft Document Normalizer.
    """
    try:
        # Convertir bytes a imagen
        nparr = np.frombuffer(image_data, np.uint8)
        img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
        
        if img is None:
            return image_data
        
        original_img = img.copy()
        height, width = img.shape[:2]
        
        # PASO 1: Preprocesamiento avanzado para mejorar detección
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # Mejorar contraste usando CLAHE
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
        enhanced = clahe.apply(gray)
        
        # Reducir ruido manteniendo bordes
        denoised = cv2.bilateralFilter(enhanced, 9, 75, 75)
        
        # PASO 2: Detección de bordes multi-escala
        # Usar múltiples escalas para detectar mejor los bordes del documento
        edges_list = []
        
        # Escala 1: Original
        blurred1 = cv2.GaussianBlur(denoised, (5, 5), 0)
        edges1 = cv2.Canny(blurred1, 50, 150)
        edges_list.append(edges1)
        
        # Escala 2: Reducida (para documentos grandes)
        if width > 1000 or height > 1000:
            scale = 0.5
            resized = cv2.resize(denoised, (int(width*scale), int(height*scale)))
            blurred2 = cv2.GaussianBlur(resized, (3, 3), 0)
            edges2 = cv2.Canny(blurred2, 30, 120)
            edges2 = cv2.resize(edges2, (width, height))
            edges_list.append(edges2)
        
        # Escala 3: Gradiente morfológico para bordes débiles
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
        gradient = cv2.morphologyEx(denoised, cv2.MORPH_GRADIENT, kernel)
        edges3 = cv2.Canny(gradient, 40, 120)
        edges_list.append(edges3)
        
        # Combinar todas las detecciones de bordes
        combined_edges = edges_list[0]
        for edges in edges_list[1:]:
            combined_edges = cv2.bitwise_or(combined_edges, edges)
        
        # PASO 3: Morfología para conectar bordes fragmentados
        # Cerrar gaps pequeños
        kernel_close = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
        closed = cv2.morphologyEx(combined_edges, cv2.MORPH_CLOSE, kernel_close)
        
        # Dilatar ligeramente para conectar bordes cercanos
        kernel_dilate = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
        processed_edges = cv2.dilate(closed, kernel_dilate, iterations=1)
        
        # PASO 4: Detección inteligente de contornos
        contours, _ = cv2.findContours(processed_edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        if not contours:
            return image_data
        
        # Filtrar y evaluar contornos
        document_contour = find_document_contour(contours, width, height)
        
        if document_contour is None:
            return image_data
        
        # PASO 5: Aplicar transformación de perspectiva
        warped_img = apply_perspective_transform(original_img, document_contour)
        
        if warped_img is None:
            return image_data
        
        # PASO 6: Post-procesamiento para mejorar calidad
        final_img = enhance_document_image(warped_img)
        
        # Convertir de vuelta a bytes
        _, buffer = cv2.imencode('.jpg', final_img, [cv2.IMWRITE_JPEG_QUALITY, 95])
        return buffer.tobytes()
        
    except Exception as e:
        print(f"Error en detección de documento: {e}")
        return image_data

def find_document_contour(contours, width, height):
    """
    Encuentra el mejor contorno que representa un documento.
    Evalúa múltiples criterios: área, forma, posición, etc.
    """
    # Ordenar por área
    contours = sorted(contours, key=cv2.contourArea, reverse=True)
    image_area = width * height
    
    for contour in contours[:10]:  # Evaluar los 10 más grandes
        area = cv2.contourArea(contour)
        area_ratio = area / image_area
        
        # Debe ocupar al menos 10% pero no más del 95% de la imagen
        if area_ratio < 0.1 or area_ratio > 0.95:
            continue
        
        # Aproximar el contorno a un polígono
        peri = cv2.arcLength(contour, True)
        
        # Probar diferentes niveles de aproximación
        for epsilon in [0.01, 0.02, 0.03, 0.05]:
            approx = cv2.approxPolyDP(contour, epsilon * peri, True)
            
            # Buscar contornos de 4 puntos (rectangulares)
            if len(approx) == 4:
                # Verificar que sea un cuadrilátero convexo
                if cv2.isContourConvex(approx):
                    return approx
            
            # Si tiene más de 4 puntos, intentar encontrar el rectángulo mínimo
            elif len(approx) > 4 and len(approx) <= 8:
                # Usar rectángulo delimitador rotado
                rect = cv2.minAreaRect(contour)
                box = cv2.boxPoints(rect)
                box = np.int0(box)
                
                # Verificar que el rectángulo tenga un tamaño razonable
                rect_area = cv2.contourArea(box)
                if rect_area / image_area > 0.1:
                    return box
    
    return None

def apply_perspective_transform(img, contour):
    """
    Aplica transformación de perspectiva para enderezar el documento.
    """
    if contour is None or len(contour) != 4:
        return None
    
    # Convertir a puntos float32
    pts = contour.reshape(4, 2).astype(np.float32)
    
    # Ordenar puntos: top-left, top-right, bottom-right, bottom-left
    rect = order_points(pts)
    
    # Calcular dimensiones del documento enderezado
    (tl, tr, br, bl) = rect
    
    # Calcular ancho
    widthA = np.sqrt(((br[0] - bl[0]) ** 2) + ((br[1] - bl[1]) ** 2))
    widthB = np.sqrt(((tr[0] - tl[0]) ** 2) + ((tr[1] - tl[1]) ** 2))
    maxWidth = max(int(widthA), int(widthB))
    
    # Calcular alto
    heightA = np.sqrt(((tr[0] - br[0]) ** 2) + ((tr[1] - br[1]) ** 2))
    heightB = np.sqrt(((tl[0] - bl[0]) ** 2) + ((tl[1] - bl[1]) ** 2))
    maxHeight = max(int(heightA), int(heightB))
    
    # Verificar dimensiones mínimas
    if maxWidth < 50 or maxHeight < 50:
        return None
    
    # Puntos de destino (rectángulo perfecto)
    dst = np.array([
        [0, 0],
        [maxWidth - 1, 0],
        [maxWidth - 1, maxHeight - 1],
        [0, maxHeight - 1]
    ], dtype="float32")
    
    # Calcular matriz de transformación
    matrix = cv2.getPerspectiveTransform(rect, dst)
    
    # Aplicar transformación
    warped = cv2.warpPerspective(img, matrix, (maxWidth, maxHeight))
    
    return warped

def order_points(pts):
    """
    Ordena los puntos en el orden: top-left, top-right, bottom-right, bottom-left
    """
    rect = np.zeros((4, 2), dtype="float32")
    
    # Suma de coordenadas: top-left tendrá la suma más pequeña, bottom-right la más grande
    s = pts.sum(axis=1)
    rect[0] = pts[np.argmin(s)]  # top-left
    rect[2] = pts[np.argmax(s)]  # bottom-right
    
    # Diferencia de coordenadas: top-right tendrá la diferencia más pequeña, bottom-left la más grande
    diff = np.diff(pts, axis=1)
    rect[1] = pts[np.argmin(diff)]  # top-right
    rect[3] = pts[np.argmax(diff)]  # bottom-left
    
    return rect

def enhance_document_image(img):
    """
    Mejora la calidad de la imagen del documento después del crop.
    """
    # Convertir a escala de grises para análisis
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # Verificar si es mejor en escala de grises o color
    # Si tiene poco color, convertir a escala de grises puede mejorar legibilidad
    
    # Mejorar contraste y brillo ligeramente
    enhanced = cv2.convertScaleAbs(img, alpha=1.1, beta=5)
    
    # Aplicar un ligero filtro de nitidez si la imagen se ve borrosa
    kernel = np.array([[-1,-1,-1], [-1,9,-1], [-1,-1,-1]]) * 0.3
    kernel[1,1] = kernel[1,1] + 0.7  # Reducir intensidad del sharpening
    sharpened = cv2.filter2D(enhanced, -1, kernel)
    
    return sharpened