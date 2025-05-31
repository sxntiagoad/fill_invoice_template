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
        print("🔍 Iniciando detección de documento...")
        # Convertir bytes a imagen
        nparr = np.frombuffer(image_data, np.uint8)
        img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
        
        if img is None:
            print("❌ No se pudo decodificar la imagen")
            return image_data
        
        print(f"📏 Imagen original: {img.shape[1]}x{img.shape[0]}")
        original_img = img.copy()
        height, width = img.shape[:2]
        
        # Redimensionar para procesamiento más rápido (manteniendo aspect ratio)
        resize_height = 800
        if height > resize_height:
            ratio = resize_height / height
            resize_width = int(width * ratio)
            img_resized = cv2.resize(img, (resize_width, resize_height))
            print(f"📐 Imagen redimensionada: {resize_width}x{resize_height}")
        else:
            img_resized = img.copy()
            ratio = 1.0
            print("📐 Imagen no redimensionada")
        
        # Convertir a escala de grises
        gray = cv2.cvtColor(img_resized, cv2.COLOR_BGR2GRAY)
        
        # Mejorar el contraste
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
        gray = clahe.apply(gray)
        
        # Aplicar múltiples métodos de detección de bordes
        
        # Método 1: Canny tradicional
        blurred = cv2.GaussianBlur(gray, (5, 5), 0)
        edged1 = cv2.Canny(blurred, 50, 150)
        
        # Método 2: Canny más agresivo
        edged2 = cv2.Canny(blurred, 30, 80)
        
        # Método 3: Gradiente morfológico
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
        gradient = cv2.morphologyEx(gray, cv2.MORPH_GRADIENT, kernel)
        edged3 = cv2.Canny(gradient, 50, 150)
        
        # Combinar todos los métodos
        combined_edges = cv2.bitwise_or(cv2.bitwise_or(edged1, edged2), edged3)
        
        # Cerrar gaps en los contornos
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5))
        closed = cv2.morphologyEx(combined_edges, cv2.MORPH_CLOSE, kernel)
        
        # Encontrar contornos
        contours, _ = cv2.findContours(closed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        if not contours:
            print("❌ No se encontraron contornos")
            return image_data
        
        print(f"🔍 Encontrados {len(contours)} contornos")
        
        # Ordenar contornos por área
        contours = sorted(contours, key=cv2.contourArea, reverse=True)
        
        document_contour = None
        
        # Intentar encontrar el mejor contorno
        for i, contour in enumerate(contours[:10]):  # Revisar los 10 contornos más grandes
            area = cv2.contourArea(contour)
            area_percentage = (area / (img_resized.shape[0] * img_resized.shape[1])) * 100
            
            print(f"📊 Contorno {i+1}: área={area:.0f} ({area_percentage:.1f}%)")
            
            # El contorno debe ocupar al menos 15% de la imagen
            if area < (img_resized.shape[0] * img_resized.shape[1] * 0.15):
                continue
                
            # Calcular el perímetro del contorno
            peri = cv2.arcLength(contour, True)
            
            # Probar diferentes niveles de aproximación
            for epsilon_factor in [0.01, 0.02, 0.03, 0.05, 0.08]:
                approx = cv2.approxPolyDP(contour, epsilon_factor * peri, True)
                
                # Si encontramos un contorno rectangular o casi rectangular
                if len(approx) >= 4 and len(approx) <= 8:
                    print(f"✅ Contorno válido encontrado: {len(approx)} puntos")
                    # Para contornos con más de 4 puntos, tomar los 4 esquinas principales
                    if len(approx) > 4:
                        # Calcular el rectángulo delimitador
                        rect = cv2.minAreaRect(contour)
                        box = cv2.boxPoints(rect)
                        approx = np.int0(box)
                    
                    document_contour = approx
                    break
            
            if document_contour is not None:
                break
        
        # Si no encontramos un buen contorno, usar el contorno más grande
        if document_contour is None and contours:
            largest_contour = contours[0]
            largest_area = cv2.contourArea(largest_contour)
            area_percentage = (largest_area / (img_resized.shape[0] * img_resized.shape[1])) * 100
            
            print(f"🔄 Usando contorno más grande: {area_percentage:.1f}%")
            
            if largest_area > (img_resized.shape[0] * img_resized.shape[1] * 0.3):
                rect = cv2.minAreaRect(largest_contour)
                box = cv2.boxPoints(rect)
                document_contour = np.int0(box)
        
        # Procesar el contorno encontrado
        if document_contour is not None and len(document_contour) >= 4:
            print("🎯 Aplicando transformación de perspectiva...")
            
            # Escalar de vuelta al tamaño original
            if ratio != 1.0:
                document_contour = document_contour / ratio
                document_contour = document_contour.astype(np.int32)
            
            # Tomar solo los primeros 4 puntos si hay más
            if len(document_contour) > 4:
                document_contour = document_contour[:4]
            
            # Ordenar los puntos del contorno
            pts = document_contour.reshape(4, 2).astype(np.float32)
            
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
            
            print(f"📐 Documento detectado: {maxWidth}x{maxHeight}")
            
            # Verificar que las dimensiones sean razonables
            if maxWidth > 50 and maxHeight > 50:
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
                
                # Mejorar la imagen resultante
                warped = cv2.convertScaleAbs(warped, alpha=1.1, beta=10)
                
                # Convertir de vuelta a bytes
                _, buffer = cv2.imencode('.jpg', warped, [cv2.IMWRITE_JPEG_QUALITY, 95])
                print("✅ Documento recortado exitosamente")
                return buffer.tobytes()
            else:
                print("❌ Dimensiones del documento muy pequeñas")
        else:
            print("❌ No se encontró contorno válido para el documento")
        
        # Si no se detectó documento, devolver imagen original
        print("🔄 Devolviendo imagen original")
        return image_data
        
    except Exception as e:
        # En caso de error, devolver imagen original
        print(f"❌ Error procesando imagen: {e}")
        return image_data