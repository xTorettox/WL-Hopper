import re
import os
from datetime import datetime, timedelta
import pandas as pd
import io

def extraer_texto_de_archivo(archivo):
    """
    Extrae texto de un archivo BytesIO o similar para alimentar el extractor de internos.
    Soporta TXT, CSV, XLSX y PDF (con OCR de fallback para PDFs escaneados).
    """
    try:
        name_lower = archivo.name.lower()
        if name_lower.endswith('.txt'):
            return archivo.getvalue().decode('utf-8')
        elif name_lower.endswith('.csv'):
            df = pd.read_csv(archivo)
            return df.to_string()
        elif name_lower.endswith('.xlsx'):
            df = pd.read_excel(archivo)
            return df.to_string()
        elif name_lower.endswith('.pdf'):
            import fitz
            # Posicionamos al inicio de los bytes
            pos_inicial = archivo.tell()
            archivo.seek(0)
            pdf_bytes = archivo.read()
            archivo.seek(pos_inicial)
            
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            texto = ""
            for pagina in doc:
                texto += pagina.get_text()
                
            # Si el PDF no contenía texto legible nativo (ej: escaneado), aplicamos OCR a las páginas
            if not texto.strip():
                print("[INFO] PDF no contiene texto nativo. Aplicando OCR...")
                import pytesseract
                from PIL import Image
                for i in range(len(doc)):
                    pagina = doc.load_page(i)
                    pix = pagina.get_pixmap(dpi=150)
                    img_data = pix.tobytes("png")
                    img = Image.open(io.BytesIO(img_data)).convert('L')
                    # Preprocesamiento de la imagen para mejorar la precisión del OCR
                    w, h = img.size
                    img = img.resize((w*2, h*2), Image.Resampling.LANCZOS)
                    texto += pytesseract.image_to_string(img, config='--psm 11')
            return texto
    except Exception as e:
        print(f"Error procesando archivo: {e}")
    return ""

def extraer_internos(texto_sucio):
    """
    Lógica Híbrida Ultra-Robusta: Busca internos nuevos ([EA] + 6 dígitos)
    y potenciales internos viejos en el texto, incluso si están concatenados,
    pegados por guiones (E061633-A040230) o sin espacios, preservando el orden original.
    """
    texto_upper = texto_sucio.upper()
    
    base_viejos = set()
    ruta_viejos = "internos_viejos.txt"
    if os.path.exists(ruta_viejos):
        try:
            with open(ruta_viejos, "r", encoding="utf-8") as f:
                base_viejos = {line.strip().upper() for line in f if line.strip()}
        except Exception as e:
            print(f"Error al leer internos_viejos.txt: {e}")

    matches = []
    
    # 1. Encontrar todos los internos nuevos [EA]\d{6} en cualquier posición
    for m in re.finditer(r'[EA]\d{6}', texto_upper):
        matches.append((m.start(), m.group()))
        
    # 2. Encontrar todos los internos viejos respetando límites de palabra para evitar matches parciales
    for viejo in base_viejos:
        patron_viejo = r'\b' + re.escape(viejo) + r'\b'
        for m in re.finditer(patron_viejo, texto_upper):
            matches.append((m.start(), m.group()))
            
    # Ordenar por el índice donde aparecieron en el texto
    matches.sort(key=lambda x: x[0])
    
    # Eliminar duplicados manteniendo el orden
    resultado = []
    vistos = set()
    for _, val in matches:
        if val not in vistos:
            resultado.append(val)
            vistos.add(val)
            
    return resultado

def analizar_fecha(fecha_str):
    """
    Determina el estado del certificado estándar.
    Retorna: (estado, color, puede_descargar)
    """
    if not fecha_str or fecha_str in ["N/A", "SIN REGISTRO"]:
        return "SIN REGISTRO", "gray", False
        
    try:
        # Formato Worklift: DD/MM/YYYY
        fecha_venc = datetime.strptime(fecha_str, "%d/%m/%Y")
        hoy = datetime.now()
        margen_30 = hoy + timedelta(days=30)
        
        if fecha_venc < hoy:
            return "VENCIDO", "red", False
        elif hoy <= fecha_venc <= margen_30:
            return "PRÓXIMO A VENCER", "orange", True
        else:
            return "VIGENTE", "green", True
    except ValueError:
        return "ERROR FORMATO", "gray", False

def calcular_vencimiento_semestral(fecha_insp_str):
    """
    Suma 180 días a la fecha de inspección y analiza el estado como semestral.
    Retorna: (estado_semestral, fecha_semestral_str)
    """
    if not fecha_insp_str or fecha_insp_str in ["N/A", "-", "SIN REGISTRO"]:
        return "-", "-"
    
    try:
        fecha_insp = datetime.strptime(fecha_insp_str, "%d/%m/%Y")
        fecha_semestral = fecha_insp + timedelta(days=180)
        fecha_semestral_str = fecha_semestral.strftime("%d/%m/%Y")
        
        hoy = datetime.now()
        margen_30 = hoy + timedelta(days=30)
        
        if fecha_semestral < hoy:
            estado = "VENCIDO (SEM)"
        elif hoy <= fecha_semestral <= margen_30:
            estado = "PRÓXIMO A VENCER (SEM)"
        else:
            estado = "VIGENTE (SEM)"
            
        return estado, fecha_semestral_str
    except ValueError:
        return "-", "-"

def asegurar_carpeta(ruta):
