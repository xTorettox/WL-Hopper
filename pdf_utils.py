import fitz  # PyMuPDF
import re
import os
import io

def _configurar_tesseract():
    try:
        import pytesseract
        rutas = [
            os.path.expanduser(r'~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'),
            r'C:\Program Files\Tesseract-OCR\tesseract.exe',
            r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
            '/usr/bin/tesseract',
            '/usr/local/bin/tesseract'
        ]
        for r in rutas:
            if os.path.exists(r):
                pytesseract.pytesseract.tesseract_cmd = r
                break
    except Exception:
        pass

_configurar_tesseract()

def extraer_texto_completo_pdf(ruta_pdf):
    """Extrae texto nativo o mediante OCR si el PDF es escaneado."""
    texto = ""
    try:
        doc = fitz.open(ruta_pdf)
        for page in doc:
            texto += page.get_text("text") + "\n"
        doc.close()

        # Si no hubo texto legible nativo, intentar OCR
        if not texto.strip():
            try:
                import pytesseract
                from PIL import Image
                _configurar_tesseract()
                doc = fitz.open(ruta_pdf)
                for page in doc:
                    pix = page.get_pixmap(dpi=150)
                    img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")
                    texto += pytesseract.image_to_string(img) + "\n"
                doc.close()
            except Exception:
                pass
    except Exception as e:
        print(f"Error extrayendo texto de PDF {ruta_pdf}: {e}")
    return texto

def analizar_informe_local(ruta_pdf):
    """
    Analiza un PDF de inspección leyendo las últimas dos páginas
    y buscando explícitamente "certifica que Cumple" o "certifica que No Cumple".
    Retorna el estado y las observaciones.
    """
    try:
        doc = fitz.open(ruta_pdf)
        texto = ""
        paginas = list(range(len(doc)))
        ultimas_dos = paginas[-2:] if len(paginas) >= 2 else paginas
        for i in ultimas_dos:
            texto += doc[i].get_text("text") + "\n"
        doc.close()
        
        if not texto.strip():
            return "ERROR LECTURA", "PDF sin texto extraíble."

        # Búsqueda usando regex case-insensitive
        if re.search(r'certifica\s+que\s+no\s+cumple', texto, re.IGNORECASE):
            return "NO CUMPLE", "Equipo NO cumple con los requerimientos."
        elif re.search(r'certifica\s+que\s+cumple', texto, re.IGNORECASE):
            return "CUMPLE", "Equipo cumple con los requerimientos."
        else:
            return "DESCONOCIDO", "No se encontró el texto esperado en el informe."

    except Exception as e:
        return "ERROR LECTURA", f"Fallo al leer PDF local: {str(e)[:40]}"

def extraer_datos_titulo_o_cedula(ruta_pdf):
    """
    Extrae Número de Chasis (Serie), Marca y Modelo de un documento oficial
    de Título de la Maquinaria o Cédula.
    """
    datos = {"chasis": "", "marca": "", "modelo": "", "tipo": ""}
    texto = extraer_texto_completo_pdf(ruta_pdf)
    if not texto.strip():
        return datos

    # 1. Extracción de Chasis / Serie
    # Ejemplos: "Nro. Chasis: KMTWA096LPA056772", "Chasis: 12345", "Nro. Serie: ..."
    m_chasis = re.search(r'Nro\.?\s*Chasis\s*[:\s]+([A-Z0-9]+)', texto, re.IGNORECASE)
    if not m_chasis:
        m_chasis = re.search(r'Chasis\s*[:\s]+([A-Z0-9]+)', texto, re.IGNORECASE)
    if not m_chasis:
        m_chasis = re.search(r'Nro\.?\s*Serie\s*[:\s]+([A-Z0-9]+)', texto, re.IGNORECASE)
    if not m_chasis:
        m_chasis = re.search(r'Serie\s*[:\s]+([A-Z0-9]+)', texto, re.IGNORECASE)

    if m_chasis:
        datos["chasis"] = m_chasis.group(1).strip()

    # 2. Extracción de Marca
    # Ejemplo: "Marca: 717-KOMATSU" -> "KOMATSU", "Marca: \nKOMATSU"
    m_marca = re.search(r'Marca\s*[:\s]+(?:\d+[-–]\s*)?([^\n\r]+)', texto, re.IGNORECASE)
    if m_marca:
        datos["marca"] = m_marca.group(1).strip()

    # 3. Extracción de Modelo
    # Ejemplo: "Modelo: 385-WA500-6" -> "WA500-6"
    m_modelo = re.search(r'Modelo\s*[:\s]+(?:\d+[-–]\s*)?([^\n\r]+)', texto, re.IGNORECASE)
    if m_modelo:
        datos["modelo"] = m_modelo.group(1).strip()

    # 4. Extracción de Tipo de Maquinaria
    # Ejemplo: "Tipo: 38-CARGADORA" -> "CARGADORA"
    m_tipo = re.search(r'Tipo\s*[:\s]+(?:\d+[-–]\s*)?([^\n\r]+)', texto, re.IGNORECASE)
    if m_tipo:
        datos["tipo"] = m_tipo.group(1).strip()

    return datos

def extraer_marca_modelo_de_certificado(ruta_pdf):
    """
    Intenta extraer Marca y Modelo desde un certificado de inspección (Lenor, Worklift o BV).
    """
    datos = {"marca": "", "modelo": ""}
    texto = extraer_texto_completo_pdf(ruta_pdf)
    if not texto.strip():
        return datos

    m_marca = re.search(r'Marca\s*[:\s]+([^\n\r]+)', texto, re.IGNORECASE)
    if m_marca:
        datos["marca"] = m_marca.group(1).strip()

    m_modelo = re.search(r'Modelo\s*[:\s]+([^\n\r]+)', texto, re.IGNORECASE)
    if m_modelo:
        datos["modelo"] = m_modelo.group(1).strip()

    return datos

def extraer_serie_de_certificado(ruta_pdf: str) -> str:
    """
    Intenta extraer el Número de Serie desde un certificado de inspección descargado (Worklift, BV o Lenor).
    """
    texto = extraer_texto_completo_pdf(ruta_pdf)
    if not texto.strip():
        return ""
    m_serie = re.search(r'Serie\s*[:\s]+([A-Za-z0-9\-]+)', texto, re.IGNORECASE)
    if m_serie:
        return m_serie.group(1).strip()
    return ""
