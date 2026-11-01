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

# Lista de marcas conocidas de maquinaria pesada e industrial para validación
MARCAS_CONOCIDAS = {
    "TOYOTA", "MANITOU", "KOMATSU", "JLG", "CATERPILLAR", "CAT",
    "BOBCAT", "SULLAIR", "GENIE", "HYSTER", "YALE", "NISSAN",
    "MITSUBISHI", "CASE", "NEW HOLLAND", "JOHN DEERE", "HAULOTTE",
    "LINDE", "CLARK", "TCM", "DOOSAN", "HELI", "HANGCHA", "HYUNDAI",
    "VOLVO", "JCB", "LIEBHERR", "GROVE", "TEREX", "TADANO", "SNORKEL",
    "SKYJACK", "DINGLI", "XCMG", "SANY", "LIUGONG", "LONKING", "KALMAR"
}

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
    de Título de la Maquinaria, Cédula MAVI o Factura de Compra.
    Diferencia estrictamente 'Mca. Chasis' (Marca) de 'Nro. Chasis' (Número real).
    """
    datos = {"chasis": "", "marca": "", "modelo": "", "tipo": ""}
    texto = extraer_texto_completo_pdf(ruta_pdf)
    if not texto.strip():
        return datos

    # 1. Extracción de Chasis / Serie
    # En títulos y cédulas DNRPA/MAVI existen dos campos diferentes:
    #   - "Mca. Chasis: TOYOTA" -> Marca del chasis (NO es el número de serie).
    #   - "Nro. Chasis: 608FD25-74043" -> Número del chasis (ES el número de serie).
    patrones_chasis = [
        # 1.1 Variantes explícitas de Nro. Chasis / N° Chasis / Nº Chasis / No. Chasis / Número de Chasis
        r'(?:Nro\.?|N[°ºo]\.?|N[uú]m(?:ero)?\.?)\s*(?:de\s*)?Chasis\s*[:\s]+([A-Za-z0-9\-]+)',
        # 1.2 Variantes explícitas de Nro. Serie / N° Serie / Nº Serie / Número de Serie
        r'(?:Nro\.?|N[°ºo]\.?|N[uú]m(?:ero)?\.?)\s*(?:de\s*)?Serie\s*[:\s]+([A-Za-z0-9\-]+)',
        # 1.3 Chasis: (con lookbehind negativo para asegurar que NO esté precedido por Mca. o Marca)
        r'(?<!Mca\.\s)(?<!Mca\.)(?<!Mca\s)(?<!Marca\s)(?<!Mca)(?<!Marca)\bChasis\s*(?:Nro\.?|N[°ºo]\.?|N[uú]m(?:ero)?\.?)?\s*[:\s]+([A-Za-z0-9\-]+)',
        # 1.4 Serie: (con lookbehind negativo)
        r'(?<!Mca\.\s)(?<!Mca\.)(?<!Mca\s)(?<!Marca\s)\bSerie\s*(?:Nro\.?|N[°ºo]\.?|N[uú]m(?:ero)?\.?)?\s*[:\s]+([A-Za-z0-9\-]+)',
        # 1.5 VIN o Bastidor
        r'(?:VIN|Bastidor|Cuadro)\s*[:\s]+([A-Za-z0-9\-]+)'
    ]

    for pat in patrones_chasis:
        for m in re.finditer(pat, texto, re.IGNORECASE):
            val = m.group(1).strip(" ,;.:-_")
            # Validación estricta:
            # - Debe tener al menos 3 caracteres
            # - Debe contener al menos un dígito
            # - NO debe ser una marca conocida (ej: "TOYOTA", "KOMATSU")
            if len(val) >= 3 and any(c.isdigit() for c in val) and val.upper() not in MARCAS_CONOCIDAS:
                datos["chasis"] = val
                break
        if datos["chasis"]:
            break

    # 2. Extracción de Marca
    # Ejemplo: "Marca: 130-TOYOTA" -> "TOYOTA", "Marca: \nKOMATSU", "Mca. Chasis: TOYOTA"
    m_marca = re.search(r'(?:^|[^\w])Marca\s*[:\s]+(?:\d+[-–]\s*)?([^\n\r,;]+)', texto, re.IGNORECASE)
    if not m_marca:
        m_marca = re.search(r'Mca\.?\s*(?:Chasis|Motor)?\s*[:\s]+([^\n\r,;]+)', texto, re.IGNORECASE)
    if m_marca:
        val_m = m_marca.group(1).strip(" ,;.:-_")
        val_m = re.sub(r'^\d{3,4}[-–]\s*', '', val_m)
        val_m = re.split(r'\s{2,}', val_m)[0].strip(" ,;.:-_")
        datos["marca"] = val_m

    # 3. Extracción de Modelo
    # Ejemplo: "Modelo: 652-62-8FD25" -> "62-8FD25"
    m_modelo = re.search(r'Modelo\s*[:\s]+([^\n\r,;]+)', texto, re.IGNORECASE)
    if m_modelo:
        val_mod = m_modelo.group(1).strip(" ,;.:-_")
        val_mod = re.sub(r'^\d{3,4}[-–]\s*', '', val_mod)
        val_mod = re.split(r'\s{2,}', val_mod)[0].strip(" ,;.:-_")
        datos["modelo"] = val_mod

    # 4. Extracción de Tipo de Maquinaria
    # Ejemplo: "Tipo: 37-AUTOELEVADORA" -> "AUTOELEVADORA"
    m_tipo = re.search(r'Tipo\s*[:\s]+([^\n\r,;]+)', texto, re.IGNORECASE)
    if m_tipo:
        val_t = m_tipo.group(1).strip(" ,;.:-_")
        val_t = re.sub(r'^\d{3,4}[-–]\s*', '', val_t)
        val_t = re.split(r'\s{2,}', val_t)[0].strip(" ,;.:-_")
        datos["tipo"] = val_t

    return datos

def extraer_marca_modelo_de_certificado(ruta_pdf):
    """
    Intenta extraer Marca y Modelo desde un certificado de inspección (Lenor, Worklift o BV).
    """
    datos = {"marca": "", "modelo": ""}
    texto = extraer_texto_completo_pdf(ruta_pdf)
    if not texto.strip():
        return datos

    m_marca = re.search(r'Marca\s*[:\s]+([^\n\r,;]+)', texto, re.IGNORECASE)
    if m_marca:
        datos["marca"] = re.split(r'\s{2,}', m_marca.group(1))[0].strip(" ,;.:-_")

    m_modelo = re.search(r'Modelo\s*[:\s]+([^\n\r,;]+)', texto, re.IGNORECASE)
    if m_modelo:
        datos["modelo"] = re.split(r'\s{2,}', m_modelo.group(1))[0].strip(" ,;.:-_")

    return datos

def extraer_serie_de_certificado(ruta_pdf: str) -> str:
    """
    Intenta extraer el Número de Serie desde un certificado de inspección descargado (Worklift, BV o Lenor).
    Valida que no sea un nombre de marca y contenga dígitos reales.
    """
    texto = extraer_texto_completo_pdf(ruta_pdf)
    if not texto.strip():
        return ""

    patrones_serie = [
        r'(?:Nro\.?|N[°ºo]\.?|N[uú]m(?:ero)?\.?)?\s*Serie\s*[:\s]+([A-Za-z0-9\-]+)',
        r'(?:Nro\.?|N[°ºo]\.?|N[uú]m(?:ero)?\.?)?\s*Chasis\s*[:\s]+([A-Za-z0-9\-]+)',
        r'N[uú]mero\s+de\s+Serie\s*[:\s]+([A-Za-z0-9\-]+)',
        r'Serie/Chasis\s*[:\s]+([A-Za-z0-9\-]+)'
    ]
    for pat in patrones_serie:
        for m in re.finditer(pat, texto, re.IGNORECASE):
            val = m.group(1).strip(" ,;.:-_")
            if len(val) >= 3 and any(c.isdigit() for c in val) and val.upper() not in MARCAS_CONOCIDAS:
                return val
    return ""
