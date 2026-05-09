import re
import os
from datetime import datetime, timedelta
import pandas as pd
import io

def extraer_texto_de_archivo(archivo):
    """
    Extrae texto de un archivo BytesIO o similar para alimentar el extractor de internos.
    Soporta TXT, CSV, y XLSX.
    """
    try:
        if archivo.name.endswith('.txt'):
            return archivo.getvalue().decode('utf-8')
        elif archivo.name.endswith('.csv'):
            df = pd.read_csv(archivo)
            return df.to_string()
        elif archivo.name.endswith('.xlsx'):
            df = pd.read_excel(archivo)
            return df.to_string()
    except Exception as e:
        print(f"Error procesando archivo: {e}")
    return ""

def extraer_internos(texto_sucio):
    """
    Lógica Híbrida preservando el orden de entrada:
    1. Usa Regex para los nuevos (E03, A04, E06, etc.)
    2. Usa internos_viejos.txt para los que no cumplen el patrón nuevo.
    """
    texto_upper = texto_sucio.upper()
    
    # Cargamos la lista de viejos
    base_viejos = set()
    ruta_viejos = "internos_viejos.txt"
    if os.path.exists(ruta_viejos):
        try:
            with open(ruta_viejos, "r") as f:
                base_viejos = {line.strip().upper() for line in f if line.strip()}
        except Exception as e:
            print(f"Error al leer internos_viejos.txt: {e}")

    resultado_lista = []
    vistos = set()

    # Buscamos usando el regex para el nuevo formato
    for match in re.finditer(r'[EA]0[346]\d{4}', texto_upper):
        id_nuevo = match.group()
        if id_nuevo not in vistos:
            vistos.add(id_nuevo)
            resultado_lista.append(id_nuevo)
            
    # Buscamos en todo el texto palabras que coincidan con la base de viejos
    # Separamos el texto sucio por cualquier cosa que no sea letra o número
    palabras_en_texto = re.split(r'[^A-Z0-9]', texto_upper)
    for palabra in palabras_en_texto:
        if palabra in base_viejos and palabra not in vistos:
            vistos.add(palabra)
            resultado_lista.append(palabra)

    return resultado_lista

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
    """ Crea la carpeta temporal de descargas si no existe """
    if not os.path.exists(ruta):
        os.makedirs(ruta)
