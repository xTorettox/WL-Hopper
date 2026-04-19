import re
import os
import json
from datetime import datetime, timedelta

def extraer_internos(texto_sucio):
    """
    Regex flexible para internos viejos (4 dígitos) y nuevos (E/A + 6 dígitos).
    Limpia ruido de Excel, comas y espacios.
    """
    # Patrón: Opcional E o A, seguido de 4 a 7 dígitos.
    patron = r"(?:[EA])?\d{4,7}"
    encontrados = re.findall(patron, texto_sucio.upper())
    
    # Retorna lista sin duplicados manteniendo orden de aparición
    return list(dict.fromkeys(encontrados))

def analizar_fecha(fecha_str):
    """
    Determina si el certificado es válido, está por vencer o vencido.
    Retorna: (estado, color, puede_descargar)
    """
    if not fecha_str or fecha_str == "N/A":
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

def gestionar_config(accion="leer", data=None):
    """ Lee o guarda la configuración en el archivo JSON """
    archivo = "config.json"
    if accion == "leer":
        if os.path.exists(archivo):
            with open(archivo, "r") as f:
                return json.load(f)
        return {"usuario": "", "clave": "", "recordar": False, "ultima_ruta": "descargas"}
    
    elif accion == "guardar" and data:
        with open(archivo, "w") as f:
            json.dump(data, f, indent=4)

def asegurar_carpeta(ruta):
    if not os.path.exists(ruta):
        os.makedirs(ruta)
