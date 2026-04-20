import re
import os
from datetime import datetime, timedelta

def extraer_internos(texto_sucio):
    """
    Lógica Híbrida:
    1. Usa Regex para los nuevos (E03, A04, E06, etc.)
    2. Usa internos_viejos.txt para los que no cumplen el patrón nuevo.
    """
    texto_upper = texto_sucio.upper()
    
    # --- 1. REGEX PARA NOMENCLATURA NUEVA ---
    # Patrón: E o A + 0 + (3, 4 o 6) + 4 dígitos. Ejemplo: E040230
    patron_nuevo = r'\b[EA]0[346]\d{4}\b'
    encontrados_nuevos = set(re.findall(patron_nuevo, texto_upper))
    
    # --- 2. BÚSQUEDA DE INTERNOS VIEJOS (Lista de Oro) ---
    encontrados_viejos = set()
    ruta_viejos = "internos_viejos.txt"
    
    if os.path.exists(ruta_viejos):
        try:
            with open(ruta_viejos, "r") as f:
                # Cargamos set limpio (sin espacios ni líneas vacías)
                base_viejos = {line.strip().upper() for line in f if line.strip()}
            
            # Separamos el texto sucio por cualquier cosa que no sea letra o número
            palabras_en_texto = set(re.split(r'[^A-Z0-9]', texto_upper))
            # Intersección: solo los que están en el texto Y en la lista de viejos
            encontrados_viejos = palabras_en_texto.intersection(base_viejos)
        except Exception as e:
            print(f"Error al leer internos_viejos.txt: {e}")

    # Unimos ambos, eliminamos duplicados y ordenamos
    resultado = sorted(list(encontrados_nuevos.union(encontrados_viejos)))
    return resultado

def analizar_fecha(fecha_str):
    """
    Determina el estado del certificado.
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

def asegurar_carpeta(ruta):
    """ Crea la carpeta temporal de descargas si no existe """
    if not os.path.exists(ruta):
        os.makedirs(ruta)
