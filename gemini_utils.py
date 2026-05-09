from google import genai
import fitz  # PyMuPDF
import json
import streamlit as st
import time

def configurar_gemini():
    """
    Configura el cliente de Gemini. 
    Busca la API Key en st.secrets de todas las formas posibles.
    """
    api_key = None
    try:
        # 1. Intento: Formato simple (Clave: GOOGLE_API_KEY, Valor: tu_key)
        api_key = st.secrets.get("GOOGLE_API_KEY")
        
        # 2. Intento: Si es un diccionario (Formato: [GOOGLE_API_KEY] -> GOOGLE_API_KEY = "tu_key")
        if isinstance(api_key, dict):
            api_key = api_key.get("GOOGLE_API_KEY")
            
        if api_key:
            return genai.Client(api_key=api_key)
            
    except Exception as e:
        st.error(f"Error técnico con los secretos: {e}")
    
    return None

def analizar_informe_gemini(ruta_pdf):
    client = configurar_gemini()
    if not client:
        return "ERROR: Sin API Key", "No se encontró la clave en st.secrets."

    try:
        doc = fitz.open(ruta_pdf)
        texto = ""
        for i in range(len(doc)):
            if i > 2: break # Primeras 3 páginas
            texto += doc[i].get_text("text") + "\n"
        doc.close()
        
        if not texto.strip():
            return "ERROR LECTURA", "PDF sin texto extraíble."

        prompt = """
        Sos un experto técnico de Sullair Argentina. Analizá el texto del informe:
        1. Estado: VIGENTE, RECHAZADO o REQUIERE REPARACION.
        2. Observaciones: Breve resumen técnico.
        Responde solo JSON: {"estado": "...", "observaciones": "..."}
        """

        # Llamada con la nueva librería
        response = client.models.generate_content(
            model='gemini-1.5-flash',
            contents=[prompt, texto]
        )
        
        resp_text = response.text.replace('```json', '').replace('```', '').strip()
        data = json.loads(resp_text)
        
        return data.get('estado', 'DESCONOCIDO'), data.get('observaciones', '-')
        
    except Exception as e:
        return "ERROR IA", f"Fallo: {str(e)[:40]}"

def extraer_internos_imagen(imagen_bytes):
    """
    Lee los internos (E040328, etc.) directamente de una foto.
    """
    client = configurar_gemini()
    if not client:
        return []

    try:
        prompt = "Listá todos los números de internos de equipos en la imagen (ej: E040328). Solo devolvé los códigos separados por coma."
        
        # Nueva sintaxis multimodal
        response = client.models.generate_content(
            model='gemini-1.5-flash',
            contents=[
                prompt,
                {"mime_type": "image/jpeg", "data": imagen_bytes}
            ]
        )
        
        return [item.strip() for item in response.text.split(',') if item.strip()]
    except Exception:
        return []

def ocr_bruto_gemini(imagen_bytes):
    """
    Usa Gemini 1.5 Flash como OCR puro. Solo transcribe texto, 
    no analiza. El resultado se pasa a extraer_internos.
    """
    client = configurar_gemini()
    if not client:
        return ""

    try:
        prompt = "Transcribe todo el texto visible en esta imagen, especialmente códigos alfanuméricos."
        
        response = client.models.generate_content(
            model='gemini-1.5-flash',
            contents=[
                prompt,
                {"mime_type": "image/jpeg", "data": imagen_bytes}
            ]
        )
        return response.text
    except Exception as e:
        print(f"Error OCR: {e}")
        return ""
