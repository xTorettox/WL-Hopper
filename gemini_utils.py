import streamlit as st
import google.generativeai as genai
import fitz  # PyMuPDF
import json

def configurar_gemini():
    """Configura Gemini de forma segura y gratuita."""
    try:
        # El error del strip suele venir de intentar manipular el objeto secrets directamente.
        # Aquí lo llamamos de la forma más estable posible.
        api_key = st.secrets["GOOGLE_API_KEY"]
        genai.configure(api_key=api_key)
        return True
    except Exception as e:
        st.error(f"Error con la API Key en Secrets: {e}")
        return False

def ocr_bruto_gemini(imagen_bytes):
    """OCR gratuito para extraer internos de fotos."""
    if not configurar_gemini(): return ""
    
    try:
        model = genai.GenerativeModel('gemini-1.5-flash')
        # Pasamos la imagen directamente como bytes para evitar líos de librerías
        response = model.generate_content([
            "Transcribe solo los códigos de internos (ej: E030193, 3797).",
            {"mime_type": "image/jpeg", "data": imagen_bytes}
        ])
        return response.text
    except Exception as e:
        print(f"Error en OCR: {e}")
        return ""

def analizar_informe_gemini(ruta_pdf):
    """Analiza el PDF para detectar equipos rechazados (Gratis)."""
    if not configurar_gemini(): return "ERROR", "Sin API Key"
    
    try:
        doc = fitz.open(ruta_pdf)
        texto = ""
        for i in range(min(len(doc), 3)): # Solo leemos las primeras 3 páginas para no gastar tokens
            texto += doc[i].get_text()
        doc.close()

        model = genai.GenerativeModel('gemini-1.5-flash')
        prompt = f"Analizá este informe técnico y decime si el equipo CUMPLE o NO CUMPLE. Si no cumple, resumí el porqué. Texto: {texto}"
        
        response = model.generate_content(prompt)
        # Una respuesta simple para evitar errores de parseo de JSON
        return "REVISIÓN IA", response.text[:200]
    except Exception as e:
        return "ERROR IA", str(e)[:50]
