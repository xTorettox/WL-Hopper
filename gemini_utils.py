from google import genai
import fitz  # PyMuPDF
import json
import streamlit as st
import time

def configurar_gemini():
    # Probamos todas las variantes posibles
    api_key = st.secrets.get("GOOGLE_API_KEY")
    
    # Si lo que trajo es un diccionario, entramos un nivel más
    if isinstance(api_key, dict):
        api_key = api_key.get("GOOGLE_API_KEY")
    
    if api_key:
        return genai.Client(api_key=api_key)
    return None
    
def analizar_informe_gemini(ruta_pdf):
    client = configurar_gemini()
    if not client:
        return "ERROR: Sin API Key", "No se encontró la clave en st.secrets."

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

        prompt = """
        Sos un experto técnico de Sullair Argentina. Analizá el texto de las últimas páginas del informe de inspección:
        1. Estado: APROBADO (si dice que cumple con los requerimientos) o RECHAZADO (si dice que no cumple).
        2. Observaciones: Extrae y escribe la ORACIÓN COMPLETA EXACTA del texto donde se indica que el equipo cumple o no cumple con los requerimientos.
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
