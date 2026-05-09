from google import genai
import fitz  # PyMuPDF
import json
import streamlit as st

def configurar_gemini():
    try:
        api_key = st.secrets["GOOGLE_API_KEY"]["GOOGLE_API_KEY"]
    except KeyError:
        # Intento de fallback por si está en la raíz
        try:
            api_key = st.secrets["GOOGLE_API_KEY"]
            if isinstance(api_key, dict):
                api_key = api_key.get("GOOGLE_API_KEY")
        except:
            api_key = None
            
    if api_key:
        return genai.Client(api_key=api_key)
    return None

import time

def analizar_informe_gemini(ruta_pdf):
    client = configurar_gemini()
    if not client:
        return "ERROR: Sin API Key", "No se encontró la clave API de Gemini en st.secrets."

    try:
        # Extraer texto del PDF
        doc = fitz.open(ruta_pdf)
        texto = ""
        for i in range(len(doc)):
            # Solo analizamos las primeras 3 páginas para ahorrar tokens y tiempo
            if i > 2: break 
            texto += doc[i].get_text("text") + "\n"
        doc.close()
        
        if not texto.strip():
            return "ERROR LECTURA", "No se pudo extraer texto del PDF."

        prompt = f"""
        Actúa como un analista de informes de inspección técnica.
        A continuación te proporciono el texto extraído de un informe de inspección de un equipo.
        Tu tarea es determinar si el equipo superó la inspección o fue rechazado, y cuáles fueron las observaciones.
        
        Responde ÚNICAMENTE con un objeto JSON con el siguiente formato, sin markdown extra:
        {{
            "estado": "RECHAZADO" | "APROBADO" | "DESCONOCIDO",
            "observaciones": "Resumen breve de los motivos de rechazo u observaciones importantes. Si no hay, pon '-'"
        }}
        
        Texto del informe:
        {texto}
        """
        
        # Reintentos por límite de cuota (429)
        max_intentos = 3
        for intento in range(max_intentos):
            try:
                response = client.models.generate_content(
                    model='gemini-2.0-flash',
                    contents=prompt
                )
                break  # Éxito
            except Exception as e:
                if "429" in str(e) and intento < max_intentos - 1:
                    time.sleep(5) # Esperamos 5 segundos y reintentamos
                    continue
                else:
                    raise e
        
        # Limpiar posible markdown
        resp_text = response.text.replace('```json', '').replace('```', '').strip()
        data = json.loads(resp_text)
        
        # Esperamos un momento para no agotar la cuota en ráfagas de varias lecturas
        time.sleep(2)
        
        return data.get('estado', 'DESCONOCIDO'), data.get('observaciones', '-')
        
    except Exception as e:
        return "ERROR GEMINI", f"Error al analizar con IA: {str(e)[:50]}"

def extraer_internos_imagen(imagen_bytes):
    client = configurar_gemini()
    if not client:
        return []

    try:
        prompt = "Identifica todos los códigos o números internos de equipos que veas en esta imagen. Devuelve una lista separada por comas, sin ningún otro texto."
        
        response = client.models.generate_content(
            model='gemini-2.0-flash',
            contents=[
                genai.types.Part.from_bytes(data=imagen_bytes, mime_type='image/jpeg'),
                prompt
            ]
        )
        
        texto_crudo = response.text
        return [item.strip() for item in texto_crudo.split(',') if item.strip()]
    except Exception as e:
        print(f"Error en vision: {e}")
        return []
