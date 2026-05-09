from google import genai
import fitz  # PyMuPDF
import json
import streamlit as st
import time

def configurar_gemini():
    """
    Configura el cliente de Gemini usando la nueva librería google-genai.
    Busca la clave en st.secrets de forma segura.
    """
    try:
        # Intentamos obtener la clave del formato estándar de Streamlit
        api_key = st.secrets.get("GOOGLE_API_KEY")
        
        # Si por alguna razón está anidada, la buscamos
        if isinstance(api_key, dict):
            api_key = api_key.get("GOOGLE_API_KEY")
            
        if api_key:
            return genai.Client(api_key=api_key)
            
    except Exception as e:
        print(f"Error al configurar Gemini: {e}")
    
    return None

def analizar_informe_gemini(ruta_pdf):
    client = configurar_gemini()
    if not client:
        return "ERROR: Sin API Key", "No se encontró la clave API en los secretos."

    try:
        # Extraer texto del PDF usando PyMuPDF
        doc = fitz.open(ruta_pdf)
        texto = ""
        for i in range(len(doc)):
            # Analizamos las primeras 3 páginas para optimizar tokens
            if i > 2: break 
            texto += doc[i].get_text("text") + "\n"
        doc.close()
        
        if not texto.strip():
            return "ERROR LECTURA", "El PDF parece estar vacío o ser una imagen pura."

        prompt = """
        Actúa como un analista técnico de Sullair Argentina. 
        Analiza el siguiente texto de un informe de inspección y determina:
        1. El estado del equipo (VIGENTE, RECHAZADO o REQUIERE REPARACION).
        2. Un resumen breve de las observaciones técnicas encontradas.
        
        Responde exclusivamente en formato JSON con esta estructura:
        {"estado": "VIGENTE", "observaciones": "Texto breve aquí"}
        """

        # Implementación con reintentos por cuota (Error 429)
        max_intentos = 3
        for intento in range(max_intentos):
            try:
                # Nueva sintaxis de google-genai
                response = client.models.generate_content(
                    model='gemini-1.5-flash',
                    contents=[prompt, texto]
                )
                break  # Éxito
            except Exception as e:
                if "429" in str(e) and intento < max_intentos - 1:
                    time.sleep(5)  # Espera de cortesía por cuota
                    continue
                else:
                    raise e
        
        # Limpieza de la respuesta para asegurar JSON válido
        resp_text = response.text.replace('```json', '').replace('```', '').strip()
        data = json.loads(resp_text)
        
        # Pequeña pausa para no saturar la API en procesos por lote
        time.sleep(1)
        
        return data.get('estado', 'DESCONOCIDO'), data.get('observaciones', '-')
        
    except Exception as e:
        return "ERROR IA", f"Error en proceso: {str(e)[:50]}"

def extraer_internos_imagen(imagen_bytes):
    """
    Usa la capacidad multimodal para leer internos desde una foto.
    """
    client = configurar_gemini()
    if not client:
        return []

    try:
        prompt = "Identifica los códigos internos de equipos (ej: E03, 5662). Devuelve solo una lista separada por comas."
        
        # Nueva forma de enviar bytes con google-genai
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
