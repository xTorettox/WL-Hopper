import streamlit as st
import pandas as pd
from datetime import datetime, date
import json
import base64
from utils import (
    procesar_pdf, 
    obtener_proximo_vencimiento, 
    obtener_estado_vencimiento,
    generar_pdf_reporte
)
from gemini_utils import configurar_gemini
import google.generativeai as genai

# Configuración de página
st.set_page_config(
    page_title="WL Hopper - Sullair Argentina",
    page_icon="🏗️",
    layout="wide"
)

# Estilos CSS personalizados
st.markdown("""
    <style>
    .main {
        background-color: #f5f5f5;
    }
    .stButton>button {
        width: stretch;
        border-radius: 5px;
        height: 3em;
        background-color: #004a99;
        color: white;
    }
    .report-table {
        width: stretch;
        border-collapse: collapse;
    }
    .report-table th, .report-table td {
        border: 1px solid #ddd;
        padding: 8px;
        text-align: left;
    }
    .report-table th {
        background-color: #004a99;
        color: white;
    }
    </style>
    """, unsafe_allow_html=True)

def main():
    st.title("🏗️ WL Hopper")
    st.subheader("Automatización de Informes de Inspección - Sullair Argentina")

    # Inicializar estado
    if 'resultados' not in st.session_state:
        st.session_state.resultados = []
    
    # Sidebar para configuración
    with st.sidebar:
        st.image("https://sullairargentina.com/wp-content/uploads/2023/04/logo-sullair-blue.png", width=200)
        st.divider()
        st.write("### Configuración")
        
        # Selección de tipo de equipo
        tipo_equipo = st.selectbox(
            "Tipo de Equipo",
            ["Hidroelevadores", "Manipuladores", "Grúas", "Plataformas"]
        )
        
        # Fecha de hoy para cálculos
        fecha_referencia = st.date_input("Fecha de Referencia", date.today())
        
        st.divider()
        if st.button("Limpiar Datos"):
            st.session_state.resultados = []
            st.rerun()

    # Área principal: Carga de archivos
    upload_files = st.file_uploader(
        "Cargar Informes (PDF)", 
        type="pdf", 
        accept_multiple_files=True
    )

    if upload_files:
        if st.button("🚀 Procesar Informes"):
            progreso = st.progress(0)
            status_text = st.empty()
            
            # Intentar configurar el cliente de IA
            client_ai = configurar_gemini()
            
            if not client_ai:
                st.error("No se pudo inicializar la IA. Revisá la API Key en los secretos de Streamlit.")
                return

            for i, file in enumerate(upload_files):
                status_text.text(f"Procesando: {file.name}...")
                
                # 1. Extraer texto del PDF
                texto_pdf = procesar_pdf(file)
                
                # 2. Enviar a Gemini para estructurar la data
                prompt = f"""
                Analizá el siguiente texto de un informe de inspección y extraé los datos en formato JSON.
                Necesito: 'id' (interno del equipo), 'insp' (fecha de inspección), 'venc' (fecha de vencimiento).
                Texto: {texto_pdf}
                """
                
                try:
                    # Usamos el cliente configurado
                    response = client_ai.models.generate_content(
                        model="gemini-2.0-flash",
                        contents=prompt
                    )
                    
                    # Limpiamos y parseamos el JSON (manejo de errores robusto)
                    raw_json = response.text.replace('```json', '').replace('```', '').strip()
                    data = json.loads(raw_json)
                    
                    # 3. Calcular estados y próximos vencimientos
                    info_procesada = {
                        "archivo": file.name,
                        "id": data.get("id", "S/N"),
                        "insp": data.get("insp", "N/A"),
                        "venc": data.get("venc", "N/A"),
                        "estado": obtener_estado_vencimiento(data.get("venc"), fecha_referencia)
                    }
                    
                    st.session_state.resultados.append(info_procesada)
                    
                except Exception as e:
                    st.error(f"Error procesando {file.name}: {e}")
                
                progreso.progress((i + 1) / len(upload_files))
            
            status_text.text("✅ Procesamiento completado")

    # Mostrar Resultados
    if st.session_state.resultados:
        st.divider()
        st.write("### 📊 Resultados de la Inspección")
        
        df = pd.DataFrame(st.session_state.resultados)
        st.dataframe(df, width="stretch") # Corregido para nuevas versiones

        # Generar Tabla HTML para reporte visual
        html = '<table class="report-table"><thead><tr><th>Interno</th><th>Últ. Inspección</th><th>Vencimiento</th><th>Estado</th></tr></thead><tbody>'
        
        for r in st.session_state.resultados:
            estado = r.get("estado", "DESCONOCIDO")
            bg_color = "#ff4b4b" if "VENCIDO" in estado.upper() else "#28a745"
            tx_color = "white"
            
            # Usamos .get() para evitar el KeyError 'insp'
            interno = r.get("id", "S/N")
            f_insp = r.get("insp", "N/A")
            f_venc = r.get("venc", "N/A")

            html += f'''
                <tr>
                    <td>{interno}</td>
                    <td>{f_insp}</td>
                    <td>{f_venc}</td>
                    <td style="background-color: {bg_color}; color: {tx_color}; font-weight: bold;">{estado}</td>
                </tr>
            '''
        
        html += '</tbody></table>'
        st.markdown(html, unsafe_allow_html=True)

        # Botón para descargar reporte final
        if st.button("📥 Descargar Reporte PDF"):
            pdf_bytes = generar_pdf_reporte(st.session_state.resultados)
            b64 = base64.b64encode(pdf_bytes).decode()
            href = f'<a href="data:application/pdf;base64,{b64}" download="reporte_inspeccion_{date.today()}.pdf">Hacé click acá para descargar</a>'
            st.markdown(href, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
