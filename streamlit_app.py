import streamlit as st
import os
import shutil
import pandas as pd
import zipfile
from io import BytesIO
from scraper import WLHopperBot
from utils import extraer_internos, asegurar_carpeta
import streamlit.components.v1 as components

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Sullair Argentina - WL Hopper", page_icon="img/favicon.png", layout="wide")

# --- ESTILO CSS (Verde Sullair y Terminal) ---
st.markdown("""
    <style>
    /* Color Verde Sullair para botones y checkboxes */
    div.stButton > button:first-child {
        background-color: #008657 !important;
        color: white !important;
        border: none !important;
    }
    div[data-testid="stCheckbox"] > label > div[data-testid="stWidgetLabel"] {
        color: #008657 !important;
    }
    
    .terminal-box {
        background-color: #212529;
        color: #f8f9fa;
        font-family: 'Consolas', monospace;
        font-size: 13px;
        padding: 15px;
        border-radius: 5px;
        height: 540px; /* Ajustado para que termine a la altura del botón */
        overflow-y: auto;
        border: 1px solid #444;
    }
    .log-entry { margin-bottom: 5px; border-bottom: 1px solid #333; padding-bottom: 2px; }
    
    /* Centrado de logo */
    .logo-container {
        display: flex;
        justify-content: center;
        align-items: center;
        flex-direction: column;
        margin-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

# --- FUNCIÓN PARA COPIAR (JS) ---
def st_copy_to_clipboard(text):
    # Solución al AttributeError: Usamos JS para copiar
    components.html(f"""
        <script>
        const text = `{text}`;
        navigator.clipboard.writeText(text).then(() => {{
            window.parent.postMessage({{type: 'copy_success'}}, '*');
        }});
        </script>
    """, height=0)
    st.toast("¡Reporte copiado! Dale Ctrl+V en Excel.")

# --- INICIALIZACIÓN ---
if "log_history" not in st.session_state: st.session_state.log_history = []
if "proceso_completo" not in st.session_state: st.session_state.proceso_completo = False
if "resultados" not in st.session_state: st.session_state.resultados = []
if "html_reporte" not in st.session_state: st.session_state.html_reporte = ""

# --- LAYOUT ---
col_left, col_right = st.columns([1, 1.2], gap="large")

with col_left:
    # Logo centrado con subtítulo (Pedido #1)
    st.markdown('<div class="logo-container">', unsafe_allow_html=True)
    st.image("img/WL Hopper Logo - nspc.png", width=350)
    st.markdown("<h5 style='text-align: center; color: #555; margin-top:-10px;'>Automatización de Descarga de Certificados</h5>", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("##### Acceso y Configuración")
    with st.container(border=True):
        user = st.text_input("Usuario (Email)", key="user_email")
        pw = st.text_input("Contraseña", type="password", key="user_pw")
        c1, c2 = st.columns(2)
        bajar_cert = c1.checkbox("Certificados", value=True)
        bajar_inf = c2.checkbox("Informes Técnicos", value=False)

    st.markdown("##### Listado de Internos")
    texto_internos = st.text_area("Pegá aquí:", height=100, placeholder="E040230, 3797...")
    btn_run = st.button("🚀 COMENZAR PROCESO", use_container_width=True)

with col_right:
    st.markdown("##### Registro de Actividad")
    terminal_placeholder = st.empty()
    
    def render_terminal():
        html = '<div class="terminal-box">'
        for entry in st.session_state.log_history:
            html += f'<div class="log-entry">{entry}</div>'
        html += '</div>'
        terminal_placeholder.markdown(html, unsafe_allow_html=True)
    render_terminal()

# --- LÓGICA DE PROCESO ---
if btn_run:
    if not user or not pw:
        st.error("Faltan credenciales.")
    elif not texto_internos.strip():
        st.warning("No hay internos.")
    else:
        ruta_temp = "descargas_temp"
        if os.path.exists(ruta_temp): shutil.rmtree(ruta_temp)
        asegurar_carpeta(ruta_temp)

        st.session_state.proceso_completo = False
        st.session_state.log_history = ["🧹 Carpeta temporal limpia.", "Conectando con Worklift..."]
        render_terminal()
        
        bot = WLHopperBot(headless=True)
        if bot.iniciar(user, pw):
            st.session_state.log_history.append("🔐 Log in exitoso.")
            render_terminal()
            
            lista = extraer_internos(texto_internos)
            st.session_state.resultados = []
            
            for int_id in lista:
                st.session_state.log_history.append(f"--- Procesando {int_id} ---")
                render_terminal()
                res = bot.procesar_interno(int_id, ruta_temp, bajar_cert, bajar_inf)
                res['id'] = int_id
                st.session_state.resultados.append(res)
                for m in res.get('log', []): st.session_state.log_history.append(f"&nbsp;&nbsp;{m}")
                render_terminal()

            bot.cerrar()
            st.session_state.log_history.append("🏁 PROCESO FINALIZADO.")
            st.session_state.proceso_completo = True
            
            # Generar HTML para Excel
            html = '<table border="1" style="font-family: Calibri; border-collapse: collapse;">'
            html += '<tr style="background-color: #008657; color: white;"><th>INTERNO</th><th>ESTADO</th><th>VENCIMIENTO</th></tr>'
            for r in st.session_state.resultados:
                bg = "#C6EFCE" if "VIGENTE" in r['status'] else "#FFC7CE"
                html += f"<tr><td>{r['id']}</td><td style='background-color:{bg}'>{r['status']}</td><td>{r['venc']}</td></tr>"
            html += '</table>'
            st.session_state.html_reporte = html
            st.rerun()

# --- BOTONES DE ACCIÓN (Siempre visibles) ---
st.divider()
dcol1, dcol2 = st.columns(2)

with dcol1:
    if st.button("📋 Copiar Reporte para Excel", disabled=not st.session_state.proceso_completo, use_container_width=True):
        st_copy_to_clipboard(st.session_state.html_reporte)

with dcol2:
    # Recuperamos el botón de descarga
    zip_buffer = BytesIO()
    if st.session_state.proceso_completo:
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            if os.path.exists("descargas_temp"):
                for root, _, files in os.walk("descargas_temp"):
                    for file in files: zip_file.write(os.path.join(root, file), file)
    
    st.download_button(
        "📂 Descargar Archivo ZIP", 
        data=zip_buffer.getvalue() if zip_buffer.tell() > 0 else b"", 
        file_name="certificados.zip", 
        disabled=not (st.session_state.proceso_completo and zip_buffer.tell() > 0),
        use_container_width=True
    )
    
st.caption("© 2026 - Desarrollado por Fede García Cendra para Sullair Argentina S.A.")