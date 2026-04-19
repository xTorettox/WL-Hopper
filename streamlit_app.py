import streamlit as st
import os
import pandas as pd
import zipfile
from io import BytesIO
from scraper import WLHopperBot
from utils import extraer_internos, asegurar_carpeta

# --- CONFIGURACIÓN DE PÁGINA E ICONO ---
# Usamos el favicon.png de la carpeta img
st.set_page_config(
    page_title="Sullair Argentina - WL Hopper", 
    page_icon="img/favicon.png", 
    layout="wide"
)

# --- ESTILO CSS (Terminal, Layout y Botones) ---
st.markdown("""
    <style>
    .terminal-box {
        background-color: #212529;
        color: #f8f9fa;
        font-family: 'Consolas', monospace;
        font-size: 13px;
        padding: 15px;
        border-radius: 5px;
        height: 500px;
        overflow-y: auto;
        border: 1px solid #444;
    }
    .log-entry { margin-bottom: 5px; border-bottom: 1px solid #333; padding-bottom: 2px; }
    .main-logo { display: block; margin-left: auto; margin-right: auto; padding-bottom: 20px; }
    </style>
    """, unsafe_allow_html=True)

# --- INICIALIZACIÓN DE ESTADOS ---
if "playwright_installed" not in st.session_state:
    os.system("playwright install chromium")
    st.session_state.playwright_installed = True

if "log_history" not in st.session_state:
    st.session_state.log_history = []

if "proceso_completo" not in st.session_state:
    st.session_state.proceso_completo = False

if "resultados" not in st.session_state:
    st.session_state.resultados = []

# --- CABECERA MEJORADA ---
# Centramos y agrandamos el logo para que no se vea "solo y chico"
st.image("img/WL Hopper Logo - nspc.png", width=450)
st.markdown("<h4 style='text-align: center; color: #666;'>Automatización de Descarga de Certificados de Worklift</h4>", unsafe_allow_html=True)
st.divider()

# --- LAYOUT PRINCIPAL ---
col_left, col_right = st.columns([1, 1.2], gap="large")

with col_left:
    st.markdown("##### Acceso y Configuración")
    with st.container(border=True):
        user = st.text_input("Usuario (Email)", key="user_email")
        pw = st.text_input("Contraseña", type="password", key="user_pw")
        c1, c2 = st.columns(2)
        bajar_cert = c1.checkbox("Certificados", value=True)
        bajar_inf = c2.checkbox("Informes Técnicos", value=False)

    st.markdown("##### Listado de Internos")
    texto_internos = st.text_area("Pegá aquí:", height=150, placeholder="E040230, 3797...")
    btn_run = st.button("🚀 COMENZAR PROCESO", use_container_width=True, type="primary")

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
        st.session_state.proceso_completo = False
        st.session_state.log_history = ["Iniciando conexión con Worklift..."]
        render_terminal()
        
        bot = WLHopperBot(headless=True)
        ruta_temp = "descargas_temp"
        asegurar_carpeta(ruta_temp)
        
        if bot.iniciar(user, pw):
            st.session_state.log_history.append("🔐 Log in exitoso.")
            render_terminal()
            
            lista = extraer_internos(texto_internos)
            st.session_state.resultados = []
            
            for int_id in lista:
                st.session_state.log_history.append(f"--- Procesando interno {int_id} ---")
                render_terminal()
                # El bot ya NO descarga si está vencido por lógica interna de scraper.py
                res = bot.procesar_interno(int_id, ruta_temp, bajar_cert, bajar_inf)
                res['id'] = int_id
                st.session_state.resultados.append(res)
                for msg in res.get('log', []):
                    st.session_state.log_history.append(f"&nbsp;&nbsp;{msg}")
                render_terminal()

            bot.cerrar()
            st.session_state.log_history.append("🏁 PROCESO FINALIZADO.")
            st.session_state.proceso_completo = True
            render_terminal()
            st.rerun() # Refrescamos para habilitar botones
        else:
            st.session_state.log_history.append("❌ ERROR: Credenciales inválidas.")
            render_terminal()

# --- BOTONES DE ACCIÓN (Siempre visibles, deshabilitados según estado) ---
st.divider()
dcol1, dcol2 = st.columns(2)

# 1. COPIAR REPORTE EXCEL (HTML con formato original)
# Generamos el HTML con el semáforo de colores
html_excel = ""
if st.session_state.proceso_completo:
    html_excel = r"""<table border="1" style="font-family: Calibri; border-collapse: collapse;">
    <tr style="background-color: #008657; color: white; font-weight: bold;">
    <th>INTERNO</th><th>ESTADO</th><th>ÚLTIMA INSPECCIÓN</th><th>VENCIMIENTO</th><th>CERTIFICADO</th><th>INFORME</th><th>DETALLE</th></tr>"""
    for r in st.session_state.resultados:
        bg, tx, st_text = "#FFFFFF", "#000000", r.get('status', '').upper()
        if "VIGENTE" in st_text: bg, tx = "#C6EFCE", "#006100"
        elif "PRÓXIMO" in st_text: bg, tx = "#FFEB9C", "#9C5700"
        elif "VENCIDO" in st_text: bg, tx = "#FFC7CE", "#9C0006"
        html_excel += f"""<tr><td>{r.get('id','-')}</td><td style="background-color: {bg}; color: {tx}; font-weight: bold;">{st_text}</td>
        <td>{r.get('insp','-')}</td><td>{r.get('venc','-')}</td><td>{r.get('cert','NO')}</td><td>{r.get('inf','NO')}</td>
        <td style="text-align: left;">{r.get('det','-')}</td></tr>"""
    html_excel += "</table>"

with dcol1:
    # En la web usamos un archivo .xls que Excel abre interpretando el HTML (mantiene formato)
    st.download_button(
        "📊 Copiar Reporte Excel", 
        data=html_excel, 
        file_name="reporte_hopper.xls", 
        mime="text/html",
        disabled=not st.session_state.proceso_completo,
        use_container_width=True
    )

# 2. DESCARGAR ARCHIVO ZIP
with dcol2:
    zip_buffer = BytesIO()
    if st.session_state.proceso_completo:
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for root, _, files in os.walk("descargas_temp"):
                for file in files: zip_file.write(os.path.join(root, file), file)
    
    st.download_button(
        "📂 Descargar Archivo ZIP", 
        data=zip_buffer.getvalue(), 
        file_name="certificados.zip", 
        mime="application/zip",
        disabled=not st.session_state.proceso_completo,
        use_container_width=True
    )

st.markdown("<br><p style='text-align: left; font-size: 12px; color: #666;'>© 2026 - Desarrollado por Fede García Cendra para Sullair Argentina S.A.</p>", unsafe_allow_html=True)