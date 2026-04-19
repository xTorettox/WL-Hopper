import streamlit as st
import os
import pandas as pd
import zipfile
from io import BytesIO
from scraper import WLHopperBot
from utils import extraer_internos, asegurar_carpeta

# --- ESTILO DE TERMINAL (CSS) ---
st.markdown("""
    <style>
    .terminal-box {
        background-color: #212529;
        color: #f8f9fa;
        font-family: 'Consolas', monospace;
        font-size: 13px;
        padding: 15px;
        border-radius: 5px;
        height: 450px;
        overflow-y: auto;
        border: 1px solid #444;
    }
    .log-entry { margin-bottom: 5px; border-bottom: 1px solid #333; padding-bottom: 2px; }
    .log-time { color: #888; }
    </style>
    """, unsafe_allow_html=True)

# --- SOLUCIÓN PLAYWRIGHT ---
if "playwright_installed" not in st.session_state:
    with st.spinner("Configurando motores..."):
        os.system("playwright install chromium")
    st.session_state.playwright_installed = True

st.set_page_config(page_title="Sullair Argentina - WL Hopper", page_icon="🚀", layout="wide")

# --- CABECERA ---
col_logo, col_vacia = st.columns([1, 3])
with col_logo:
    try: st.image("img/WL Hopper Logo - nspc.png", width=220)
    except: st.title("🚀 WL Hopper")

# --- LAYOUT PRINCIPAL (IZQUIERDA: INPUT / DERECHA: TERMINAL) ---
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
    # Inicializamos el log en el estado de la sesión
    if "log_history" not in st.session_state:
        st.session_state.log_history = []

    # Función para renderizar la terminal
    def render_terminal():
        html_content = '<div class="terminal-box">'
        for entry in st.session_state.log_history:
            html_content += f'<div class="log-entry">{entry}</div>'
        html_content += '</div>'
        terminal_placeholder.markdown(html_content, unsafe_allow_html=True)

    render_terminal()

# --- PROCESO ---
if btn_run:
    if not user or not pw:
        st.error("Faltan credenciales.")
    elif not texto_internos.strip():
        st.warning("No hay internos.")
    else:
        lista = extraer_internos(texto_internos)
        ruta_temp = "descargas_temp"
        asegurar_carpeta(ruta_temp)
        
        st.session_state.log_history = [f"Iniciando conexión con Worklift..."]
        render_terminal()
        
        bot = WLHopperBot(headless=True)
        resultados = []

        if bot.iniciar(user, pw):
            st.session_state.log_history.append("🔐 Log in exitoso.")
            render_terminal()

            for i, int_id in enumerate(lista):
                st.session_state.log_history.append(f"--- Procesando interno {int_id} ---")
                render_terminal()
                
                res = bot.procesar_interno(int_id, ruta_temp, bajar_cert, bajar_inf)
                res['id'] = int_id
                resultados.append(res)
                
                for msg in res.get('log', []):
                    st.session_state.log_history.append(f"&nbsp;&nbsp;{msg}")
                render_terminal()

            bot.cerrar()
            st.session_state.log_history.append("🏁 PROCESO FINALIZADO.")
            render_terminal()

            # --- SECCIÓN DE DESCARGAS (Aparece abajo al terminar) ---
            st.success("¡Listo el pollo! Bajate los archivos acá abajo:")
            dcol1, dcol2 = st.columns(2)
            
            with dcol1:
                df = pd.DataFrame(resultados)
                if 'log' in df.columns: df = df.drop(columns=['log'])
                output_excel = BytesIO()
                with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Reporte')
                st.download_button("📊 Copiar Reporte Excel", output_excel.getvalue(), "reporte_hopper.xlsx", use_container_width=True)

            with dcol2:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for root, _, files in os.walk(ruta_temp):
                        for file in files: zip_file.write(os.path.join(root, file), file)
                st.download_button("📂 Descargar Archivos (ZIP)", zip_buffer.getvalue(), "certificados.zip", use_container_width=True)
        else:
            st.session_state.log_history.append("❌ ERROR: Credenciales inválidas.")
            render_terminal()

st.caption("© 2026 - Desarrollado por Fede García Cendra para Sullair Argentina S.A.")