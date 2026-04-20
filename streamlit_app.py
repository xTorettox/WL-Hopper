import streamlit as st
import os
import shutil
import pandas as pd
import zipfile
from io import BytesIO
from scraper import WLHopperBot
from utils import extraer_internos, asegurar_carpeta
import streamlit.components.v1 as components

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Sullair Argentina - WL Hopper", page_icon="img/favicon.png", layout="wide")

# --- ESTILOS CSS (Ajuste de márgenes y Verde Sullair) ---
VERDE_SULLAIR = "#008657"
st.markdown(f"""
    <style>
    /* Reducir margen superior */
    .block-container {{ padding-top: 0.5rem !important; padding-bottom: 0rem !important; }}
    
    /* Verde Sullair en Botones y Checkboxes */
    div.stButton > button:first-child {{ background-color: {VERDE_SULLAIR} !important; color: white !important; }}
    [data-testid="stCheckbox"] [data-testid="stWidgetLabel"] p {{ color: {VERDE_SULLAIR}; font-weight: bold; }}
    [data-testid="stCheckbox"] div[role="checkbox"][aria-checked="true"] {{ background-color: {VERDE_SULLAIR} !important; border-color: {VERDE_SULLAIR} !important; }}
    
    .terminal-box {{
        background-color: #212529; color: #f8f9fa; font-family: 'Consolas', monospace;
        font-size: 13px; padding: 15px; border-radius: 5px; height: 520px; overflow-y: auto; border: 1px solid #444;
    }}
    .log-entry {{ margin-bottom: 5px; border-bottom: 1px solid #333; padding-bottom: 2px; }}
    .logo-container {{ display: flex; justify-content: center; align-items: center; flex-direction: column; margin-bottom: 10px; }}
    </style>
    """, unsafe_allow_html=True)

# --- INICIALIZACIÓN ---
if "log_history" not in st.session_state: st.session_state.log_history = []
if "proceso_completo" not in st.session_state: st.session_state.proceso_completo = False
if "html_excel" not in st.session_state: st.session_state.html_excel = ""

# --- LAYOUT ---
col_left, col_right = st.columns([1, 1.2], gap="large")

with col_left:
    st.markdown('<div class="logo-container">', unsafe_allow_html=True)
    st.image("img/WL Hopper Logo - nspc.png", width=350)
    st.markdown("<h5 style='text-align: center; color: #555; margin-top:-10px;'>Automatización de Descarga de Certificados</h5>", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    
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
        html = f'<div class="terminal-box">'
        for entry in st.session_state.log_history: html += f'<div class="log-entry">{entry}</div>'
        html += '</div>'
        terminal_placeholder.markdown(html, unsafe_allow_html=True)
    render_terminal()

# --- LÓGICA ---
if btn_run:
    if not user or not pw: st.error("Faltan credenciales.")
    else:
        ruta_temp = "descargas_temp"
        if os.path.exists(ruta_temp): shutil.rmtree(ruta_temp)
        asegurar_carpeta(ruta_temp)

        st.session_state.proceso_completo = False
        st.session_state.log_history = ["🧹 Carpeta temporal limpia.", "Iniciando conexión con Worklift..."]
        render_terminal()
        
        bot = WLHopperBot(headless=True)
        if bot.iniciar(user, pw):
            lista = extraer_internos(texto_internos)
            resultados = []
            for int_id in lista:
                st.session_state.log_history.append(f"--- Procesando {int_id} ---")
                render_terminal()
                res = bot.procesar_interno(int_id, ruta_temp, bajar_cert, bajar_inf)
                res['id'] = int_id
                resultados.append(res)
                for m in res.get('log', []): st.session_state.log_history.append(f"&nbsp;&nbsp;{m}")
                render_terminal()

            bot.cerrar()
            st.session_state.log_history.append("🏁 PROCESO FINALIZADO.")
            st.session_state.proceso_completo = True
            
            # GENERAR HTML EXACTO (Pedido #1)
            html = """<style>table { border-collapse: collapse; } td { white-space: nowrap; text-align: center; vertical-align: middle; mso-number-format: "\\@"; padding: 5px 15px; } th { white-space: nowrap; text-align: center; padding: 10px 20px; }</style>
            <table width="100%" border="1" style="font-family: Calibri;">
            <tr style="background-color: #008657; color: white; font-weight: bold;"><th>INTERNO</th><th>ESTADO</th><th>ÚLTIMA INSPECCIÓN</th><th>VENCIMIENTO</th><th>CERTIFICADO</th><th>INFORME</th><th>DETALLE</th></tr>"""
            
            for r in resultados:
                bg, tx, st_text = "#FFFFFF", "#000000", r['status'].upper()
                if "VIGENTE" in st_text: bg, tx = "#C6EFCE", "#006100"
                elif "PRÓXIMO" in st_text: bg, tx = "#FFEB9C", "#9C5700"
                elif "VENCIDO" in st_text: bg, tx = "#FFC7CE", "#9C0006"
                
                html += f'<tr><td>{r["id"]}</td><td style="background-color: {bg}; color: {tx}; font-weight: bold;">{st_text}</td>'
                html += f'<td>{r["insp"]}</td><td>{r["venc"]}</td><td>{r["cert"]}</td><td>{r["inf"]}</td>'
                html += f'<td style="text-align: left; white-space: normal; padding-right: 30px;">{r["det"]}</td></tr>'
            html += "</table>"
            st.session_state.html_excel = html.replace("\n", "")
            st.rerun()

# --- BOTONES DE ACCIÓN ---
st.divider()
dcol1, dcol2 = st.columns(2)

with dcol1:
    if st.session_state.proceso_completo:
        # COMPONENTE DE COPIADO JAVASCRIPT (Infalible)
        components.html(f"""
            <button id="copyBtn" style="width: 100%; height: 45px; background-color: {VERDE_SULLAIR}; color: white; border: none; border-radius: 4px; font-weight: bold; cursor: pointer;">
                📋 Copiar Reporte para Excel
            </button>
            <script>
            document.getElementById('copyBtn').onclick = function() {{
                const blob = new Blob([`{st.session_state.html_excel}`], {{ type: 'text/html' }});
                const item = new ClipboardItem({{ 'text/html': blob }});
                navigator.clipboard.write([item]).then(() => {{
                    alert('¡Reporte copiado! Dale Ctrl+V en Excel.');
                }});
            }};
            </script>
        """, height=60)
    else:
        st.button("📋 Copiar Reporte para Excel", disabled=True, use_container_width=True)

with dcol2:
    zip_buffer = BytesIO()
    if st.session_state.proceso_completo:
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            if os.path.exists("descargas_temp"):
                for root, _, files in os.walk("descargas_temp"):
                    for file in files: zip_file.write(os.path.join(root, file), file)
    
    st.download_button(
        "📂 Descargar Archivo ZIP", data=zip_buffer.getvalue(), file_name="certificados.zip", 
        disabled=not (st.session_state.proceso_completo and zip_buffer.tell() > 0), use_container_width=True
    )

st.caption("© 2026 - Desarrollado por Fede García Cendra para Sullair Argentina S.A.")