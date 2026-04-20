import streamlit as st
import os
import shutil
import pandas as pd
import zipfile
from io import BytesIO
from scraper import WLHopperBot
from utils import extraer_internos, asegurar_carpeta
import streamlit.components.v1 as components
import base64

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="WL Hopper - Sullair Argentina", page_icon="img/favicon.png", layout="wide")

# --- FUNCIÓN DE LOGIN ---
def check_password():
    def password_entered():
        if st.session_state["username"] in st.secrets["passwords"] and \
           st.session_state["password"] == st.secrets["passwords"][st.session_state["username"]]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.image("img/WL Hopper Logo - nspc.png", width=300)
        with st.form("login_form"):
            st.text_input("Usuario", key="username")
            st.text_input("Contraseña", type="password", key="password")
            st.form_submit_button("Ingresar", on_click=password_entered)
        return False
    return True

# --- FLUJO PRINCIPAL ---
if check_password():
    st.sidebar.button("Cerrar Sesión", on_click=lambda: st.session_state.clear())

    # --- ESTILOS CSS ---
    VERDE_SULLAIR = "#008657"
    st.markdown(f"""
        <style>
        [data-testid="stColumn"] {{ display: flex; flex-direction: column; justify-content: flex-end; }}
        div.stButton > button:first-child {{ background-color: {VERDE_SULLAIR} !important; color: white !important; font-weight: bold; }}
        .terminal-box {{ 
            background-color: #212529; color: #f8f9fa; font-family: 'Consolas', monospace; 
            font-size: 13px; padding: 15px; border-radius: 5px; height: 535px; 
            overflow-y: auto; border: 1px solid #444; 
        }}
        .log-entry {{ margin-bottom: 5px; border-bottom: 1px solid #333; padding-bottom: 2px; }}
        </style>
    """, unsafe_allow_html=True)

    # --- ESTADO DE SESIÓN ---
    if "log_history" not in st.session_state: st.session_state.log_history = []
    if "proceso_completo" not in st.session_state: st.session_state.proceso_completo = False
    if "html_excel" not in st.session_state: st.session_state.html_excel = ""
    if "hay_archivos" not in st.session_state: st.session_state.hay_archivos = False
    if "zip_base64" not in st.session_state: st.session_state.zip_base64 = ""

    def render_terminal(placeholder):
        html = f'<div class="terminal-box">'
        for entry in st.session_state.log_history: html += f'<div class="log-entry">{entry}</div>'
        html += '</div>'
        placeholder.markdown(html, unsafe_allow_html=True)

    # --- LAYOUT ---
    col_left, col_right = st.columns([1, 1.2], gap="large")

    with col_left:
        st.image("img/WL Hopper Logo - nspc.png", width=250)
        
        with st.container(border=True):
            user = st.text_input("Usuario Worklift", key="user_email")
            pw = st.text_input("Contraseña Worklift", type="password", key="user_pw")
            c1, c2, c3 = st.columns(3)
            bajar_cert = c1.checkbox("Certificados", value=True)
            bajar_inf = c2.checkbox("Informes", value=False)
            modo_6meses = c3.toggle("Modo 6m")

        st.markdown("##### Listado de Internos")
        # VOLVIÓ LA VENTANITA DE TEXTO (Prioridad 1)
        texto_puro = st.text_area("Pegá el texto del mail acá:", height=100, placeholder="E040230, 3797...")
        
        # Opción de Excel (Prioridad 2)
        archivo_excel = st.file_uploader("O subí un archivo:", type=["xlsx", "csv"])
        
        btn_run = st.button("🚀 COMENZAR PROCESO", use_container_width=True)

    with col_right:
        st.markdown("##### Registro de Actividad")
        terminal_placeholder = st.empty()
        render_terminal(terminal_placeholder)

    # --- LÓGICA DE PROCESO ---
    if btn_run:
        if not user or not pw:
            st.error("Faltan credenciales de Worklift.")
        else:
            # Notificación permission
            components.html("<script>Notification.requestPermission();</script>", height=0)
            
            # Decidir de dónde sacar los internos
            if archivo_excel:
                # Si subió archivo, leemos el contenido para la regex
                df_temp = pd.read_csv(archivo_excel) if archivo_excel.name.endswith('csv') else pd.read_excel(archivo_excel)
                contenido_para_regex = df_temp.to_string()
            else:
                contenido_para_regex = texto_puro

            lista_internos = extraer_internos(contenido_para_regex)

            if not lista_internos:
                st.warning("No encontré internos en el texto o archivo.")
            else:
                ruta_temp = "descargas_temp"
                if os.path.exists(ruta_temp): shutil.rmtree(ruta_temp)
                asegurar_carpeta(ruta_temp)

                st.session_state.proceso_completo = False
                st.session_state.log_history = [f"Detectados {len(lista_internos)} internos. Iniciando..."]
                render_terminal(terminal_placeholder)
                
                bot = WLHopperBot(headless=True)
                if bot.iniciar(user, pw):
                    st.session_state.log_history.append("🔐 Login exitoso.")
                    res_lista = []
                    
                    for int_id in lista_internos:
                        st.session_state.log_history.append(f"--- Procesando {int_id} ---")
                        render_terminal(terminal_placeholder)
                        res = bot.procesar_interno(int_id, ruta_temp, bajar_cert, bajar_inf)
                        res['id'] = int_id
                        res_lista.append(res)
                        for m in res.get('log', []): st.session_state.log_history.append(f"&nbsp;&nbsp;{m}")
                        render_terminal(terminal_placeholder)

                    bot.cerrar()
                    st.session_state.log_history.append("🏁 PROCESO FINALIZADO.")
                    st.session_state.proceso_completo = True
                    
                    # Generar ZIP y HTML (Lógica semáforos igual que antes)
                    z_buf = BytesIO()
                    if os.path.exists(ruta_temp) and os.listdir(ruta_temp):
                        with zipfile.ZipFile(z_buf, "a", zipfile.ZIP_DEFLATED, False) as zf:
                            for r_dir, d_dir, f_list in os.walk(ruta_temp):
                                for f in f_list: zf.write(os.path.join(r_dir, f), f)
                        st.session_state.hay_archivos = True
                        st.session_state.zip_base64 = base64.b64encode(z_buf.getvalue()).decode()
                    
                    # (Aquí va tu lógica de generación de st.session_state.html_excel que ya teníamos)
                    # ... [Omitido para no hacer el post infinito, pero mantené tu tabla de colores] ...
                    
                    st.rerun()

    # --- BOTONES FINALES ---
    # Mantené el bloque de components.html con el botón de Copiar y el ZIP que ya funcionaban
    # ... 

    st.divider()
    st.caption("© 2026 - Desarrollado por Fede García Cendra para Sullair Argentina S.A.")
    st.caption("Consultas: fcendra@sullair.com.ar")
