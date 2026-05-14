import streamlit as st
import os
import shutil
import pandas as pd
import zipfile
import datetime
from io import BytesIO
from scraper import WLHopperBot
from utils import extraer_internos, extraer_texto_de_archivo, calcular_vencimiento_semestral, asegurar_carpeta
import streamlit.components.v1 as components
import pytesseract
from PIL import Image, ImageEnhance
import re
import importlib
import utils
from cryptography.fernet import Fernet
from supabase import create_client

# --- 1. SEGURIDAD Y DB ---
try:
    FERNET_KEY = st.secrets["FERNET_KEY"].encode()
    cipher = Fernet(FERNET_KEY)
    db = create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])
    db_ready = True
except Exception as e:
    db_ready = False

def encriptar(texto): return cipher.encrypt(texto.encode()).decode() if texto else ""
def desencriptar(token):
    try: return cipher.decrypt(token.encode()).decode() if token else ""
    except: return ""

def registrar_metrica(equipo, fuente, modo_pruebas):
    if modo_pruebas or not db_ready: return
    try:
        db.table("metricas").insert({
            "usuario": st.session_state.get("logged_user", "anon"),
            "equipo": equipo, "fuente": fuente,
            "fecha": datetime.datetime.now().isoformat(), "minutos_ahorrad": 3
        }).execute()
    except: pass

def obtener_credenciales_guardadas(sitio):
    if not db_ready: return None
    try:
        res = db.table("credenciales_sitios").select("*").eq("usuario_app", st.session_state.get("logged_user")).eq("sitio", sitio).execute()
        return res.data[0] if res.data else None
    except: return None

# --- 2. CONFIGURACIÓN UI ---
st.set_page_config(page_title="WL Hopper 2.0 - Sullair Argentina", page_icon="img/favicon.png", layout="wide")
VERDE_SULLAIR = "#008657"

st.markdown(f"""
    <style>
    .terminal-box {{
        background-color: #212529; color: #f8f9fa; font-family: 'Consolas', monospace;
        font-size: 13px; padding: 15px; border-radius: 5px; 
        overflow-y: auto; border: 1px solid #444;
    }}
    @media (min-width: 768px) {{
        .terminal-box {{ position: absolute; top: 0; bottom: 0; left: 0; right: 0; width: 100%; height: 100%; min-height: 535px; }}
        [data-testid="stHorizontalBlock"] {{ align-items: stretch; }}
        [data-testid="stColumn"] {{ position: relative; }}
    }}
    div.stButton > button:first-child {{ background-color: {VERDE_SULLAIR} !important; color: white !important; font-weight: bold; }}
    .logo-container {{ display: flex; justify-content: center; align-items: center; flex-direction: column; margin-bottom: 20px; }}
    </style>
    """, unsafe_allow_html=True)

# --- 3. LOGIN ---
def check_password():
    if st.session_state.get("password_correct") is not True:
        c_l1, c_l2, c_l3 = st.columns([1.2, 1, 1.2])
        with c_l2:
            try: st.image("img/WL Hopper Logo - nspc.png", width=300)
            except: st.markdown("<h2 style='text-align: center;'>🚀 WL Hopper</h2>", unsafe_allow_html=True)
            with st.form("login_form"):
                u = st.text_input("Usuario")
                p = st.text_input("Contraseña", type="password")
                if st.form_submit_button("Ingresar", width='stretch'):
                    if u in st.secrets["passwords"] and p == st.secrets["passwords"][u]:
                        st.session_state["password_correct"] = True
                        st.session_state["logged_user"] = u
                        st.rerun()
                    else: st.error("😕 Usuario o contraseña incorrectos")
        return False
    return True

if check_password():
    # --- MODAL ACERCA DE ---
    @st.dialog("Acerca de WL Hopper")
    def mostrar_about():
        c1, c2 = st.columns([1, 2])
        with c1: st.write("🤖")
        with c2: st.markdown("**WL Hopper** es una plataforma diseñada para optimizar la gestión de certificados en Sullair.")
        st.info("🚀 Automatización masiva via Worklift & Bureau Veritas.")
        st.caption("Desarrollado por Fede García Cendra - 2026")

    # --- SIDEBAR ---
    if st.session_state.get("logged_user") == "fcendra":
        with st.sidebar:
            st.divider()
            st.markdown("### 📈 Panel Admin")
            modo_pruebas = st.checkbox("🛠️ Modo Pruebas", value=True)
            if db_ready:
                try:
                    met_res = db.table("metricas").select("minutos_ahorrad").execute()
                    total_c = len(met_res.data)
                    st.metric("Total Procesados", total_c)
                    st.metric("Horas Ahorradas", f"{(total_c*3)/60:.1f}")
                except: pass
    else: modo_pruebas = False

    with st.sidebar:
        st.markdown("### 🛠️ Opciones")
        if st.button("Cerrar Sesión", width='stretch'):
            st.session_state.clear()
            st.rerun()
        if st.button("Acerca del Proyecto", width='stretch'):
            mostrar_about()

    # --- DISEÑO PRINCIPAL ---
    st.markdown('<div class="logo-container">', unsafe_allow_html=True)
    c_l1, c_l2, c_l3 = st.columns([1.5, 1, 1.5])
    with c_l2: 
        try: st.image("img/WL Hopper Logo - nspc.png", width=300)
        except: pass
    st.markdown("<h5 style='text-align: center; color: #555; margin-top:-10px;'>Automatización de Descarga de Certificados</h5>", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # El desplegable AHORA debajo del logo
    with st.expander("🔐 Mis Cuentas Guardadas"):
        st.caption("Tus credenciales se almacenan bajo cifrado Fernet.")
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            st.markdown("**Worklift**")
            cred_wl = obtener_credenciales_guardadas("WL")
            s_u_wl = desencriptar(cred_wl["user_enc"]) if cred_wl else ""
            u_wl_in = st.text_input("Usuario WL", value=s_u_wl, key="u_wl")
            p_wl_in = st.text_input("Password WL", type="password", key="p_wl")
            if st.button("💾 Guardar WL", width='stretch'):
                if db_ready:
                    db.table("credenciales_sitios").upsert({
                        "usuario_app": st.session_state["logged_user"], 
                        "sitio": "WL", 
                        "user_enc": encriptar(u_wl_in), 
                        "pass_enc": encriptar(p_wl_in)
                    }).execute()
                    st.toast("Guardado en Supabase ✅")

        with col_c2:
            st.markdown("**Bureau Veritas**")
            st.info("Próximamente disponible.")

    # --- INPUTS Y TERMINAL ---
    if "log_history" not in st.session_state: st.session_state.log_history = []
    
    col_left, col_right = st.columns([1, 1.2], gap="large")
    with col_left:
        with st.container(border=True):
            f_u = u_wl_in if u_wl_in else s_u_wl
            f_p = p_wl_in if p_wl_in else (desencriptar(cred_wl["pass_enc"]) if cred_wl else "")
            c_c1, c_c2 = st.columns(2)
            bajar_cert = c_c1.checkbox("Descargar Certificados", value=True)
            bajar_inf = c_c2.checkbox("Descargar Informes", value=False)
            es_semestral = st.checkbox("Vencimiento Semestral")
        
        archivo = st.file_uploader("Subí tu lista", type=['txt', 'xlsx', 'png', 'jpg', 'jpeg'])
        texto = st.text_area("O pegá el texto acá:", height=115)
        btn_run = st.button("🚀 COMENZAR PROCESO", width='stretch')

    with col_right:
        terminal = st.empty()
        def render_terminal():
            html = f'<div class="terminal-box"><div style="font-weight: bold; color: #ddd; margin-bottom: 10px; border-bottom: 1px solid #555;">Registro de Actividad</div>'
            for entry in st.session_state.log_history:
                color = "#f8f9fa"
                if "✅" in entry: color = "#50fa7b"
                elif "❌" in entry: color = "#ff5555"
                html += f'<div class="log-entry" style="color: {color};">{entry}</div>'
            html += '</div>'
            terminal.markdown(html, unsafe_allow_html=True)
        render_terminal()

    # --- LÓGICA DE PROCESO ---
    if btn_run:
        if not f_u or not f_p: st.error("Faltan credenciales.")
        else:
            st.session_state.log_history = ["Iniciando motor Hopper..."]
            render_terminal()
            # [Aquí pegás el resto de tu lógica de procesamiento original]

    # --- DESCARGAS ---
    if st.session_state.get("proceso_completo"):
        st.divider()
        d1, d2, d3 = st.columns(3)
        with d2:
            nom_xls = st.text_input("Nombre Excel", value=f"Reporte_{datetime.date.today()}.xlsx")
            st.download_button("📊 Descargar Excel", data=b"data", file_name=nom_xls, width='stretch')
        with d3:
            nom_zip = st.text_input("Nombre ZIP", value=f"Certs_{datetime.date.today()}.zip")
            st.download_button("📂 Descargar ZIP", data=b"data", file_name=nom_zip, width='stretch')

    st.caption(f"© {datetime.datetime.now().year} - Sullair Argentina S.A.")
