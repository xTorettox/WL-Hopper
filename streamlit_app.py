import streamlit as st
import os
import shutil
import pandas as pd
import zipfile
from io import BytesIO
from scraper import WLHopperBot
from utils import extraer_internos, extraer_texto_de_archivo, calcular_vencimiento_semestral, asegurar_carpeta
import streamlit.components.v1 as components
import pytesseract
from PIL import Image, ImageEnhance
import re
import importlib
import utils
importlib.reload(utils)
from utils import extraer_internos, extraer_texto_de_archivo, calcular_vencimiento_semestral, asegurar_carpeta
# --- NUEVAS IMPORTACIONES ---
from cryptography.fernet import Fernet
from supabase import create_client
import datetime

# --- 1. CONEXIÓN SILENCIOSA A DB ---
try:
    cipher = Fernet(st.secrets["FERNET_KEY"].encode())
    db = create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])
    db_ready = True
except:
    db_ready = False

def encriptar(t): return cipher.encrypt(t.encode()).decode() if t else ""
def desencriptar(t):
    try: return cipher.decrypt(t.encode()).decode() if t else ""
    except: return ""

# --- CONFIGURACIÓN (ORIGINAL) ---
st.set_page_config(page_title="WL Hopper - Sullair Argentina", page_icon="img/favicon.png", layout="wide")

# --- ESTILOS CSS (ORIGINAL) ---
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
    @media (max-width: 767px) {{ .terminal-box {{ height: 400px; margin-top: 10px; }} }}
    div.stButton > button:first-child {{ background-color: {VERDE_SULLAIR} !important; color: white !important; font-weight: bold; }}
    .log-entry {{ margin-bottom: 5px; border-bottom: 1px solid #333; padding-bottom: 2px; }}
    .logo-container {{ display: flex; justify-content: center; align-items: center; flex-direction: column; margin-bottom: 10px; }}
    </style>
    """, unsafe_allow_html=True)

# --- CONTROL DE ACCESO (MODIFICADO PARA PERSISTENCIA) ---
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if not st.session_state["password_correct"]:
        c1, c2, c3 = st.columns([1.2, 1, 1.2])
        with c2:
            st.image("img/WL Hopper Logo - nspc.png", use_container_width=True)
            with st.form("login"):
                u = st.text_input("Usuario")
                p = st.text_input("Contraseña", type="password")
                if st.form_submit_button("Ingresar", use_container_width=True):
                    if u in st.secrets["passwords"] and p == st.secrets["passwords"][u]:
                        st.session_state["password_correct"] = True
                        st.session_state["logged_user"] = u
                        # Carga inicial de credenciales guardadas
                        if db_ready:
                            try:
                                res = db.table("credenciales_sitios").select("*").eq("usuario_app", u).eq("sitio", "WL").execute()
                                if res.data:
                                    st.session_state["u_wl_saved"] = desencriptar(res.data[0]["user_enc"])
                                    st.session_state["p_wl_saved"] = desencriptar(res.data[0]["pass_enc"])
                            except: pass
                        st.rerun()
                    else: st.error("Usuario o contraseña incorrectos")
        return False
    return True

if check_password():
    # --- SIDEBAR (ORIGINAL + MODO PRUEBAS) ---
    with st.sidebar:
        st.image("img/WL Hopper Logo - nspc.png", width=150)
        st.title("Opciones")
        if st.session_state["logged_user"] == "fcendra":
            st.divider()
            modo_pruebas = st.checkbox("🛠️ Modo Pruebas", value=True)
            if db_ready:
                try:
                    total = len(db.table("metricas").select("id").execute().data)
                    st.metric("Total Procesados", total)
                except: pass
        else: modo_pruebas = False
        
        if st.button("Cerrar Sesión", use_container_width=True):
            st.session_state.clear()
            st.rerun()

    # --- DISEÑO PRINCIPAL (ORIGINAL) ---
    st.markdown('<div class="logo-container">', unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1.5, 1, 1.5])
    with c2: st.image("img/WL Hopper Logo - nspc.png", use_container_width=True)
    st.markdown("<h5 style='text-align: center; color: #555; margin-top:-10px; margin-bottom: 25px;'>Automatización de Descarga de Certificados</h5>", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # --- CREDENCIALES (SÓLO SI ES NECESARIO) ---
    with st.expander("🔐 Mis Cuentas Guardadas"):
        col_wl, col_bv = st.columns(2)
        with col_wl:
            st.caption("Worklift")
            u_in = st.text_input("Usuario WL", value=st.session_state.get("u_wl_saved", ""))
            p_in = st.text_input("Password WL", type="password", value=st.session_state.get("p_wl_saved", ""))
            if st.button("💾 Guardar WL", use_container_width=True):
                if db_ready:
                    db.table("credenciales_sitios").upsert({
                        "usuario_app": st.session_state["logged_user"], "sitio": "WL",
                        "user_enc": encriptar(u_in), "pass_enc": encriptar(p_in)
                    }).execute()
                    st.session_state["u_wl_saved"], st.session_state["p_wl_saved"] = u_in, p_in
                    st.toast("Guardado!")

    # --- LÓGICA DE PROCESO (TU ORIGINAL REFORZADA) ---
    if "log_history" not in st.session_state: st.session_state.log_history = []
    
    col_left, col_right = st.columns([1, 1.2], gap="large")
    with col_left:
        with st.container(border=True):
            c_cert, c_inf = st.columns(2)
            bajar_cert = c_cert.checkbox("Descargar Certificados", value=True)
            bajar_inf = c_inf.checkbox("Descargar Informes", value=False)
            es_semestral = st.checkbox("Vencimiento Semestral (180 días)")
        
        archivo = st.file_uploader("Subí tu lista", type=['txt', 'xlsx', 'png', 'jpg'])
        texto = st.text_area("O pegá los internos acá:", height=115)
        btn_run = st.button("🚀 COMENZAR PROCESO", use_container_width=True)

    with col_right:
        terminal = st.empty()
        def render_terminal():
            html = f'<div class="terminal-box">'
            for entry in st.session_state.log_history:
                html += f'<div class="log-entry">{entry}</div>'
            html += '</div>'
            terminal.markdown(html, unsafe_allow_html=True)
        render_terminal()

    if btn_run:
        # Usamos las del input o las guardadas
        user_wl = u_in if u_in else st.session_state.get("u_wl_saved")
        pass_wl = p_in if p_in else st.session_state.get("p_wl_saved")
        
        if not user_wl or not pass_wl:
            st.error("Faltan credenciales de Worklift")
        else:
            # AQUÍ COMIENZA TU BOT (Aseguramos que el log se limpie)
            st.session_state.log_history = ["🛰️ Iniciando motor Hopper..."]
            render_terminal()
            
            # --- LÓGICA DE SCRAPER (Aquí integrás tu loop original) ---
            # Por cada equipo procesado exitosamente, agregamos esto:
            # if not modo_pruebas and db_ready:
            #     db.table("metricas").insert({"usuario": st.session_state["logged_user"], "equipo": id_equipo, "fuente": "WL"}).execute()
