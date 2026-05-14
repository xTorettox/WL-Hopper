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
        overflow-y: auto; border: 1px solid #444; min-height: 535px;
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
                        # CARGA INICIAL DE CREDENCIALES AL LOGUEAR
                        cred = obtener_credenciales_guardadas("WL")
                        if cred:
                            st.session_state["u_wl_saved"] = desencriptar(cred["user_enc"])
                            st.session_state["p_wl_saved"] = desencriptar(cred["pass_enc"])
                        st.rerun()
                    else: st.error("😕 Usuario o contraseña incorrectos")
        return False
    return True

if check_password():
    @st.dialog("Acerca de WL Hopper")
    def mostrar_about():
        st.markdown("**WL Hopper 2.0**")
        st.info("Plataforma de automatización masiva para Sullair Argentina.")
        st.caption("Desarrollado por Fede García Cendra - 2026")

    # --- SIDEBAR ---
    with st.sidebar:
        st.markdown(f"Bienvenido, **{st.session_state['logged_user']}**")
        if st.session_state.get("logged_user") == "fcendra":
            st.divider()
            modo_pruebas = st.checkbox("🛠️ Modo Pruebas", value=True)
            if db_ready:
                try:
                    met_res = db.table("metricas").select("minutos_ahorrad").execute()
                    total_c = len(met_res.data)
                    st.metric("Total Procesados", total_c)
                except: pass
        else: modo_pruebas = False
        
        if st.button("Cerrar Sesión", width='stretch'):
            st.session_state.clear()
            st.rerun()
        if st.button("Acerca del Proyecto", width='stretch'):
            mostrar_about()

    # --- DISEÑO PRINCIPAL ---
    st.markdown('<div class="logo-container">', unsafe_allow_html=True)
    c_l1, c_l2, c_l3 = st.columns([1.5, 1, 1.5])
    with c_l2: st.image("img/WL Hopper Logo - nspc.png", width=300)
    st.markdown("<h5 style='text-align: center; color: #555; margin-top:-10px;'>Automatización de Descarga de Certificados</h5>", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # --- CREDENCIALES CON PERSISTENCIA ---
    with st.expander("🔐 Mis Cuentas Guardadas", expanded=False):
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            st.markdown("**Worklift**")
            # Usamos session_state para que el valor persista entre clics
            u_wl_val = st.text_input("Usuario WL", value=st.session_state.get("u_wl_saved", ""))
            p_wl_val = st.text_input("Password WL", type="password", value=st.session_state.get("p_wl_saved", ""))
            
            if st.button("💾 Guardar y Aplicar", width='stretch'):
                if db_ready:
                    db.table("credenciales_sitios").upsert({
                        "usuario_app": st.session_state["logged_user"], 
                        "sitio": "WL", 
                        "user_enc": encriptar(u_wl_val), 
                        "pass_enc": encriptar(p_wl_val)
                    }).execute()
                    st.session_state["u_wl_saved"] = u_wl_val
                    st.session_state["p_wl_saved"] = p_wl_val
                    st.success("Credenciales actualizadas en DB ✅")
                    st.rerun()

        with col_c2:
            st.markdown("**Bureau Veritas**")
            st.info("Próximamente disponible.")

    # --- LÓGICA DE PROCESO ---
    if "log_history" not in st.session_state: st.session_state.log_history = []
    
    col_left, col_right = st.columns([1, 1.2], gap="large")
    with col_left:
        with st.container(border=True):
            c_c1, c_c2 = st.columns(2)
            bajar_cert = c_c1.checkbox("Descargar Certificados", value=True)
            bajar_inf = c_c2.checkbox("Descargar Informes", value=False)
            es_semestral = st.checkbox("Vencimiento Semestral")
        
        archivo = st.file_uploader("Subí tu lista", type=['txt', 'xlsx', 'png', 'jpg', 'jpeg'])
        texto = st.text_area("O pegá el texto acá:", height=115)
        btn_run = st.button("🚀 COMENZAR PROCESO", width='stretch')

    with col_right:
        terminal_placeholder = st.empty()
        def render_terminal():
            html = f'<div class="terminal-box"><b>Registro de Actividad</b><br>'
            for entry in st.session_state.log_history:
                html += f'<div style="color:#50fa7b">{entry}</div>'
            html += '</div>'
            terminal_placeholder.markdown(html, unsafe_allow_html=True)
        render_terminal()

    if btn_run:
        # Recuperamos las credenciales finales (las del input o las guardadas en session)
        final_u = u_wl_val if u_wl_val else st.session_state.get("u_wl_saved")
        final_p = p_wl_val if p_wl_val else st.session_state.get("p_wl_saved")
        
        if not final_u or not final_p:
            st.error("❌ No hay credenciales para iniciar el motor.")
        else:
            # LIMPIAR LOG ANTES DE EMPEZAR
            st.session_state.log_history = ["Iniciando motor Hopper..."]
            render_terminal()
            
            # --- AQUÍ VA TU LLAMADA AL SCRAPER ---
            # bot = WLHopperBot()
            # bot.iniciar(final_u, final_p)
            # ... resto del código ...
