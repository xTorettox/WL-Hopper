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

# --- NUEVOS IMPORTS SEGURIDAD/DB ---
from supabase import create_client, Client
from cryptography.fernet import Fernet

importlib.reload(utils)

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="WL Hopper - Sullair Argentina", page_icon="img/favicon.png", layout="wide")

# --- BÚNKER DE SEGURIDAD Y CONEXIÓN (Punto 2) ---
@st.cache_resource
def init_supabase() -> Client:
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])

supabase = init_supabase()
cipher_suite = Fernet(st.secrets["FERNET_KEY"])

def encriptar(texto: str) -> str:
    return cipher_suite.encrypt(texto.encode()).decode()

def desencriptar(token: str) -> str:
    try:
        return cipher_suite.decrypt(token.encode()).decode()
    except:
        return ""

def registrar_metrica(interno, usuario):
    # Solo registramos si no es modo pruebas (puedes ajustar esta lógica)
    if not st.session_state.get("modo_pruebas", False):
        supabase.table("metricas_descargas").insert({
            "interno": interno, 
            "usuario_ejecutor": usuario,
            "herramienta": "WL Hopper"
        }).execute()

# --- ESTILOS CSS (Sin cambios) ---
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
    }}
    div.stButton > button:first-child {{ background-color: {VERDE_SULLAIR} !important; color: white !important; font-weight: bold; }}
    .log-entry {{ margin-bottom: 5px; border-bottom: 1px solid #333; padding-bottom: 2px; }}
    </style>
    """, unsafe_allow_html=True)

# --- EVOLUCIÓN DEL LOGIN (Punto 3) ---
def check_password():
    def password_entered():
        if st.session_state["username"] in st.secrets["passwords"] and \
           st.session_state["password"] == st.secrets["passwords"][st.session_state["username"]]:
            st.session_state["password_correct"] = True
            st.session_state["logged_user"] = st.session_state["username"]
            
            # Fetch inicial de credenciales guardadas
            res = supabase.table("credenciales_sitios").select("*").eq("usuario_app", st.session_state["logged_user"]).execute()
            if res.data:
                st.session_state["u_saved"] = res.data[0].get("user_worklift", "")
                st.session_state["p_saved"] = desencriptar(res.data[0].get("pass_worklift", ""))
            
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct") is not True:
        c_l1, c_l2, c_l3 = st.columns([1.2, 1, 1.2])
        with c_l2:
            st.markdown("<h2 style='text-align: center;'>🚀 WL Hopper</h2>", unsafe_allow_html=True)
            with st.form("login_form"):
                st.text_input("Usuario", key="username")
                st.text_input("Contraseña", type="password", key="password")
                st.form_submit_button("Ingresar", on_click=password_entered, use_container_width=True)
            if st.session_state.get("password_correct") == False:
                st.error("😕 Usuario o contraseña incorrectos")
        return False
    return True

# --- FLUJO PRINCIPAL ---
if check_password():
    # Sidebar con Dashboard de Admin (Punto 6)
    with st.sidebar:
        st.markdown("### 🛠️ Opciones")
        if st.session_state.get("logged_user") == "fcendra":
            st.session_state["modo_pruebas"] = st.checkbox("🛠️ Modo Pruebas", value=False)
            try:
                count_res = supabase.table("metricas_descargas").select("*", count="exact").execute()
                total = count_res.count if count_res.count else 0
                st.metric("Certificados Hopper", total, f"+{total*2} min ahorrados")
            except: pass
            
        st.button("Cerrar Sesión", on_click=lambda: st.session_state.clear(), use_container_width=True)

    # UI Principal
    c_l1, c_l2, c_l3 = st.columns([1.5, 1, 1.5])
    with c_l2: 
        try: st.image("img/WL Hopper Logo - nspc.png", use_container_width=True)
        except: pass
    
    # INTERFAZ DE CREDENCIALES (Punto 4)
    with st.expander("🔐 Mis Credenciales de Worklift", expanded=not st.session_state.get("u_saved")):
        u_work = st.text_input("Usuario Worklift", value=st.session_state.get("u_saved", ""))
        p_work = st.text_input("Contraseña Worklift", type="password", value=st.session_state.get("p_saved", ""))
        if st.button("💾 Guardar Credenciales"):
            p_enc = encriptar(p_work)
            supabase.table("credenciales_sitios").upsert({
                "usuario_app": st.session_state["logged_user"],
                "user_worklift": u_work,
                "pass_worklift": p_enc
            }).execute()
            st.session_state["u_saved"] = u_work
            st.session_state["p_saved"] = p_work
            st.success("Credenciales actualizadas.")

    col_left, col_right = st.columns([1, 1.2], gap="large")
    
    with col_left:
        # Usamos las credenciales guardadas por defecto
        user = st.text_input("Usuario (Email)", value=st.session_state.get("u_saved", ""), key="user_email")
        pw = st.text_input("Contraseña", type="password", value=st.session_state.get("p_saved", ""), key="user_pw")
        
        c1, c2 = st.columns(2)
        bajar_cert = c1.checkbox("Descargar Certificados", value=True)
        bajar_inf = c2.checkbox("Descargar Informes", value=False)
        es_semestral = st.checkbox("Vencimiento Semestral (180 días)")
        
        archivo_subido = st.file_uploader("Subí tu archivo o Foto", type=['txt', 'csv', 'xlsx', 'png', 'jpg', 'jpeg'])
        texto_internos = st.text_area("O pegá el texto acá:", height=115)
        btn_run = st.button("🚀 COMENZAR PROCESO", use_container_width=True)

    with col_right:
        terminal_placeholder = st.empty()
        def render_terminal():
            html = f'<div class="terminal-box"><div style="font-weight: bold; color: #ddd; margin-bottom: 10px;">Registro de Actividad</div>'
            for entry in st.session_state.get("log_history", []):
                html += f'<div class="log-entry">{entry}</div>'
            html += '</div>'
            terminal_placeholder.markdown(html, unsafe_allow_html=True)

    if btn_run:
        if not user or not pw: st.error("Faltan credenciales.")
        else:
            # (Lógica de procesamiento igual a la anterior...)
            # [...] 
            bot = WLHopperBot(headless=True)
            if bot.iniciar(user, pw):
                res_lista = []
                lista = extraer_internos(texto_internos + (extraer_texto_de_archivo(archivo_subido) if archivo_subido else ""))
                
                for int_id in lista:
                    res = bot.procesar_interno(int_id, "descargas_temp", bajar_cert, bajar_inf, es_semestral=es_semestral)
                    res['id'] = int_id
                    res_lista.append(res)
                    
                    # INYECTOR DE MÉTRICAS (Punto 5)
                    if res.get("status") and "ERROR" not in res["status"].upper():
                        registrar_metrica(int_id, st.session_state["logged_user"])
                    
                bot.cerrar()
                st.session_state.proceso_completo = True
                st.rerun()
