import streamlit as st
import os
import shutil
import pandas as pd
import zipfile
import datetime
import base64
from io import BytesIO
from scraper import WLHopperBot
from utils import extraer_internos, extraer_texto_de_archivo, asegurar_carpeta
import streamlit.components.v1 as components
import pytesseract
from PIL import Image, ImageEnhance
from cryptography.fernet import Fernet
from supabase import create_client

# --- DIAGNÓSTICO DE EMERGENCIA ---
try:
    st.write("🔍 Verificando llaves...")
    keys = list(st.secrets.keys())
    st.write(f"Llaves encontradas: {keys}")
    
    # Probamos el cifrado rápido
    from cryptography.fernet import Fernet
    test_key = st.secrets["FERNET_KEY"].encode()
    test_cipher = Fernet(test_key)
    st.success("✅ Cifrado OK")
    
    # Probamos conexión a DB
    st.write("📡 Conectando a Supabase...")
    from supabase import create_client
    test_db = create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])
    st.success("✅ Supabase conectado")

except Exception as e:
    st.error(f"❌ ERROR CRÍTICO EN EL ARRANQUE: {e}")
    st.stop() # Detenemos todo para ver el error

# ACA TERMINA EL DEBUG

# --- CONFIGURACIÓN DE SEGURIDAD Y DB ---
# La FERNET_KEY debe estar en st.secrets
FERNET_KEY = st.secrets["FERNET_KEY"].encode()
cipher = Fernet(FERNET_KEY)

supabase_url = st.secrets["SUPABASE_URL"]
supabase_key = st.secrets["SUPABASE_KEY"]
db = create_client(supabase_url, supabase_key)

def encriptar(texto): return cipher.encrypt(texto.encode()).decode()
def desencriptar(token): return cipher.decrypt(token.encode()).decode()

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="WL Hopper 2.0 - Sullair Argentina", page_icon="img/favicon.png", layout="wide")

# [Mantener aquí tus bloques de CSS originales]
VERDE_SULLAIR = "#008657"

# --- FUNCIONES DE BANCO DE DATOS ---
def registrar_metrica(equipo, fuente, modo_pruebas):
    if modo_pruebas: return
    try:
        db.table("metricas").insert({
            "usuario": st.session_state["username"],
            "equipo": equipo,
            "fuente": fuente,
            "fecha": datetime.datetime.now().isoformat(),
            "minutos_ahorrad": 3
        }).execute()
    except: pass

def obtener_credenciales_guardadas(sitio):
    try:
        res = db.table("credenciales_sitios").select("*").eq("usuario_app", st.session_state["username"]).eq("sitio", sitio).execute()
        return res.data[0] if res.data else None
    except: return None

# --- CONTROL DE ACCESO ---
def check_password():
    def password_entered():
        if st.session_state["username"] in st.secrets["passwords"] and \
           st.session_state["password"] == st.secrets["passwords"][st.session_state["username"]]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct") is not True:
        # [Mantener aquí tu bloque visual de Login original]
        return False
    return True

if check_password():
    # --- DASHBOARD ADMIN Y MODO PRUEBAS ---
    if st.session_state["username"] == "fcendra":
        with st.sidebar:
            st.divider()
            st.markdown("### 📈 Panel de Control Admin")
            modo_pruebas = st.checkbox("🛠️ Modo Pruebas", value=True, help="No registra métricas en la DB.")
            try:
                met_res = db.table("metricas").select("minutos_ahorrad").execute()
                total_certs = len(met_res.data)
                total_horas = (total_certs * 3) / 60
                st.metric("Certificados Procesados", total_certs)
                st.metric("Tiempo Ahorrado (Hs)", f"{total_horas:.1f}")
            except: st.caption("Error al conectar con métricas.")
    else:
        modo_pruebas = False

    # --- SIDEBAR ESTÁNDAR ---
    with st.sidebar:
        st.markdown("### 🛠️ Opciones")
        if st.button("Cerrar Sesión", use_container_width=True):
            st.session_state.clear()
            st.rerun()

    # --- BANCO DE DATOS DE CREDENCIALES ---
    with st.expander("🔐 Mis Cuentas Guardadas"):
        st.markdown("<small>Tus credenciales se almacenan bajo cifrado Fernet de grado militar.</small>", unsafe_allow_html=True)
        col_c1, col_c2 = st.columns(2)
        
        with col_c1:
            st.caption("Worklift")
            cred_wl = obtener_credenciales_guardadas("WL")
            saved_u_wl = desencriptar(cred_wl["user_enc"]) if cred_wl else ""
            u_wl = st.text_input("Usuario WL", value=saved_u_wl)
            p_wl = st.text_input("Password WL", type="password")
            if st.button("💾 Guardar WL"):
                db.table("credenciales_sitios").upsert({
                    "usuario_app": st.session_state["username"], "sitio": "WL",
                    "user_enc": encriptar(u_wl), "pass_enc": encriptar(p_wl)
                }, on_conflict="usuario_app,sitio").execute()
                st.toast("Credenciales WL blindadas ✅")

        with col_c2:
            st.caption("Bureau Veritas (Próximamente)")
            st.info("Espacio reservado para el motor BV.")

    # --- CUERPO DE LA APP ---
    st.markdown("<h5 style='text-align: center; color: #555;'>Automatización de Descarga de Certificados</h5>", unsafe_allow_html=True)
    
    col_left, col_right = st.columns([1, 1.2], gap="large")
    
    with col_left:
        with st.container(border=True):
            bajar_cert = st.checkbox("Descargar Certificados", value=True)
            bajar_inf = st.checkbox("Descargar Informes", value=False)
            es_semestral = st.checkbox("Vencimiento Semestral (180 días)")
            
            # Recuperar credenciales para el proceso
            f_u_wl = u_wl if u_wl else saved_u_wl
            f_p_wl = p_wl if p_wl else (desencriptar(cred_wl["pass_enc"]) if cred_wl else "")
            
        archivo_subido = st.file_uploader("Subí tu lista", type=['txt', 'csv', 'xlsx', 'png', 'jpg', 'jpeg'])
        texto_internos = st.text_area("O pegá el texto acá:", height=100)
        btn_run = st.button("🚀 COMENZAR PROCESO", use_container_width=True)

    with col_right:
        terminal_placeholder = st.empty()
        # [Mantener tu función render_terminal() aquí]

    if btn_run:
        if not f_u_wl or not f_p_wl:
            st.error("Faltan credenciales de Worklift.")
        else:
            # [Lógica de procesamiento igual a tu original, pero agregando métricas]
            # Ejemplo al final del loop exitoso:
            # registrar_metrica(int_id, "Worklift", modo_pruebas)
            pass

    # --- ZONA DE DESCARGAS ---
    if st.session_state.get("proceso_completo"):
        st.divider()
        dcol1, dcol2, dcol3 = st.columns(3)
        # [En dcol2 y dcol3 agregar st.text_input para nombres de archivo antes del download_button]
