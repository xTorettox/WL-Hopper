import streamlit as st
import os
import shutil
import pandas as pd
import zipfile
import datetime
from io import BytesIO
from scraper import WLHopperBot
from utils import extraer_internos, extraer_texto_de_archivo, asegurar_carpeta
import streamlit.components.v1 as components
import pytesseract
from PIL import Image, ImageEnhance
from cryptography.fernet import Fernet
from supabase import create_client

# --- 1. CONFIGURACIÓN DE SEGURIDAD Y DB ---
try:
    # La FERNET_KEY debe estar en st.secrets
    FERNET_KEY = st.secrets["FERNET_KEY"].encode()
    cipher = Fernet(FERNET_KEY)

    supabase_url = st.secrets["SUPABASE_URL"]
    supabase_key = st.secrets["SUPABASE_KEY"]
    db = create_client(supabase_url, supabase_key)
    db_ready = True
except Exception as e:
    db_ready = False
    st.error(f"Error de conexión inicial: {e}")

def encriptar(texto): 
    return cipher.encrypt(texto.encode()).decode() if texto else ""

def desencriptar(token): 
    try:
        return cipher.decrypt(token.encode()).decode() if token else ""
    except:
        return ""

# --- 2. CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="WL Hopper 2.0 - Sullair Argentina", page_icon="img/favicon.png", layout="wide")

# Estilos CSS (Versión simplificada para evitar conflictos de renderizado)
st.markdown("""
    <style>
    .terminal-box {
        background-color: #212529; color: #f8f9fa; font-family: monospace;
        padding: 15px; border-radius: 5px; border: 1px solid #444;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 3. FUNCIONES DE BANCO DE DATOS ---
def registrar_metrica(equipo, fuente, modo_pruebas):
    if modo_pruebas or not db_ready: return
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
    if not db_ready: return None
    try:
        res = db.table("credenciales_sitios").select("*").eq("usuario_app", st.session_state["username"]).eq("sitio", sitio).execute()
        return res.data[0] if res.data else None
    except: return None

# --- 4. CONTROL DE ACCESO ---
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if not st.session_state["password_correct"]:
        st.title("🔐 Acceso WL Hopper")
        user = st.text_input("Usuario", key="username")
        password = st.text_input("Contraseña", type="password", key="password")
        if st.button("Ingresar"):
            if user in st.secrets["passwords"] and password == st.secrets["passwords"][user]:
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos")
        return False
    return True

# --- 5. CUERPO DE LA APLICACIÓN ---
if check_password():
    
    # Lógica de Modo Pruebas para Admin
    if st.session_state["username"] == "fcendra":
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
                except: st.caption("Error cargando métricas")
    else:
        modo_pruebas = False

    with st.sidebar:
        st.markdown("### 🛠️ Opciones")
        if st.button("Cerrar Sesión"):
            st.session_state.clear()
            st.rerun()

    # Gestión de Credenciales
    with st.expander("🔐 Mis Cuentas Guardadas"):
        st.info("Tus datos se guardan encriptados (Fernet).")
        col1, col2 = st.columns(2)
        
        with col1:
            st.caption("Worklift")
            cred_wl = obtener_credenciales_guardadas("WL")
            s_u_wl = desencriptar(cred_wl["user_enc"]) if cred_wl else ""
            input_u_wl = st.text_input("Usuario WL", value=s_u_wl)
            input_p_wl = st.text_input("Password WL", type="password")
            if st.button("💾 Guardar WL"):
                if db_ready:
                    db.table("credenciales_sitios").upsert({
                        "usuario_app": st.session_state["username"], "sitio": "WL",
                        "user_enc": encriptar(input_u_wl), "pass_enc": encriptar(input_p_wl)
                    }, on_conflict="usuario_app,sitio").execute()
                    st.toast("Guardado con éxito ✅")

        with col2:
            st.caption("Bureau Veritas")
            st.write("Próximamente disponible.")

    # Interfaz de Proceso
    col_l, col_r = st.columns([1, 1.2])
    
    with col_l:
        with st.container(border=True):
            b_cert = st.checkbox("Descargar Certificados", value=True)
            b_inf = st.checkbox("Descargar Informes", value=False)
            
            # Credenciales finales
            final_u = input_u_wl if input_u_wl else s_u_wl
            final_p = input_p_wl if input_p_wl else (desencriptar(cred_wl["pass_enc"]) if cred_wl else "")
            
        archivo = st.file_uploader("Subir lista", type=['txt', 'xlsx', 'png', 'jpg'])
        texto = st.text_area("O pegar internos aquí:")
        btn_run = st.button("🚀 COMENZAR PROCESO", use_container_width=True)

    with col_r:
        st.markdown("**Estado del Proceso**")
        terminal = st.empty()
        terminal.markdown("<div class='terminal-box'>Esperando inicio...</div>", unsafe_allow_html=True)

    if btn_run:
        if not final_u or not final_p:
            st.error("Cargá tus credenciales de Worklift primero.")
        else:
            lista = extraer_internos(texto + (extraer_texto_de_archivo(archivo) if archivo else ""))
            if not lista:
                st.warning("No se encontraron números de internos.")
            else:
                # Aquí llamarías a tu bot original
                st.info(f"Procesando {len(lista)} equipos...")
                # Ejemplo de registro al finalizar cada equipo con éxito:
                # registrar_metrica(equipo_id, "WL", modo_pruebas)

    # Botones de descarga (Aparecen si hay resultados)
    if st.session_state.get("proceso_completo"):
        st.divider()
        d1, d2, d3 = st.columns(3)
        with d2:
            nom_xls = st.text_input("Nombre Excel", value=f"Reporte_{datetime.date.today()}.xlsx")
            st.button("📊 Descargar Excel") # Aquí va tu st.download_button original
        with d3:
            nom_zip = st.text_input("Nombre ZIP", value=f"Certs_{datetime.date.today()}.zip")
            st.button("📂 Descargar ZIP") # Aquí va tu st.download_button original

    st.caption(f"© {datetime.datetime.now().year} - Sullair Argentina")
