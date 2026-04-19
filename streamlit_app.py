import streamlit as st
import os
import pandas as pd
import zipfile
from io import BytesIO
from scraper import WLHopperBot
from utils import extraer_internos, asegurar_carpeta

# --- SOLUCIÓN PARA ENTORNOS DE NUBE (Codespaces/Streamlit Cloud) ---
if "playwright_installed" not in st.session_state:
    with st.spinner("Configurando motores de navegación..."):
        os.system("playwright install chromium")
    st.session_state.playwright_installed = True

# --- CONFIGURACIÓN ESTÉTICA ---
st.set_page_config(
    page_title="Sullair Argentina - WL Hopper", 
    page_icon="🚀",
    layout="centered"
)

# Intentamos cargar el logo (asegurate que la carpeta img esté en tu repo)
try:
    st.image("img/WL Hopper Logo - nspc.png", width=300)
except:
    st.title("🚀 WL Hopper v1.0")

st.markdown("### Automatización de Certificados Worklift")
st.info("Esta versión web procesa los internos y te permite descargar los resultados al finalizar.")

# --- SIDEBAR: CREDENCIALES Y OPCIONES ---
with st.sidebar:
    st.header("Configuración")
    user = st.text_input("Usuario (Email)", key="user_email")
    pw = st.text_input("Contraseña", type="password", key="user_pw")
    
    st.divider()
    st.subheader("Opciones de descarga")
    bajar_cert = st.checkbox("Descargar Certificados", value=True)
    bajar_inf = st.checkbox("Descargar Informes", value=False)
    
    st.divider()
    st.caption("Sullair Argentina - WL Hopper Web Edition")

# --- CUERPO PRINCIPAL ---
texto_internos = st.text_area(
    "Pegá los internos acá (pueden venir de Excel, separados por comas, espacios o saltos de línea):", 
    height=150,
    placeholder="E040230, 3797, A001234..."
)

if st.button("🚀 Iniciar Proceso"):
    if not user or not pw:
        st.error("Che, faltan las credenciales en la barra lateral.")
    elif not texto_internos.strip():
        st.warning("No pusiste ningún interno para procesar.")
    else:
        lista_internos = extraer_internos(texto_internos)
        st.info(f"Se detectaron {len(lista_internos)} internos únicos para procesar.")
        
        # Carpeta temporal para guardar los PDFs antes del ZIP
        ruta_temp = "descargas_temp"
        asegurar_carpeta(ruta_temp)
        
        resultados = []
        progreso = st.progress(0)
        status_text = st.empty()
        log_container = st.expander("Ver Log de Actividad", expanded=True)

        # Iniciamos el Bot con Headless=True (obligatorio en la nube)
        bot = WLHopperBot(headless=True)
        
        if bot.iniciar(user, pw):
            for i, interno in enumerate(lista_internos):
                status_text.text(f"Procesando interno: {interno}")
                
                # Ejecutamos la lógica de tu scraper.py
                res = bot.procesar_interno(interno, ruta_temp, bajar_cert, bajar_inf)
                res['id'] = interno  # Seteamos el ID para el Excel
                resultados.append(res)
                
                # Mostramos progreso en la interfaz
                with log_container:
                    icono = "✅" if "OK" in res['status'] or "VIGENTE" in res['status'] else "⚠️"
                    st.write(f"{icono} **{interno}:** {res['status']} - {res['det']}")
                
                progreso.progress((i + 1) / len(lista_internos))
            
            bot.cerrar()
            status_text.success("🏁 ¡Proceso completado!")