import streamlit as st
import os
import pandas as pd
from scraper import WLHopperBot
from utils import extraer_internos, asegurar_carpeta
import zipfile
from io import BytesIO

# Configuración de página
st.set_page_config(page_title="Sullair Argentina - WL Hopper", page_icon="🚀")

st.title("🚀 WL Hopper v1.0 (Web Edition)")
st.markdown("Automatización de certificados Worklift")

# --- SIDEBAR: Credenciales ---
with st.sidebar:
    st.header("Configuración")
    user = st.text_input("Usuario (Email)", value="")
    pw = st.text_input("Contraseña", type="password")
    
    st.divider()
    bajar_cert = st.checkbox("Descargar Certificados", value=True)
    bajar_inf = st.checkbox("Descargar Informes", value=False)

# --- CUERPO: Entrada de Internos ---
texto_internos = st.text_area("Pegá los internos acá (pueden venir de Excel, separados por comas o espacios):", height=150)

if st.button("Iniciar Proceso"):
    if not user or not pw:
        st.error("Che, faltan las credenciales.")
    elif not texto_internos.strip():
        st.warning("No pusiste ningún interno para procesar.")
    else:
        lista_internos = extraer_internos(texto_internos)
        st.info(f"Detectados {len(lista_internos)} internos únicos.")
        
        # Carpeta temporal para esta sesión
        ruta_temp = "descargas_temp"
        asegurar_carpeta(ruta_temp)
        
        resultados = []
        progreso = st.progress(0)
        status_text = st.empty()
        log_placeholder = st.expander("Ver Log de Consola", expanded=True)

        # Iniciamos el Bot (pasamos headless=True para el servidor)
        bot = WLHopperBot(headless=True)
        
        if bot.iniciar(user, pw):
            for i, interno in enumerate(lista_internos):
                status_text.text(f"Procesando: {interno}...")
                
                # Llamamos a tu función original de scraper.py
                res = bot.procesar_interno(interno, ruta_temp, bajar_cert, bajar_inf)
                
                # Guardar resultado para el Excel final
                res['interno'] = interno
                resultados.append(res)
                
                # Mostrar logs en la web
                with log_placeholder:
                    st.write(f"**{interno}:** {res['status']} - {res['det']}")
                
                progreso.progress((i + 1) / len(lista_internos))
            
            bot.cerrar()
            st.success("¡Proceso finalizado!")

            # --- DESCARGAS ---
            # 1. Botón para el Excel de resultados
            df = pd.DataFrame(resultados)
            excel_data = BytesIO()
            df.to_excel(excel_data, index=False)
            st.download_button("📊 Descargar Reporte Excel", excel_data.getvalue(), "reporte_hopper.xlsx")

            # 2. ZIP con los PDFs
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for root, dirs, files in os.walk(ruta_temp):
                    for file in files:
                        zip_file.write(os.path.join(root, file), file)
            
            st.download_button("📂 Descargar Certificados (ZIP)", zip_buffer.getvalue(), "certificados.zip")
            
        else:
            st.error("Error de login. Revisá el usuario y la clave.")
