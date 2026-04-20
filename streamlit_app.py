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
st.set_page_config(page_title="WL Hopper - Sullair Argentina", page_icon="img/favicon.png", layout="wide")

# --- FUNCIÓN DE LOGIN (Control de Acceso) ---
def check_password():
    """Devuelve True si el usuario ingresó credenciales válidas."""
    def password_entered():
        if st.session_state["username"] in st.secrets["passwords"] and \
           st.session_state["password"] == st.secrets["passwords"][st.session_state["username"]]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Limpieza de seguridad
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # Pantalla de Login
        st.write("# 🚀 Acceso a WL Hopper")
        with st.form("login_form"):
            st.text_input("Usuario", key="username")
            st.text_input("Contraseña", type="password", key="password")
            st.form_submit_button("Ingresar", on_click=password_entered)
        
        if "password_correct" in st.session_state and not st.session_state["password_correct"]:
            st.error("😕 Usuario o contraseña incorrectos")
        return False
    return True

# --- FLUJO PRINCIPAL ---
if check_password():
    # Solo se ejecuta si el login es exitoso
    
    # Botón para cerrar sesión en la sidebar
    st.sidebar.button("Cerrar Sesión", on_click=lambda: st.session_state.clear())

    # --- ESTILOS CSS ---
    VERDE_SULLAIR = "#008657"
    st.markdown(f"""
        <style>
        .terminal-box {{
            background-color: #212529; color: #f8f9fa; font-family: 'Consolas', monospace;
            font-size: 13px; padding: 15px; border-radius: 5px; height: 522px; 
            overflow-y: auto; border: 1px solid #444;
        }}
        [data-testid="stColumn"] {{
            position: relative;
            display: flex;
            flex-direction: column;
            justify-content: flex-end;
        }}
        .stButton>button {{
            background-color: {VERDE_SULLAIR};
            color: white;
        }}
        </style>
    """, unsafe_allow_html=True)

    # --- ENCABEZADO ---
    col1, col2 = st.columns([1, 3])
    with col1:
        # Intenta cargar el logo si ya creaste la carpeta img
        try:
            st.image("img/WL Hopper Logo - nspc.png", width=250)
        except:
            st.write("### WL HOPPER")
            
    with col2:
        st.title("Gestor de Certificados Worklift")
        st.subheader("Automatización de descargas y validación de vencimientos")

    # --- ESTADO DE SESIÓN ---
    if "log_history" not in st.session_state: st.session_state.log_history = []
    if "proceso_completo" not in st.session_state: st.session_state.proceso_completo = False
    if "reporte_excel" not in st.session_state: st.session_state.reporte_excel = ""
    if "hay_archivos" not in st.session_state: st.session_state.hay_archivos = False

    # --- LAYOUT DE COLUMNAS ---
    c1, c2 = st.columns([1, 1.5], gap="large")

    with c1:
        st.info("📌 Paso 1: Configurar credenciales")
        user = st.text_input("Usuario Worklift", value="", placeholder="ejemplo@sullair.com.ar")
        pw = st.text_input("Contraseña", type="password")
        
        st.divider()
        st.info("⚙️ Paso 2: Listado de Internos")
        # ACÁ ESTÁ TU VENTANITA DE SIEMPRE PARA PEGAR TEXTO
        texto_sucio = st.text_area("Pegá acá el texto del mail o la lista:", height=150, placeholder="E040230, 3797...")
        
        c_cert, c_inf = st.columns(2)
        b_cert = c_cert.checkbox("Descargar Certificados", value=True)
        b_inf = c_inf.checkbox("Descargar Informes", value=False)
        
        if st.button("🚀 COMENZAR PROCESO", use_container_width=True):
            if not user or not pw or not texto_sucio:
                st.error("Faltan datos obligatorios (Usuario, Clave o Internos).")
            else:
                asegurar_carpeta("descargas_temp")
                internos = extraer_internos(texto_sucio)
                
                if not internos:
                    st.warning("No se detectaron números de internos válidos.")
                else:
                    st.session_state.log_history = [f"🔍 Detectados {len(internos)} internos."]
                    bot = WLHopperBot(headless=True)
                    
                    if bot.iniciar(user, pw):
                        st.session_state.log_history.append("✅ Login exitoso en Worklift.")
                        resultados = []
                        
                        progress_bar = st.progress(0)
                        for i, interno in enumerate(internos):
                            st.session_state.log_history.append(f"⚙️ Procesando: {interno}...")
                            res = bot.procesar_interno(interno, "descargas_temp", b_cert, b_inf)
                            st.session_state.log_history += res["log"]
                            
                            resultados.append({
                                "Interno": interno,
                                "Vencimiento": res["venc"],
                                "Inspección": res["insp"],
                                "Certificado": res["cert"],
                                "Informe": res["inf"],
                                "Estado": res["status"],
                                "Detalle": res["det"]
                            })
                            progress_bar.progress((i + 1) / len(internos))
                        
                        bot.cerrar()
                        
                        # Generar Reporte para copiar
                        df = pd.DataFrame(resultados)
                        st.session_state.reporte_excel = df.to_csv(index=False, sep='\t')
                        st.session_state.proceso_completo = True
                        st.session_state.hay_archivos = any(res["cert"] == "SI" or res["inf"] == "SI" for res in resultados)
                        st.session_state.log_history.append("🏁 ¡PROCESO FINALIZADO!")
                    else:
                        st.session_state.log_history.append("❌ Error de Login en Worklift.")

    with c2:
        st.write("📟 Consola de Proceso")
        log_content = "\n".join(st.session_state.log_history)
        st.markdown(f'<div class="terminal-box"><pre>{log_content}</pre></div>', unsafe_allow_html=True)
        
        st.divider()
        dcol1, dcol2 = st.columns(2)
        
        with dcol1:
            if st.session_state.proceso_completo:
                reporte_js = st.session_state.reporte_excel.replace("'", "\\'").replace("\n", "\\n")
                components.html(f"""
                    <button id="copyBtn" style="width:100%; height:45px; background-color:{VERDE_SULLAIR}; color:white; border:none; border-radius:5px; cursor:pointer; font-weight:bold;">
                        📋 Copiar Reporte para Excel
                    </button>
                    <script>
                    const btn = document.getElementById('copyBtn');
                    btn.onclick = () => {{
                        const data = [new ClipboardItem({{ "text/plain": new Blob(['{reporte_js}'], {{ type: "text/plain" }}) }})];
                        navigator.clipboard.write(data).then(() => {{
                            btn.innerHTML = "✅ ¡REPORTE COPIADO!";
                            btn.style.backgroundColor = "#28a745";
                            setTimeout(() => {{
                                btn.innerHTML = "📋 Copiar Reporte para Excel";
                                btn.style.backgroundColor = "{VERDE_SULLAIR}";
                            }}, 2000);
                        }});
                    }};
                    </script>
                """, height=45) 
            else:
                st.button("📋 Copiar Reporte para Excel", disabled=True, use_container_width=True)

        with dcol2:
            z_buf = BytesIO()
            if st.session_state.proceso_completo and st.session_state.hay_archivos:
                with zipfile.ZipFile(z_buf, "a", zipfile.ZIP_DEFLATED, False) as zf:
                    for r, d, files in os.walk("descargas_temp"):
                        for f in files: zf.write(os.path.join(r, f), f)
            
            st.download_button(
                "📂 Descargar Archivo ZIP", 
                data=z_buf.getvalue(), 
                file_name="certificados.zip", 
                disabled=not (st.session_state.proceso_completo and st.session_state.hay_archivos),
                use_container_width=True
            )

    st.divider()
    st.caption("© 2026 - Desarrollado por Fede García Cendra para Sullair Argentina S.A.")
    st.caption("Consultas: fcendra@sullair.com.ar")
