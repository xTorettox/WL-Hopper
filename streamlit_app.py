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

# --- ESTILOS CSS ---
    
VERDE_SULLAIR = "#008657"
st.markdown(f"""
    <style>
    /* Terminal alineada a la base del botón 'Comenzar' */
    .terminal-box {{
        background-color: #212529; color: #f8f9fa; font-family: 'Consolas', monospace;
        font-size: 13px; padding: 15px; border-radius: 5px; height: 522px; 
        overflow-y: auto; border: 1px solid #444;
    }}
        
    /* Forzar a las columnas a ser contenedores relativos */
    [data-testid="stColumn"] {{
        position: relative;
        display: flex;
        flex-direction: column;
        justify-content: flex-end;
    }}
    
    /* Estilo para que el botón deshabilitado no flote */
    .stDownloadButton button {{
        margin-bottom: 0px !important;
        height: 45px !important;
    }}
    
    div.stButton > button:first-child {{ background-color: {VERDE_SULLAIR} !important; color: white !important; font-weight: bold; }}
    .log-entry {{ margin-bottom: 5px; border-bottom: 1px solid #333; padding-bottom: 2px; }}
    .logo-container {{ display: flex; justify-content: center; align-items: center; flex-direction: column; margin-bottom: 10px; }}
    </style>
    """, unsafe_allow_html=True)

# --- FUNCIÓN DE LOGIN (Control de Acceso) ---
# --- FUNCIÓN DE LOGIN (Control de Acceso) ---
def check_password():
    """Devuelve True si el usuario ingresó credenciales válidas."""
    def password_entered():
        # Verificamos contra los secrets de Streamlit
        if st.session_state["username"] in st.secrets["passwords"] and \
           st.session_state["password"] == st.secrets["passwords"][st.session_state["username"]]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Borramos la clave por seguridad
        else:
            st.session_state["password_correct"] = False

    # Si NO es True (es decir, es False o None), mostramos el login
    if st.session_state.get("password_correct") is not True:
        st.markdown("<br><br>", unsafe_allow_html=True)
        c_l1, c_l2, c_l3 = st.columns([1.2, 1, 1.2])
        
        with c_l2:
            # El TRY evita que la app muera si el nombre del archivo está mal
            try:
                # REVISÁ ESTE NOMBRE: Debe ser idéntico al de tu carpeta img
                st.image("img/WL Hopper Logo - nspc.png", use_container_width=True)
            except:
                st.markdown("<h2 style='text-align: center;'>🚀 WL Hopper</h2>", unsafe_allow_html=True)
            
            st.markdown("<h4 style='text-align: center;'>Acceso al Sistema</h4>", unsafe_allow_html=True)
            
            with st.form("login_form"):
                st.text_input("Usuario", key="username")
                st.text_input("Contraseña", type="password", key="password")
                st.form_submit_button("Ingresar", on_click=password_entered, use_container_width=True)
            
            # Solo mostramos el error si el usuario ya intentó y falló
            if st.session_state.get("password_correct") == False:
                st.error("😕 Usuario o contraseña incorrectos")
        
        return False  # <--- Corta la ejecución aquí, no deja pasar al resto de la app
        
    return True # <--- Solo llega acá si el login fue exitoso

# --- INICIO DE LA APLICACIÓN (Si el login es exitoso) ---
if check_password():
    # Botón de Logout en el sidebar
    st.sidebar.button("Cerrar Sesión", on_click=lambda: st.session_state.clear())

    # --- ENCABEZADO PRINCIPAL ---
    col1, col2 = st.columns([1, 3])
    with col1:
        try:
            st.image("img/WL Hopper Logo - nspc.png", width=200)
        except:
            st.write("### WL HOPPER")
            
    with col2:
        st.title("Gestor de Certificados Worklift")
        st.subheader("Automatización de descargas y validación de vencimientos")

    # --- INICIALIZAR ESTADO DE SESIÓN ---
    if "log_history" not in st.session_state:
        st.session_state.log_history = []
    if "proceso_completo" not in st.session_state:
        st.session_state.proceso_completo = False
    if "reporte_excel" not in st.session_state:
        st.session_state.reporte_excel = ""
    if "hay_archivos" not in st.session_state:
        st.session_state.hay_archivos = False

    # --- LAYOUT DE COLUMNAS ---
    c1, c2 = st.columns([1, 1.5], gap="large")

    with c1:
        st.info("📌 Paso 1: Configurar credenciales")
        user = st.text_input("Usuario Worklift", value="", placeholder="ejemplo@sullair.com.ar")
        pw = st.text_input("Contraseña Worklift", type="password")
        
        st.divider()
        st.info("⚙️ Paso 2: Listado de Internos")
        texto_sucio = st.text_area("Pegá acá el texto del mail o la lista de internos:", height=150, placeholder="E040230, 3797, A060124...")
        
        c_cert, c_inf = st.columns(2)
        b_cert = c_cert.checkbox("Descargar Certificados", value=True)
        b_inf = c_inf.checkbox("Descargar Informes", value=False)
        
        if st.button("🚀 COMENZAR PROCESO", use_container_width=True):
            if not user or not pw or not texto_sucio:
                st.error("Por favor, completa todos los campos obligatorios.")
            else:
                asegurar_carpeta("descargas_temp")
                internos = extraer_internos(texto_sucio)
                
                if not internos:
                    st.warning("No se detectaron números de interno válidos en el texto.")
                else:
                    st.session_state.log_history = [f"🔍 Se detectaron {len(internos)} internos para procesar."]
                    bot = WLHopperBot(headless=True)
                    
                    if bot.iniciar(user, pw):
                        st.session_state.log_history.append("✅ Login exitoso en Worklift.")
                        resultados = []
                        progress_bar = st.progress(0)
                        
                        for i, interno in enumerate(internos):
                            st.session_state.log_history.append(f"⚙️ Procesando: {interno}...")
                            res = bot.procesar_interno(interno, "descargas_temp", b_cert, b_inf)
                            
                            # Alertas visuales en terminal
                            st_text = res['status'].upper()
                            if "VIGENTE" in st_text:
                                st.session_state.log_history.append(f"  ✅ Certificado VIGENTE (Vence: {res['venc']})")
                            elif "PRÓXIMO" in st_text:
                                st.session_state.log_history.append(f"  ⚠️ ¡ALERTA! PRÓXIMO A VENCER ({res['venc']})")
                            elif "VENCIDO" in st_text:
                                st.session_state.log_history.append(f"  ❌ CERTIFICADO VENCIDO ({res['venc']})")
                            
                            st.session_state.log_history += res.get("log", [])
                            
                            resultados.append({
                                "Interno": interno,
                                "Vencimiento": res["venc"],
                                "Estado": res["status"],
                                "Cert": res["cert"],
                                "Inf": res["inf"]
                            })
                            progress_bar.progress((i + 1) / len(internos))
                        
                        bot.cerrar()
                        df = pd.DataFrame(resultados)
                        st.session_state.reporte_excel = df.to_csv(index=False, sep='\t')
                        st.session_state.proceso_completo = True
                        st.session_state.hay_archivos = any(r["Cert"] == "SI" or r["Inf"] == "SI" for r in resultados)
                        st.session_state.log_history.append("🏁 ¡PROCESO FINALIZADO!")
                    else:
                        st.session_state.log_history.append("❌ Error de Login: Credenciales de Worklift incorrectas.")

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
                            const originalText = btn.innerHTML;
                            btn.innerHTML = "✅ ¡REPORTE COPIADO!";
                            btn.style.backgroundColor = "#28a745";
                            setTimeout(() => {{
                                btn.innerHTML = originalText;
                                btn.style.backgroundColor = "{VERDE_SULLAIR}";
                            }}, 2000);
                        }});
                    }};
                    </script>
                """, height=45) 
            else:
                st.button("📋 Copiar Reporte para Excel", disabled=True, use_container_width=True, key="btn_copy_off")

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
                use_container_width=True,
                key="btn_zip"
            )

    st.divider()
    st.caption("© 2026 - Desarrollado por Fede García Cendra para Sullair Argentina S.A.")
    st.caption("Consultas a: fcendra@sullair.com.ar")
