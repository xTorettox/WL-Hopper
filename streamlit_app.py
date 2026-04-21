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

# --- FLUJO PRINCIPAL ---
if check_password():

    # --- FUNCIÓN DEL MODAL ACERCA DE ---
    @st.dialog("Acerca de WL Hopper")
    def mostrar_about():
        c1, c2 = st.columns([1, 2])
        with c1:
            
            try:
                st.image("img/robot_diapos.png", use_container_width=True)
            except:
                st.write("🤖") # Placeholder
        with c2:
            st.markdown("""
            **WL Hopper** es una solución de automatización diseñada para optimizar la descarga de certificados en PDF desde el sitio de **Worklift**.
            
            Inspirada en una tarea repetitiva que no quería seguirlo siendo, esta herramienta utiliza bots de navegación y un poquitito de inteligencia artificial para centralizar la descarga de certificados y validar vencimientos de internos de forma masiva.
            """)
        
        st.info("🚀 **Misión:** Acelerar la tarea de descarga y/o recuperación de información desde el sitio web de Worklift.")
        
        st.divider()
        st.caption("Desarrollado por Fede García Cendra - 2026")

    # --- SIDEBAR (Menú Lateral) ---
    with st.sidebar:
        st.markdown("### 🛠️ Opciones")
        
        # Botón 1: Cerrar Sesión
        st.button("Cerrar Sesión", on_click=lambda: st.session_state.clear(), use_container_width=True)
        
        # Botón 2: Acerca de
        if st.button("Acerca del Proyecto", use_container_width=True):
            mostrar_about()
    
    # --- INICIALIZACIÓN ---
    if "log_history" not in st.session_state: st.session_state.log_history = []
    if "proceso_completo" not in st.session_state: st.session_state.proceso_completo = False
    if "html_excel" not in st.session_state: st.session_state.html_excel = ""
    if "hay_archivos" not in st.session_state: st.session_state.hay_archivos = False
    
    # --- LAYOUT ---
    col_left, col_right = st.columns([1, 1.2], gap="large")
    
    with col_left:
        st.markdown('<div class="logo-container">', unsafe_allow_html=True)
        c_l1, c_l2, c_l3 = st.columns([1, 2, 1])
        with c_l2: st.image("img/WL Hopper Logo - nspc.png", use_container_width=True)
        st.markdown("<h5 style='text-align: center; color: #555; margin-top:-10px;'>Automatización de Descarga de Certificados</h5>", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        
        with st.container(border=True):
            user = st.text_input("Usuario (Email)", key="user_email", help="Acá podés introducir tu usuario de Worklift")
            pw = st.text_input("Contraseña", type="password", key="user_pw", help="Acá va tu clave de Worklift")
            c1, c2 = st.columns(2)
            bajar_cert = c1.checkbox("Descargar Certificados", value=True)
            bajar_inf = c2.checkbox("Descargar Informes de Inspección", value=False)
    
        st.markdown("##### Listado de Internos")
        texto_internos = st.text_area("Pegá acá:", height=115, placeholder="E040230, 3797...", help="En este cuadro de texto podés pegar tu listado de internos. Acepta texto simple con cualquier separador, listado pegado de Excel, y puede reconocer internos dentro de cualquier texto mientras no estén pegados.")
        btn_run = st.button("🚀 COMENZAR PROCESO", use_container_width=True)
    
    with col_right:
        st.markdown("##### Registro de Actividad")
        terminal_placeholder = st.empty()
        def render_terminal():
            html = f'<div class="terminal-box">'
            for entry in st.session_state.log_history: html += f'<div class="log-entry">{entry}</div>'
            html += '</div>'
            terminal_placeholder.markdown(html, unsafe_allow_html=True)
        render_terminal()
    
    # --- LÓGICA DE PROCESO ---
    if btn_run:
        if not user or not pw: st.error("Faltan credenciales.")
        else:
            ruta_temp = "descargas_temp"
            if os.path.exists(ruta_temp): shutil.rmtree(ruta_temp)
            asegurar_carpeta(ruta_temp)
    
            st.session_state.proceso_completo = False
            st.session_state.log_history = ["Iniciando conexión con Worklift..."]
            render_terminal()
            
            bot = WLHopperBot(headless=True)
            if bot.iniciar(user, pw):
                st.session_state.log_history.append("🔐 Login exitoso.")
                render_terminal()
                
                lista = extraer_internos(texto_internos)
                res_lista = []
                for int_id in lista:
                    st.session_state.log_history.append(f"--- Procesando {int_id} ---")
                    render_terminal()
                    res = bot.procesar_interno(int_id, ruta_temp, bajar_cert, bajar_inf)
                    res['id'] = int_id
                    res_lista.append(res)
                    for m in res.get('log', []): st.session_state.log_history.append(f"&nbsp;&nbsp;{m}")
                    render_terminal()
    
                bot.cerrar()
                st.session_state.log_history.append("🔓 Sesión cerrada correctamente.")
                st.session_state.log_history.append("🏁 PROCESO FINALIZADO.")
                st.session_state.proceso_completo = True
                st.session_state.hay_archivos = len(os.listdir(ruta_temp)) > 0 if os.path.exists(ruta_temp) else False
    
                # Generación de HTML
                html = """<style>table { border-collapse: collapse; } td { white-space: nowrap; text-align: center; vertical-align: middle; mso-number-format: "\\@"; padding: 5px 15px; } th { white-space: nowrap; text-align: center; padding: 10px 20px; }</style>
                <table width="100%" border="1" style="font-family: Calibri;">
                <tr style="background-color: #008657; color: white; font-weight: bold;"><th>INTERNO</th><th>ESTADO</th><th>ÚLTIMA INSPECCIÓN</th><th>VENCIMIENTO</th><th>CERTIFICADO</th><th>INFORME</th><th>DETALLE</th></tr>"""
                for r in res_lista:
                    bg, tx, st_text = "#FFFFFF", "#000000", r['status'].upper()
                    cert_val = "SI" if "VIGENTE" in st_text or "PRÓXIMO" in st_text else "NO"
                    if "VIGENTE" in st_text: bg, tx = "#C6EFCE", "#006100"
                    elif "PRÓXIMO" in st_text: bg, tx = "#FFEB9C", "#9C5700"
                    elif "VENCIDO" in st_text: bg, tx = "#FFC7CE", "#9C0006"
                    html += f'<tr><td>{r["id"]}</td><td style="background-color: {bg}; color: {tx}; font-weight: bold;">{st_text}</td>'
                    html += f'<td>{r["insp"]}</td><td>{r["venc"]}</td><td>{cert_val}</td><td>{r["inf"]}</td>'
                    html += f'<td style="text-align: left; white-space: normal; padding-right: 30px;">{r["det"]}</td></tr>'
                html += "</table>"
                st.session_state.html_excel = html.replace("\n", "")
                st.rerun()
            else:
                # --- Manejo del error de login ---
                st.session_state.log_history.append("❌ ERROR: Credenciales de Worklift incorrectas.")
                render_terminal()
                st.error("No se pudo iniciar sesión. Verificá tu usuario y contraseña de Worklift.")
                
    # --- BOTONES DE ACCIÓN ---
    st.divider()
    dcol1, dcol2 = st.columns(2)
    
    with dcol1:
        if st.session_state.proceso_completo:
            # Usamos un div con padding:0 para que el iframe se pegue al piso
            components.html(f"""
                <div style="margin:0; padding:0; height: 45px; display: flex; align-items: center;">
                    <button id="cBtn" style="
                        width: 100%; 
                        height: 45px; 
                        background-color: {VERDE_SULLAIR}; 
                        color: white; 
                        border: none; 
                        border-radius: 4px; 
                        font-weight: bold; 
                        cursor: pointer; 
                        font-family: sans-serif;
                        box-sizing: border-box;
                    ">
                        📋 Copiar Reporte para Excel
                    </button>
                    <textarea id="hiddenTable" style="position:fixed; top:-1000px; opacity:0;">{st.session_state.html_excel}</textarea>
                </div>
                <script>
                document.getElementById('cBtn').onclick = function() {{
                    const btn = this;
                    const html = document.getElementById('hiddenTable').value;
                    const blob = new Blob([html], {{ type: 'text/html' }});
                    const data = [new ClipboardItem({{ 'text/html': blob }})];
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
        # Definimos z_buf siempre para que Pylance no llore
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
    
    # --- CAPTIONS ---
    st.divider()
    st.caption("© 2026 - Desarrollado por Fede García Cendra para Sullair Argentina S.A.")
    st.caption("Consultas a: fcendra@sullair.com.ar")
