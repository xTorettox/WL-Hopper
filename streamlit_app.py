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

from supabase import create_client, Client
from cryptography.fernet import Fernet
from datetime import datetime

# --- CONFIGURACIÓN BÚNKER DE SEGURIDAD Y BD ---
try:
    SUPABASE_URL = st.secrets["SUPABASE_URL"]
    SUPABASE_KEY = st.secrets["SUPABASE_KEY"]
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
    
    FERNET_KEY = st.secrets["FERNET_KEY"].encode()
    cipher_suite = Fernet(FERNET_KEY)
except Exception as e:
    st.error(f"Error cargando secrets: {e}")
    supabase = None
    cipher_suite = None

def encriptar(texto):
    if not texto or not cipher_suite: return ""
    return cipher_suite.encrypt(texto.encode()).decode()

def desencriptar(texto_cifrado):
    if not texto_cifrado or not cipher_suite: return ""
    try:
        return cipher_suite.decrypt(texto_cifrado.encode()).decode()
    except:
        return ""

def registrar_metrica(interno, fuente):
    if not supabase: return
    try:
        minutos = 5
        data = {
            "usuario": st.session_state.get("logged_user", "desconocido"),
            "equipo": interno,
            "fuente": fuente,
            "fecha": datetime.now().isoformat(),
            "minutos_ahorrados": minutos
        }
        supabase.table("metricas").insert(data).execute()
    except Exception as e:
        print(f"Error registrando métrica: {e}")


# --- CONFIGURACIÓN ---
st.set_page_config(page_title="WL Hopper - Sullair Argentina", page_icon="img/favicon.png", layout="wide")

# --- ESTILOS CSS ---
    
VERDE_SULLAIR = "#008657"
st.markdown(f"""
    <style>
    /* Base de la terminal */
    .terminal-box {{
        background-color: #212529; color: #f8f9fa; font-family: 'Consolas', monospace;
        font-size: 13px; padding: 15px; border-radius: 5px; 
        overflow-y: auto; border: 1px solid #444;
    }}
    
    /* Layout exclusivo para Desktop (pantallas anchas) */
    @media (min-width: 768px) {{
        .terminal-box {{
            position: absolute; top: 0; bottom: 0; left: 0; right: 0; width: 100%; height: 100%; min-height: 535px;
        }}
        [data-testid="stHorizontalBlock"] {{
            align-items: stretch;
        }}
        [data-testid="stColumn"] {{
            position: relative;
        }}
    }}
    
    /* Layout para Móviles */
    @media (max-width: 767px) {{
        .terminal-box {{
            height: 400px;
            margin-top: 10px;
        }}
    }}
    
    /* Estilo para que el botón deshabilitado no flote */
    .stDownloadButton button {{
        margin-bottom: 0px !important;
        height: 45px !important;
    }}
    
    div.stButton > button:first-child {{ background-color: {VERDE_SULLAIR} !important; color: white !important; font-weight: bold; }}
    .log-entry {{ margin-bottom: 5px; border-bottom: 1px solid #333; padding-bottom: 2px; }}
    .logo-container {{ display: flex; justify-content: center; align-items: center; flex-direction: column; margin-bottom: 10px; }}
    
    @keyframes ellipsis {{
      0% {{ content: ""; }}
      25% {{ content: "."; }}
      50% {{ content: ".."; }}
      75% {{ content: "..."; }}
      100% {{ content: ""; }}
    }}
    .loading-dots::after {{
      content: "";
      animation: ellipsis 1.5s infinite;
    }}
    
    @keyframes blink {{
        0% {{ opacity: 1; }}
        50% {{ opacity: 0; }}
        100% {{ opacity: 1; }}
    }}
    .blink-cursor {{
        animation: blink 1s step-end infinite;
    }}
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
            st.session_state["logged_user"] = st.session_state["username"]
            del st.session_state["password"]  # Borramos la clave por seguridad
            
            # Fetch credenciales desde Supabase
            if supabase:
                try:
                    res = supabase.table("credenciales_sitios").select("*").eq("usuario_app", st.session_state["logged_user"]).execute()
                    wl_creds = {}
                    if res.data:
                        for cred in res.data:
                            s_name = cred.get("sitio", "")
                            if s_name.startswith("WL"):
                                user_dec = desencriptar(cred.get("user_enc", ""))
                                pass_dec = desencriptar(cred.get("pass_enc", ""))
                                if user_dec:
                                    wl_creds[user_dec] = pass_dec
                    
                    st.session_state["wl_creds_dict"] = wl_creds
                    if wl_creds:
                        # Seleccionar el primero por defecto si existe
                        first_u = list(wl_creds.keys())[0]
                        st.session_state["wl_user"] = first_u
                        st.session_state["wl_pw"] = wl_creds[first_u]
                    else:
                        st.session_state["wl_user"] = ""
                        st.session_state["wl_pw"] = ""
                except Exception as e:
                    print(f"Error fetching credentials: {e}")
        else:
            st.session_state["password_correct"] = False

    # Si NO es True (es decir, es False o None), mostramos el login
    if st.session_state.get("password_correct") is not True:
        c_l1, c_l2, c_l3 = st.columns([1.2, 1, 1.2])
        
        with c_l2:
            try:
                st.image("img/WL Hopper Logo - nspc.png", use_container_width=True)
            except:
                st.markdown("<h2 style='text-align: center;'>🚀 WL Hopper</h2>", unsafe_allow_html=True)
            
            st.markdown("<h4 style='text-align: center;'>Acceso al Sistema</h4>", unsafe_allow_html=True)
            
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

    # Detección de móvil
    components.html("""
        <script>
        const isMobile = window.innerWidth < 768;
        window.parent.postMessage({type: 'streamlit:setComponentValue', value: isMobile}, '*');
        </script>
    """, height=0)

    # --- ADMIN DASHBOARD (SIDEBAR) ---
    if st.session_state.get("logged_user") == "fcendra" and supabase:
        with st.sidebar:
            st.markdown("### ⚙️ Panel de Admin")
            try:
                # supabase python no siempre retorna count de manera facil, traemos data
                res_met = supabase.table("metricas").select("minutos_ahorrados").execute()
                total_items = len(res_met.data) if res_met.data else 0
                total_minutos = sum(item.get("minutos_ahorrados", 0) for item in res_met.data) if res_met.data else 0
                horas_ahorradas = total_minutos / 60
                
                st.metric("🤖 Total Ejecuciones", total_items)
                st.metric("⏳ Tiempo Ahorrado", f"{horas_ahorradas:.1f} hs")
            except Exception as e:
                st.error(f"Error cargando métricas: {e}")

    @st.dialog("Acerca de WL Hopper")
    def mostrar_about():
        c1, c2 = st.columns([1, 2])
        with c1:
            try:
                st.image("img/robot_diapos.png")
            except:
                st.write("🤖")
        with c2:
            st.markdown("""
            **WL Hopper** es una app diseñada para optimizar la descarga de certificados desde el sitio de **Worklift**.
            
            Inspirada en una tarea repetitiva que no quería seguirlo siendo, esta herramienta usa bots de navegación para descargar PDFs en segundo plano.
            """)
        
        st.info("🚀 **Misión:** Automatizar y acelerar la tarea de descarga masiva de certificados e informes, y recuperar y estructurar la información de nuestros equipos desde el sitio web de Worklift.")
        
        st.divider()
        st.caption("Desarrollado por Fede García Cendra - 2026")

    with st.sidebar:
        st.markdown("### 🛠️ Opciones")
        st.button("Cerrar Sesión", on_click=lambda: st.session_state.clear(), use_container_width=True)
        if st.button("Acerca del Proyecto", use_container_width=True):
            mostrar_about()
    
    if "log_history" not in st.session_state: st.session_state.log_history = []
    if "proceso_completo" not in st.session_state: st.session_state.proceso_completo = False
    if "html_excel" not in st.session_state: st.session_state.html_excel = ""
    if "df_excel" not in st.session_state: st.session_state.df_excel = None
    if "hay_archivos" not in st.session_state: st.session_state.hay_archivos = False
    if "res_lista" not in st.session_state: st.session_state.res_lista = []
    if "texto_area" not in st.session_state: st.session_state.texto_area = ""
    if "ultimo_archivo_procesado" not in st.session_state: st.session_state.ultimo_archivo_procesado = None

    
    st.markdown('<div class="logo-container">', unsafe_allow_html=True)
    c_l1, c_l2, c_l3 = st.columns([1.5, 1, 1.5])
    with c_l2: 
        try:
            st.image("img/WL Hopper Logo - nspc.png", use_container_width=True)
        except:
            pass
    st.markdown("<h5 style='text-align: center; color: #555; margin-top:-10px; margin-bottom: 25px;'>Automatización de Descarga de Certificados</h5>", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    
    # --- INTERFAZ DE CREDENCIALES (EXPANDER) ---
    with st.expander("🔐 Credenciales de Worklift"):
        st.write("Gestiona tus credenciales de Worklift guardadas de forma cifrada.")
        
        wl_creds_dict = st.session_state.get("wl_creds_dict", {})
        opciones_cred = list(wl_creds_dict.keys()) + ["➕ Nueva Credencial..."]
        
        # Seleccion de credencial
        default_idx = 0
        if st.session_state.get("wl_user") in opciones_cred:
            default_idx = opciones_cred.index(st.session_state["wl_user"])
            
        selected_cred = st.selectbox("Seleccionar credencial guardada:", opciones_cred, index=default_idx)
        
        is_new = (selected_cred == "➕ Nueva Credencial...")
        
        c_cred1, c_cred2 = st.columns(2)
        new_user = c_cred1.text_input("Usuario Worklift", value="" if is_new else selected_cred, key="input_wl_user")
        new_pw = c_cred2.text_input("Contraseña Worklift", value="" if is_new else wl_creds_dict.get(selected_cred, ""), type="password", key="input_wl_pw")
        
        if st.button("💾 Guardar / Actualizar Credencial"):
            if not new_user or not new_pw:
                st.error("Debes ingresar un usuario y contraseña.")
            elif supabase:
                try:
                    enc_u = encriptar(new_user)
                    enc_p = encriptar(new_pw)
                    logged_usr = st.session_state.get("logged_user", "")
                    # Bypass global unique constraint by including user in sitio name
                    sitio_val = f"WL_{logged_usr}_{new_user}"
                    
                    res_check = supabase.table("credenciales_sitios").select("id").eq("sitio", sitio_val).execute()
                    
                    if res_check.data and len(res_check.data) > 0:
                        supabase.table("credenciales_sitios").update({
                            "user_enc": enc_u,
                            "pass_enc": enc_p
                        }).eq("sitio", sitio_val).execute()
                    else:
                        supabase.table("credenciales_sitios").insert({
                            "usuario_app": logged_usr,
                            "sitio": sitio_val,
                            "user_enc": enc_u,
                            "pass_enc": enc_p
                        }).execute()
                    
                    # Update local state
                    wl_creds_dict[new_user] = new_pw
                    st.session_state["wl_creds_dict"] = wl_creds_dict
                    st.session_state["wl_user"] = new_user
                    st.session_state["wl_pw"] = new_pw
                    
                    st.success("Credencial cifrada y guardada correctamente.")
                except Exception as e:
                    st.error(f"Error al guardar credenciales: {e}")
            else:
                st.error("No hay conexión con la base de datos.")
                
        # Update session state with selected credential so the bot uses it
        if not is_new and selected_cred in wl_creds_dict:
            st.session_state["wl_user"] = selected_cred
            st.session_state["wl_pw"] = wl_creds_dict[selected_cred]
    
    col_left, col_right = st.columns([1, 1.2], gap="large")
    
    with col_left:
        
        with st.container(border=True):
            plataforma = st.selectbox("🌐 Plataforma", ["Worklift", "Bureau Veritas (Próximamente)"], help="Selecciona de dónde intentar descargar los certificados.")
            c1, c2 = st.columns(2)
            bajar_cert = c1.checkbox("Descargar Certificados", value=True)
            bajar_inf = c2.checkbox("Descargar Informes", value=False)
            es_semestral = st.checkbox("Vencimiento Semestral (180 días)", help="Calcula una alerta extra a los 6 meses.")
            modo_pruebas = False
            if st.session_state.get("logged_user") == "fcendra":
                modo_pruebas = st.checkbox("🧪 Modo Pruebas (No inyecta métricas)", value=True)
                
        with st.expander("⚙️ Configuración de Salida"):
            nombre_excel = st.text_input("Nombre del Excel", value="Reporte_Hopper")
            prefijo_cert = st.text_input("Prefijo de Certificados", value="Certificado")
    
        st.markdown("##### Listado de Internos")
        archivo_subido = st.file_uploader("Subí tu Excel, TXT, CSV o Foto", type=['txt', 'csv', 'xlsx', 'png', 'jpg', 'jpeg'], help="También podés arrastrar el archivo.")
        
        # --- LÓGICA DE COMPONENTE PORTAPAPELES (LIBRERÍA EXTERNA) ---
        try:
            from streamlit_paste_button import paste_image_button
            paste_result = paste_image_button(
                label="📋 Pegar Imagen del Portapapeles",
                background_color="#008657",
                hover_background_color="#006644",
                key=f"paste_btn_{st.session_state.get('paste_key', 0)}"
            )
            if paste_result and paste_result.image_data is not None:
                if not archivo_subido:
                    import io
                    buf = io.BytesIO()
                    paste_result.image_data.save(buf, format="PNG")
                    buf.name = "pasted_image.png"
                    buf.type = "image/png"
                    archivo_subido = buf
                    st.success("✅ Imagen pegada cargada correctamente.")
        except ImportError:
            pass

        if archivo_subido and archivo_subido.name.lower().endswith(('.png', '.jpg', '.jpeg')):
            st.info("⚠️ **Función Experimental:** La extracción de texto desde imagen (OCR) puede requerir revisión manual.")
            
        # LÓGICA DE EXTRACCIÓN AUTOMÁTICA
        if archivo_subido is not None:
            archivo_id = archivo_subido.name + str(archivo_subido.size)
            if st.session_state.ultimo_archivo_procesado != archivo_id:
                with st.spinner("Procesando archivo..."):
                    texto_extraido = ""
                    if archivo_subido.name.lower().endswith(('.png', '.jpg', '.jpeg')):
                        try:
                            image = Image.open(archivo_subido).convert('L')
                            w, h = image.size
                            image = image.resize((w*2, h*2), Image.Resampling.LANCZOS)
                            enhancer = ImageEnhance.Contrast(image)
                            image = enhancer.enhance(2.0)
                            
                            texto_imagen = pytesseract.image_to_string(image, config='--psm 11')
                            palabras = texto_imagen.upper().split()
                            texto_corregido = []
                            for p in palabras:
                                for char in ['£', '€', 'È', 'É']: p = p.replace(char, 'E')
                                p = re.sub(r'[^A-Z0-9]', '', p)
                                if p and p.startswith('3'): p = 'E' + p[1:]
                                elif p and p[0] in ['4', '^', '@']: p = 'A' + p[1:]
                                if p.startswith('E') or p.startswith('A'):
                                    p_resto = p[1:].replace('O', '0').replace('I', '1').replace('L', '1').replace('S', '5')
                                    texto_corregido.append(p[0] + p_resto)
                                else:
                                    texto_corregido.append(p)
                            texto_extraido = " ".join(texto_corregido)
                        except Exception as e:
                            st.error(f"Error OCR: {e}")
                    else:
                        texto_extraido = extraer_texto_de_archivo(archivo_subido)
                    
                    lista_nuevos = extraer_internos(texto_extraido)
                    if lista_nuevos:
                        nuevo_texto = " ".join(lista_nuevos)
                        if st.session_state.texto_area:
                            st.session_state.texto_area += " " + nuevo_texto
                        else:
                            st.session_state.texto_area = nuevo_texto
                        st.success(f"✅ Se agregaron {len(lista_nuevos)} internos extraídos del archivo.")
                    else:
                        st.warning("No se detectaron internos válidos en el archivo.")
                
                st.session_state.ultimo_archivo_procesado = archivo_id
                st.rerun()
            
        texto_internos = st.text_area("Revisá o pegá los internos acá:", height=115, placeholder="E040230, 3797...", key="texto_area")
        btn_run = st.button("🚀 COMENZAR PROCESO", use_container_width=True)

    
    with col_right:
        terminal_placeholder = st.empty()
        def render_terminal():
            html = f'<div class="terminal-box">'
            html += f'<div style="font-family: \'Consolas\', monospace; font-weight: bold; font-size: 1.1rem; color: #00ff00; margin-bottom: 10px; border-bottom: 1px dashed #555; padding-bottom: 5px;">&gt;<span class="blink-cursor">_</span> REGISTRO DE ACTIVIDAD</div>'
            for entry in st.session_state.log_history:
                # Colores basados en íconos
                color = "#f8f9fa"
                if "✅" in entry or "VIGENTE" in entry: color = "#50fa7b"
                elif "❌" in entry or "ERROR" in entry or "VENCIDO" in entry: color = "#ff5555"
                elif "⚠️" in entry or "⏳" in entry or "PRÓXIMO" in entry: color = "#f1fa8c"
                elif "🤖" in entry: color = "#8be9fd"
                html += f'<div class="log-entry" style="color: {color};">{entry}</div>'
            html += '</div>'
            terminal_placeholder.markdown(html, unsafe_allow_html=True)
        render_terminal()
    
    if btn_run:
        wl_usr = st.session_state.get("wl_user", "")
        wl_pass = st.session_state.get("wl_pw", "")
        
        if not wl_usr or not wl_pass: st.error("Faltan credenciales de Worklift. Selecciona o guarda una.")
        else:
            ruta_temp = "descargas_temp"
            if os.path.exists(ruta_temp): shutil.rmtree(ruta_temp)
            asegurar_carpeta(ruta_temp)
    
            st.session_state.proceso_completo = False
            
            # Formato de log inicial enriquecido
            si_no = lambda b: "Sí" if b else "No"
            st.session_state.log_history = [
                "--- PARÁMETROS ---",
                f"Descargar certificados: {si_no(bajar_cert)}",
                f"Descargar informes de inspección: {si_no(bajar_inf)}",
                f"Modo vencimiento semestral: {si_no(es_semestral)}",
                "------------------",
                "Iniciando conexión con Worklift..."
            ]
            render_terminal()
            
            # Solo se extrae del cuadro de texto final (ya combinado)
            lista = extraer_internos(texto_internos)
            
            if not lista:
                st.session_state.log_history.append("❌ No se encontraron internos para procesar en el cuadro de texto.")
                render_terminal()
            else:
                st.session_state.log_history.append("⏳ Iniciando sesión en Worklift<span class='loading-dots'></span>")
                render_terminal()
                
                bot = WLHopperBot(headless=True)
                if bot.iniciar(wl_usr, wl_pass):
                    st.session_state.log_history[-1] = "🔐 Login exitoso en Worklift."
                    render_terminal()
                    
                    res_lista = []
                    for int_id in lista:
                        st.session_state.log_history.append(f"--- Procesando {int_id} ---")
                        render_terminal()
                        res = bot.procesar_interno(int_id, ruta_temp, bajar_cert, bajar_inf, es_semestral=es_semestral, prefijo_cert=prefijo_cert)
                        res['id'] = int_id
                        
                        # --- CONTINGENCIA BUREAU VERITAS ---
                        if res.get('cert') == 'NO' or res.get('status') in ['VENCIDO', 'RECHAZADO', 'No se pudo encontrar']:
                            st.session_state.log_history.append("⚠️ Certificado inválido/ausente en WL. Iniciando contingencia en Bureau Veritas...")
                            render_terminal()
                            try:
                                from scraper import BureauVeritasBot
                                bv_bot = BureauVeritasBot(headless=True)
                                bv_usr = "SULLAIRNQN"
                                bv_pw = st.secrets.get("bv_password", "TU_PASSWORD_ACA") # Placeholder u obtenido de secrets
                                if bv_bot.iniciar(bv_usr, bv_pw):
                                    bv_res = bv_bot.procesar_interno(int_id, ruta_temp, prefijo_cert=prefijo_cert)
                                    if bv_res.get('descargado'):
                                        st.session_state.log_history.append(f"✅ ¡Certificado encontrado y descargado en Bureau Veritas!")
                                        res['cert'] = "Descargado (BV)"
                                        # Si WL decía vencido pero BV lo bajó, actualizamos el estado para no alarmar
                                        res['status'] = "VIGENTE (BV)"
                                        res['color'] = "VERDE"
                                        res['obs_final'] = "Certificado obtenido de Bureau Veritas."
                                        res['accion_final'] = "-"
                                    else:
                                        st.session_state.log_history.append(f"❌ Tampoco se encontró en Bureau Veritas.")
                                    bv_bot.cerrar()
                                else:
                                    st.session_state.log_history.append("❌ Falló inicio de sesión en Bureau Veritas.")
                            except Exception as e:
                                st.session_state.log_history.append(f"❌ Error en contingencia BV: {e}")
                                
                        res_lista.append(res)
                        for m in res.get('log', []): st.session_state.log_history.append(f"&nbsp;&nbsp;{m}")
                        
                        # Inyección de métricas (si no es prueba y no falló por completo)
                        if not modo_pruebas and res.get('status') != "No se pudo encontrar":
                            registrar_metrica(int_id, "Archivo/Texto")
                            
                        render_terminal()
        
                    bot.cerrar()
                    st.session_state.log_history.append("🔓 Sesión cerrada correctamente.")
                    st.session_state.log_history.append("🏁 PROCESO FINALIZADO.")
                    st.session_state.res_lista = res_lista
                    st.session_state.proceso_completo = True
                    st.session_state.hay_archivos = len(os.listdir(ruta_temp)) > 0 if os.path.exists(ruta_temp) else False
                    st.session_state.paste_key = st.session_state.get('paste_key', 0) + 1 # Resetear portapapeles
        
                    # Generación de HTML
                    html = f"""<style>table {{ border-collapse: collapse; }} td {{ white-space: nowrap; text-align: center; vertical-align: middle; mso-number-format: "\\@"; padding: 5px 15px; }} th {{ white-space: nowrap; text-align: center; vertical-align: middle; padding: 10px 20px; }}</style>
                    <table id="hopperTable" width="100%" border="1" style="font-family: Calibri;">
                    <tr style="background-color: #008657; color: white; font-weight: bold;"><th>INTERNO</th><th>ESTADO</th><th>ÚLTIMA INSPECCIÓN</th><th>VENCIMIENTO<br>ÚLTIMO CERTIFICADO</th><th>CERTIFICADO</th><th>INFORME</th><th>OBSERVACIONES</th><th>ACCIONES</th></tr>"""

                    
                    excel_data = []
                    for r in res_lista:
                        bg, tx, st_text = "#FFFFFF", "#000000", r['status'].upper()
                        cert_val = r['cert']
                        color_code = r.get('color', '').upper()
                        if color_code == "VERDE": bg, tx = "#C6EFCE", "#006100"
                        elif color_code == "AMARILLO": bg, tx = "#FFEB9C", "#9C5700"
                        elif color_code == "ROJO": bg, tx = "#FFC7CE", "#9C0006"
                        else:
                            if "VERDE" in st_text or "VIGENTE" in st_text or "APROBADO" in st_text: bg, tx = "#C6EFCE", "#006100"
                            elif "AMARILLO" in st_text or "PRÓXIMO" in st_text or "GESTIÓN" in st_text or "REINSPECCIONAR" in st_text: bg, tx = "#FFEB9C", "#9C5700"
                            elif "ROJO" in st_text or "VENCIDO" in st_text or "RECHAZADO" in st_text: bg, tx = "#FFC7CE", "#9C0006"
                        
                        html += f'<tr><td>{r["id"]}</td><td style="background-color: {bg}; color: {tx}; font-weight: bold;">{st_text}</td>'
                        html += f'<td>{r.get("insp", "N/A")}</td><td>{r.get("venc", "N/A")}</td>'
                        html += f'<td>{cert_val}</td><td>{r["inf"]}</td>'
                        html += f'<td style="text-align: left; max-width: 250px; white-space: normal;">{r.get("obs_final", "-")}</td>'
                        html += f'<td style="text-align: left; white-space: normal; padding-right: 30px;">{r.get("accion_final", "-")}</td></tr>'
                        
                        row_dict = {
                            "INTERNO": r["id"], "ESTADO": st_text, "ÚLTIMA INSPECCIÓN": r["insp"], "VENCIMIENTO ÚLTIMO CERTIFICADO": r["venc"]
                        }
                            
                        row_dict.update({
                            "CERTIFICADO": cert_val, "INFORME": r["inf"], 
                            "OBSERVACIONES": r.get("obs_final", "-"), "ACCIONES": r.get("accion_final", "-")
                        })
                        excel_data.append(row_dict)
                        
                    html += "</table>"
                    st.session_state.html_excel = html.replace("\n", "")
                    st.session_state.df_excel = pd.DataFrame(excel_data)
                    st.rerun()
                else:
                    st.session_state.log_history.append("❌ ERROR: Credenciales de Worklift incorrectas.")
                    render_terminal()
                    st.error("No se pudo iniciar sesión. Verificá tu usuario y contraseña de Worklift.")
                
    st.divider()
    
    # Preparar el archivo Excel
    excel_buffer = BytesIO()
    if st.session_state.proceso_completo and st.session_state.df_excel is not None:
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            st.session_state.df_excel.to_excel(writer, index=False, sheet_name='Reporte')
            
            # Formato Condicional y Ancho de Columnas
            worksheet = writer.sheets['Reporte']
            from openpyxl.styles import PatternFill, Font, Alignment
            
            # Ajustar ancho de columnas y alinear
            col_letras = [col[0].column_letter for col in worksheet.columns]
            headers_list = [cell.value for cell in worksheet[1]]
            
            for column_cells in worksheet.columns:
                col_letra = column_cells[0].column_letter
                header_val = column_cells[0].value
                
                # Calcular longitud para auto-fit básico
                length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
                ancho_final = min(length + 2, 50)
                
                if header_val == "ESTADO":
                    ancho_final = max(ancho_final, 22) # Más ancho
                elif header_val == "OBSERVACIONES":
                    ancho_final = ancho_final * 0.7 # Reducir 30%
                    if ancho_final < 15: ancho_final = 15
                elif header_val == "INFORME":
                    ancho_final = 9.5
                
                worksheet.column_dimensions[col_letra].width = ancho_final
                
                # Aplicar centrado y wrap a todas las celdas
                for cell in column_cells:
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                
            # Colores
            green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            green_font = Font(color="006100", bold=True)
            yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            yellow_font = Font(color="9C5700", bold=True)
            red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            red_font = Font(color="9C0006", bold=True)
            
            header_fill = PatternFill(start_color="008657", end_color="008657", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            
            if es_semestral:
                worksheet.insert_rows(1)
                worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(worksheet[2]))
                title_cell = worksheet.cell(row=1, column=1)
                title_cell.value = "⚠️ REPORTE DE VENCIMIENTOS SEMESTRALES (180 DÍAS)"
                title_cell.font = Font(bold=True, size=14, color="FFFFFF")
                title_cell.fill = header_fill
                title_cell.alignment = Alignment(horizontal="center", vertical="center")
                header_row = 2
                worksheet.row_dimensions[1].height = 25
            else:
                header_row = 1
            
            for cell in worksheet[header_row]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
            headers = [cell.value for cell in worksheet[header_row]]
            estado_idx = headers.index("ESTADO") + 1 if "ESTADO" in headers else -1
                
            for row in range(header_row + 1, worksheet.max_row + 1):
                res_idx = row - (header_row + 1)
                color_code = ""
                if res_idx >= 0 and res_idx < len(st.session_state.res_lista):
                    color_code = st.session_state.res_lista[res_idx].get('color', '').upper()
                    
                for idx in [estado_idx]:
                    if idx != -1:
                        cell = worksheet.cell(row=row, column=idx)
                        if cell.value:
                            val = str(cell.value).upper()
                            
                            # Si estamos en la columna de estado principal, usamos el color explícito si existe
                            if idx == estado_idx and color_code:
                                if color_code == "VERDE":
                                    cell.fill = green_fill
                                    cell.font = green_font
                                elif color_code == "AMARILLO":
                                    cell.fill = yellow_fill
                                    cell.font = yellow_font
                                elif color_code == "ROJO":
                                    cell.fill = red_fill
                                    cell.font = red_font
                            else:
                                # Fallback o para la columna de estado semestral
                                if "VERDE" in val or "VIGENTE" in val or "APROBADO" in val:
                                    cell.fill = green_fill
                                    cell.font = green_font
                                elif "AMARILLO" in val or "PRÓXIMO" in val or "GESTIÓN" in val or "REINSPECCIONAR" in val:
                                    cell.fill = yellow_fill
                                    cell.font = yellow_font
                                elif "ROJO" in val or "VENCIDO" in val or "RECHAZADO" in val or "ERROR" in val:
                                    cell.fill = red_fill
                                    cell.font = red_font


    excel_data = excel_buffer.getvalue()

    dcol1, dcol2, dcol3 = st.columns(3)
    
    with dcol1:
        if st.session_state.proceso_completo:
            components.html(f"""
                <div id="desktopBtnContainer" style="display: none; margin:0; padding:0; height: 45px; align-items: center;">
                    <button id="cBtn" style="width: 100%; height: 45px; background-color: {VERDE_SULLAIR}; color: white; border: none; border-radius: 4px; font-weight: bold; cursor: pointer; font-family: sans-serif; box-sizing: border-box;">
                        📋 Copiar Tabla Excel
                    </button>
                    <textarea id="hiddenTable" style="position:fixed; top:-1000px; opacity:0;">{st.session_state.html_excel}</textarea>
                </div>
                
                <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
                <div id="mobileBtnContainer" style="display: none; margin:0; padding:0; height: 45px; align-items: center;">
                    <button id="shareBtn" style="width: 100%; height: 45px; background-color: #25D366; color: white; border: none; border-radius: 4px; font-weight: bold; cursor: pointer; font-family: sans-serif; box-sizing: border-box;">
                        📱 Compartir Tabla como Imagen
                    </button>
                    <div id="captureArea" style="position: absolute; left: -9999px; background: white; padding: 10px;">
                        {st.session_state.html_excel}
                    </div>
                </div>

                <script>
                // Detección de dispositivo
                const isMobile = /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent);
                if (isMobile) {{
                    document.getElementById('mobileBtnContainer').style.display = 'flex';
                }} else {{
                    document.getElementById('desktopBtnContainer').style.display = 'flex';
                }}

                // Lógica de copiar (Desktop)
                document.getElementById('cBtn').onclick = function() {{
                    const btn = this;
                    const html = document.getElementById('hiddenTable').value;
                    const blob = new Blob([html], {{ type: 'text/html' }});
                    const data = [new ClipboardItem({{ 'text/html': blob }})];
                    navigator.clipboard.write(data).then(() => {{
                        btn.innerHTML = "✅ ¡COPIADO! (pegar con ctrl+v en Excel)";
                        btn.style.backgroundColor = "#28a745";
                        setTimeout(() => {{ btn.innerHTML = "📋 Copiar Tabla Excel"; btn.style.backgroundColor = "{VERDE_SULLAIR}"; }}, 2000);
                    }});
                }};

                // Lógica de compartir imagen (Móvil)
                document.getElementById('shareBtn').onclick = function() {{
                    const btn = this;
                    const originalText = btn.innerHTML;
                    btn.innerHTML = "⏳ Generando...";
                    
                    html2canvas(document.getElementById('captureArea')).then(canvas => {{
                        canvas.toBlob(blob => {{
                            const file = new File([blob], "reporte.png", {{ type: "image/png" }});
                            if (navigator.canShare && navigator.canShare({{ files: [file] }})) {{
                                navigator.share({{
                                    files: [file],
                                    title: 'Reporte WL Hopper',
                                    text: 'Reporte de Certificados'
                                }}).then(() => {{
                                    btn.innerHTML = "✅ Compartido";
                                    setTimeout(() => btn.innerHTML = originalText, 2000);
                                }}).catch(err => {{
                                    btn.innerHTML = "❌ Error al compartir";
                                    setTimeout(() => btn.innerHTML = originalText, 2000);
                                }});
                            }} else {{
                                const url = URL.createObjectURL(blob);
                                const a = document.createElement('a');
                                a.href = url;
                                a.download = "reporte.png";
                                a.click();
                                URL.revokeObjectURL(url);
                                btn.innerHTML = "✅ Descargado";
                                setTimeout(() => btn.innerHTML = originalText, 2000);
                            }}
                        }});
                    }});
                }};
                </script>
            """, height=45) 
        else:
            st.button("📋 Copiar Tabla / Imagen", disabled=True, use_container_width=True)
            
    with dcol2:
        if st.session_state.proceso_completo:
            st.download_button("📊 Descargar Excel", data=excel_data, file_name="Reporte_WLHopper.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        else:
            st.button("📊 Descargar Excel", disabled=True, use_container_width=True)
    
    with dcol3:
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

    # El bloque de compartir como imagen se movió a dcol1 y se intercala por CSS.

    st.divider()
    st.caption("© 2026 - Desarrollado por Fede García Cendra para Sullair Argentina S.A.")
    st.caption("Consultas a: fcendra@sullair.com.ar")
