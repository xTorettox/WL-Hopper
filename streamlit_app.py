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
            **WL Hopper** automatiza la descarga de certificados de Worklift.
            """)
        st.info("🚀 **Misión:** Recuperar y estructurar la información de equipos.")
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
    
    col_left, col_right = st.columns([1, 1.2], gap="large")
    
    with col_left:
        st.markdown('<div class="logo-container">', unsafe_allow_html=True)
        c_l1, c_l2, c_l3 = st.columns([1, 2, 1])
        with c_l2: 
            try:
                st.image("img/WL Hopper Logo - nspc.png", use_container_width=True)
            except:
                pass
        st.markdown("<h5 style='text-align: center; color: #555; margin-top:-10px;'>Automatización de Descarga de Certificados</h5>", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        
        with st.container(border=True):
            user = st.text_input("Usuario (Email)", key="user_email")
            pw = st.text_input("Contraseña", type="password", key="user_pw")
            c1, c2 = st.columns(2)
            bajar_cert = c1.checkbox("Descargar Certificados", value=True)
            bajar_inf = c2.checkbox("Descargar Informes", value=False)
            es_semestral = st.checkbox("Vencimiento Semestral (180 días)", help="Calcula una alerta extra a los 6 meses.")
    
        st.markdown("##### Listado de Internos")
        archivo_subido = st.file_uploader("Subí tu Excel, TXT, CSV o Foto", type=['txt', 'csv', 'xlsx', 'png', 'jpg', 'jpeg'], help="También podés arrastrar el archivo.")
        if archivo_subido and archivo_subido.name.lower().endswith(('.png', '.jpg', '.jpeg')):
            st.info("⚠️ **Función Experimental:** La extracción de texto desde imagen (OCR) puede requerir revisión manual.")
        texto_internos = st.text_area("O pegá el texto acá:", height=115, placeholder="E040230, 3797...")
        btn_run = st.button("🚀 COMENZAR PROCESO", use_container_width=True)
    
    with col_right:
        st.markdown("##### Registro de Actividad")
        terminal_placeholder = st.empty()
        def render_terminal():
            html = f'<div class="terminal-box">'
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
        if not user or not pw: st.error("Faltan credenciales.")
        else:
            ruta_temp = "descargas_temp"
            if os.path.exists(ruta_temp): shutil.rmtree(ruta_temp)
            asegurar_carpeta(ruta_temp)
    
            st.session_state.proceso_completo = False
            st.session_state.log_history = ["Iniciando conexión con Worklift..."]
            render_terminal()
            
            # --- Procesamiento Multimodal ---
            texto_base = texto_internos
            if archivo_subido is not None:
                if archivo_subido.name.lower().endswith(('.png', '.jpg', '.jpeg')):
                    st.session_state.log_history.append("🤖 Analizando imagen con OCR...")
                    render_terminal()
                    try:
                        # Preprocesamiento de imagen para mejorar la precisión del OCR
                        image = Image.open(archivo_subido).convert('L') # Escala de grises
                        w, h = image.size
                        image = image.resize((w*2, h*2), Image.Resampling.LANCZOS) # Aumentar tamaño
                        enhancer = ImageEnhance.Contrast(image)
                        image = enhancer.enhance(2.0) # Aumentar contraste
                        
                        texto_imagen = pytesseract.image_to_string(image, config='--psm 11')
                        
                        # Limpiar errores típicos de OCR SOLO para imágenes
                        palabras = texto_imagen.upper().split()
                        texto_corregido = []
                        for p in palabras:
                            # 1. Corregir prefijo OCR
                            if p and p[0] in ['£', '€', '3', 'È', 'É']:
                                p = 'E' + p[1:]
                            elif p and p[0] in ['4', '^', '@']:
                                p = 'A' + p[1:]
                                
                            # 2. Corregir letras confundidas con números dentro del interno
                            if p.startswith('E') or p.startswith('A'):
                                p_resto = p[1:].replace('O', '0').replace('I', '1').replace('L', '1').replace('S', '5')
                                texto_corregido.append(p[0] + p_resto)
                            else:
                                texto_corregido.append(p)
                                
                        texto_imagen_limpio = " ".join(texto_corregido)
                        texto_base += " " + texto_imagen_limpio
                        st.session_state.log_history.append("🤖 Texto extraído de la imagen exitosamente.")
                        st.session_state.log_history.append(f"🐛 Debug OCR: {texto_imagen_limpio}")
                    except Exception as e:
                        st.session_state.log_history.append(f"❌ Error de OCR: {e}")
                else:
                    st.session_state.log_history.append("📄 Extrayendo texto de archivo...")
                    texto_base += " " + extraer_texto_de_archivo(archivo_subido)
            
            lista = extraer_internos(texto_base)
            
            if not lista:
                st.session_state.log_history.append("❌ No se encontraron internos para procesar.")
                render_terminal()
            else:
                bot = WLHopperBot(headless=True)
                if bot.iniciar(user, pw):
                    st.session_state.log_history.append("🔐 Login exitoso.")
                    render_terminal()
                    
                    res_lista = []
                    for int_id in lista:
                        st.session_state.log_history.append(f"--- Procesando {int_id} ---")
                        render_terminal()
                        res = bot.procesar_interno(int_id, ruta_temp, bajar_cert, bajar_inf)
                        res['id'] = int_id
                        
                        # Vencimiento Semestral
                        if es_semestral:
                            est_sem, fec_sem = calcular_vencimiento_semestral(res["insp"])
                            res["est_sem"] = est_sem
                            res["venc_sem"] = fec_sem
                        else:
                            res["est_sem"] = "-"
                            res["venc_sem"] = "-"
                            
                        res_lista.append(res)
                        for m in res.get('log', []): st.session_state.log_history.append(f"&nbsp;&nbsp;{m}")
                        render_terminal()
        
                    bot.cerrar()
                    st.session_state.log_history.append("🔓 Sesión cerrada correctamente.")
                    st.session_state.log_history.append("🏁 PROCESO FINALIZADO.")
                    st.session_state.proceso_completo = True
                    st.session_state.hay_archivos = len(os.listdir(ruta_temp)) > 0 if os.path.exists(ruta_temp) else False
        
                    # Generación de HTML
                    th_semestral = "<th>VENC. SEMESTRAL</th><th>ESTADO SEMESTRAL</th>" if es_semestral else ""
                    html = f"""<style>table {{ border-collapse: collapse; }} td {{ white-space: nowrap; text-align: center; vertical-align: middle; mso-number-format: "\\@"; padding: 5px 15px; }} th {{ white-space: nowrap; text-align: center; padding: 10px 20px; }}</style>
                    <table id="hopperTable" width="100%" border="1" style="font-family: Calibri;">
                    <tr style="background-color: #008657; color: white; font-weight: bold;"><th>INTERNO</th><th>ESTADO</th><th>ÚLTIMA INSPECCIÓN</th><th>VENCIMIENTO</th>{th_semestral}<th>CERTIFICADO</th><th>INFORME</th><th>OBSERVACIONES</th><th>DETALLE</th></tr>"""
                    
                    excel_data = []
                    for r in res_lista:
                        bg, tx, st_text = "#FFFFFF", "#000000", r['status'].upper()
                        cert_val = r['cert']
                        if "VERDE" in st_text or "VIGENTE" in st_text or "APROBADO" in st_text: bg, tx = "#C6EFCE", "#006100"
                        elif "AMARILLO" in st_text or "PRÓXIMO" in st_text: bg, tx = "#FFEB9C", "#9C5700"
                        elif "ROJO" in st_text or "VENCIDO" in st_text or "RECHAZADO" in st_text: bg, tx = "#FFC7CE", "#9C0006"
                        
                        html += f'<tr><td>{r["id"]}</td><td style="background-color: {bg}; color: {tx}; font-weight: bold;">{st_text}</td>'
                        html += f'<td>{r.get("insp", "N/A")}</td><td>{r.get("venc", "N/A")}</td>'
                        
                        if es_semestral:
                            bg_sem, tx_sem = "#FFFFFF", "#000000"
                            if "VIGENTE" in r["est_sem"]: bg_sem, tx_sem = "#C6EFCE", "#006100"
                            elif "PRÓXIMO" in r["est_sem"]: bg_sem, tx_sem = "#FFEB9C", "#9C5700"
                            elif "VENCIDO" in r["est_sem"]: bg_sem, tx_sem = "#FFC7CE", "#9C0006"
                            html += f'<td>{r["venc_sem"]}</td><td style="background-color: {bg_sem}; color: {tx_sem}; font-weight: bold;">{r["est_sem"]}</td>'
                        
                        html += f'<td>{cert_val}</td><td>{r["inf"]}</td>'
                        html += f'<td style="text-align: left; max-width: 250px; white-space: normal;">{r.get("obs", "-")}</td>'
                        html += f'<td style="text-align: left; white-space: normal; padding-right: 30px;">{r["det"]}</td></tr>'
                        
                        row_dict = {
                            "INTERNO": r["id"], "ESTADO": st_text, "ÚLTIMA INSPECCIÓN": r["insp"], "VENCIMIENTO": r["venc"]
                        }
                        if es_semestral:
                            row_dict["VENC. SEMESTRAL"] = r["venc_sem"]
                            row_dict["ESTADO SEMESTRAL"] = r["est_sem"]
                            
                        row_dict.update({
                            "CERTIFICADO": cert_val, "INFORME": r["inf"], 
                            "OBSERVACIONES": r.get("obs", "-"), "DETALLE": r["det"]
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
            
            # Ajustar ancho de columnas
            for column_cells in worksheet.columns:
                length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
                worksheet.column_dimensions[column_cells[0].column_letter].width = min(length + 2, 50)
                
            # Colores
            green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            green_font = Font(color="006100", bold=True)
            yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            yellow_font = Font(color="9C5700", bold=True)
            red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            red_font = Font(color="9C0006", bold=True)
            
            header_fill = PatternFill(start_color="008657", end_color="008657", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            
            for cell in worksheet[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center")
            
            headers = [cell.value for cell in worksheet[1]]
            estado_idx = headers.index("ESTADO") + 1 if "ESTADO" in headers else -1
            estado_sem_idx = headers.index("ESTADO SEMESTRAL") + 1 if "ESTADO SEMESTRAL" in headers else -1
                
            for row in range(2, worksheet.max_row + 1):
                for idx in [estado_idx, estado_sem_idx]:
                    if idx != -1:
                        cell = worksheet.cell(row=row, column=idx)
                        if cell.value:
                            val = str(cell.value).upper()
                            if "VERDE" in val or "VIGENTE" in val or "APROBADO" in val:
                                cell.fill = green_fill
                                cell.font = green_font
                            elif "AMARILLO" in val or "PRÓXIMO" in val:
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
