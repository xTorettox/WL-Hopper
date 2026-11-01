import os
import requests
import logging
from playwright.sync_api import sync_playwright
from datetime import datetime, timedelta
from utils import analizar_fecha, calcular_vencimiento_semestral
from pdf_utils import analizar_informe_local

# --- PARÁMETROS DE AHORRO DE MEMORIA EXTREMO (STREAMLIT CLOUD 1GB RAM) ---
CHROMIUM_LOW_MEM_ARGS = [
    "--no-sandbox",
    "--disable-dev-shm-usage",
    "--disable-gpu",
    "--disable-extensions",
    "--disable-background-networking",
    "--disable-background-timer-throttling",
    "--disable-backgrounding-occluded-windows",
    "--disable-breakpad",
    "--disable-component-extensions-with-background-pages",
    "--disable-ipc-flooding-protection",
    "--disable-renderer-backgrounding",
    "--metrics-recording-only",
    "--mute-audio",
    "--no-first-run",
    "--no-default-browser-check",
    "--renderer-process-limit=2",
    "--js-flags=--max-old-space-size=128",  # Limita el heap de JS a 128MB
]

def lanzar_navegador_compartido(pw, headless=True):
    """
    Lanza una única instancia global de Chromium optimizada para entornos
    con recursos limitados (Streamlit Cloud ~1GB RAM).
    """
    rutas_posibles = [
        "/usr/bin/chromium",
        "/usr/bin/chromium-browser",
        "/usr/lib/chromium/chromium",
        "/usr/bin/google-chrome"
    ]
    exec_path = None
    for r in rutas_posibles:
        if os.path.exists(r):
            exec_path = r
            break

    launch_kwargs = {
        "headless": headless,
        "args": CHROMIUM_LOW_MEM_ARGS
    }
    if exec_path:
        launch_kwargs["executable_path"] = exec_path

    return pw.chromium.launch(**launch_kwargs)

class WLHopperBot:
    def __init__(self, headless=True):
        self.headless = headless
        self.pw = None
        self.browser = None
        self.context = None
        self.page = None
        self.headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
        self.pw_started_here = False
        self.browser_started_here = False

    def iniciar(self, usuario, clave, pw_instance=None, browser_instance=None):
        usuario = (usuario or "").strip()
        clave = (clave or "").strip()
        if not usuario or not clave:
            return False, "Usuario o contraseña de Worklift vacíos."

        self.pw_started_here = False
        if pw_instance:
            self.pw = pw_instance
        else:
            self.pw = sync_playwright().start()
            self.pw_started_here = True

        self.browser_started_here = False
        if browser_instance:
            self.browser = browser_instance
        else:
            try:
                self.browser = lanzar_navegador_compartido(self.pw, headless=self.headless)
                self.browser_started_here = True
            except Exception as e:
                msg = f"Error crítico al lanzar navegador: {e}"
                print(msg)
                return False, msg
            
        try:
            self.context = self.browser.new_context()
            self.page = self.context.new_page()
            
            self.page.goto("https://certifica.worklift.com.ar/login/", wait_until="domcontentloaded", timeout=60000)
            self.page.wait_for_selector('input[type="email"]', timeout=20000)
            self.page.locator('input[type="email"]').fill(usuario)
            self.page.locator('input[type="password"]').fill(clave)
            self.page.click('button:has-text("Ingresar")')
            
            # --- VALIDACIÓN ROBUSTA DE LOGIN ---
            # Esperamos activamente hasta 25 segundos (útil en Streamlit Cloud con latencia alta)
            for _ in range(25):
                self.page.wait_for_timeout(1000)
                
                # 1. ¿Apareció el buscador del panel de equipos?
                if self.page.locator('input[placeholder*="Buscar"]').count() > 0:
                    return True, ""
                
                # 2. ¿Apareció una alerta de credenciales incorrectas en el sitio?
                alertas = self.page.locator('.alert, [role="alert"]').all_inner_texts()
                for a in alertas:
                    a_limpia = a.strip()
                    if a_limpia and ("no puede iniciar" in a_limpia.lower() or "credenciales" in a_limpia.lower() or "incorrect" in a_limpia.lower()):
                        return False, f"Credenciales rechazadas por Worklift: '{a_limpia}'"
                        
                # 3. ¿El URL navegó fuera de la ruta de login?
                if "/login" not in self.page.url:
                    try:
                        self.page.wait_for_selector('input[placeholder*="Buscar"]', timeout=8000)
                        return True, ""
                    except Exception:
                        return True, ""
                        
            # Si pasaron los 25s y seguimos en la URL de login
            alertas = self.page.locator('.alert, [role="alert"]').all_inner_texts()
            if alertas:
                return False, f"Error en Worklift: {'; '.join(a.strip() for a in alertas if a.strip())}"
                
            if "/login" in self.page.url:
                return False, "Timeout esperando respuesta de Worklift (la página no avanzó desde /login/)."
                
            return True, ""
                
        except Exception as e:
            msg = f"Error en conexión con Worklift: {e}"
            print(msg)
            return False, msg

    def procesar_interno(self, interno, ruta_base, bajar_certificado, bajar_informe, es_semestral=False, prefijo_cert="Certificado"):
        try:
            # Limpieza y búsqueda del interno
            buscador = self.page.get_by_role("textbox", name="Buscar")
            buscador.click()
            self.page.keyboard.press("Control+A")
            self.page.keyboard.press("Backspace")
            buscador.type(interno, delay=50)
            buscador.press("Enter")
            self.page.wait_for_timeout(3000)

            filas = self.page.query_selector_all("table tbody tr")
            candidatos = []
            
            for fila in filas:
                celdas = fila.query_selector_all("td")
                if len(celdas) < 6: continue
                # Validación estricta: el interno debe estar en la celda correspondiente
                if interno.upper() in celdas[4].inner_text().strip().upper(): 
                    candidatos.append({
                        "venc": celdas[3].inner_text().strip(),
                        "insp": celdas[2].inner_text().strip(),
                        "menu": fila.query_selector("[id^='__BVID__'][id$='__BV_toggle_']"),
                        "id_inf": celdas[0].inner_text().strip(),
                        "fila_obj": fila
                    })

            if not candidatos:
                return {
                    "status": "No se pudo encontrar", 
                    "venc": "-", "venc_real": "-", "insp": "-", "cert": "NO", "inf": "NO", 
                    "det": "No se halló el interno en Worklift", 
                    "log": ["❌ Interno no encontrado en la tabla."]
                }

            # Ordenamos: el más reciente primero, basándonos en la fecha de inspección
            def parse_date(d_str):
                try:
                    return datetime.strptime(d_str, "%d/%m/%Y")
                except:
                    return datetime.min
            
            candidatos.sort(key=lambda x: parse_date(x["insp"]), reverse=True)
            fila_reciente = candidatos[0]
            
            # --- AUDITORÍA DE PDFS ---
            mejor_cert = None
            mejor_inf = None
            
            # Buscamos en la fila más reciente (la primera tras ordenar)
            fila_reciente["fila_obj"].scroll_into_view_if_needed()
            fila_reciente["menu"].click()
            self.page.wait_for_selector(".dropdown-menu.show", timeout=2000)
            links = self.page.query_selector_all(".dropdown-menu.show a")
            
            tiene_cert_reciente = any("CERTIFICADO" in l.inner_text().upper() for l in links)
            tiene_inf_reciente = any("INFORME" in l.inner_text().upper() for l in links)
            self.page.keyboard.press("Escape")

            if tiene_cert_reciente:
                mejor_cert = fila_reciente
            else:
                # Si la fila más reciente NO tiene certificado, revisamos las anteriores por si hay alguno válido,
                # pero mantenemos la fila más reciente para los reportes
                for cand in candidatos[1:]:
                    cand["fila_obj"].scroll_into_view_if_needed()
                    cand["menu"].click()
                    self.page.wait_for_selector(".dropdown-menu.show", timeout=2000)
                    links_cand = self.page.query_selector_all(".dropdown-menu.show a")
                    if any("CERTIFICADO" in l.inner_text().upper() for l in links_cand):
                        mejor_cert = cand
                        self.page.keyboard.press("Escape")
                        break
                    self.page.keyboard.press("Escape")

            vencimiento_real = mejor_cert["venc"] if mejor_cert else fila_reciente["venc"]
            estado_f, _, permitir_f = analizar_fecha(vencimiento_real)
            
            # Permitimos descarga solo si el más reciente es el que tiene el PDF
            permitir_descarga_cert = permitir_f and (mejor_cert == fila_reciente)

            hubo_descarga_forzada = False
            # --- PROCESO DE DESCARGA ---
            descargo_cert, descargo_inf = False, False
            existe_inf = False
            
            fila_reciente["menu"].click()
            self.page.wait_for_selector(".dropdown-menu.show", timeout=2000)
            links = self.page.query_selector_all(".dropdown-menu.show a")
            cookies = {c['name']: c['value'] for c in self.context.cookies()}

            ruta_informe_descargado = None
            ruta_cert_descargado = None
            archivos_wl = []

            for link in links:
                txt = link.inner_text().upper()
                url = link.get_attribute("href")
                if not url: continue

                es_cert = "CERTIFICADO" in txt
                es_inf = "INFORME" in txt

                if es_inf: existe_inf = True

                # Siempre descargamos el informe si NO hay certificado en la fila reciente, para analizarlo
                forzar_descarga_informe = existe_inf and not tiene_cert_reciente
                if forzar_descarga_informe: hubo_descarga_forzada = True

                if (es_inf and (bajar_informe or forzar_descarga_informe)) or (es_cert and bajar_certificado and permitir_descarga_cert):
                    full_url = url if url.startswith("http") else f"https://certifica.worklift.com.ar{url if url.startswith('/') else '/' + url}"
                    r = requests.get(full_url, cookies=cookies, headers=self.headers, timeout=25)
                    if r.status_code == 200:
                        tipo = "Certificado" if es_cert else "Informe"
                        nombre_base = f"{interno}_{tipo}_Vence_{vencimiento_real.replace('/','-')}.pdf"
                        prefijo = f"{prefijo_cert}_" if prefijo_cert else ""
                        nombre = f"{prefijo}{nombre_base}"
                        
                        ruta_archivo = os.path.join(ruta_base, nombre)
                        with open(ruta_archivo, "wb") as f:
                            f.write(r.content)
                            
                        archivos_wl.append(ruta_archivo)
                        
                        if es_cert: 
                            descargo_cert = True
                            ruta_cert_descargado = ruta_archivo
                        if es_inf: 
                            descargo_inf = True
                            ruta_informe_descargado = ruta_archivo

            self.page.keyboard.press("Escape")

            # --- ANÁLISIS PDF (AUDITORÍA DE SEGURIDAD) ---
            observaciones = "-"
            estado_final = "DESCONOCIDO"
            det_final = "-"
            
            def _parse_date(d_str):
                try: return datetime.strptime(d_str, "%d/%m/%Y")
                except: return datetime.min
                
            fecha_insp_reciente = _parse_date(fila_reciente["insp"])
            fecha_insp_cert = _parse_date(mejor_cert["insp"]) if mejor_cert else datetime.min
            fecha_venc_cert = _parse_date(mejor_cert["venc"]) if mejor_cert else datetime.min
            
            es_rechazado = False
            if fecha_insp_reciente > fecha_insp_cert and not tiene_cert_reciente:
                es_rechazado = True
                
            # Analizar el informe si fue descargado
            estado_ocr = None
            obs_ocr = ""
            if es_rechazado and ruta_informe_descargado:
                estado_ocr, obs_ocr = analizar_informe_local(ruta_informe_descargado)
                
            # --- MODIFICACIÓN LÓGICA SEMESTRAL Y CORRECCIÓN DE FECHAS ---
            # Si la inspección falló, Worklift carga una fecha falsa +1 año.
            # Lo corregimos usando la fecha del último certificado real.
            if es_rechazado and estado_ocr != "CUMPLE":
                fecha_venc_base = fecha_venc_cert
                venc_mostrar = mejor_cert["venc"] if mejor_cert else "-"
            else:
                fecha_venc_base = _parse_date(fila_reciente["venc"])
                venc_mostrar = fila_reciente["venc"]
            
            venc_real = venc_mostrar
                
            if es_semestral:
                # La vigencia es 180 días desde la inspección válida
                if not es_rechazado or (es_rechazado and estado_ocr == "CUMPLE"):
                    base_insp = fecha_insp_reciente
                else:
                    base_insp = fecha_insp_cert
                    
                if base_insp != datetime.min:
                    fecha_venc_base = base_insp + timedelta(days=180)
                    venc_mostrar = fecha_venc_base.strftime("%d/%m/%Y")
                else:
                    fecha_venc_base = datetime.min
                    venc_mostrar = "-"
                    
            # Recalcular días de vigencia con la fecha corregida
            if fecha_venc_base != datetime.min:
                dias_restantes = (fecha_venc_base - datetime.now()).days
                cert_vigente = dias_restantes >= 0
                txt_dias = f"{dias_restantes} días de vigencia" if cert_vigente else f"vencido en {venc_mostrar}"
            else:
                dias_restantes = -1
                cert_vigente = False
                txt_dias = "sin registro"
                
            if not es_rechazado:
                # Caso A: INFORME Y CERTIFICADO EN ÚLTIMA COLUMNA
                if dias_restantes > 30:
                    estado_final = "VIGENTE"
                    obs_final = f"{txt_dias}"
                    accion_final = "-"
                    color_final = "VERDE"
                elif 0 <= dias_restantes <= 30:
                    estado_final = "PRÓXIMO A VENCER"
                    obs_final = f"{txt_dias}"
                    accion_final = "Coordinar recertificación"
                    color_final = "AMARILLO"
                else:
                    estado_final = "VENCIDO"
                    obs_final = f"Último certificado vencido en {venc_mostrar}." if venc_mostrar != "-" else "Último certificado vencido."
                    accion_final = "Coordinar recertificación urgente"
                    color_final = "ROJO"
            else:
                # Caso B: INFORME EN ÚLTIMA COLUMNA SIN CERTIFICADO
                if cert_vigente:
                    # B.1) Último certificado está vigente
                    if estado_ocr == "CUMPLE":
                        estado_final = "EN GESTIÓN"
                        obs_final = f"{txt_dias}. Inspección superada."
                        accion_final = "Esperar carga nuevo certificado"
                        color_final = "VERDE"
                    elif estado_ocr == "NO CUMPLE":
                        estado_final = "REINSPECCIONAR"
                        obs_final = f"Último certificado: {txt_dias}. Interno NO superó la inspección."
                        accion_final = "Verificar informe, contactar a ST"
                        color_final = "AMARILLO"
                    else:
                        estado_final = "EN GESTIÓN"
                        obs_final = f"Último certificado: {txt_dias}. No se pudo leer el informe automáticamente."
                        accion_final = "Verificar informe para comprobar resultado"
                        color_final = "AMARILLO"
                else:
                    # B.2) Último certificado está vencido o no hay registro
                    if estado_ocr == "CUMPLE":
                        estado_final = "VIGENTE"
                        obs_final = f"Último certificado vencido. Inspección superada."
                        accion_final = "Esperar nuevo certificado/solicitar provisorio"
                        color_final = "VERDE" # Si es VIGENTE, lógicamente es VERDE
                    elif estado_ocr == "NO CUMPLE":
                        estado_final = "VENCIDO"
                        obs_final = f"Último certificado vencido. Interno NO superó la inspección."
                        accion_final = "Verificar informe, contactar a ST"
                        color_final = "ROJO"
                    else:
                        estado_final = "VENCIDO"
                        obs_final = f"Último certificado vencido. No se pudo leer el informe automáticamente."
                        accion_final = "Verificar informe"
                        color_final = "ROJO"
                        
            # --- CONSTRUCCIÓN DEL REGISTRO DE ACTIVIDAD (LOG) ---
            log_pasos = []
            log_pasos.append("✅ Interno encontrado")
            log_pasos.append(f"📄 Último Informe de Inspección: {fila_reciente['insp']}")
            
            estado_cert_str = ""
            if tiene_cert_reciente:
                estado_cert_str = "Vigente (descargado)" if descargo_cert else "Vigente (no descargado)" if permitir_descarga_cert else "Vencido"
            else:
                if mejor_cert:
                    estado_cert_str = f"Vencido (anterior, no descargado)"
                else:
                    estado_cert_str = "No hallado"
                    
            if descargo_cert and not tiene_cert_reciente:
                # Caso atípico, descargado desde uno anterior
                estado_cert_str = "Vigente (anterior, descargado)"
                
            log_pasos.append(f"📜 Último Certificado: {estado_cert_str}")
            
            if fecha_venc_base != datetime.min:
                log_pasos.append(f"📅 Fecha vencimiento certificado: {venc_mostrar} ({txt_dias})")
                
            if hubo_descarga_forzada:
                log_pasos.append("⚠️ Certificado faltante o vencido. Se forzó la descarga del informe de inspección.")
                
            if es_rechazado and ruta_informe_descargado:
                log_pasos.append("🤖 Certificado nuevo no encontrado. Analizando Informe de Inspección mediante OCR...")
                if estado_ocr == "NO CUMPLE":
                    log_pasos.append(f"🤖 Reporte: interno {interno} NO PASÓ la inspección")
                elif estado_ocr == "CUMPLE":
                    log_pasos.append(f"🤖 Reporte: interno {interno} PASÓ la inspección")
                else:
                    log_pasos.append(f"🤖 Reporte: no se pudo leer el resultado del informe")
                    
            if accion_final != "-":
                log_pasos.append(f"💡 Sugerencia: {accion_final}")
                        
            return {
                "status": estado_final,
                "insp": fila_reciente["insp"],
                "venc": venc_mostrar,
                "venc_real": venc_real,
                "cert": "SI" if descargo_cert else "NO",
                "inf": "SI" if existe_inf else "NO",
                "obs_final": obs_final,
                "accion_final": accion_final,
                "color": color_final,
                "log": log_pasos,
                "archivos_descargados": archivos_wl
            }

        except Exception as e:
            return {"status": "Error", "insp": "-", "venc": "-", "venc_real": "-", "cert": "NO", "inf": "NO", "obs_final": str(e)[:30], "accion_final": "-", "color": "ROJO", "log": [f"❌ Error: {str(e)[:50]}"], "archivos_descargados": []}
        
    def cerrar(self):
        try:
            if self.page:
                self.page.close()
                self.page = None
        except Exception: pass
        try:
            if self.context:
                self.context.close()
                self.context = None
        except Exception: pass
        if getattr(self, "browser_started_here", False) and self.browser:
            try:
                self.browser.close()
                self.browser = None
            except Exception: pass
        if getattr(self, "pw_started_here", False) and self.pw:
            try:
                self.pw.stop()
                self.pw = None
            except Exception: pass

class BureauVeritasBot:
    def __init__(self, headless=True):
        self.headless = headless
        self.pw = None
        self.browser = None
        self.context = None
        self.page = None
        self.pw_started_here = False
        self.browser_started_here = False

    def iniciar(self, usuario, clave, pw_instance=None, browser_instance=None):
        try:
            self.pw_started_here = False
            if pw_instance:
                self.pw = pw_instance
            else:
                self.pw = sync_playwright().start()
                self.pw_started_here = True

            self.browser_started_here = False
            if browser_instance:
                self.browser = browser_instance
            else:
                self.browser = lanzar_navegador_compartido(self.pw, headless=self.headless)
                self.browser_started_here = True

            self.context = self.browser.new_context(
                accept_downloads=True,
                extra_http_headers={"Content-Disposition": "attachment"}
            )
            self.page = self.context.new_page()
            
            self.page.goto("https://iip.bureauveritas.com.ar/Login.aspx", wait_until="load", timeout=60000)
            self.page.fill('input[name="txtUsuario"]', usuario)
            self.page.fill('input[name="txtPassword"]', clave)
            self.page.click('input[name="btnAcceso"]')

            self.page.click('a.ctl00_menua_1[href="Busquedas.aspx"]')
            self.page.click('a#ctl00_ContentPlaceHolder1_lkBuscaEQ')
            return True, ""
        except Exception as e:
            print(f"Error BV Login: {e}")
            return False, str(e)

    def procesar_interno(self, interno, ruta_base, bajar_cert=True, bajar_inf=False, prefijo_cert=""):
        res = {"status": "No encontrado en BV", "cert": "NO", "descargado": False, "insp": "-", "venc": "-", "observaciones": "", "informe": "NO", "log": [], "archivos_descargados": []}
        log_pasos = []
        try:
            if not self.page or self.page.is_closed():
                log_pasos.append("⚠️ Página de BV no disponible.")
                res["log"] = log_pasos
                return res

            if "BuscaEQ.aspx" not in self.page.url:
                self.page.goto("https://iip.bureauveritas.com.ar/BuscaEQ.aspx", timeout=20000)

            log_pasos.append(f"🔍 Iniciando búsqueda del interno '{interno}' en Bureau Veritas...")
            self.page.check('input#ctl00_ContentPlaceHolder1_RBOpciones_1')
            self.page.fill('input#ctl00_ContentPlaceHolder1_txtNroBuscado', interno)
            self.page.click('input#ctl00_ContentPlaceHolder1_btnBusca')

            # Esperamos a que la tabla se dibuje si hay registros
            hay_tabla = False
            try:
                self.page.wait_for_selector('table#ctl00_ContentPlaceHolder1_gvInformes', timeout=4000)
                hay_tabla = True
            except Exception:
                pass
            
            if not hay_tabla:
                log_pasos.append(f"❌ No se hallaron registros en Bureau Veritas para el interno '{interno}'.")
                res["status"] = "No encontrado en BV"
                res["log"] = log_pasos
                return res

            # Recuperamos todas las filas
            rows = self.page.locator("table#ctl00_ContentPlaceHolder1_gvInformes tr")
            row_count = rows.count()
            
            # Filtramos únicamente las filas que contienen celdas de datos (td)
            candidatas = []
            for i in range(row_count):
                r = rows.nth(i)
                if r.locator("td").count() > 0:
                    candidatas.append(r)
                    
            log_pasos.append(f"📊 Se recuperaron {len(candidatas)} registros en la tabla para el interno '{interno}'.")
            
            # Listamos las filas en los logs para visibilidad completa del usuario
            for idx, r in enumerate(candidatas):
                celdas = r.locator("td")
                txt_suministro = celdas.nth(1).inner_text().strip() if celdas.count() > 1 else "-"
                txt_insp = celdas.nth(2).inner_text().strip() if celdas.count() > 2 else "-"
                txt_venc = celdas.nth(3).inner_text().strip() if celdas.count() > 3 else "-"
                log_pasos.append(f"   👉 Fila #{idx + 1}: Suministro='{txt_suministro}' | Inspección={txt_insp} | Vencimiento={txt_venc}")
            
            # Buscamos si alguna fila tiene el código de interno en su texto visible
            matching_candidatas = []
            for r in candidatas:
                text = r.inner_text().upper()
                if interno.upper() in text:
                    matching_candidatas.append(r)
                    
            fila = None
            if matching_candidatas:
                fila = matching_candidatas[-1]
                log_pasos.append(f"🎯 Fila coincidente encontrada para '{interno}' (seleccionando la más reciente).")
            elif candidatas:
                fila = candidatas[-1]
                log_pasos.append(f"⚠️ El código de interno '{interno}' no figura textualmente en las celdas.")
                log_pasos.append(f"👉 Aplicando fallback automático: Seleccionando la última fila del equipo buscado.")
                
            if fila:
                celdas = fila.locator("td")
                fecha_insp = celdas.nth(2).inner_text().strip() if celdas.count() > 3 else "-"
                fecha_venc = celdas.nth(3).inner_text().strip() if celdas.count() > 3 else "-"
                
                res["insp"] = fecha_insp
                res["venc"] = fecha_venc
                
                venc_format = fecha_venc.replace('/', '-') if fecha_venc != "-" else "SinFecha"
                prefijo = f"{prefijo_cert}_" if prefijo_cert else ""
                archivos_bv = []
                
                # --- PASO 1: DETALLE DEL INFORME (PRIMERO) ---
                boton_informe = fila.locator('input[id*="BtnInforme" i], input[name*="BtnInforme" i], input[type="submit"]') if fila else None
                if boton_informe and boton_informe.count() > 0:
                    res["informe"] = "SI"
                    log_pasos.append("📂 Accediendo a los detalles del informe...")
                    try:
                        boton_informe.first.click()
                        self.page.wait_for_selector('textarea#ctl00_ContentPlaceHolder1_txtConclusion', timeout=10000)
                        
                        obs_text = self.page.locator('textarea#ctl00_ContentPlaceHolder1_txtConclusion').input_value()
                        res["observaciones"] = obs_text.strip()
                        log_pasos.append(f"💬 Conclusión del informe en BV: '{res['observaciones']}'")
                        
                        if bajar_inf:
                            img_pdf = self.page.locator('input#ctl00_ContentPlaceHolder1_imgGeneraPDF')
                            if img_pdf.count() > 0:
                                log_pasos.append("📄 Iniciando descarga del archivo del informe PDF...")
                                try:
                                    with self.page.expect_download(timeout=15000) as download_info:
                                        img_pdf.click()
                                    
                                    download = download_info.value
                                    file_path_inf = os.path.join(ruta_base, f"{prefijo}{interno}_Informe_Vence_{venc_format}.pdf")
                                    download.save_as(file_path_inf)
                                    
                                    archivos_bv.append(file_path_inf)
                                    res["informe"] = "SI"
                                    log_pasos.append(f"💾 Informe PDF guardado en: {file_path_inf}")
                                except Exception as e_inf:
                                    log_pasos.append(f"❌ Error descargando PDF del informe: {e_inf}")
                            else:
                                log_pasos.append("⚠️ El botón de generación de PDF de informe no está presente.")
                    except Exception as e_detail:
                        log_pasos.append(f"❌ Error accediendo al detalle del informe: {e_detail}")
                        
                # --- PASO 2: DESCARGAR EL CERTIFICADO ---
                if bajar_cert:
                    # Volvemos a buscar el interno para recargar la grilla limpia y descargar el certificado
                    log_pasos.append("🔄 Volviendo a buscar el equipo para descargar el certificado...")
                    try:
                        self.page.goto("https://iip.bureauveritas.com.ar/BuscaEQ.aspx", timeout=20000)
                        self.page.check('input#ctl00_ContentPlaceHolder1_RBOpciones_1')
                        self.page.fill('input#ctl00_ContentPlaceHolder1_txtNroBuscado', interno)
                        self.page.click('input#ctl00_ContentPlaceHolder1_btnBusca')
                        self.page.wait_for_selector('table#ctl00_ContentPlaceHolder1_gvInformes', timeout=8000)
                        
                        # Re-localizar la fila
                        rows_rec = self.page.locator("table#ctl00_ContentPlaceHolder1_gvInformes tr")
                        candidatas_rec = []
                        for i in range(rows_rec.count()):
                            r_rec = rows_rec.nth(i)
                            if r_rec.locator("td").count() > 0:
                                candidatas_rec.append(r_rec)
                                
                        matching_candidatas_rec = []
                        for r_rec in candidatas_rec:
                            text_rec = r_rec.inner_text().upper()
                            if interno.upper() in text_rec:
                                matching_candidatas_rec.append(r_rec)
                                
                        fila_rec = None
                        if matching_candidatas_rec:
                            fila_rec = matching_candidatas_rec[-1]
                        elif candidatas_rec:
                            fila_rec = candidatas_rec[-1]
                            
                        if fila_rec:
                            boton_pdf = fila_rec.locator('input[id*="ImgbtnCertificado"]')
                            if boton_pdf.count() > 0:
                                log_pasos.append("📜 Botón de Certificado PDF disponible. Descargando...")
                                try:
                                    with self.page.expect_response(lambda r: "Prepara_PDF.aspx" in r.url, timeout=20000) as response_info:
                                        boton_pdf.click()
                                    
                                    response = response_info.value
                                    pdf_url = response.url
                                    
                                    cookies = {c['name']: c['value'] for c in self.context.cookies()}
                                    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
                                    
                                    r = requests.get(pdf_url, cookies=cookies, headers=headers, timeout=25)
                                    if r.status_code == 200:
                                        file_path = os.path.join(ruta_base, f"{prefijo}{interno}_Certificado_Vence_{venc_format}.pdf")
                                        with open(file_path, "wb") as f:
                                            f.write(r.content)
                                        
                                        archivos_bv.append(file_path)
                                        res["cert"] = "SI"
                                        res["descargado"] = True
                                        log_pasos.append(f"💾 Certificado PDF descargado con éxito y guardado en: {file_path}")
                                    else:
                                        log_pasos.append(f"❌ Error al descargar el PDF de red (HTTP {r.status_code})")
                                except Exception as e_pdf:
                                    log_pasos.append(f"❌ Error al capturar certificado PDF: {e_pdf}")
                            else:
                                log_pasos.append("⚠️ El certificado no está disponible para este equipo.")
                        else:
                            log_pasos.append("❌ No se pudo re-localizar la fila para descargar el certificado.")
                    except Exception as e_rec:
                        log_pasos.append(f"❌ Error al re-buscar para descargar certificado: {e_rec}")
                            
                res["status"] = "VIGENTE (BV)" if res["descargado"] else "Encontrado en BV"
            else:
                log_pasos.append("❌ No se encontraron registros de datos en la tabla de BV.")
        except Exception as e:
            log_pasos.append(f"❌ Error en búsqueda de BV: {e}")
        finally:
            try:
                # Volvemos a la sección de búsqueda limpia para el siguiente interno
                if self.page:
                    self.page.goto("https://iip.bureauveritas.com.ar/BuscaEQ.aspx", timeout=20000)
            except Exception:
                pass
            
        res["archivos_descargados"] = archivos_bv if 'archivos_bv' in locals() else []
        res["log"] = log_pasos
        return res

    def cerrar(self):
        try:
            if self.page:
                self.page.close()
                self.page = None
        except Exception: pass
        try:
            if self.context:
                self.context.close()
                self.context = None
        except Exception: pass
        if getattr(self, "browser_started_here", False) and self.browser:
            try:
                self.browser.close()
                self.browser = None
            except Exception: pass
        if getattr(self, "pw_started_here", False) and self.pw:
            try:
                self.pw.stop()
                self.pw = None
            except Exception: pass
