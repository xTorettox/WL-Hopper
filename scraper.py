import os
import requests
import logging
from playwright.sync_api import sync_playwright
from datetime import datetime
from utils import analizar_fecha, calcular_vencimiento_semestral
from pdf_utils import analizar_informe_local

class WLHopperBot:
    def __init__(self, headless=True):
        self.headless = headless
        self.pw = None
        self.browser = None
        self.context = None
        self.page = None
        self.headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

    def iniciar(self, usuario, clave):
        self.pw = sync_playwright().start()
        
        # --- LÓGICA DE DETECCIÓN DE NAVEGADOR PARA CLOUD ---
        rutas_posibles = [
            "/usr/bin/chromium",
            "/usr/bin/chromium-browser",
            "/usr/lib/chromium/chromium",
            "/usr/bin/google-chrome"
        ]
        
        self.browser = None
        for ruta in rutas_posibles:
            if os.path.exists(ruta):
                try:
                    self.browser = self.pw.chromium.launch(
                        executable_path=ruta,
                        headless=self.headless,
                        args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"]
                    )
                    break
                except:
                    continue
        
        # Si no encuentra rutas manuales, intenta el lanzamiento estándar
        if not self.browser:
            try:
                self.browser = self.pw.chromium.launch(
                    headless=self.headless,
                    args=["--no-sandbox", "--disable-dev-shm-usage"]
                )
            except Exception as e:
                print(f"Error crítico al lanzar navegador: {e}")
                return False
            
        self.context = self.browser.new_context()
        self.page = self.context.new_page()
        
        try:
            self.page.goto("https://certifica.worklift.com.ar/login/", wait_until="load", timeout=60000)
            self.page.locator('input[type="email"]').fill(usuario)
            self.page.locator('input[type="password"]').fill(clave)
            self.page.click('button:has-text("Ingresar")')
            
            # --- VALIDACIÓN REAL DE LOGIN ---
            # Esperamos a ver el buscador de equipos que solo aparece si el login fue exitoso.
            try:
                # Si en 12 segundos no aparece el buscador, devolvemos False.
                self.page.wait_for_selector('input[placeholder="Buscar"]', timeout=12000)
                return True
            except:
                return False
                
        except Exception as e:
            print(f"Error en Login: {e}")
            return False

    def procesar_interno(self, interno, ruta_base, bajar_certificado, bajar_informe):
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
                    "venc": "-", "insp": "-", "cert": "NO", "inf": "NO", 
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
            log_pasos = [f"✅ Encontrado: {interno}", f"📅 Última Insp: {fila_reciente['insp']}"]
            
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

            # --- PROCESO DE DESCARGA ---
            descargo_cert, descargo_inf = False, False
            existe_inf = False
            
            fila_reciente["menu"].click()
            self.page.wait_for_selector(".dropdown-menu.show", timeout=2000)
            links = self.page.query_selector_all(".dropdown-menu.show a")
            cookies = {c['name']: c['value'] for c in self.context.cookies()}

            ruta_informe_descargado = None

            for link in links:
                txt = link.inner_text().upper()
                url = link.get_attribute("href")
                if not url: continue

                es_cert = "CERTIFICADO" in txt
                es_inf = "INFORME" in txt

                if es_inf: existe_inf = True

                # Siempre descargamos el informe si NO hay certificado en la fila reciente, para analizarlo
                forzar_descarga_informe = existe_inf and not tiene_cert_reciente

                if (es_inf and (bajar_informe or forzar_descarga_informe)) or (es_cert and bajar_certificado and permitir_descarga_cert):
                    full_url = url if url.startswith("http") else f"https://certifica.worklift.com.ar{url if url.startswith('/') else '/' + url}"
                    r = requests.get(full_url, cookies=cookies, headers=self.headers, timeout=25)
                    if r.status_code == 200:
                        tipo = "Certificado" if es_cert else "Informe"
                        nombre = f"{interno}_{tipo}_Vence_{vencimiento_real.replace('/','-')}.pdf"
                        ruta_archivo = os.path.join(ruta_base, nombre)
                        with open(ruta_archivo, "wb") as f:
                            f.write(r.content)
                        if es_cert: descargo_cert = True
                        if es_inf: 
                            descargo_inf = True
                            ruta_informe_descargado = ruta_archivo

            self.page.keyboard.press("Escape")

            log_pasos.append(f"📄 Informe: {'Descargado' if descargo_inf else ('No hallado' if not existe_inf else 'Omitido')}")
            log_pasos.append(f"📜 Certificado: {'Descargado' if descargo_cert else 'Omitido o Vencido'}")

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
                
            # Calcular días de vigencia del último certificado
            dias_restantes = (fecha_venc_cert - datetime.now()).days if mejor_cert else -1
            cert_vigente = dias_restantes >= 0
            txt_dias = f"{dias_restantes} días de vigencia" if cert_vigente else f"vencido en {mejor_cert['venc']}" if mejor_cert else "sin registro"
            
            # Analizar el informe si fue descargado
            estado_ocr = None
            obs_ocr = ""
            if es_rechazado and ruta_informe_descargado:
                log_pasos.append("🤖 Certificado faltante. Analizando informe localmente...")
                estado_ocr, obs_ocr = analizar_informe_local(ruta_informe_descargado)
                log_pasos.append(f"🤖 Resultado OCR: {estado_ocr}")
                
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
                    obs_final = f"Último certificado vencido en {mejor_cert['venc']}." if mejor_cert else "Último certificado vencido."
                    accion_final = "Coordinar recertificación urgente"
                    color_final = "ROJO"
            else:
                # Caso B: INFORME EN ÚLTIMA COLUMNA SIN CERTIFICADO
                if cert_vigente:
                    # B.1) Último certificado está vigente
                    if estado_ocr == "CUMPLE":
                        estado_final = "EN GESTIÓN"
                        obs_final = f"{txt_dias}. Reporte: {obs_ocr}."
                        accion_final = "Esperar carga nuevo certificado"
                        color_final = "VERDE"
                    elif estado_ocr == "NO CUMPLE":
                        estado_final = "REINSPECCIONAR"
                        obs_final = f"Último certificado: {txt_dias}. Reporte: {obs_ocr}."
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
                        obs_final = f"Último certificado vencido. Reporte: {obs_ocr}."
                        accion_final = "Esperar carga nuevo certificado/solicitar provisorio"
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
                        
            return {
                "status": estado_final,
                "insp": fila_reciente["insp"],
                "venc": fila_reciente["venc"],
                "cert": "SI" if descargo_cert else "NO",
                "inf": "SI" if existe_inf else "NO",
                "obs_final": obs_final,
                "accion_final": accion_final,
                "color": color_final,
                "log": log_pasos
            }

        except Exception as e:
            return {"status": "Error", "det": str(e)[:30], "obs": "-", "log": [f"❌ Error: {str(e)[:50]}"]}
        
    def cerrar(self):
        if self.browser: self.browser.close()
        if self.pw: self.pw.stop()
