import os
import requests
import logging
from playwright.sync_api import sync_playwright
from datetime import datetime
from utils import analizar_fecha

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
        self.browser = self.pw.chromium.launch(headless=self.headless)
        self.context = self.browser.new_context()
        self.page = self.context.new_page()
        try:
            self.page.goto("https://certifica.worklift.com.ar/login/", wait_until="load")
            self.page.locator('input[type="email"]').press_sequentially(usuario, delay=50)
            self.page.locator('input[type="password"]').press_sequentially(clave, delay=50)
            self.page.click('button:has-text("Ingresar")')
            self.page.wait_for_load_state("networkidle")
            return True
        except: return False

    def procesar_interno(self, interno, ruta_base, bajar_certificado, bajar_informe):
        try:
            buscador = self.page.get_by_role("textbox", name="Buscar")
            buscador.fill("")
            buscador.type(interno, delay=30)
            buscador.press("Enter")
            self.page.wait_for_timeout(2000)

            filas = self.page.query_selector_all("table tbody tr")
            candidatos = []
            for fila in filas:
                celdas = fila.query_selector_all("td")
                if len(celdas) < 6: continue
                if interno.upper() in celdas[4].inner_text().strip().upper(): 
                    candidatos.append({
                        "venc": celdas[3].inner_text().strip(),
                        "insp": celdas[2].inner_text().strip(),
                        "menu": fila.query_selector("[id^='__BVID__'][id$='__BV_toggle_']"),
                        "id_inf": celdas[0].inner_text().strip(),
                        "fila_obj": fila
                    })

            if not candidatos:
                return {"status": "No hallado", "venc": "-", "insp": "-", "cert": "NO", "inf": "NO", "det": "Interno no coincide", "log": ["No se hallaron resultados"]}

            candidatos.sort(
                key=lambda x: (
                    datetime.strptime(x["venc"], "%d/%m/%Y"), 
                    datetime.strptime(x["insp"], "%d/%m/%Y")
                ), 
                reverse=True
            )

            # --- LÓGICA DE AUDITORÍA AVANZADA ---
            # 1. El informe a descargar es SIEMPRE el de la primera fila (el más reciente)
            fila_reciente = candidatos[0]
            
            # 2. Buscamos el último certificado que REALMENTE tuvo un PDF
            mejor_cert = None
            for cand in candidatos:
                self.page.keyboard.press("Escape")
                cand["fila_obj"].scroll_into_view_if_needed()
                cand["menu"].click()
                
                tiene_pdf_cert = False
                try:
                    self.page.wait_for_selector(".dropdown-menu.show", timeout=2000)
                    links = self.page.query_selector_all(".dropdown-menu.show a")
                    for link in links:
                        if "CERTIFICADO" in link.inner_text().upper() and cand["id_inf"] in link.get_attribute("href"):
                            tiene_pdf_cert = True
                            break
                except: pass
                self.page.keyboard.press("Escape")

                if tiene_pdf_cert:
                    mejor_cert = cand
                    break

            # Determinamos vencimiento real
            # Si la fila más reciente NO tiene cert, el vencimiento real es el del certificado anterior.
            vencimiento_real = mejor_cert["venc"] if mejor_cert else fila_reciente["venc"]
            estado_f, _, permitir_f = analizar_fecha(vencimiento_real)
            
            # Solo permitimos descarga de certificado si el 'mejor_cert' es el más reciente y está vigente
            permitir_cert = permitir_f and (mejor_cert == fila_reciente)

            log_pasos = [f"Última Inspección: {fila_reciente['insp']}", f"Vencimiento: {vencimiento_real}"]
            existe_cert, existe_inf = (mejor_cert is not None), False
            descargo_cert, descargo_inf = False, False

            # --- PROCESO DE DESCARGA ---
            # Siempre operamos sobre la fila más reciente para el Informe
            fila_reciente["fila_obj"].scroll_into_view_if_needed()
            fila_reciente["menu"].click()
            
            try:
                self.page.wait_for_selector(".dropdown-menu.show", timeout=2000)
                links = self.page.query_selector_all(".dropdown-menu.show a")
                cookies = {c['name']: c['value'] for c in self.context.cookies()}
                
                for link in links:
                    txt = link.inner_text().upper()
                    url = link.get_attribute("href")
                    if not url or fila_reciente["id_inf"] not in url: continue
                    
                    es_cert = "CERTIFICADO" in txt
                    es_inf = "INFORME" in txt
                    
                    if es_inf: existe_inf = True

                    # Descarga de Informe (Siempre que el usuario quiera)[cite: 2]
                    # Descarga de Certificado (Solo si es el vigente real)[cite: 2]
                    if (es_inf and bajar_informe) or (es_cert and bajar_certificado and permitir_cert):
                        full_url = url if url.startswith("http") else f"https://certifica.worklift.com.ar{url if url.startswith('/') else '/' + url}"
                        r = requests.get(full_url, cookies=cookies, headers=self.headers, timeout=20)
                        if r.status_code == 200:
                            tipo_s = "Certificado" if es_cert else "Informe"
                            # Usamos la fecha de la fila de donde proviene el archivo
                            nombre = f"{interno}_{tipo_s}_Vence_{fila_reciente['venc'].replace('/','-')}.pdf"
                            with open(os.path.join(ruta_base, nombre), "wb") as f: f.write(r.content)
                            if es_cert: descargo_cert = True
                            if es_inf: descargo_inf = True
            except: pass
            self.page.keyboard.press("Escape")

            # Logs para consola
            def formatear_log(existe, descargo, permitir):
                if not existe: return "No hallado. No se descargó."
                st = "OK" if permitir else "Vencido"
                ac = "Archivo descargado." if descargo else "No se descargó."
                return f"{st}. {ac}"

            log_pasos.append(f"Certificado: {formatear_log((mejor_cert == fila_reciente), descargo_cert, permitir_cert)}")
            log_pasos.append(f"Informe: {formatear_log(existe_inf, descargo_inf, True)}") # Informe siempre OK si existe[cite: 2]

            # Estado final para Excel (Imagen 2)[cite: 2]
            estado_final = "VENCIDO" if not permitir_cert else estado_f
            
            detalle = f"Certificado {estado_final}"
            if mejor_cert != fila_reciente:
                detalle = f"VENCIDO (No se halló cert. vigente con PDF) - Vencido el {vencimiento_real}"

            return {
                "status": estado_final, "venc": vencimiento_real, "insp": fila_reciente["insp"],
                "cert": "SI" if (mejor_cert == fila_reciente) else "NO", 
                "inf": "SI" if existe_inf else "NO",
                "det": detalle, "log": log_pasos
            }
        except Exception as e:
            return {"status": "Error", "det": str(e)[:20], "log": ["Error crítico"]}
        
    def cerrar(self):
        if self.browser: self.browser.close()
        if self.pw: self.pw.stop()
