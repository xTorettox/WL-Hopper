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
            self.page.wait_for_load_state("networkidle")
            return True
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

            # Ordenamos: el más reciente primero
            candidatos.sort(key=lambda x: datetime.strptime(x["venc"], "%d/%m/%Y"), reverse=True)
            fila_reciente = candidatos[0]
            log_pasos = [f"✅ Hallado: {interno}", f"📅 Última Insp: {fila_reciente['insp']}"]
            
            # --- AUDITORÍA DE PDFS ---
            mejor_cert = None
            for cand in candidatos:
                cand["fila_obj"].scroll_into_view_if_needed()
                cand["menu"].click()
                self.page.wait_for_selector(".dropdown-menu.show", timeout=2000)
                links = self.page.query_selector_all(".dropdown-menu.show a")
                tiene_pdf = any("CERTIFICADO" in l.inner_text().upper() for l in links)
                self.page.keyboard.press("Escape")
                
                if tiene_pdf:
                    mejor_cert = cand
                    break

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

            for link in links:
                txt = link.inner_text().upper()
                url = link.get_attribute("href")
                if not url: continue

                es_cert = "CERTIFICADO" in txt
                es_inf = "INFORME" in txt

                if es_inf: existe_inf = True

                if (es_inf and bajar_informe) or (es_cert and bajar_certificado and permitir_descarga_cert):
                    full_url = url if url.startswith("http") else f"https://certifica.worklift.com.ar{url if url.startswith('/') else '/' + url}"
                    r = requests.get(full_url, cookies=cookies, headers=self.headers, timeout=25)
                    if r.status_code == 200:
                        tipo = "Certificado" if es_cert else "Informe"
                        nombre = f"{interno}_{tipo}_Vence_{vencimiento_real.replace('/','-')}.pdf"
                        with open(os.path.join(ruta_base, nombre), "wb") as f:
                            f.write(r.content)
                        if es_cert: descargo_cert = True
                        if es_inf: descargo_inf = True

            self.page.keyboard.press("Escape")

            log_pasos.append(f"📄 Informe: {'Descargado' if descargo_inf else ('No hallado' if not existe_inf else 'Omitido')}")
            log_pasos.append(f"📜 Certificado: {'Descargado' if descargo_cert else 'Omitido o Vencido'}")

            return {
                "status": estado_f if permitir_descarga_cert else "Vencido",
                "venc": vencimiento_real,
                "insp": fila_reciente["insp"],
                "cert": "SI" if descargo_cert else "NO",
                "inf": "SI" if existe_inf else "NO",
                "det": f"Certificado {estado_f}" if permitir_descarga_cert else f"VENCIDO el {vencimiento_real}",
                "log": log_pasos
            }

        except Exception as e:
            return {"status": "Error", "det": str(e)[:30], "log": [f"❌ Error: {str(e)[:50]}"]}
        
    def cerrar(self):
        if self.browser: self.browser.close()
        if self.pw: self.pw.stop()
