import os
import requests
import logging
import base64
from playwright.sync_api import sync_playwright
from datetime import datetime, timedelta

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
        rutas_posibles = ["/usr/bin/chromium", "/usr/bin/chromium-browser", "/usr/lib/chromium/chromium", "/usr/bin/google-chrome"]
        
        self.browser = None
        for ruta in rutas_posibles:
            if os.path.exists(ruta):
                try:
                    self.browser = self.pw.chromium.launch(executable_path=ruta, headless=self.headless, args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"])
                    break
                except: continue
        
        if not self.browser:
            self.browser = self.pw.chromium.launch(headless=self.headless, args=["--no-sandbox", "--disable-dev-shm-usage"])
            
        self.context = self.browser.new_context()
        self.page = self.context.new_page()
        
        try:
            self.page.goto("https://certifica.worklift.com.ar/login/", wait_until="load", timeout=60000)
            self.page.locator('input[type="email"]').fill(usuario)
            self.page.locator('input[type="password"]').fill(clave)
            self.page.click('button:has-text("Ingresar")')
            
            try:
                self.page.wait_for_selector('input[placeholder="Buscar"]', timeout=12000)
                return True
            except: return False
        except Exception as e:
            print(f"Error en Login WL: {e}")
            return False

    def cerrar(self):
        if self.browser: self.browser.close()
        if self.pw: self.pw.stop()

class BureauVeritasBot:
    def __init__(self, headless=True):
        self.headless = headless
        self.pw = None
        self.browser = None
        self.context = None
        self.page = None

    def iniciar(self, usuario, clave, pw_instance=None):
        try:
            self.pw_started_here = False
            if pw_instance: self.pw = pw_instance
            else:
                self.pw = sync_playwright().start()
                self.pw_started_here = True
            
            rutas_posibles = ["/usr/bin/chromium", "/usr/bin/chromium-browser", "/usr/lib/chromium/chromium", "/usr/bin/google-chrome"]
            self.browser = None
            for ruta in rutas_posibles:
                if os.path.exists(ruta):
                    try:
                        self.browser = self.pw.chromium.launch(executable_path=ruta, headless=self.headless, args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu", "--disable-extensions"])
                        break
                    except: continue
            
            if not self.browser:
                self.browser = self.pw.chromium.launch(headless=self.headless, args=["--disable-extensions", "--no-sandbox", "--disable-dev-shm-usage"])
                
            self.context = self.browser.new_context(accept_downloads=True)
            self.page = self.context.new_page()
            
            self.page.goto("https://iip.bureauveritas.com.ar/Login.aspx", wait_until="load", timeout=60000)
            self.page.fill('input[name="txtUsuario"]', usuario)
            self.page.fill('input[name="txtPassword"]', clave)
            self.page.click('input[name="btnAcceso"]')
            return True, ""
        except Exception as e:
            print(f"Error BV Login: {e}")
            return False, str(e)

    def procesar_interno(self, interno, ruta_base, bajar_cert=True, bajar_inf=False, prefijo_cert=""):
        # Retorno de seguridad por si algo falla antes del try
        res = {"status": "Error", "cert": "NO", "descargado": False, "insp": "-", "venc": "-", "observaciones": "Fallo ejecución", "informe": "NO", "archivos_descargados": []}
        try:
            if "BuscaEQ.aspx" not in self.page.url:
                self.page.goto("https://iip.bureauveritas.com.ar/BuscaEQ.aspx")
            
            self.page.check('input#ctl00_ContentPlaceHolder1_RBOpciones_1')
            self.page.fill('input#ctl00_ContentPlaceHolder1_txtNroBuscado', interno)
            self.page.click('input#ctl00_ContentPlaceHolder1_btnBusca')

            # Espera aumentada a 20s para prevenir timeouts en carga lenta
            self.page.wait_for_selector('table#ctl00_ContentPlaceHolder1_gvInformes', timeout=20000)
            fila = self.page.locator(f"tr:has-text('{interno}')").last
            
            if fila.count() > 0:
                celdas = fila.locator("td")
                fecha_insp = celdas.nth(2).inner_text().strip()
                fecha_venc = celdas.nth(3).inner_text().strip()
                res.update({"insp": fecha_insp, "venc": fecha_venc})
                venc_format = fecha_venc.replace('/', '-')
                prefijo = f"{prefijo_cert}_" if prefijo_cert else ""
                archivos_bv = []
                
                boton_pdf = fila.locator('input[id*="ImgbtnCertificado"]')
                if bajar_cert and boton_pdf.count() > 0:
                    with self.page.expect_download(timeout=25000) as download_info:
                        boton_pdf.click()
                    download_info.value.save_as(os.path.join(ruta_base, f"{prefijo}{interno}_Certificado_Vence_{venc_format}.pdf"))
                    archivos_bv.append(os.path.join(ruta_base, f"{prefijo}{interno}_Certificado_Vence_{venc_format}.pdf"))
                    res.update({"cert": "SI", "descargado": True})
                    self.page.goto("https://iip.bureauveritas.com.ar/BuscaEQ.aspx")
                    self.page.fill('input#ctl00_ContentPlaceHolder1_txtNroBuscado', interno)
                    self.page.click('input#ctl00_ContentPlaceHolder1_btnBusca')
                    fila = self.page.locator(f"tr:has-text('{interno}')").last
                    
                boton_informe = fila.locator('input[id*="BtnInforme"]')
                if boton_informe.count() > 0:
                    res["informe"] = "SI"
                    boton_informe.click()
                    self.page.wait_for_selector('textarea#ctl00_ContentPlaceHolder1_txtConclusion', timeout=10000)
                    res["observaciones"] = self.page.locator('textarea#ctl00_ContentPlaceHolder1_txtConclusion').input_value().strip()
                    if bajar_inf:
                        img_pdf = self.page.locator('input#ctl00_ContentPlaceHolder1_imgGeneraPDF')
                        if img_pdf.count() > 0:
                            with self.page.expect_download(timeout=15000) as down_inf:
                                img_pdf.click()
                            down_inf.value.save_as(os.path.join(ruta_base, f"{prefijo}{interno}_Informe_Vence_{venc_format}.pdf"))
                            archivos_bv.append(os.path.join(ruta_base, f"{prefijo}{interno}_Informe_Vence_{venc_format}.pdf"))
                    self.page.goto("https://iip.bureauveritas.com.ar/BuscaEQ.aspx")
                res["status"] = "VIGENTE (BV)" if res["descargado"] else "Encontrado en BV"
                res["archivos_descargados"] = archivos_bv
        except Exception as e:
            res["observaciones"] = f"Error: {str(e)[:30]}"
        return res

    def cerrar(self):
        if self.browser: self.browser.close()
        if getattr(self, "pw_started_here", False) and self.pw: self.pw.stop()
