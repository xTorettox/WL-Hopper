import os
import re
import base64
from playwright.sync_api import sync_playwright
from utils import analizar_fecha


class BureauVeritasBot:
    def __init__(self, headless=True):
        self.headless = headless
        self.pw = None
        self.browser = None
        self.context = None
        self.page = None

    def iniciar(self, usuario, clave):
        self.pw = sync_playwright().start()

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

        if not self.browser:
            try:
                self.browser = self.pw.chromium.launch(
                    headless=self.headless,
                    args=["--no-sandbox", "--disable-dev-shm-usage"]
                )
            except Exception as e:
                print(f"Error cr\u00edtico al lanzar navegador BV: {e}")
                return False

        self.context = self.browser.new_context(accept_downloads=True)
        self.page = self.context.new_page()

        try:
            self.page.goto("https://iip.bureauveritas.com.ar/Login.aspx", wait_until="load", timeout=60000)
            self.page.fill('input[name="txtUsuario"]', usuario)
            self.page.fill('input[name="txtPassword"]', clave)
            self.page.click('input[name="btnAcceso"]')

            try:
                self.page.wait_for_selector('a.ctl00_menua_1[href="Busquedas.aspx"]', timeout=15000)
                self.page.click('a.ctl00_menua_1[href="Busquedas.aspx"]')
                self.page.wait_for_selector('a#ctl00_ContentPlaceHolder1_lkBuscaEQ', timeout=10000)
                self.page.click('a#ctl00_ContentPlaceHolder1_lkBuscaEQ')
                return True
            except:
                return False
        except Exception as e:
            print(f"Error en login BV: {e}")
            return False

    def _extraer_fechas(self, texto):
        return re.findall(r'\d{2}/\d{2}/\d{4}', texto)

    def procesar_interno(self, interno, ruta_base):
        log_pasos = []
        try:
            self.page.check('input#ctl00_ContentPlaceHolder1_RBOpciones_1')
            self.page.fill('input#ctl00_ContentPlaceHolder1_txtNroBuscado', interno)
            self.page.click('input#ctl00_ContentPlaceHolder1_btnBusca')
            self.page.wait_for_timeout(2000)

            try:
                self.page.wait_for_selector('table#ctl00_ContentPlaceHolder1_gvInformes', timeout=10000)
            except:
                return {
                    "status": "No disponible en BV",
                    "venc": "-", "insp": "-", "cert": "NO", "inf": "NO",
                    "obs_final": "No encontrado en Bureau Veritas",
                    "accion_final": "-", "color": "ROJO",
                    "log": ["\u274c Interno no encontrado en Bureau Veritas."]
                }

            filas = self.page.query_selector_all('table#ctl00_ContentPlaceHolder1_gvInformes tr')
            fila_target = None
            celdas_target = []
            for fila in filas:
                celdas = fila.query_selector_all("td")
                if len(celdas) < 3:
                    continue
                texto_fila = " ".join(c.inner_text().strip() for c in celdas)
                if interno in texto_fila:
                    fila_target = fila
                    celdas_target = celdas
                    break

            if not fila_target:
                return {
                    "status": "No disponible en BV",
                    "venc": "-", "insp": "-", "cert": "NO", "inf": "NO",
                    "obs_final": "No encontrado en Bureau Veritas",
                    "accion_final": "-", "color": "ROJO",
                    "log": ["\u274c Interno no hallado en tabla BV."]
                }

            boton_pdf = fila_target.query_selector('input[id*="ImgbtnCertificado"]')
            if not boton_pdf:
                return {
                    "status": "Sin certificado en BV",
                    "venc": "-", "insp": "-", "cert": "NO", "inf": "NO",
                    "obs_final": "Equipo listado sin certificado disponible en BV",
                    "accion_final": "-", "color": "ROJO",
                    "log": ["\u26a0\ufe0f Interno en BV pero sin certificado."]
                }

            fechas = self._extraer_fechas(" ".join(c.inner_text().strip() for c in celdas_target))
            insp_fecha = fechas[0] if len(fechas) > 0 else "-"
            venc_fecha = fechas[1] if len(fechas) > 1 else "-"

            with self.page.expect_response(lambda r: "Prepara_PDF.aspx" in r.url) as response_info:
                boton_pdf.click()

            response_info.value
            self.page.wait_for_timeout(3000)

            pdf_b64 = self.page.evaluate("""
                (async () => {
                    try {
                        const resp = await fetch(window.location.href);
                        const blob = await resp.blob();
                        return await new Promise((resolve) => {
                            const reader = new FileReader();
                            reader.onloadend = () => resolve(reader.result.split(',')[1]);
                            reader.readAsDataURL(blob);
                        });
                    } catch(e) { return null; }
                })()
            """)

            if not pdf_b64:
                return {
                    "status": "Error descarga BV",
                    "venc": venc_fecha, "insp": insp_fecha,
                    "cert": "NO", "inf": "NO",
                    "obs_final": "Error al obtener PDF desde BV",
                    "accion_final": "-", "color": "ROJO",
                    "log": ["\u274c Error al descargar PDF de BV."]
                }

            file_path = os.path.join(ruta_base, f"{interno}_Certificado_BV.pdf")
            with open(file_path, "wb") as f:
                f.write(base64.b64decode(pdf_b64))

            log_pasos.append("\u2705 Certificado descargado desde Bureau Veritas")

            estado_f, _, _ = analizar_fecha(venc_fecha) if venc_fecha != "-" else ("SIN REGISTRO", "gray", False)

            if estado_f in ["VIGENTE", "PR\u00d3XIMO A VENCER"]:
                return {
                    "status": estado_f,
                    "venc": venc_fecha, "insp": insp_fecha,
                    "cert": "SI", "inf": "NO",
                    "obs_final": "Certificado vigente. Descargado desde Bureau Veritas.",
                    "accion_final": "-",
                    "color": "VERDE" if estado_f == "VIGENTE" else "AMARILLO",
                    "log": log_pasos
                }
            else:
                return {
                    "status": "VENCIDO (BV)",
                    "venc": venc_fecha, "insp": insp_fecha,
                    "cert": "SI", "inf": "NO",
                    "obs_final": "Certificado descargado pero vencido.",
                    "accion_final": "Coordinar recertificaci\u00f3n",
                    "color": "ROJO",
                    "log": log_pasos
                }

        except Exception as e:
            return {"status": "Error BV", "cert": "NO", "inf": "NO",
                    "log": [f"\u274c Error en BV: {str(e)[:50]}"]}

        finally:
            try:
                self.page.goto("https://iip.bureauveritas.com.ar/BuscaEQ.aspx",
                               wait_until="load", timeout=30000)
            except:
                pass

    def cerrar(self):
        if self.browser:
            self.browser.close()
        if self.pw:
            self.pw.stop()
