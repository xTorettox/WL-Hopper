"""
Scraper de Lenor Group para WL Hopper (Sullair Argentina).
Permite buscar certificados e informes de inspección a partir del número de serie / chasis
extraído del título de propiedad del equipo y valida que el interno coincida en el PDF.
"""
import os
import re
import logging
import urllib.parse
from datetime import datetime, timedelta
from typing import Dict, Any, List, Optional
import requests
from playwright.sync_api import sync_playwright

from pdf_utils import extraer_marca_modelo_de_certificado, extraer_texto_completo_pdf

logger = logging.getLogger("LenorBot")

class LenorBot:
    def __init__(self, headless: bool = True):
        self.headless = headless
        self.pw = None
        self.browser = None
        self.context = None
        self.page = None
        self.cookies_dict = {}
        self.pw_started_here = False
        self.base_url = "https://sistema.lenor.com.ar"
        self.panel_url = f"{self.base_url}/LenorMVC/PanelCliente"

    def iniciar(self, usuario: str = "ingresosnqn@sullair.com.ar", clave: str = "Sullair2025", pw_instance=None) -> tuple[bool, str]:
        """Inicia sesión en Lenor Group y valida el acceso."""
        try:
            if pw_instance:
                self.pw = pw_instance
                self.pw_started_here = False
            else:
                self.pw = sync_playwright().start()
                self.pw_started_here = True

            # Rutas para entornos Linux/Cloud
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
                "headless": self.headless,
                "args": ["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"]
            }
            if exec_path:
                launch_kwargs["executable_path"] = exec_path

            self.browser = self.pw.chromium.launch(**launch_kwargs)
            self.context = self.browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
                viewport={"width": 1366, "height": 768}
            )
            self.page = self.context.new_page()

            # Navegación a Login
            self.page.goto(self.panel_url, timeout=30000)

            # Si nos manda al login de Lenor
            if "Account/Login" in self.page.url:
                self.page.wait_for_selector('input#Email', timeout=15000)
                self.page.fill('input#Email', usuario)
                self.page.fill('input#Password', clave)
                self.page.click('input.lenor_btn_login')
                self.page.wait_for_load_state("networkidle", timeout=30000)

            # Verificar si logueó correctamente
            if "PanelCliente" in self.page.url and "Account/Login" not in self.page.url:
                self.cookies_dict = {c['name']: c['value'] for c in self.context.cookies()}
                logger.info("Login en Lenor exitoso.")
                return True, "Login exitoso en Lenor."
            else:
                return False, f"Credenciales de Lenor incorrectas o bloqueo: {self.page.url}"

        except Exception as e:
            logger.error(f"Error iniciando LenorBot: {e}")
            return False, str(e)

    def _obtener_digitos_serie(self, serie_completo: str) -> List[str]:
        """
        Extrae secuencias de búsqueda de 4 y 5 dígitos a partir del número de serie/chasis.
        Regla general del usuario: buscar primero con 4 dígitos, luego 5 dígitos.
        Ejemplo: KMTWA096LPA056772 -> ['6772', '56772', '056772']
        """
        if not serie_completo:
            return []

        candidatos = []
        # Buscar todos los grupos de dígitos
        todos_digitos = re.findall(r'\d+', serie_completo)
        if todos_digitos:
            ultimo_grupo = todos_digitos[-1]
            if len(ultimo_grupo) >= 4:
                candidatos.append(ultimo_grupo[-4:])
            if len(ultimo_grupo) >= 5:
                candidatos.append(ultimo_grupo[-5:])
            if len(ultimo_grupo) not in [4, 5]:
                candidatos.append(ultimo_grupo)
        
        # Si no hubo grupos numéricos puros, probar con los últimos caracteres alfanuméricos
        if not candidatos and len(serie_completo) >= 4:
            candidatos.append(serie_completo[-4:])
            if len(serie_completo) >= 5:
                candidatos.append(serie_completo[-5:])

        # Eliminar duplicados manteniendo orden
        vistos = set()
        resultado = []
        for c in candidatos:
            if c not in vistos and len(c) >= 3:
                resultado.append(c)
                vistos.add(c)
        return resultado

    def buscar_equipo(self, interno: str, chasis_o_serie: str, ruta_base: str, bajar_cert: bool = True, bajar_inf: bool = False, prefijo_cert: str = "Certificado") -> Dict[str, Any]:
        """
        Busca el equipo en Lenor utilizando los dígitos del chasis/serie y valida el interno
        directamente en el texto del certificado PDF descargado.
        """
        log_pasos = []
        interno_clean = interno.strip().upper()
        res = {
            "proveedor": "Lenor",
            "interno": interno_clean,
            "status": "No se pudo encontrar",
            "venc": "-",
            "venc_real": "-",
            "insp": "-",
            "marca": "",
            "modelo": "",
            "cert": "NO",
            "informe": "NO",
            "descargado": False,
            "observaciones": "",
            "archivos_descargados": [],
            "log": log_pasos
        }

        digitos_busqueda = self._obtener_digitos_serie(chasis_o_serie)
        if not digitos_busqueda:
            log_pasos.append(f"⚠️ No se pudo extraer número de serie del título para buscar en Lenor ({chasis_o_serie}).")
            res["log"] = log_pasos
            return res

        log_pasos.append(f"🔍 Lenor: Buscando interno {interno_clean} mediante serie/chasis '{chasis_o_serie}'...")

        encontrado = False
        for termino in digitos_busqueda:
            if encontrado:
                break

            log_pasos.append(f"   👉 Probando búsqueda en Lenor con dígitos '{termino}'...")
            
            # Consultar el endpoint de listado mediante POST
            url_list = f"{self.base_url}/LenorMVC/PanelCliente/ListAll"
            payload = {
                "EquipoIndustrialNumeroSerie": termino,
                "pageSize": "50",
                "page": "1",
                "sort": "NUMERO_SET",
                "sortDir": "DESC"
            }
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                "X-Requested-With": "XMLHttpRequest"
            }

            try:
                r = requests.post(url_list, data=payload, cookies=self.cookies_dict, headers=headers, timeout=25)
                html_resp = r.text

                # Si no hay registros
                if "No se encontraron registros" in html_resp or "Mostrando 0 a 0" in html_resp:
                    continue

                # Parsear filas individuales de la tabla para aislar cada equipo
                filas = re.findall(r'<tr[^>]*>.*?</tr>', html_resp, re.DOTALL)
                # Si no vinieron filas explícitas con tr, procesar todo el bloque
                if not filas:
                    filas = [html_resp]

                for fila in filas:
                    if encontrado:
                        break

                    patron_links = re.findall(r'href="(/LenorMVC/DocumentProvider/f/[a-zA-Z0-9\-]+)"[^>]*>.*?([0-9a-zA-Z_\-]+\.pdf)', fila, re.DOTALL)
                    if not patron_links:
                        continue

                    # Extraer Marca y Modelo del equipo en esta fila
                    m_equipo = re.search(r'<i class="icofont-cog"></i>\s*([^<]+)', fila)
                    marca_lenor = ""
                    modelo_lenor = ""
                    if m_equipo:
                        partes = [p.strip() for p in m_equipo.group(1).split("/")]
                        if len(partes) >= 2:
                            marca_lenor = partes[0]
                            modelo_lenor = partes[1]

                    # Descargar y auditar los PDFs de esta fila
                    archivos_descargados_lenor = []
                    fila_coincide = False

                    for link_rel, nombre_archivo in patron_links:
                        es_certificado = "CR.PDF" in nombre_archivo.upper() or "CERT" in nombre_archivo.upper()
                        es_informe = "TRF.PDF" in nombre_archivo.upper() or "INF" in nombre_archivo.upper()

                        doc_url = f"{self.base_url}{link_rel}"
                        temp_pdf_path = os.path.join(ruta_base, f"temp_{nombre_archivo}")

                        r_pdf = requests.get(doc_url, cookies=self.cookies_dict, headers=headers, timeout=30)
                        if r_pdf.status_code == 200:
                            with open(temp_pdf_path, "wb") as f_out:
                                f_out.write(r_pdf.content)

                            # VALIDACIÓN CRÍTICA: ¿El PDF o la fila contiene el número de interno buscado?
                            texto_pdf = extraer_texto_completo_pdf(temp_pdf_path)
                            tiene_interno = (interno_clean in texto_pdf.upper()) or (interno_clean in fila.upper())

                            if tiene_interno:
                                fila_coincide = True
                                log_pasos.append(f"🎯 ¡Coincidencia verificada en Lenor! El interno {interno_clean} figura en el documento/registro.")

                                # Extraer fecha de inspección del PDF de Lenor
                                m_insp = re.search(r'Fecha\s+de\s+inspecci[oó\W]+n\s*[:\s]+(\d{2}/\d{2}/\d{4})', texto_pdf, re.IGNORECASE)
                                fecha_insp_str = m_insp.group(1) if m_insp else "-"
                                
                                # Vencimiento: a 12 meses de la fecha de inspección
                                fecha_venc_str = "-"
                                if fecha_insp_str != "-":
                                    try:
                                        dt_insp = datetime.strptime(fecha_insp_str, "%d/%m/%Y")
                                        dt_venc = dt_insp + timedelta(days=365)
                                        fecha_venc_str = dt_venc.strftime("%d/%m/%Y")
                                    except Exception:
                                        pass

                                # Extraer Marca y Modelo del PDF
                                datos_cert = extraer_marca_modelo_de_certificado(temp_pdf_path)
                                res["marca"] = datos_cert.get("marca") or marca_lenor
                                res["modelo"] = datos_cert.get("modelo") or modelo_lenor
                                res["insp"] = fecha_insp_str
                                res["venc"] = fecha_venc_str
                                res["venc_real"] = fecha_venc_str

                                venc_format = fecha_venc_str.replace('/', '-') if fecha_venc_str != "-" else "SinFecha"
                                prefijo = f"{prefijo_cert}_" if prefijo_cert else ""

                                if es_certificado and bajar_cert:
                                    final_path = os.path.join(ruta_base, f"{prefijo}{interno_clean}_Certificado_Vence_{venc_format}_Lenor.pdf")
                                    os.replace(temp_pdf_path, final_path)
                                    archivos_descargados_lenor.append(final_path)
                                    res["cert"] = "SI"
                                    res["descargado"] = True
                                    log_pasos.append(f"💾 Certificado Lenor guardado: {os.path.basename(final_path)}")
                                elif es_informe and bajar_inf:
                                    final_path = os.path.join(ruta_base, f"{prefijo}{interno_clean}_Informe_Lenor_{venc_format}.pdf")
                                    os.replace(temp_pdf_path, final_path)
                                    archivos_descargados_lenor.append(final_path)
                                    res["informe"] = "SI"
                                    log_pasos.append(f"💾 Informe Lenor guardado: {os.path.basename(final_path)}")
                                else:
                                    if os.path.exists(temp_pdf_path):
                                        os.remove(temp_pdf_path)
                            else:
                                if os.path.exists(temp_pdf_path):
                                    os.remove(temp_pdf_path)

                    if fila_coincide:
                        encontrado = True
                        res["status"] = "VIGENTE (LENOR)" if res["descargado"] else "Encontrado en Lenor"
                        res["archivos_descargados"] = archivos_descargados_lenor
                        break

            except Exception as e_req:
                log_pasos.append(f"⚠️ Error consultando Lenor con término '{termino}': {e_req}")

        if not encontrado:
            log_pasos.append(f"❌ No se hallaron certificados coincidentes en Lenor para el interno {interno_clean}.")

        res["log"] = log_pasos
        return res

    def cerrar(self):
        """Cierra el navegador y contexto de Playwright."""
        try:
            if self.context:
                self.context.close()
            if self.browser:
                self.browser.close()
            if self.pw_started_here and self.pw:
                self.pw.stop()
        except Exception as e:
            logger.warning(f"Error cerrando LenorBot: {e}")
