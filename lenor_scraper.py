"""
Scraper de Lenor Group para WL Hopper (Sullair Argentina).
Busca certificados (CR.pdf) e informes técnicos (TRF.pdf) priorizando ID de Equipo,
con fallback por Número de Serie/Chasis y Motor, filtro estricto de cliente Sullair,
deduplicación por UUID y validación en memoria del campo 'Número de Interno:'.
"""
import os
import re
import logging
import urllib.parse
from datetime import datetime, timedelta
from typing import Dict, Any, List, Optional
import requests
from playwright.sync_api import sync_playwright

from pdf_utils import extraer_texto_pdf_bytes

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
        self.browser_started_here = False
        self.base_url = "https://sistema.lenor.com.ar"
        self.panel_url = f"{self.base_url}/LenorMVC/PanelCliente"
        # Cache intra-corrida: cuando un Set trae 2+ equipos (ej. SET #619812 -> E030429 + E030591),
        # guardamos los CR verificados en memoria por su Número de Interno real.
        self._cache_por_interno: Dict[str, Dict[str, Any]] = {}
        self._uuids_auditados_global: set[str] = set()
        self._recent_sets_scanned: bool = False

    def iniciar(self, usuario: str = "ingresosnqn@sullair.com.ar", clave: str = "Sullair2025", pw_instance=None, browser_instance=None) -> tuple[bool, str]:
        """Inicia sesión en Lenor Group y valida el acceso."""
        try:
            if pw_instance:
                self.pw = pw_instance
                self.pw_started_here = False
            else:
                self.pw = sync_playwright().start()
                self.pw_started_here = True

            self.browser_started_here = False
            if browser_instance:
                self.browser = browser_instance
            else:
                from scraper import lanzar_navegador_compartido
                self.browser = lanzar_navegador_compartido(self.pw, headless=self.headless)
                self.browser_started_here = True
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
                
                # ¡LIBERACIÓN INMEDIATA DE MEMORIA PARA STREAMLIT CLOUD!
                # Lenor realiza todas sus búsquedas y descargas vía peticiones REST (requests).
                # Cerramos página y contexto inmediatamente para ahorrar ~180MB de RAM.
                try:
                    if self.page:
                        self.page.close()
                        self.page = None
                    if self.context:
                        self.context.close()
                        self.context = None
                    if getattr(self, "browser_started_here", False) and self.browser:
                        self.browser.close()
                        self.browser = None
                except Exception:
                    pass

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

    @staticmethod
    def _normalizar_interno(val: str) -> str:
        return re.sub(r'\s+', '', (val or '').strip().upper())

    @staticmethod
    def _extraer_interno_de_texto_pdf(texto_pdf: str) -> str:
        """Extrae el valor exacto de 'Número de Interno:' del texto nativo del PDF de Lenor."""
        if not texto_pdf:
            return ""
        m = re.search(
            r'N[uú]mero\s+de\s+Interno\s*[:\s]+([A-Za-z0-9\-/]+)',
            texto_pdf,
            re.IGNORECASE,
        )
        return m.group(1).strip(" ,;.:-_").upper() if m else ""

    @staticmethod
    def _extraer_metadatos_pdf(texto_pdf: str) -> Dict[str, str]:
        """Extrae Fecha de inspección, vencimiento (+12 meses), Marca y Modelo desde texto nativo."""
        m_insp = re.search(
            r'Fecha\s+de\s+inspecci[oó\W]+n\s*[:\s]+(\d{2}/\d{2}/\d{4})',
            texto_pdf,
            re.IGNORECASE,
        )
        fecha_insp_str = m_insp.group(1) if m_insp else "-"
        fecha_venc_str = "-"
        if fecha_insp_str != "-":
            try:
                dt_insp = datetime.strptime(fecha_insp_str, "%d/%m/%Y")
                fecha_venc_str = (dt_insp + timedelta(days=365)).strftime("%d/%m/%Y")
            except Exception:
                pass

        marca = ""
        m_marca = re.search(r'(?:^|\n)\s*Marca\s*[:\s]+([^\n\r,;]+)', texto_pdf, re.IGNORECASE)
        if m_marca:
            marca = re.split(r'\s{2,}', m_marca.group(1))[0].strip(" ,;.:-_")

        modelo = ""
        m_modelo = re.search(r'(?:^|\n)\s*Modelo\s*[:\s]+([^\n\r,;]+)', texto_pdf, re.IGNORECASE)
        if m_modelo:
            modelo = re.split(r'\s{2,}', m_modelo.group(1))[0].strip(" ,;.:-_")

        return {
            "insp": fecha_insp_str,
            "venc": fecha_venc_str,
            "marca": marca,
            "modelo": modelo,
        }

    @staticmethod
    def _extraer_stem_documento(nombre_archivo: str) -> str:
        """Extrae el identificador numérico común entre CR y TRF (ej. '..._120257CR.pdf' -> '120257')."""
        m = re.search(r'_([0-9]+)(?:CR|TRF)\.pdf$', nombre_archivo, re.IGNORECASE)
        if m:
            return m.group(1)
        m2 = re.search(r'([0-9]+)(?:CR|TRF)\.pdf$', nombre_archivo, re.IGNORECASE)
        return m2.group(1) if m2 else ""

    @staticmethod
    def _es_fila_sullair(tds: List[str]) -> bool:
        """Valida que la columna 2 (td[1]: Cliente / Descripción del servicio) corresponda a SULLAIR o prefijo SUL-."""
        if len(tds) < 2:
            return False
        col2_txt = re.sub(r'<[^>]+>', ' ', tds[1]).upper()
        return ("SULLAIR" in col2_txt) or ("SUL-" in col2_txt)

    @staticmethod
    def _extraer_links_columna_equipo(tds: List[str]) -> tuple[List[tuple[str, str]], Dict[str, tuple[str, str]]]:
        """
        Lee ÚNICAMENTE la columna 3 (td[2]: Operador / Equipo), ignora columna 5 (Informe de Ensayo),
        deduplica por UUID (/DocumentProvider/f/{uuid}) y separa CR.pdf (Certificados) vs TRF.pdf (Informes).
        """
        if len(tds) < 3:
            return [], {}
        col3_html = tds[2]
        raw_links = re.findall(
            r'href="(/LenorMVC/DocumentProvider/f/([a-zA-Z0-9\-]+))"[^>]*>.*?([0-9a-zA-Z_\-]+\.pdf)',
            col3_html,
            re.DOTALL | re.IGNORECASE,
        )
        seen_uuids = set()
        cr_links: List[tuple[str, str]] = []
        trf_by_stem: Dict[str, tuple[str, str]] = {}

        for link_rel, doc_uuid, nombre_archivo in raw_links:
            uuid_norm = doc_uuid.lower()
            if uuid_norm in seen_uuids:
                continue
            seen_uuids.add(uuid_norm)

            nombre_up = nombre_archivo.upper()
            if nombre_up.endswith("CR.PDF"):
                cr_links.append((link_rel, nombre_archivo))
            elif nombre_up.endswith("TRF.PDF"):
                stem = LenorBot._extraer_stem_documento(nombre_archivo)
                if stem and stem not in trf_by_stem:
                    trf_by_stem[stem] = (link_rel, nombre_archivo)

        return cr_links, trf_by_stem

    def _guardar_match_lenor(
        self,
        res: Dict[str, Any],
        log_pasos: List[str],
        interno_clean: str,
        match_info: Dict[str, Any],
        ruta_base: str,
        bajar_cert: bool,
        bajar_inf: bool,
        prefijo_cert: str,
        headers: Dict[str, str],
    ) -> None:
        """Persiste en disco el CR.pdf (en memoria) y opcionalmente el TRF.pdf emparejado."""
        meta = match_info["meta"]
        res["marca"] = meta.get("marca") or match_info.get("marca_dom", "")
        res["modelo"] = meta.get("modelo") or match_info.get("modelo_dom", "")
        res["insp"] = meta.get("insp", "-")
        res["venc"] = meta.get("venc", "-")
        res["venc_real"] = meta.get("venc", "-")

        venc_format = res["venc"].replace("/", "-") if res["venc"] != "-" else "SinFecha"
        prefijo = f"{prefijo_cert}_" if prefijo_cert else ""
        archivos = []

        if bajar_cert and match_info.get("cr_bytes"):
            final_cert = os.path.join(ruta_base, f"{prefijo}{interno_clean}_Certificado_Vence_{venc_format}_Lenor.pdf")
            with open(final_cert, "wb") as f_out:
                f_out.write(match_info["cr_bytes"])
            archivos.append(final_cert)
            res["cert"] = "SI"
            res["descargado"] = True
            log_pasos.append(f"💾 Certificado Lenor guardado: {os.path.basename(final_cert)}")

        if bajar_inf and match_info.get("trf_link"):
            trf_url = f"{self.base_url}{match_info['trf_link']}"
            try:
                r_trf = requests.get(trf_url, cookies=self.cookies_dict, headers=headers, timeout=30)
                if r_trf.status_code == 200 and r_trf.content:
                    final_inf = os.path.join(ruta_base, f"{prefijo}{interno_clean}_Informe_Lenor_{venc_format}.pdf")
                    with open(final_inf, "wb") as f_out:
                        f_out.write(r_trf.content)
                    archivos.append(final_inf)
                    res["informe"] = "SI"
                    log_pasos.append(f"💾 Informe técnico Lenor (TRF) guardado: {os.path.basename(final_inf)}")
            except Exception as e_trf:
                log_pasos.append(f"⚠️ No se pudo descargar TRF de Lenor: {e_trf}")

        res["status"] = "VIGENTE (LENOR)" if res["descargado"] else "Encontrado en Lenor"
        res["archivos_descargados"] = archivos

    def _auditar_html_lenor(
        self,
        html_resp: str,
        interno_clean: str,
        headers: Dict[str, str],
        detener_al_encontrar: bool = True,
    ) -> Optional[Dict[str, Any]]:
        """
        Recorre filas HTML de ListAll filtrando estrictamente por SULLAIR/SUL- (td[1]),
        lee solo columna 3 (td[2]), deduplica por UUID, descarga CR.pdf en memoria,
        valida 'Número de Interno:' y alimenta self._cache_por_interno.
        """
        if "No se encontraron registros" in html_resp or "Mostrando 0 a 0" in html_resp:
            return None

        filas = [
            f for f in re.findall(r'<tr[^>]*>(.*?)</tr>', html_resp, re.DOTALL | re.IGNORECASE)
            if '<td' in f.lower()
        ]
        if not filas:
            return None

        match_objetivo: Optional[Dict[str, Any]] = None

        for fila in filas:
            if match_objetivo and detener_al_encontrar:
                break

            tds = re.findall(r'<td[^>]*>(.*?)</td>', fila, re.DOTALL | re.IGNORECASE)
            # Filtro estricto por cliente SULLAIR / SUL- en columna 2 (td[1])
            if not self._es_fila_sullair(tds):
                continue

            # Extraer solo links únicos (deduplicados por UUID) desde columna 3 (td[2])
            cr_links, trf_by_stem = self._extraer_links_columna_equipo(tds)
            if not cr_links:
                continue

            # Marca/modelo de respaldo desde el DOM de columna 3
            marca_dom, modelo_dom = "", ""
            m_equipo = re.search(r'<i class="icofont-cog"></i>\s*([^<]+)', tds[2])
            if m_equipo:
                partes = [p.strip() for p in m_equipo.group(1).split("/")]
                if len(partes) >= 2:
                    marca_dom, modelo_dom = partes[0], partes[1]

            for link_rel, nombre_cr in cr_links:
                if link_rel in self._uuids_auditados_global:
                    continue
                self._uuids_auditados_global.add(link_rel)

                doc_url = f"{self.base_url}{link_rel}"
                r_pdf = requests.get(doc_url, cookies=self.cookies_dict, headers=headers, timeout=30)
                if r_pdf.status_code != 200 or not r_pdf.content:
                    continue

                # Validar en memoria el campo exacto 'Número de Interno:' del CR.pdf
                texto_pdf = extraer_texto_pdf_bytes(r_pdf.content)
                interno_pdf = self._normalizar_interno(self._extraer_interno_de_texto_pdf(texto_pdf))
                if not interno_pdf:
                    continue

                stem_cr = self._extraer_stem_documento(nombre_cr)
                trf_pair = trf_by_stem.get(stem_cr)
                meta_pdf = self._extraer_metadatos_pdf(texto_pdf)

                match_info = {
                    "cr_bytes": r_pdf.content,
                    "cr_name": nombre_cr,
                    "trf_link": trf_pair[0] if trf_pair else None,
                    "trf_name": trf_pair[1] if trf_pair else None,
                    "meta": meta_pdf,
                    "marca_dom": marca_dom,
                    "modelo_dom": modelo_dom,
                }

                # Cachear en memoria por si otro interno del mismo Set se consulta luego en la corrida
                if interno_pdf not in self._cache_por_interno:
                    self._cache_por_interno[interno_pdf] = match_info

                if interno_pdf == interno_clean and not match_objetivo:
                    match_objetivo = match_info
                    if detener_al_encontrar:
                        break

        return match_objetivo or self._cache_por_interno.get(interno_clean)

    def buscar_equipo(
        self,
        interno: str,
        chasis_o_serie: str = "",
        motor: str = "",
        ruta_base: str = ".",
        bajar_cert: bool = True,
        bajar_inf: bool = False,
        prefijo_cert: str = "Certificado",
    ) -> Dict[str, Any]:
        """
        Estrategia robusta para Lenor:
        1. Busca primero por `ID equipo` (`NUMERO_ID_INTERNO_EQUIPO`) con el interno directo (sin depender de SharePoint).
        2. Solo si devuelve 0 coincidencias, usa `# Serie` (`EquipoIndustrialNumeroSerie`) con chasis/serie y motor como fallback.
        3. Filtra filas del DOM cuya columna 2 (Cliente) contenga SULLAIR o SUL-.
        4. Lee solo columna 3 (Operador/Equipo), deduplica por UUID, separa CR.pdf vs TRF.pdf.
        5. Valida `Número de Interno:` en memoria sobre el CR.pdf antes de escribir a disco.
        """
        log_pasos = []
        interno_clean = self._normalizar_interno(interno)
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
            "log": log_pasos,
        }

        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "X-Requested-With": "XMLHttpRequest",
        }

        # 0. Verificar si un Set multi-equipo ya escaneado en esta misma corrida contenía este interno
        cached_match = self._cache_por_interno.get(interno_clean)
        if cached_match:
            log_pasos.append(f"⚡ Lenor: Interno {interno_clean} resuelto desde Set ya auditado en esta corrida ({cached_match.get('cr_name', 'CR.pdf')}).")
            self._guardar_match_lenor(res, log_pasos, interno_clean, cached_match, ruta_base, bajar_cert, bajar_inf, prefijo_cert, headers)
            res["log"] = log_pasos
            return res

        # 1. Construir cola de consultas priorizadas: primero ID equipo, luego fallback por Serie/Chasis y Motor
        consultas: List[tuple[str, str, str]] = [
            ("ID equipo", "NUMERO_ID_INTERNO_EQUIPO", interno_clean),
        ]

        terminos_serie_vistos = set()
        for origen_lbl, valor_serie in [("Serie/Chasis", chasis_o_serie), ("Motor", motor)]:
            for dig in self._obtener_digitos_serie(valor_serie):
                if dig not in terminos_serie_vistos:
                    terminos_serie_vistos.add(dig)
                    consultas.append((origen_lbl, "EquipoIndustrialNumeroSerie", dig))

        log_pasos.append(f"🔍 Lenor: Buscando interno {interno_clean} (prioridad: ID equipo; fallback: Serie/Chasis/Motor)...")

        url_list = f"{self.base_url}/LenorMVC/PanelCliente/ListAll"
        encontrado = False

        for etiqueta_modo, campo_post, termino in consultas:
            if encontrado:
                break

            log_pasos.append(f"   👉 Consultando Lenor por {etiqueta_modo} = '{termino}'...")
            payload = {
                campo_post: termino,
                "pageSize": "50",
                "page": "1",
                "sort": "NUMERO_SET",
                "sortDir": "DESC",
            }

            try:
                r = requests.post(url_list, data=payload, cookies=self.cookies_dict, headers=headers, timeout=25)
                match_info = self._auditar_html_lenor(r.text, interno_clean, headers, detener_al_encontrar=True)
                if match_info:
                    meta_pdf = match_info["meta"]
                    log_pasos.append(
                        f"🎯 ¡Coincidencia verificada en PDF Lenor ({match_info['cr_name']})! "
                        f"Número de Interno: {interno_clean} | Inspección: {meta_pdf['insp']}"
                    )
                    self._guardar_match_lenor(
                        res, log_pasos, interno_clean, match_info,
                        ruta_base, bajar_cert, bajar_inf, prefijo_cert, headers
                    )
                    encontrado = True
                    break

            except Exception as e_req:
                log_pasos.append(f"⚠️ Error consultando Lenor ({etiqueta_modo}='{termino}'): {e_req}")

        # 2. Fallback ligero (una sola vez por corrida): si Lenor no indexó el ID en la grilla
        # (pasa en Sets multi-equipo donde solo indexan el 1er equipo, ej. E030591 en SET #619812),
        # auditamos los Sets Sullair más recientes (pageSize=15) y poblamos la caché en memoria.
        if not encontrado and not self._recent_sets_scanned:
            self._recent_sets_scanned = True
            log_pasos.append("   🔎 Fallback Lenor: Auditando últimos Sets Sullair recientes por si el ID no fue indexado en grilla...")
            try:
                r_rec = requests.post(
                    url_list,
                    data={"pageSize": "15", "page": "1", "sort": "NUMERO_SET", "sortDir": "DESC"},
                    cookies=self.cookies_dict,
                    headers=headers,
                    timeout=25,
                )
                match_info = self._auditar_html_lenor(r_rec.text, interno_clean, headers, detener_al_encontrar=True)
                if match_info:
                    meta_pdf = match_info["meta"]
                    log_pasos.append(
                        f"🎯 ¡Coincidencia recuperada en Sets recientes ({match_info['cr_name']})! "
                        f"Número de Interno: {interno_clean} | Inspección: {meta_pdf['insp']}"
                    )
                    self._guardar_match_lenor(
                        res, log_pasos, interno_clean, match_info,
                        ruta_base, bajar_cert, bajar_inf, prefijo_cert, headers
                    )
                    encontrado = True
            except Exception as e_rec:
                log_pasos.append(f"⚠️ Error en fallback de Sets recientes de Lenor: {e_rec}")

        if not encontrado:
            log_pasos.append(f"❌ No se hallaron certificados Sullair en Lenor para el interno {interno_clean}.")

        res["log"] = log_pasos
        return res

    def cerrar(self):
        """Cierra el navegador y contexto de Playwright si quedaron activos."""
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
