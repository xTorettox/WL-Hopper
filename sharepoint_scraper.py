"""
Scraper de SharePoint para WL Hopper (Sullair Argentina).
Accede de forma segura a la carpeta blindada de BACK OFFICE ARG,
busca y descarga por número de interno:
  - Título + Cédula (ej: título + cédula E040293.pdf)
  - Acta Provisoria (ej: ACTA PROVISORIA 8906 - A040247.pdf)
  - Póliza de Seguro Técnico (ej: 207087 POLIZA - SEGURO TECNICO ...pdf)
  - Libre Deuda de Seguro (ej: 207087 POLIZA - LIBRE DEUDA SEGURO TECNICO ...pdf)
"""
import os
import re
import urllib.parse
import logging
from typing import List, Dict, Any, Optional
import requests
from playwright.sync_api import sync_playwright

from sharepoint_config import (
    SHAREPOINT_BASE_URL,
    SHAREPOINT_SITE_URL,
    SHAREPOINT_SERVER_RELATIVE_FOLDER,
    FOLDER_DOC_EQUIPOS,
    FOLDER_DOC_TITULOS,
    FOLDER_DOC_VTV,
    FOLDER_SEGUROS,
    FOLDER_CERTIFICACIONES,
    SHAREPOINT_TARGET_BROWSER_URL,
    STORAGE_STATE_PATH
)
from security_guard import HopperSecurityGuard, KillSwitchException
from auth_manager import SharePointAuthManager
from pdf_utils import extraer_datos_titulo_o_cedula

logger = logging.getLogger("SharePointScraper")

class SharePointHopperBot:
    def __init__(self, security_guard: HopperSecurityGuard, auth_manager: SharePointAuthManager, headless: bool = True):
        self.guard = security_guard
        self.auth = auth_manager
        self.headless = headless
        self.pw = None
        self.pw_started_here = False
        self.browser = None
        self.context = None
        self.page = None
        self.cookies_dict = {}
        self.cached_titulos_files = None # Cache de archivos de títulos/cédulas
        self.cached_vtv_files = None     # Cache de actas provisorias y técnicas
        self.cached_seguros_files = None # Cache de pólizas y libres deudas

    def iniciar(self, app_user: str = "fcendra", pw_instance=None) -> tuple[bool, str]:
        """
        Lanza el navegador con protecciones stealth, valida sesión existente o ejecuta
        login con Kill Switch estricto. Reutiliza la instancia existente de Playwright si se provee.
        """
        # 1. Comprobar si el Kill Switch está activo antes de cualquier acción
        locked, reason = self.guard.is_locked()
        if locked:
            msg = f"⛔ Scraper bloqueado preventivamente por Kill Switch: {reason}"
            logger.critical(msg)
            return False, msg

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

            browser_args = [
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-blink-features=AutomationControlled",
                "--disable-gpu"
            ]

            launch_kwargs = {
                "headless": self.headless,
                "args": browser_args
            }
            if exec_path:
                launch_kwargs["executable_path"] = exec_path

            self.browser = self.pw.chromium.launch(**launch_kwargs)

            # 2. Restaurar storageState si existe
            tiene_sesion = self.auth.restore_session_state()

            context_kwargs = {
                "user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
                "viewport": {"width": 1366, "height": 768},
                "locale": "es-ES"
            }
            if tiene_sesion and os.path.exists(STORAGE_STATE_PATH):
                context_kwargs["storage_state"] = STORAGE_STATE_PATH

            self.context = self.browser.new_context(**context_kwargs)
            self.context.add_init_script("delete Object.getPrototypeOf(navigator).webdriver")
            self.page = self.context.new_page()

            # 3. Comprobar si la sesión es válida o necesita login
            sesion_ok = False
            if tiene_sesion:
                sesion_ok = self.auth.is_session_valid(self.page)

            if not sesion_ok:
                logger.info("Sesión no válida o inexistente. Iniciando login automatizado...")
                u, p = self.auth.get_credentials(app_user)
                if not u or not p:
                    return False, "No se encontraron credenciales de SharePoint en la base de datos."
                
                # Ejecutar login paso a paso con Kill Switch
                self.auth.perform_automated_login(self.page, u, p, self.context)

            # 4. Confirmar navegación a SharePoint (usando domcontentloaded para evitar timeouts en SPA)
            self.page.goto(SHAREPOINT_TARGET_BROWSER_URL, wait_until="domcontentloaded", timeout=30000)
            if "sharepoint.com" not in self.page.url:
                self.guard.trigger_kill_switch("Desvío de URL", f"URL fuera de SharePoint: {self.page.url}")

            # Extraer cookies para requests REST ultrarrápidos
            self._actualizar_cookies()
            logger.info("SharePointHopperBot iniciado exitosamente.")
            return True, "Conexión exitosa a SharePoint."

        except KillSwitchException as kse:
            self.cerrar()
            return False, f"KILL SWITCH DISPARADO: {kse}"
        except Exception as e:
            self.cerrar()
            logger.error(f"Error al iniciar SharePointHopperBot: {e}")
            return False, str(e)

    def _actualizar_cookies(self):
        """Captura cookies del contexto de Playwright para uso en requests."""
        if self.context:
            self.cookies_dict = {c['name']: c['value'] for c in self.context.cookies()}

    def _get_api_headers(self) -> dict:
        return {
            "Accept": "application/json;odata=verbose",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        }

    def _listar_archivos_carpeta_rest(self, folder_relative_path: str) -> List[Dict[str, Any]]:
        """
        Consulta los archivos de una carpeta específica de SharePoint usando la API REST.
        """
        folder_clean = folder_relative_path.strip().replace("'", "''")
        encoded_folder = urllib.parse.quote(folder_clean, safe="/()+-")
        api_url = f"{SHAREPOINT_SITE_URL}/_api/web/GetFolderByServerRelativeUrl('{encoded_folder}')/Files?$select=Name,ServerRelativeUrl&$top=5000"

        try:
            r = requests.get(api_url, cookies=self.cookies_dict, headers=self._get_api_headers(), timeout=20)
            if r.status_code == 200:
                data = r.json()
                results = data.get("d", {}).get("results", [])
                return results
            else:
                logger.warning(f"API REST status {r.status_code} para carpeta {folder_relative_path}")
        except Exception as e:
            logger.warning(f"Error en consulta REST de carpeta {folder_relative_path}: {e}")

        return []

    def _listar_subcarpetas_rest(self, folder_relative_path: str) -> List[Dict[str, Any]]:
        """Consulta subcarpetas de una carpeta de SharePoint."""
        folder_clean = folder_relative_path.strip().replace("'", "''")
        encoded_folder = urllib.parse.quote(folder_clean, safe="/()+-")
        api_url = f"{SHAREPOINT_SITE_URL}/_api/web/GetFolderByServerRelativeUrl('{encoded_folder}')/Folders?$select=Name,ServerRelativeUrl"
        try:
            r = requests.get(api_url, cookies=self.cookies_dict, headers=self._get_api_headers(), timeout=15)
            if r.status_code == 200:
                return r.json().get("d", {}).get("results", [])
        except Exception as e:
            logger.warning(f"Error listando subcarpetas en {folder_relative_path}: {e}")
        return []

    def _descargar_archivo_rest(self, server_relative_url: str, destino_path: str) -> bool:
        """
        Descarga directamente un archivo binario desde su ruta relativa en SharePoint.
        """
        file_url = f"{SHAREPOINT_BASE_URL}{server_relative_url}"
        try:
            r = requests.get(file_url, cookies=self.cookies_dict, headers=self._get_api_headers(), timeout=30, stream=True)
            if r.status_code == 200:
                with open(destino_path, "wb") as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        f.write(chunk)
                logger.info(f"Archivo descargado exitosamente: {destino_path}")
                return True
        except Exception as e:
            logger.error(f"Error descargando {file_url}: {e}")
        return False

    def obtener_seguros_vigentes(self, ruta_descargas: str) -> List[str]:
        """
        Obtiene y descarga los documentos generales de Seguro Técnico y Libre Deuda
        presentes en la carpeta SEGUROS de BACK OFFICE Sullair Arg y sus subcarpetas.
        """
        if self.cached_seguros_files is not None:
            return self.cached_seguros_files

        archivos_seguros = []

        # Buscar en raíz de SEGUROS y en sus subcarpetas (Sancor, San Cristóbal, etc.)
        subcarpetas = self._listar_subcarpetas_rest(FOLDER_SEGUROS)
        rutas_a_explorar = [FOLDER_SEGUROS] + [sf.get("ServerRelativeUrl") for sf in subcarpetas if sf.get("ServerRelativeUrl")]
        
        # También incluir subcarpetas de segundo nivel como "Documentos Generales"
        for sf in subcarpetas:
            s_url = sf.get("ServerRelativeUrl")
            if s_url:
                sub2 = self._listar_subcarpetas_rest(s_url)
                for s2 in sub2:
                    if s2.get("ServerRelativeUrl"):
                        rutas_a_explorar.append(s2.get("ServerRelativeUrl"))

        todos_los_archivos = []
        for r_path in rutas_a_explorar:
            todos_los_archivos.extend(self._listar_archivos_carpeta_rest(r_path))

        poliza_file = None
        libre_deuda_file = None

        for f in todos_los_archivos:
            name = f.get("Name", "")
            name_upper = name.upper()
            
            # Identificar Libre Deuda vigente (priorizar Sancor 207087 o San Cristobal)
            if ("LIBRE DEUDA" in name_upper or "LD " in name_upper or name_upper.startswith("LD")) and ("SEGURO" in name_upper or "ST" in name_upper or "RC" in name_upper or "207087" in name_upper):
                if not libre_deuda_file or "207087" in name_upper:
                    libre_deuda_file = f
            # Identificar Póliza de Seguro Técnico (priorizar 207087 o Poliza Completa)
            elif "POLIZA" in name_upper and ("SEGURO" in name_upper or "ST" in name_upper or "207087" in name_upper):
                if not poliza_file or "COMPLETA" in name_upper or "207087" in name_upper:
                    poliza_file = f

        for candidato in [poliza_file, libre_deuda_file]:
            if candidato:
                file_name = candidato.get("Name")
                dest = os.path.join(ruta_descargas, file_name)
                if not os.path.exists(dest):
                    if self._descargar_archivo_rest(candidato.get("ServerRelativeUrl"), dest):
                        archivos_seguros.append(dest)
                else:
                    archivos_seguros.append(dest)

        self.cached_seguros_files = archivos_seguros
        return archivos_seguros

    def procesar_interno(self, interno: str, ruta_descargas: str) -> Dict[str, Any]:
        """
        Busca y descarga los archivos correspondientes a un número de interno en SharePoint:
        - Título + Cédula (en TITULO Y CEDULA - FACTURAS)
        - Acta Provisoria / Técnica (en VTV-RTO o CERTIFICACIONES)
        - Seguro Técnico y Libre Deuda (en SEGUROS)
        """
        log_pasos = []
        interno_clean = interno.strip().upper()
        res = {
            "interno": interno_clean,
            "titulo_cedula": False,
            "acta_provisoria": False,
            "seguro_poliza": False,
            "libre_deuda": False,
            "chasis": "",
            "marca": "",
            "modelo": "",
            "archivos_descargados": [],
            "log": log_pasos
        }

        log_pasos.append(f"🔍 SharePoint: Buscando documentación para interno {interno_clean}...")

        # 1. BÚSQUEDA DE TÍTULO + CÉDULA
        if self.cached_titulos_files is None:
            self.cached_titulos_files = self._listar_archivos_carpeta_rest(FOLDER_DOC_TITULOS)
            if not self.cached_titulos_files:
                self.cached_titulos_files = self._listar_archivos_carpeta_rest(FOLDER_DOC_EQUIPOS)

        encontrado_titulo = False
        for f in (self.cached_titulos_files or []):
            name = f.get("Name", "")
            name_upper = name.upper()
            if interno_clean in name_upper and ("TITULO" in name_upper or "TÍTULO" in name_upper or "CEDULA" in name_upper or "CÉDULA" in name_upper or "FC" in name_upper):
                dest = os.path.join(ruta_descargas, name)
                if not os.path.exists(dest):
                    descargado = self._descargar_archivo_rest(f.get("ServerRelativeUrl"), dest)
                else:
                    descargado = True

                if descargado:
                    res["titulo_cedula"] = True
                    if dest not in res["archivos_descargados"]:
                        res["archivos_descargados"].append(dest)
                    log_pasos.append(f"📄 Título + Cédula encontrado: {name}")
                    
                    # Extraer Chasis, Marca y Modelo directamente del PDF
                    datos_tit = extraer_datos_titulo_o_cedula(dest)
                    res["chasis"] = datos_tit.get("chasis", "")
                    res["marca"] = datos_tit.get("marca", "")
                    res["modelo"] = datos_tit.get("modelo", "")
                    if res["chasis"]:
                        log_pasos.append(f"🔢 Datos del Título: Chasis={res['chasis']} | Marca={res['marca']} | Modelo={res['modelo']}")
                    
                    encontrado_titulo = True
                    break

        if not encontrado_titulo:
            log_pasos.append("ℹ️ Título + Cédula no hallado en SharePoint.")

        # 2. BÚSQUEDA DE ACTA PROVISORIA O CERTIFICADO MANUAL
        if self.cached_vtv_files is None:
            vtv_f = self._listar_archivos_carpeta_rest(FOLDER_DOC_VTV)
            cert_f = self._listar_archivos_carpeta_rest(FOLDER_CERTIFICACIONES)
            self.cached_vtv_files = vtv_f + cert_f

        for f in (self.cached_vtv_files or []):
            name = f.get("Name", "")
            name_upper = name.upper()
            if interno_clean in name_upper and ("ACTA" in name_upper or "PROVISORIA" in name_upper or "TECNICA" in name_upper or "TÉCNICA" in name_upper or "CERT" in name_upper):
                dest = os.path.join(ruta_descargas, name)
                if not os.path.exists(dest):
                    descargado = self._descargar_archivo_rest(f.get("ServerRelativeUrl"), dest)
                else:
                    descargado = True

                if descargado:
                    res["acta_provisoria"] = True
                    if dest not in res["archivos_descargados"]:
                        res["archivos_descargados"].append(dest)
                    log_pasos.append(f"📜 Certificado/Acta Provisoria encontrada: {name}")
                    break

        # 3. VINCULACIÓN DE SEGUROS (Póliza y Libre Deuda)
        try:
            seguros_descargados = self.obtener_seguros_vigentes(ruta_descargas)
            for s_path in seguros_descargados:
                base_name = os.path.basename(s_path).upper()
                if "LIBRE DEUDA" in base_name or "LD " in base_name or base_name.startswith("LD"):
                    res["libre_deuda"] = True
                if "POLIZA" in base_name or "ENDOSO" in base_name:
                    res["seguro_poliza"] = True
                if s_path not in res["archivos_descargados"]:
                    res["archivos_descargados"].append(s_path)
            if seguros_descargados:
                log_pasos.append(f"🛡️ Seguros vinculados: {len(seguros_descargados)} archivo(s) de póliza/libre deuda.")
        except Exception as e_seg:
            log_pasos.append(f"⚠️ Error vinculando seguros de SharePoint: {e_seg}")

        return res

    def cerrar(self):
        """Cierra el contexto y navegador de Playwright limpiamente."""
        try:
            if self.page:
                try: self.page.close()
                except Exception: pass
            if self.context:
                try: self.context.close()
                except Exception: pass
            if self.browser:
                try: self.browser.close()
                except Exception: pass
            if getattr(self, "pw_started_here", False) and self.pw:
                try: self.pw.stop()
                except Exception: pass
        except Exception as e:
            logger.warning(f"Error cerrando SharePointHopperBot: {e}")
