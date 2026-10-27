"""
Manejador de Autenticación, Sesión y Login Sigiloso para SharePoint (Microsoft 365).
Integra almacenamiento encriptado de storageState, delays realistas y disparo
inmediato del Kill Switch ante cualquier desafío no previsto.
"""
import os
import json
import time
import random
import logging
from datetime import datetime
from typing import Tuple, Optional
from playwright.sync_api import BrowserContext, Page

from sharepoint_config import (
    SHAREPOINT_BASE_URL,
    SHAREPOINT_SITE_URL,
    STORAGE_STATE_PATH,
    STORAGE_STATE_DB_KEY,
    CREDENTIALS_DB_KEY,
    MS_LOGIN_URL_PREFIX,
    SEL_EMAIL_INPUT,
    SEL_PASSWORD_INPUT,
    SEL_SUBMIT_BTN,
    SEL_KMSI_CHECK,
    DANGER_SELECTORS,
)
from security_guard import HopperSecurityGuard, KillSwitchException

logger = logging.getLogger("HopperAuthManager")

class SharePointAuthManager:
    def __init__(self, security_guard: HopperSecurityGuard, cipher_suite=None, supabase_client=None):
        self.guard = security_guard
        self.cipher = cipher_suite
        self.supabase = supabase_client

    def human_delay(self, min_s: float = 1.0, max_s: float = 2.5):
        """Pausa aleatoria entre acciones para emular cadencia humana."""
        time.sleep(random.uniform(min_s, max_s))

    def type_realistic(self, locator, text: str):
        """Tipea caracter a caracter con micro-variaciones aleatorias."""
        for char in text:
            locator.type(char, delay=random.randint(50, 150))
            if random.random() < 0.08:
                time.sleep(random.uniform(0.12, 0.30))

    def get_credentials(self, app_user: str = "") -> Tuple[str, str]:
        """
        Recupera y desencripta el usuario y contraseña corporativos desde Supabase.
        """
        if not self.supabase or not self.cipher:
            self.guard.trigger_kill_switch("Configuración incompleta", "Falta cliente de Supabase o clave Fernet.")

        try:
            # 1. Buscar credencial específica de SharePoint
            res = self.supabase.table("credenciales_sitios").select("*").eq("sitio", CREDENTIALS_DB_KEY).execute()
            if not res.data:
                # 2. Buscar por patrón de sitio SP_ o usuario logueado
                res = self.supabase.table("credenciales_sitios").select("*").like("sitio", "SP_%").execute()

            if res.data:
                for row in res.data:
                    # Ignorar registros de estado o sesión
                    if row.get("sitio") in ["SP_KILL_SWITCH", "SP_STORAGE_STATE"]:
                        continue
                    u_enc = row.get("user_enc", "")
                    p_enc = row.get("pass_enc", "")
                    if u_enc and p_enc:
                        u = self.cipher.decrypt(u_enc.encode()).decode()
                        p = self.cipher.decrypt(p_enc.encode()).decode()
                        return u, p
        except Exception as e:
            self.guard.trigger_kill_switch("Error al obtener credenciales de la BD", str(e))

        self.guard.trigger_kill_switch(
            "Credenciales de SharePoint no encontradas",
            "No se hallaron credenciales configuradas en la tabla credenciales_sitios para SharePoint."
        )
        return "", ""

    def restore_session_state(self) -> bool:
        """
        Restaura el storage_state.json desde el disco local o desde Supabase.
        Retorna True si existe un estado para inyectar en el contexto de Playwright.
        """
        # 1. Revisar archivo local
        if os.path.exists(STORAGE_STATE_PATH) and os.path.getsize(STORAGE_STATE_PATH) > 10:
            logger.info("StorageState localizado en disco local.")
            return True

        # 2. Si no está local, recuperarlo de Supabase (respaldo cifrado)
        if self.supabase and self.cipher:
            try:
                res = self.supabase.table("credenciales_sitios").select("*").eq("sitio", STORAGE_STATE_DB_KEY).execute()
                if res.data:
                    row = res.data[0]
                    state_enc = row.get("pass_enc", "")
                    if state_enc:
                        state_json = self.cipher.decrypt(state_enc.encode()).decode()
                        with open(STORAGE_STATE_PATH, "w", encoding="utf-8") as f:
                            f.write(state_json)
                        logger.info("StorageState recuperado y sincronizado desde Supabase.")
                        return True
            except Exception as e:
                logger.warning(f"No se pudo restaurar StorageState desde Supabase: {e}")

        return False

    def save_session_state(self, context: BrowserContext):
        """
        Guarda las cookies y tokens del navegador en disco y en Supabase cifrado.
        """
        try:
            context.storage_state(path=STORAGE_STATE_PATH)
            logger.info(f"StorageState guardado en {STORAGE_STATE_PATH}.")

            if self.supabase and self.cipher and os.path.exists(STORAGE_STATE_PATH):
                with open(STORAGE_STATE_PATH, "r", encoding="utf-8") as f:
                    state_raw = f.read()
                
                state_enc = self.cipher.encrypt(state_raw.encode()).decode()
                timestamp = datetime.now().isoformat()
                payload = {
                    "usuario_app": "system_auth",
                    "sitio": STORAGE_STATE_DB_KEY,
                    "user_enc": timestamp,
                    "pass_enc": state_enc
                }
                
                res = self.supabase.table("credenciales_sitios").select("id").eq("sitio", STORAGE_STATE_DB_KEY).execute()
                if res.data:
                    self.supabase.table("credenciales_sitios").update(payload).eq("sitio", STORAGE_STATE_DB_KEY).execute()
                else:
                    self.supabase.table("credenciales_sitios").insert(payload).execute()
                logger.info("StorageState respaldado y cifrado en Supabase.")
        except Exception as e:
            logger.error(f"Error al persistir storage_state: {e}")

    def is_session_valid(self, page: Page) -> bool:
        """
        Verifica de forma no invasiva si la sesión actual permite acceso a SharePoint.
        """
        try:
            logger.info(f"Comprobando validez de sesión en {SHAREPOINT_SITE_URL}...")
            page.goto(SHAREPOINT_SITE_URL, wait_until="domcontentloaded", timeout=25000)
            self.human_delay(1.5, 2.5)

            # Si la URL sigue en SharePoint y no fue redirigida a login.microsoftonline
            if "sharepoint.com" in page.url and MS_LOGIN_URL_PREFIX not in page.url:
                logger.info("Sesión VÁLIDA. Acceso confirmado a SharePoint.")
                return True
        except Exception as e:
            logger.warning(f"Error comprobando sesión previa: {e}")

        return False

    def perform_automated_login(self, page: Page, user: str, password: str, context: BrowserContext):
        """
        Ejecuta el flujo de login headless estricto en Microsoft 365.
        REGLA DE ORO: Al primer error o pantalla imprevista, ABORTA SIN REINTENTOS Y DISPARA EL KILL SWITCH.
        """
        logger.info("Iniciando flujo de login estricto en Microsoft Online...")

        try:
            page.goto(SHAREPOINT_SITE_URL, wait_until="domcontentloaded", timeout=30000)
            self.human_delay(2.0, 3.5)

            # Si ya ingresó directamente
            if "sharepoint.com" in page.url and MS_LOGIN_URL_PREFIX not in page.url:
                self.save_session_state(context)
                return

            # --- PASO 1: Ingreso de Correo Electrónico ---
            self._check_danger_screens(page, "Antes de ingresar email")
            
            email_field = page.locator(SEL_EMAIL_INPUT)
            if email_field.count() == 0 or not email_field.first.is_visible():
                self.guard.trigger_kill_switch("Campo de email no visible en la página de login")

            self.human_delay(0.8, 1.6)
            email_field.first.click()
            self.type_realistic(email_field.first, user)
            self.human_delay(1.0, 2.0)

            # Clic en botón Siguiente
            submit_btn = page.locator(SEL_SUBMIT_BTN)
            if submit_btn.count() == 0:
                self.guard.trigger_kill_switch("Botón de siguiente no encontrado tras ingresar email")
            submit_btn.first.click()
            self.human_delay(2.5, 4.0)

            # --- PASO 2: Verificación de Desafíos Intermedios / Pantalla Contraseña ---
            self._check_danger_screens(page, "Tras ingresar email")

            pwd_field = page.locator(SEL_PASSWORD_INPUT)
            try:
                pwd_field.first.wait_for(state="visible", timeout=10000)
            except Exception:
                # Si no apareció el campo de contraseña, probablemente Microsoft pidió MFA o reportó error
                body_text = page.locator("body").inner_text()[:400]
                self.guard.trigger_kill_switch(
                    "Desafío no soportado o campo de contraseña no disponible",
                    body_text
                )

            # --- PASO 3: Ingreso de Contraseña ---
            self.human_delay(0.8, 1.5)
            pwd_field.first.click()
            self.type_realistic(pwd_field.first, password)
            self.human_delay(1.0, 2.0)

            # Clic en Iniciar Sesión
            submit_btn = page.locator(SEL_SUBMIT_BTN)
            submit_btn.first.click()
            self.human_delay(3.0, 5.0)

            # --- PASO 4: Post-Password Check (MFA / Error de clave / Bloqueo) ---
            self._check_danger_screens(page, "Tras ingresar contraseña")

            # --- PASO 5: Pantalla '¿Quiere mantener la sesión iniciada?' (KMSI) ---
            kmsi_btn = page.locator(SEL_KMSI_CHECK)
            if kmsi_btn.count() > 0 and kmsi_btn.first.is_visible():
                logger.info("Detectada pantalla 'Mantener sesión iniciada'. Aceptando...")
                self.human_delay(1.0, 2.0)
                kmsi_btn.first.click()
                self.human_delay(3.0, 5.0)

            # --- PASO 6: Validación Final de Redirección ---
            try:
                page.wait_for_url(lambda u: "sharepoint.com" in u and MS_LOGIN_URL_PREFIX not in u, timeout=25000)
            except Exception:
                pass

            if "sharepoint.com" in page.url and MS_LOGIN_URL_PREFIX not in page.url:
                logger.info("Login exitoso en SharePoint. Guardando nuevo StorageState...")
                self.save_session_state(context)
            else:
                self.guard.trigger_kill_switch(
                    "Fallo de redirección final a SharePoint",
                    f"URL actual no esperada: {page.url}"
                )

        except KillSwitchException:
            raise
        except Exception as e:
            self.guard.trigger_kill_switch("Error crítico no contemplado durante el login", str(e))

    def _check_danger_screens(self, page: Page, checkpoint: str):
        """
        Inspecciona el DOM en busca de cualquier pantalla de MFA, captcha,
        reset de clave o error de cuenta. Si encuentra alguna, dispara Kill Switch.
        """
        for selector in DANGER_SELECTORS:
            loc = page.locator(selector)
            if loc.count() > 0 and loc.first.is_visible():
                try:
                    text_detail = loc.first.inner_text()[:200]
                except Exception:
                    text_detail = "Elemento visible detectado"
                self.guard.trigger_kill_switch(
                    f"Validación de seguridad imprevista de Microsoft ({checkpoint})",
                    f"Selector: {selector} | Detalle: {text_detail}"
                )
