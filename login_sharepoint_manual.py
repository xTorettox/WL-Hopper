"""
Script de Login Interactivo Asistido para SharePoint (Sullair Argentina).
Permite al usuario autenticarse interactivamente (headless=False) resolviendo
MFA o cualquier validación de Microsoft de forma segura.
Genera el 'storageState' válido, lo encripta y lo sincroniza con Supabase,
desbloqueando el Kill Switch automáticamente.
"""
import os
import sys
import time
import logging
from playwright.sync_api import sync_playwright

from sharepoint_config import (
    SHAREPOINT_BASE_URL,
    SHAREPOINT_SITE_URL,
    STORAGE_STATE_PATH,
    STORAGE_STATE_DB_KEY,
    KILL_SWITCH_DB_KEY,
    CREDENTIALS_DB_KEY
)
from security_guard import HopperSecurityGuard
from auth_manager import SharePointAuthManager

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("LoginManual")

def inicializar_servicios():
    """Inicializa Supabase y Fernet cargando secretos de Streamlit."""
    try:
        import streamlit as st
        # Forzar carga de secretos de streamlit si existe
        supabase_url = st.secrets.get("SUPABASE_URL")
        supabase_key = st.secrets.get("SUPABASE_KEY")
        fernet_key = st.secrets.get("FERNET_KEY")

        supabase_client = None
        cipher_suite = None

        if supabase_url and supabase_key:
            from supabase import create_client
            supabase_client = create_client(supabase_url, supabase_key)

        if fernet_key:
            from cryptography.fernet import Fernet
            cipher_suite = Fernet(fernet_key.encode())

        return supabase_client, cipher_suite
    except Exception as e:
        logger.warning(f"No se pudieron cargar secretos de streamlit: {e}")
        return None, None

def ejecutar_login_manual():
    print("=" * 70)
    print("🚀 HOPPER SHAREPOINT - SCRIPT DE AUTENTICACIÓN ASISTIDA")
    print("=" * 70)
    print("Este script abrirá un navegador visible para que puedas iniciar sesión")
    print("corporativa en Microsoft 365 y validar MFA/Authenticator si es requerido.")
    print("Una vez dentro de SharePoint, la sesión se guardará y encriptará")
    print("para que el scraper automático en Streamlit funcione sin interrupciones.")
    print("=" * 70)

    supabase_client, cipher_suite = inicializar_servicios()
    guard = HopperSecurityGuard(supabase_client, cipher_suite)
    auth_mgr = SharePointAuthManager(guard, cipher_suite, supabase_client)

    usuario_sugerido = ""
    clave_sugerida = ""
    try:
        u, p = auth_mgr.get_credentials("manual_runner")
        usuario_sugerido = u
        clave_sugerida = p
    except Exception:
        pass

    with sync_playwright() as pw:
        print("\n⏳ Lanzando navegador en modo visible...")
        browser = pw.chromium.launch(
            headless=False,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--start-maximized"
            ]
        )
        context = browser.new_context(
            viewport=None,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        )
        context.add_init_script("delete Object.getPrototypeOf(navigator).webdriver")
        page = context.new_page()

        print(f"🔗 Navegando a SharePoint: {SHAREPOINT_SITE_URL}")
        page.goto(SHAREPOINT_SITE_URL)

        # Si tenemos usuario sugerido, ayudamos a precargar el input si aparece
        try:
            page.wait_for_selector('input[type="email"]', timeout=4000)
            if usuario_sugerido and page.locator('input[type="email"]').is_visible():
                print(f"💡 Sugiriendo usuario: {usuario_sugerido}")
                page.fill('input[type="email"]', usuario_sugerido)
        except Exception:
            pass

        print("\n👉 Por favor, completá el inicio de sesión y el desafío MFA en el navegador...")
        print("Esperando acceso a 'BACK OFFICE ARG' de SharePoint (tiempo límite: 4 minutos)...\n")

        # Esperar a que la URL regrese a sharepoint.com y contenga el sitio
        try:
            page.wait_for_url(
                lambda url: "tefsullairargentina.sharepoint.com" in url and "login.microsoftonline.com" not in url,
                timeout=240000
            )
            print("⏳ Acceso detectado. Esperando estabilización de cookies y red...")
            time.sleep(5)

            # Guardar el estado
            auth_mgr.save_session_state(context)

            # Desbloquear el Kill Switch preventivo si estaba activo
            guard.unlock_system(authorized_user="fcendra_login_asistido")

            print("\n" + "=" * 70)
            print("🎉 ¡ÉXITO TOTAL! SESIÓN DE SHAREPOINT GUARDADA Y SINCRONIZADA")
            print(f"Archivo generado: {STORAGE_STATE_PATH}")
            if supabase_client:
                print("✅ Respaldo cifrado actualizado en Supabase (tabla 'credenciales_sitios').")
                print("✅ Kill Switch desactivado. El sistema está 100% OPERATIVO.")
            print("=" * 70)
            print("\nYa podés cerrar esta ventana y usar Hopper con normalidad en Streamlit.\n")

        except Exception as ex:
            print(f"\n❌ Tiempo de espera agotado o error: {ex}")
        finally:
            context.close()
            browser.close()

if __name__ == "__main__":
    ejecutar_login_manual()
