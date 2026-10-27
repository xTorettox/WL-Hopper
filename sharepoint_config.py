"""
Módulo de Configuración Blindada para SharePoint Scraper (Hopper).
Regla de seguridad: Las rutas son constantes inmutables en backend.
El scraper NO acepta rutas dinámicas inyectadas desde la UI.
"""
from typing import Final

# URL BASE DEL TENANT Y SITIO DE SHAREPOINT (HARDCODEADO)
SHAREPOINT_BASE_URL: Final[str] = "https://tefsullairargentina.sharepoint.com"
SHAREPOINT_SITE_URL: Final[str] = "https://tefsullairargentina.sharepoint.com/sites/BACKOFFICEARG"

# RUTA RELATIVA DEL SERVIDOR A LA CARPETA OBJETIVO ESPECÍFICA
SHAREPOINT_SERVER_RELATIVE_ROOT: Final[str] = "/sites/BACKOFFICEARG/Documentos Compartidos"
SHAREPOINT_SERVER_RELATIVE_FOLDER: Final[str] = "/sites/BACKOFFICEARG/Documentos Compartidos/BACK OFFICE Sullair Arg"

# SUB-CARPETAS ESPECÍFICAS DE BÚSQUEDA DENTRO DE "BACK OFFICE Sullair Arg"
FOLDER_DOC_EQUIPOS: Final[str] = f"{SHAREPOINT_SERVER_RELATIVE_FOLDER}/DOCUMENTACION EQUIPOS"
FOLDER_SEGUROS: Final[str] = f"{SHAREPOINT_SERVER_RELATIVE_FOLDER}/SEGUROS"
FOLDER_CERTIFICACIONES: Final[str] = f"{SHAREPOINT_SERVER_RELATIVE_FOLDER}/CERTIFICACIONES"

# RUTA COMPLETA DE NAVEGACIÓN DIRECTA PARA PLAYWRIGHT
SHAREPOINT_TARGET_BROWSER_URL: Final[str] = (
    f"{SHAREPOINT_SITE_URL}/Documentos%20Compartidos/Forms/AllItems.aspx"
    f"?RootFolder=%2Fsites%2FBACKOFFICEARG%2FDocumentos%20Compartidos%2FBACK%20OFFICE%20Sullair%20Arg"
)

# PERSISTENCIA LOCAL Y CLAVES DE BASE DE DATOS (SUPABASE)
STORAGE_STATE_PATH: Final[str] = "session_sharepoint_state.json"
STORAGE_STATE_DB_KEY: Final[str] = "SP_STORAGE_STATE"
KILL_SWITCH_DB_KEY: Final[str] = "SP_KILL_SWITCH"
CREDENTIALS_DB_KEY: Final[str] = "SP_CRED"

# SELECTORES ESTRICTOS DE MICROSOFT ONLINE
MS_LOGIN_URL_PREFIX: Final[str] = "https://login.microsoftonline.com"
SEL_EMAIL_INPUT: Final[str] = 'input[type="email"], input[name="loginfmt"]'
SEL_PASSWORD_INPUT: Final[str] = 'input[type="password"], input[name="passwd"]'
SEL_SUBMIT_BTN: Final[str] = 'input[type="submit"], button[type="submit"]'
SEL_KMSI_CHECK: Final[str] = 'input[id="idSIButton9"], input[value="Sí"], input[value="Yes"]'

# SELECTORES DE PELIGRO / DESAFÍO (Disparan el Kill Switch inmediatamente sin reintentos)
DANGER_SELECTORS: Final[list[str]] = [
    'div#idDiv_SAOTCS_Title',              # Desafío MFA / Authenticator
    'input[name="otc"]',                   # Código OTP / SMS
    'div#idDiv_SAOTCC_Description',        # Aprobación en app móvil
    'input[id="ProofConfirmationOption"]', # Selección de método alternativo de verificación
    'div#serviceError',                    # Error de servicio o bloqueo de cuenta
    'div#passwordError',                   # Contraseña incorrecta
    'div#accountError',                    # Cuenta bloqueada o deshabilitada
    'iframe[title*="recaptcha"]',          # Captcha interactivo
    'div[id*="arkose"]',                   # Arkose Labs Bot Challenge
    'div#idA_PWD_ForgotPassword',         # Forzado de cambio / reseteo de clave
]
