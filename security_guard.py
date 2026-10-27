"""
Módulo de Seguridad Anti-Bloqueo (Kill Switch) para Hopper.
Garantiza que al primer síntoma de validación anómala o error en Microsoft 365,
el scraper se detenga inmediatamente sin reintentos y deshabilite ejecuciones futuras.
"""
import os
import json
import logging
from datetime import datetime
from typing import Optional, Tuple
from sharepoint_config import KILL_SWITCH_DB_KEY

logger = logging.getLogger("HopperSecurityGuard")
LOCK_FILE_FALLBACK = ".hopper_sp_lock"

class KillSwitchException(Exception):
    """Excepción irrecuperable que corta la ejecución del scraper en 0ms."""
    pass

class HopperSecurityGuard:
    def __init__(self, supabase_client=None, cipher_suite=None):
        self.supabase = supabase_client
        self.cipher = cipher_suite

    def is_locked(self) -> Tuple[bool, Optional[str]]:
        """
        Verifica si el sistema se encuentra bloqueado preventivamente.
        Consulta primero Supabase (tabla credenciales_sitios) y luego fallback local.
        Retorna: (esta_bloqueado: bool, motivo: str)
        """
        # 1. Chequeo en Supabase
        if self.supabase:
            try:
                res = self.supabase.table("credenciales_sitios").select("*").eq("sitio", KILL_SWITCH_DB_KEY).execute()
                if res.data:
                    row = res.data[0]
                    status = row.get("user_enc", "active")
                    if status == "locked":
                        raw_reason = row.get("pass_enc", "")
                        # Intentar desencriptar el motivo si está cifrado
                        if self.cipher and raw_reason:
                            try:
                                motivo = self.cipher.decrypt(raw_reason.encode()).decode()
                            except Exception:
                                motivo = raw_reason
                        else:
                            motivo = raw_reason or "Bloqueo preventivo de seguridad activado."
                        return True, motivo
            except Exception as e:
                logger.warning(f"Error consultando Kill Switch en Supabase: {e}")

        # 2. Fallback de archivo local (si no hay red o reinició)
        if os.path.exists(LOCK_FILE_FALLBACK):
            try:
                with open(LOCK_FILE_FALLBACK, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if data.get("status") == "locked":
                        return True, data.get("reason", "Bloqueo local preventivo.")
            except Exception:
                return True, "Bloqueo local detectado en archivo testigo."

        return False, None

    def trigger_kill_switch(self, reason: str, details: str = ""):
        """
        Activa el Kill Switch:
        1. Persiste el estado 'locked' en Supabase (tabla credenciales_sitios).
        2. Guarda el archivo testigo local.
        3. Lanza KillSwitchException sin permitir reintentos.
        """
        timestamp = datetime.now().isoformat()
        motivo_completo = f"[{timestamp}] {reason}"
        if details:
            motivo_completo += f" | Detalle: {details[:300]}"

        logger.critical(f"🛑 KILL SWITCH ACTIVADO: {motivo_completo}")

        # 1. Guardar archivo local preventivo
        try:
            with open(LOCK_FILE_FALLBACK, "w", encoding="utf-8") as f:
                json.dump({
                    "status": "locked",
                    "reason": motivo_completo,
                    "timestamp": timestamp
                }, f)
        except Exception as e_file:
            logger.error(f"Error al escribir lock file local: {e_file}")

        # 2. Guardar en Supabase (credenciales_sitios)
        if self.supabase:
            try:
                reason_to_save = motivo_completo
                if self.cipher:
                    try:
                        reason_to_save = self.cipher.encrypt(motivo_completo.encode()).decode()
                    except Exception:
                        pass

                res = self.supabase.table("credenciales_sitios").select("id").eq("sitio", KILL_SWITCH_DB_KEY).execute()
                payload = {
                    "usuario_app": "system_guard",
                    "sitio": KILL_SWITCH_DB_KEY,
                    "user_enc": "locked",
                    "pass_enc": reason_to_save
                }
                if res.data:
                    self.supabase.table("credenciales_sitios").update(payload).eq("sitio", KILL_SWITCH_DB_KEY).execute()
                else:
                    self.supabase.table("credenciales_sitios").insert(payload).execute()
                logger.info("Kill switch persistido en Supabase.")
            except Exception as e_db:
                logger.error(f"Error al actualizar Kill Switch en Supabase: {e_db}")

        # 3. Interrupción inmediata
        raise KillSwitchException(motivo_completo)

    def unlock_system(self, authorized_user: str) -> bool:
        """
        Desbloquea el sistema manualmente tras verificación del administrador.
        Restaura el estado a 'active' en Supabase y remueve el lock local.
        """
        # 1. Remover lock local
        if os.path.exists(LOCK_FILE_FALLBACK):
            try:
                os.remove(LOCK_FILE_FALLBACK)
            except Exception as e:
                logger.warning(f"No se pudo eliminar el lock file local: {e}")

        # 2. Actualizar Supabase
        if self.supabase:
            try:
                res = self.supabase.table("credenciales_sitios").select("id").eq("sitio", KILL_SWITCH_DB_KEY).execute()
                payload = {
                    "usuario_app": authorized_user,
                    "sitio": KILL_SWITCH_DB_KEY,
                    "user_enc": "active",
                    "pass_enc": f"Desbloqueado por {authorized_user} el {datetime.now().isoformat()}"
                }
                if res.data:
                    self.supabase.table("credenciales_sitios").update(payload).eq("sitio", KILL_SWITCH_DB_KEY).execute()
                else:
                    self.supabase.table("credenciales_sitios").insert(payload).execute()
                logger.info(f"Kill switch desactivado en Supabase por {authorized_user}.")
                return True
            except Exception as e:
                logger.error(f"Error al desbloquear en Supabase: {e}")
                return False

        return True
