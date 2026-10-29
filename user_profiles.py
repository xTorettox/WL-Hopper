"""
Módulo de Gestión de Perfiles y Control de Acceso (RBAC) para WL Hopper.
Define los roles, permisos y usuarios habilitados para el sistema, con
soporte de sincronización persistente en Supabase y disco local.
"""

import json
import os
from typing import Dict, Any, Optional

PROFILES_FILE = "user_profiles.json"
PROFILES_DB_KEY = "USER_PROFILES_CONFIG"

PERFILES_USUARIOS: Dict[str, Dict[str, Any]] = {
    "fcendra": {
        "nombre": "Federico Cendra",
        "email": "fcendra@sullair.com.ar",
        "rol": "ADMIN",
        "permiso_documentacion_completa": True,
        "permiso_dashboard": True,
        "permiso_gestion_credenciales": True,
    },
    "admin": {
        "nombre": "Administrador",
        "email": "admin@sullair.com.ar",
        "rol": "ADMIN",
        "permiso_documentacion_completa": True,
        "permiso_dashboard": True,
        "permiso_gestion_credenciales": True,
    },
    "ingnqn": {
        "nombre": "Ingresos Neuquén",
        "email": "ingresosnqn@sullair.com.ar",
        "rol": "INGRESOS",
        "permiso_documentacion_completa": True,
        "permiso_dashboard": False,
        "permiso_gestion_credenciales": True,
    },
    "usrtest": {
        "nombre": "Usuario Tester",
        "email": "tester@sullair.com.ar",
        "rol": "INGRESOS",
        "permiso_documentacion_completa": True,
        "permiso_dashboard": False,
        "permiso_gestion_credenciales": True,
    },
}

PERFIL_DEFECTO: Dict[str, Any] = {
    "nombre": "Usuario Estándar",
    "email": "",
    "rol": "ESTANDAR",
    "permiso_documentacion_completa": False,
    "permiso_dashboard": False,
    "permiso_gestion_credenciales": True,
}

def cargar_perfiles(supabase_client=None) -> Dict[str, Dict[str, Any]]:
    """
    Carga perfiles desde Supabase o archivo local, sincronizando el estado global.
    """
    global PERFILES_USUARIOS

    # 1. Intentar desde Supabase
    if supabase_client:
        try:
            res = supabase_client.table("credenciales_sitios").select("*").eq("sitio", PROFILES_DB_KEY).execute()
            if res.data:
                raw_json = res.data[0].get("pass_enc", "")
                if raw_json:
                    data = json.loads(raw_json)
                    if isinstance(data, dict):
                        PERFILES_USUARIOS.update(data)
                        return PERFILES_USUARIOS
        except Exception:
            pass

    # 2. Intentar desde archivo local
    if os.path.exists(PROFILES_FILE):
        try:
            with open(PROFILES_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    PERFILES_USUARIOS.update(data)
                    return PERFILES_USUARIOS
        except Exception:
            pass

    return PERFILES_USUARIOS

def guardar_perfil_usuario(usuario: str, datos: Dict[str, Any], supabase_client=None) -> bool:
    """
    Guarda o actualiza un perfil de usuario en memoria, archivo local y Supabase.
    """
    global PERFILES_USUARIOS
    u = usuario.strip().lower()
    PERFILES_USUARIOS[u] = datos

    # Guardar en archivo local
    try:
        with open(PROFILES_FILE, "w", encoding="utf-8") as f:
            json.dump(PERFILES_USUARIOS, f, indent=2, ensure_ascii=False)
    except Exception:
        pass

    # Guardar en Supabase
    if supabase_client:
        try:
            raw_json = json.dumps(PERFILES_USUARIOS, ensure_ascii=False)
            payload = {
                "usuario_app": "system_rbac",
                "sitio": PROFILES_DB_KEY,
                "user_enc": "profiles_sync",
                "pass_enc": raw_json
            }
            res = supabase_client.table("credenciales_sitios").select("id").eq("sitio", PROFILES_DB_KEY).execute()
            if res.data:
                supabase_client.table("credenciales_sitios").update(payload).eq("sitio", PROFILES_DB_KEY).execute()
            else:
                supabase_client.table("credenciales_sitios").insert(payload).execute()
            return True
        except Exception as e:
            print(f"Error guardando perfiles en Supabase: {e}")
            return False

    return True

def eliminar_perfil_usuario(usuario: str, supabase_client=None) -> bool:
    """
    Elimina un usuario del registro de perfiles.
    Protege cuentas maestras (fcendra y admin).
    """
    global PERFILES_USUARIOS
    u = usuario.strip().lower()
    if u in ["fcendra", "admin"]:
        return False
    
    if u in PERFILES_USUARIOS:
        del PERFILES_USUARIOS[u]

    # Guardar en archivo local
    try:
        with open(PROFILES_FILE, "w", encoding="utf-8") as f:
            json.dump(PERFILES_USUARIOS, f, indent=2, ensure_ascii=False)
    except Exception:
        pass

    # Guardar en Supabase
    if supabase_client:
        try:
            raw_json = json.dumps(PERFILES_USUARIOS, ensure_ascii=False)
            payload = {
                "usuario_app": "system_rbac",
                "sitio": PROFILES_DB_KEY,
                "user_enc": "profiles_sync",
                "pass_enc": raw_json
            }
            supabase_client.table("credenciales_sitios").update(payload).eq("sitio", PROFILES_DB_KEY).execute()
            return True
        except Exception:
            return False
            
    return True

def listar_perfiles() -> Dict[str, Dict[str, Any]]:
    """Devuelve el diccionario completo de perfiles de usuario."""
    return PERFILES_USUARIOS.copy()

def obtener_perfil(usuario: str) -> Dict[str, Any]:
    """Obtiene el perfil configurado para el usuario o asigna el perfil estándar por defecto."""
    if not usuario:
        return PERFIL_DEFECTO.copy()
    u = usuario.strip().lower()
    if u in PERFILES_USUARIOS:
        return PERFILES_USUARIOS[u].copy()
    
    # Heurística flexible para nuevos usuarios que contengan roles en su nombre
    if "admin" in u:
        p = PERFILES_USUARIOS.get("admin", PERFIL_DEFECTO).copy()
        p["nombre"] = usuario
        return p
    if "ingresos" in u or "ingnqn" in u:
        p = PERFILES_USUARIOS.get("ingnqn", PERFIL_DEFECTO).copy()
        p["nombre"] = usuario
        return p
        
    p = PERFIL_DEFECTO.copy()
    p["nombre"] = usuario
    return p

def tiene_permiso_documentacion_completa(usuario: str) -> bool:
    """Verifica si el usuario tiene permiso para descargar documentación completa (Título, Cédula, Seguros)."""
    return obtener_perfil(usuario).get("permiso_documentacion_completa", False)

def tiene_permiso_dashboard(usuario: str) -> bool:
    """Verifica si el usuario tiene permiso para ver el dashboard de métricas."""
    return obtener_perfil(usuario).get("permiso_dashboard", False)
