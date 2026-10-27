"""
Módulo de Gestión de Perfiles y Control de Acceso (RBAC) para WL Hopper.
Define los roles, permisos y usuarios habilitados para el sistema.
"""

from typing import Dict, Any

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

def obtener_perfil(usuario: str) -> Dict[str, Any]:
    """Obtiene el perfil configurado para el usuario o asigna el perfil estándar por defecto."""
    if not usuario:
        return PERFIL_DEFECTO.copy()
    u = usuario.strip().lower()
    if u in PERFILES_USUARIOS:
        return PERFILES_USUARIOS[u].copy()
    
    # Heurística flexible para nuevos usuarios que contengan roles en su nombre
    if "admin" in u:
        p = PERFILES_USUARIOS["admin"].copy()
        p["nombre"] = usuario
        return p
    if "ingresos" in u or "ingnqn" in u:
        p = PERFILES_USUARIOS["ingnqn"].copy()
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
