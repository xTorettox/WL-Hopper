"""
Componente visual de Streamlit para el Panel de Administrador y Alerta Crítica del Kill Switch.
"""
import streamlit as st
from security_guard import HopperSecurityGuard

def render_kill_switch_banner(security_guard: HopperSecurityGuard, current_user: str) -> bool:
    """
    Renderiza la alerta crítica si el Kill Switch se encuentra activado.
    Retorna True si el sistema está bloqueado, False si está activo.
    """
    is_locked, motivo = security_guard.is_locked()

    if is_locked:
        st.error(f"""
        ### 🚨 ALERTA CRÍTICA: HOPPER SHAREPOINT BLOQUEADO (KILL SWITCH PREVENTIVO)
        
        **Motivo del bloqueo:** `{motivo}`
        
        *Las ejecuciones automáticas de SharePoint han sido suspendidas para blindar y proteger la cuenta de correo corporativa de Microsoft 365 contra bloqueos por contraseña o desafíos imprevistos.*
        """)

        # Si el usuario es el administrador
        if current_user in ["fcendra", "admin"] or current_user.startswith("fcendra"):
            with st.expander("🛠️ Panel de Recuperación del Administrador", expanded=True):
                st.markdown("""
                **Pasos recomendados para reanudar el servicio:**
                1. Si Microsoft solicitó MFA o cambio de clave, ejecutá en tu terminal local:
                   ```bash
                   python login_sharepoint_manual.py
                   ```
                   *Esto abrirá una ventana para que inicies sesión tú mismo y se sincronizará la nueva sesión cifrada automáticamente.*
                2. Si fue una falsa alarma o ya validaste tu cuenta en Microsoft, podés reanudar el sistema aquí abajo:
                """)
                
                col1, col2 = st.columns([2, 1])
                with col1:
                    confirma = st.checkbox("Confirmo que la cuenta está segura y deseo desbloquear", key="chk_unlock_sp_banner")
                with col2:
                    if st.button("🔓 Desbloquear Hopper", disabled=not confirma, type="primary", use_container_width=True):
                        if security_guard.unlock_system(authorized_user=current_user):
                            st.success("✅ Hopper SharePoint desbloqueado correctamente.")
                            st.rerun()
                        else:
                            st.error("❌ No se pudo actualizar el estado en Supabase.")

        return True

    return False

import pandas as pd
from user_profiles import (
    listar_perfiles,
    guardar_perfil_usuario,
    eliminar_perfil_usuario,
    cargar_perfiles
)

def render_user_management(supabase_client=None):
    """
    Renderiza el editor y gestor integral de usuarios y permisos (RBAC) para el Administrador.
    """
    st.title("👥 Gestión de Usuarios y Permisos - WL Hopper")
    st.markdown("Administrá qué usuarios tienen acceso al sistema y qué funciones específicas pueden ejecutar.")

    col_nav1, col_nav2 = st.columns([1, 4])
    with col_nav1:
        if st.button("⬅️ Volver a la App", type="primary", use_container_width=True):
            st.session_state.show_user_editor = False
            st.rerun()

    # Cargar los perfiles más recientes
    perfiles = cargar_perfiles(supabase_client)

    # Métricas superiores
    total_usuarios = len(perfiles)
    total_ingresos = sum(1 for p in perfiles.values() if p.get("rol") == "INGRESOS")
    total_admins = sum(1 for p in perfiles.values() if p.get("rol") == "ADMIN")

    m1, m2, m3 = st.columns(3)
    m1.metric("👥 Usuarios Registrados", total_usuarios)
    m2.metric("📋 Perfil Ingresos", total_ingresos)
    m3.metric("🛡️ Administradores", total_admins)

    st.markdown("---")

    tab_lista, tab_editar = st.tabs(["📋 Directorio de Usuarios", "✏️ Asignar Permisos / Crear Usuario"])

    # --- TAB 1: LISTA ---
    with tab_lista:
        filas = []
        for u, d in perfiles.items():
            filas.append({
                "Usuario": u,
                "Nombre": d.get("nombre", "-"),
                "Email": d.get("email", "-"),
                "Rol": d.get("rol", "ESTANDAR"),
                "Doc. Completa (SharePoint)": "🟢 Habilitado" if d.get("permiso_documentacion_completa") else "⚪ No",
                "Dashboard": "🟢 Habilitado" if d.get("permiso_dashboard") else "⚪ No",
                "Gestión Credenciales": "🟢 Habilitado" if d.get("permiso_gestion_credenciales") else "⚪ No",
            })
        df_users = pd.DataFrame(filas)
        st.dataframe(df_users, use_container_width=True, hide_index=True)

    # --- TAB 2: EDITOR / ASIGNADOR ---
    with tab_editar:
        col_sel, col_form = st.columns([1, 2], gap="large")

        with col_sel:
            st.markdown("##### Seleccionar Usuario")
            opciones_usuarios = list(perfiles.keys())
            opciones_display = opciones_usuarios + ["➕ Crear Nuevo Usuario..."]
            
            # Guardar selección en sesión
            sel_usuario = st.selectbox(
                "Elegir cuenta para editar permisos:",
                opciones_display,
                key="admin_user_select"
            )
            is_new = (sel_usuario == "➕ Crear Nuevo Usuario...")

        with col_form:
            with st.container(border=True):
                if is_new:
                    st.markdown("##### ➕ Crear Nuevo Usuario")
                    u_key = st.text_input("Identificador de Usuario (Username):", key="new_u_key", placeholder="ej: jdoe, ingresosbue, etc.").strip().lower()
                    u_nombre = st.text_input("Nombre y Apellido:", key="new_u_nombre", placeholder="ej: Juan Pérez")
                    u_email = st.text_input("Correo Corporativo:", key="new_u_email", placeholder="ej: jperez@sullair.com.ar")
                    u_rol = st.selectbox("Rol Base:", ["INGRESOS", "ESTANDAR", "ADMIN"], index=0, key="new_u_rol")
                    
                    st.markdown("###### Permisos Específicos:")
                    def_doc = (u_rol in ["INGRESOS", "ADMIN"])
                    def_dash = (u_rol == "ADMIN")
                    def_cred = True
                    
                    p_doc = st.checkbox("📄 Descargar Documentación Completa (Título, Cédula y Seguros en SharePoint)", value=def_doc, key="new_p_doc")
                    p_dash = st.checkbox("📊 Acceso al Dashboard de Métricas y KPIs", value=def_dash, key="new_p_dash")
                    p_cred = st.checkbox("🔑 Configuración de Credenciales de Sitios", value=def_cred, key="new_p_cred")

                    if st.button("💾 Guardar y Crear Usuario", type="primary", use_container_width=True):
                        if not u_key:
                            st.error("Debes ingresar un nombre de usuario.")
                        else:
                            nuevo_perfil = {
                                "nombre": u_nombre or u_key,
                                "email": u_email or "",
                                "rol": u_rol,
                                "permiso_documentacion_completa": p_doc,
                                "permiso_dashboard": p_dash,
                                "permiso_gestion_credenciales": p_cred,
                            }
                            if guardar_perfil_usuario(u_key, nuevo_perfil, supabase_client):
                                st.success(f"✅ Usuario '{u_key}' creado con éxito.")
                                st.rerun()
                            else:
                                st.error("❌ Error guardando el perfil.")
                else:
                    perfil_actual = perfiles[sel_usuario]
                    st.markdown(f"##### ✏️ Editando: `{sel_usuario}`")
                    st.text_input("Usuario (Username):", value=sel_usuario, disabled=True)
                    e_nombre = st.text_input("Nombre y Apellido:", value=perfil_actual.get("nombre", ""), key=f"e_nombre_{sel_usuario}")
                    e_email = st.text_input("Correo Corporativo:", value=perfil_actual.get("email", ""), key=f"e_email_{sel_usuario}")
                    
                    roles = ["INGRESOS", "ADMIN", "ESTANDAR"]
                    rol_idx = roles.index(perfil_actual.get("rol", "ESTANDAR")) if perfil_actual.get("rol", "ESTANDAR") in roles else 0
                    e_rol = st.selectbox("Rol:", roles, index=rol_idx, key=f"e_rol_{sel_usuario}")

                    st.markdown("###### Configuración de Permisos:")
                    e_doc = st.checkbox(
                        "📄 Descargar Documentación Completa (Título, Cédula y Seguros en SharePoint)",
                        value=perfil_actual.get("permiso_documentacion_completa", False),
                        key=f"e_doc_{sel_usuario}"
                    )
                    e_dash = st.checkbox(
                        "📊 Acceso al Dashboard de Métricas y KPIs",
                        value=perfil_actual.get("permiso_dashboard", False),
                        key=f"e_dash_{sel_usuario}"
                    )
                    e_cred = st.checkbox(
                        "🔑 Configuración de Credenciales de Sitios",
                        value=perfil_actual.get("permiso_gestion_credenciales", True),
                        key=f"e_cred_{sel_usuario}"
                    )

                    c_act1, c_act2 = st.columns(2)
                    with c_act1:
                        if st.button("💾 Guardar Permisos", type="primary", use_container_width=True, key=f"btn_save_{sel_usuario}"):
                            actualizado = {
                                "nombre": e_nombre,
                                "email": e_email,
                                "rol": e_rol,
                                "permiso_documentacion_completa": e_doc,
                                "permiso_dashboard": e_dash,
                                "permiso_gestion_credenciales": e_cred,
                            }
                            if guardar_perfil_usuario(sel_usuario, actualizado, supabase_client):
                                st.success("✅ Permisos actualizados correctamente.")
                                st.rerun()
                            else:
                                st.error("❌ Error guardando los cambios.")

                    with c_act2:
                        es_admin_fijo = sel_usuario in ["fcendra", "admin"]
                        if not es_admin_fijo:
                            if st.button("🗑️ Eliminar Usuario", type="secondary", use_container_width=True, key=f"btn_del_{sel_usuario}"):
                                if eliminar_perfil_usuario(sel_usuario, supabase_client):
                                    st.warning(f"Usuario '{sel_usuario}' eliminado.")
                                    st.rerun()
                                else:
                                    st.error("No se pudo eliminar el usuario.")
                        else:
                            st.caption("🔒 *Cuenta de administrador protegida.*")

        st.info("💡 **Aviso sobre contraseñas:** Recordá que si creás un nuevo usuario para Streamlit Cloud, debés registrar su contraseña en la sección `[passwords]` de los Secrets de Streamlit para que pueda ingresar al sistema.")
