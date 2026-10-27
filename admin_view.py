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
