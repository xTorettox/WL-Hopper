"""
SullApps Suite - Shared Python UI Helper Module
Sullair Argentina S.A.

Provides standardized enterprise styling, live ecosystem sync (`api/ecosystem.php`),
contextual loading spinners powered by `shared/quotes/loading_quotes.json`,
celebration notifications, iframe user-activity heartbeats (for the 2-hour idle timeout),
and unified Single Sign-On (SSO) verification across all SullApps.
"""

import os
import json
import html
import hmac
import hashlib
import time
import random
import urllib.parse
import urllib.request
from contextlib import contextmanager
import streamlit as st
import streamlit.components.v1 as components

SULLAIR_PRIMARY = "#006837"
SULLAIR_DARK = "#004D26"
SULLAIR_LIGHT = "#E6F4EA"
SULLAIR_BG = "#F8F9FA"
SULLAIR_CARD = "#FFFFFF"
SULLAIR_TEXT = "#1F2937"

ECOSYSTEM_API_URL = os.environ.get(
    "SULLAPPS_ECOSYSTEM_API_URL",
    "https://sullapps.fedev.com.ar/api/ecosystem.php",
)

DEFAULT_QUOTES = {
    "general": [
        "Cebando un mate mientras bajan los datos...",
        "Ajustando tuercas y revisando niveles...",
        "Validando que todo esté 10 impecable...",
    ],
    "scraping": [
        "Peleando con el login de Worklift (vamos ganando)...",
        "Buscando la póliza en el servidor...",
        "Cazando certificados (es temporada!)...",
    ],
    "reports": [
        "Armando el paquete de documentación flama 🔥...",
        "Firmando digitalmente y foliando hojas...",
        "Dejando todo prístino para salir...",
    ],
}


def _resolve_quotes_paths() -> list[str]:
    """Return candidate paths for loading_quotes.json in priority order."""
    here = os.path.dirname(os.path.abspath(__file__))
    cwd = os.getcwd()
    return [
        os.path.normpath(os.path.join(here, "quotes", "loading_quotes.json")),
        os.path.normpath(os.path.join(here, "..", "..", "shared", "quotes", "loading_quotes.json")),
        os.path.normpath(os.path.join(here, "..", "shared", "quotes", "loading_quotes.json")),
        os.path.normpath(os.path.join(cwd, "quotes", "loading_quotes.json")),
        os.path.normpath(os.path.join(cwd, "shared", "quotes", "loading_quotes.json")),
        os.path.normpath(os.path.join(here, "loading_quotes.json")),
    ]


def _resolve_env_paths() -> list[str]:
    """Return candidate paths for .env in local development."""
    here = os.path.dirname(os.path.abspath(__file__))
    cwd = os.getcwd()
    return [
        os.path.normpath(os.path.join(here, ".env")),
        os.path.normpath(os.path.join(here, "..", ".env")),
        os.path.normpath(os.path.join(here, "..", "..", ".env")),
        os.path.normpath(os.path.join(cwd, ".env")),
        os.path.normpath(os.path.join(cwd, "..", "..", ".env")),
    ]


def _get_sso_secret() -> str:
    """
    Resolve SSO HMAC secret from `st.secrets["SSO_SECRET"]`,
    `os.environ["SULLAPPS_SSO_SECRET"]`, or local `.env` file.
    """
    try:
        if hasattr(st, "secrets"):
            if "SSO_SECRET" in st.secrets and st.secrets["SSO_SECRET"]:
                return str(st.secrets["SSO_SECRET"]).strip()
            if "SULLAPPS_SSO_SECRET" in st.secrets and st.secrets["SULLAPPS_SSO_SECRET"]:
                return str(st.secrets["SULLAPPS_SSO_SECRET"]).strip()
    except Exception:
        pass

    env_val = os.environ.get("SULLAPPS_SSO_SECRET") or os.environ.get("SSO_SECRET")
    if env_val:
        return str(env_val).strip()

    for env_path in _resolve_env_paths():
        if os.path.isfile(env_path):
            try:
                with open(env_path, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if line.startswith("SULLAPPS_SSO_SECRET="):
                            return line.split("=", 1)[1].strip().strip("\"'")
            except Exception:
                continue

    return ""


@st.cache_data(ttl=300, show_spinner=False)
def fetch_ecosystem_config() -> dict:
    """
    Fetches live corporate theme, active apps catalog, and loading quotes from
    `https://sullapps.fedev.com.ar/api/ecosystem.php` (cached 300s).
    Falls back gracefully if offline.
    """
    try:
        req = urllib.request.Request(
            ECOSYSTEM_API_URL,
            headers={"User-Agent": "SullApps-Streamlit-Client/2.0"},
        )
        with urllib.request.urlopen(req, timeout=3.5) as resp:
            if resp.status == 200:
                payload = json.loads(resp.read().decode("utf-8"))
                if isinstance(payload, dict) and payload.get("status") == "ok":
                    return payload
    except Exception:
        pass
    return {}


def verify_sullapps_sso(max_age_seconds: int = 86400) -> dict | None:
    """
    Verifies the HMAC-SHA256 signed SSO parameters passed by the SullApps Hub
    (`sso_user`, `sso_role`, `sso_name`, `sso_ts`, `sso_sig`).
    Caches valid identity in `st.session_state['_sullapps_sso_payload']`.
    """
    cached = st.session_state.get("_sullapps_sso_payload")
    if isinstance(cached, dict) and cached.get("username"):
        st.session_state.pop("_sullapps_sso_error", None)
        return cached

    st.session_state.pop("_sullapps_sso_error", None)
    try:
        qp = st.query_params
        sso_user = str(qp.get("sso_user", "")).strip()
        sso_role = str(qp.get("sso_role", "")).strip()
        sso_name = str(qp.get("sso_name", "")).strip()
        sso_ts = str(qp.get("sso_ts", "")).strip()
        sso_sig = str(qp.get("sso_sig", "")).strip()
    except Exception:
        return None

    if not sso_user or not sso_ts or not sso_sig:
        return None

    try:
        ts_int = int(sso_ts)
        now = int(time.time())
        if abs(now - ts_int) > max_age_seconds:
            st.session_state["_sullapps_sso_error"] = "expired"
            return None
    except ValueError:
        st.session_state["_sullapps_sso_error"] = "bad_sig"
        return None

    secret = _get_sso_secret()
    if not secret:
        st.session_state["_sullapps_sso_error"] = "missing_secret"
        return None

    payload_formats = [
        f"{sso_user}|{sso_role}|{sso_ts}",
        f"{sso_user.lower()}|{sso_role.lower()}|{sso_ts}",
    ]
    sig_valid = False
    for pf in payload_formats:
        expected_sig = hmac.new(
            secret.encode("utf-8"),
            pf.encode("utf-8"),
            hashlib.sha256,
        ).hexdigest()
        if hmac.compare_digest(expected_sig, sso_sig):
            sig_valid = True
            break

    if not sig_valid:
        st.session_state["_sullapps_sso_error"] = "bad_sig"
        return None

    identity = {
        "username": sso_user,
        "role": sso_role or "user",
        "display_name": sso_name or sso_user,
    }
    st.session_state["_sullapps_sso_payload"] = identity
    return identity


def clear_sullapps_sso() -> None:
    """Clears cached SullApps SSO session state."""
    st.session_state.pop("_sullapps_sso_payload", None)


def render_sullapps_portal_notice(
    portal_url: str = "https://sullapps.fedev.com.ar/default.php",
    app_slug: str = "",
) -> None:
    """
    Bloquea el acceso directo no autenticado a los subdominios de Streamlit Cloud y ofrece
    retorno directo a la app dentro de SullApps Hub (`redirect=/apps/<slug>/`).
    Intenta redirección automática tras 4.5s (sujeta al sandbox de Streamlit Cloud) y provee
    el botón manual con `target="_blank"`.
    """
    if not isinstance(portal_url, str) or not portal_url.strip().startswith(("http://", "https://")):
        portal_url = "https://sullapps.fedev.com.ar/default.php"

    clean_slug = str(app_slug or "").strip().strip("/").lower()
    if clean_slug:
        qs = urllib.parse.urlencode({
            "login_required": "1",
            "redirect": f"/apps/{clean_slug}/",
            "app": clean_slug,
        })
    else:
        qs = urllib.parse.urlencode({"login_required": "1"})

    sep = "&" if "?" in portal_url else "?"
    target_url = f"{portal_url.rstrip('/')}{sep}{qs}"
    safe_url = html.escape(target_url)

    sso_err = st.session_state.get("_sullapps_sso_error")
    if sso_err == "missing_secret":
        st.warning(
            "⚠️ **Configuración pendiente en Streamlit Cloud:** falta cargar `SSO_SECRET` en "
            "**Settings → Secrets** de esta app (debe coincidir con `SULLAPPS_SSO_SECRET` del `.env` en Hostinger)."
        )
    elif sso_err == "bad_sig":
        st.warning(
            "⚠️ **Firma SSO inválida:** el valor `SSO_SECRET` en Streamlit Cloud no coincide con "
            "`SULLAPPS_SSO_SECRET` configurado en Hostinger."
        )
    elif sso_err == "expired":
        st.warning("⚠️ **Token SSO expirado:** volvé a ingresar desde SullApps Hub para renovar tu acceso.")

    st.markdown(
        f"""
        <div style="
            background: linear-gradient(135deg, #F0FDF4 0%, #DCFCE7 100%);
            border: 1.5px solid #86EFAC;
            border-left: 6px solid #006837;
            border-radius: 12px;
            padding: 18px 22px;
            margin-bottom: 18px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 14px;
            flex-wrap: wrap;
        ">
            <div style="font-size: 0.94rem; color: #064E3B; line-height: 1.5;">
                <strong>🔐 Acceso Exclusivo vía SullApps Hub:</strong><br>
                Por seguridad corporativa, debés ingresar desde el portal principal <b>SullApps</b> con tu clave única.
                <div style="font-size: 0.78rem; color: #047857; margin-top: 4px;">
                    Si el navegador bloquea la redirección automática desde Streamlit, presioná el botón para iniciar sesión y volver a esta app.
                </div>
            </div>
            <a href="{safe_url}" target="_blank" rel="noopener noreferrer" style="
                background: #006837;
                color: #FFFFFF;
                text-decoration: none;
                font-weight: 700;
                font-size: 0.88rem;
                padding: 10px 18px;
                border-radius: 8px;
                white-space: nowrap;
                box-shadow: 0 2px 6px rgba(0, 104, 55, 0.25);
            ">Ir a SullApps Hub →</a>
        </div>
        """,
        unsafe_allow_html=True,
    )

    try:
        js_url = json.dumps(target_url)
        components.html(
            f"""
            <script>
            setTimeout(function() {{
                try {{
                    if (window.top && window.top !== window.self) {{
                        window.top.location.replace({js_url});
                    }} else {{
                        window.location.replace({js_url});
                    }}
                }} catch (e) {{
                    // Streamlit Community Cloud sandbox bloquea navegación automática cross-origin;
                    // el usuario dispone del botón manual con target="_blank".
                }}
            }}, 4500);
            </script>
            """,
            height=0,
        )
    except Exception:
        pass
    st.stop()


def load_quotes_dict() -> dict[str, list[str]]:
    """
    Loads quotes from the central `/api/ecosystem.php` endpoint first (cached 300s),
    falling back to local `quotes/loading_quotes.json` and `DEFAULT_QUOTES` if offline.
    """
    live_cfg = fetch_ecosystem_config()
    live_quotes = live_cfg.get("quotes") if isinstance(live_cfg, dict) else None
    if isinstance(live_quotes, dict) and live_quotes:
        return live_quotes

    for candidate in _resolve_quotes_paths():
        if os.path.isfile(candidate):
            try:
                with open(candidate, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict) and data:
                    return data
            except Exception:
                continue
    return DEFAULT_QUOTES


def get_loading_quote(task_type: str = "default") -> str:
    """
    Returns a random contextual quote for the given `task_type`
    ('general', 'scraping', 'reports', or 'default').
    """
    quotes_map = load_quotes_dict()
    normalized = (task_type or "default").strip().lower()
    if normalized == "default":
        normalized = "general"

    bucket = quotes_map.get(normalized) or quotes_map.get("general") or DEFAULT_QUOTES["general"]
    if not isinstance(bucket, list) or not bucket:
        bucket = DEFAULT_QUOTES["general"]

    return str(random.choice(bucket))


@contextmanager
def sullapps_spinner(task_type: str = "default", custom_prefix: str = ""):
    """
    Context manager that reads `loading_quotes.json` (or live `/api/ecosystem.php`)
    and displays `st.spinner(...)` with a fun/relatable Sullair quote.
    """
    quote = get_loading_quote(task_type)
    display_text = f"{custom_prefix} — {quote}" if custom_prefix else quote
    with st.spinner(display_text):
        yield quote


def celebrate_success(message: str, subtitle: str = "", use_balloons: bool = True) -> None:
    """
    Triggers st.balloons() + st.toast() and renders a polished Sullair Enterprise success card.
    """
    if use_balloons:
        try:
            st.balloons()
        except Exception:
            pass

    try:
        st.toast(message, icon="🎉")
    except Exception:
        pass

    safe_msg = html.escape(str(message))
    safe_sub = html.escape(str(subtitle)) if subtitle else ""
    sub_html = f'<p style="margin: 6px 0 0 0; color: #065F46; font-size: 0.92rem;">{safe_sub}</p>' if safe_sub else ""

    st.markdown(
        f"""
        <div class="sullapps-celebration-card" style="
            background: linear-gradient(135deg, #ECFDF5 0%, #D1FAE5 100%);
            border: 1.5px solid #10B981;
            border-left: 6px solid #006837;
            border-radius: 12px;
            padding: 16px 20px;
            margin: 12px 0;
            box-shadow: 0 4px 14px rgba(0, 104, 55, 0.10);
            display: flex;
            align-items: center;
            gap: 14px;
        ">
            <div style="
                width: 42px;
                height: 42px;
                border-radius: 10px;
                background: #006837;
                color: #FFFFFF;
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 1.35rem;
                flex-shrink: 0;
            ">✅</div>
            <div style="flex-grow: 1;">
                <div style="font-weight: 800; font-size: 1.05rem; color: #064E3B; line-height: 1.3;">
                    {safe_msg}
                </div>
                {sub_html}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _inject_activity_heartbeat() -> None:
    """
    Injects a zero-height script that listens for user interactions inside the Streamlit
    iframe (`click`, `keydown`, `mousemove`, `scroll`, `touchstart`, throttled to 15s) and emits
    `window.parent.postMessage({ type: 'SULLAPPS_USER_ACTIVITY' }, '*')` to reset the parent Hub's
    2-hour idle session countdown.
    """
    try:
        components.html(
            """
            <script>
            (function() {
                var lastSent = 0;
                var THROTTLE_MS = 15000;
                function notifyParentActivity() {
                    var now = Date.now();
                    if ((now - lastSent) < THROTTLE_MS) return;
                    lastSent = now;
                    try {
                        if (window.parent && window.parent !== window) {
                            window.parent.postMessage({ type: 'SULLAPPS_USER_ACTIVITY', ts: now }, '*');
                        }
                        if (window.top && window.top !== window) {
                            window.top.postMessage({ type: 'SULLAPPS_USER_ACTIVITY', ts: now }, '*');
                        }
                    } catch (e) {}
                }
                notifyParentActivity();
                var targetDoc = document;
                try {
                    if (window.parent && window.parent.document) {
                        targetDoc = window.parent.document;
                    }
                } catch (e) {
                    targetDoc = document;
                }
                ['click', 'keydown', 'mousemove', 'scroll', 'touchstart'].forEach(function(evt) {
                    targetDoc.addEventListener(evt, notifyParentActivity, { passive: true });
                    document.addEventListener(evt, notifyParentActivity, { passive: true });
                });
            })();
            </script>
            """,
            height=0,
        )
    except Exception:
        pass


def inject_custom_css() -> None:
    """
    Polishes cards, buttons, tabs, tables, and form controls to match Sullair Enterprise style,
    hides redundant Streamlit default headers/footers, and activates the parent idle heartbeat.
    """
    _inject_activity_heartbeat()
    st.markdown(
        """
        <style>
        /* =====================================================================
           SULLAPPS SUITE - ENTERPRISE STREAMLIT THEME & UI POLISH
           Primary: #006837 | Surface: #FFFFFF | Background: #F8F9FA
           ===================================================================== */
        :root {
            --sullapps-green: #006837;
            --sullapps-green-hover: #004D26;
            --sullapps-green-soft: #E6F4EA;
            --sullapps-bg: #F8F9FA;
            --sullapps-surface: #FFFFFF;
            --sullapps-text: #1F2937;
            --sullapps-border: #E5E7EB;
        }

        /* Hide redundant Streamlit chrome (footer, deploy button, menu, decoration bar) */
        #MainMenu,
        footer,
        [data-testid="stFooter"],
        [data-testid="stDecoration"],
        .stDeployButton,
        [data-testid="stAppDeployButton"],
        [data-testid="stToolbarActions"] {
            display: none !important;
        }

        /* Keep header minimal & transparent while guaranteeing sidebar toggle stays visible & clickable on desktop & mobile */
        header[data-testid="stHeader"] {
            background: transparent !important;
            height: 2.85rem !important;
            pointer-events: none !important;
            z-index: 999990 !important;
        }
        header[data-testid="stHeader"] [data-testid="stToolbar"],
        [data-testid="collapsedControl"],
        [data-testid="stSidebarCollapsedControl"],
        [data-testid="stExpandSidebarButton"],
        [data-testid="stSidebarCollapseButton"] {
            pointer-events: auto !important;
            visibility: visible !important;
            opacity: 1 !important;
            z-index: 999999 !important;
        }
        [data-testid="collapsedControl"],
        [data-testid="stSidebarCollapsedControl"],
        [data-testid="stExpandSidebarButton"] {
            display: inline-flex !important;
            align-items: center !important;
            justify-content: center !important;
            position: fixed !important;
            top: 0.55rem !important;
            left: 0.65rem !important;
            background: #FFFFFF !important;
            border: 1.5px solid #006837 !important;
            border-radius: 10px !important;
            box-shadow: 0 3px 10px rgba(0, 104, 55, 0.22) !important;
            color: #006837 !important;
            min-width: 2.35rem !important;
            min-height: 2.35rem !important;
        }
        header[data-testid="stHeader"] button,
        [data-testid="collapsedControl"] button,
        [data-testid="stSidebarCollapsedControl"] button,
        [data-testid="stExpandSidebarButton"] button,
        [data-testid="stSidebarCollapseButton"] button {
            pointer-events: auto !important;
            visibility: visible !important;
            opacity: 1 !important;
            color: #006837 !important;
        }
        [data-testid="collapsedControl"] svg,
        [data-testid="stSidebarCollapsedControl"] svg,
        [data-testid="stExpandSidebarButton"] svg,
        [data-testid="stSidebarCollapseButton"] svg {
            fill: #006837 !important;
            color: #006837 !important;
        }

        /* Clean container spacing for embedded & full-screen views */
        .block-container {
            padding-top: 1.75rem !important;
            padding-bottom: 2.5rem !important;
        }
        @media (max-width: 768px) {
            .block-container {
                padding-top: 3.35rem !important;
                padding-left: 0.85rem !important;
                padding-right: 0.85rem !important;
                padding-bottom: 2.25rem !important;
            }
        }

        /* Primary Action Buttons */
        button[kind="primary"],
        button[data-testid="stBaseButton-primary"],
        button[data-testid="baseButton-primary"],
        .stFormSubmitButton > button[kind="primary"] {
            background: linear-gradient(135deg, #006837 0%, #00522B 100%) !important;
            border-color: #006837 !important;
            color: #FFFFFF !important;
            font-weight: 700 !important;
            border-radius: 10px !important;
            box-shadow: 0 3px 10px rgba(0, 104, 55, 0.18) !important;
            transition: all 0.18s ease !important;
        }

        button[kind="primary"]:hover,
        button[data-testid="stBaseButton-primary"]:hover,
        button[data-testid="baseButton-primary"]:hover {
            background: linear-gradient(135deg, #00522B 0%, #003D20 100%) !important;
            border-color: #004D26 !important;
            transform: translateY(-1px) !important;
            box-shadow: 0 6px 14px rgba(0, 104, 55, 0.25) !important;
        }

        /* Secondary Action Buttons */
        button[kind="secondary"],
        button[data-testid="stBaseButton-secondary"],
        button[data-testid="baseButton-secondary"] {
            border-radius: 10px !important;
            font-weight: 600 !important;
            border: 1px solid #D1D5DB !important;
            transition: all 0.18s ease !important;
        }

        button[kind="secondary"]:hover,
        button[data-testid="stBaseButton-secondary"]:hover,
        button[data-testid="baseButton-secondary"]:hover {
            border-color: #006837 !important;
            color: #006837 !important;
            background-color: #F0FDF4 !important;
        }

        /* Metric Cards, Expanders & Form Containers */
        [data-testid="stMetric"],
        [data-testid="stForm"],
        [data-testid="stExpander"] {
            background-color: #FFFFFF !important;
            border: 1px solid #E5E7EB !important;
            border-radius: 12px !important;
            box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04) !important;
        }

        [data-testid="stMetric"] {
            padding: 14px 18px !important;
        }

        /* Tabs Styling */
        .stTabs [data-baseweb="tab-list"] {
            gap: 8px;
            border-bottom: 2px solid #E5E7EB;
        }

        .stTabs [data-baseweb="tab"] {
            border-radius: 8px 8px 0 0 !important;
            padding: 8px 16px !important;
            font-weight: 600 !important;
        }

        .stTabs [aria-selected="true"] {
            color: #006837 !important;
            border-bottom-color: #006837 !important;
            background-color: #F0FDF4 !important;
        }

        /* Sidebar Polish */
        section[data-testid="stSidebar"] {
            background-color: #FFFFFF !important;
            border-right: 1px solid #E5E7EB !important;
        }

        /* Dataframes & Tables */
        [data-testid="stDataFrame"] {
            border-radius: 10px !important;
            overflow: hidden !important;
            border: 1px solid #E5E7EB !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
