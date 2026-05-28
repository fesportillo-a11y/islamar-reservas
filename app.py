import streamlit as st
import pandas as pd
from supabase import create_client, Client
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, date, timedelta
import calendar
import re
import bcrypt
import streamlit_authenticator as stauth

# ─────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="ESTEASUR 2015 - ISLAMAR",
    page_icon="🏖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

MESES = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
         "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]

FUENTES  = ["DIRECTA", "BOOKING.COM"]
ESTADOS  = ["", "PAGADO", "PENDIENTE", "SEÑAL PAGADA", "Pago mediante Booking.com", "EFECTIVO", "RESERVA ANULADA"]
DORMS    = ["1", "2", "3", "Estudio"]

APTOS = [
    # ── Apartamentos propios ──────────────────
    "APTO 2 - 1 DORM", "APTO 9 - 1 DORM", "APTO 10 - 1 DORM", "APTO 109 - 1 DORM",
    "APTO 201 - 1 DORM", "APTO 208 - 1 DORM", "APTO 209 - 1 DORM",
    "APTO 7 - 2 DORM", "APTO 14 - 2 DORM", "APTO 15 - 2 DORM",
    # ── Apartamentos JUANMA ───────────────────
    "APTO 215 - 2 DORM", "ESTUDIO 105", "ESTUDIO 216", "ESTUDIO 217",
]

APTOS_JUANMA = {"APTO 215 - 2 DORM", "ESTUDIO 105", "ESTUDIO 216", "ESTUDIO 217"}

# Apartamentos agrupados por tipo de dormitorio
APTOS_POR_TIPO = {
    "1":       [a for a in APTOS if "1 DORM"  in a],
    "2":       [a for a in APTOS if "2 DORM"  in a],
    "Estudio": [a for a in APTOS if "ESTUDIO" in a],
}

# ── Helpers de fecha y estado ─────────────────────────────────────────
_DATE_FMTS = ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y"]

def parse_date_safe(s) -> "date | None":
    """
    Parsea una fecha probando múltiples formatos (dd/mm/yyyy, yyyy-mm-dd, etc.).
    Devuelve None si no puede parsear — NUNCA lanza excepción.
    """
    txt = str(s).strip()
    for t in [txt, txt[:10]]:          # intenta cadena completa y solo los 10 primeros chars
        for fmt in _DATE_FMTS:
            try:
                return datetime.strptime(t, fmt).date()
            except Exception:
                pass
    return None

_ESTADOS_CANCELADOS = {
    "cancel", "anula", "no show", "no-show", "noshow",
    "cancelled", "canceled", "cancelled_by_guest", "cancelled_by_hotel",
    "guest_cancelled", "annul",
}

def parse_eur(v) -> "float | None":
    """Convierte '211,20 €' / '211.2 EUR' / '0' / 211.2 → float. None si no se puede."""
    if v is None:
        return None
    try:
        s = str(v).strip().replace("€", "").replace("EUR", "").replace(" ", "")
        if not s or s.lower() in ("nan", "none"):
            return None
        return float(s.replace(",", "."))
    except Exception:
        return None

def format_eur(v) -> str:
    """Convierte 211.2 → '211,20 €'. Devuelve '' si v es None/vacío/no numérico."""
    f = parse_eur(v)
    if f is None:
        return ""
    return f"{f:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") + " €"

def es_cancelada(estado_str: str) -> bool:
    """True si el estado indica que la reserva está cancelada o es un no-show."""
    t = str(estado_str).lower().strip()
    return any(x in t for x in _ESTADOS_CANCELADOS)

def clasificar_dormitorios(tipo_unidad_str: str) -> str:
    """
    Detecta el tipo de apartamento a partir del campo 'Tipo de unidad' de Booking.com.
    Devuelve: 'Estudio', '2' o '1'.
    NO usa el número de personas — solo el tipo de unidad.
    Maneja tanto valores en español como en inglés (Two-Bedroom Apartment, etc.).
    """
    t     = str(tipo_unidad_str).lower().strip()
    t_nor = t.replace("-", " ").replace("_", " ")   # normaliza guiones → espacios
    # Estudio / Studio / Loft
    if any(x in t_nor for x in ["estudio", "studio", "loft", "monoamb"]):
        return "Estudio"
    # Número explícito de dormitorios: "2 dormitorios", "2 bedroom", etc.
    m = re.search(r'(\d+)\s*(?:dorm|hab|bed|bdr|room)', t_nor)
    if m:
        return "2" if int(m.group(1)) >= 2 else "1"
    # Palabras escritas (inglés/español): "two-bedroom", "two bedroom", "dos dorm"…
    if any(x in t_nor for x in [
        "two bed", "two room", "two dorm",
        "dos dorm", "duplex", "dúplex", "2 dorm",
    ]):
        return "2"
    # Explícito 1 dormitorio en inglés (one-bedroom → "1")
    if any(x in t_nor for x in ["one bed", "one room", "one dorm", "1 dorm"]):
        return "1"
    return "1"  # defecto: 1 dormitorio

def dorm_desde_nombre_apto(nombre_apto: str) -> str:
    """Extrae el tipo de dormitorios directamente del nombre del apartamento."""
    n = nombre_apto.upper()
    if "ESTUDIO" in n: return "Estudio"
    if "2 DORM"  in n: return "2"
    if "1 DORM"  in n: return "1"
    return clasificar_dormitorios(nombre_apto)   # fallback genérico

def match_apto_directo(tipo_unidad_str: str) -> "str | None":
    """
    Intenta hacer match del valor 'Tipo de unidad' del Excel con un apartamento de APTOS.
    Estrategia (en orden de prioridad):
      1. Igualdad exacta (case-insensitive y sin espacios extra)
      2. El valor está contenido en el nombre del apartamento
      3. El nombre del apartamento está contenido en el valor
    Devuelve el nombre del apartamento si hay match, None si no.
    """
    t = str(tipo_unidad_str).strip()
    if not t or t.lower() in ("nan", "none", ""):
        return None
    tu = t.upper()
    # 1. Igualdad exacta
    for apto in APTOS:
        if tu == apto.upper():
            return apto
    # 2/3. Contención parcial (en ambas direcciones)
    for apto in APTOS:
        au = apto.upper()
        if tu in au or au in tu:
            return apto
    return None

def apto_libre(nombre_apto: str, f_ent: date, f_sal: date, reservas_df) -> bool:
    """
    True si el apartamento está libre en [f_ent, f_sal).
    Regla mismo día: salida ≤12h · entrada ≥16h → compatible.
    Conflicto real: f_ent < salida_existente  AND  f_sal > entrada_existente
    Maneja cualquier formato de fecha (ISO, dd/mm/yyyy, etc.) sin errores silenciosos.
    """
    for _, r in reservas_df.iterrows():
        if str(r.get("apartamento", "")).strip() != nombre_apto:
            continue
        re_d = parse_date_safe(r.get("entrada", ""))
        rs_d = parse_date_safe(r.get("salida",  ""))
        if re_d is None or rs_d is None:
            continue                   # fecha no parseable → ignorar fila
        if f_ent < rs_d and f_sal > re_d:   # solapamiento real
            return False
    return True

def asignar_aptos_auto(tipo_dorm: str, f_ent: date, f_sal: date,
                       n: int, reservas_df) -> list:
    """
    Devuelve lista de hasta n apartamentos libres del tipo solicitado.
    Tiene en cuenta tanto la BD como las reservas ya asignadas en el batch actual.
    """
    candidatos = APTOS_POR_TIPO.get(tipo_dorm, APTOS_POR_TIPO["1"])
    asignados  = []
    for c in candidatos:
        if len(asignados) >= n:
            break
        if apto_libre(c, f_ent, f_sal, reservas_df):
            asignados.append(c)
    return asignados

# ─────────────────────────────────────────────
# CONEXIÓN SUPABASE
# ─────────────────────────────────────────────
@st.cache_resource
def get_supabase() -> Client:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

supabase = get_supabase()

# ─────────────────────────────────────────────
# USUARIOS (en BD)
# ─────────────────────────────────────────────
# La tabla `usuarios` permite gestionar el acceso desde dentro de la app sin
# tocar los Secrets de Streamlit. Los usuarios definidos en st.secrets["auth"]
# siguen funcionando como "admins de rescate" — útiles para no quedarte fuera
# si la BD falla.

def cargar_usuarios_bd() -> list[dict]:
    """Lista de usuarios definidos en la tabla `usuarios` de Supabase.
    Devuelve [] si la tabla aún no existe (primer despliegue)."""
    try:
        resp = supabase.table("usuarios").select("*").order("username").execute()
        return resp.data or []
    except Exception:
        return []

def _hash_password(plain: str) -> str:
    return bcrypt.hashpw(plain.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")

def crear_usuario_bd(username: str, nombre: str, email: str,
                     password_plain: str, rol: str = "usuario") -> None:
    supabase.table("usuarios").insert({
        "username":      username,
        "nombre":        nombre,
        "email":         email,
        "password_hash": _hash_password(password_plain),
        "rol":           rol,
        "activo":        True,
    }).execute()

def actualizar_usuario_bd(user_id: int, datos: dict) -> None:
    supabase.table("usuarios").update(datos).eq("id", user_id).execute()

def cambiar_password_usuario_bd(user_id: int, password_plain: str) -> None:
    supabase.table("usuarios").update(
        {"password_hash": _hash_password(password_plain)}
    ).eq("id", user_id).execute()

def eliminar_usuario_bd(user_id: int) -> None:
    supabase.table("usuarios").delete().eq("id", user_id).execute()

# ─────────────────────────────────────────────
# AUTENTICACIÓN
# ─────────────────────────────────────────────
# La configuración de usuarios vive en dos sitios:
#   1) st.secrets["auth"] — "admins de rescate", definidos en Misterios.
#   2) Tabla `usuarios` de Supabase — gestionables desde la propia app.
# Si st.secrets["auth"] no existe, la app se bloquea con un mensaje claro:
# nadie entra sin credenciales correctas, ni siquiera por accidente.

def _build_authenticator():
    try:
        auth_cfg = st.secrets["auth"]
    except (KeyError, FileNotFoundError):
        st.error(
            "⚠️ La autenticación no está configurada todavía. "
            "Añade la sección [auth] en los Secrets de Streamlit Cloud. "
            "Consulta GUIA_DESPLIEGUE.md, sección 5."
        )
        st.stop()

    # Streamlit Secrets devuelve objetos AttrDict; los pasamos a dict planos
    # porque streamlit-authenticator espera diccionarios mutables.
    def _to_dict(obj):
        if hasattr(obj, "to_dict"):
            return obj.to_dict()
        if isinstance(obj, dict):
            return {k: _to_dict(v) for k, v in obj.items()}
        return obj

    credentials = _to_dict(auth_cfg["credentials"])
    if "usernames" not in credentials or not isinstance(credentials["usernames"], dict):
        credentials = {"usernames": {}}
    bootstrap_admins = set(credentials["usernames"].keys())

    # Fusionar usuarios de la BD (sobrescriben los de secrets si comparten nombre)
    for u in cargar_usuarios_bd():
        if not u.get("activo", True):
            continue
        credentials["usernames"][u["username"]] = {
            "email":    u.get("email") or "",
            "name":     u.get("nombre") or u["username"],
            "password": u["password_hash"],
        }

    cookie = _to_dict(auth_cfg["cookie"])
    auth = stauth.Authenticate(
        credentials,
        cookie["name"],
        cookie["key"],
        int(cookie.get("expiry_days", 30)),
    )
    return auth, bootstrap_admins

authenticator, BOOTSTRAP_ADMINS = _build_authenticator()

# Renderiza el formulario de login en el área principal
try:
    authenticator.login(
        location="main",
        fields={
            "Form name": "Iniciar sesión",
            "Username":  "Usuario",
            "Password":  "Contraseña",
            "Login":     "Entrar",
        },
    )
except Exception as ex:
    st.error(f"Error en el sistema de login: {ex}")
    st.stop()

auth_status = st.session_state.get("authentication_status")

if auth_status is False:
    st.error("Usuario o contraseña incorrectos.")
    st.stop()
elif auth_status is None:
    st.info("🔒 Esta aplicación es privada. Introduce tus credenciales para entrar.")
    st.stop()

# A partir de aquí, el usuario está autenticado.
USER_NAME     = st.session_state.get("name", "")
USER_USERNAME = st.session_state.get("username", "")

def _es_admin(username: str) -> bool:
    """Admin si está en st.secrets["auth"] o tiene rol='admin' en la BD."""
    if username in BOOTSTRAP_ADMINS:
        return True
    for u in cargar_usuarios_bd():
        if u["username"] == username and u.get("rol") == "admin" and u.get("activo", True):
            return True
    return False

IS_ADMIN = _es_admin(USER_USERNAME)

# ─────────────────────────────────────────────
# FUNCIONES DE DATOS
# ─────────────────────────────────────────────
def cargar_reservas() -> pd.DataFrame:
    resp = supabase.table("reservas").select("*").order("mes_num").order("entrada").execute()
    if resp.data:
        df = pd.DataFrame(resp.data)
        return df
    return pd.DataFrame()

def guardar_reserva(datos: dict):
    supabase.table("reservas").insert(datos).execute()

def actualizar_reserva(id_reserva: int, datos: dict):
    supabase.table("reservas").update(datos).eq("id", id_reserva).execute()

def eliminar_reserva(id_reserva: int):
    supabase.table("reservas").delete().eq("id", id_reserva).execute()

def borrar_todas_las_reservas():
    """Elimina TODAS las reservas de la base de datos."""
    supabase.table("reservas").delete().gt("id", 0).execute()

def mes_num(mes: str) -> int:
    try:
        return MESES.index(mes) + 1
    except ValueError:
        return 99

def calcular_noches(entrada_str: str, salida_str: str) -> int:
    try:
        e = datetime.strptime(entrada_str, "%d/%m/%Y")
        s = datetime.strptime(salida_str, "%d/%m/%Y")
        return max((s - e).days, 0)
    except:
        return 0

# ─────────────────────────────────────────────
# EXPORTAR EXCEL
# ─────────────────────────────────────────────
def exportar_excel(df: pd.DataFrame) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "RESERVAS COMBINADAS"

    thin = Side(style="thin", color="BBBBBB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    COLS = [
        ("Nº RESERVA", 15), ("FUENTE", 15), ("NOMBRE", 30), ("DORMITORIOS", 13),
        ("ENTRADA", 13), ("SALIDA", 13), ("NOCHES", 8), ("PERSONAS", 9),
        ("PRECIO (€)", 13), ("PAGO A CTA", 13), ("FECHA INGRESO", 15),
        ("RESTO PDTE.", 13), ("ESTADO PAGO", 22), ("COMENTARIOS", 38),
    ]

    # Cabecera
    for ci, (h, w) in enumerate(COLS, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font      = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
        cell.fill      = PatternFill("solid", fgColor="1F4E79")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = border
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 32

    # Leyenda
    ws.cell(row=2, column=1, value="Azul = Reserva Directa    Verde = Booking.com").font = Font(italic=True, color="444444", name="Calibri", size=9)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(COLS))

    row_idx = 3
    current_mes = None

    for _, r in df.iterrows():
        mes = str(r.get("mes", "")).upper()
        if mes != current_mes:
            current_mes = mes
            cell = ws.cell(row=row_idx, column=1, value=mes)
            cell.font      = Font(bold=True, color="1F4E79", name="Calibri", size=11)
            cell.fill      = PatternFill("solid", fgColor="BDD7EE")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=len(COLS))
            ws.row_dimensions[row_idx].height = 20
            row_idx += 1

        bg = "E8F5E9" if str(r.get("fuente","")) == "BOOKING.COM" else "D6E4F0"
        vals = [r.get("nro_reserva",""), r.get("fuente",""), r.get("nombre",""),
                r.get("dormitorios",""), r.get("entrada",""), r.get("salida",""),
                r.get("noches",""), r.get("personas",""), r.get("precio",""),
                r.get("pago_cta",""), r.get("fecha_ingreso",""), r.get("resto_pdte",""),
                r.get("estado_pago",""), r.get("comentarios","")]

        for ci, val in enumerate(vals, 1):
            cell = ws.cell(row=row_idx, column=ci, value=val)
            cell.fill      = PatternFill("solid", fgColor=bg)
            cell.font      = Font(name="Calibri", size=10)
            cell.alignment = Alignment(vertical="center", wrap_text=(ci == 14))
            cell.border    = border
        ws.row_dimensions[row_idx].height = 16
        row_idx += 1

    ws.freeze_panes = "A3"
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ─────────────────────────────────────────────
# ESTILOS CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
/* ── Fondo general ── */
[data-testid="stAppViewContainer"] { background: #f0f4f8; }

/* ══════════════════════════════════════
   SIDEBAR
══════════════════════════════════════ */
[data-testid="stSidebar"] {
    background: linear-gradient(170deg, #071a2e 0%, #0f2f52 45%, #1a4370 100%) !important;
    border-right: 1px solid rgba(100,181,246,0.10) !important;
}
[data-testid="stSidebar"] * { color: rgba(255,255,255,0.82) !important; }

/* Logo */
.sb-logo {
    text-align: center;
    padding: 24px 12px 20px;
    border-bottom: 1px solid rgba(255,255,255,0.07);
    margin-bottom: 4px;
}
.sb-logo-icon { font-size: 2.6rem; line-height: 1; display: block; }
.sb-logo-title {
    font-size: 1.18rem; font-weight: 800; color: white !important;
    letter-spacing: 3px; margin-top: 8px; display: block;
}
.sb-logo-sub {
    font-size: 0.65rem; color: rgba(100,181,246,0.65) !important;
    letter-spacing: 2px; margin-top: 3px; display: block; text-transform: uppercase;
}

/* Etiqueta de sección */
.sb-label {
    font-size: 0.58rem; font-weight: 700; letter-spacing: 2.5px;
    color: rgba(100,181,246,0.55) !important;
    padding: 16px 18px 5px; text-transform: uppercase; display: block;
}

/* ── Navegación: radio → botones ── */
[data-testid="stSidebar"] .stRadio > div[role="radiogroup"] { gap: 1px !important; }
[data-testid="stSidebar"] .stRadio > label { display: none !important; }

[data-testid="stSidebar"] .stRadio label {
    display: flex !important;
    align-items: center !important;
    padding: 10px 18px !important;
    border-radius: 9px !important;
    margin: 1px 10px !important;
    cursor: pointer !important;
    font-size: 0.875rem !important;
    font-weight: 500 !important;
    color: rgba(255,255,255,0.65) !important;
    border-left: 3px solid transparent !important;
    transition: all 0.18s ease !important;
    background: transparent !important;
}
[data-testid="stSidebar"] .stRadio label > div:first-child {
    display: none !important;
}
[data-testid="stSidebar"] .stRadio label > div:last-child {
    margin-left: 0 !important;
}
[data-testid="stSidebar"] .stRadio label:hover {
    background: rgba(255,255,255,0.06) !important;
    color: rgba(255,255,255,0.92) !important;
    border-left-color: rgba(100,181,246,0.35) !important;
}
[data-testid="stSidebar"] .stRadio label:has(input:checked) {
    background: rgba(100,181,246,0.14) !important;
    color: white !important;
    font-weight: 700 !important;
    border-left: 3px solid #64B5F6 !important;
    box-shadow: inset 0 0 0 1px rgba(100,181,246,0.12) !important;
}

/* ── Filtros ── */
[data-testid="stSidebar"] .stMultiSelect > label,
[data-testid="stSidebar"] .stTextInput > label {
    font-size: 0.68rem !important;
    color: rgba(255,255,255,0.45) !important;
    letter-spacing: 1px !important;
    text-transform: uppercase !important;
    margin-bottom: 2px !important;
}
[data-testid="stSidebar"] .stMultiSelect [data-baseweb="select"] > div {
    background: rgba(255,255,255,0.08) !important;
    border: 1px solid rgba(255,255,255,0.18) !important;
    border-radius: 7px !important;
    color: #ffffff !important;
    font-size: 1rem !important;
}
/* ── Buscar nombre: cubrir TODOS los niveles del DOM del input ── */
/* Contenedor baseweb (el que tiene el fondo blanco de Streamlit) */
[data-testid="stSidebar"] [data-testid="stTextInput"] [data-baseweb="input"],
[data-testid="stSidebar"] [data-testid="stTextInput"] [data-baseweb="input"] > div,
[data-testid="stSidebar"] [data-testid="stTextInput"] [data-baseweb="base-input"] {
    background-color: #12325a !important;
    border: 1px solid rgba(255,255,255,0.30) !important;
    border-radius: 7px !important;
    box-shadow: none !important;
}
/* El <input> propiamente dicho */
[data-testid="stSidebar"] [data-testid="stTextInput"] input {
    background-color: #12325a !important;
    color: #ffffff !important;
    -webkit-text-fill-color: #ffffff !important;
    caret-color: #ffffff !important;
    font-size: 1rem !important;
    font-weight: 500 !important;
    border: none !important;
    box-shadow: none !important;
}
[data-testid="stSidebar"] [data-testid="stTextInput"] input::placeholder {
    color: rgba(255,255,255,0.45) !important;
    -webkit-text-fill-color: rgba(255,255,255,0.45) !important;
}
/* Foco: borde azul */
[data-testid="stSidebar"] [data-testid="stTextInput"] [data-baseweb="input"]:focus-within {
    border-color: rgba(100,181,246,0.70) !important;
    box-shadow: 0 0 0 2px rgba(100,181,246,0.20) !important;
}
[data-testid="stSidebar"] .stMultiSelect [data-baseweb="tag"] {
    background: rgba(100,181,246,0.25) !important;
    border-radius: 5px !important;
}
[data-testid="stSidebar"] hr {
    border-color: rgba(255,255,255,0.07) !important;
    margin: 6px 0 !important;
}

/* Pie del sidebar */
.sb-footer {
    margin-top: 20px;
    padding: 12px 16px;
    border-top: 1px solid rgba(255,255,255,0.07);
    font-size: 0.68rem;
    color: rgba(255,255,255,0.28) !important;
    text-align: center;
    line-height: 1.8;
}

/* ══════════════════════════════════════
   CONTENIDO PRINCIPAL
══════════════════════════════════════ */
.metric-card {
    background: white; border-radius: 12px; padding: 18px 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.07); text-align: center;
    border-top: 3px solid #1F4E79;
}
.metric-num  { font-size: 2rem; font-weight: 800; color: #1F4E79; }
.metric-lab  { font-size: 0.82rem; color: #888; margin-top: 3px; }
.badge-directa { background:#D6E4F0; color:#1F4E79; padding:2px 9px; border-radius:12px; font-size:0.78rem; font-weight:600; }
.badge-booking { background:#E8F5E9; color:#2E7D32; padding:2px 9px; border-radius:12px; font-size:0.78rem; font-weight:600; }
.stDataFrame { border-radius: 10px; overflow: hidden; }
h1 { color: #1F4E79 !important; }
h2, h3 { color: #2C5F8A !important; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:

    # Logo
    st.image("logo.png", width=140)
    st.markdown(f"""
    <div style="text-align:center;padding:4px 10px 12px;border-bottom:1px solid rgba(255,255,255,0.07);margin-bottom:4px;">
        <span class="sb-logo-title">ESTEASUR 2015</span>
        <span style="font-size:0.72rem;color:rgba(100,181,246,0.8)!important;letter-spacing:1px;display:block;margin-top:2px;">ISLAMAR</span>
        <span class="sb-logo-sub">Gestión de Reservas</span>
        <span style="display:block;margin-top:10px;font-size:0.75rem;color:rgba(255,255,255,0.7);">
            👤 {USER_NAME or USER_USERNAME}
        </span>
    </div>
    """, unsafe_allow_html=True)

    # Botón de cierre de sesión
    authenticator.logout("🚪 Cerrar sesión", location="sidebar")

    # Navegación
    st.markdown('<span class="sb-label">Navegación</span>', unsafe_allow_html=True)
    _secciones_nav = [
        "📊 Reservas",
        "💰 Resumen de ventas",
        "📅 Plantilla mensual",
        "📥 Importar Booking",
        "➕ Nueva reserva",
        "✏️ Editar reserva",
    ]
    if IS_ADMIN:
        _secciones_nav.append("👥 Usuarios")
    seccion = st.radio("nav", _secciones_nav, label_visibility="collapsed")

    # Filtros
    st.markdown('<span class="sb-label">Filtros</span>', unsafe_allow_html=True)
    filtro_mes        = st.multiselect("Mes", MESES, placeholder="Todos los meses")
    filtro_fuente     = st.multiselect("Fuente", FUENTES, placeholder="Todas las fuentes")
    filtro_nombre     = st.text_input("Buscar nombre", placeholder="Nombre del cliente...")
    filtro_dorm       = st.multiselect("Dormitorios", DORMS, placeholder="Todos")
    mostrar_canceladas = st.checkbox("Mostrar canceladas / anuladas", value=False)

# ─────────────────────────────────────────────
# CARGAR DATOS
# ─────────────────────────────────────────────
df = cargar_reservas()

# Pie del sidebar con estadísticas
with st.sidebar:
    total_res   = len(df) if not df.empty else 0
    directas_n  = len(df[df["fuente"] == "DIRECTA"]) if not df.empty else 0
    booking_n   = len(df[df["fuente"] == "BOOKING.COM"]) if not df.empty else 0
    st.markdown(f"""
    <div class="sb-footer">
        📋 {total_res} reservas totales<br>
        🔵 {directas_n} directas &nbsp;·&nbsp; 🟢 {booking_n} Booking<br>
        <span style="opacity:.5;">ESTEASUR 2015 · ISLAMAR · 2026</span>
    </div>
    """, unsafe_allow_html=True)

if df.empty:
    st.warning("⚠️ No hay reservas cargadas. Añade la primera desde '➕ Nueva reserva'.")
    df = pd.DataFrame(columns=["id","nro_reserva","fuente","mes","mes_num","nombre",
                                "dormitorios","entrada","salida","noches","personas",
                                "precio","pago_cta","fecha_ingreso","resto_pdte",
                                "estado_pago","comentarios"])

# Aplicar filtros
df_filtrado = df.copy()
if not mostrar_canceladas and not df_filtrado.empty:
    df_filtrado = df_filtrado[~df_filtrado["estado_pago"].apply(es_cancelada)]
if filtro_mes:
    df_filtrado = df_filtrado[df_filtrado["mes"].isin(filtro_mes)]
if filtro_fuente:
    df_filtrado = df_filtrado[df_filtrado["fuente"].isin(filtro_fuente)]
if filtro_nombre:
    df_filtrado = df_filtrado[df_filtrado["nombre"].str.contains(filtro_nombre, case=False, na=False)]
if filtro_dorm:
    df_filtrado = df_filtrado[df_filtrado["dormitorios"].astype(str).isin(filtro_dorm)]

# ─────────────────────────────────────────────
# SECCIÓN: RESERVAS
# ─────────────────────────────────────────────
if seccion == "📊 Reservas":

    # KPIs
    total      = len(df_filtrado)
    directas   = len(df_filtrado[df_filtrado["fuente"] == "DIRECTA"])
    booking    = len(df_filtrado[df_filtrado["fuente"] == "BOOKING.COM"])

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f'<div class="metric-card"><div class="metric-num">{total}</div><div class="metric-lab">Reservas totales</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="metric-card"><div class="metric-num" style="color:#2C5F8A">{directas}</div><div class="metric-lab">Reservas directas</div></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="metric-card"><div class="metric-num" style="color:#2E7D32">{booking}</div><div class="metric-lab">Booking.com</div></div>', unsafe_allow_html=True)
    with c4:
        meses_activos = df_filtrado["mes"].nunique() if not df_filtrado.empty else 0
        st.markdown(f'<div class="metric-card"><div class="metric-num" style="color:#6A1B9A">{meses_activos}</div><div class="metric-lab">Meses con reservas</div></div>', unsafe_allow_html=True)

    st.markdown("")

    # Tabla editable
    st.markdown(f"### 📋 Listado de reservas ({total})  <span style='font-size:0.8rem;color:#888;font-weight:400'>— doble clic en cualquier celda para editar</span>", unsafe_allow_html=True)

    if not df_filtrado.empty:
        COLS_EDIT = ["fuente","nombre","apartamento","entrada","salida",
                     "noches","personas","precio","pago_cta","fecha_ingreso",
                     "resto_pdte","estado_pago","mes","comentarios"]
        cols_exist = [c for c in COLS_EDIT if c in df_filtrado.columns]

        # Guardamos IDs por separado para actualizar correctamente
        id_map = df_filtrado["id"].reset_index(drop=True)
        df_show = df_filtrado[cols_exist].copy().reset_index(drop=True)

        edited = st.data_editor(
            df_show,
            use_container_width=True,
            height=700,
            column_config={
                "fuente":       st.column_config.SelectboxColumn("Fuente", options=FUENTES, width=130),
                "nombre":       st.column_config.TextColumn("Nombre", width=200),
                "apartamento":  st.column_config.SelectboxColumn("Apartamento", options=[""] + APTOS, width=180),
                "entrada":      st.column_config.TextColumn("Entrada", width=100),
                "salida":       st.column_config.TextColumn("Salida", width=100),
                "noches":       st.column_config.NumberColumn("Noches", width=75),
                "personas":     st.column_config.TextColumn("Pers.", width=65),
                "precio":       st.column_config.TextColumn("Precio €", width=95),
                "pago_cta":     st.column_config.TextColumn("Pago cta €", width=100),
                "fecha_ingreso":st.column_config.TextColumn("F. Ingreso", width=110),
                "resto_pdte":   st.column_config.TextColumn("Resto pdte €", width=110),
                "estado_pago":  st.column_config.SelectboxColumn("Estado pago", options=ESTADOS, width=190),
                "mes":          st.column_config.SelectboxColumn("Mes", options=MESES, width=120),
                "comentarios":  st.column_config.TextColumn("Comentarios", width=250),
            },
            hide_index=True,
            num_rows="fixed",
            key="tabla_editable",
        )

        # Botones
        col_save, col_dl = st.columns([1, 2])
        with col_save:
            if st.button("💾 Guardar cambios", type="primary", use_container_width=True):
                cambios = 0
                for i in range(len(edited)):
                    if not df_show.iloc[i].equals(edited.iloc[i]):
                        id_r  = int(id_map.iloc[i])
                        datos = edited.iloc[i].to_dict()
                        # Recalcular noches si cambiaron fechas
                        noches = calcular_noches(
                            str(datos.get("entrada", "")),
                            str(datos.get("salida", ""))
                        )
                        if noches:
                            datos["noches"] = noches
                        datos["mes_num"] = mes_num(str(datos.get("mes", "")))
                        actualizar_reserva(id_r, datos)
                        cambios += 1
                if cambios:
                    st.success(f"✅ {cambios} registro(s) actualizados correctamente.")
                    st.rerun()
                else:
                    st.info("No hay cambios que guardar.")
        with col_dl:
            excel_bytes = exportar_excel(df_filtrado)
            st.download_button(
                label="⬇️ Descargar Excel actualizado",
                data=excel_bytes,
                file_name=f"Reservas_ESTEASUR_ISLAMAR_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    else:
        st.info("No hay reservas con los filtros seleccionados.")

# ─────────────────────────────────────────────
# SECCIÓN: NUEVA RESERVA
# ─────────────────────────────────────────────
elif seccion == "➕ Nueva reserva":
    st.markdown("### ➕ Añadir nueva reserva")

    with st.form("form_nueva", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            fuente      = st.selectbox("Fuente *", FUENTES)
            nombre      = st.text_input("Nombre del cliente *")
            nro_reserva = st.text_input("Nº de reserva")
            apartamento = st.selectbox("Apartamento *", [""] + APTOS)
            dormitorios = st.selectbox("Dormitorios", DORMS)
            mes         = st.selectbox("Mes *", MESES)
        with c2:
            entrada     = st.date_input("Fecha entrada *", value=None, format="DD/MM/YYYY")
            salida      = st.date_input("Fecha salida *",  value=None, format="DD/MM/YYYY")
            personas    = st.text_input("Nº personas")
            precio      = st.text_input("Precio (€)")
            estado_pago = st.selectbox("Estado de pago", ESTADOS)

        c3, c4 = st.columns(2)
        with c3:
            pago_cta    = st.text_input("Pago a cuenta (€)")
            fecha_ing   = st.text_input("Fecha ingreso")
        with c4:
            resto_pdte  = st.text_input("Resto pendiente (€)")

        comentarios = st.text_area("Comentarios", height=80)

        submitted = st.form_submit_button("💾 Guardar reserva", type="primary", use_container_width=True)

    if submitted:
        errores = []
        if not nombre:  errores.append("El nombre es obligatorio.")
        if not entrada: errores.append("La fecha de entrada es obligatoria.")
        if not salida:  errores.append("La fecha de salida es obligatoria.")
        if entrada and salida and salida <= entrada:
            errores.append("La fecha de salida debe ser posterior a la de entrada.")

        if errores:
            for e in errores:
                st.error(e)
        else:
            entrada_str = entrada.strftime("%d/%m/%Y") if entrada else ""
            salida_str  = salida.strftime("%d/%m/%Y")  if salida  else ""
            noches      = (salida - entrada).days if entrada and salida else 0

            datos = {
                "nro_reserva": nro_reserva,
                "fuente":      fuente,
                "mes":         mes,
                "mes_num":     mes_num(mes),
                "nombre":      nombre,
                "apartamento": apartamento,
                "dormitorios": dormitorios,
                "entrada":     entrada_str,
                "salida":      salida_str,
                "noches":      noches,
                "personas":    personas,
                "precio":      precio,
                "pago_cta":    pago_cta,
                "fecha_ingreso": fecha_ing,
                "resto_pdte":  resto_pdte,
                "estado_pago": estado_pago,
                "comentarios": comentarios,
            }
            guardar_reserva(datos)
            st.success(f"✅ Reserva de **{nombre}** guardada correctamente.")
            st.cache_resource.clear()
            st.rerun()

# ─────────────────────────────────────────────
# SECCIÓN: EDITAR RESERVA
# ─────────────────────────────────────────────
elif seccion == "✏️ Editar reserva":
    st.markdown("### ✏️ Editar o eliminar una reserva")

    if df.empty:
        st.info("No hay reservas cargadas todavía.")
    else:
        # Selector de reserva
        opciones = {
            f"{row['nombre']}  |  {row.get('entrada','')} → {row.get('salida','')}  |  {row.get('mes','')}": row["id"]
            for _, row in df.iterrows()
        }
        seleccion = st.selectbox("Selecciona la reserva a editar:", list(opciones.keys()))
        id_sel    = opciones[seleccion]
        reserva   = df[df["id"] == id_sel].iloc[0]

        st.markdown("---")

        def parse_date(s):
            try: return datetime.strptime(str(s), "%d/%m/%Y").date()
            except: return None

        with st.form("form_editar"):
            c1, c2 = st.columns(2)
            with c1:
                fuente      = st.selectbox("Fuente", FUENTES, index=FUENTES.index(reserva["fuente"]) if reserva["fuente"] in FUENTES else 0)
                nombre      = st.text_input("Nombre del cliente *", value=str(reserva.get("nombre","")))
                nro_reserva = st.text_input("Nº de reserva", value=str(reserva.get("nro_reserva","")))
                apto_val    = str(reserva.get("apartamento",""))
                apto_opts   = [""] + APTOS
                apartamento = st.selectbox("Apartamento", apto_opts, index=apto_opts.index(apto_val) if apto_val in apto_opts else 0)
                dorm_val    = str(reserva.get("dormitorios","1"))
                dormitorios = st.selectbox("Dormitorios", DORMS, index=DORMS.index(dorm_val) if dorm_val in DORMS else 0)
                mes_val     = str(reserva.get("mes","ENERO")).upper()
                mes         = st.selectbox("Mes", MESES, index=MESES.index(mes_val) if mes_val in MESES else 0)
            with c2:
                entrada     = st.date_input("Fecha entrada", value=parse_date(reserva.get("entrada")), format="DD/MM/YYYY")
                salida      = st.date_input("Fecha salida",  value=parse_date(reserva.get("salida")),  format="DD/MM/YYYY")
                personas    = st.text_input("Nº personas", value=str(reserva.get("personas","")))
                precio      = st.text_input("Precio (€)",  value=str(reserva.get("precio","")))
                est_val     = str(reserva.get("estado_pago",""))
                estado_pago = st.selectbox("Estado de pago", ESTADOS, index=ESTADOS.index(est_val) if est_val in ESTADOS else 0)

            c3, c4 = st.columns(2)
            with c3:
                pago_cta  = st.text_input("Pago a cuenta (€)", value=str(reserva.get("pago_cta","")))
                fecha_ing = st.text_input("Fecha ingreso",     value=str(reserva.get("fecha_ingreso","")))
            with c4:
                resto_pdte = st.text_input("Resto pendiente (€)", value=str(reserva.get("resto_pdte","")))

            comentarios = st.text_area("Comentarios", value=str(reserva.get("comentarios","")), height=80)

            col_save, col_del = st.columns([3, 1])
            with col_save:
                submitted = st.form_submit_button("💾 Guardar cambios", type="primary", use_container_width=True)
            with col_del:
                eliminar  = st.form_submit_button("🗑️ Eliminar", use_container_width=True)

        if submitted:
            entrada_str = entrada.strftime("%d/%m/%Y") if entrada else ""
            salida_str  = salida.strftime("%d/%m/%Y")  if salida  else ""
            noches      = (salida - entrada).days if entrada and salida else 0
            datos = {
                "nro_reserva": nro_reserva, "fuente": fuente, "mes": mes,
                "mes_num": mes_num(mes), "nombre": nombre, "apartamento": apartamento,
                "dormitorios": dormitorios, "entrada": entrada_str, "salida": salida_str,
                "noches": noches, "personas": personas, "precio": precio,
                "pago_cta": pago_cta, "fecha_ingreso": fecha_ing, "resto_pdte": resto_pdte,
                "estado_pago": estado_pago, "comentarios": comentarios,
            }
            actualizar_reserva(id_sel, datos)
            st.success(f"✅ Reserva de **{nombre}** actualizada.")
            st.cache_resource.clear()
            st.rerun()

        if eliminar:
            eliminar_reserva(id_sel)
            st.success("🗑️ Reserva eliminada.")
            st.cache_resource.clear()
            st.rerun()

# ─────────────────────────────────────────────
# SECCIÓN: RESUMEN DE VENTAS
# ─────────────────────────────────────────────
elif seccion == "💰 Resumen de ventas":

    st.markdown("### 💰 Resumen de ventas")

    def to_eur(v):
        try:
            return float(str(v).replace(",", ".").replace("€", "").replace(" ", "").strip())
        except:
            return 0.0

    def fmt_es(v, dec=2):
        """Formato español: 1.234,56 €"""
        try:
            s = f"{float(v):,.{dec}f}"
            return s.replace(",", "X").replace(".", ",").replace("X", ".")
        except:
            return "—"

    if df.empty:
        st.info("No hay reservas cargadas todavía.")
    else:
        # ── Tasa de comisión configurable ─────────────────
        col_tasa, col_void = st.columns([1, 3])
        with col_tasa:
            tasa_com = st.number_input(
                "Comisión Booking.com (%)", min_value=0.0, max_value=30.0,
                value=15.0, step=0.5, format="%.1f",
                help="Porcentaje que cobra Booking.com sobre el precio bruto"
            )
        tasa = tasa_com / 100.0

        # Preparar datos numéricos
        df_v = df.copy()
        df_v["precio_eur"] = df_v["precio"].apply(to_eur)
        df_v["comision_eur"] = df_v.apply(
            lambda r: r["precio_eur"] * tasa if str(r.get("fuente", "")) == "BOOKING.COM" else 0.0,
            axis=1
        )
        df_v["neto_eur"] = df_v["precio_eur"] - df_v["comision_eur"]

        # ── KPIs ──────────────────────────────────────────
        st.markdown("")
        total_bruto    = df_v["precio_eur"].sum()
        total_com      = df_v["comision_eur"].sum()
        total_neto     = df_v["neto_eur"].sum()
        n_booking      = len(df_v[df_v["fuente"] == "BOOKING.COM"])
        n_directa      = len(df_v[df_v["fuente"] == "DIRECTA"])

        k1, k2, k3, k4 = st.columns(4)
        k1.markdown(f'<div class="metric-card"><div class="metric-num">{fmt_es(total_bruto, 0)} €</div><div class="metric-lab">Ingresos brutos</div></div>', unsafe_allow_html=True)
        k2.markdown(f'<div class="metric-card"><div class="metric-num" style="color:#c0392b">{fmt_es(total_com, 0)} €</div><div class="metric-lab">Comisiones Booking.com</div></div>', unsafe_allow_html=True)
        k3.markdown(f'<div class="metric-card"><div class="metric-num" style="color:#27ae60">{fmt_es(total_neto, 0)} €</div><div class="metric-lab">Ingreso neto</div></div>', unsafe_allow_html=True)
        k4.markdown(f'<div class="metric-card"><div class="metric-num" style="color:#8e44ad">{n_booking}</div><div class="metric-lab">Reservas Booking ({n_directa} directas)</div></div>', unsafe_allow_html=True)

        # ── KPI split: Propios vs JUANMA ──────────────────
        df_v_propios = df_v[~df_v["apartamento"].isin(APTOS_JUANMA)]
        df_v_juanma  = df_v[ df_v["apartamento"].isin(APTOS_JUANMA)]
        kp1, kp2 = st.columns(2)
        kp1.markdown(
            f'<div class="metric-card">'
            f'<div class="metric-num" style="color:#1a5276">{fmt_es(df_v_propios["neto_eur"].sum(), 0)} €</div>'
            f'<div class="metric-lab">🏠 Apartamentos propios — ingreso neto</div></div>',
            unsafe_allow_html=True,
        )
        kp2.markdown(
            f'<div class="metric-card">'
            f'<div class="metric-num" style="color:#1a5276">{fmt_es(df_v_juanma["neto_eur"].sum(), 0)} €</div>'
            f'<div class="metric-lab">👤 JUANMA — ingreso neto</div></div>',
            unsafe_allow_html=True,
        )

        st.markdown("---")

        # ── Tabla: Ingresos por mes ────────────────────────
        st.markdown("#### 📅 Ingresos por mes")

        meses_ord = [m for m in MESES if m in df_v["mes"].values]
        resumen_mes = []
        for m in meses_ord:
            sub = df_v[df_v["mes"] == m]
            sub_d = sub[sub["fuente"] == "DIRECTA"]
            sub_b = sub[sub["fuente"] == "BOOKING.COM"]
            bruto_d = sub_d["precio_eur"].sum()
            bruto_b = sub_b["precio_eur"].sum()
            com_b   = sub_b["comision_eur"].sum()
            resumen_mes.append({
                "Mes":                m,
                "Reservas":           len(sub),
                "Ingresos directa €": round(bruto_d, 2),
                "Ingresos Booking €": round(bruto_b, 2),
                "Comisión Booking €": round(com_b, 2),
                "Total bruto €":      round(bruto_d + bruto_b, 2),
                "Total neto €":       round(bruto_d + bruto_b - com_b, 2),
            })

        df_mes = pd.DataFrame(resumen_mes)
        # Fila de totales
        totales_row = {
            "Mes": "▶ TOTAL",
            "Reservas":           df_mes["Reservas"].sum(),
            "Ingresos directa €": round(df_mes["Ingresos directa €"].sum(), 2),
            "Ingresos Booking €": round(df_mes["Ingresos Booking €"].sum(), 2),
            "Comisión Booking €": round(df_mes["Comisión Booking €"].sum(), 2),
            "Total bruto €":      round(df_mes["Total bruto €"].sum(), 2),
            "Total neto €":       round(df_mes["Total neto €"].sum(), 2),
        }
        df_mes_tot = pd.concat([df_mes, pd.DataFrame([totales_row])], ignore_index=True)

        # Formatear columnas € en español
        cols_eur_mes = ["Ingresos directa €", "Ingresos Booking €", "Comisión Booking €", "Total bruto €", "Total neto €"]
        df_mes_show = df_mes_tot.copy()
        for c in cols_eur_mes:
            df_mes_show[c] = df_mes_show[c].apply(lambda x: fmt_es(x) if isinstance(x, (int, float)) else x)

        st.dataframe(
            df_mes_show,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Mes":                st.column_config.TextColumn("Mes",                width=130),
                "Reservas":           st.column_config.NumberColumn("Reservas",         width=90, format="%d"),
                "Ingresos directa €": st.column_config.TextColumn("Directa €",          width=130),
                "Ingresos Booking €": st.column_config.TextColumn("Booking bruto €",    width=145),
                "Comisión Booking €": st.column_config.TextColumn("Comisión Booking €", width=155),
                "Total bruto €":      st.column_config.TextColumn("Total bruto €",      width=130),
                "Total neto €":       st.column_config.TextColumn("Total neto €",       width=130),
            },
        )

        st.markdown("---")

        # ── Tabla pivot: Ingresos por apartamento ─────────
        st.markdown("#### 🏠 Ingresos por apartamento y mes")

        aptos_con_datos = [a for a in APTOS if a in df_v["apartamento"].values]
        aptos_propios_d = [a for a in aptos_con_datos if a not in APTOS_JUANMA]
        aptos_juanma_d  = [a for a in aptos_con_datos if a in APTOS_JUANMA]
        meses_piv = [m for m in MESES if m in df_v["mes"].values]
        df_pivot  = pd.DataFrame()   # para el export aunque no haya datos

        def _piv_filas(aptos_list):
            filas = []
            for apto in aptos_list:
                sub_a = df_v[df_v["apartamento"] == apto]
                fila = {"Apartamento": apto}
                total_apto = 0.0
                for m in meses_piv:
                    val = sub_a[sub_a["mes"] == m]["neto_eur"].sum()
                    fila[m] = round(val, 2)
                    total_apto += val
                fila["TOTAL €"] = round(total_apto, 2)
                filas.append(fila)
            return filas

        def _subtotal_fila(label, aptos_list):
            sub_df = df_v[df_v["apartamento"].isin(aptos_list)]
            fila = {"Apartamento": label}
            for m in meses_piv:
                fila[m] = round(sub_df[sub_df["mes"] == m]["neto_eur"].sum(), 2)
            fila["TOTAL €"] = round(sub_df["neto_eur"].sum(), 2)
            return fila

        def _fmt_piv(df_in):
            df_out = df_in.copy()
            for c in meses_piv + ["TOTAL €"]:
                if c in df_out.columns:
                    df_out[c] = df_out[c].apply(lambda x: fmt_es(x) if isinstance(x, (int, float)) else x)
            return df_out

        col_cfg_pivot = {
            "Apartamento": st.column_config.TextColumn("Apartamento", width=200),
            "TOTAL €":     st.column_config.TextColumn("TOTAL €",     width=115),
        }
        for m in meses_piv:
            col_cfg_pivot[m] = st.column_config.TextColumn(m[:3], width=85)

        if aptos_con_datos and meses_piv:
            all_filas_export = []

            # ── Apartamentos propios ──
            if aptos_propios_d:
                st.markdown("**🏠 Apartamentos propios**")
                fp = _piv_filas(aptos_propios_d)
                sub_p = _subtotal_fila("▶ SUBTOTAL PROPIOS", aptos_propios_d)
                fp.append(sub_p)
                df_pp = pd.DataFrame(fp)
                st.dataframe(_fmt_piv(df_pp), use_container_width=True,
                             hide_index=True, column_config=col_cfg_pivot)
                all_filas_export += fp

            # ── Apartamentos JUANMA ──
            if aptos_juanma_d:
                st.markdown("**👤 Apartamentos JUANMA**")
                fj = _piv_filas(aptos_juanma_d)
                sub_j = _subtotal_fila("▶ SUBTOTAL JUANMA", aptos_juanma_d)
                fj.append(sub_j)
                df_jj = pd.DataFrame(fj)
                st.dataframe(_fmt_piv(df_jj), use_container_width=True,
                             hide_index=True, column_config=col_cfg_pivot)
                all_filas_export += fj

            # ── Total general ──
            fila_gen = _subtotal_fila("▶ TOTAL GENERAL", aptos_con_datos)
            all_filas_export.append(fila_gen)
            st.dataframe(_fmt_piv(pd.DataFrame([fila_gen])), use_container_width=True,
                         hide_index=True, column_config=col_cfg_pivot)

            df_pivot = pd.DataFrame(all_filas_export)
        else:
            st.info("Asigna apartamentos a las reservas para ver este desglose.")

        st.markdown("---")

        # ── Detalle comisiones Booking ─────────────────────
        st.markdown("#### 🏷️ Detalle comisiones Booking.com por mes")

        df_bk_v = df_v[df_v["fuente"] == "BOOKING.COM"].copy()
        if df_bk_v.empty:
            st.info("No hay reservas de Booking.com con precio registrado.")
        else:
            det_com = []
            for m in meses_ord:
                sub_bm = df_bk_v[df_bk_v["mes"] == m]
                if sub_bm.empty:
                    continue
                bruto = sub_bm["precio_eur"].sum()
                com   = sub_bm["comision_eur"].sum()
                neto  = sub_bm["neto_eur"].sum()
                det_com.append({
                    "Mes":            m,
                    "Reservas Bk.":   len(sub_bm),
                    "Precio bruto €": round(bruto, 2),
                    f"Comisión ({tasa_com:.0f}%) €": round(com, 2),
                    "Ingreso neto €": round(neto, 2),
                })
            df_com = pd.DataFrame(det_com)
            tot_com_row = {
                "Mes": "▶ TOTAL",
                "Reservas Bk.":   df_com["Reservas Bk."].sum(),
                "Precio bruto €": round(df_com["Precio bruto €"].sum(), 2),
                f"Comisión ({tasa_com:.0f}%) €": round(df_com[f"Comisión ({tasa_com:.0f}%) €"].sum(), 2),
                "Ingreso neto €": round(df_com["Ingreso neto €"].sum(), 2),
            }
            df_com = pd.concat([df_com, pd.DataFrame([tot_com_row])], ignore_index=True)

            # Formatear columnas € en español
            cols_eur_com = ["Precio bruto €", f"Comisión ({tasa_com:.0f}%) €", "Ingreso neto €"]
            df_com_show = df_com.copy()
            for c in cols_eur_com:
                if c in df_com_show.columns:
                    df_com_show[c] = df_com_show[c].apply(lambda x: fmt_es(x) if isinstance(x, (int, float)) else x)

            st.dataframe(
                df_com_show,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Mes":            st.column_config.TextColumn("Mes",                        width=130),
                    "Reservas Bk.":   st.column_config.NumberColumn("Reservas",                width=90, format="%d"),
                    "Precio bruto €": st.column_config.TextColumn("Bruto €",                   width=130),
                    f"Comisión ({tasa_com:.0f}%) €": st.column_config.TextColumn(f"Comisión {tasa_com:.0f}% €", width=150),
                    "Ingreso neto €": st.column_config.TextColumn("Neto €",                    width=130),
                },
            )

            # Nota informativa
            st.markdown(
                f"<span style='font-size:0.78rem;color:#888;'>💡 La comisión aplicada es del <b>{tasa_com:.1f}%</b> sobre el precio bruto de cada reserva de Booking.com. "
                f"Modifica el porcentaje arriba si tu contrato con Booking establece otro porcentaje.</span>",
                unsafe_allow_html=True,
            )

        # ── Botón exportar resumen ─────────────────────────
        st.markdown("---")

        def exportar_resumen(df_mes_t, df_piv, df_com_t):
            wb2 = openpyxl.Workbook()
            thin2 = Side(style="thin", color="CCCCCC")
            brd2  = Border(left=thin2, right=thin2, top=thin2, bottom=thin2)

            def hdr_cell(ws, row, col, val, bg="1F4E79"):
                c = ws.cell(row=row, column=col, value=val)
                c.font      = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
                c.fill      = PatternFill("solid", fgColor=bg)
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border    = brd2
                return c

            def data_cell(ws, row, col, val, bold=False, bg=None):
                c = ws.cell(row=row, column=col, value=val)
                c.font   = Font(bold=bold, name="Calibri", size=10)
                c.border = brd2
                if bg:
                    c.fill = PatternFill("solid", fgColor=bg)
                return c

            # Hoja 1: Por mes
            ws1 = wb2.active
            ws1.title = "Por Mes"
            cols1 = list(df_mes_t.columns)
            for ci, h in enumerate(cols1, 1):
                hdr_cell(ws1, 1, ci, h)
                ws1.column_dimensions[get_column_letter(ci)].width = 18
            for ri, row in df_mes_t.iterrows():
                is_tot = str(row.get("Mes","")).startswith("▶")
                bg = "BDD7EE" if is_tot else None
                for ci, col in enumerate(cols1, 1):
                    data_cell(ws1, ri + 2, ci, row[col], bold=is_tot, bg=bg if is_tot else None)

            # Hoja 2: Por apartamento
            ws2 = wb2.create_sheet("Por Apartamento")
            if df_piv is not None:
                cols2 = list(df_piv.columns)
                for ci, h in enumerate(cols2, 1):
                    hdr_cell(ws2, 1, ci, h)
                    ws2.column_dimensions[get_column_letter(ci)].width = 14
                ws2.column_dimensions["A"].width = 22
                for ri, row in df_piv.iterrows():
                    is_tot = str(row.get("Apartamento","")).startswith("▶")
                    for ci, col in enumerate(cols2, 1):
                        data_cell(ws2, ri + 2, ci, row[col], bold=is_tot, bg="BDD7EE" if is_tot else None)

            # Hoja 3: Comisiones
            ws3 = wb2.create_sheet("Comisiones Booking")
            if df_com_t is not None:
                cols3 = list(df_com_t.columns)
                for ci, h in enumerate(cols3, 1):
                    hdr_cell(ws3, 1, ci, h)
                    ws3.column_dimensions[get_column_letter(ci)].width = 18
                for ri, row in df_com_t.iterrows():
                    is_tot = str(row.get("Mes","")).startswith("▶")
                    for ci, col in enumerate(cols3, 1):
                        data_cell(ws3, ri + 2, ci, row[col], bold=is_tot, bg="BDD7EE" if is_tot else None)

            buf2 = BytesIO()
            wb2.save(buf2)
            return buf2.getvalue()

        df_piv_exp  = df_pivot if not df_pivot.empty else None
        df_com_exp  = df_com if not df_bk_v.empty else None
        excel_res   = exportar_resumen(df_mes_tot, df_piv_exp, df_com_exp)
        st.download_button(
            label="⬇️ Descargar resumen en Excel",
            data=excel_res,
            file_name=f"Resumen_Ventas_ISLAMAR_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=False,
        )

# ─────────────────────────────────────────────
# SECCIÓN: PLANTILLA MENSUAL
# ─────────────────────────────────────────────
elif seccion == "📅 Plantilla mensual":

    # ── Selectores ────────────────────────────
    col_mes, col_anio = st.columns([2, 1])
    with col_mes:
        mes_sel  = st.selectbox("Mes", MESES, index=datetime.now().month - 1, key="pm_mes")
    with col_anio:
        anio_sel = int(st.number_input("Año", min_value=2024, max_value=2030, value=2026, key="pm_anio"))

    mes_n      = MESES.index(mes_sel) + 1
    n_dias     = calendar.monthrange(anio_sel, mes_n)[1]
    dias       = list(range(1, n_dias + 1))
    primer_dia = date(anio_sel, mes_n, 1).weekday()   # 0=Lunes
    DIAS_SEM   = ["L","M","X","J","V","S","D"]

    # ── Construir grid, salida_map y reverse_map ──────────
    grid       = {apto: {d: None for d in dias} for apto in APTOS}
    salida_map = {}   # (apto, día) → datos del cliente que SALE ese día
    reverse_map = {}  # (apto, día) → reservation_id

    dias_set = set(dias)
    for _, r in df.iterrows():
        # ── Ignorar reservas canceladas / anuladas ──────────────────────
        if es_cancelada(r.get("estado_pago", "")):
            continue

        # ── Parsear fechas con soporte multi-formato ────────────────────
        entrada = parse_date_safe(r.get("entrada", ""))
        salida  = parse_date_safe(r.get("salida",  ""))
        apto    = str(r.get("apartamento", "")).strip()

        if not entrada or not salida or apto not in APTOS:
            continue
        try:
            rid = int(r.get("id"))
        except Exception:
            continue

        edia = entrada.day if entrada.month == mes_n and entrada.year == anio_sel else 0
        sdia = salida.day  if salida.month  == mes_n and salida.year  == anio_sel else n_dias + 1

        # Fechas siempre en dd/mm/yyyy para el display (independiente de cómo estén en BD)
        ent_str = entrada.strftime("%d/%m/%Y")
        sal_str = salida.strftime("%d/%m/%Y")

        data = {
            "id": rid, "nombre": str(r.get("nombre", "")),
            "fuente": str(r.get("fuente", "")),
            "entrada": ent_str, "salida": sal_str,
            "precio": str(r.get("precio", "")), "estado_pago": str(r.get("estado_pago", "")),
            "edia": edia, "sdia": sdia,
        }
        for d in dias:
            curr = date(anio_sel, mes_n, d)
            if entrada <= curr < salida:
                grid[apto][d] = data
                reverse_map[(apto, d)] = rid
        # Registrar día de salida para casilla compartida / media casilla checkout
        if salida.month == mes_n and salida.year == anio_sel and salida.day in dias_set:
            salida_map[(apto, salida.day)] = data

    # ── Versión de edición (para refrescar data_editor tras guardar) ──
    if "edit_ver" not in st.session_state:
        st.session_state["edit_ver"] = 0

    # ── Tabs ─────────────────────────────────
    tab_vista, tab_edit = st.tabs(["📅 Vista calendario", "✏️ Editar cuadro"])

    # ══════════════════════════════════════════
    # TAB 1: VISTA CALENDARIO (HTML - lectura)
    # ══════════════════════════════════════════
    with tab_vista:
        st.markdown("""
        <style>
        .cal-wrap{overflow-x:auto;border-radius:10px;box-shadow:0 3px 12px rgba(0,0,0,0.15);margin-bottom:8px;}
        .cal-tbl{border-collapse:collapse;font-family:'Segoe UI',Arial,sans-serif;width:100%;}
        .th-apto{background:#1a3f5c;color:white;padding:6px 14px;text-align:left;font-size:0.78rem;
                 position:sticky;left:0;z-index:3;white-space:nowrap;min-width:170px;
                 border-right:2px solid #0d2a3d;border-bottom:1px solid #0d2a3d;}
        .th-day{background:#1F4E79;color:white;padding:5px 2px;text-align:center;
                font-size:0.75rem;min-width:85px;border:1px solid #144070;line-height:1.3;}
        .th-day.we{background:#163d5e;}
        .th-day.sun{background:#163d5e;min-width:85px;}
        .td.sun{background:#f0f2f5;}
        .dow{font-size:0.62rem;color:#90CAF9;display:block;}
        .td-apto{background:#2C5F8A;color:white;font-weight:700;padding:5px 14px;white-space:nowrap;
                 font-size:0.82rem;position:sticky;left:0;z-index:1;
                 border-right:2px solid #144070;border-bottom:1px solid #1a4a72;}
        .td{padding:0;border:1px solid #dde2ea;height:58px;vertical-align:middle;
            overflow:hidden;position:relative;box-sizing:border-box;}
        .td.we{background:#eceff1;}
        .td.sun{background:#eceff1;}
        .td.libre{background:#fafbfd;}
        .sep td{background:#D0E8F7;color:#1F4E79;font-weight:700;padding:4px 10px;
                font-size:0.78rem;border-top:2px solid #1F4E79;letter-spacing:.5px;}
        </style>
        """, unsafe_allow_html=True)

        html = '<div class="cal-wrap"><table class="cal-tbl">'
        html += f'<tr><th class="th-apto" style="font-size:0.88rem;font-weight:700;">{mes_sel} {anio_sel}</th>'
        for d in dias:
            wd = (primer_dia + d - 1) % 7
            we = " sun" if wd == 6 else (" we" if wd == 5 else "")
            html += f'<th class="th-day{we}">{d}<span class="dow">{DIAS_SEM[wd]}</span></th>'
        html += '</tr>'

        # Paleta de 14 colores distintos y legibles — se asignan por ID de reserva
        # Estos colores se aplican AL TEXTO del nombre (no al fondo). Todos son
        # tonos oscuros y saturados, legibles sobre el fondo azul claro de la barra.
        _PALETA = [
            "#1F4E79",  # azul corporativo (oscuro)
            "#C0622A",  # naranja tostado
            "#2E8B6E",  # verde esmeralda
            "#7B3FA0",  # violeta
            "#B5452A",  # rojo ladrillo
            "#1A7A6E",  # verde azulado oscuro
            "#A0522D",  # sienna
            "#1B3A6B",  # azul marino
            "#7A5C00",  # ocre dorado
            "#5B3A8A",  # índigo
            "#2B7A4B",  # verde bosque
            "#8B3A62",  # frambuesa
            "#2C4E70",  # azul pizarra oscuro
            "#6B5C2E",  # marrón kaki
        ]

        # Fondo azul claro uniforme para TODAS las barras de reservas.
        # El color por reserva ya no es el fondo sino el TEXTO del nombre.
        BAR_BG       = "#7FB3DC"   # azul medio
        BAR_BORDER   = "#4A82A8"   # azul oscuro para el contorno

        def _color_reserva(rid=0):
            """Color del TEXTO del nombre del cliente, único por reserva."""
            return _PALETA[int(rid) % len(_PALETA)]

        for i, apto in enumerate(APTOS):
            if apto == "APTO 215 - 2 DORM":
                html += f'<tr class="sep"><td colspan="{n_dias+1}">▸ JUANMA</td></tr>'
            html += f'<tr><td class="td-apto">{apto}</td>'

            d = 1
            while d <= n_dias:
                c     = grid[apto][d]
                c_out = salida_map.get((apto, d))
                split = c and c_out and c.get("id") != c_out.get("id")
                wd    = (primer_dia + d - 1) % 7
                wc    = " sun" if wd == 6 else (" we" if wd == 5 else "")

                if split:
                    # ── Casilla dividida: checkout arriba / checkin abajo ──
                    txt_out = _color_reserva(c_out["id"])
                    txt_in  = _color_reserva(c["id"])
                    tip = f"SALE: {c_out['nombre']} ({c_out['salida']}) / ENTRA: {c['nombre']} ({c['entrada']})"
                    html += (
                        f'<td class="td{wc}" style="padding:0;position:relative;overflow:hidden;" title="{tip}">'
                        f'<div style="position:absolute;top:0;left:0;right:0;height:50%;background:{BAR_BG};'
                        f'display:flex;align-items:center;overflow:hidden;">'
                        f'<span style="color:{txt_out};font-size:0.74rem;font-weight:800;padding:0 6px;'
                        f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">◀ {c_out["nombre"]}</span></div>'
                        f'<div style="position:absolute;top:50%;left:0;right:0;height:50%;background:{BAR_BG};'
                        f'border-top:2px solid {BAR_BORDER};display:flex;align-items:center;overflow:hidden;">'
                        f'<span style="color:{txt_in};font-size:0.74rem;font-weight:800;padding:0 6px;'
                        f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">▶ {c["nombre"]}</span></div>'
                        f'</td>'
                    )
                    d += 1

                elif c:
                    # ── Barra con colspan: agrupa todos los días consecutivos en un solo TD ──
                    curr_rid = c["id"]
                    # Calcular cuántos días consecutivos pertenecen a esta reserva
                    span_end = d
                    while span_end < n_dias:
                        nc = grid[apto].get(span_end + 1)
                        if nc is None or nc.get("id") != curr_rid:
                            break
                        span_end += 1
                    colspan = span_end - d + 1

                    txt_color      = _color_reserva(curr_rid)
                    started_before = (c["edia"] == 0)       # empezó antes del mes
                    ends_after     = (c["sdia"] > n_dias)   # termina después del mes

                    # Márgenes y radio según si la barra viene/va más allá del mes
                    left_px  = "0"   if started_before else "4px"
                    right_px = "0"   if ends_after     else "4px"
                    if   started_before and ends_after:     brad = "3px"
                    elif started_before:                    brad = "0 7px 7px 0"
                    elif ends_after:                        brad = "7px 0 0 7px"
                    else:                                   brad = "7px"

                    prefix    = "↩ " if started_before else ""
                    name_text = f"{prefix}{c['nombre']}"
                    tip       = f"{c['nombre']} | {c['entrada']} → {c['salida']}"

                    html += (
                        f'<td class="td{wc}" colspan="{colspan}" '
                        f'style="padding:0;position:relative;overflow:hidden;" title="{tip}">'
                        f'<div style="position:absolute;top:6px;bottom:6px;'
                        f'left:{left_px};right:{right_px};background:{BAR_BG};'
                        f'border:1px solid {BAR_BORDER};border-radius:{brad};overflow:hidden;'
                        f'display:flex;align-items:center;padding:0 10px;">'
                        f'<span style="font-size:0.83rem;font-weight:800;color:{txt_color};'
                        f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">'
                        f'{name_text}</span></div></td>'
                    )
                    d = span_end + 1

                elif c_out:
                    # ── Solo checkout ese día (sin nueva entrada) ──
                    txt_out = _color_reserva(c_out["id"])
                    fbg = "#eaecef" if wd >= 5 else "#fafbfd"
                    tip = f"SALE: {c_out['nombre']} ({c_out['salida']})"
                    html += (
                        f'<td class="td{wc}" style="padding:0;position:relative;overflow:hidden;" title="{tip}">'
                        f'<div style="position:absolute;top:0;left:0;right:0;height:50%;background:{BAR_BG};'
                        f'display:flex;align-items:center;overflow:hidden;">'
                        f'<span style="color:{txt_out};font-size:0.74rem;font-weight:800;padding:0 6px;'
                        f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">◀ {c_out["nombre"]}</span></div>'
                        f'<div style="position:absolute;top:50%;left:0;right:0;height:50%;'
                        f'background:{fbg};border-top:1px solid #bbb;overflow:hidden;"></div>'
                        f'</td>'
                    )
                    d += 1

                else:
                    # ── Casilla libre ──
                    fbg = "#eaecef" if wd >= 5 else "#fafbfd"
                    html += f'<td class="td libre{wc}" style="background:{fbg};"></td>'
                    d += 1

            html += '</tr>'
        html += '</table></div>'

        st.markdown("""
        <div style="display:flex;gap:14px;align-items:center;font-size:0.78rem;margin-bottom:6px;flex-wrap:wrap;">
          <span style="color:#888;">El color del texto identifica cada reserva &nbsp;|&nbsp;
          ↩ Entró mes anterior &nbsp;|&nbsp; ◀/▶ Casilla dividida (salida/entrada mismo día) &nbsp;|&nbsp; Gris = fin de semana</span>
        </div>
        """, unsafe_allow_html=True)
        st.markdown(html, unsafe_allow_html=True)

        # ── Panel consultar / editar individual ──
        st.divider()
        st.markdown("### 🔍 Consultar · editar · crear reserva")
        pi_c1, pi_c2 = st.columns([2, 2])
        with pi_c1:
            apto_pi  = st.selectbox("Apartamento", [""] + APTOS, key="pi_apto")
        with pi_c2:
            fecha_pi = st.date_input(
                "Fecha", value=date(anio_sel, mes_n, 1),
                min_value=date(anio_sel, mes_n, 1),
                max_value=date(anio_sel, mes_n, n_dias),
                format="DD/MM/YYYY", key="pi_fecha",
            )
        if apto_pi:
            d_sel = fecha_pi.day
            celda = grid.get(apto_pi, {}).get(d_sel)
            if celda:
                badge = "🔵 Directa" if celda["fuente"] == "DIRECTA" else "🟢 Booking.com"
                st.success(f"**{apto_pi}** — {fecha_pi.strftime('%d/%m/%Y')}: **{celda['nombre']}** &nbsp; {badge}")
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Entrada", celda["entrada"])
                m2.metric("Salida",  celda["salida"])
                m3.metric("Precio",  format_eur(celda['precio']) or "—")
                m4.metric("Estado",  celda["estado_pago"] or "—")
                with st.expander("✏️ Editar esta reserva", expanded=False):
                    r_data = df[df["id"] == celda["id"]]
                    if not r_data.empty:
                        rv = r_data.iloc[0]
                        def parse_d(s):
                            try: return datetime.strptime(str(s), "%d/%m/%Y").date()
                            except: return None
                        with st.form("form_plant_edit"):
                            pe1, pe2 = st.columns(2)
                            with pe1:
                                pf  = st.selectbox("Fuente", FUENTES, index=FUENTES.index(rv["fuente"]) if rv["fuente"] in FUENTES else 0)
                                pn  = st.text_input("Nombre *", value=str(rv.get("nombre","")))
                                pa_opts = [""] + APTOS
                                pa  = st.selectbox("Apartamento", pa_opts, index=pa_opts.index(apto_pi) if apto_pi in pa_opts else 0)
                                pm_v = str(rv.get("mes","")).upper()
                                pm  = st.selectbox("Mes", MESES, index=MESES.index(pm_v) if pm_v in MESES else 0)
                            with pe2:
                                pe_in  = st.date_input("Entrada", value=parse_d(rv.get("entrada")), format="DD/MM/YYYY")
                                pe_out = st.date_input("Salida",  value=parse_d(rv.get("salida")),  format="DD/MM/YYYY")
                                pp  = st.text_input("Precio €", value=str(rv.get("precio","")))
                                pst_v = str(rv.get("estado_pago",""))
                                pst = st.selectbox("Estado pago", ESTADOS, index=ESTADOS.index(pst_v) if pst_v in ESTADOS else 0)
                            pcom = st.text_area("Comentarios", value=str(rv.get("comentarios","")), height=60)
                            sc2, sd2 = st.columns([3,1])
                            with sc2: psave = st.form_submit_button("💾 Guardar", type="primary", use_container_width=True)
                            with sd2: pdel  = st.form_submit_button("🗑️ Eliminar", use_container_width=True)
                        if psave:
                            e2 = pe_in.strftime("%d/%m/%Y")  if pe_in  else ""
                            s2 = pe_out.strftime("%d/%m/%Y") if pe_out else ""
                            actualizar_reserva(celda["id"], {
                                "fuente": pf, "nombre": pn, "apartamento": pa,
                                "mes": pm, "mes_num": mes_num(pm),
                                "entrada": e2, "salida": s2,
                                "noches": (pe_out - pe_in).days if pe_in and pe_out else 0,
                                "precio": pp, "estado_pago": pst, "comentarios": pcom,
                            })
                            st.success("✅ Reserva actualizada.")
                            st.rerun()
                        if pdel:
                            eliminar_reserva(celda["id"])
                            st.success("🗑️ Reserva eliminada.")
                            st.rerun()
            else:
                st.info(f"**{apto_pi}** está libre el {fecha_pi.strftime('%d/%m/%Y')}.")
                with st.expander("➕ Crear reserva aquí", expanded=True):
                    with st.form("form_plant_new", clear_on_submit=True):
                        nn1, nn2 = st.columns(2)
                        with nn1:
                            nf  = st.selectbox("Fuente *", FUENTES, key="nf")
                            nn  = st.text_input("Nombre del cliente *")
                            na_opts = [""] + APTOS
                            na  = st.selectbox("Apartamento *", na_opts, index=na_opts.index(apto_pi) if apto_pi in na_opts else 0)
                            nm  = st.selectbox("Mes *", MESES, index=mes_n - 1)
                        with nn2:
                            ni  = st.date_input("Entrada *", value=fecha_pi, format="DD/MM/YYYY")
                            no  = st.date_input("Salida *",  value=None,     format="DD/MM/YYYY")
                            np_ = st.text_input("Precio €")
                            nst = st.selectbox("Estado pago", ESTADOS, key="nst")
                        ncom = st.text_area("Comentarios", height=60)
                        nsub = st.form_submit_button("💾 Guardar reserva", type="primary", use_container_width=True)
                    if nsub:
                        if not nn:
                            st.error("El nombre es obligatorio.")
                        elif not ni or not no:
                            st.error("Las fechas son obligatorias.")
                        elif no <= ni:
                            st.error("La salida debe ser posterior a la entrada.")
                        else:
                            guardar_reserva({
                                "fuente": nf, "nombre": nn, "apartamento": na,
                                "mes": nm, "mes_num": mes_num(nm),
                                "entrada": ni.strftime("%d/%m/%Y"),
                                "salida":  no.strftime("%d/%m/%Y"),
                                "noches":  (no - ni).days,
                                "precio": np_, "estado_pago": nst, "comentarios": ncom,
                            })
                            st.success(f"✅ Reserva de **{nn}** en **{na}** guardada.")
                            st.rerun()

    # ══════════════════════════════════════════
    # TAB 2: EDITAR CUADRO (tipo Excel)
    # ══════════════════════════════════════════
    with tab_edit:
        st.markdown(
            "<div style='font-size:0.84rem;color:#555;padding:4px 0 10px;'>"
            "✏️ <b>Doble clic</b> en una celda para editar. "
            "<b>Renombra</b> el huésped cambiando el texto. "
            "<b>Mueve</b> una reserva vaciando su celda y escribiendo el nombre en la nueva fila (mismo día). "
            "<b>Crea</b> una nueva reserva escribiendo en una celda vacía. "
            "Pulsa <b>💾 Guardar cambios</b> al terminar.</div>",
            unsafe_allow_html=True,
        )

        # ── Construir DataFrame base desde la BD ───
        rows_edit = []
        for apto in APTOS:
            row = {"Apartamento": apto}
            for d in dias:
                celda = grid[apto][d]
                row[str(d)] = celda["nombre"] if celda else ""
            rows_edit.append(row)
        df_base = pd.DataFrame(rows_edit)   # referencia limpia de la BD actual

        # ── Config columnas ────────────────────────
        col_cfg_edit = {
            "Apartamento": st.column_config.TextColumn("Apartamento", width=160, disabled=True),
        }
        for d in dias:
            wd = (primer_dia + d - 1) % 7
            col_cfg_edit[str(d)] = st.column_config.TextColumn(
                str(d),
                width=85,
                help=f"Día {d} — {DIAS_SEM[wd]}{'  (fin de semana)' if wd >= 5 else ''}",
            )

        # ── Data editor ────────────────────────────
        # La clave cambia tras cada guardado para resetear el estado interno
        edited_grid = st.data_editor(
            df_base,
            use_container_width=True,
            height=600,
            column_config=col_cfg_edit,
            num_rows="fixed",
            disabled=["Apartamento"],
            hide_index=True,
            key=f"ged_{mes_sel}_{anio_sel}_{st.session_state.edit_ver}",
        )

        # ── Botones ────────────────────────────────
        btn1, btn2, _ = st.columns([1, 1, 4])
        with btn1:
            guardar_grid = st.button("💾 Guardar cambios", type="primary", use_container_width=True)
        with btn2:
            if st.button("↺ Descartar", use_container_width=True):
                st.session_state["edit_ver"] += 1
                st.rerun()

        # ── Lógica de guardado ─────────────────────
        if guardar_grid:
            actualizaciones = {}   # rid → {campo: valor}
            clears          = {}   # rid → {"apto": str, "days": [int]}
            nuevos          = {}   # (apto, day) → nombre
            advertencias    = []

            for i, apto in enumerate(APTOS):
                for d in dias:
                    col = str(d)

                    # Valor original (BD actual) y valor editado
                    base_serie   = df_base.iloc[i]
                    edited_serie = edited_grid.iloc[i]
                    old_val = str(base_serie[col]   if col in base_serie.index   else "")
                    new_val = str(edited_serie[col] if col in edited_serie.index else "")
                    # Limpiar None / nan
                    old_val = "" if old_val in ("None", "nan", "NaN") else old_val.strip()
                    new_val = "" if new_val in ("None", "nan", "NaN") else new_val.strip()

                    if old_val == new_val:
                        continue  # sin cambio

                    rid = reverse_map.get((apto, d))

                    if rid is not None and new_val:
                        # ✏️ Renombrar: celda con reserva, nombre cambiado
                        if rid not in actualizaciones:
                            actualizaciones[rid] = {}
                        actualizaciones[rid]["nombre"] = new_val

                    elif rid is not None and not new_val:
                        # ⬜ Vaciada: posible inicio de movimiento
                        if rid not in clears:
                            clears[rid] = {"apto": apto, "days": []}
                        clears[rid]["days"].append(d)

                    elif rid is None and new_val:
                        # 🆕 Texto nuevo en celda vacía: posible destino de movimiento o reserva nueva
                        nuevos[(apto, d)] = new_val

            # ── Detectar movimientos ───────────────
            # Si una celda se vació en apto A y en el mismo día apareció texto en apto B → movimiento
            for rid, info in clears.items():
                old_apto     = info["apto"]
                cleared_days = info["days"]
                matched_apto = None

                for d in cleared_days:
                    for (na, nd) in list(nuevos.keys()):
                        if nd == d and na != old_apto:
                            matched_apto = na
                            break
                    if matched_apto:
                        break

                if matched_apto:
                    # Mover reserva a otro apartamento
                    if rid not in actualizaciones:
                        actualizaciones[rid] = {}
                    actualizaciones[rid]["apartamento"] = matched_apto
                    # Eliminar del dict nuevos los días del movimiento (evitar creación duplicada)
                    for d in cleared_days:
                        nuevos.pop((matched_apto, d), None)
                else:
                    dias_str = (f"días {min(cleared_days)}–{max(cleared_days)}"
                                if len(cleared_days) > 1 else f"día {cleared_days[0]}")
                    advertencias.append(
                        f"⚠️ Celda vaciada en **{old_apto}** ({dias_str}). "
                        "Si quieres eliminar la reserva usa '✏️ Editar reserva'."
                    )

            # ── Crear reservas de 1 día para entradas nuevas sin match ──
            creaciones = []
            for (apto, d), nombre in nuevos.items():
                f_ent = date(anio_sel, mes_n, d)
                f_sal = f_ent + timedelta(days=1)
                creaciones.append({
                    "fuente":      "DIRECTA",
                    "nombre":      nombre,
                    "apartamento": apto,
                    "mes":         mes_sel,
                    "mes_num":     mes_n,
                    "entrada":     f_ent.strftime("%d/%m/%Y"),
                    "salida":      f_sal.strftime("%d/%m/%Y"),
                    "noches":      1,
                    "estado_pago": "",
                    "comentarios": "",
                })

            # ── Aplicar ────────────────────────────
            total = 0
            for rid, datos in actualizaciones.items():
                actualizar_reserva(rid, datos)
                total += 1
            for datos in creaciones:
                guardar_reserva(datos)
                total += 1

            for adv in advertencias:
                st.warning(adv)

            if total:
                st.success(f"✅ {total} cambio(s) guardados correctamente.")
                st.session_state["edit_ver"] += 1   # resetea el data_editor con datos frescos
                st.rerun()
            elif not advertencias:
                st.info("No hay cambios que guardar.")

# ─────────────────────────────────────────────
# SECCIÓN: IMPORTAR BOOKING
# ─────────────────────────────────────────────
elif seccion == "📥 Importar Booking":
    st.markdown("### 📥 Importar reservas desde Booking.com")
    st.markdown("Sube el Excel de **Check-in** que descarga Booking.com y se importarán automáticamente las reservas nuevas.")

    # ── Zona peligrosa: borrar toda la BD ─────────────────────────────
    with st.expander("🗑️ Borrar base de datos completa", expanded=False):
        n_reservas = len(df) if not df.empty else 0
        st.error(
            f"⚠️ **Acción irreversible.** Se eliminarán las **{n_reservas} reservas** "
            f"almacenadas actualmente. Úsalo solo para empezar desde cero antes de una importación nueva."
        )
        if not st.session_state.get("_confirm_borrar_bd", False):
            if st.button("🗑️ Borrar toda la base de datos", use_container_width=True):
                st.session_state["_confirm_borrar_bd"] = True
                st.rerun()
        else:
            st.warning("¿Estás SEGURO? No hay vuelta atrás.")
            cb1, cb2 = st.columns(2)
            if cb1.button("✅ Sí, borrar TODO ahora", type="primary", use_container_width=True, key="btn_confirm_borrar"):
                borrar_todas_las_reservas()
                st.session_state["_confirm_borrar_bd"] = False
                st.cache_data.clear()
                st.success("✅ Base de datos vaciada correctamente. Ya puedes importar desde cero.")
                st.rerun()
            if cb2.button("❌ Cancelar", use_container_width=True, key="btn_cancel_borrar"):
                st.session_state["_confirm_borrar_bd"] = False
                st.rerun()

    st.markdown("---")
    archivo = st.file_uploader("Selecciona el archivo Excel de Booking.com", type=["xls","xlsx"], key="bk_upload")

    if archivo:
        try:
            bk = pd.read_excel(archivo, header=0)
            bk.columns = [str(c).strip() for c in bk.columns]

            # Mapeo de columnas Booking → nuestra BD
            COL_MAP = {
                "nro_reserva":   ["Número de reserva", "Numero de reserva"],
                "nombre":        ["Nombre del cliente (o clientes)", "Nombre del cliente"],
                "entrada":       ["Entrada"],
                "salida":        ["Salida"],
                "fecha_reserva": ["Fecha de reserva", "Booking date", "Reservation date"],
                "noches":        ["Duración (noches)", "Duracion (noches)"],
                "personas":      ["Personas", "Adultos"],
                "habitaciones":  ["Habitaciones", "Rooms", "Unidades", "Nº de habitaciones",
                                  "Numero de habitaciones", "Número de habitaciones"],
                "precio":        ["Precio"],
                "estado_pago":   ["Estado del pago", "Estado de pago",
                                  "Payment status"],
                "estado_reserva":["Estado", "Status", "Booking status",
                                  "Reservation status", "Estado reserva"],
                "fecha_cancel":  ["Fecha de cancelación", "Fecha de cancelacion",
                                  "Cancellation date"],
                "comentarios":   ["Comentarios"],
                "tipo_unidad":   ["Tipo de unidad", "Tipo de habitación", "Room type",
                                  "Tipo de alojamiento"],
            }

            def get_col(df_bk, opciones):
                # 1) igualdad exacta (case-insensitive)
                for op in opciones:
                    for c in df_bk.columns:
                        if op.lower() == c.lower():
                            return c
                # 2) substring (fallback tolerante)
                for op in opciones:
                    for c in df_bk.columns:
                        if op.lower() in c.lower():
                            return c
                return None

            def limpiar_precio(v):
                try:
                    return str(v).replace("EUR","").replace("€","").strip().replace(",",".")
                except:
                    return ""

            def fmt_fecha(v):
                """Acepta '2026-05-25', '2026-05-18 09:14:02', Timestamps, etc.
                Devuelve dd/mm/yyyy o '' si v es vacío/None/NaN."""
                if v is None or v == "" or (isinstance(v, float) and pd.isna(v)):
                    return ""
                try:
                    if isinstance(v, str):
                        d = datetime.strptime(v.strip()[:10], "%Y-%m-%d")
                    else:
                        d = pd.Timestamp(v).to_pydatetime()
                    return d.strftime("%d/%m/%Y")
                except Exception:
                    return ""

            canceladas_excel = []   # filas del Excel marcadas como canceladas

            # ── Construir filas con auto-asignación de apartamento ──────────
            filas = []
            # df_asignados: base de reservas activas para detectar conflictos.
            # Solo incluye filas con apartamento asignado y fechas parseables.
            # Se va ampliando con cada asignación del batch actual.
            if not df.empty and "apartamento" in df.columns:
                _mask_valid = (
                    df["apartamento"].notna() &
                    (df["apartamento"].astype(str).str.strip() != "")
                )
                df_asignados = df[_mask_valid][["apartamento","entrada","salida"]].copy()
            else:
                df_asignados = pd.DataFrame(columns=["apartamento","entrada","salida"])

            for _, row in bk.iterrows():
                def g(key):
                    c = get_col(bk, COL_MAP.get(key, []))
                    return row[c] if c and not pd.isna(row.get(c, float("nan"))) else ""

                nro = str(g("nro_reserva")).strip()
                if not nro or nro in ("nan", ""):
                    continue

                # ── Filtrar canceladas ANTES de procesar ─────────────────
                # Booking marca canceladas en la columna "Estado" (cancelled_by_guest…),
                # NO en "Estado del pago" (que viene vacío). También miramos "Fecha de
                # cancelación": si tiene valor, la reserva está cancelada.
                estado_pago_raw     = str(g("estado_pago")).strip()
                estado_reserva_raw  = str(g("estado_reserva")).strip()
                fecha_cancel_raw    = str(g("fecha_cancel")).strip()
                tiene_fecha_cancel  = bool(fecha_cancel_raw) and fecha_cancel_raw.lower() not in ("nan", "none", "")
                if (es_cancelada(estado_reserva_raw)
                        or es_cancelada(estado_pago_raw)
                        or tiene_fecha_cancel):
                    canceladas_excel.append({
                        "nro_reserva": nro,
                        "nombre":      str(g("nombre")).strip().title(),
                        "entrada":     fmt_fecha(g("entrada")),
                        "salida":      fmt_fecha(g("salida")),
                        "estado":      estado_reserva_raw or estado_pago_raw or "cancelled",
                    })
                    continue   # ← no importar reservas canceladas

                # Para el resto del procesamiento, conservar nombre interno
                estado_raw = estado_pago_raw

                entrada_str    = fmt_fecha(g("entrada"))
                salida_str     = fmt_fecha(g("salida"))
                fecha_reserva  = fmt_fecha(g("fecha_reserva"))

                try:
                    e_date  = datetime.strptime(entrada_str, "%d/%m/%Y")
                    mes_n2  = e_date.month
                    mes_str = MESES[mes_n2 - 1]
                    f_ent   = e_date.date()
                    f_sal   = datetime.strptime(salida_str, "%d/%m/%Y").date()
                except:
                    mes_n2, mes_str, f_ent, f_sal = 0, "", None, None

                noches_raw = g("noches")
                try:
                    noches_val = int(float(str(noches_raw))) if noches_raw != "" else 0
                except:
                    noches_val = calcular_noches(entrada_str, salida_str)

                precio_raw   = limpiar_precio(g("precio"))
                precio_total = parse_eur(precio_raw)   # float o None

                # estado_raw ya leído arriba; traducir al valor interno
                # Criterio entrevista: "Pago mediante Booking.com" → "PAGADO";
                # vacío → "PENDIENTE". Mantiene la convención mayúscula del resto
                # de la app (ver lista ESTADOS).
                if estado_raw and "booking" in estado_raw.lower():
                    estado_val = "PAGADO"
                    pagado     = True
                elif estado_raw.lower() in ("ok", "pagado"):
                    estado_val = "PAGADO"
                    pagado     = True
                elif not estado_raw or estado_raw.lower() in ("nan", "none", ""):
                    estado_val = "PENDIENTE"
                    pagado     = False
                else:
                    estado_val = estado_raw
                    pagado     = False

                # Número de habitaciones reservadas
                try:
                    n_hab = max(1, int(float(str(g("habitaciones")))))
                except:
                    n_hab = 1

                # ── Tipo de unidad → puede traer varios separados por coma ─────
                # Ej: "Two-Bedroom Apartment, One-Bedroom Apartment". Si la lista
                # de tipos es menor que n_hab, repetimos el último para cubrir.
                tipo_unidad_raw = str(g("tipo_unidad")).strip()
                tipos_unidad = [t.strip() for t in tipo_unidad_raw.split(",") if t.strip()]
                if not tipos_unidad:
                    tipos_unidad = [""]
                while len(tipos_unidad) < n_hab:
                    tipos_unidad.append(tipos_unidad[-1])

                # Personas totales (las repartimos: total en fila 1, 0 en las demás)
                personas_raw = str(g("personas")).replace(".0", "").strip()
                try:
                    personas_total = int(float(personas_raw)) if personas_raw not in ("", "nan") else 0
                except Exception:
                    personas_total = 0

                base = {
                    "nro_reserva": nro,
                    "fuente":      "BOOKING.COM",
                    "nombre":      str(g("nombre")).strip().title(),
                    "mes":         mes_str,
                    "mes_num":     mes_n2,
                    "entrada":     entrada_str,
                    "salida":      salida_str,
                    "noches":      noches_val,
                    "estado_pago": estado_val,
                    "comentarios": str(g("comentarios")) if g("comentarios") else "",
                }

                # ── Una fila por habitación ─────────────────────────────────
                # Cada fila se asigna individualmente para respetar tipos distintos.
                # idx_hab == 0 lleva valores reales (precio, personas, pago…);
                # las siguientes llevan "0,00 €" y 0 personas (criterio entrevista).
                for idx_hab in range(n_hab):
                    tipo_unit_i = tipos_unidad[idx_hab]
                    apto_directo_i = match_apto_directo(tipo_unit_i)
                    if apto_directo_i:
                        tipo_dorm_i = dorm_desde_nombre_apto(apto_directo_i)
                    else:
                        tipo_dorm_i = clasificar_dormitorios(tipo_unit_i)

                    # Asignar apartamento concreto
                    if apto_directo_i and f_ent and f_sal:
                        apto_i = apto_directo_i if apto_libre(apto_directo_i, f_ent, f_sal, df_asignados) else ""
                        if not apto_i:
                            libres = asignar_aptos_auto(tipo_dorm_i, f_ent, f_sal, 1, df_asignados)
                            apto_i = libres[0] if libres else ""
                    elif f_ent and f_sal:
                        libres = asignar_aptos_auto(tipo_dorm_i, f_ent, f_sal, 1, df_asignados)
                        apto_i = libres[0] if libres else (apto_directo_i or "")
                    else:
                        apto_i = apto_directo_i or ""

                    # Valores que dependen de la posición de la fila
                    es_primera = (idx_hab == 0)
                    if es_primera:
                        precio_fila   = format_eur(precio_total) if precio_total is not None else ""
                        personas_fila = str(personas_total) if personas_total else ""
                        if pagado:
                            pago_cta_fila   = format_eur(precio_total) if precio_total is not None else ""
                            resto_pdte_fila = "0,00 €"
                            fecha_ing_fila  = fecha_reserva
                        else:
                            pago_cta_fila   = ""
                            resto_pdte_fila = ""
                            fecha_ing_fila  = ""
                    else:
                        precio_fila     = "0,00 €"
                        personas_fila   = "0"
                        pago_cta_fila   = "0,00 €" if pagado else ""
                        resto_pdte_fila = "0,00 €" if pagado else ""
                        fecha_ing_fila  = fecha_reserva if pagado else ""

                    nro_ext = f"{nro}-{idx_hab+1}" if n_hab > 1 else nro
                    fila = {**base,
                            "nro_reserva":   nro_ext,
                            "apartamento":   apto_i,
                            "dormitorios":   tipo_dorm_i,
                            "precio":        precio_fila,
                            "personas":      personas_fila,
                            "pago_cta":      pago_cta_fila,
                            "resto_pdte":    resto_pdte_fila,
                            "fecha_ingreso": fecha_ing_fila}
                    filas.append(fila)
                    # Registrar en df_asignados para que el siguiente apto del batch lo tenga en cuenta
                    if apto_i:
                        nuevo_reg = pd.DataFrame([{
                            "apartamento": apto_i,
                            "entrada":     entrada_str,
                            "salida":      salida_str,
                        }])
                        df_asignados = pd.concat([df_asignados, nuevo_reg], ignore_index=True)

            df_bk = pd.DataFrame(filas) if filas else pd.DataFrame()

            # Detectar duplicados (nro_reserva ya en BD)
            nros_bd = set(str(r) for r in df["nro_reserva"].tolist()) if not df.empty else set()
            if not df_bk.empty:
                df_bk["_nuevo"] = ~df_bk["nro_reserva"].astype(str).isin(nros_bd)
                nuevas   = df_bk[df_bk["_nuevo"]].drop(columns=["_nuevo"])
                ya_exist = df_bk[~df_bk["_nuevo"]].drop(columns=["_nuevo"])
            else:
                nuevas   = pd.DataFrame()
                ya_exist = pd.DataFrame()

            # ── Sobrescritura selectiva: detectar cuáles ya existentes han cambiado ─
            # Solo se comparan/actualizan las columnas "de Booking". Los campos manuales
            # (pago_cta, fecha_ingreso, resto_pdte, estado_pago, comentarios, apartamento)
            # se respetan SIEMPRE en la reserva existente.
            CAMPOS_UPDATE = ["nombre", "entrada", "salida", "noches", "personas",
                             "precio", "dormitorios", "mes", "mes_num"]

            def _norm(campo, v):
                """Normaliza para comparar entre BD y nuevo import."""
                if v is None:
                    return ""
                if campo == "precio":
                    f = parse_eur(v)
                    return round(f, 2) if f is not None else ""
                if campo in ("noches", "mes_num", "personas"):
                    try:
                        s = str(v).strip().replace(".0", "")
                        return int(float(s)) if s and s.lower() != "nan" else 0
                    except Exception:
                        return 0
                return str(v).strip()

            updates_plan = []   # [{ "id": ..., "cambios": {campo: (antes, ahora)}, "fila_nueva": ... }]
            sin_cambios  = []   # filas que ya existen y coinciden
            if not ya_exist.empty and not df.empty:
                df_idx = df.set_index(df["nro_reserva"].astype(str))
                for _, fila_nueva in ya_exist.iterrows():
                    nro_e = str(fila_nueva["nro_reserva"])
                    if nro_e not in df_idx.index:
                        continue
                    fila_bd = df_idx.loc[nro_e]
                    if isinstance(fila_bd, pd.DataFrame):
                        fila_bd = fila_bd.iloc[0]   # por si hay varias filas con el mismo Nº
                    cambios = {}
                    for c in CAMPOS_UPDATE:
                        antes = _norm(c, fila_bd.get(c))
                        ahora = _norm(c, fila_nueva.get(c))
                        if antes != ahora:
                            cambios[c] = (antes, ahora)
                    if cambios:
                        updates_plan.append({
                            "id":          int(fila_bd["id"]),
                            "nro_reserva": nro_e,
                            "nombre":      fila_bd.get("nombre", ""),
                            "cambios":     cambios,
                            "fila_nueva":  fila_nueva,
                        })
                    else:
                        sin_cambios.append(fila_nueva)
            else:
                sin_cambios = [r for _, r in ya_exist.iterrows()]

            # ── Canceladas del Excel que ya están guardadas en la BD ───────
            canceladas_en_bd = []
            if canceladas_excel and not df.empty:
                nros_cancel = {c["nro_reserva"] for c in canceladas_excel}
                for _, r in df.iterrows():
                    stored = str(r.get("nro_reserva", ""))
                    # match exacto O sufijo -1/-2 (reservas multi-habitación)
                    base   = re.sub(r'-\d+$', '', stored)
                    if stored in nros_cancel or base in nros_cancel:
                        canceladas_en_bd.append(r)

            # ── Resumen métricas ────────────────────────────────────────────
            col_r1, col_r2, col_r3, col_r4, col_r5 = st.columns(5)
            col_r1.metric("Válidas en el archivo",     len(df_bk))
            col_r2.metric("✅ Nuevas a importar",       len(nuevas),  delta=f"+{len(nuevas)}")
            col_r3.metric("🔄 Con cambios",             len(updates_plan))
            col_r4.metric("= Sin cambios",              len(sin_cambios))
            col_r5.metric("🚫 Canceladas (excluidas)",  len(canceladas_excel))

            # ── Panel de canceladas que están en la BD ──────────────────────
            if canceladas_en_bd:
                df_cancel_bd = pd.DataFrame(canceladas_en_bd)
                st.markdown("---")
                st.error(
                    f"🚨 **{len(df_cancel_bd)} reserva(s)** ya guardadas en la aplicación "
                    f"aparecen ahora como **CANCELADAS** en Booking.com. ¿Deseas eliminarlas?"
                )
                cols_c = [c for c in ["nro_reserva","nombre","apartamento","entrada","salida"]
                          if c in df_cancel_bd.columns]
                st.dataframe(df_cancel_bd[cols_c], use_container_width=True, hide_index=True,
                             column_config={
                                 "nro_reserva": st.column_config.TextColumn("Nº Reserva",  width=130),
                                 "nombre":       st.column_config.TextColumn("Nombre",       width=190),
                                 "apartamento":  st.column_config.TextColumn("Apartamento",  width=175),
                                 "entrada":      st.column_config.TextColumn("Entrada",       width=90),
                                 "salida":       st.column_config.TextColumn("Salida",        width=90),
                             })
                c_si, c_no = st.columns(2)
                ids_eliminar = [int(r["id"]) for _, r in df_cancel_bd.iterrows()]
                if c_si.button("🗑️ Sí, eliminar reservas canceladas",
                               type="primary", use_container_width=True, key="btn_del_cancel"):
                    for rid in ids_eliminar:
                        eliminar_reserva(rid)
                    st.success(f"✅ {len(ids_eliminar)} reserva(s) cancelada(s) eliminadas.")
                    st.rerun()
                if c_no.button("Mantener en la aplicación",
                               use_container_width=True, key="btn_keep_cancel"):
                    st.info("Las reservas canceladas se han mantenido en la aplicación.")
            elif canceladas_excel:
                st.info(
                    f"ℹ️ {len(canceladas_excel)} reserva(s) cancelada(s) en el archivo "
                    f"— ninguna estaba guardada en la aplicación."
                )

            # ── Panel de reservas existentes con cambios ────────────────────
            if updates_plan:
                st.markdown("---")
                st.warning(
                    f"🔄 **{len(updates_plan)} reserva(s)** ya guardadas tienen cambios en Booking.com. "
                    f"Solo se actualizarán las columnas de Booking — los campos manuales "
                    f"(pago a cuenta, fecha de ingreso, resto pendiente, estado de pago, comentarios, "
                    f"apartamento asignado) se respetarán."
                )
                resumen_updates = []
                for u in updates_plan:
                    descripcion = " · ".join(
                        f"{k}: {a!r} → {n!r}" for k, (a, n) in u["cambios"].items()
                    )
                    resumen_updates.append({
                        "Nº Reserva":  u["nro_reserva"],
                        "Nombre":      u["nombre"],
                        "Cambios":     descripcion,
                    })
                st.dataframe(pd.DataFrame(resumen_updates),
                             use_container_width=True, hide_index=True,
                             column_config={
                                 "Nº Reserva": st.column_config.TextColumn(width=130),
                                 "Nombre":     st.column_config.TextColumn(width=190),
                                 "Cambios":    st.column_config.TextColumn(width=600),
                             })
                if st.button(f"🔄 Aplicar cambios a {len(updates_plan)} reserva(s)",
                             type="primary", use_container_width=True, key="btn_apply_updates"):
                    aplicados = 0
                    errores_upd = []
                    for u in updates_plan:
                        try:
                            payload = {c: u["fila_nueva"].get(c) for c in CAMPOS_UPDATE
                                       if c in u["fila_nueva"]}
                            actualizar_reserva(u["id"], payload)
                            aplicados += 1
                        except Exception as ex:
                            errores_upd.append(f"{u['nro_reserva']}: {ex}")
                    if aplicados:
                        st.success(f"✅ {aplicados} reserva(s) actualizada(s).")
                        st.rerun()
                    for err in errores_upd:
                        st.error(f"Error al actualizar: {err}")

            # ── Vista previa editable ─────────────────────────────────
            if not nuevas.empty:
                st.markdown("#### Vista previa de reservas nuevas")

                # Aviso si alguna fila quedó sin apartamento asignado
                sin_apto = nuevas[nuevas["apartamento"].astype(str).str.strip() == ""]
                if not sin_apto.empty:
                    st.warning(
                        f"⚠️ **{len(sin_apto)} reserva(s) sin apartamento** — no había disponible del tipo "
                        f"requerido. Selecciona uno manualmente en la columna **Apartamento ✏️** antes de importar."
                    )
                else:
                    st.success("✅ Todos los apartamentos asignados correctamente. Revisa y confirma.")

                cols_edit = ["apartamento","nro_reserva","nombre","entrada","salida",
                             "noches","personas","precio","estado_pago"]
                df_edit = nuevas[[c for c in cols_edit if c in nuevas.columns]].reset_index(drop=True)

                edited_preview = st.data_editor(
                    df_edit,
                    use_container_width=True,
                    height=min(60 + 35 * len(df_edit), 440),
                    hide_index=True,
                    column_config={
                        "apartamento": st.column_config.SelectboxColumn(
                            "Apartamento ✏️",
                            options=[""] + APTOS,
                            width=185,
                        ),
                        "nro_reserva": st.column_config.TextColumn("Nº Reserva",  width=130, disabled=True),
                        "nombre":      st.column_config.TextColumn("Nombre",       width=175, disabled=True),
                        "entrada":     st.column_config.TextColumn("Entrada",      width=90,  disabled=True),
                        "salida":      st.column_config.TextColumn("Salida",       width=90,  disabled=True),
                        "noches":      st.column_config.NumberColumn("Noches",     width=62,  disabled=True),
                        "personas":    st.column_config.TextColumn("Pers.",        width=55,  disabled=True),
                        "precio":      st.column_config.TextColumn("Precio €",     width=90,  disabled=True),
                        "estado_pago": st.column_config.TextColumn("Estado pago",  width=175, disabled=True),
                    },
                    num_rows="fixed",
                    key="preview_import_editor",
                )
                st.caption("💡 La columna **Apartamento ✏️** es editable — puedes cambiar cualquier asignación antes de importar.")

                st.markdown("")
                if st.button(f"📥 Importar {len(nuevas)} reserva(s) nueva(s)", type="primary", use_container_width=True):
                    # Aplicar cambios manuales del editor
                    nuevas_import = nuevas.copy().reset_index(drop=True)
                    nuevas_import["apartamento"] = edited_preview["apartamento"].values

                    # Validación final: cargar BD fresca y verificar disponibilidad
                    df_fresh    = cargar_reservas()
                    importadas  = 0
                    conflictos  = []
                    errores_imp = []

                    for _, row in nuevas_import.iterrows():
                        apto = str(row.get("apartamento", "")).strip()
                        f_e  = parse_date_safe(row.get("entrada", ""))
                        f_s  = parse_date_safe(row.get("salida",  ""))

                        # Verificación de conflicto antes de insertar
                        if apto and f_e and f_s:
                            if not apto_libre(apto, f_e, f_s, df_fresh):
                                conflictos.append(
                                    f"**{apto}** · {row.get('entrada','')} → {row.get('salida','')} "
                                    f"({row.get('nombre','')})"
                                )
                                continue   # NO importar esta fila

                        try:
                            guardar_reserva(row.to_dict())
                            importadas += 1
                            # Añadir a df_fresh para que el próximo check la tenga en cuenta
                            if apto and f_e and f_s:
                                df_fresh = pd.concat([df_fresh, pd.DataFrame([{
                                    "apartamento": apto,
                                    "entrada":     row.get("entrada",""),
                                    "salida":      row.get("salida",""),
                                }])], ignore_index=True)
                        except Exception as ex:
                            errores_imp.append(f"{row.get('nro_reserva','?')}: {ex}")

                    if conflictos:
                        st.error(
                            f"⛔ **{len(conflictos)} reserva(s) con conflicto de disponibilidad** "
                            f"(no se importaron):\n\n" + "\n\n".join(f"- {c}" for c in conflictos)
                        )
                    if importadas:
                        st.success(f"✅ {importadas} reserva(s) importadas correctamente.")
                        st.rerun()
                    for err in errores_imp:
                        st.error(f"Error al guardar: {err}")
            else:
                st.info("✅ Todas las reservas del archivo ya están en la base de datos. No hay nada nuevo que importar.")

            if not ya_exist.empty:
                with st.expander(f"Ver {len(ya_exist)} reserva(s) ya existentes"):
                    st.dataframe(
                        ya_exist[["nro_reserva","nombre","entrada","salida"]],
                        use_container_width=True, hide_index=True,
                    )

        except Exception as e:
            st.error(f"Error al procesar el archivo: {e}")

# ─────────────────────────────────────────────
# SECCIÓN: USUARIOS (solo admins)
# ─────────────────────────────────────────────
elif seccion == "👥 Usuarios":
    if not IS_ADMIN:
        st.error("⛔ No tienes permiso para acceder a esta sección.")
        st.stop()

    st.markdown("### 👥 Gestión de usuarios")
    st.caption(
        "Da de alta, edita o elimina usuarios. Solo los administradores ven esta sección. "
        "Los administradores de rescate (los definidos en Streamlit Cloud → Misterios) no se "
        "tocan desde aquí — para modificarlos hay que editar los Secrets."
    )

    usuarios_bd = cargar_usuarios_bd()
    usernames_bd = {u["username"] for u in usuarios_bd}

    # Métricas rápidas
    n_total   = len(usuarios_bd) + len(BOOTSTRAP_ADMINS)
    n_admins  = sum(1 for u in usuarios_bd if u.get("rol") == "admin" and u.get("activo", True)) + len(BOOTSTRAP_ADMINS)
    n_activos = sum(1 for u in usuarios_bd if u.get("activo", True)) + len(BOOTSTRAP_ADMINS)
    c1, c2, c3 = st.columns(3)
    c1.metric("Usuarios totales", n_total)
    c2.metric("Administradores",  n_admins)
    c3.metric("Activos",          n_activos)

    tab_lista, tab_alta = st.tabs(["📋 Lista de usuarios", "➕ Dar de alta"])

    # ── TAB: LISTA ─────────────────────────────
    with tab_lista:
        if BOOTSTRAP_ADMINS:
            st.markdown("**🛡️ Administradores de rescate** (configurados en Misterios)")
            for u in sorted(BOOTSTRAP_ADMINS):
                marca = "  · ⭐ tú" if u == USER_USERNAME else ""
                st.markdown(f"- `{u}`{marca}")
            st.caption(
                "Para editar o eliminar estos usuarios, ve a Streamlit Cloud → "
                "Settings → Misterios."
            )
            st.divider()

        st.markdown("**👤 Usuarios en base de datos**")
        if not usuarios_bd:
            st.info(
                "No hay usuarios en la BD todavía. Crea el primero en la "
                "pestaña **'Dar de alta'**."
            )
        else:
            for u in usuarios_bd:
                emoji_estado = "🟢" if u.get("activo", True) else "🔴"
                rol_lbl = (u.get("rol") or "usuario").upper()
                titulo  = (f"{emoji_estado} **{u['username']}** — "
                           f"{u.get('nombre') or '(sin nombre)'} · "
                           f"{rol_lbl}")
                if u["username"] == USER_USERNAME:
                    titulo += "  · ⭐ tú"

                with st.expander(titulo):
                    col_a, col_b = st.columns(2)
                    with col_a:
                        nuevo_nombre = st.text_input(
                            "Nombre", value=u.get("nombre") or "",
                            key=f"nom_{u['id']}",
                        )
                        nuevo_email  = st.text_input(
                            "Email", value=u.get("email") or "",
                            key=f"em_{u['id']}",
                        )
                    with col_b:
                        rol_actual = (u.get("rol") or "usuario")
                        nuevo_rol = st.selectbox(
                            "Rol", ["usuario", "admin"],
                            index=0 if rol_actual == "usuario" else 1,
                            key=f"rol_{u['id']}",
                        )
                        nuevo_activo = st.toggle(
                            "Activo", value=u.get("activo", True),
                            key=f"act_{u['id']}",
                        )

                    if st.button("💾 Guardar cambios",
                                 key=f"save_{u['id']}",
                                 use_container_width=True):
                        cambios = {}
                        if nuevo_nombre != (u.get("nombre") or ""):
                            cambios["nombre"] = nuevo_nombre
                        if nuevo_email != (u.get("email") or ""):
                            cambios["email"] = nuevo_email
                        if nuevo_rol != rol_actual:
                            cambios["rol"] = nuevo_rol
                        if nuevo_activo != u.get("activo", True):
                            cambios["activo"] = nuevo_activo

                        # Protección: que no se quede la app sin admins activos
                        admins_restantes = (
                            sum(1 for x in usuarios_bd
                                if x["id"] != u["id"]
                                and x.get("rol") == "admin"
                                and x.get("activo", True))
                            + len(BOOTSTRAP_ADMINS)
                        )
                        peligro = (
                            (rol_actual == "admin" and cambios.get("rol") == "usuario") or
                            (u.get("activo", True) and cambios.get("activo") is False and rol_actual == "admin")
                        )
                        if peligro and admins_restantes == 0:
                            st.error(
                                "⛔ No puedes hacer este cambio: dejaría la app sin ningún "
                                "administrador. Crea otro admin primero."
                            )
                        elif cambios:
                            try:
                                actualizar_usuario_bd(u["id"], cambios)
                                st.success("✅ Cambios guardados.")
                                st.rerun()
                            except Exception as ex:
                                st.error(f"Error al guardar: {ex}")
                        else:
                            st.info("No hay cambios que guardar.")

                    st.markdown("---")
                    st.markdown("**🔑 Cambiar contraseña**")
                    new_pwd1 = st.text_input(
                        "Nueva contraseña", type="password",
                        key=f"pwd1_{u['id']}",
                    )
                    new_pwd2 = st.text_input(
                        "Repite la contraseña", type="password",
                        key=f"pwd2_{u['id']}",
                    )
                    if st.button("🔑 Cambiar contraseña",
                                 key=f"chpwd_{u['id']}",
                                 use_container_width=True):
                        if not new_pwd1:
                            st.warning("La contraseña no puede estar vacía.")
                        elif new_pwd1 != new_pwd2:
                            st.warning("Las dos contraseñas no coinciden.")
                        elif len(new_pwd1) < 6:
                            st.warning("La contraseña debe tener al menos 6 caracteres.")
                        else:
                            try:
                                cambiar_password_usuario_bd(u["id"], new_pwd1)
                                st.success("✅ Contraseña actualizada.")
                            except Exception as ex:
                                st.error(f"Error: {ex}")

                    st.markdown("---")
                    admins_si_borro = (
                        sum(1 for x in usuarios_bd
                            if x["id"] != u["id"]
                            and x.get("rol") == "admin"
                            and x.get("activo", True))
                        + len(BOOTSTRAP_ADMINS)
                    )
                    if u["username"] == USER_USERNAME:
                        st.info("No puedes eliminarte a ti mismo.")
                    elif rol_actual == "admin" and admins_si_borro == 0:
                        st.info(
                            "No puedes eliminar al único administrador activo. "
                            "Crea otro admin primero."
                        )
                    else:
                        confirm_key = f"_confirm_del_user_{u['id']}"
                        if not st.session_state.get(confirm_key):
                            if st.button(
                                f"🗑️ Eliminar usuario {u['username']}",
                                key=f"del_{u['id']}",
                                use_container_width=True,
                            ):
                                st.session_state[confirm_key] = True
                                st.rerun()
                        else:
                            st.warning(
                                f"¿Seguro que quieres eliminar a **{u['username']}**? "
                                "Esta acción no se puede deshacer."
                            )
                            col_si, col_no = st.columns(2)
                            if col_si.button("✅ Sí, eliminar",
                                             type="primary",
                                             use_container_width=True,
                                             key=f"delok_{u['id']}"):
                                try:
                                    eliminar_usuario_bd(u["id"])
                                    st.session_state[confirm_key] = False
                                    st.success(f"✅ Usuario {u['username']} eliminado.")
                                    st.rerun()
                                except Exception as ex:
                                    st.error(f"Error: {ex}")
                            if col_no.button("Cancelar",
                                             use_container_width=True,
                                             key=f"delno_{u['id']}"):
                                st.session_state[confirm_key] = False
                                st.rerun()

    # ── TAB: DAR DE ALTA ──────────────────────
    with tab_alta:
        st.markdown("**Crear un usuario nuevo**")
        with st.form("crear_usuario_form", clear_on_submit=True):
            col_x, col_y = st.columns(2)
            with col_x:
                new_username = st.text_input(
                    "Usuario", placeholder="ej. juana",
                    help="En minúsculas, sin espacios. Solo letras, números y _."
                )
                new_nombre = st.text_input(
                    "Nombre completo", placeholder="Juana López",
                )
            with col_y:
                new_email = st.text_input(
                    "Email", placeholder="juana@ejemplo.com",
                )
                new_rol = st.selectbox(
                    "Rol", ["usuario", "admin"],
                    help="'admin' puede entrar a esta sección y gestionar usuarios.",
                )

            col_p1, col_p2 = st.columns(2)
            with col_p1:
                new_pwd_a = st.text_input(
                    "Contraseña", type="password",
                    help="Mínimo 6 caracteres.",
                )
            with col_p2:
                new_pwd_b = st.text_input(
                    "Repite la contraseña", type="password",
                )

            submitted = st.form_submit_button(
                "➕ Crear usuario",
                type="primary",
                use_container_width=True,
            )
            if submitted:
                username_clean = (new_username or "").strip().lower()
                if not username_clean:
                    st.error("El usuario es obligatorio.")
                elif not re.match(r"^[a-z0-9_]+$", username_clean):
                    st.error(
                        "El usuario solo puede contener letras minúsculas, números y _."
                    )
                elif username_clean in usernames_bd or username_clean in BOOTSTRAP_ADMINS:
                    st.error(f"El usuario '{username_clean}' ya existe.")
                elif not new_pwd_a:
                    st.error("La contraseña es obligatoria.")
                elif new_pwd_a != new_pwd_b:
                    st.error("Las dos contraseñas no coinciden.")
                elif len(new_pwd_a) < 6:
                    st.error("La contraseña debe tener al menos 6 caracteres.")
                else:
                    try:
                        crear_usuario_bd(
                            username=username_clean,
                            nombre=(new_nombre or "").strip(),
                            email=(new_email or "").strip(),
                            password_plain=new_pwd_a,
                            rol=new_rol,
                        )
                        st.success(
                            f"✅ Usuario **{username_clean}** creado. "
                            "Ya puede iniciar sesión con la contraseña que has indicado."
                        )
                        st.balloons()
                    except Exception as ex:
                        st.error(f"Error al crear usuario: {ex}")
