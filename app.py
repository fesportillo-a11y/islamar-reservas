import streamlit as st
import pandas as pd
from supabase import create_client, Client
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, date
import calendar
import re

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
    "APTO 2", "APTO 9", "APTO 10", "APTO 109",
    "APTO 14- 2 DORM", "APTO 15- 2 DORM", "APTO 7- 2 DORM",
    "APTO 201", "APTO 208", "APTO 209",
    "ESTUDIO 105", "ESTUDIO 106", "APTO 204- 2 DORM", "APTO 1", "APTO 210",
]

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
[data-testid="stSidebar"] .stMultiSelect [data-baseweb="select"] > div,
[data-testid="stSidebar"] .stTextInput input {
    background: rgba(255,255,255,0.05) !important;
    border: 1px solid rgba(255,255,255,0.10) !important;
    border-radius: 7px !important;
    color: white !important;
    font-size: 0.82rem !important;
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
# CABECERA
# ─────────────────────────────────────────────
st.markdown("# ESTEASUR 2015 - ISLAMAR")
st.markdown("<span style='color:#888;font-size:0.9rem'>Gestión de Reservas · Apartamentos Islamar · 2026</span>", unsafe_allow_html=True)
col_c1, col_c2, col_c3 = st.columns([2, 1, 2])
with col_c2:
    st.image("logo.png", width=120)

st.divider()

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:

    # Logo
    st.image("logo.png", width=140)
    st.markdown("""
    <div style="text-align:center;padding:4px 10px 16px;border-bottom:1px solid rgba(255,255,255,0.07);margin-bottom:4px;">
        <span class="sb-logo-title">ESTEASUR 2015</span>
        <span style="font-size:0.72rem;color:rgba(100,181,246,0.8)!important;letter-spacing:1px;display:block;margin-top:2px;">ISLAMAR</span>
        <span class="sb-logo-sub">Gestión de Reservas</span>
    </div>
    """, unsafe_allow_html=True)

    # Navegación
    st.markdown('<span class="sb-label">Navegación</span>', unsafe_allow_html=True)
    seccion = st.radio("nav", [
        "📊 Reservas",
        "📅 Plantilla mensual",
        "📥 Importar Booking",
        "➕ Nueva reserva",
        "✏️ Editar reserva",
    ], label_visibility="collapsed")

    # Filtros
    st.markdown('<span class="sb-label">Filtros</span>', unsafe_allow_html=True)
    filtro_mes    = st.multiselect("Mes", MESES, placeholder="Todos los meses")
    filtro_fuente = st.multiselect("Fuente", FUENTES, placeholder="Todas las fuentes")
    filtro_nombre = st.text_input("Buscar nombre", placeholder="Nombre del cliente...")
    filtro_dorm   = st.multiselect("Dormitorios", DORMS, placeholder="Todos")

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
        COLS_EDIT = ["fuente","nombre","dormitorios","entrada","salida",
                     "noches","personas","precio","pago_cta","fecha_ingreso",
                     "resto_pdte","estado_pago","mes","comentarios"]
        cols_exist = [c for c in COLS_EDIT if c in df_filtrado.columns]

        # Guardamos IDs por separado para actualizar correctamente
        id_map = df_filtrado["id"].reset_index(drop=True)
        df_show = df_filtrado[cols_exist].copy().reset_index(drop=True)

        edited = st.data_editor(
            df_show,
            use_container_width=True,
            height=520,
            column_config={
                "fuente":       st.column_config.SelectboxColumn("Fuente", options=FUENTES, width=130),
                "nombre":       st.column_config.TextColumn("Nombre", width=200),
                "dormitorios":  st.column_config.SelectboxColumn("Dorm.", options=DORMS, width=80),
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

    # ── Construir grid ────────────────────────
    grid = {apto: {d: None for d in dias} for apto in APTOS}
    for _, r in df.iterrows():
        try:
            entrada = datetime.strptime(str(r.get("entrada","")), "%d/%m/%Y").date()
            salida  = datetime.strptime(str(r.get("salida", "")), "%d/%m/%Y").date()
            apto    = str(r.get("apartamento","")).strip()
            if apto not in APTOS:
                continue
            edia = entrada.day if entrada.month == mes_n and entrada.year == anio_sel else 0
            sdia = salida.day  if salida.month  == mes_n and salida.year  == anio_sel else n_dias + 1
            for d in dias:
                curr = date(anio_sel, mes_n, d)
                if entrada <= curr < salida:
                    grid[apto][d] = {
                        "id": r.get("id"), "nombre": str(r.get("nombre","")),
                        "fuente": str(r.get("fuente","")),
                        "entrada": str(r.get("entrada","")), "salida": str(r.get("salida","")),
                        "precio": str(r.get("precio","")), "estado_pago": str(r.get("estado_pago","")),
                        "edia": edia, "sdia": sdia,
                    }
        except:
            pass

    # ── CSS ───────────────────────────────────
    st.markdown("""
    <style>
    .cal-wrap{overflow-x:auto;border-radius:10px;box-shadow:0 3px 12px rgba(0,0,0,0.15);margin-bottom:8px;}
    .cal-tbl{border-collapse:collapse;font-family:'Segoe UI',Arial,sans-serif;width:100%;}
    .th-apto{background:#1a3f5c;color:white;padding:6px 14px;text-align:left;font-size:0.73rem;
             position:sticky;left:0;z-index:3;white-space:nowrap;min-width:155px;
             border-right:2px solid #0d2a3d;border-bottom:1px solid #0d2a3d;}
    .th-mes{background:linear-gradient(135deg,#1F4E79,#2C5F8A);color:white;text-align:center;
            font-size:1rem;font-weight:700;padding:9px;letter-spacing:2px;}
    .th-day{background:#1F4E79;color:white;padding:3px 1px;text-align:center;
            font-size:0.7rem;min-width:33px;border:1px solid #144070;line-height:1.2;}
    .th-day.we{background:#163d5e;}
    .dow{font-size:0.58rem;color:#90CAF9;display:block;}
    .td-apto{background:#2C5F8A;color:white;font-weight:700;padding:5px 14px;white-space:nowrap;
             font-size:0.78rem;position:sticky;left:0;z-index:1;
             border-right:2px solid #144070;border-bottom:1px solid #1a4a72;}
    .td{padding:0;border:1px solid #dde2ea;height:30px;vertical-align:middle;overflow:hidden;cursor:pointer;}
    .td.we{background:#f0f2f5 !important;}
    .td.libre{background:#fafbfd;}
    .res{display:block;height:100%;line-height:30px;padding:0 4px;font-size:0.67rem;
         font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
    .sep td{background:#D0E8F7;color:#1F4E79;font-weight:700;padding:4px 10px;
            font-size:0.78rem;border-top:2px solid #1F4E79;letter-spacing:.5px;}
    </style>
    """, unsafe_allow_html=True)

    # ── HTML calendario ───────────────────────
    html = '<div class="cal-wrap"><table class="cal-tbl">'

    # Fila título
    html += f'<tr><th class="th-apto" style="font-size:0.88rem;font-weight:700;">{mes_sel} {anio_sel}</th>'
    for d in dias:
        wd  = (primer_dia + d - 1) % 7
        we  = " we" if wd >= 5 else ""
        html += f'<th class="th-day{we}">{d}<span class="dow">{DIAS_SEM[wd]}</span></th>'
    html += '</tr>'

    # Filas apartamentos
    for i, apto in enumerate(APTOS):
        if apto == "ESTUDIO 105":
            html += f'<tr class="sep"><td colspan="{n_dias+1}">▸ JUANMA</td></tr>'

        rbg = "#f5f8fc" if i % 2 == 0 else "#ffffff"
        html += f'<tr><td class="td-apto">{apto}</td>'

        for d in dias:
            c   = grid[apto][d]
            wd  = (primer_dia + d - 1) % 7
            wec = " we" if wd >= 5 else ""

            if c:
                f   = c["fuente"]
                bg  = "#1565C0" if f == "DIRECTA" else ("#2E7D32" if f == "BOOKING.COM" else "#6A1B9A")
                nom = c["nombre"]
                tip = f"{nom} | {c['entrada']} → {c['salida']}"
                pre = "▶ " if d == c["edia"] else ""
                suf = " ◀" if d == c["sdia"] - 1 and 0 < c["sdia"] <= n_dias + 1 else ""
                lbl = f"{pre}{nom}{suf}"
                html += (f'<td class="td{wec}" style="background:{bg};" title="{tip}">'
                         f'<span class="res" style="color:white;">{lbl}</span></td>')
            else:
                fbg = "#eaecef" if wd >= 5 else rbg
                html += f'<td class="td libre{wec}" style="background:{fbg};"></td>'

        html += '</tr>'

    html += '</table></div>'

    # ── Leyenda ───────────────────────────────
    st.markdown("""
    <div style="display:flex;gap:14px;align-items:center;font-size:0.78rem;margin-bottom:6px;flex-wrap:wrap;">
      <span style="color:#888;">▶ Entrada &nbsp;|&nbsp; ◀ Salida &nbsp;|&nbsp; Gris = fin de semana</span>
      <span style="background:#1565C0;color:white;padding:2px 12px;border-radius:12px;">■ Directa</span>
      <span style="background:#2E7D32;color:white;padding:2px 12px;border-radius:12px;">■ Booking.com</span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(html, unsafe_allow_html=True)

    # ── Panel interactivo ─────────────────────
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
        d_sel  = fecha_pi.day
        celda  = grid.get(apto_pi, {}).get(d_sel)

        if celda:
            badge = "🔵 Directa" if celda["fuente"] == "DIRECTA" else "🟢 Booking.com"
            st.success(f"**{apto_pi}** — {fecha_pi.strftime('%d/%m/%Y')}: **{celda['nombre']}** &nbsp; {badge}")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Entrada", celda["entrada"])
            m2.metric("Salida",  celda["salida"])
            m3.metric("Precio",  f"{celda['precio']} €" if celda["precio"] else "—")
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

# ─────────────────────────────────────────────
# SECCIÓN: IMPORTAR BOOKING
# ─────────────────────────────────────────────
elif seccion == "📥 Importar Booking":
    st.markdown("### 📥 Importar reservas desde Booking.com")
    st.markdown("Sube el Excel de **Check-in** que descarga Booking.com y se importarán automáticamente las reservas nuevas.")

    archivo = st.file_uploader("Selecciona el archivo Excel de Booking.com", type=["xls","xlsx"], key="bk_upload")

    if archivo:
        try:
            bk = pd.read_excel(archivo, header=0)
            bk.columns = [str(c).strip() for c in bk.columns]

            # Mapeo de columnas Booking → nuestra BD
            COL_MAP = {
                "nro_reserva":  ["Número de reserva", "Numero de reserva"],
                "nombre":       ["Nombre del cliente (o clientes)", "Nombre del cliente"],
                "entrada":      ["Entrada"],
                "salida":       ["Salida"],
                "noches":       ["Duración (noches)", "Duracion (noches)"],
                "personas":     ["Personas", "Adultos"],
                "precio":       ["Precio"],
                "estado_pago":  ["Estado del pago"],
                "comentarios":  ["Comentarios"],
                "tipo_unidad":  ["Tipo de unidad"],
            }

            def get_col(df_bk, opciones):
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
                try:
                    if isinstance(v, str):
                        d = datetime.strptime(v.strip()[:10], "%Y-%m-%d")
                    else:
                        d = pd.Timestamp(v).to_pydatetime()
                    return d.strftime("%d/%m/%Y")
                except:
                    return str(v)

            # Construir dataframe normalizado
            filas = []
            for _, row in bk.iterrows():
                def g(key):
                    c = get_col(bk, COL_MAP.get(key,[]))
                    return row[c] if c and not pd.isna(row.get(c, float("nan"))) else ""

                nro   = str(g("nro_reserva")).strip()
                if not nro or nro in ("nan",""):
                    continue

                entrada_raw = g("entrada")
                salida_raw  = g("salida")
                entrada_str = fmt_fecha(entrada_raw)
                salida_str  = fmt_fecha(salida_raw)

                try:
                    e_date = datetime.strptime(entrada_str, "%d/%m/%Y")
                    mes_n2 = e_date.month
                    mes_str = MESES[mes_n2 - 1]
                except:
                    mes_n2, mes_str = 0, ""

                precio_raw = limpiar_precio(g("precio"))
                noches_raw = g("noches")
                try:
                    noches_val = int(float(str(noches_raw))) if noches_raw != "" else 0
                except:
                    noches_val = calcular_noches(entrada_str, salida_str)

                estado_raw = str(g("estado_pago")).strip()
                if "booking" in estado_raw.lower():
                    estado_val = "Pago mediante Booking.com"
                elif estado_raw.lower() in ("ok","pagado"):
                    estado_val = "PAGADO"
                else:
                    estado_val = estado_raw

                filas.append({
                    "nro_reserva":  nro,
                    "fuente":       "BOOKING.COM",
                    "nombre":       str(g("nombre")).strip().title(),
                    "mes":          mes_str,
                    "mes_num":      mes_n2,
                    "entrada":      entrada_str,
                    "salida":       salida_str,
                    "noches":       noches_val,
                    "personas":     str(g("personas")).replace(".0",""),
                    "precio":       precio_raw,
                    "estado_pago":  estado_val,
                    "comentarios":  str(g("comentarios")) if g("comentarios") else "",
                    "apartamento":  "",
                })

            df_bk = pd.DataFrame(filas)

            # Detectar duplicados (nro_reserva ya en BD)
            nros_bd = set(str(r) for r in df["nro_reserva"].tolist()) if not df.empty else set()
            df_bk["_nuevo"] = ~df_bk["nro_reserva"].astype(str).isin(nros_bd)
            nuevas   = df_bk[df_bk["_nuevo"]].drop(columns=["_nuevo"])
            ya_exist = df_bk[~df_bk["_nuevo"]].drop(columns=["_nuevo"])

            # Resumen
            col_r1, col_r2, col_r3 = st.columns(3)
            col_r1.metric("Reservas en el archivo", len(df_bk))
            col_r2.metric("✅ Nuevas a importar",   len(nuevas),   delta=f"+{len(nuevas)}")
            col_r3.metric("⚠️ Ya existentes",       len(ya_exist))

            # Vista previa
            if not nuevas.empty:
                st.markdown("#### Vista previa de reservas nuevas")
                cols_vista = ["nro_reserva","nombre","entrada","salida","noches","personas","precio","estado_pago"]
                st.dataframe(
                    nuevas[[c for c in cols_vista if c in nuevas.columns]],
                    use_container_width=True, height=300, hide_index=True,
                    column_config={
                        "nro_reserva":  st.column_config.TextColumn("Nº Reserva", width=130),
                        "nombre":       st.column_config.TextColumn("Nombre", width=200),
                        "entrada":      st.column_config.TextColumn("Entrada", width=100),
                        "salida":       st.column_config.TextColumn("Salida", width=100),
                        "noches":       st.column_config.NumberColumn("Noches", width=70),
                        "personas":     st.column_config.TextColumn("Pers.", width=60),
                        "precio":       st.column_config.TextColumn("Precio €", width=90),
                        "estado_pago":  st.column_config.TextColumn("Estado pago", width=200),
                    },
                )

                st.markdown("")
                if st.button(f"📥 Importar {len(nuevas)} reserva(s) nueva(s)", type="primary", use_container_width=True):
                    importadas = 0
                    errores_imp = []
                    for _, row in nuevas.iterrows():
                        try:
                            guardar_reserva(row.to_dict())
                            importadas += 1
                        except Exception as ex:
                            errores_imp.append(str(ex))

                    if importadas:
                        st.success(f"✅ {importadas} reserva(s) importadas correctamente.")
                        st.rerun()
                    for err in errores_imp:
                        st.error(f"Error: {err}")
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
