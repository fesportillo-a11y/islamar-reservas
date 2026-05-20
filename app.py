import streamlit as st
import pandas as pd
from supabase import create_client, Client
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, date
import re

# ─────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Apartamentos Islamar – Reservas",
    page_icon="🏖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

MESES = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
         "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]

FUENTES  = ["DIRECTA", "BOOKING.COM"]
ESTADOS  = ["", "PAGADO", "PENDIENTE", "SEÑAL PAGADA", "Pago mediante Booking.com", "EFECTIVO", "RESERVA ANULADA"]
DORMS    = ["1", "2", "3", "Estudio"]

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
[data-testid="stAppViewContainer"] { background: #f5f7fa; }
[data-testid="stSidebar"] { background: #1F4E79; }
[data-testid="stSidebar"] * { color: white !important; }
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label { color: #BDD7EE !important; }
.metric-card {
    background: white; border-radius: 10px; padding: 16px 20px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.08); text-align: center;
}
.metric-num  { font-size: 2rem; font-weight: 700; color: #1F4E79; }
.metric-lab  { font-size: 0.85rem; color: #666; margin-top: 2px; }
.badge-directa { background:#D6E4F0; color:#1F4E79; padding:2px 8px; border-radius:12px; font-size:0.8rem; font-weight:600; }
.badge-booking { background:#E8F5E9; color:#2E7D32; padding:2px 8px; border-radius:12px; font-size:0.8rem; font-weight:600; }
.stDataFrame { border-radius: 8px; overflow: hidden; }
h1 { color: #1F4E79 !important; }
h2, h3 { color: #2C5F8A !important; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# CABECERA
# ─────────────────────────────────────────────
col_logo, col_titulo = st.columns([1, 6])
with col_logo:
    st.markdown("## 🏖️")
with col_titulo:
    st.markdown("# Apartamentos Islamar")
    st.markdown("<span style='color:#666;font-size:0.9rem'>Gestión de Reservas 2026</span>", unsafe_allow_html=True)

st.divider()

# ─────────────────────────────────────────────
# SIDEBAR – FILTROS + NAVEGACIÓN
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🔍 Filtros")
    filtro_mes    = st.multiselect("Mes", MESES, placeholder="Todos los meses")
    filtro_fuente = st.multiselect("Fuente", FUENTES, placeholder="Todas las fuentes")
    filtro_nombre = st.text_input("Buscar nombre")
    filtro_dorm   = st.multiselect("Dormitorios", DORMS, placeholder="Todos")

    st.divider()
    st.markdown("### 📋 Sección")
    seccion = st.radio("", ["📊 Reservas", "➕ Nueva reserva", "✏️ Editar reserva"], label_visibility="collapsed")

# ─────────────────────────────────────────────
# CARGAR DATOS
# ─────────────────────────────────────────────
df = cargar_reservas()

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
                file_name=f"Reservas_Islamar_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
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
                "mes_num": mes_num(mes), "nombre": nombre, "dormitorios": dormitorios,
                "entrada": entrada_str, "salida": salida_str, "noches": noches,
                "personas": personas, "precio": precio, "pago_cta": pago_cta,
                "fecha_ingreso": fecha_ing, "resto_pdte": resto_pdte,
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
