"""
Dashboard de Inventarios - Huevos Kikes
=======================================
Aplicación de dos módulos:

  1) Inventario de Edades  -> vista actual de edades por planta/CEDI.
                              Fuente: 'Inventario Hoy.xlsx', hoja 'INV. EDADES'.

  2) Análisis de Rotación PEPS -> auditoría de rotación comparando el inventario
                              inicial de ayer, las ventas del día y el inventario
                              inicial de hoy. Detecta rupturas de rotación PEPS.
                              Fuentes (raíz del repositorio de GitHub):
                                - 'Inventario Hoy.xlsx'   (corte de anoche  = inicial de hoy)
                                - 'Inventario Ayer.xlsx'  (corte de antenoche = inicial de ayer)
                                - 'ventas.xlsx'           (ventas del día de ayer)

Sube los tres archivos al repositorio y la app se actualizará al refrescar.
"""

import re
import unicodedata
from collections import defaultdict

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ===========================================================================
# CONFIGURACIÓN GENERAL Y MARCA
# ===========================================================================
st.set_page_config(
    page_title="Inventarios | Huevos Kikes",
    page_icon="🥚",
    layout="wide",
)

COLOR_PRIMARIO = "#3DAE2B"   # verde Kikes
COLOR_ACENTO = "#F7941D"     # naranja (yema del logo)
COLOR_TEXTO = "#1A1A1A"
COLOR_CRITICO = "#D0021B"
COLOR_ADV = "#F5A623"

# Archivos esperados en la raíz del repositorio.
# Se aceptan variantes con espacio o guion bajo (p.ej. "Inventario Hoy.xlsx"
# o "Inventario_Hoy.xlsx") para no depender de cómo se suban al repo.
import os


def resolver_archivo(*nombres):
    """Devuelve la primera ruta existente entre las variantes dadas.
    Si ninguna existe, devuelve la primera (para que el mensaje de error la nombre)."""
    for n in nombres:
        if os.path.exists(n):
            return n
    return nombres[0]


ARCHIVO_HOY = resolver_archivo("Inventario Hoy.xlsx", "Inventario_Hoy.xlsx")
ARCHIVO_AYER = resolver_archivo("Inventario Ayer.xlsx", "Inventario_Ayer.xlsx")
ARCHIVO_VENTAS = resolver_archivo("ventas.xlsx", "Ventas.xlsx", "VENTAS.xlsx")
HOJA_EDADES_DASH = "INV. EDADES"   # hoja para el módulo 1 (dashboard de edades)
HOJA_INV_ANALISIS = "INV. EDADES"  # hoja para el módulo 2 (análisis de rotación)


def mtime(ruta: str) -> float:
    """Fecha de última modificación del archivo. Se usa como parte de la clave de
    caché: si el archivo cambia (aunque conserve el nombre), el caché se invalida
    y los datos se releen. Resuelve el problema de ver datos viejos tras reemplazar
    un archivo en el repositorio."""
    try:
        return os.path.getmtime(ruta)
    except OSError:
        return 0.0

# Estilos compartidos
st.markdown(
    f"""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;700;800;900&display=swap');
        html, body, [class*="css"] {{
            font-family: 'Nunito', sans-serif;
        }}
        .titulo-principal {{
            color: {COLOR_PRIMARIO};
            font-family: 'Nunito', sans-serif;
            font-size: 4.2rem;
            font-weight: 900;
            line-height: 1.05;
            letter-spacing: -1.5px;
            margin-bottom: 0.2rem;
            display: inline-block;
            border-bottom: 7px solid {COLOR_ACENTO};
            padding-bottom: 0.15rem;
        }}
        .titulo-modulo {{
            color: {COLOR_PRIMARIO};
            font-family: 'Nunito', sans-serif;
            font-size: 3.2rem;
            font-weight: 900;
            line-height: 1.05;
            letter-spacing: -1.2px;
            margin-bottom: 0.2rem;
            display: inline-block;
            border-bottom: 6px solid {COLOR_ACENTO};
            padding-bottom: 0.15rem;
        }}

        /* ----- Tarjetas KPI personalizadas con acento condicional ----- */
        .kpi-card {{
            background-color: #FFFFFF;
            border: 1px solid #E6E6E6;
            border-left: 7px solid #9AA0A6;
            border-radius: 10px;
            padding: 16px 18px;
            height: 100%;
            box-shadow: 0 1px 3px rgba(0,0,0,0.06);
        }}
        .kpi-card .kpi-label {{
            color: #4A4A4A;
            font-size: 0.95rem;
            font-weight: 700;
            margin: 0 0 6px 0;
            line-height: 1.2;
        }}
        .kpi-card .kpi-value {{
            color: {COLOR_TEXTO};
            font-size: 1.9rem;
            font-weight: 900;
            margin: 0;
            line-height: 1.1;
        }}
        .kpi-neutral {{ border-left-color: {COLOR_PRIMARIO}; }}
        .kpi-warning {{ border-left-color: {COLOR_ADV}; }}
        .kpi-critical {{ border-left-color: {COLOR_CRITICO}; }}
        .kpi-reina {{
            background: linear-gradient(135deg, #F2FAF0 0%, #FFFFFF 100%);
            border-left: 9px solid {COLOR_PRIMARIO};
            border-top: 1px solid {COLOR_PRIMARIO}33;
        }}
        .kpi-reina .kpi-label {{ color: {COLOR_PRIMARIO}; font-size: 1rem; }}
        .kpi-reina .kpi-value {{ font-size: 2.2rem; }}
    </style>
    """,
    unsafe_allow_html=True,
)


def tarjeta_kpi(label, value, estado="neutral", reina=False):
    """Genera el HTML de una tarjeta KPI con acento condicional."""
    clases = "kpi-card"
    if reina:
        clases += " kpi-reina"
    else:
        clases += f" kpi-{estado}"
    return (
        f'<div class="{clases}">'
        f'<p class="kpi-label">{label}</p>'
        f'<p class="kpi-value">{value}</p></div>'
    )


# ===========================================================================
# UTILIDADES COMPARTIDAS
# ===========================================================================
def norm(s):
    """Normaliza texto: mayúsculas, sin acentos, sin espacios dobles."""
    if s is None:
        return ""
    s = str(s).strip().upper()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s)


def norm_item(serie: pd.Series) -> pd.Series:
    """Normaliza códigos de item a string entero limpio.

    El inventario trae el item como float (p.ej. 112357.0) y ventas como
    string (p.ej. '2110302'). Unificamos ambos a '112357' / '2110302' para que
    el cruce por código funcione correctamente.
    """
    num = pd.to_numeric(serie, errors="coerce")
    out = num.astype("Int64").astype("string")            # 112357.0 -> '112357'
    no_num = num.isna() & serie.notna()                   # valores no numéricos (raros)
    out[no_num] = serie[no_num].astype("string").str.strip()
    return out


# ===========================================================================
# BASE DE DATOS DE GESTIÓN DE RUPTURAS  (backend: Google Sheets)
# ===========================================================================
# La persistencia vive FUERA del repositorio porque el sistema de archivos de
# Streamlit Cloud es efímero. Usamos un Google Sheet como base de datos:
#   - hoja 'rupturas'  : registro automático e idempotente de las rupturas por fecha
#   - hoja 'gestiones' : explicaciones que el líder de zona consigna desde la app
#
# El usuario NUNCA abre Google Sheets: interactúa solo con el dashboard. El Sheet
# es el almacén invisible (y tu ventana de administración como dueña del proceso).

import datetime as _dt

# Categorías de causa disponibles para el líder (definidas con Camila).
CATEGORIAS_RUPTURA = [
    "Producto en vehículo (sugerido/stock a bordo)",
    "Error de conteo / registro",
    "Reingreso o devolución",
    "Otra",
]

HOJA_RUPTURAS = "rupturas"
HOJA_GESTIONES = "gestiones"

COLS_RUPTURAS = ["llave", "fecha_corte", "fecha_registro", "destino", "item",
                 "referencia", "unds_varadas"]
COLS_GESTIONES = ["llave", "fecha_corte", "destino", "item", "lider_zona",
                  "categoria", "razon", "accion_correctiva", "fecha_gestion"]


def llave_ruptura(fecha_corte, destino, item):
    """Identidad estable de una ruptura: fecha de corte + destino + item."""
    return f"{fecha_corte}|{destino}|{item}"


@st.cache_resource(show_spinner=False)
def _conectar_sheet():
    """Abre el Google Sheet de la BD. Devuelve el objeto Spreadsheet o None si no
    hay credenciales configuradas (modo sin BD). Usa st.secrets para las
    credenciales de la cuenta de servicio (nunca van al repositorio)."""
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except Exception:
        return None
    if "gcp_service_account" not in st.secrets or "sheet_bd" not in st.secrets:
        return None
    try:
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]), scopes=scopes)
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(st.secrets["sheet_bd"]["spreadsheet_id"])
        return sh
    except Exception as e:
        st.session_state["_bd_error"] = str(e)
        return None


def _hoja(sh, nombre, cols):
    """Devuelve la worksheet, creándola con encabezados si no existe."""
    try:
        ws = sh.worksheet(nombre)
    except Exception:
        ws = sh.add_worksheet(title=nombre, rows=1000, cols=max(10, len(cols)))
        ws.append_row(cols)
    # Asegura encabezados si la hoja está vacía
    if not ws.row_values(1):
        ws.append_row(cols)
    return ws


def bd_disponible():
    return _conectar_sheet() is not None


def leer_tabla(nombre, cols):
    """Lee una hoja completa como DataFrame. Devuelve df vacío si no hay BD."""
    sh = _conectar_sheet()
    if sh is None:
        return pd.DataFrame(columns=cols)
    try:
        ws = _hoja(sh, nombre, cols)
        registros = ws.get_all_records()
        df = pd.DataFrame(registros)
        if df.empty:
            return pd.DataFrame(columns=cols)
        for c in cols:
            if c not in df.columns:
                df[c] = None
        return df[cols]
    except Exception as e:
        st.session_state["_bd_error"] = str(e)
        return pd.DataFrame(columns=cols)


def registrar_rupturas(rupturas_df, fecha_corte):
    """Registro AUTOMÁTICO e IDEMPOTENTE de las rupturas de una fecha de corte.
    Si ya existe registro para esa fecha, no hace nada (respeta lo congelado).
    Devuelve (n_registradas, ya_existia)."""
    sh = _conectar_sheet()
    if sh is None:
        return 0, False
    fc = str(fecha_corte)
    try:
        ws = _hoja(sh, HOJA_RUPTURAS, COLS_RUPTURAS)
        existentes = pd.DataFrame(ws.get_all_records())
        # ¿Ya se congeló esta fecha de corte?
        if not existentes.empty and "fecha_corte" in existentes.columns:
            if (existentes["fecha_corte"].astype(str) == fc).any():
                return 0, True
        if rupturas_df.empty:
            return 0, False
        hoy = _dt.date.today().isoformat()
        filas = []
        for r in rupturas_df.itertuples():
            filas.append([
                llave_ruptura(fc, r.destino, r.item),
                fc, hoy, r.destino, str(r.item), r.referencia,
                int(round(r.unds_varadas)),
            ])
        ws.append_rows(filas, value_input_option="USER_ENTERED")
        return len(filas), False
    except Exception as e:
        st.session_state["_bd_error"] = str(e)
        return 0, False


def guardar_gestion(llave, fecha_corte, destino, item, lider, categoria,
                    razon, accion):
    """Consigna la explicación del líder para una ruptura. Devuelve True/False."""
    sh = _conectar_sheet()
    if sh is None:
        return False
    try:
        ws = _hoja(sh, HOJA_GESTIONES, COLS_GESTIONES)
        ws.append_row([
            llave, str(fecha_corte), destino, str(item), lider, categoria,
            razon, accion, _dt.datetime.now().isoformat(timespec="seconds"),
        ], value_input_option="USER_ENTERED")
        return True
    except Exception as e:
        st.session_state["_bd_error"] = str(e)
        return False


# ===========================================================================
# MÓDULO 1 — DASHBOARD DE INVENTARIO DE EDADES  (sin cambios funcionales)
# ===========================================================================
@st.cache_data(ttl=3600)
def cargar_datos_edades(ruta: str, hoja: str, cache_key: float = 0.0) -> pd.DataFrame:
    df = pd.read_excel(ruta, sheet_name=hoja)
    df.columns = [str(c).strip().lower() for c in df.columns]
    renombres = {}
    for col in df.columns:
        if col in ("tipo huevo", "tipo_huevo", "tipohuevo"):
            renombres[col] = "tipo huevo"
    df = df.rename(columns=renombres)
    if "edad" in df.columns:
        df["edad"] = pd.to_numeric(df["edad"], errors="coerce").astype("Int64")
    if "cantidad" in df.columns:
        df["cantidad"] = pd.to_numeric(df["cantidad"], errors="coerce").fillna(0)
    return df


def render_modulo_edades():
    try:
        df = cargar_datos_edades(ARCHIVO_HOY, HOJA_EDADES_DASH, mtime(ARCHIVO_HOY))
    except FileNotFoundError:
        st.error(
            f"No se encontró el archivo **{ARCHIVO_HOY}**. "
            "Asegúrate de que esté en la raíz del repositorio de GitHub."
        )
        st.stop()
    except ValueError:
        st.error(
            f"No se encontró la hoja **{HOJA_EDADES_DASH}** dentro de {ARCHIVO_HOY}. "
            "Verifica el nombre exacto de la hoja."
        )
        st.stop()

    st.markdown('<p class="titulo-principal">🥚 Inventario de Edades</p>', unsafe_allow_html=True)
    st.divider()

    def opciones(col: str):
        if col in df.columns:
            return sorted(df[col].dropna().unique().tolist())
        return []

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        f_zona = st.multiselect("Zona", opciones("zona"), placeholder="Todas")
    with c2:
        f_edad = st.multiselect("Edad", opciones("edad"), placeholder="Todas")
    with c3:
        f_tipo = st.multiselect("Tipo", opciones("tipo"), placeholder="Todos")
    with c4:
        f_destino = st.multiselect("Destino", opciones("destino"), placeholder="Todos")

    dff = df.copy()
    if f_zona:
        dff = dff[dff["zona"].isin(f_zona)]
    if f_edad:
        dff = dff[dff["edad"].isin(f_edad)]
    if f_tipo:
        dff = dff[dff["tipo"].isin(f_tipo)]
    if f_destino:
        dff = dff[dff["destino"].isin(f_destino)]

    st.divider()

    inv_total = dff["cantidad"].sum()
    und_mas_6 = dff.loc[dff["edad"] > 6, "cantidad"].sum()
    und_mas_10 = dff.loc[dff["edad"] >= 10, "cantidad"].sum()
    _val = dff.dropna(subset=["edad"])
    _peso = _val["cantidad"].sum()
    edad_prom = (
        (_val["edad"].astype(float) * _val["cantidad"]).sum() / _peso if _peso > 0 else 0
    )
    pct_mas_6 = (und_mas_6 / inv_total * 100) if inv_total > 0 else 0

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.markdown(
            tarjeta_kpi("Edad promedio ponderada", f"{edad_prom:,.1f} días", reina=True),
            unsafe_allow_html=True,
        )
    with k2:
        st.markdown(
            tarjeta_kpi("Inventario total (und)", f"{inv_total:,.0f}", estado="neutral"),
            unsafe_allow_html=True,
        )
    with k3:
        st.markdown(
            tarjeta_kpi("Unidades con más de 6 días", f"{und_mas_6:,.0f}", estado="warning"),
            unsafe_allow_html=True,
        )
    with k4:
        st.markdown(
            tarjeta_kpi("% con más de 6 días", f"{pct_mas_6:,.1f}%", estado="warning"),
            unsafe_allow_html=True,
        )
    with k5:
        st.markdown(
            tarjeta_kpi("Unidades con 10+ días", f"{und_mas_10:,.0f}", estado="critical"),
            unsafe_allow_html=True,
        )

    st.divider()
    st.subheader("Distribución del inventario por edad")

    dist = dff.dropna(subset=["edad"]).copy()
    dist["edad_int"] = dist["edad"].astype(int)
    dist["bucket"] = dist["edad_int"].apply(lambda x: "10+" if x >= 10 else str(x))
    orden_buckets = [str(i) for i in range(0, 10)] + ["10+"]
    serie = dist.groupby("bucket")["cantidad"].sum().reindex(orden_buckets, fill_value=0)

    def color_barra(bucket):
        if bucket == "10+":
            return COLOR_CRITICO
        val = int(bucket)
        if val <= 5:
            return COLOR_PRIMARIO
        elif val <= 9:
            return COLOR_ADV
        return COLOR_CRITICO

    colores = [color_barra(b) for b in serie.index]
    fig = go.Figure(
        go.Bar(
            x=list(serie.index),
            y=serie.values,
            marker_color=colores,
            text=[f"{v:,.0f}" if v > 0 else "" for v in serie.values],
            textposition="outside",
            hovertemplate="Edad: %{x} días<br>Cantidad: %{y:,.0f} und<extra></extra>",
        )
    )
    fig.update_layout(
        height=360,
        margin=dict(l=10, r=10, t=10, b=10),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Nunito, sans-serif", size=13),
        xaxis=dict(title="Días de edad del producto", tickmode="linear"),
        yaxis=dict(title="Unidades", showgrid=True, gridcolor="#EEEEEE"),
        bargap=0.25,
    )
    st.plotly_chart(fig, use_container_width=True)

    st.divider()
    st.subheader("Detalle de inventario")

    cols_tabla = [c for c in ["destino", "edad", "item", "referencia", "cantidad"] if c in dff.columns]
    tabla = (
        dff[cols_tabla]
        .groupby([c for c in cols_tabla if c != "cantidad"], as_index=False)["cantidad"]
        .sum()
        .sort_values(["destino", "edad"], ascending=[True, False])
        .reset_index(drop=True)
    )
    tabla = tabla.rename(
        columns={
            "destino": "Destino",
            "edad": "Edad",
            "item": "Item",
            "referencia": "Referencia",
            "cantidad": "Suma de Cantidad",
        }
    )

    def color_edad(val):
        try:
            v = float(val)
        except (TypeError, ValueError):
            return ""
        if 1 <= v <= 5:
            return "background-color: #C6EFCE; color: #006100; font-weight: 700;"
        elif v == 6:
            return "background-color: #FFE08A; color: #7A5200; font-weight: 700;"
        elif 7 <= v <= 9:
            return "background-color: #FFB84D; color: #7A3E00; font-weight: 800;"
        elif v >= 10:
            return "background-color: #FF8A80; color: #7A0006; font-weight: 800;"
        return ""

    if "Item" in tabla.columns:
        tabla["Item"] = pd.to_numeric(tabla["Item"], errors="coerce").astype("Int64")

    styler = (
        tabla.style
        .map(color_edad, subset=["Edad"])
        .bar(subset=["Suma de Cantidad"], color="#BFE3B5", align="left", height=70, vmin=0)
        .format({"Suma de Cantidad": "{:,.0f}", "Edad": "{:.0f}", "Item": "{:.0f}"})
    )
    st.dataframe(styler, use_container_width=True, hide_index=True, height=600)

    st.markdown(
        """
        **Convención de edades:**
        🟢 1–5 días: óptimo &nbsp;&nbsp; 🟡 6 días: alerta &nbsp;&nbsp; 🟠 7–9 días: preocupante &nbsp;&nbsp; 🔴 10+ días: crítico
        &nbsp;&nbsp;|&nbsp;&nbsp; Las barras en *Suma de Cantidad* son proporcionales al volumen de cada fila.
        """
    )
    st.caption(f"Fuente: {ARCHIVO_HOY} — hoja {HOJA_EDADES_DASH}")


# ===========================================================================
# MÓDULO 2 — ANÁLISIS DE ROTACIÓN PEPS
# ===========================================================================

# --- Mapeo bodega de venta -> destino de inventario --------------------------
def map_bodega(desc):
    """Devuelve (destino_inv, motivo). destino=None si no mapea."""
    b = norm(desc)
    if b == "BODEGA KIKES":
        return "LANZA", "regla BODEGA KIKES=LANZA"
    if "MONTEVIDEO" in b:
        return "TAT BOGOTA MONTEVIDEO", "bogota montevideo"
    if "SIBERIA" in b:
        return "TAT BOGOTA SIBERIA", "bogota siberia"
    ciudades = {
        "BARRANQUILLA": "TAT BARRANQUILLA", "BUCARAMANGA": "TAT BUCARAMANGA",
        "CARTAGENA": "TAT CARTAGENA", "CUCUTA": "TAT CUCUTA",
        "MEDELLIN": "TAT MEDELLIN", "MONTERIA": "TAT MONTERIA",
        "PASTO": "TAT PASTO", "POPAYAN": "TAT POPAYAN",
        "SANTAMARTA": "TAT SANTA MARTA", "SANTA MARTA": "TAT SANTA MARTA",
        "SINCELEJO": "TAT SINCELEJO", "VALLEDUPAR": "TAT VALLEDUPAR",
        "VILLAVICENCIO": "TAT VILLAVICENCIO", "CALI": "TAT CALI",
    }
    for k, v in ciudades.items():
        if k in b:
            return v, f"ciudad={k}"
    if "BELLAVISTA" in b:
        return "BELLAVISTA", "planta"
    if "PALMAS" in b:
        return "PALMAS", "planta"
    if "EGIPTO" in b:
        return None, "EGIPTO (no en EDADES)"
    if b.startswith("BODEGA BOGOTA") or "BEES BOGOTA" in b:
        return None, "bodega bogota genérica (ambiguo Montevideo/Siberia)"
    if "PEREIRA" in b:
        return "OL. PEREIRA", "pereira"
    return None, "sin mapeo"


@st.cache_data(ttl=3600)
def leer_inventario(ruta: str, hoja: str, cache_key: float = 0.0) -> pd.DataFrame:
    """Lee la hoja INV. EDADES y devuelve columnas estandarizadas."""
    df = pd.read_excel(ruta, sheet_name=hoja)
    df.columns = [str(c).strip().lower() for c in df.columns]
    col = {c: c for c in df.columns}
    # columnas esperadas: destino, edad, item, cantidad, referencia
    out = pd.DataFrame({
        "destino": df.get("destino").map(norm) if "destino" in df else "",
        "edad": pd.to_numeric(df.get("edad"), errors="coerce"),
        "item": norm_item(df.get("item")),
        "referencia": df.get("referencia"),
        "cantidad": pd.to_numeric(df.get("cantidad"), errors="coerce").fillna(0.0),
    })
    out = out.dropna(subset=["edad", "item"])
    out = out[out["destino"] != ""]
    out["edad"] = out["edad"].astype(int)
    return out


@st.cache_data(ttl=3600)
def leer_fecha_corte(ruta: str, hoja: str, cache_key: float = 0.0):
    """Lee la fecha de corte más reciente del inventario (columna 'Fecha').

    Se mantiene separada de leer_inventario porque st.cache_data no preserva
    de forma fiable los atributos (df.attrs) al serializar el DataFrame.
    """
    df = pd.read_excel(ruta, sheet_name=hoja, usecols=None)
    df.columns = [str(c).strip().lower() for c in df.columns]
    if "fecha" not in df.columns:
        return None
    fechas = pd.to_datetime(df["fecha"], errors="coerce").dropna()
    if fechas.empty:
        return None
    return fechas.max().normalize()


def es_pet(descripcion):
    """Categoriza un SKU: PET si la descripción contiene 'PET', si no SUELTO."""
    return "PET" if "PET" in str(descripcion).upper() else "SUELTO"


# Patrón de producto de TIENDA: "HUEVO (TALLA) X (N) CARTON VERDE CANASTA".
# Todo lo demás se asume pedido para preventa. Tolera espaciado variable (X 30 / X20)
# y sufijos después de CANASTA (p.ej. "- BUCAROS").
_PATRON_TIENDA = re.compile(r"HUEVO\s+\w+\s*X\s*\d+\s+CARTON VERDE CANASTA")


def es_tienda(referencia):
    """True si la referencia tiene la estructura de producto de tienda."""
    return bool(_PATRON_TIENDA.search(str(referencia).upper()))


@st.cache_data(ttl=3600)
def leer_ventas(ruta: str, cache_key: float = 0.0) -> pd.DataFrame:
    """Lee ventas, filtra línea HU, mapea bodega->destino, categoriza PET/SUELTO
    y extrae la fecha de venta (columna fec_venta). El archivo puede contener
    varios días."""
    raw = pd.read_excel(ruta, sheet_name=0)
    raw.columns = [str(c).strip().lower() for c in raw.columns]
    df = raw[raw["codigo_linea"] == "HU"].copy()
    df["item"] = norm_item(df["id_item"])
    df["cantidad"] = pd.to_numeric(df["cantidad"], errors="coerce").fillna(0.0)
    mapeo = df["descripcion"].apply(map_bodega)
    df["destino"] = mapeo.apply(lambda t: t[0])
    df["motivo_map"] = mapeo.apply(lambda t: t[1])
    df["bodega_raw"] = df["descripcion"]
    df["categoria"] = df["descripcion_articulo"].apply(es_pet)
    # Fecha de venta (columna fec_venta). Puede haber varios días en el archivo.
    df["fecha_venta"] = pd.to_datetime(df.get("fec_venta"), errors="coerce").dt.date
    return df[["item", "cantidad", "destino", "motivo_map", "bodega_raw",
               "categoria", "fecha_venta"]]


def preparar_ventas_peps(ventas, fecha_corte_hoy, dias):
    """Construye la venta que alimentará el teórico PEPS por (destino, item):
       - SUELTO: promedio diario (total ÷ días operativos del rango) × días de ventana.
                 'Días operativos' = días distintos del rango que NO son domingo, ya que
                 el domingo es día estructural sin despacho y contarlo deprimiría el
                 promedio (falsos positivos los lunes). La venta de un domingo, si la
                 hubiera, sí se suma al total; solo se excluye del divisor.
       - PET: venta exacta del día analizado (= fecha de corte 'hoy').
    Devuelve (dict {(destino,item): venta_peps}, dict {(destino,item): categoria},
              dict {(destino,item): venta_diaria}, n_dias_operativos).
    'venta_diaria' es el ritmo por día (promedio para SUELTO; venta del día para PET),
    usado en la alerta de vida útil."""
    ven_ok = ventas.dropna(subset=["destino"]).copy()
    # Días operativos del rango = días distintos que NO son domingo (weekday()==6).
    fechas_unicas = pd.Series(ven_ok["fecha_venta"].dropna().unique())
    dias_operativos = int(sum(1 for f in fechas_unicas if f.weekday() != 6))
    if dias_operativos == 0:
        dias_operativos = 1
    dias_rango = dias_operativos   # nombre conservado para el resto del código

    # Categoría dominante por (destino, item) — un SKU es PET o SUELTO de forma estable
    cat_map = (ven_ok.groupby(["destino", "item"])["categoria"]
               .agg(lambda s: "PET" if (s == "PET").any() else "SUELTO").to_dict())

    venta_peps = {}
    venta_diaria = {}
    for (dest, item), cat in cat_map.items():
        sub = ven_ok[(ven_ok["destino"] == dest) & (ven_ok["item"] == item)]
        total = sub["cantidad"].sum()
        if cat == "SUELTO":
            prom = total / dias_operativos            # promedio diario (sin domingos)
            venta_diaria[(dest, item)] = prom
            venta_peps[(dest, item)] = prom * dias    # ventana del análisis
        else:  # PET: venta exacta del día = fecha de corte 'hoy'
            vdia = sub[sub["fecha_venta"] == fecha_corte_hoy]["cantidad"].sum()
            venta_diaria[(dest, item)] = vdia
            venta_peps[(dest, item)] = vdia
    return venta_peps, cat_map, venta_diaria, dias_rango


def proyectar_vida_util(inv_vec, venta_diaria, umbral=5):
    """Proyección PEPS lote por lote, separando dos fenómenos distintos:
       - YA VENCIDO HOY: el lote ya tiene edad > umbral en el inventario actual,
         sin importar la venta. Es producto a retirar/vender ya, no un problema futuro.
       - RIESGO PROYECTADO: el lote nace con edad <= umbral hoy, pero la venta lo
         lleva a superar el umbral antes de agotarse (este es el propósito de la alerta).
    Lo viejo sale primero; cada lote espera a que se vendan los más viejos.
    inv_vec: dict edad->cantidad (inventario actual del SKU/destino).
    Devuelve dict con: riesgo_proyectado(bool), ya_vencido(bool),
      edad_max_proyectada(float), unds_riesgo, unds_vencidas, detalle(list).
    """
    base = {"riesgo_proyectado": False, "ya_vencido": False,
            "edad_max_proyectada": 0.0, "unds_riesgo": 0.0,
            "unds_vencidas": 0.0, "detalle": []}
    if not inv_vec:
        return base
    acumulado = 0.0
    detalle = []
    edad_max_proy = 0.0
    unds_riesgo = 0.0
    unds_vencidas = 0.0
    # PEPS: procesar de más viejo a más nuevo
    for edad in sorted(inv_vec.keys(), reverse=True):
        cant = inv_vec[edad]
        if cant <= 0:
            continue
        ya_vencido_lote = edad > umbral          # ya superó la vida útil hoy
        if venta_diaria > 0:
            acumulado += cant
            dias_hasta_agotar = acumulado / venta_diaria
            edad_proyectada = edad + dias_hasta_agotar
        else:
            dias_hasta_agotar = None
            edad_proyectada = float(edad)        # sin venta no se proyecta avance
        edad_max_proy = max(edad_max_proy, edad_proyectada)

        # Riesgo proyectado SOLO si el lote nace sano (<=umbral) pero la venta lo cruza.
        riesgo_lote = (not ya_vencido_lote) and (venta_diaria > 0) and (edad_proyectada > umbral)
        if ya_vencido_lote:
            unds_vencidas += cant
        elif riesgo_lote:
            unds_riesgo += cant

        detalle.append({
            "edad_actual": edad, "cantidad": cant,
            "dias_para_vender": round(dias_hasta_agotar, 1) if dias_hasta_agotar is not None else None,
            "edad_proyectada": round(edad_proyectada, 1),
            "ya_vencido": ya_vencido_lote,
            "riesgo": riesgo_lote,
        })
    return {
        "riesgo_proyectado": unds_riesgo > 0,
        "ya_vencido": unds_vencidas > 0,
        "edad_max_proyectada": round(edad_max_proy, 1),
        "unds_riesgo": round(unds_riesgo, 0),
        "unds_vencidas": round(unds_vencidas, 0),
        "detalle": detalle,
    }


def vec_por_edad(df, dest, item):
    """dict edad->cantidad para un destino/item."""
    sub = df[(df["destino"] == dest) & (df["item"] == item)]
    d = defaultdict(float)
    for r in sub.itertuples():
        d[int(r.edad)] += float(r.cantidad)
    return d


def edad_ponderada(d):
    t = sum(d.values())
    return sum(e * c for e, c in d.items()) / t if t > 0 else 0.0


def peps_consumir(inv_envejecido, vendido):
    """Consume 'vendido' por PEPS (edad mayor primero). Devuelve (remanente, faltante)."""
    rem = dict(inv_envejecido)
    restante = vendido
    for edad in sorted(rem.keys(), reverse=True):
        if restante <= 0:
            break
        toma = min(rem[edad], restante)
        rem[edad] -= toma
        restante -= toma
    rem = {e: c for e, c in rem.items() if c > 0.5}
    return rem, restante


def construir_analisis(inv_ayer, inv_hoy, ventas, dias=1, venta_peps=None, cat_map=None):
    """Motor PEPS consciente de los días transcurridos.

    'dias' es el número de días entre el corte de inv_ayer y el de inv_hoy
    (1 en un día normal; 3 tras un puente). Cada lote del inventario inicial
    se envejece +dias antes de consumir las ventas por PEPS.

    venta_peps: dict {(destino,item): venta} a usar en el teórico PEPS. Para SUELTO
    es el promedio diario × días de ventana; para PET la venta exacta del día. Si es
    None, se usa la venta total del archivo (comportamiento anterior).
    cat_map: dict {(destino,item): 'PET'|'SUELTO'} para etiquetar cada resultado.
    """
    dias = max(1, int(dias))
    ven_ok = ventas.dropna(subset=["destino"])
    # Venta que alimenta el teórico PEPS: si se pasa venta_peps (promedio para SUELTO,
    # venta exacta del día para PET) se usa esa; si no, cae a la venta total (compat.).
    if venta_peps is None:
        ven_map = ven_ok.groupby(["destino", "item"])["cantidad"].sum().to_dict()
    else:
        ven_map = venta_peps
    if cat_map is None:
        cat_map = {}

    ref_map = {}
    for df in (inv_ayer, inv_hoy):
        for r in df.itertuples():
            if r.item not in ref_map and pd.notna(r.referencia) and str(r.referencia) != "#N/D":
                ref_map[r.item] = r.referencia

    claves = inv_ayer.groupby(["destino", "item"]).size().index.tolist()
    filas = []
    for dest, item in claves:
        va = vec_por_edad(inv_ayer, dest, item)
        vh = vec_por_edad(inv_hoy, dest, item)
        vendido = float(ven_map.get((dest, item), 0.0))
        tot_ayer = sum(va.values())
        tot_hoy = sum(vh.values())

        va_env = defaultdict(float)
        for e, c in va.items():
            va_env[e + dias] += c
        teorico, faltante = peps_consumir(va_env, vendido)

        edad_max_teorica = max(teorico.keys()) if teorico else (max(va_env.keys()) if va_env else 0)
        edad_mas_vieja_ayer = max(va_env.keys()) if va_env else 0

        # --- Detección a nivel de cohorte: separa VARADO de INFLADO ---
        # Recorremos solo las edades que YA existían ayer (cohortes reales).
        #   - INFLADO: la cohorte creció respecto a ayer (real > ayer +10%): imposible
        #     por envejecimiento, entró producto o hubo error de registro. No es rotación.
        #   - VARADO (ruptura de ORDEN PEPS): un lote VIEJO quedó con exceso sobre su
        #     teórico Y, al mismo tiempo, salió (rotó) producto de un lote MÁS NUEVO.
        #     Esa es la verdadera inversión de orden. Si lo más viejo se agotó, o si solo
        #     hay un lote, NO es ruptura (el excedente suele ser producto a bordo de los
        #     vehículos: sugeridos + stock del vehículo, que el modelo no rastrea por edad).
        UMBRAL_INFLADO = 1.10
        varado = 0.0
        inflado = 0.0
        detalle_varado = []

        # Cuánto salió (rotó) de cada cohorte = max(0, ayer_envejecido - real).
        # Una cohorte "más nueva" que rotó es la que tiene MENOR edad y salida > 0.
        salida_por_edad = {}
        for e_hoy_coh in va_env.keys():
            salida_por_edad[e_hoy_coh] = max(0.0, va_env.get(e_hoy_coh, 0.0) - vh.get(e_hoy_coh, 0.0))

        if vendido < tot_ayer - 0.5:
            for e_ayer, c_ayer_coh in va.items():
                e_hoy_coh = e_ayer + dias
                c_real = vh.get(e_hoy_coh, 0.0)
                c_teo = teorico.get(e_hoy_coh, 0.0)
                if c_real > c_ayer_coh * UMBRAL_INFLADO + 0.5:
                    # La cohorte creció: imposible por envejecimiento -> inflado.
                    inflado += c_real - c_ayer_coh
                    continue
                exceso = c_real - c_teo
                if exceso <= 0.5:
                    continue
                # ¿Salió producto de algún lote MÁS NUEVO que este (edad menor)?
                rotó_un_lote_mas_nuevo = any(
                    e_otro < e_hoy_coh and sal > 0.5
                    for e_otro, sal in salida_por_edad.items()
                )
                if rotó_un_lote_mas_nuevo:
                    varado += exceso            # ruptura de orden real
                    detalle_varado.append((e_hoy_coh, exceso))
                # Si NO rotó nada más nuevo, el exceso es producto a bordo / no rotación:
                # no se cuenta como varado.

        ep_teo = edad_ponderada(teorico)
        ep_real = edad_ponderada(vh)
        ep_ayer = edad_ponderada(va)
        ruptura = (varado > 0.5) and (vendido > 0)
        hay_inflado = inflado > 0.5

        # Diagnóstico textual. Con una ventana de 'dias', un lote puede envejecer
        # legítimamente hasta +dias; solo un salto MAYOR a eso es anómalo.
        edades_real = list(vh.keys())
        edad_max_real = max(edades_real) if edades_real else 0
        salto = edad_max_real - (edad_mas_vieja_ayer - dias)  # edad original más vieja de ayer
        if not ruptura:
            diag = ""
        elif edad_max_real > edad_mas_vieja_ayer:
            diag = (f"Apareció lote de {edad_max_real}d, más viejo de lo posible incluso tras "
                    f"{dias} día(s) de envejecimiento: posible reingreso/devolución o conteo inconsistente")
        elif teorico and edad_max_real > max(teorico.keys()):
            diag = "Lote viejo no rotó: salió producto más nuevo dejando varado el antiguo"
        else:
            diag = "Producto antiguo permanece pese a haber ventas del día"

        # Conjuntos de edades (hoy) clasificadas, para que las tablas usen exactamente
        # el mismo criterio que el motor (incluida la condición de inversión de orden).
        edades_varadas = {e for e, _ in detalle_varado}
        edades_infladas = set()
        if vendido < tot_ayer - 0.5:
            for e_ayer, c_ayer_coh in va.items():
                e_hoy_coh = e_ayer + dias
                if vh.get(e_hoy_coh, 0.0) > c_ayer_coh * UMBRAL_INFLADO + 0.5:
                    edades_infladas.add(e_hoy_coh)

        filas.append({
            "destino": dest, "item": item, "referencia": ref_map.get(item, ""),
            "categoria": cat_map.get((dest, item), "SUELTO"),
            "cant_ayer": tot_ayer, "edad_ayer": round(ep_ayer, 1),
            "vendido": vendido,
            "cant_hoy": tot_hoy, "edad_hoy": round(ep_real, 1),
            "edad_pond_teorica": round(ep_teo, 1),
            "delta_edad": round(ep_real - ep_teo, 1),
            "edad_max_teorica": edad_max_teorica,
            "unds_varadas": round(varado, 0),
            "unds_infladas": round(inflado, 0),
            "hay_inflado": hay_inflado,
            "edades_varadas": edades_varadas,
            "edades_infladas": edades_infladas,
            "faltante_peps": round(faltante, 0),
            "ruptura": ruptura, "diagnostico": diag,
            "vec_ayer": dict(va), "vec_teorico": dict(teorico), "vec_real": dict(vh),
            "edad_mas_vieja_ayer": edad_mas_vieja_ayer,
        })

    res = pd.DataFrame(filas)
    no_map = ventas[ventas["destino"].isna()].groupby(
        ["bodega_raw", "motivo_map"])["cantidad"].sum().reset_index()
    no_map = no_map.sort_values("cantidad", ascending=False)
    return res, no_map


def render_modulo_rotacion():
    st.markdown('<p class="titulo-modulo">🔄 Análisis de Rotación PEPS</p>', unsafe_allow_html=True)
    st.caption(
        "Compara el inventario inicial de ayer + ventas del período contra el inventario inicial de hoy. "
        "Lógica: se envejece cada lote según los días transcurridos entre cortes, se consume la venta por "
        "PEPS (más viejo primero) y se contrasta contra el inventario real de hoy."
    )
    st.divider()

    # Carga
    faltantes = []
    try:
        inv_ayer = leer_inventario(ARCHIVO_AYER, HOJA_INV_ANALISIS, mtime(ARCHIVO_AYER))
    except FileNotFoundError:
        faltantes.append(ARCHIVO_AYER)
    try:
        inv_hoy = leer_inventario(ARCHIVO_HOY, HOJA_INV_ANALISIS, mtime(ARCHIVO_HOY))
    except FileNotFoundError:
        faltantes.append(ARCHIVO_HOY)
    try:
        ventas = leer_ventas(ARCHIVO_VENTAS, mtime(ARCHIVO_VENTAS))
    except FileNotFoundError:
        faltantes.append(ARCHIVO_VENTAS)

    if faltantes:
        st.error(
            "No se encontraron estos archivos en la raíz del repositorio: "
            + ", ".join(f"**{f}**" for f in faltantes)
            + ". Súbelos para activar el análisis."
        )
        st.stop()

    # --- Días transcurridos entre el corte de ayer y el de hoy ---
    f_ayer = leer_fecha_corte(ARCHIVO_AYER, HOJA_INV_ANALISIS, mtime(ARCHIVO_AYER))
    f_hoy = leer_fecha_corte(ARCHIVO_HOY, HOJA_INV_ANALISIS, mtime(ARCHIVO_HOY))
    dias = 1
    fechas_ok = (f_ayer is not None) and (f_hoy is not None)
    if fechas_ok:
        dias = int((f_hoy - f_ayer).days)

    # Validaciones de la ventana antes de correr el motor
    if fechas_ok and dias < 0:
        st.error(
            f"Las fechas de corte están invertidas: **Ayer = {f_ayer.date()}**, "
            f"**Hoy = {f_hoy.date()}**. El inventario de hoy debe ser posterior al de ayer. "
            "Parece que intercambiaste los archivos."
        )
        st.stop()

    if fechas_ok and dias == 0:
        st.error(
            f"Ambos inventarios tienen la **misma fecha de corte ({f_hoy.date()})**, "
            "así que no hay un período que analizar. Esto suele pasar por una de dos razones:\n\n"
            "1. Subiste el mismo archivo (o una copia) en *Ayer* y *Hoy*.\n"
            "2. Reemplazaste un archivo en el repositorio pero la app está mostrando datos "
            "en caché. Usa el botón **🔄 Actualizar datos** del panel lateral y vuelve a intentar."
        )
        st.stop()

    if not fechas_ok:
        st.warning(
            "No pude leer la fecha de corte de uno de los inventarios (columna *Fecha*). "
            "Asumiré **1 día** de diferencia; si vienes de un puente, los resultados no serán confiables."
        )

    # Fecha de corte 'hoy' (la del inventario de hoy). Para PET, su venta exacta
    # se toma de este día dentro del archivo de ventas.
    fecha_corte_obj = f_hoy.date() if fechas_ok else _dt.date.today()
    fecha_corte_hoy = fecha_corte_obj.isoformat()

    # Venta para el teórico PEPS: SUELTO usa promedio diario; PET la venta del día.
    venta_peps, cat_map, venta_diaria, dias_rango = preparar_ventas_peps(
        ventas, fecha_corte_obj, dias)

    res, no_map = construir_analisis(inv_ayer, inv_hoy, ventas, dias=dias,
                                     venta_peps=venta_peps, cat_map=cat_map)
    rupturas = res[res["ruptura"]].copy()

    if dias_rango > 1:
        st.info(
            f"📅 El promedio diario de **SUELTO** se calcula sobre **{dias_rango} días "
            "operativos** del rango (se excluyen domingos por ser días sin despacho). "
            f"Para **PET** se usa la venta exacta del día de corte ({fecha_corte_hoy})."
        )

    # --- Registro automático e idempotente a la BD ---
    # La primera vez que se corre el análisis para esta fecha de corte, las rupturas
    # quedan congeladas en la BD. Reaperturas el mismo día no duplican ni recalculan.
    if bd_disponible():
        n_reg, ya_existia = registrar_rupturas(rupturas, fecha_corte_hoy)
        if n_reg > 0:
            st.success(f"✅ {n_reg} ruptura(s) del corte {fecha_corte_hoy} registradas en seguimiento.")
    # Guardamos para que el formulario y el módulo de seguimiento los reutilicen
    st.session_state["_fecha_corte_actual"] = fecha_corte_hoy

    # Aviso de ventana multi-día
    if dias > 1:
        rango = ""
        if fechas_ok:
            rango = f" (corte {f_ayer.date()} → {f_hoy.date()})"
        st.warning(
            f"**Ventana de {dias} días{rango}.** El inventario se envejeció +{dias} días y las ventas "
            "del archivo se consumieron de forma agregada. El análisis indica **si hubo ruptura** en la "
            "ventana completa, no en qué día ocurrió. En ventanas largas, las alertas de *lote reaparecido* "
            "deben leerse con más cautela, ya que el envejecimiento normal acerca las edades al umbral."
        )

    # --- KPIs ---
    k1, k2 = st.columns(2)
    with k1:
        st.markdown(tarjeta_kpi("Rupturas de rotación", f"{len(rupturas):,}",
                                estado="critical" if len(rupturas) else "neutral", reina=True),
                    unsafe_allow_html=True)
    with k2:
        st.markdown(tarjeta_kpi("Unidades viejas varadas", f"{rupturas['unds_varadas'].sum():,.0f}",
                                estado="critical"), unsafe_allow_html=True)

    st.divider()

    # ----- Sub-secciones por pestañas -----
    tab1, tab2, tab3 = st.tabs(
        ["🚨 Rupturas PEPS", "📋 Detalle por destino (ayer vs hoy)",
         "⏳ Riesgo de vida útil"]
    )

    # ===== TAB 1: RUPTURAS PEPS =====
    with tab1:
        st.subheader("Rupturas de rotación detectadas")
        st.caption(
            "Producto viejo que, según PEPS, debió salir y permanece en inventario "
            "(o reapareció más viejo de lo posible)."
        )
        if rupturas.empty:
            st.success("✅ No se detectaron rupturas de rotación con los datos cargados.")
        else:
            cols = ["destino", "item", "referencia", "cant_ayer", "vendido", "cant_hoy",
                    "edad_max_teorica", "edad_pond_teorica", "edad_hoy", "unds_varadas",
                    "diagnostico"]
            t = rupturas.sort_values("unds_varadas", ascending=False)[cols].copy()
            t["item"] = pd.to_numeric(t["item"], errors="coerce").astype("Int64")
            t = t.rename(columns={
                "destino": "Destino", "item": "Item", "referencia": "Referencia",
                "cant_ayer": "Inv. ayer", "vendido": "Vendido", "cant_hoy": "Inv. hoy",
                "edad_max_teorica": "Edad máx. teórica", "edad_pond_teorica": "Edad pond. teórica",
                "edad_hoy": "Edad pond. real", "unds_varadas": "Unds varadas",
                "diagnostico": "Diagnóstico",
            })
            styler = (
                t.style
                .format({"Inv. ayer": "{:,.0f}", "Vendido": "{:,.0f}", "Inv. hoy": "{:,.0f}",
                         "Unds varadas": "{:,.0f}", "Edad pond. teórica": "{:.1f}",
                         "Edad pond. real": "{:.1f}", "Item": "{:.0f}"})
                .map(lambda _: f"color:{COLOR_CRITICO}; font-weight:800;", subset=["Unds varadas"])
            )
            st.dataframe(styler, use_container_width=True, hide_index=True)

            st.markdown("##### Detalle por lote de la ruptura seleccionada")
            opciones_rup = [
                f"{r.destino} — {int(r.item) if str(r.item).isdigit() else r.item} — {r.referencia}"
                for r in rupturas.sort_values("unds_varadas", ascending=False).itertuples()
            ]
            sel = st.selectbox("Selecciona una ruptura para ver el detalle lote a lote", opciones_rup)
            if sel:
                idx = opciones_rup.index(sel)
                r = rupturas.sort_values("unds_varadas", ascending=False).iloc[idx]

                st.caption(
                    f"Cada fila sigue un **lote** desde su edad de ayer hasta hoy (envejece +{dias} día(s)). "
                    f"**Teórico PEPS** = lo que debería quedar de ese lote si rotara bien; "
                    f"**Real** = lo observado hoy a esa edad. La reconstrucción por lote es un modelo PEPS "
                    "(el inventario no etiqueta lotes individuales)."
                )

                # Venta del período como KPI destacado
                kv1, kv2 = st.columns([1, 2])
                with kv1:
                    st.markdown(
                        tarjeta_kpi("Venta del período", f"{r['vendido']:,.0f} unds", reina=True),
                        unsafe_allow_html=True,
                    )

                # --- SECCIÓN 1: Lotes que ya existían ayer (su evolución) ---
                edades_ayer = sorted(r["vec_ayer"].keys(), reverse=True)
                vendio_todo = r["vendido"] >= r["cant_ayer"] - 0.5
                filas_lote = []
                for e in edades_ayer:
                    e_hoy = e + dias                      # edad de ese lote hoy
                    c_ayer = r["vec_ayer"].get(e, 0)
                    c_teo = r["vec_teorico"].get(e_hoy, 0)    # lo que PEPS dejaría
                    c_real = r["vec_real"].get(e_hoy, 0)      # lo observado a esa edad envejecida
                    # Clasificación tomada del motor (mismo criterio, incluida la
                    # condición de inversión de orden para el varado).
                    es_inflado = e_hoy in r["edades_infladas"]
                    es_varado = e_hoy in r["edades_varadas"]
                    if es_inflado:
                        estado = "📦 Inventario inflado (hoy > ayer)"
                    elif es_varado:
                        estado = "⚠️ Varado (debió salir)"
                    else:
                        estado = ""
                    filas_lote.append({
                        "Lote (edad ayer → hoy)": f"{e}d → {e_hoy}d",
                        "Cantidad ayer": c_ayer,
                        "Teórico hoy (PEPS)": c_teo,
                        "Real hoy": c_real,
                        "Estado del lote": estado,
                        "_color": "varado" if es_varado else ("inflado" if es_inflado else ""),
                    })
                df_lotes = pd.DataFrame(filas_lote)

                def estilo_lote(row):
                    if row["_color"] == "varado":
                        base = "background-color:#FFE08A; font-weight:700;"   # ámbar
                    elif row["_color"] == "inflado":
                        base = "background-color:#D6E9F8; font-weight:700;"   # azul claro
                    else:
                        base = ""
                    return [base] * len(row)

                st.markdown("**Lotes que venían de ayer**")
                styler_lotes = (
                    df_lotes.style
                    .apply(estilo_lote, axis=1)
                    .format({"Cantidad ayer": "{:,.0f}", "Teórico hoy (PEPS)": "{:,.0f}",
                             "Real hoy": "{:,.0f}"})
                    .hide(axis="columns", subset=["_color"])
                )
                st.dataframe(styler_lotes, use_container_width=True, hide_index=True)

                # --- SECCIÓN 2: Entradas nuevas del período ---
                # Edades reales hoy que NO corresponden a ninguna cohorte de ayer
                edades_cohorte = {e + dias for e in r["vec_ayer"].keys()}
                edad_max_cohorte = max(edades_cohorte) if edades_cohorte else dias
                filas_nuevas = []
                for e_hoy in sorted(r["vec_real"].keys()):
                    if e_hoy not in edades_cohorte:
                        anomala = e_hoy > edad_max_cohorte
                        filas_nuevas.append({
                            "Edad hoy": f"⁉️ {e_hoy}d (reaparecido)" if anomala else f"{e_hoy}d",
                            "Cantidad hoy": r["vec_real"].get(e_hoy, 0),
                            "Nota": "Más viejo de lo posible: revisar reingreso/conteo" if anomala else "",
                        })
                if filas_nuevas:
                    df_nuevas = pd.DataFrame(filas_nuevas)
                    st.markdown("**Entradas nuevas del período** (producto que no existía ayer)")
                    st.dataframe(
                        df_nuevas.style.format({"Cantidad hoy": "{:,.0f}"}),
                        use_container_width=True, hide_index=True,
                    )

        # ----- Formulario del líder de zona: explicar una ruptura -----
        st.divider()
        st.markdown("#### 📝 Registrar explicación de una ruptura")
        if not bd_disponible():
            st.info(
                "La base de datos de seguimiento no está configurada todavía. "
                "Cuando se conecte el Google Sheet, aquí el líder de zona podrá explicar "
                "cada ruptura y alimentar el módulo de seguimiento."
            )
        elif rupturas.empty:
            st.success("No hay rupturas de rotación en este corte para explicar. 🎉")
        else:
            fecha_corte_act = st.session_state.get("_fecha_corte_actual", fecha_corte_hoy)
            gestiones = leer_tabla(HOJA_GESTIONES, COLS_GESTIONES)
            llaves_gestionadas = set(gestiones["llave"].astype(str)) if not gestiones.empty else set()

            # Rupturas pendientes (registradas hoy y aún sin explicación)
            pendientes = []
            for r in rupturas.itertuples():
                lv = llave_ruptura(fecha_corte_act, r.destino, r.item)
                if lv not in llaves_gestionadas:
                    pendientes.append((lv, r.destino, r.item, r.referencia, r.unds_varadas))

            if not pendientes:
                st.success("Todas las rupturas de este corte ya fueron explicadas. ✅")
            else:
                etiquetas = [f"{d} — {it} — {ref} ({uv:,.0f} unds)"
                             for (_, d, it, ref, uv) in pendientes]
                sel_idx = st.selectbox(
                    f"Rupturas pendientes de explicar ({len(pendientes)}):",
                    range(len(pendientes)), format_func=lambda i: etiquetas[i],
                )
                lv, d_sel, it_sel, ref_sel, uv_sel = pendientes[sel_idx]

                with st.form("form_gestion", clear_on_submit=True):
                    lider = st.text_input("Líder de zona *", placeholder="Tu nombre completo")
                    categoria = st.selectbox("Categoría de la causa *", CATEGORIAS_RUPTURA)
                    razon = st.text_area("¿Por qué ocurrió la ruptura? *",
                                         placeholder="Describe la causa concreta…")
                    accion = st.text_area("Acción correctiva",
                                          placeholder="¿Qué se hará para evitar que se repita? (opcional)")
                    enviar = st.form_submit_button("Guardar explicación", use_container_width=True)
                    if enviar:
                        if not lider.strip() or not razon.strip():
                            st.error("El nombre del líder y la razón son obligatorios.")
                        else:
                            ok = guardar_gestion(lv, fecha_corte_act, d_sel, it_sel,
                                                 lider.strip(), categoria, razon.strip(),
                                                 accion.strip())
                            if ok:
                                st.success("Explicación guardada. ¡Gracias! Se reflejará en el seguimiento.")
                                st.rerun()
                            else:
                                st.error("No se pudo guardar. Revisa la conexión a la base de datos.")

    # ===== TAB 2: DETALLE POR DESTINO — UNA FILA POR LOTE (cohorte ayer → hoy) =====
    with tab2:
        st.subheader("Detalle por destino — seguimiento por lote")
        st.caption(
            f"Cada fila sigue un **lote** desde su edad de ayer hasta hoy (envejece +{dias} día(s)). "
            "Los lotes que venían de ayer muestran su evolución; las entradas nuevas del período "
            "aparecen marcadas aparte. La reconstrucción por lote es un modelo PEPS."
        )

        destinos = sorted(res["destino"].unique().tolist())
        cfilt1, cfilt2 = st.columns([2, 1])
        with cfilt1:
            f_dest = st.multiselect("Destino (CEDI/Planta)", destinos, placeholder="Todos")
        with cfilt2:
            solo_rup = st.toggle("Solo items con ruptura", value=False)

        sub = res.copy()
        if f_dest:
            sub = sub[sub["destino"].isin(f_dest)]
        if solo_rup:
            sub = sub[sub["ruptura"]]

        # Explota a nivel de lote/cohorte: por cada (destino, item) seguimos los lotes
        # de ayer (edad ayer -> edad hoy = ayer + dias) y, aparte, las entradas nuevas.
        registros = []
        for r in sub.itertuples():
            edades_cohorte = {e + dias for e in r.vec_ayer.keys()}
            vendio_todo = r.vendido >= r.cant_ayer - 0.5
            # 1) Lotes que venían de ayer
            for e in sorted(r.vec_ayer.keys(), reverse=True):
                e_hoy = e + dias
                c_ayer = r.vec_ayer.get(e, 0)
                c_teo = r.vec_teorico.get(e_hoy, 0)
                c_real = r.vec_real.get(e_hoy, 0)
                es_inflado = e_hoy in r.edades_infladas
                es_varado = e_hoy in r.edades_varadas
                if es_inflado:
                    estado = "📦 Inflado"
                    color = "inflado"
                elif es_varado:
                    estado = "⚠️ Varado"
                    color = "varado"
                else:
                    estado = ""
                    color = ""
                registros.append({
                    "Destino": r.destino,
                    "Item": int(r.item) if str(r.item).isdigit() else r.item,
                    "Referencia": r.referencia,
                    "Lote (ayer → hoy)": f"{e}d → {e_hoy}d",
                    "Cant. ayer": c_ayer,
                    "Teórico hoy": c_teo,
                    "Real hoy": c_real,
                    "Estado": estado,
                    "_orden": 0, "_color": color,
                })
            # 2) Entradas nuevas del período (edades reales que no vienen de ayer)
            edad_max_cohorte = max(edades_cohorte) if edades_cohorte else dias
            for e_hoy in sorted(r.vec_real.keys()):
                if e_hoy not in edades_cohorte:
                    c_real = r.vec_real.get(e_hoy, 0)
                    # Una "entrada nueva" más vieja que cualquier cohorte posible de ayer
                    # no puede ser producto fresco: es una reaparición anómala.
                    anomala = e_hoy > edad_max_cohorte and r.ruptura
                    registros.append({
                        "Destino": r.destino,
                        "Item": int(r.item) if str(r.item).isdigit() else r.item,
                        "Referencia": r.referencia,
                        "Lote (ayer → hoy)": f"⁉️ {e_hoy}d (reaparecido)" if anomala else f"nuevo → {e_hoy}d",
                        "Cant. ayer": 0,
                        "Teórico hoy": 0,
                        "Real hoy": c_real,
                        "Estado": "⚠️ Reaparecido (anómalo)" if anomala else "🆕 Entrada nueva",
                        "_orden": 1, "_color": "varado" if anomala else "nueva",
                    })

        if not registros:
            st.info("No hay registros para los filtros seleccionados.")
        else:
            tabla = pd.DataFrame(registros).sort_values(
                ["Destino", "Item", "_orden", "Lote (ayer → hoy)"],
                ascending=[True, True, True, False]
            ).reset_index(drop=True)

            def estilo_fila(row):
                if row["_color"] == "varado":
                    return ["background-color:#FDEDEC;"] * len(row)          # rojo claro
                if row["_color"] == "inflado":
                    return ["background-color:#D6E9F8;"] * len(row)          # azul claro
                if row["_color"] == "nueva":
                    return ["background-color:#F4F9F2; color:#4A4A4A;"] * len(row)
                return [""] * len(row)

            vista_cols = ["Destino", "Item", "Referencia", "Lote (ayer → hoy)",
                          "Cant. ayer", "Teórico hoy", "Real hoy", "Estado", "_color"]
            styler = (
                tabla[vista_cols].style
                .apply(estilo_fila, axis=1)
                .format({"Cant. ayer": "{:,.0f}", "Teórico hoy": "{:,.0f}",
                         "Real hoy": "{:,.0f}", "Item": "{:.0f}"})
                .hide(axis="columns", subset=["_color"])
            )
            st.dataframe(styler, use_container_width=True, hide_index=True, height=600)
            st.markdown(
                "⚠️ **Varado** (rojo) = lote viejo que debió salir y no salió → revisar rotación.  "
                "📦 **Inflado** (azul) = hoy hay más que ayer, imposible por envejecimiento → "
                "revisar ingreso o conteo.  🆕 Entradas nuevas en verde.  Filas en blanco rotaron bien."
            )
            st.caption(f"{len(tabla):,} lotes mostrados.")

    # ===== TAB 3: RIESGO DE VIDA ÚTIL (proyección PEPS lote por lote) =====
    with tab3:
        st.subheader("Riesgo de superar 5 días de vida útil")
        st.caption(
            "Proyección a futuro: con el ritmo de venta diaria (promedio para SUELTO, "
            "venta del día para PET) y la edad actual de cada lote, se estima la edad que "
            "tendría al venderse bajo PEPS (lo más viejo primero, cada lote espera a los "
            "anteriores). Se alerta si algún lote proyecta superar **5 días**."
        )

        UMBRAL_VIDA = 5
        filas_riesgo = []
        # Inventario de hoy por (destino, item)
        claves_hoy = inv_hoy.groupby(["destino", "item"]).size().index.tolist()
        for dest, item in claves_hoy:
            vdiaria = venta_diaria.get((dest, item), 0.0)
            inv_vec = vec_por_edad(inv_hoy, dest, item)
            tot = sum(inv_vec.values())
            if tot <= 0:
                continue
            proy = proyectar_vida_util(inv_vec, vdiaria, umbral=UMBRAL_VIDA)
            en_riesgo = proy["riesgo_proyectado"]
            ya_vencido = proy["ya_vencido"]
            edad_max_proy = proy["edad_max_proyectada"]
            ref = inv_hoy[(inv_hoy["destino"] == dest) & (inv_hoy["item"] == item)]["referencia"].iloc[0] \
                if not inv_hoy[(inv_hoy["destino"] == dest) & (inv_hoy["item"] == item)].empty else ""
            filas_riesgo.append({
                "Destino": dest,
                "Item": int(item) if str(item).isdigit() else item,
                "Referencia": ref,
                "Categoría": cat_map.get((dest, item), "SUELTO"),
                "Inv. hoy": tot,
                "Venta diaria": round(vdiaria, 0),
                "Días cobertura": round(tot / vdiaria, 1) if vdiaria > 0 else None,
                "Edad máx. proyectada": edad_max_proy if vdiaria > 0 else None,
                "Unds en riesgo": proy["unds_riesgo"],
                "_riesgo": en_riesgo,
                "_vencido": ya_vencido,
                "_tienda": es_tienda(ref),
            })

        df_riesgo = pd.DataFrame(filas_riesgo)
        if df_riesgo.empty:
            st.info("No hay inventario para proyectar.")
        else:
            # Filtros (primero, para que los KPIs respeten el de tienda)
            cfa, cfb, cfc = st.columns([2, 1, 1])
            with cfa:
                f_dest_r = st.multiselect("Destino", sorted(df_riesgo["Destino"].unique()),
                                          placeholder="Todos", key="riesgo_dest")
            with cfb:
                solo_riesgo = st.toggle("Solo con alerta", value=True, key="riesgo_toggle",
                                        help="Muestra solo los que van a vencer o ya están vencidos hoy.")
            with cfc:
                solo_tienda = st.toggle("Solo productos de tienda", value=False,
                                        key="riesgo_tienda",
                                        help="Producto de tienda: HUEVO (talla) X (n) CARTON VERDE CANASTA. "
                                             "El resto suele ser preventa.")

            # Base para KPIs: respeta destino y tienda, pero no 'solo en riesgo'
            base = df_riesgo.copy()
            if f_dest_r:
                base = base[base["Destino"].isin(f_dest_r)]
            if solo_tienda:
                base = base[base["_tienda"]]

            n_riesgo = int(base["_riesgo"].sum())
            unds_tot_riesgo = base.loc[base["_riesgo"], "Unds en riesgo"].sum()
            kr1, kr2 = st.columns(2)
            with kr1:
                st.markdown(tarjeta_kpi("SKU/destino en riesgo (va a vencer)", f"{n_riesgo:,}",
                                        estado="warning" if n_riesgo else "neutral", reina=True),
                            unsafe_allow_html=True)
            with kr2:
                st.markdown(tarjeta_kpi("Unidades en riesgo", f"{unds_tot_riesgo:,.0f}",
                                        estado="warning"), unsafe_allow_html=True)
            st.caption(
                "**Va a vencer** = el lote nace sano (≤5 días) pero el ritmo de venta lo "
                "lleva a superar 5 días antes de agotarse → accionable bajando inventario o "
                "acelerando rotación. El producto ya vencido se revisa en el módulo de Edades."
            )

            st.divider()

            # Tabla: base + el filtro de 'solo en riesgo'
            vista = base.copy()
            if solo_riesgo:
                vista = vista[vista["_riesgo"] | vista["_vencido"]]
            vista = vista.sort_values(["_riesgo", "_vencido", "Edad máx. proyectada"],
                                      ascending=[False, False, False])

            def estilo_riesgo(row):
                # Solo se resalta en ámbar lo que VA a vencer (accionable). El producto
                # ya vencido no se pinta aquí porque ya es visible en el módulo de Edades.
                if row["_riesgo"]:
                    return ["background-color:#FFE08A;"] * len(row)
                return [""] * len(row)

            cols_v = ["Destino", "Item", "Referencia", "Categoría", "Inv. hoy",
                      "Venta diaria", "Días cobertura", "Edad máx. proyectada",
                      "Unds en riesgo", "_riesgo", "_vencido"]
            styler = (
                vista[cols_v].style
                .apply(estilo_riesgo, axis=1)
                .format({"Inv. hoy": "{:,.0f}", "Venta diaria": "{:,.0f}",
                         "Días cobertura": "{:.1f}", "Edad máx. proyectada": "{:.1f}",
                         "Unds en riesgo": "{:,.0f}",
                         "Item": "{:.0f}"}, na_rep="—")
                .hide(axis="columns", subset=["_riesgo", "_vencido"])
            )
            st.dataframe(styler, use_container_width=True, hide_index=True, height=520)
            st.markdown(
                "🟡 **Va a vencer** (ámbar): rota lento y superará 5 días antes de agotarse. "
                "El producto **ya vencido hoy** se puede revisar en el módulo de Edades. "
                "**Edad máx. proyectada** = edad del último lote al venderse bajo PEPS."
            )
            st.caption(f"{len(vista):,} SKU/destino mostrados.")

    st.caption(
        f"Fuentes: {ARCHIVO_AYER} (inicial ayer) · {ARCHIVO_VENTAS} (ventas del período) · "
        f"{ARCHIVO_HOY} (inicial hoy) — hoja {HOJA_INV_ANALISIS}"
    )


# ===========================================================================
# MÓDULO 3 — SEGUIMIENTO DE RUPTURAS (gestión del proceso)
# ===========================================================================
def render_modulo_seguimiento():
    st.markdown('<p class="titulo-modulo">📈 Seguimiento de Rupturas</p>', unsafe_allow_html=True)
    st.caption(
        "Mide el nivel de gestión: cuántas rupturas de rotación se detectaron día a día "
        "y cuántas fueron explicadas por los líderes de zona."
    )
    st.divider()

    if not bd_disponible():
        st.warning(
            "La base de datos de seguimiento aún no está configurada. Sigue la guía de "
            "conexión a Google Sheets para activar este módulo. Una vez conectado, aquí verás "
            "la línea de tiempo de rupturas y tu indicador de gestión."
        )
        if st.session_state.get("_bd_error"):
            st.caption(f"Detalle técnico: {st.session_state['_bd_error']}")
        return

    rup = leer_tabla(HOJA_RUPTURAS, COLS_RUPTURAS)
    ges = leer_tabla(HOJA_GESTIONES, COLS_GESTIONES)

    if rup.empty:
        st.info("Todavía no hay rupturas registradas en la base de datos.")
        return

    # Tipos y marca de gestión
    rup["unds_varadas"] = pd.to_numeric(rup["unds_varadas"], errors="coerce").fillna(0)
    llaves_gest = set(ges["llave"].astype(str)) if not ges.empty else set()
    rup["gestionada"] = rup["llave"].astype(str).isin(llaves_gest)

    total = len(rup)
    gestionadas = int(rup["gestionada"].sum())
    pendientes = total - gestionadas
    pct = (gestionadas / total * 100) if total else 0

    # --- KPIs de gestión ---
    k1, k2, k3 = st.columns(3)
    with k1:
        st.markdown(tarjeta_kpi("Nivel de gestión", f"{pct:,.0f}%",
                                estado="critical" if pct < 50 else "neutral", reina=True),
                    unsafe_allow_html=True)
    with k2:
        st.markdown(tarjeta_kpi("Rupturas gestionadas", f"{gestionadas:,}", estado="neutral"),
                    unsafe_allow_html=True)
    with k3:
        st.markdown(tarjeta_kpi("Rupturas pendientes", f"{pendientes:,}",
                                estado="warning" if pendientes else "neutral"),
                    unsafe_allow_html=True)
    st.divider()

    # --- Línea de tiempo: rupturas por fecha de corte, gestionadas vs pendientes ---
    st.subheader("Línea de tiempo de rupturas")
    serie = (rup.groupby(["fecha_corte", "gestionada"]).size()
             .reset_index(name="n"))
    pivot = serie.pivot(index="fecha_corte", columns="gestionada", values="n").fillna(0)
    pivot = pivot.rename(columns={True: "Gestionadas", False: "Pendientes"})
    for col in ("Gestionadas", "Pendientes"):
        if col not in pivot.columns:
            pivot[col] = 0
    pivot = pivot.sort_index()

    fig = go.Figure()
    fig.add_trace(go.Bar(x=pivot.index, y=pivot["Gestionadas"], name="Gestionadas",
                         marker_color=COLOR_PRIMARIO))
    fig.add_trace(go.Bar(x=pivot.index, y=pivot["Pendientes"], name="Pendientes",
                         marker_color="#D0021B"))
    # Línea de total por día
    total_dia = pivot["Gestionadas"] + pivot["Pendientes"]
    fig.add_trace(go.Scatter(x=pivot.index, y=total_dia, name="Total rupturas",
                             mode="lines+markers", line=dict(color=COLOR_ACENTO, width=3)))
    fig.update_layout(
        barmode="stack", height=380,
        margin=dict(l=10, r=10, t=10, b=10),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Nunito, sans-serif", size=13),
        xaxis=dict(title="Fecha de corte"),
        yaxis=dict(title="N° de rupturas", showgrid=True, gridcolor="#EEEEEE"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # --- Detalle de gestiones registradas ---
    st.subheader("Explicaciones registradas")
    if ges.empty:
        st.info("Aún no se han registrado explicaciones de los líderes.")
    else:
        vista = ges.rename(columns={
            "fecha_corte": "Fecha corte", "destino": "Destino", "item": "Item",
            "lider_zona": "Líder de zona", "categoria": "Categoría", "razon": "Razón",
            "accion_correctiva": "Acción correctiva", "fecha_gestion": "Registrado",
        })[["Fecha corte", "Destino", "Item", "Líder de zona", "Categoría",
            "Razón", "Acción correctiva", "Registrado"]]
        st.dataframe(vista, use_container_width=True, hide_index=True, height=360)

        # Causas más frecuentes
        st.subheader("Causas más frecuentes")
        causas = ges["categoria"].value_counts().reset_index()
        causas.columns = ["Categoría", "N° de casos"]
        figc = go.Figure(go.Bar(
            x=causas["N° de casos"], y=causas["Categoría"], orientation="h",
            marker_color=COLOR_PRIMARIO,
            text=causas["N° de casos"], textposition="outside",
        ))
        figc.update_layout(
            height=260, margin=dict(l=10, r=10, t=10, b=10),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Nunito, sans-serif", size=13),
            xaxis=dict(title="Casos", showgrid=True, gridcolor="#EEEEEE"),
            yaxis=dict(autorange="reversed"),
        )
        st.plotly_chart(figc, use_container_width=True)


# ===========================================================================
# NAVEGACIÓN
# ===========================================================================
with st.sidebar:
    st.markdown(f"<h2 style='color:{COLOR_PRIMARIO}; font-weight:900;'>🥚 Huevos Kikes</h2>",
                unsafe_allow_html=True)
    st.markdown("**Panel de Inventarios**")
    modulo = st.radio(
        "Módulo",
        ["Inventario de Edades", "Análisis de Rotación PEPS", "Seguimiento de Rupturas"],
        label_visibility="collapsed",
    )
    st.divider()
    if st.button("🔄 Actualizar datos", use_container_width=True,
                 help="Limpia el caché y vuelve a leer los archivos del repositorio."):
        st.cache_data.clear()
        st.rerun()
    st.caption(
        "Sube a la raíz del repositorio (con espacio o guion bajo):\n\n"
        f"• **{ARCHIVO_HOY}**\n\n"
        f"• **{ARCHIVO_AYER}**\n\n"
        f"• **{ARCHIVO_VENTAS}**"
    )

if modulo == "Inventario de Edades":
    render_modulo_edades()
elif modulo == "Análisis de Rotación PEPS":
    render_modulo_rotacion()
else:
    render_modulo_seguimiento()
