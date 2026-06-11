"""
Dashboard de Inventario de Edades - Huevos Kikes
================================================
Muestra el inventario actual de edades en plantas y centros de distribución.

Fuente de datos: archivo 'Inventario Hoy.xlsx', hoja 'INV. EDADES'
(debe estar en el mismo repositorio de GitHub que este archivo).

Columnas utilizadas:
    CEDI, fecha, edad, item, cantidad, referencia, destino, tipo, tipo huevo, zona
"""

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ---------------------------------------------------------------------------
# CONFIGURACIÓN GENERAL
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Dashboard Inventario de Edades | Huevos Kikes",
    page_icon="🥚",
    layout="wide",
)

# Paleta de marca Huevos Kikes
COLOR_PRIMARIO = "#3DAE2B"   # verde Kikes
COLOR_ACENTO = "#F7941D"     # naranja (yema del logo)
COLOR_TEXTO = "#1A1A1A"

ARCHIVO_DATOS = "Inventario Hoy.xlsx"
HOJA_DATOS = "INV. EDADES"

# Estilos
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
        /* Acentos por estado */
        .kpi-neutral {{ border-left-color: {COLOR_PRIMARIO}; }}
        .kpi-warning {{ border-left-color: #F5A623; }}
        .kpi-critical {{ border-left-color: #D0021B; }}
        /* Tarjeta reina: edad promedio ponderada */
        .kpi-reina {{
            background: linear-gradient(135deg, #F2FAF0 0%, #FFFFFF 100%);
            border-left: 9px solid {COLOR_PRIMARIO};
            border-top: 1px solid {COLOR_PRIMARIO}33;
        }}
        .kpi-reina .kpi-label {{ color: {COLOR_PRIMARIO}; font-size: 1rem; }}
        .kpi-reina .kpi-value {{ font-size: 2.2rem; }}
        .kpi-crown {{
            font-size: 0.75rem;
            font-weight: 800;
            color: {COLOR_ACENTO};
            letter-spacing: 0.5px;
            text-transform: uppercase;
            margin: 0 0 2px 0;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)


def tarjeta_kpi(label, value, estado="neutral", reina=False):
    """Genera el HTML de una tarjeta KPI con acento condicional."""
    clases = "kpi-card"
    if reina:
        clases += " kpi-reina"
        corona = '<p class="kpi-crown">★ Métrica clave</p>'
    else:
        clases += f" kpi-{estado}"
        corona = ""
    return (
        f'<div class="{clases}">{corona}'
        f'<p class="kpi-label">{label}</p>'
        f'<p class="kpi-value">{value}</p></div>'
    )


# ---------------------------------------------------------------------------
# CARGA DE DATOS
# ---------------------------------------------------------------------------
@st.cache_data(ttl=3600)
def cargar_datos(ruta: str, hoja: str) -> pd.DataFrame:
    df = pd.read_excel(ruta, sheet_name=hoja)

    # Normaliza los nombres de columnas: minúsculas y sin espacios extra
    df.columns = [str(c).strip().lower() for c in df.columns]

    # Renombra "tipo huevo" si viniera con variaciones
    renombres = {}
    for col in df.columns:
        if col in ("tipo huevo", "tipo_huevo", "tipohuevo"):
            renombres[col] = "tipo huevo"
    df = df.rename(columns=renombres)

    # Tipos numéricos seguros
    if "edad" in df.columns:
        df["edad"] = pd.to_numeric(df["edad"], errors="coerce").astype("Int64")
    if "cantidad" in df.columns:
        df["cantidad"] = pd.to_numeric(df["cantidad"], errors="coerce").fillna(0)

    return df


try:
    df = cargar_datos(ARCHIVO_DATOS, HOJA_DATOS)
except FileNotFoundError:
    st.error(
        f"No se encontró el archivo **{ARCHIVO_DATOS}**. "
        "Asegúrate de que esté en la raíz del repositorio de GitHub."
    )
    st.stop()
except ValueError:
    st.error(
        f"No se encontró la hoja **{HOJA_DATOS}** dentro de {ARCHIVO_DATOS}. "
        "Verifica el nombre exacto de la hoja."
    )
    st.stop()


# ---------------------------------------------------------------------------
# ENCABEZADO
# ---------------------------------------------------------------------------
st.markdown('<p class="titulo-principal">🥚 Inventario de Edades</p>', unsafe_allow_html=True)
st.divider()


# ---------------------------------------------------------------------------
# FILTROS (multiselección en el encabezado)
# ---------------------------------------------------------------------------
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

# Aplica filtros
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


# ---------------------------------------------------------------------------
# KPIs
# ---------------------------------------------------------------------------
inv_total = dff["cantidad"].sum()
und_mas_6 = dff.loc[dff["edad"] > 6, "cantidad"].sum()
und_mas_10 = dff.loc[dff["edad"] >= 10, "cantidad"].sum()
# Edad ponderada: ignora filas sin edad para no romper el cálculo
_val = dff.dropna(subset=["edad"])
_peso = _val["cantidad"].sum()
edad_prom = (
    (_val["edad"].astype(float) * _val["cantidad"]).sum() / _peso if _peso > 0 else 0
)
pct_mas_6 = (und_mas_6 / inv_total * 100) if inv_total > 0 else 0

k1, k2, k3, k4, k5 = st.columns(5)

# Métrica reina al extremo izquierdo (lo primero que se lee)
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


# ---------------------------------------------------------------------------
# HISTOGRAMA DE DISTRIBUCIÓN DE INVENTARIO POR EDAD
# ---------------------------------------------------------------------------
st.subheader("Distribución del inventario por edad")

# Agrupa cantidad por edad; agrupa 10+ en una sola barra
dist = dff.dropna(subset=["edad"]).copy()
dist["edad_int"] = dist["edad"].astype(int)
dist["bucket"] = dist["edad_int"].apply(lambda x: "10+" if x >= 10 else str(x))

# Orden correcto del eje X
orden_buckets = [str(i) for i in range(0, 10)] + ["10+"]
serie = (
    dist.groupby("bucket")["cantidad"].sum()
    .reindex(orden_buckets, fill_value=0)
)


def color_barra(bucket):
    if bucket == "10+":
        return "#D0021B"          # rojo crítico
    val = int(bucket)
    if val <= 5:
        return COLOR_PRIMARIO     # verde óptimo
    elif val <= 9:
        return "#F5A623"          # amarillo/naranja advertencia
    return "#D0021B"


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


# ---------------------------------------------------------------------------
# TABLA DETALLADA CON SEMÁFORO POR EDAD
# ---------------------------------------------------------------------------
st.subheader("Detalle de inventario")

# Columnas a mostrar (en el orden de la imagen de referencia)
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
    """Semáforo por rango de edad, con intensidad creciente."""
    try:
        v = float(val)
    except (TypeError, ValueError):
        return ""
    if 1 <= v <= 5:
        return "background-color: #C6EFCE; color: #006100; font-weight: 700;"   # verde - óptimo
    elif v == 6:
        return "background-color: #FFE08A; color: #7A5200; font-weight: 700;"    # amarillo
    elif 7 <= v <= 9:
        return "background-color: #FFB84D; color: #7A3E00; font-weight: 800;"    # naranja vivo (día 7+)
    elif v >= 10:
        return "background-color: #FF8A80; color: #7A0006; font-weight: 800;"    # rojo claro - crítico
    return ""


# Asegura que Item se muestre como entero (es un código, sin decimales)
if "Item" in tabla.columns:
    tabla["Item"] = pd.to_numeric(tabla["Item"], errors="coerce").astype("Int64")

styler = (
    tabla.style
    .map(color_edad, subset=["Edad"])
    .bar(
        subset=["Suma de Cantidad"],
        color="#BFE3B5",          # verde Kikes suave para las barras de datos
        align="left",
        height=70,
        vmin=0,
    )
    .format({"Suma de Cantidad": "{:,.0f}", "Edad": "{:.0f}", "Item": "{:.0f}"})
)

st.dataframe(styler, use_container_width=True, hide_index=True, height=600)

# Leyenda del semáforo
st.markdown(
    """
    **Convención de edades:**
    🟢 1–5 días: óptimo &nbsp;&nbsp; 🟡 6 días: alerta &nbsp;&nbsp; 🟠 7–9 días: preocupante &nbsp;&nbsp; 🔴 10+ días: crítico
    &nbsp;&nbsp;|&nbsp;&nbsp; Las barras en *Suma de Cantidad* son proporcionales al volumen de cada fila.
    """
)

st.caption(f"Fuente: {ARCHIVO_DATOS} — hoja {HOJA_DATOS}")
