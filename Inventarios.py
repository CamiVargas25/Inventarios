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
            font-size: 3.4rem;
            font-weight: 900;
            margin-bottom: 0.4rem;
        }}
        div[data-testid="stMetric"] {{
            background-color: #F2FAF0;
            border-left: 5px solid {COLOR_PRIMARIO};
            border-radius: 8px;
            padding: 12px 16px;
        }}
    </style>
    """,
    unsafe_allow_html=True,
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
k1.metric("Inventario total (und)", f"{inv_total:,.0f}")
k2.metric("Unidades con más de 6 días", f"{und_mas_6:,.0f}")
k3.metric("% con más de 6 días", f"{pct_mas_6:,.1f}%")
k4.metric("Unidades con 10+ días", f"{und_mas_10:,.0f}")
k5.metric("Edad promedio ponderada", f"{edad_prom:,.1f} días")

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
    """Semáforo por rango de edad."""
    try:
        v = float(val)
    except (TypeError, ValueError):
        return ""
    if 1 <= v <= 5:
        return "background-color: #C6EFCE; color: #006100;"   # verde claro - óptimo
    elif 6 <= v <= 9:
        return "background-color: #FFEB9C; color: #9C6500;"    # amarillo - preocupante
    elif v >= 10:
        return "background-color: #FFC7CE; color: #9C0006;"    # rojo - preocupante
    return ""


# Asegura que Item se muestre como entero (es un código, sin decimales)
if "Item" in tabla.columns:
    tabla["Item"] = pd.to_numeric(tabla["Item"], errors="coerce").astype("Int64")

styler = (
    tabla.style
    .map(color_edad, subset=["Edad"])
    .format({"Suma de Cantidad": "{:,.0f}", "Edad": "{:.0f}", "Item": "{:.0f}"})
)

st.dataframe(styler, use_container_width=True, hide_index=True, height=600)

# Leyenda del semáforo
st.markdown(
    """
    **Convención de edades:**
    🟢 1–5 días: óptimo &nbsp;&nbsp; 🟡 6–9 días: preocupante &nbsp;&nbsp; 🔴 10+ días: crítico
    """
)

st.caption(f"Fuente: {ARCHIVO_DATOS} — hoja {HOJA_DATOS}")
