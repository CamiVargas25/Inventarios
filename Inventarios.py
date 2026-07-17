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

import base64
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
ARCHIVO_PEDIDOS = resolver_archivo("19.1 Pedidos.xlsx", "19.1_Pedidos.xlsx")
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

# Enlace al Google Sheet histórico de rupturas (BD manual administrada por Camila).
# Los líderes de zona entran con sus credenciales corporativas y escriben la
# explicación de cada ruptura directamente en este Sheet. El dashboard NO escribe
# en él por API: solo exporta el CSV de rupturas del día y enlaza aquí.
# >>> Pega aquí la URL de tu Sheet histórico cuando lo tengas creado <<<
URL_SHEET_HISTORICO = "https://docs.google.com/spreadsheets/d/1lp14wEJ0kbf1FTspk70HXniBcz6VrGYSVgOuy9rF4Y8/edit?usp=sharing"

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


HOJA_INV_DETALLE = "inv"  # hoja de inventario detallado (crudo ERP): trae las
                          # referencias "DESECHO" que no aparecen en INV. EDADES.


@st.cache_data(ttl=3600)
def cargar_desecho_por_planta(ruta: str, cache_key: float = 0.0) -> pd.DataFrame:
    """Lee la hoja 'inv' y agrupa por planta/CEDI ('id_item_bodega2') el inventario
    cuya referencia contiene la palabra 'desecho'."""
    df = pd.read_excel(ruta, sheet_name=HOJA_INV_DETALLE)
    df.columns = [str(c).strip() for c in df.columns]
    out = pd.DataFrame({
        "planta": df.get("id_item_bodega2"),
        "referencia": df.get("descripcion_articulo"),
        "cantidad": pd.to_numeric(df.get("cantidad"), errors="coerce").fillna(0.0),
    })
    out = out.dropna(subset=["planta"])
    es_desecho = out["referencia"].astype(str).str.upper().str.contains("DESECHO", na=False)
    out = out[es_desecho & (out["cantidad"] > 0)]
    return (
        out.groupby("planta", as_index=False)["cantidad"]
        .sum()
        .sort_values("cantidad", ascending=False)
        .reset_index(drop=True)
    )


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
    und_mas_6 = dff.loc[dff["edad"] >= 5, "cantidad"].sum()
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
            tarjeta_kpi("Unidades con 5 días o más", f"{und_mas_6:,.0f}", estado="warning"),
            unsafe_allow_html=True,
        )
    with k4:
        st.markdown(
            tarjeta_kpi("% con 5 días o más", f"{pct_mas_6:,.1f}%", estado="warning"),
            unsafe_allow_html=True,
        )
    with k5:
        st.markdown(
            tarjeta_kpi("Unidades con 10+ días", f"{und_mas_10:,.0f}", estado="critical"),
            unsafe_allow_html=True,
        )

    st.divider()
    st.subheader("🗑️ Alerta de Inventario Desecho")
    st.caption(
        "Referencias cuyo nombre contiene la palabra **desecho**, agrupadas por planta/CEDI."
    )

    try:
        desecho = cargar_desecho_por_planta(ARCHIVO_HOY, mtime(ARCHIVO_HOY))
    except Exception as e:
        desecho = pd.DataFrame()
        st.warning(
            f"⚠️ No se pudo leer el inventario DESECHO desde '{ARCHIVO_HOY}': "
            f"{type(e).__name__}: {e}"
        )

    if desecho.empty:
        st.success("No se detectó inventario DESECHO. 🎉")
    else:
        total_desecho = desecho["cantidad"].sum()
        n_plantas_desecho = desecho["planta"].nunique()
        st.markdown(
            f'<div style="background-color:#FDEDEC; border-left:6px solid {COLOR_CRITICO}; '
            f'padding:12px 16px; border-radius:6px; margin-bottom:12px;">'
            f'🚨 <b>Inventario DESECHO detectado</b> en {n_plantas_desecho} planta(s)/CEDI: '
            f'<b>{total_desecho:,.0f} und.</b></div>',
            unsafe_allow_html=True,
        )
        tabla_desecho = desecho.rename(columns={"planta": "Planta/CEDI", "cantidad": "Cantidad"})
        st.dataframe(
            tabla_desecho.style.format({"Cantidad": "{:,.0f}"}).bar(
                subset=["Cantidad"], color="#F5B7B1", align="left", vmin=0
            ),
            use_container_width=True,
            hide_index=True,
        )

    st.divider()
    st.subheader("Distribución del inventario por edad")

    dist = dff.dropna(subset=["edad"]).copy()
    dist["edad_int"] = dist["edad"].astype(int)
    dist["bucket"] = dist["edad_int"].apply(lambda x: "10+" if x >= 10 else str(x))
    orden_buckets = [str(i) for i in range(1, 10)] + ["10+"]
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
        if 1 <= v <= 4:
            return "background-color: #C6EFCE; color: #006100; font-weight: 700;"
        elif v == 5:
            return "background-color: #FFE08A; color: #7A5200; font-weight: 700;"
        elif 6 <= v <= 9:
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
        🟢 1–4 días: óptimo &nbsp;&nbsp; 🟡 5 días: alerta &nbsp;&nbsp; 🟠 6–9 días: preocupante &nbsp;&nbsp; 🔴 10+ días: crítico
        &nbsp;&nbsp;|&nbsp;&nbsp; Las barras en *Suma de Cantidad* son proporcionales al volumen de cada fila.
        """
    )
    st.divider()
    st.subheader("⏳ Riesgo de Vida Útil")
    st.caption(
        "Proyección a futuro: con el ritmo de venta diaria y la edad actual de cada lote, "
        "se estima la edad que tendría al venderse bajo PEPS (más viejo primero). "
        "Se alerta si algún lote proyecta superar **5 días**."
    )

    try:
        inv_hoy_r = leer_inventario(ARCHIVO_HOY, HOJA_INV_ANALISIS, mtime(ARCHIVO_HOY))
    except (FileNotFoundError, ValueError):
        inv_hoy_r = None

    ventas_r = None
    try:
        ventas_r = leer_ventas(ARCHIVO_VENTAS, mtime(ARCHIVO_VENTAS))
    except FileNotFoundError:
        pass

    if inv_hoy_r is None:
        st.warning("No se pudo cargar el inventario de hoy para la proyección.")
    elif ventas_r is None:
        st.info(f"Sube **{ARCHIVO_VENTAS}** para activar la proyección de riesgo de vida útil.")
    else:
        f_hoy_r = leer_fecha_corte(ARCHIVO_HOY, HOJA_INV_ANALISIS, mtime(ARCHIVO_HOY))
        fecha_corte_obj_r = f_hoy_r.date() if f_hoy_r is not None else _dt.date.today()
        fecha_corte_hoy_r = fecha_corte_obj_r.isoformat()
        _, cat_map_r, venta_diaria_r, dias_rango_r = preparar_ventas_peps(
            ventas_r, fecha_corte_obj_r, 1)

        if dias_rango_r > 1:
            st.info(
                f"📅 El promedio diario de **SUELTO** se calcula sobre **{dias_rango_r} días "
                "operativos** del rango (se excluyen domingos). "
                f"Para **PET** se usa la venta exacta del día de corte ({fecha_corte_hoy_r})."
            )

        UMBRAL_VIDA = 5
        filas_riesgo = []
        claves_hoy_r = inv_hoy_r.groupby(["destino", "item"]).size().index.tolist()
        for dest, item in claves_hoy_r:
            vdiaria = venta_diaria_r.get((dest, item), 0.0)
            inv_vec = vec_por_edad(inv_hoy_r, dest, item)
            tot = sum(inv_vec.values())
            if tot <= 0:
                continue
            edad_max_actual = max(inv_vec.keys()) if inv_vec else 0
            proy = proyectar_vida_util(inv_vec, vdiaria, umbral=UMBRAL_VIDA)
            ref_rows = inv_hoy_r[(inv_hoy_r["destino"] == dest) & (inv_hoy_r["item"] == item)]
            ref = ref_rows["referencia"].iloc[0] if not ref_rows.empty else ""
            filas_riesgo.append({
                "Destino": dest,
                "Item": int(item) if str(item).isdigit() else item,
                "Referencia": ref,
                "Categoría": cat_map_r.get((dest, item), "SUELTO"),
                "Inv. hoy": tot,
                "Venta diaria": round(vdiaria, 0),
                "Días cobertura": round(tot / vdiaria, 1) if vdiaria > 0 else None,
                "Edad máx. actual": edad_max_actual,
                "Edad máx. proyectada": proy["edad_max_proyectada"] if vdiaria > 0 else None,
                "Unds en riesgo": proy["unds_riesgo"],
                "_riesgo": proy["riesgo_proyectado"],
                "_tienda": es_tienda(ref),
            })

        df_riesgo = pd.DataFrame(filas_riesgo)
        if df_riesgo.empty:
            st.info("No hay inventario para proyectar.")
        else:
            cfa, cfc = st.columns([2, 1])
            with cfa:
                f_dest_r = st.multiselect("Destino", sorted(df_riesgo["Destino"].unique()),
                                          placeholder="Todos", key="riesgo_dest")
            with cfc:
                solo_tienda_r = st.toggle("Solo productos de tienda", value=False,
                                          key="riesgo_tienda",
                                          help="Producto de tienda: HUEVO (talla) X (n) CARTON VERDE CANASTA.")

            base_r = df_riesgo.copy()
            if f_dest_r:
                base_r = base_r[base_r["Destino"].isin(f_dest_r)]
            if solo_tienda_r:
                base_r = base_r[base_r["_tienda"]]

            n_riesgo = int(base_r["_riesgo"].sum())
            unds_tot_riesgo = base_r.loc[base_r["_riesgo"], "Unds en riesgo"].sum()
            unds_actuales_r = base_r["Inv. hoy"].sum()
            pct_riesgo = (unds_tot_riesgo / unds_actuales_r * 100) if unds_actuales_r > 0 else 0
            kr1, kr2, kr3, kr4 = st.columns(4)
            with kr1:
                st.markdown(tarjeta_kpi("SKU/destino en riesgo (va a vencer)", f"{n_riesgo:,}",
                                        estado="warning" if n_riesgo else "neutral", reina=True),
                            unsafe_allow_html=True)
            with kr2:
                st.markdown(tarjeta_kpi("Unidades en riesgo", f"{unds_tot_riesgo:,.0f}",
                                        estado="warning"), unsafe_allow_html=True)
            with kr3:
                st.markdown(tarjeta_kpi("Unds actuales", f"{unds_actuales_r:,.0f}",
                                        estado="neutral"), unsafe_allow_html=True)
            with kr4:
                st.markdown(tarjeta_kpi("% en riesgo", f"{pct_riesgo:,.1f}%",
                                        estado="warning" if pct_riesgo > 0 else "neutral"),
                            unsafe_allow_html=True)
            st.caption(
                "**Va a vencer** = el lote nace sano (≤5 días) pero el ritmo de venta lo lleva "
                "a superar 5 días antes de agotarse → accionable bajando inventario o acelerando rotación."
            )

            st.divider()

            vista_r = base_r[base_r["_riesgo"]].copy()
            vista_r = vista_r.sort_values("Edad máx. proyectada", ascending=False)

            if vista_r.empty:
                st.success("No hay producto en riesgo de vencer con los filtros actuales. 🎉")
            else:
                def _color_semaforo(val):
                    try:
                        v = float(val)
                    except (TypeError, ValueError):
                        return ""
                    if v <= 4:
                        return "background-color: #C6EFCE; color: #006100; font-weight: 700;"
                    elif v == 5:
                        return "background-color: #FFE08A; color: #7A5200; font-weight: 700;"
                    elif 6 <= v <= 9:
                        return "background-color: #FFB84D; color: #7A3E00; font-weight: 800;"
                    else:
                        return "background-color: #FF8A80; color: #7A0006; font-weight: 800;"

                cols_v = ["Destino", "Item", "Referencia", "Categoría", "Inv. hoy",
                          "Venta diaria", "Días cobertura", "Edad máx. actual",
                          "Edad máx. proyectada"]
                styler_r = (
                    vista_r[cols_v].style
                    .map(_color_semaforo, subset=["Edad máx. actual", "Edad máx. proyectada"])
                    .format({"Inv. hoy": "{:,.0f}", "Venta diaria": "{:,.0f}",
                             "Días cobertura": "{:.1f}", "Edad máx. actual": "{:.0f}",
                             "Edad máx. proyectada": "{:.1f}",
                             "Item": "{:.0f}"}, na_rep="—")
                )
                st.dataframe(styler_r, use_container_width=True, hide_index=True, height=520)

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


# --- Clasificación planta vs CEDI, y despacho de plantas hacia otros CEDIs -------
# Las plantas despachan producto hacia otros CEDIs y ese movimiento no queda
# registrado como venta; para esos destinos el "vendido" del PEPS se completa con
# lo despachado según 19.1 Pedidos.xlsx (ver leer_despachos_planta).
PLANTAS = {"ALKA1", "ALKA2", "BELLAVISTA", "BODEGA EVENTUALIDAD", "LANZA", "PALMAS"}


def tipo_destino(destino: str) -> str:
    """'CEDI' si el destino es un TAT; 'PLANTA' en caso contrario."""
    return "CEDI" if "TAT" in norm(destino) else "PLANTA"


def planta_por_bodega(codigo_bodega) -> str | None:
    """Clasifica un código crudo de 'id_bodega_inventario' (Pedidos) en su planta,
    por coincidencia de texto: cubre variantes como EMALKA2, EMINKI, PAL01/02 que no
    aparecen en la lista base de códigos pero son sub-bodegas de la misma planta
    (validado contra los códigos reales de 19.1 Pedidos.xlsx). None si no es de planta."""
    c = norm(codigo_bodega)
    if "ALKA2" in c:
        return "ALKA2"
    if "ALKA" in c:
        return "ALKA1"
    if "KIGT" in c:
        return "BODEGA EVENTUALIDAD"
    if "KI" in c:
        return "LANZA"
    if "BE" in c:
        return "BELLAVISTA"
    if "PAL" in c:
        return "PALMAS"
    return None


@st.cache_data(ttl=3600)
def leer_despachos_planta(ruta: str, cache_key: float = 0.0) -> pd.DataFrame:
    """Lee 19.1 Pedidos.xlsx y calcula lo despachado desde cada planta hacia otros
    CEDIs: exige documento de entrega (id_doc_entrega no vacío = sí se despachó),
    agrupa el 'id_bodega_inventario' crudo en su planta y normaliza el item para
    cruzar con el inventario. La fecha/hora usada es 'fec_doc_entrega' (fecha del
    documento de entrega), no 'fecha_despacho' (fecha planeada), porque es la que
    tiene más datos y refleja cuándo realmente se confirmó la salida. Se conserva
    la hora completa para poder filtrar luego solo lo confirmado después de las 8am."""
    df = pd.read_excel(ruta, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    df = df[df["id_doc_entrega"].notna()].copy()
    df["destino"] = df["id_bodega_inventario"].apply(planta_por_bodega)
    df = df.dropna(subset=["destino"])
    df["item"] = norm_item(df["id_item"])
    df["cantidad"] = pd.to_numeric(df["cantidad_entregada"], errors="coerce").fillna(0.0)
    df["fecha_doc_entrega"] = pd.to_datetime(df["fec_doc_entrega"], errors="coerce")
    return df[["destino", "item", "descripcion_articulo", "cantidad", "fecha_doc_entrega"]].rename(
        columns={"descripcion_articulo": "referencia"}
    )


def lider_por_destino(destino: str) -> str:
    """Devuelve el líder responsable del destino según la clasificación de zonas."""
    d = norm(str(destino))
    if "BARRANQUILLA" in d or "CARTAGENA" in d:
        return "Juan Carlos Ortega"
    if "BOGOTA" in d or "MONTEVIDEO" in d or "SIBERIA" in d:
        return "Ariel Baez"
    if "BUCARAMANGA" in d or "CUCUTA" in d or "PASTO" in d:
        return "Johanna Olave"
    if "MEDELLIN" in d or "MONTERIA" in d:
        return "Sergio Clavijo"
    return ""


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

        # En CEDI (solo distribuye) se asume que si vendió >= todo lo de ayer, no puede
        # quedar nada viejo varado. En PLANTA ese supuesto no aplica: produce huevo
        # fresco el mismo día y lo despacha sin que pase por el inventario de la
        # mañana, así que el despacho puede superar el inventario de ayer sin que el
        # lote viejo se haya movido un solo huevo. Por eso en planta siempre se revisa
        # lote a lote, y además cuenta como "rotó" cuando el despacho no se explica con
        # los lotes de ayer (salida_por_edad): salió producto fresco en vez del viejo.
        es_planta = dest in PLANTAS
        salio_producto_fresco = es_planta and (vendido - sum(salida_por_edad.values()) > 0.5)

        if es_planta or vendido < tot_ayer - 0.5:
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
                # ¿Salió producto de algún lote MÁS NUEVO que este (edad menor), o
                # (solo en planta) producto fresco del día que nunca pasó por inventario?
                rotó_un_lote_mas_nuevo = any(
                    e_otro < e_hoy_coh and sal > 0.5
                    for e_otro, sal in salida_por_edad.items()
                )
                if rotó_un_lote_mas_nuevo or salio_producto_fresco:
                    varado += exceso            # ruptura de orden real
                    detalle_varado.append((e_hoy_coh, exceso))
                # Si NO rotó nada más nuevo (ni salió producto fresco en planta), el
                # exceso es producto a bordo / no rotación: no se cuenta como varado.

        ep_teo = edad_ponderada(teorico)
        ep_real = edad_ponderada(vh)
        ep_ayer = edad_ponderada(va)
        ruptura = (varado > 0.5) and (vendido > 0)
        hay_inflado = inflado > 0.5

        # Diagnóstico textual. Con una ventana de 'dias', un lote puede envejecer
        # legítimamente hasta +dias; solo un salto MAYOR a eso es anómalo.
        edades_real = list(vh.keys())
        node_max_real = max(edades_real) if edades_real else 0
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
    venta_peps, cat_map, _, dias_rango = preparar_ventas_peps(
        ventas, fecha_corte_obj, dias)

    # --- Despacho de plantas hacia otros CEDIs (movimiento que la venta no ve) ---
    # Para destinos de planta, el "vendido" del PEPS = SOLO lo despachado según
    # 19.1 Pedidos.xlsx (fec_doc_entrega en la fecha de corte, después de las 8am).
    # Se descarta la venta directa que ventas.xlsx pudiera mapear a estos destinos
    # (p.ej. BODEGA KIKES->LANZA), para que refleje exactamente lo registrado en
    # Pedidos. CEDI no cambia: sigue usando venta_peps (ventas.xlsx) tal cual.
    try:
        despachos = leer_despachos_planta(ARCHIVO_PEDIDOS, mtime(ARCHIVO_PEDIDOS))
        es_fecha_corte = despachos["fecha_doc_entrega"].dt.date == fecha_corte_obj
        es_despues_8am = despachos["fecha_doc_entrega"].dt.time > _dt.time(8, 0)
        desp_hoy = despachos[es_fecha_corte & es_despues_8am]
        despacho_map = desp_hoy.groupby(["destino", "item"])["cantidad"].sum().to_dict()
    except FileNotFoundError:
        despacho_map = {}
        st.warning(
            f"No se encontró **{ARCHIVO_PEDIDOS}**: las rupturas de planta se calcularán "
            "con 0 despacho (no hay venta directa de respaldo)."
        )
    venta_peps = {clave: cant for clave, cant in venta_peps.items() if clave[0] not in PLANTAS}
    venta_peps.update(despacho_map)

    res, no_map = construir_analisis(inv_ayer, inv_hoy, ventas, dias=dias,
                                     venta_peps=venta_peps, cat_map=cat_map)

    f_tipo_destino = st.radio(
        "Tipo de destino",
        ["Todos", "CEDI (TAT)", "Planta"],
        horizontal=True,
        help="CEDI = destinos TAT, rotación por venta (sin cambios). Planta = ALKA1/ALKA2/"
             "BELLAVISTA/BODEGA EVENTUALIDAD/LANZA/PALMAS, rotación por venta directa + "
             "despacho hacia otros CEDIs (19.1 Pedidos.xlsx).",
    )
    if f_tipo_destino == "CEDI (TAT)":
        res = res[res["destino"].apply(tipo_destino) == "CEDI"]
    elif f_tipo_destino == "Planta":
        res = res[res["destino"].apply(tipo_destino) == "PLANTA"]

    rupturas = res[res["ruptura"]].copy()

    if dias_rango > 1:
        st.info(
            f"📅 El promedio diario de **SUELTO** se calcula sobre **{dias_rango} días "
            "operativos** del rango (se excluyen domingos por ser días sin despacho). "
            f"Para **PET** se usa la venta exacta del día de corte ({fecha_corte_hoy})."
        )

    # Guardamos la fecha de corte por si otras secciones la reutilizan.
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

    # Pre-computar CSV de rupturas (incluye líder responsable) para el enlace oculto del encabezado.
    _csv_bytes_rup = None
    if not rupturas.empty:
        _exp = rupturas.sort_values("unds_varadas", ascending=False).copy()
        _exp["item"] = pd.to_numeric(_exp["item"], errors="coerce").astype("Int64")
        _exp_csv = pd.DataFrame({
            "Fecha corte": fecha_corte_hoy,
            "Destino": _exp["destino"],
            "Líder responsable": _exp["destino"].apply(lider_por_destino),
            "Item": _exp["item"],
            "Referencia": _exp["referencia"],
            "Inv Ayer": _exp["cant_ayer"].round(0).astype("Int64"),
            "Vendido": _exp["vendido"].round(0).astype("Int64"),
            "Inv hoy": _exp["cant_hoy"].round(0).astype("Int64"),
            "Explicación": "",
        })
        _csv_bytes_rup = _exp_csv.to_csv(index=False).encode("utf-8-sig")

    # ----- Sub-secciones por pestañas -----
    tab1, tab2 = st.tabs(
        ["🚨 Rupturas PEPS", "📋 Detalle por destino (ayer vs hoy)"]
    )

    # ===== TAB 1: RUPTURAS PEPS =====
    with tab1:
        if _csv_bytes_rup is not None:
            _b64_rup = base64.b64encode(_csv_bytes_rup).decode()
            _dl_href = (
                f'<a href="data:text/csv;charset=utf-8-sig;base64,{_b64_rup}" '
                f'download="rupturas_{fecha_corte_hoy}.csv" '
                f'style="text-decoration:none; color:{COLOR_CRITICO}; '
                f'font-size:1.25rem; vertical-align:middle; cursor:pointer;" '
                f'title="Descargar rupturas CSV">💔</a>'
            )
            st.markdown(
                f'<h3 style="margin-bottom:0.2rem; font-size:1.4rem; font-weight:700;">'
                f'Rupturas de rotación detectadas {_dl_href}</h3>',
                unsafe_allow_html=True,
            )
        else:
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
            rupturas_orden = rupturas.sort_values(["destino", "unds_varadas"], ascending=[True, False])
            opciones_rup = [
                f"{r.destino} — {int(r.item) if str(r.item).isdigit() else r.item} — {r.referencia}"
                for r in rupturas_orden.itertuples()
            ]
            sel = st.selectbox("Selecciona una ruptura para ver el detalle lote a lote", opciones_rup)
            if sel:
                idx = opciones_rup.index(sel)
                r = rupturas_orden.iloc[idx]

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
                        estado = "⚠️ Lote Omitido (debió salir)"
                    else:
                        estado = ""
                    filas_lote.append({
                        "Lote (edad ayer → hoy)": f"{e}d → {e_hoy}d",
                        "Cantidad ayer": c_ayer,
                        "Real hoy": c_real,
                        "Teórico hoy (PEPS)": c_teo,
                        "Estado del lote": estado,
                    })
                df_lotes = pd.DataFrame(filas_lote)

                def estilo_lote(row):
                    estado = str(row.get("Estado del lote", ""))
                    if "Omitido" in estado:
                        base = "background-color:#FFE08A; font-weight:700;"
                    elif "Inflado" in estado:
                        base = "background-color:#D6E9F8; font-weight:700;"
                    else:
                        base = ""
                    return [base] * len(row)

                st.markdown("**Lotes que venían de ayer**")
                styler_lotes = (
                    df_lotes.style
                    .apply(estilo_lote, axis=1)
                    .format({"Cantidad ayer": "{:,.0f}", "Teórico hoy (PEPS)": "{:,.0f}",
                             "Real hoy": "{:,.0f}"})
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
                        })
                if filas_nuevas:
                    df_nuevas = pd.DataFrame(filas_nuevas)
                    st.markdown("**Entradas nuevas del período** (producto que no existía ayer)")
                    st.dataframe(
                        df_nuevas.style.format({"Cantidad hoy": "{:,.0f}"}),
                        use_container_width=True, hide_index=True,
                    )

        # ----- Explicación de rupturas: se hace en el Sheet histórico -----
        st.divider()
        st.markdown("#### 📝 Registrar explicación de una ruptura")
        if rupturas.empty:
            st.success("No hay rupturas de rotación en este corte para explicar. 🎉")
        else:
            st.markdown(
                "Si tu ciudad aparece en la lista de rupturas, ingresa al histórico y registra "
                "el motivo en la columna **Explicación** de cada fila que te corresponde. "
                "Usa tu cuenta corporativa para editarlo."
            )
            st.link_button(
                "📄 Ir al histórico de rupturas",
                URL_SHEET_HISTORICO,
                use_container_width=True,
            )

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
                    estado = "⚠️ Lote Omitido"
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
                estado = str(row.get("Estado", ""))
                if "Lote Omitido" in estado or "Reaparecido" in estado:
                    return ["background-color:#FDEDEC;"] * len(row)          # rojo claro
                if "Inflado" in estado:
                    return ["background-color:#D6E9F8;"] * len(row)          # azul claro
                if "🆕" in estado:
                    return ["background-color:#F4F9F2; color:#4A4A4A;"] * len(row)
                return [""] * len(row)

            vista_cols = ["Destino", "Item", "Referencia", "Lote (ayer → hoy)",
                          "Cant. ayer", "Teórico hoy", "Real hoy", "Estado"]
            styler = (
                tabla[vista_cols].style
                .apply(estilo_fila, axis=1)
                .format({"Cant. ayer": "{:,.0f}", "Teórico hoy": "{:,.0f}",
                         "Real hoy": "{:,.0f}", "Item": "{:.0f}"})
            )
            st.dataframe(styler, use_container_width=True, hide_index=True, height=600)
            st.markdown(
                "⚠️ **Lote Omitido** (rojo) = lote viejo que debió salir y no salió → revisar rotación.  "
                "📦 **Inflado** (azul) = hoy hay más que ayer, imposible por envejecimiento → "
                "revisar ingreso o conteo.  🆕 Entradas nuevas en verde.  Filas en blanco rotaron bien."
            )
            st.caption(f"{len(tabla):,} lotes mostrados.")


    st.caption(
        f"Fuentes: {ARCHIVO_AYER} (inicial ayer) · {ARCHIVO_VENTAS} (ventas del período) · "
        f"{ARCHIVO_HOY} (inicial hoy) — hoja {HOJA_INV_ANALISIS}"
    )


# ===========================================================================
# MÓDULO 3 — SEGUIMIENTO DE RUPTURAS (gestión del proceso)
# ===========================================================================
def _numero_causa(causa: str) -> float:
    """Extrae el número de prefijo de una causa estandarizada (p.ej. '2. Error...' -> 2)
    para poder ordenarla numéricamente. Las causas sin número (como 'Sin explicación')
    van al final."""
    m = re.match(r"\s*(\d+)\.", str(causa))
    return int(m.group(1)) if m else float("inf")


def _colores_causas(causas: list) -> dict:
    """Asigna un color estable a cada causa, respetando su orden numérico."""
    paleta = [COLOR_PRIMARIO, COLOR_ACENTO, COLOR_ADV, COLOR_CRITICO, "#4A90D9"]
    colores = {"Sin explicación": "#C9CDD1"}
    resto = [c for c in causas if c not in colores]
    for i, c in enumerate(resto):
        colores[c] = paleta[i % len(paleta)]
    return colores


# Agrupación comercial de los CEDI/destino por regional (definida con Camila).
REGIONALES = {
    "CENTRO ORIENTE": ["BOGOTA MONTEVIDEO", "BOGOTA SIBERIA", "VILLAVICENCIO"],
    "COSTA ORIENTE": ["BARRANQUILLA", "BUCARAMANGA", "CARTAGENA", "CUCUTA",
                      "SANTA MARTA", "VALLEDUPAR"],
    "OCCIDENTE": ["CALI", "MEDELLIN", "MONTERIA", "PASTO", "POPAYAN", "SINCELEJO"],
}
ORDEN_REGIONALES = ["CENTRO ORIENTE", "COSTA ORIENTE", "OCCIDENTE", "OTROS"]
# Misma paleta que _colores_causas() (gráfico de rupturas por fecha y causa). Se
# usa CRITICO en vez de ADV para Occidente porque ADV (ámbar) y ACENTO (naranja)
# son tonos casi idénticos y, al oscurecerlos por regional, se veían iguales.
COLOR_REGIONAL = {
    "CENTRO ORIENTE": COLOR_PRIMARIO,
    "COSTA ORIENTE": COLOR_ACENTO,
    "OCCIDENTE": COLOR_CRITICO,
    "OTROS": "#4A90D9",
}


def regional_por_destino(destino: str) -> str:
    """Regional comercial a la que pertenece un CEDI/destino. Los CEDI que no calzan
    con ninguna regional (p.ej. plantas como LANZA/BELLAVISTA/PALMAS) caen en 'OTROS'."""
    d = norm(str(destino))
    for regional, claves in REGIONALES.items():
        if any(clave in d for clave in claves):
            return regional
    return "OTROS"


HOJA_INV_DESECHO = "inv"  # pestaña de inventario detallado (para detectar DESECHO)


@st.cache_data(ttl=3600)
def leer_desecho_destinos(ruta: str, cache_key: float = 0.0) -> pd.DataFrame:
    """Lee la pestaña 'inv' y filtra referencias de huevo DESECHO en todos los destinos[cite: 1]."""
    df = pd.read_excel(ruta, sheet_name=HOJA_INV_DESECHO)
    df.columns = [str(c).strip() for c in df.columns]
    out = pd.DataFrame({
        "destino": df.get("DESTINO"),
        "referencia": df.get("descripcion_articulo"),
        "cantidad": pd.to_numeric(df.get("cantidad"), errors="coerce").fillna(0.0),
    })
    out = out.dropna(subset=["destino"])
    es_desecho = out["referencia"].astype(str).str.upper().str.contains("DESECHO")
    out = out[es_desecho & (out["cantidad"] > 0)]
    return (
        out.groupby(["destino", "referencia"], as_index=False)["cantidad"]
        .sum()
        .sort_values("cantidad", ascending=False)
    )


def render_modulo_seguimiento():
    st.markdown('<p class="titulo-modulo">📈 Seguimiento de Rupturas</p>', unsafe_allow_html=True)
    st.caption(
        "Evolución histórica de las rupturas de rotación y nivel de gestión por zona."
    )
    st.divider()

    # ----- Alerta de inventario DESECHO en todos los destinos -----
    try:
        desecho = leer_desecho_destinos(ARCHIVO_HOY, mtime(ARCHIVO_HOY))
    except Exception as e:
        desecho = pd.DataFrame()
        st.warning(
            f"⚠️ No se pudo leer el inventario DESECHO desde '{ARCHIVO_HOY}': "
            f"{type(e).__name__}: {e}"
        )



    if not desecho.empty:
        total_desecho = desecho["cantidad"].sum()
        n_destinos_desecho = desecho["destino"].nunique()
        st.markdown(
            f'<div style="background-color:#FDEDEC; border-left:6px solid {COLOR_CRITICO}; '
            f'border-radius:8px; padding:14px 18px; margin-bottom:12px;">'
            f'<span style="color:{COLOR_CRITICO}; font-weight:800; font-size:1.1rem;">'
            f'🚨 Inventario DESECHO detectado en {n_destinos_desecho} destino(s): '
            f'{total_desecho:,.0f} unds.</span></div>',
            unsafe_allow_html=True,
        )
        tabla_desecho = desecho.rename(
            columns={"destino": "Destino", "referencia": "Referencia", "cantidad": "Cantidad"}
        )
        st.dataframe(
            tabla_desecho.style.format({"Cantidad": "{:,.0f}"}),
            use_container_width=True, hide_index=True,
        )
        st.divider()

    ARCHIVO_BD_RUP = resolver_archivo(
        "BD Rupturas - Hoja 1.csv",
        "BD_Rupturas_-_Hoja_1.csv",
        "BD Rupturas Hoja 1.csv",
    )

    if not os.path.exists(ARCHIVO_BD_RUP):
        st.info(
            f"Sube el archivo **BD Rupturas - Hoja 1.csv** a la raíz del repositorio "
            "para activar el seguimiento histórico."
        )
        return

    try:
        for _enc in ("utf-8-sig", "latin-1", "cp1252", "utf-8"):
            try:
                # sep=None + engine='python': pandas detecta automáticamente el separador
                # (, ; \t etc.) sin importar cómo se exportó desde Excel o Google Sheets.
                rup_raw = pd.read_csv(ARCHIVO_BD_RUP, encoding=_enc,
                                      sep=None, engine="python")
                break
            except UnicodeDecodeError:
                continue
        else:
            st.error("No pude decodificar el archivo CSV. Guárdalo como UTF-8 desde Excel.")
            return
    except Exception as e:
        st.error(f"No pude leer el archivo: {e}")
        return

    # Normaliza columnas (tolerante a mayúsculas/acentos/espacios).
    norm_map = {}
    for c in rup_raw.columns:
        cn = norm(c)
        if cn == "FECHA CORTE":
            norm_map[c] = "fecha_corte"
        elif cn == "DESTINO":
            norm_map[c] = "destino"
        elif cn in ("RESPONSABLE INV", "LIDER RESPONSABLE"):
            norm_map[c] = "responsable_inv"
        elif cn == "ITEM":
            norm_map[c] = "item"
        elif cn == "REFERENCIA":
            norm_map[c] = "referencia"
        elif cn in ("INV AYER", "INV. AYER"):
            norm_map[c] = "inv_ayer"
        elif cn == "VENDIDO":
            norm_map[c] = "vendido"
        elif cn in ("INV HOY", "INV. HOY"):
            norm_map[c] = "inv_hoy"
        elif cn in ("EXPLICACION", "EXPLICACION"):
            norm_map[c] = "explicacion"
    rup = rup_raw.rename(columns=norm_map)

    if "responsable_inv" not in rup.columns:
        rup["responsable_inv"] = ""
    if "explicacion" not in rup.columns:
        rup["explicacion"] = ""

    faltan = [c for c in ("fecha_corte", "destino") if c not in rup.columns]
    if faltan:
        st.error(
            "Al archivo le faltan columnas obligatorias: **"
            + "**, **".join(faltan)
            + f"**. Encabezados leídos: `{list(rup_raw.columns)}`"
        )
        return

    if rup.empty:
        st.info("El archivo no tiene filas de rupturas todavía.")
        return

    rup["explicacion"] = rup["explicacion"].fillna("").astype(str).str.strip()
    rup["responsable_inv"] = rup["responsable_inv"].fillna("").astype(str).str.strip()
    rup["gestionada"] = rup["explicacion"] != ""

    # --- Filtros ---
    destinos_disp = sorted(rup["destino"].dropna().unique().tolist())
    f_dest = st.multiselect("Destino", destinos_disp, placeholder="Todos")

    base = rup.copy()
    base["responsable_inv"] = base["responsable_inv"].replace("", "Sin asignar")
    if f_dest:
        base = base[base["destino"].isin(f_dest)]

    if base.empty:
        st.info("No hay registros con los filtros seleccionados.")
        return

    total = len(base)
    gestionadas = int(base["gestionada"].sum())
    pendientes = total - gestionadas
    pct = (gestionadas / total * 100) if total else 0

    st.divider()

    # --- KPI principal: donut de gestión + tarjetas ---
    col_donut, col_kpis = st.columns([1, 2], gap="large")

    with col_donut:
        fig_d = go.Figure(go.Pie(
            values=[max(pct, 0.5), max(100 - pct, 0)],
            hole=0.72,
            marker_colors=[COLOR_PRIMARIO, "#D4EDD0"],
            textinfo="none",
            hoverinfo="skip",
            direction="clockwise",
            sort=False,
            rotation=90,
        ))
        fig_d.update_layout(
            showlegend=False,
            annotations=[{
                "text": f"<b>{pct:.0f}%</b>",
                "x": 0.5, "y": 0.5,
                "font": {"size": 36, "color": COLOR_TEXTO, "family": "Nunito, sans-serif"},
                "showarrow": False,
                "xanchor": "center",
                "yanchor": "middle",
            }],
            title=dict(
                text="Nivel de Gestión",
                x=0.5, xanchor="center",
                font=dict(size=13, color="#4A4A4A", family="Nunito, sans-serif"),
            ),
            margin=dict(l=20, r=20, t=45, b=10),
            height=230,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(fig_d, use_container_width=True)

    with col_kpis:
        st.markdown("<br><br>", unsafe_allow_html=True)
        ck1, ck2, ck3 = st.columns(3)
        with ck1:
            st.markdown(
                tarjeta_kpi("Total rupturas", f"{total:,}", estado="neutral"),
                unsafe_allow_html=True,
            )
        with ck2:
            st.markdown(
                tarjeta_kpi("Gestionadas", f"{gestionadas:,}", estado="neutral"),
                unsafe_allow_html=True,
            )
        with ck3:
            st.markdown(
                tarjeta_kpi("Pendientes", f"{pendientes:,}",
                            estado="warning" if pendientes else "neutral"),
                unsafe_allow_html=True,
            )

    st.divider()

    # --- Evolución de rupturas (líneas) ---
    st.subheader("Evolución de rupturas")
    if "fecha_corte" in base.columns:
        serie = base.groupby(["fecha_corte", "gestionada"]).size().reset_index(name="n")
        pivot = serie.pivot(index="fecha_corte", columns="gestionada", values="n").fillna(0)
        pivot = pivot.rename(columns={True: "Gestionadas", False: "Pendientes"})
        for col in ("Gestionadas", "Pendientes"):
            if col not in pivot.columns:
                pivot[col] = 0
        pivot = pivot.sort_index()
        pivot["Total"] = pivot["Gestionadas"] + pivot["Pendientes"]

        fig_l = go.Figure()
        fig_l.add_trace(go.Scatter(
            x=pivot.index, y=pivot["Total"],
            name="Total", mode="lines+markers",
            line=dict(color=COLOR_ACENTO, width=3),
            marker=dict(size=7),
        ))
        fig_l.add_trace(go.Scatter(
            x=pivot.index, y=pivot["Gestionadas"],
            name="Gestionadas", mode="lines+markers",
            line=dict(color=COLOR_PRIMARIO, width=2),
            marker=dict(size=6),
        ))
        fig_l.add_trace(go.Scatter(
            x=pivot.index, y=pivot["Pendientes"],
            name="Pendientes", mode="lines+markers",
            line=dict(color=COLOR_CRITICO, width=2),
            marker=dict(size=6),
        ))
        fig_l.update_layout(
            height=360,
            margin=dict(l=10, r=10, t=10, b=10),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Nunito, sans-serif", size=13),
            xaxis=dict(title="Fecha de corte", showgrid=True, gridcolor="#EEEEEE"),
            yaxis=dict(title="N° de rupturas", showgrid=True, gridcolor="#EEEEEE"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        )
        st.plotly_chart(fig_l, use_container_width=True)
    else:
        st.info("El archivo no contiene columna de fecha de corte.")

    st.divider()

    # --- Rupturas por fecha, desglosado por causa ---
    st.subheader("📅 Rupturas por fecha y causa")
    st.caption(
        "Número de rupturas por fecha de corte, desglosado por causa. Incluye "
        "todas las rupturas del periodo filtrado; las que aún no tienen "
        "explicación registrada se agrupan como 'Sin explicación'."
    )
    fecha_causa = base.copy()
    fecha_causa["causa_mostrada"] = fecha_causa["explicacion"].replace("", "Sin explicación")
    fecha_causa["fecha_dt"] = pd.to_datetime(fecha_causa["fecha_corte"], dayfirst=True, errors="coerce")
    fecha_causa = fecha_causa.dropna(subset=["fecha_dt"])

    if fecha_causa.empty:
        st.info("No se pudieron interpretar las fechas de corte para graficar.")
    else:
        orden_fechas = sorted(fecha_causa["fecha_dt"].unique())
        etiquetas_fecha = [pd.Timestamp(f).strftime("%d-%b") for f in orden_fechas]

        causas_f = sorted(fecha_causa["causa_mostrada"].unique().tolist(), key=_numero_causa)
        color_causa_f = _colores_causas(causas_f)
        max_n = int(fecha_causa.groupby("fecha_dt").size().max())

        fig_fecha = go.Figure()
        for causa in causas_f:
            conteo = (
                fecha_causa[fecha_causa["causa_mostrada"] == causa]
                .groupby("fecha_dt").size()
                .reindex(orden_fechas, fill_value=0)
            )
            valores = conteo.values.astype(int)
            fig_fecha.add_trace(go.Bar(
                x=etiquetas_fecha, y=valores, name=causa,
                marker_color=color_causa_f[causa],
                text=[str(v) if v > 0 else "" for v in valores],
                textposition="inside",
                insidetextfont=dict(color="white", size=11),
                constraintext="none",
                hovertemplate=f"{causa}<br>%{{x}}: %{{y}} rupturas<extra></extra>",
            ))
        fig_fecha.update_layout(
            barmode="stack",
            height=380,
            margin=dict(l=10, r=10, t=10, b=10),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Nunito, sans-serif", size=13),
            xaxis=dict(title="Fecha de corte", type="category"),
            yaxis=dict(title="N° de rupturas", showgrid=True, gridcolor="#EEEEEE",
                       tick0=0, dtick=1, range=[0, max_n + 1]),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
                        traceorder="normal"),
        )
        st.plotly_chart(fig_fecha, use_container_width=True)

    st.divider()

    # --- Ranking de rupturas por causa (raíz), desglosado por regional ---
    st.subheader("🔎 Rupturas por causa y regional")
    st.caption(
        "Ranking de causas por número de rupturas (la más frecuente arriba), "
        "desglosado por regional (Centro Oriente, Costa Oriente, Occidente). "
        "Cada barra suma todos los CEDIs de esa regional. Incluye todas las rupturas "
        "del periodo filtrado; las que aún no tienen explicación registrada se "
        "agrupan como 'Sin explicación'."
    )
    causas_base = base.copy()
    causas_base["causa_mostrada"] = causas_base["explicacion"].replace("", "Sin explicación")
    causas_base["regional"] = causas_base["destino"].apply(regional_por_destino)

    # Orden por la escala de numeración de las causas (1, 2, 3... de arriba hacia abajo);
    # se pasa en reversa porque en un bar horizontal el primer elemento de la lista
    # queda abajo del todo.
    orden_causas = sorted(
        causas_base["causa_mostrada"].unique().tolist(), key=_numero_causa, reverse=True
    )
    max_n_causa = int(causas_base.groupby("causa_mostrada").size().max())

    regionales_presentes = [r for r in ORDEN_REGIONALES
                             if r in causas_base["regional"].unique()]

    fig_causas = go.Figure()
    for regional in regionales_presentes:
        conteo = (
            causas_base[causas_base["regional"] == regional]
            .groupby("causa_mostrada").size()
            .reindex(orden_causas, fill_value=0)
        )
        fig_causas.add_trace(go.Bar(
            y=orden_causas, x=conteo.values, name=regional, orientation="h",
            marker_color=COLOR_REGIONAL[regional],
            text=[str(v) if v > 0 else "" for v in conteo.values],
            textposition="inside",
            insidetextfont=dict(color="white", size=11),
            constraintext="none",
            hovertemplate=f"{regional}<br>%{{y}}: %{{x}} rupturas<extra></extra>",
        ))
    fig_causas.update_layout(
        barmode="stack",
        height=max(320, 50 * len(orden_causas)),
        margin=dict(l=10, r=10, t=10, b=10),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Nunito, sans-serif", size=13),
        xaxis=dict(title="N° de rupturas", showgrid=True, gridcolor="#EEEEEE",
                   tick0=0, dtick=1, range=[0, max_n_causa + 1]),
        yaxis=dict(title=""),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
                    traceorder="normal"),
    )
    st.plotly_chart(fig_causas, use_container_width=True)

    gest_causas = base[base["gestionada"]]
    if gest_causas.empty:
        st.info("Aún no hay rupturas gestionadas con causa registrada.")
    else:
        top_causa = gest_causas["explicacion"].value_counts().idxmax()
        kc1, _ = st.columns(2)
        with kc1:
            st.markdown(tarjeta_kpi("Causa más frecuente", top_causa, estado="warning"),
                        unsafe_allow_html=True)

    st.divider()

    # --- Tablas de detalle ---
    pend_df = base[~base["gestionada"]].copy()
    gest_df = base[base["gestionada"]].copy()

    if not pend_df.empty:
        st.subheader(f"⏳ Rupturas pendientes de gestión ({len(pend_df):,})")
        cols_p = [c for c in ["fecha_corte", "destino", "item",
                               "referencia", "inv_ayer", "vendido", "inv_hoy"]
                  if c in pend_df.columns]
        vista_p = pend_df[cols_p].rename(columns={
            "fecha_corte": "Fecha corte",
            "destino": "Destino", "item": "Item", "referencia": "Referencia",
            "inv_ayer": "Inv Ayer", "vendido": "Vendido", "inv_hoy": "Inv hoy",
        })
        st.dataframe(vista_p, use_container_width=True, hide_index=True, height=280)

    if not gest_df.empty:
        st.subheader(f"✅ Rupturas gestionadas ({len(gest_df):,})")
        cols_g = [c for c in ["fecha_corte", "destino", "item",
                               "referencia", "responsable_inv", "explicacion"] if c in gest_df.columns]
        vista_g = gest_df[cols_g].rename(columns={
            "fecha_corte": "Fecha corte",
            "destino": "Destino", "item": "Item", "referencia": "Referencia",
            "responsable_inv": "Responsable Inv", "explicacion": "Explicación",
        })
        st.dataframe(vista_g, use_container_width=True, hide_index=True, height=280)


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
        f"• **{ARCHIVO_VENTAS}**\n\n"
        f"• **{ARCHIVO_PEDIDOS}**"
    )

if modulo == "Inventario de Edades":
    render_modulo_edades()
elif modulo == "Análisis de Rotación PEPS":
    render_modulo_rotacion()
else:
    render_modulo_seguimiento()