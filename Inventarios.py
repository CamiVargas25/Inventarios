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

# --- Paleta de marca (identidad Huevos Kikes) ------------------------------
COLOR_PRIMARIO = "#3DAE2B"   # verde Kikes
COLOR_ACENTO = "#F7941D"     # naranja (yema del logo)

# --- Tintas y superficies (capa neutra sobre la que respira la marca) ------
# Los grises llevan una pizca de verde para que el neutro no pelee con la marca.
TINTA = "#16211B"            # texto principal
TINTA_2 = "#5A635D"          # texto secundario
TINTA_3 = "#8A928C"          # texto tenue (ejes, notas al pie)
SUPERFICIE = "#FFFFFF"       # superficie de tarjeta / gráfico
PLANO = "#F5F7F4"            # fondo de página (deja flotar las tarjetas)
BORDE = "#E3E8E1"            # filete de 1px
REJILLA = "#EDF1EB"          # línea de rejilla de los gráficos

COLOR_TEXTO = TINTA          # se conserva el nombre por compatibilidad

# --- Escala de severidad (semáforo de edades) ------------------------------
# Rampa ordinal de "calor semántico": el orden lo da la posición (eje de días)
# y el número siempre está visible, así que el color solo refuerza.
# Validada con scripts/validate_palette.js del skill dataviz sobre superficie
# blanca: separación CVD 11.3 y de visión normal 16.7 entre pares adyacentes
# (ambas por encima del umbral). El amarillo queda fuera de la banda de
# luminosidad y por debajo de 3:1 de contraste —inevitable en un amarillo—,
# cubierto por la regla de relieve: la cifra de edad siempre se ve y hay
# leyenda de convención.
SEV_OPTIMO = "#3DAE2B"       # 1–4 días
SEV_ALERTA = "#F2C230"       # 5 días
SEV_PREOCUPANTE = "#E2611A"  # 6–9 días
SEV_CRITICO = "#B3141F"      # 10+ días

COLOR_CRITICO = SEV_CRITICO  # rojo único en toda la app (UI + datos)
COLOR_ADV = SEV_PREOCUPANTE  # ámbar de advertencia

# Tintes suaves para fondos de banner y celdas (texto oscuro encima).
TINTE_CRITICO = "#FCEEEF"
TINTE_ADV = "#FDF2E6"
TINTE_OK = "#EFF8EC"
TINTE_INFO = "#EEF3F8"

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
ARCHIVO_INVENTARIOS = resolver_archivo("Inventarios.xlsx", "inventarios.xlsx")
ARCHIVO_KARDEX = resolver_archivo("kardex.xlsx", "Kardex.xlsx", "KARDEX.xlsx")
HOJA_EDADES_DASH = "INV. EDADES"   # hoja para el módulo 1 (dashboard de edades)
HOJA_INV_ANALISIS = "INV. EDADES"  # hoja para el módulo 2 (análisis de rotación)
HOJA_TRANSITO = "Inventarios"      # hoja de Inventarios.xlsx con el inventario en tránsito
HOJA_KARDEX = "Kardex"             # hoja de kardex.xlsx con los movimientos E/S de las VT
UMBRAL_DIAS_TRANSITO = 3           # días en tránsito a partir de los cuales se considera varado


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
# Inter para interfaz y cifras (lectura neutra, ejecutiva); Nunito se reserva
# para el logotipo del panel, que es donde la marca debe sonar.
st.markdown(
    f"""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Nunito:wght@800;900&display=swap');

        html, body, [class*="css"], .stMarkdown, .stDataFrame {{
            font-family: 'Inter', system-ui, -apple-system, 'Segoe UI', sans-serif;
            -webkit-font-smoothing: antialiased;
        }}

        /* ----- Plano de página: las tarjetas blancas flotan sobre él ----- */
        .stApp {{ background-color: {PLANO}; }}
        .block-container {{ padding-top: 2.4rem; padding-bottom: 3.5rem; max-width: 1500px; }}

        h1, h2, h3, h4, h5, h6 {{ color: {TINTA}; letter-spacing: -0.015em; }}

        /* ----- Encabezado de módulo (kicker + título + bajada) ----- */
        .enc-wrap {{ margin: 0 0 1.6rem 0; }}
        .enc-kicker {{
            display: inline-flex; align-items: center; gap: 7px;
            font-size: 0.7rem; font-weight: 700; letter-spacing: 0.13em;
            text-transform: uppercase; color: {COLOR_PRIMARIO};
            margin-bottom: 0.5rem;
        }}
        .enc-kicker::before {{
            content: ""; width: 16px; height: 3px; border-radius: 2px;
            background: {COLOR_ACENTO};
        }}
        .enc-titulo {{
            color: {TINTA}; font-size: 2.15rem; font-weight: 800;
            letter-spacing: -0.028em; line-height: 1.12; margin: 0;
        }}
        .enc-sub {{
            color: {TINTA_2}; font-size: 0.95rem; font-weight: 400;
            margin: 0.4rem 0 0 0; max-width: 78ch; line-height: 1.5;
        }}

        /* ----- Título de sección ----- */
        .sec-titulo {{
            display: flex; align-items: center; gap: 9px;
            font-size: 1.12rem; font-weight: 700; color: {TINTA};
            letter-spacing: -0.012em; margin: 0 0 0.15rem 0;
        }}
        .sec-titulo::before {{
            content: ""; width: 4px; height: 17px; border-radius: 2px;
            background: {COLOR_PRIMARIO}; flex: none;
        }}
        .sec-sub {{
            color: {TINTA_2}; font-size: 0.875rem; line-height: 1.5;
            margin: 0 0 0.85rem 13px; max-width: 92ch;
        }}

        /* ----- Tarjetas KPI (stat tiles) ----- */
        .kpi-card {{
            position: relative; background-color: {SUPERFICIE};
            border: 1px solid {BORDE}; border-radius: 14px;
            padding: 18px 18px 16px 18px; height: 100%;
            box-shadow: 0 1px 2px rgba(22,33,27,0.04);
            overflow: hidden;
        }}
        /* El acento es una barra superior fina, no un bloque lateral grueso. */
        .kpi-card::before {{
            content: ""; position: absolute; top: 0; left: 0; right: 0;
            height: 3px; background: {TINTA_3};
        }}
        .kpi-card .kpi-label {{
            color: {TINTA_2}; font-size: 0.735rem; font-weight: 600;
            letter-spacing: 0.075em; text-transform: uppercase;
            margin: 0 0 9px 0; line-height: 1.35;
        }}
        .kpi-card .kpi-value {{
            color: {TINTA}; font-size: 2rem; font-weight: 700;
            letter-spacing: -0.03em; margin: 0; line-height: 1.05;
        }}
        .kpi-neutral::before {{ background: {COLOR_PRIMARIO}; }}
        .kpi-warning::before {{ background: {COLOR_ADV}; }}
        .kpi-critical::before {{ background: {COLOR_CRITICO}; }}

        /* Tarjeta protagonista: una sola por vista. */
        .kpi-reina {{
            background: linear-gradient(160deg, {TINTE_OK} 0%, {SUPERFICIE} 62%);
            border-color: {COLOR_PRIMARIO}2E;
        }}
        .kpi-reina::before {{ background: {COLOR_PRIMARIO}; height: 3px; }}
        .kpi-reina .kpi-label {{ color: {COLOR_PRIMARIO}; }}
        .kpi-reina .kpi-value {{ font-size: 2.5rem; }}

        /* ----- Banners de alerta ----- */
        .banner {{
            display: flex; gap: 12px; align-items: flex-start;
            border: 1px solid {BORDE}; border-left-width: 4px;
            border-radius: 11px; padding: 13px 16px; margin-bottom: 13px;
            font-size: 0.93rem; line-height: 1.5; color: {TINTA};
        }}
        .banner .banner-ico {{ font-size: 1.05rem; line-height: 1.35; flex: none; }}
        .banner b {{ font-weight: 700; }}
        .banner-critico {{ background: {TINTE_CRITICO}; border-left-color: {COLOR_CRITICO}; }}
        .banner-adv {{ background: {TINTE_ADV}; border-left-color: {COLOR_ADV}; }}
        .banner-ok {{ background: {TINTE_OK}; border-left-color: {COLOR_PRIMARIO}; }}
        .banner-info {{ background: {TINTE_INFO}; border-left-color: #4A7FB5; }}

        /* ----- Barra lateral ----- */
        section[data-testid="stSidebar"] {{
            background-color: {SUPERFICIE};
            border-right: 1px solid {BORDE};
        }}
        section[data-testid="stSidebar"] .block-container {{ padding-top: 1.6rem; }}
        .marca {{
            display: flex; align-items: center; gap: 10px; margin-bottom: 2px;
        }}
        .marca-logo {{
            width: 36px; height: 36px; border-radius: 10px; flex: none;
            background: linear-gradient(140deg, {COLOR_PRIMARIO} 0%, #2E9020 100%);
            display: flex; align-items: center; justify-content: center;
            font-size: 1.15rem;
        }}
        .marca-txt {{ line-height: 1.15; }}
        .marca-nombre {{
            font-family: 'Nunito', sans-serif; font-size: 1.05rem;
            font-weight: 900; color: {TINTA}; letter-spacing: -0.02em;
        }}
        .marca-sub {{
            font-size: 0.69rem; font-weight: 600; letter-spacing: 0.1em;
            text-transform: uppercase; color: {TINTA_3};
        }}
        .nav-rotulo {{
            font-size: 0.68rem; font-weight: 700; letter-spacing: 0.12em;
            text-transform: uppercase; color: {TINTA_3};
            margin: 0.3rem 0 0.5rem 0;
        }}

        /* Navegación: el radio se lee como lista de secciones, no como formulario. */
        section[data-testid="stSidebar"] div[role="radiogroup"] {{ gap: 3px; }}
        section[data-testid="stSidebar"] div[role="radiogroup"] > label {{
            padding: 9px 12px; border-radius: 9px; width: 100%;
            transition: background-color .13s ease;
        }}
        section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover {{
            background-color: {PLANO};
        }}
        section[data-testid="stSidebar"] div[role="radiogroup"] > label p {{
            font-size: 0.9rem; font-weight: 600; color: {TINTA_2};
        }}

        /* ----- Controles ----- */
        .stButton > button {{
            border-radius: 9px; border: 1px solid {BORDE};
            font-weight: 600; font-size: 0.88rem; color: {TINTA};
            background: {SUPERFICIE}; transition: all .13s ease;
        }}
        .stButton > button:hover {{
            border-color: {COLOR_PRIMARIO}; color: {COLOR_PRIMARIO};
            background: {TINTE_OK};
        }}
        .stTabs [data-baseweb="tab-list"] {{ gap: 4px; border-bottom: 1px solid {BORDE}; }}
        .stTabs [data-baseweb="tab"] {{
            font-weight: 600; font-size: 0.92rem; color: {TINTA_2};
        }}
        .stTabs [aria-selected="true"] {{ color: {COLOR_PRIMARIO}; }}

        /* ----- Tablas ----- */
        .stDataFrame {{
            border: 1px solid {BORDE}; border-radius: 12px; overflow: hidden;
        }}
        .stDataFrame [data-testid="stTable"] {{ font-variant-numeric: tabular-nums; }}

        /* ----- Separadores y notas ----- */
        hr {{ border-color: {BORDE}; margin: 1.9rem 0 1.5rem 0; }}
        [data-testid="stCaptionContainer"] p {{
            color: {TINTA_2}; font-size: 0.855rem; line-height: 1.5;
        }}
        div[data-testid="stMetricValue"] {{ font-weight: 700; letter-spacing: -0.02em; }}
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


def encabezado_modulo(kicker: str, titulo: str, subtitulo: str = "") -> str:
    """Encabezado de módulo: antetítulo de marca + título + bajada explicativa."""
    sub = f'<p class="enc-sub">{subtitulo}</p>' if subtitulo else ""
    return (
        f'<div class="enc-wrap">'
        f'<div class="enc-kicker">{kicker}</div>'
        f'<h1 class="enc-titulo">{titulo}</h1>{sub}</div>'
    )


def titulo_seccion(titulo: str, subtitulo: str = "") -> str:
    """Título de sección con filete verde y bajada opcional."""
    sub = f'<p class="sec-sub">{subtitulo}</p>' if subtitulo else ""
    return f'<div class="sec-titulo">{titulo}</div>{sub}'


def banner(tono: str, icono: str, texto: str) -> str:
    """Banner de alerta. tono: 'critico' | 'adv' | 'ok' | 'info'."""
    return (
        f'<div class="banner banner-{tono}">'
        f'<span class="banner-ico">{icono}</span><div>{texto}</div></div>'
    )


def tinta_sobre(fondo_hex: str) -> str:
    """Blanco o tinta oscura según la luminancia del relleno, para que una
    etiqueta puesta DENTRO de una marca de color siempre tenga contraste."""
    h = str(fondo_hex).lstrip("#")
    if len(h) != 6:
        return "#FFFFFF"
    try:
        r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    except ValueError:
        return "#FFFFFF"

    def lin(c):
        return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4

    lum = 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b)
    return "#FFFFFF" if lum < 0.42 else "#16211B"


def color_severidad_edad(edad) -> str:
    """Color del semáforo de edades. Fuente única de verdad para el gráfico y la
    tabla, para que ninguno de los dos se desvíe de la convención publicada:
    1–4 óptimo · 5 alerta · 6–9 preocupante · 10+ crítico."""
    try:
        v = float(edad)
    except (TypeError, ValueError):
        return TINTA_3
    if v >= 10:
        return SEV_CRITICO
    if v >= 6:
        return SEV_PREOCUPANTE
    if v >= 5:
        return SEV_ALERTA
    return SEV_OPTIMO


# --- Tema compartido de gráficos -------------------------------------------
FUENTE_GRAFICO = "Inter, system-ui, -apple-system, 'Segoe UI', sans-serif"


def estilo_grafico(fig, alto=340, mostrar_leyenda=False, margen=None):
    """Aplica el cromado común a una figura Plotly: ejes y rejilla recesivos
    (filete sólido de 1px, nunca punteado), tipografía de interfaz y fondo
    transparente para que herede la superficie de la tarjeta."""
    fig.update_layout(
        height=alto,
        margin=margen or dict(l=8, r=18, t=8, b=8),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family=FUENTE_GRAFICO, size=12.5, color=TINTA_2),
        showlegend=mostrar_leyenda,
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
            font=dict(size=12, color=TINTA_2), bgcolor="rgba(0,0,0,0)",
        ),
        hoverlabel=dict(
            bgcolor=SUPERFICIE, bordercolor=BORDE,
            font=dict(family=FUENTE_GRAFICO, size=12.5, color=TINTA),
        ),
    )
    eje = dict(
        showgrid=False, zeroline=False,
        linecolor=BORDE, linewidth=1,
        tickfont=dict(size=11.5, color=TINTA_3),
        title_font=dict(size=12, color=TINTA_3),
    )
    fig.update_xaxes(**eje)
    fig.update_yaxes(**eje, gridcolor=REJILLA)
    fig.update_yaxes(showgrid=True, griddash="solid")
    return fig


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


@st.cache_data(ttl=3600)
def cargar_transito_varado_kardex(
    ruta: str,
    ruta_snapshot: str,
    umbral_dias: int,
    cache_key: float = 0.0,
    cache_key_snapshot: float = 0.0,
) -> tuple:
    """Reconstruye, con lógica PEPS, los lotes de huevo (línea HU, por descripción
    'HUEVO...') que siguen sin salir de cada bodega de tránsito ('VT...'), a partir
    del historial de movimientos E/S de 'kardex.xlsx'.

    Por cada (bodega, código_artículo) se procesan los movimientos en orden
    cronológico real (columna 'datetime'): cada entrada (E) abre un lote nuevo con
    la fecha de movimiento (columna 'fecha'); cada salida (S) consume primero el
    lote más antiguo (FIFO/PEPS). Una salida que no encuentra lote que la respalde
    (p. ej. la primera transacción de la serie es una salida, señal de que hubo una
    entrada anterior al rango del kardex) se descarta sin efecto: no se resta de
    ningún lote ni genera saldo negativo. Los lotes que quedan con saldo > 0 al
    final son el inventario que sigue varado en tránsito, con su fecha de entrada
    original.

    Cada saldo remanente se cruza además contra el snapshot actual
    ('Inventarios.xlsx'): si esa combinación bodega+artículo no aparece ahí con
    cantidad > 0, se marca como 'no confirmada' en vez de descartarse, ya que el
    kardex por sí solo (con su ventana de historia limitada) no basta para
    garantizar que el lote sigue físicamente ahí.

    Devuelve (dataframe, n_salidas_huerfanas, snapshot_disponible).
    """
    df = pd.read_excel(ruta, sheet_name=HOJA_KARDEX)
    df.columns = [str(c).strip() for c in df.columns]
    base = pd.DataFrame({
        "bodega": df.get("descripcion"),
        "codigo_bodega": df.get("bodega"),
        "codigo_articulo": df.get("codigo_articulo").astype(str),
        "producto": df.get("descripcion_articulo"),
        "fecha": pd.to_datetime(df.get("fecha"), errors="coerce"),
        "orden": pd.to_datetime(df.get("datetime"), errors="coerce"),
        "cantidad": pd.to_numeric(df.get("cantidad"), errors="coerce").fillna(0.0),
        "tipo_mov": df.get("tipo_mov").astype(str).str.upper().str.strip(),
    })
    base = base.dropna(subset=["fecha", "orden"])
    base = base[base["producto"].astype(str).str.upper().str.startswith("HUEVO")]
    base = base.sort_values(["codigo_bodega", "codigo_articulo", "orden"])

    filas = []
    n_huerfanas = 0
    for (cb, ca), grupo in base.groupby(["codigo_bodega", "codigo_articulo"], sort=False):
        lotes = []  # cada lote: [fecha_entrada, cantidad_restante]
        for _, mov in grupo.iterrows():
            if mov["tipo_mov"] == "E":
                lotes.append([mov["fecha"], mov["cantidad"]])
            elif mov["tipo_mov"] == "S":
                restante = mov["cantidad"]
                while restante > 0.001 and lotes:
                    lote = lotes[0]
                    consumido = min(lote[1], restante)
                    lote[1] -= consumido
                    restante -= consumido
                    if lote[1] <= 0.001:
                        lotes.pop(0)
                if restante > 0.001:
                    n_huerfanas += 1  # salida sin lote que la respalde: se ignora
        bodega_nombre = grupo["bodega"].iloc[-1]
        producto_nombre = grupo["producto"].iloc[-1]
        for fecha_lote, cantidad_lote in lotes:
            if cantidad_lote > 0.001:
                filas.append({
                    "bodega": bodega_nombre,
                    "codigo_bodega": cb,
                    "producto": producto_nombre,
                    "codigo_articulo": ca,
                    "cantidad": cantidad_lote,
                    "fecha": fecha_lote,
                })

    out = pd.DataFrame(filas)
    if out.empty:
        return out, n_huerfanas, True

    hoy = pd.Timestamp(_dt.date.today())
    out["dias_varado"] = (hoy - out["fecha"].dt.normalize()).dt.days
    out = out[out["dias_varado"] >= umbral_dias].sort_values("dias_varado", ascending=False).reset_index(drop=True)

    try:
        snap = pd.read_excel(ruta_snapshot, sheet_name=HOJA_TRANSITO)
        snap.columns = [str(c).strip() for c in snap.columns]
        snap_cant = pd.to_numeric(snap.get("cantidad"), errors="coerce").fillna(0.0)
        snap_key = pd.DataFrame({
            "codigo_bodega": snap.get("codigo_bodega"),
            "codigo_articulo": snap.get("codigo_articulo").astype(str),
            "cantidad": snap_cant,
        })
        vigentes = snap_key.groupby(["codigo_bodega", "codigo_articulo"])["cantidad"].sum()
        confirmados = set(vigentes[vigentes > 0].index)
        snapshot_ok = True
    except Exception:
        confirmados = set()
        snapshot_ok = False

    out["confirmado"] = [
        (cb, ca) in confirmados for cb, ca in zip(out["codigo_bodega"], out["codigo_articulo"])
    ]
    return out, n_huerfanas, snapshot_ok


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

    st.markdown(
        encabezado_modulo(
            "Panel de inventarios",
            "Inventario de Edades",
            "Estado actual del inventario por edad, planta y CEDI, con alertas de "
            "desecho, tránsito varado y riesgo de vida útil.",
        ),
        unsafe_allow_html=True,
    )

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
    st.markdown(
        titulo_seccion(
            "🗑️ Alerta de inventario desecho",
            "Referencias cuyo nombre contiene la palabra <b>desecho</b>, agrupadas por planta/CEDI.",
        ),
        unsafe_allow_html=True,
    )

    try:
        desecho = cargar_desecho_por_planta(ARCHIVO_HOY, mtime(ARCHIVO_HOY))
    except Exception as e:
        desecho = pd.DataFrame()
        st.markdown(
            banner("adv", "⚠️",
                   f"No se pudo leer el inventario DESECHO desde '{ARCHIVO_HOY}': "
                   f"{type(e).__name__}: {e}"),
            unsafe_allow_html=True,
        )

    if desecho.empty:
        st.markdown(
            banner("ok", "✅", "No se detectó inventario DESECHO."),
            unsafe_allow_html=True,
        )
    else:
        total_desecho = desecho["cantidad"].sum()
        n_plantas_desecho = desecho["planta"].nunique()
        st.markdown(
            banner("critico", "🚨",
                   f"<b>Inventario DESECHO detectado</b> en {n_plantas_desecho} planta(s)/CEDI: "
                   f"<b>{total_desecho:,.0f} und.</b>"),
            unsafe_allow_html=True,
        )
        tabla_desecho = desecho.rename(columns={"planta": "Planta/CEDI", "cantidad": "Cantidad"})
        st.dataframe(
            tabla_desecho.style.format({"Cantidad": "{:,.0f}"}).bar(
                subset=["Cantidad"], color=TINTE_CRITICO, align="left", vmin=0
            ),
            use_container_width=True,
            hide_index=True,
        )

    st.divider()
    st.markdown(
        titulo_seccion(
            "🚚 Alerta de inventario varado en tránsito",
            f"Reconstrucción PEPS de huevo (por descripción 'HUEVO...') a partir de los "
            f"movimientos E/S de <b>{ARCHIVO_KARDEX}</b>. Por cada bodega+artículo se consume primero "
            f"el lote más antiguo; lo que queda sin salida y lleva <b>{UMBRAL_DIAS_TRANSITO} días o más</b> "
            f"desde su entrada se marca como varado. Las salidas huérfanas (sin entrada previa en el "
            f"kardex) se ignoran sin afectar el resto de la serie. Cada saldo remanente se cruza además "
            f"contra <b>{ARCHIVO_INVENTARIOS}</b>: si esa combinación bodega+artículo no aparece ahí con "
            f"cantidad &gt; 0, la fila se muestra igual pero marcada como <b>no confirmada</b>.",
        ),
        unsafe_allow_html=True,
    )

    transito_kdx = pd.DataFrame()
    n_huerfanas = 0
    cargo_ok_kdx = False
    try:
        transito_kdx, n_huerfanas, snapshot_ok_kdx = cargar_transito_varado_kardex(
            ARCHIVO_KARDEX, ARCHIVO_INVENTARIOS, UMBRAL_DIAS_TRANSITO,
            mtime(ARCHIVO_KARDEX), mtime(ARCHIVO_INVENTARIOS),
        )
        cargo_ok_kdx = True
    except FileNotFoundError:
        st.info(f"Sube **{ARCHIVO_KARDEX}** para activar esta segunda versión (PEPS) de la alerta.")
    except Exception as e:
        st.warning(
            f"⚠️ No se pudo leer el kardex desde '{ARCHIVO_KARDEX}': {type(e).__name__}: {e}"
        )

    if cargo_ok_kdx:
        if n_huerfanas:
            st.caption(f"🧹 {n_huerfanas} salida(s) huérfana(s) detectada(s) e ignorada(s) en el kardex.")
        if not snapshot_ok_kdx:
            st.caption(
                f"⚠️ No se pudo cruzar contra **{ARCHIVO_INVENTARIOS}**; ninguna fila quedó confirmada."
            )

    if transito_kdx.empty:
        if cargo_ok_kdx:
            st.markdown(
                banner("ok", "✅", "No hay lotes de huevo varados en tránsito según el kardex."),
                unsafe_allow_html=True,
            )
    else:
        n_lotes = len(transito_kdx)
        n_bodegas_kdx = transito_kdx["bodega"].nunique()
        max_dias_kdx = int(transito_kdx["dias_varado"].max())
        n_no_confirmados = int((~transito_kdx["confirmado"]).sum())
        st.markdown(
            banner("critico", "🚨",
                   f"<b>{n_lotes} lote(s) varados</b> en {n_bodegas_kdx} bodega(s) de tránsito "
                   f"(hasta <b>{max_dias_kdx} días</b> sin resolver) — "
                   f"<b>{n_no_confirmados}</b> sin confirmar contra el snapshot actual."),
            unsafe_allow_html=True,
        )
        tabla_kdx = transito_kdx.copy()
        tabla_kdx["confirmado"] = tabla_kdx["confirmado"].map({True: "✅ Sí", False: "⚠️ No confirmado"})
        tabla_kdx = tabla_kdx.rename(columns={
            "bodega": "Bodega tránsito",
            "codigo_bodega": "Código",
            "producto": "Producto",
            "codigo_articulo": "Código artículo",
            "cantidad": "Cantidad",
            "fecha": "Fecha entrada (lote)",
            "dias_varado": "Días varado",
            "confirmado": "Confirmado en snapshot",
        })
        st.dataframe(
            tabla_kdx.style.format({
                "Cantidad": "{:,.0f}",
                "Fecha entrada (lote)": lambda d: d.strftime("%Y-%m-%d %H:%M"),
                "Días varado": "{:,.0f}",
            }).bar(subset=["Días varado"], color=TINTE_CRITICO, align="left", vmin=0),
            use_container_width=True,
            hide_index=True,
        )

    st.divider()
    st.markdown(
        titulo_seccion(
            "Distribución del inventario por edad",
            "Unidades en stock por día de edad. El color refuerza el semáforo de "
            "vida útil; el eje y la cifra sobre cada barra llevan el dato exacto.",
        ),
        unsafe_allow_html=True,
    )

    dist = dff.dropna(subset=["edad"]).copy()
    dist["edad_int"] = dist["edad"].astype(int)
    dist["bucket"] = dist["edad_int"].apply(lambda x: "10+" if x >= 10 else str(x))
    orden_buckets = [str(i) for i in range(1, 10)] + ["10+"]
    serie = dist.groupby("bucket")["cantidad"].sum().reindex(orden_buckets, fill_value=0)

    colores = [color_severidad_edad(10 if b == "10+" else int(b)) for b in serie.index]
    fig = go.Figure(
        go.Bar(
            x=list(serie.index),
            y=serie.values,
            marker=dict(color=colores, cornerradius=4),
            text=[f"{v:,.0f}" if v > 0 else "" for v in serie.values],
            textposition="outside",
            textfont=dict(size=11.5, color=TINTA_2),
            cliponaxis=False,
            hovertemplate="Edad: %{x} días<br>Cantidad: %{y:,.0f} und<extra></extra>",
        )
    )
    fig.update_layout(bargap=0.42)
    estilo_grafico(fig, alto=350, margen=dict(l=8, r=18, t=22, b=8))
    fig.update_xaxes(title="Días de edad del producto", tickmode="linear")
    fig.update_yaxes(title="Unidades", separatethousands=True)
    st.plotly_chart(fig, use_container_width=True)

    st.divider()
    st.markdown(titulo_seccion("Detalle de inventario"), unsafe_allow_html=True)

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

    # Celda del semáforo: tinte suave de fondo + tinta oscura de la misma familia,
    # para que el número siempre se lea (el color acompaña, no sustituye al dato).
    TINTA_SEV = {
        SEV_OPTIMO: ("#E8F5E4", "#1F6B15"),
        SEV_ALERTA: ("#FDF3D6", "#7A5A00"),
        SEV_PREOCUPANTE: ("#FCEBDF", "#8A3708"),
        SEV_CRITICO: ("#FAE4E6", "#8A0F17"),
    }

    def color_edad(val):
        sev = color_severidad_edad(val)
        if sev not in TINTA_SEV:
            return ""
        fondo, tinta = TINTA_SEV[sev]
        return f"background-color: {fondo}; color: {tinta}; font-weight: 700;"

    if "Item" in tabla.columns:
        tabla["Item"] = pd.to_numeric(tabla["Item"], errors="coerce").astype("Int64")

    styler = (
        tabla.style
        .map(color_edad, subset=["Edad"])
        .bar(subset=["Suma de Cantidad"], color=TINTE_OK, align="left", height=70, vmin=0)
        .format({"Suma de Cantidad": "{:,.0f}", "Edad": "{:.0f}", "Item": "{:.0f}"})
    )
    st.dataframe(styler, use_container_width=True, hide_index=True, height=600)

    st.markdown(
        f"""
        <div style="display:flex; flex-wrap:wrap; gap:18px; align-items:center;
                    font-size:0.84rem; color:{TINTA_2}; margin-top:10px;">
          <span style="font-weight:600; color:{TINTA};">Convención de edades</span>
          <span><span style="display:inline-block;width:10px;height:10px;border-radius:3px;
                background:{SEV_OPTIMO};margin-right:6px;"></span>1–4 días · óptimo</span>
          <span><span style="display:inline-block;width:10px;height:10px;border-radius:3px;
                background:{SEV_ALERTA};margin-right:6px;"></span>5 días · alerta</span>
          <span><span style="display:inline-block;width:10px;height:10px;border-radius:3px;
                background:{SEV_PREOCUPANTE};margin-right:6px;"></span>6–9 días · preocupante</span>
          <span><span style="display:inline-block;width:10px;height:10px;border-radius:3px;
                background:{SEV_CRITICO};margin-right:6px;"></span>10+ días · crítico</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.divider()
    st.markdown(
        titulo_seccion(
            "⏳ Riesgo de vida útil",
            "Proyección a futuro: con el ritmo de venta diaria y la edad actual de cada lote, "
            "se estima la edad que tendría al venderse bajo PEPS (más viejo primero). "
            "Se alerta si algún lote proyecta superar <b>5 días</b>.",
        ),
        unsafe_allow_html=True,
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
        # Tolerancia solo en planta: el modelo ahí es más ruidoso (producción del mismo
        # día, ajustes no vistos), así que se ignora un exceso menor al 10% de lo que
        # había ayer en ese lote específico. CEDI no cambia (tolerancia = 0).
        UMBRAL_TOLERANCIA_PLANTA = 0.10

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
                tolerancia = c_ayer_coh * UMBRAL_TOLERANCIA_PLANTA if es_planta else 0.0
                if exceso <= max(0.5, tolerancia):
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
    st.markdown(
        encabezado_modulo(
            "Auditoría de rotación",
            "Análisis de Rotación PEPS",
            "Compara el inventario inicial de ayer, el consumo del período y el "
            "inventario de hoy para detectar lotes viejos que no rotaron.",
        ),
        unsafe_allow_html=True,
    )
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
                titulo_seccion(f"Rupturas de rotación detectadas {_dl_href}"),
                unsafe_allow_html=True,
            )
        else:
            st.markdown(titulo_seccion("Rupturas de rotación detectadas"),
                        unsafe_allow_html=True)
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
        st.markdown(
            titulo_seccion(
                "Detalle por destino — seguimiento por lote",
                f"Cada fila sigue un <b>lote</b> desde su edad de ayer hasta hoy (envejece "
                f"+{dias} día(s)). Los lotes que venían de ayer muestran su evolución; las "
                "entradas nuevas del período aparecen marcadas aparte. La reconstrucción por "
                "lote es un modelo PEPS.",
            ),
            unsafe_allow_html=True,
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


def _es_motivo_otro(causa: str) -> bool:
    """True si la causa corresponde al motivo '7. Otro' (tolerante a numeración,
    acentos y mayúsculas)."""
    return norm(re.sub(r"^\s*\d+\.\s*", "", str(causa))) == "OTRO"


def _colores_causas(causas: list) -> dict:
    """Asigna un color estable a cada causa, respetando su orden numérico.

    Paleta categórica de 8 tonos (azul/naranja/aqua/amarillo/magenta/verde/violeta/rojo),
    validada con el validador de paleta del skill dataviz: separación CVD y de
    visión normal por encima del umbral entre pares adyacentes, para que causas
    consecutivas nunca se vean parecidas (a diferencia de la paleta anterior, donde
    naranja y ámbar eran casi indistinguibles). `causas` debe venir ordenada de forma
    estable (p.ej. por _numero_causa ascendente) para que la misma causa reciba
    siempre el mismo color en todos los gráficos.
    """
    paleta = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100",
              "#e87ba4", "#008300", "#4a3aa7", "#e34948"]
    colores = {"Sin explicación": "#C9CDD1"}
    resto = [c for c in causas if c not in colores]
    for i, c in enumerate(resto):
        colores[c] = paleta[i % len(paleta)]
    return colores


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
    st.markdown(
        encabezado_modulo(
            "Gestión del proceso",
            "Seguimiento de Rupturas",
            "Evolución, causas y estado de gestión de las rupturas registradas por "
            "los líderes de zona.",
        ),
        unsafe_allow_html=True,
    )
    st.caption(
        "Evolución histórica de las rupturas de rotación y nivel de gestión por zona."
    )
    st.divider()

    # ----- Alerta de inventario DESECHO en todos los destinos -----
    try:
        desecho = leer_desecho_destinos(ARCHIVO_HOY, mtime(ARCHIVO_HOY))
    except Exception as e:
        desecho = pd.DataFrame()
        st.markdown(
            banner("adv", "⚠️",
                   f"No se pudo leer el inventario DESECHO desde '{ARCHIVO_HOY}': "
                   f"{type(e).__name__}: {e}"),
            unsafe_allow_html=True,
        )

    if not desecho.empty:
        total_desecho = desecho["cantidad"].sum()
        n_destinos_desecho = desecho["destino"].nunique()
        st.markdown(
            banner("critico", "🚨",
                   f"<b>Inventario DESECHO detectado</b> en {n_destinos_desecho} destino(s): "
                   f"<b>{total_desecho:,.0f} unds.</b>"),
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
        "BD Rupturas.xlsx",
        "BD_Rupturas.xlsx",
        "BD Rupturas.xls",
    )
    ARCHIVO_BD_RUP_CSV = resolver_archivo(
        "BD Rupturas - Hoja 1.csv",
        "BD_Rupturas_-_Hoja_1.csv",
        "BD Rupturas Hoja 1.csv",
    )
    usa_excel = os.path.exists(ARCHIVO_BD_RUP)

    if not usa_excel and not os.path.exists(ARCHIVO_BD_RUP_CSV):
        st.info(
            f"Sube el archivo **BD Rupturas.xlsx** (hojas 'Cedis' y 'Plantas') a la "
            "raíz del repositorio para activar el seguimiento histórico."
        )
        return

    if usa_excel:
        # Archivo con dos hojas: 'Cedis' y 'Plantas'. Se etiqueta cada fila con su
        # origen para poder filtrar por vista (Todo/Cedis/Plantas) más abajo.
        try:
            xls = pd.ExcelFile(ARCHIVO_BD_RUP)
        except Exception as e:
            st.error(f"No pude leer el archivo: {e}")
            return
        hojas_por_origen = {}
        for hoja in xls.sheet_names:
            hn = norm(hoja)
            if hn == "CEDIS":
                hojas_por_origen["Cedis"] = hoja
            elif hn == "PLANTAS":
                hojas_por_origen["Plantas"] = hoja
        if not hojas_por_origen:
            st.error(
                f"El archivo **{os.path.basename(ARCHIVO_BD_RUP)}** no tiene hojas "
                f"'Cedis' ni 'Plantas'. Hojas encontradas: `{xls.sheet_names}`"
            )
            return
        partes = []
        for origen, hoja in hojas_por_origen.items():
            dfo = pd.read_excel(xls, sheet_name=hoja)
            dfo["origen"] = origen
            partes.append(dfo)
        rup_raw = pd.concat(partes, ignore_index=True)
    else:
        try:
            for _enc in ("utf-8-sig", "latin-1", "cp1252", "utf-8"):
                try:
                    # sep=None + engine='python': pandas detecta automáticamente el separador
                    # (, ; \t etc.) sin importar cómo se exportó desde Excel o Google Sheets.
                    rup_raw = pd.read_csv(ARCHIVO_BD_RUP_CSV, encoding=_enc,
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
        rup_raw["origen"] = "Cedis"

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
        elif cn == "OTRO":
            norm_map[c] = "otro"
    rup = rup_raw.rename(columns=norm_map)

    if "responsable_inv" not in rup.columns:
        rup["responsable_inv"] = ""
    if "explicacion" not in rup.columns:
        rup["explicacion"] = ""
    if "otro" not in rup.columns:
        rup["otro"] = ""

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
    rup["otro"] = rup["otro"].fillna("").astype(str).str.strip()
    rup["responsable_inv"] = rup["responsable_inv"].fillna("").astype(str).str.strip()
    rup["gestionada"] = rup["explicacion"] != ""
    rup["fecha_dt"] = pd.to_datetime(rup["fecha_corte"], dayfirst=True, errors="coerce")
    # "Fecha corte" se deja como fecha corta (sin hora) en todo el módulo: tablas,
    # gráfico de evolución, etc.
    rup["fecha_corte"] = rup["fecha_dt"].dt.date

    # --- Selector de vista (Todo / Cedis / Plantas) ---
    vista = "Todo"
    if "origen" in rup.columns and rup["origen"].nunique() > 1:
        vista = st.radio("Vista", ["Todo", "Cedis", "Plantas"], horizontal=True)
        if vista != "Todo":
            rup = rup[rup["origen"] == vista]
    etiqueta_vista = {"Cedis": "CEDIs", "Plantas": "Plantas", "Todo": "Destinos"}[vista]

    # --- Filtros (afectan todo el módulo, incluyendo los KPIs iniciales) ---
    fechas_validas = rup["fecha_dt"].dropna()
    hoy = _dt.date.today()
    fecha_min = fechas_validas.min().date() if not fechas_validas.empty else hoy
    fecha_max = fechas_validas.max().date() if not fechas_validas.empty else hoy

    col_fecha, col_dest = st.columns([1, 2])
    with col_fecha:
        rango_fecha = st.date_input(
            "Rango de fecha (fecha de corte)",
            value=(max(hoy.replace(day=1), fecha_min), min(hoy, fecha_max)),
            min_value=fecha_min,
            max_value=fecha_max,
        )
    if isinstance(rango_fecha, (tuple, list)) and len(rango_fecha) == 2:
        f_ini, f_fin = rango_fecha
        rup = rup[(rup["fecha_dt"].dt.date >= f_ini) & (rup["fecha_dt"].dt.date <= f_fin)]

    with col_dest:
        destinos_disp = sorted(rup["destino"].dropna().unique().tolist())
        f_dest = st.multiselect("Destino", destinos_disp, placeholder="Todos")

    base = rup.copy()
    base["responsable_inv"] = base["responsable_inv"].replace("", "Sin asignar")
    if f_dest:
        base = base[base["destino"].isin(f_dest)]

    if base.empty:
        st.info("No hay registros con los filtros seleccionados.")
        return

    # Mapa de color por causa, calculado UNA vez sobre todas las causas del periodo
    # filtrado (orden ascendente por número de causa) para que cada causa conserve
    # siempre el mismo color sin importar el gráfico u orden de despliegue.
    causas_todas = sorted(
        base["explicacion"].replace("", "Sin explicación").unique().tolist(),
        key=_numero_causa,
    )
    color_causa_map = _colores_causas(causas_todas)

    total = len(base)
    gestionadas = int(base["gestionada"].sum())
    pendientes = total - gestionadas
    pct = (gestionadas / total * 100) if total else 0

    st.divider()

    # --- KPI principal: donut de gestión + tarjetas ---
    col_donut, col_kpis = st.columns([1, 2], gap="large")

    with col_donut:
        # Medidor: el relleno lleva el avance; la pista es un paso claro del
        # mismo verde, para que el estado se lea en todo el anillo.
        fig_d = go.Figure(go.Pie(
            values=[max(pct, 0.5), max(100 - pct, 0)],
            hole=0.78,
            marker=dict(colors=[COLOR_PRIMARIO, "#E4F1DF"],
                        line=dict(color=SUPERFICIE, width=2)),
            textinfo="none",
            hoverinfo="skip",
            direction="clockwise",
            sort=False,
            rotation=90,
        ))
        fig_d.update_layout(
            showlegend=False,
            annotations=[
                {
                    "text": f"<b>{pct:.0f}%</b>",
                    "x": 0.5, "y": 0.54,
                    "font": {"size": 38, "color": TINTA, "family": FUENTE_GRAFICO},
                    "showarrow": False, "xanchor": "center", "yanchor": "middle",
                },
                {
                    "text": "gestionado",
                    "x": 0.5, "y": 0.34,
                    "font": {"size": 11.5, "color": TINTA_3, "family": FUENTE_GRAFICO},
                    "showarrow": False, "xanchor": "center", "yanchor": "middle",
                },
            ],
            title=dict(
                text="Nivel de gestión",
                x=0.5, xanchor="center", y=0.97,
                font=dict(size=12.5, color=TINTA_2, family=FUENTE_GRAFICO),
            ),
            margin=dict(l=16, r=16, t=40, b=8),
            height=232,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(family=FUENTE_GRAFICO),
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
    st.markdown(
        titulo_seccion(
            "Evolución de rupturas",
            "Rupturas por fecha de corte. <b>Total</b> va en tinta neutra por ser la "
            "envolvente (siempre la suma de las otras dos), no una serie más.",
        ),
        unsafe_allow_html=True,
    )
    if "fecha_corte" in base.columns:
        serie = base.groupby(["fecha_corte", "gestionada"]).size().reset_index(name="n")
        pivot = serie.pivot(index="fecha_corte", columns="gestionada", values="n").fillna(0)
        pivot = pivot.rename(columns={True: "Gestionadas", False: "Pendientes"})
        for col in ("Gestionadas", "Pendientes"):
            if col not in pivot.columns:
                pivot[col] = 0
        pivot = pivot.sort_index()
        pivot["Total"] = pivot["Gestionadas"] + pivot["Pendientes"]

        # 'Total' es la envolvente, no una categoría par: va en tinta neutra.
        # Antes iba en naranja de marca, que frente al verde de 'Gestionadas'
        # medía ΔE 0.3 en protanopia (indistinguibles); el neutro lo resuelve y
        # además lee mejor como línea de referencia.
        fig_l = go.Figure()
        fig_l.add_trace(go.Scatter(
            x=pivot.index, y=pivot["Total"],
            name="Total", mode="lines+markers",
            line=dict(color="#454F49", width=2, shape="linear"),
            marker=dict(size=8, line=dict(color=SUPERFICIE, width=2)),
            hovertemplate="%{x|%Y-%m-%d}<br>Total: %{y:,.0f}<extra></extra>",
        ))
        fig_l.add_trace(go.Scatter(
            x=pivot.index, y=pivot["Gestionadas"],
            name="Gestionadas", mode="lines+markers",
            line=dict(color=COLOR_PRIMARIO, width=2),
            marker=dict(size=8, line=dict(color=SUPERFICIE, width=2)),
            hovertemplate="%{x|%Y-%m-%d}<br>Gestionadas: %{y:,.0f}<extra></extra>",
        ))
        fig_l.add_trace(go.Scatter(
            x=pivot.index, y=pivot["Pendientes"],
            name="Pendientes", mode="lines+markers",
            line=dict(color=COLOR_CRITICO, width=2),
            marker=dict(size=8, line=dict(color=SUPERFICIE, width=2)),
            hovertemplate="%{x|%Y-%m-%d}<br>Pendientes: %{y:,.0f}<extra></extra>",
        ))
        estilo_grafico(fig_l, alto=350, mostrar_leyenda=True,
                       margen=dict(l=8, r=18, t=34, b=8))
        fig_l.update_layout(hovermode="x unified")
        fig_l.update_xaxes(title="Fecha de corte")
        fig_l.update_yaxes(title="N° de rupturas", separatethousands=True)
        st.plotly_chart(fig_l, use_container_width=True)
    else:
        st.info("El archivo no contiene columna de fecha de corte.")

    st.divider()

    # --- Rupturas por fecha, desglosado por causa ---
    st.markdown(
        titulo_seccion(
            "📅 Rupturas por fecha y causa",
            "Número de rupturas por fecha de corte, desglosado por causa. Incluye "
            "todas las rupturas del periodo filtrado; las que aún no tienen "
            "explicación registrada se agrupan como 'Sin explicación'.",
        ),
        unsafe_allow_html=True,
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
        max_n = int(fecha_causa.groupby("fecha_dt").size().max())

        fig_fecha = go.Figure()
        for causa in causas_f:
            conteo = (
                fecha_causa[fecha_causa["causa_mostrada"] == causa]
                .groupby("fecha_dt").size()
                .reindex(orden_fechas, fill_value=0)
            )
            valores = conteo.values.astype(int)
            relleno = color_causa_map.get(causa, "#C9CDD1")
            fig_fecha.add_trace(go.Bar(
                x=etiquetas_fecha, y=valores, name=causa,
                # Línea de 2px en color de superficie = separación entre segmentos
                # apilados (el hueco es lo que separa, no un borde de contraste).
                marker=dict(color=relleno,
                            line=dict(color=SUPERFICIE, width=2)),
                # Solo se rotula el segmento que tiene altura para contener el
                # número; el resto lo llevan el tooltip y las tablas de abajo.
                text=[str(v) if v >= 2 else "" for v in valores],
                textposition="inside",
                insidetextfont=dict(color=tinta_sobre(relleno), size=11),
                constraintext="both",
                hovertemplate=f"{causa}<br>%{{x}}: %{{y}} rupturas<extra></extra>",
            ))
        fig_fecha.update_layout(barmode="stack", bargap=0.4)
        estilo_grafico(fig_fecha, alto=380, mostrar_leyenda=True,
                       margen=dict(l=8, r=18, t=52, b=8))
        fig_fecha.update_layout(legend=dict(traceorder="normal"))
        fig_fecha.update_xaxes(title="Fecha de corte", type="category")
        fig_fecha.update_yaxes(title="N° de rupturas", tick0=0, dtick=1,
                               range=[0, max_n + 1])
        st.plotly_chart(fig_fecha, use_container_width=True)

    st.divider()

    # --- Ranking de destinos y de causas con más rupturas (acumulado del registro) ---
    if vista == "Plantas":
        st.markdown(
            titulo_seccion(
                "🔎 Rupturas por planta y top 5 causas",
                "Rupturas por cada planta y ranking de las 5 causas con más rupturas de lo "
                "que va del periodo filtrado; las que aún no tienen explicación registrada se "
                "agrupan como 'Sin explicación'.",
            ),
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            titulo_seccion(
                f"🔎 Top 5 {etiqueta_vista.lower()} y causas con más rupturas",
                f"Ranking de los 5 {etiqueta_vista.lower()} y las 5 causas con más rupturas de lo "
                "que va del periodo filtrado; las que aún no tienen explicación registrada se "
                "agrupan como 'Sin explicación'.",
            ),
            unsafe_allow_html=True,
        )
    causas_base = base.copy()
    causas_base["causa_mostrada"] = causas_base["explicacion"].replace("", "Sin explicación")

    col_top_cedi, col_top_causa = st.columns(2, gap="large")

    with col_top_cedi:
        conteo_cedis = causas_base.groupby("destino").size().sort_values(ascending=False)
        if vista == "Plantas":
            st.markdown("##### 🏭 Rupturas por Planta")
        else:
            st.markdown(f"##### 🏭 Top 5 {etiqueta_vista} con más rupturas")
            conteo_cedis = conteo_cedis.head(5)
        st.caption("De lo que va del registro, con los filtros de destino aplicados.")
        orden_cedis = conteo_cedis.index.tolist()[::-1]
        valores_cedis = conteo_cedis.values[::-1]
        fig_top_cedi = go.Figure(go.Bar(
            y=orden_cedis, x=valores_cedis, orientation="h",
            marker=dict(color=COLOR_PRIMARIO, cornerradius=4),
            text=[str(v) for v in valores_cedis],
            textposition="outside",
            textfont=dict(size=11.5, color=TINTA_2),
            cliponaxis=False,
            hovertemplate="%{y}: %{x} rupturas<extra></extra>",
        ))
        fig_top_cedi.update_layout(bargap=0.45)
        estilo_grafico(fig_top_cedi, alto=max(260, 46 * len(orden_cedis)),
                       margen=dict(l=8, r=34, t=8, b=8))
        fig_top_cedi.update_xaxes(title="N° de rupturas", showgrid=True,
                                  gridcolor=REJILLA, tick0=0, dtick=1,
                                  range=[0, int(conteo_cedis.max()) + 1])
        fig_top_cedi.update_yaxes(title="", showgrid=False, linecolor=BORDE,
                                  tickfont=dict(size=12, color=TINTA_2))
        st.plotly_chart(fig_top_cedi, use_container_width=True)

    with col_top_causa:
        st.markdown("##### 🔎 Top 5 causas con más rupturas")
        st.caption("De lo que va del registro, con los filtros de destino aplicados.")
        top_causas_n = (
            causas_base.groupby("causa_mostrada").size().sort_values(ascending=False).head(5)
        )
        orden_top_causas = top_causas_n.index.tolist()[::-1]
        valores_top_causas = top_causas_n.values[::-1]
        fig_top_causa = go.Figure(go.Bar(
            y=orden_top_causas, x=valores_top_causas, orientation="h",
            marker=dict(
                color=[color_causa_map.get(c, "#C9CDD1") for c in orden_top_causas],
                cornerradius=4,
            ),
            text=[str(v) for v in valores_top_causas],
            textposition="outside",
            textfont=dict(size=11.5, color=TINTA_2),
            cliponaxis=False,
            hovertemplate="%{y}: %{x} rupturas<extra></extra>",
        ))
        fig_top_causa.update_layout(bargap=0.45)
        estilo_grafico(fig_top_causa, alto=max(260, 46 * len(orden_top_causas)),
                       margen=dict(l=8, r=34, t=8, b=8))
        fig_top_causa.update_xaxes(title="N° de rupturas", showgrid=True,
                                   gridcolor=REJILLA, tick0=0, dtick=1,
                                   range=[0, int(top_causas_n.max()) + 1])
        fig_top_causa.update_yaxes(title="", showgrid=False, linecolor=BORDE,
                                   tickfont=dict(size=12, color=TINTA_2))
        st.plotly_chart(fig_top_causa, use_container_width=True)

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
    gest_df = base[base["gestionada"] & base["explicacion"].apply(_es_motivo_otro)].copy()

    if not pend_df.empty:
        st.markdown(
            titulo_seccion(f"⏳ Rupturas pendientes de gestión ({len(pend_df):,})"),
            unsafe_allow_html=True,
        )
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
        st.markdown(
            titulo_seccion(f"✅ Rupturas gestionadas con motivo Otro ({len(gest_df):,})"),
            unsafe_allow_html=True,
        )
        cols_g = [c for c in ["fecha_corte", "destino", "item",
                               "referencia", "responsable_inv", "otro"] if c in gest_df.columns]
        vista_g = gest_df[cols_g].rename(columns={
            "fecha_corte": "Fecha corte",
            "destino": "Destino", "item": "Item", "referencia": "Referencia",
            "responsable_inv": "Responsable Inv", "otro": "Otro",
        })
        st.dataframe(vista_g, use_container_width=True, hide_index=True, height=280)


# ===========================================================================
# NAVEGACIÓN
# ===========================================================================
with st.sidebar:
    st.markdown(
        '<div class="marca">'
        '<div class="marca-logo">🥚</div>'
        '<div class="marca-txt">'
        '<div class="marca-nombre">Huevos Kikes</div>'
        '<div class="marca-sub">Panel de inventarios</div>'
        '</div></div>',
        unsafe_allow_html=True,
    )
    st.markdown('<div class="nav-rotulo">Módulos</div>', unsafe_allow_html=True)
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

    # Origen de datos: mismos archivos de siempre, con la fecha en que se
    # cargaron por última vez, para saber de un vistazo qué tan fresco está cada uno.
    st.markdown('<div class="nav-rotulo">Origen de datos</div>', unsafe_allow_html=True)
    filas_archivos = []
    for etiqueta, ruta in [
        ("Inventario hoy", ARCHIVO_HOY),
        ("Inventario ayer", ARCHIVO_AYER),
        ("Ventas", ARCHIVO_VENTAS),
        ("Pedidos", ARCHIVO_PEDIDOS),
        ("Inventarios (tránsito)", ARCHIVO_INVENTARIOS),
        ("Kardex", ARCHIVO_KARDEX),
    ]:
        ts = mtime(ruta)
        if ts:
            sello = _dt.datetime.fromtimestamp(ts).strftime("%d %b · %H:%M")
            punto, color_pt = "●", COLOR_PRIMARIO
        else:
            sello, punto, color_pt = "sin cargar", "○", TINTA_3
        filas_archivos.append(
            f'<div style="display:flex; align-items:center; gap:8px; padding:3px 0;">'
            f'<span style="color:{color_pt}; font-size:0.6rem;">{punto}</span>'
            f'<span style="flex:1; color:{TINTA_2};">{etiqueta}</span>'
            f'<span style="color:{TINTA_3}; font-variant-numeric:tabular-nums;">{sello}</span>'
            f'</div>'
        )
    st.markdown(
        f'<div style="font-size:0.755rem; line-height:1.45;">{"".join(filas_archivos)}</div>',
        unsafe_allow_html=True,
    )
    st.caption("Sube los archivos a la raíz del repositorio (con espacio o guion bajo).")

if modulo == "Inventario de Edades":
    render_modulo_edades()
elif modulo == "Análisis de Rotación PEPS":
    render_modulo_rotacion()
else:
    render_modulo_seguimiento()