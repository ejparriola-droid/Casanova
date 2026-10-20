import re
import unicodedata
import io
import hashlib
from difflib import get_close_matches

import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path
from PIL import Image, ImageOps
import pytesseract


# =========================================================
# CONFIGURACIÓN
# =========================================================

st.set_page_config(
    page_title="Casanova | Unidades en piso | Centenario / Lala",
    page_icon="🚗",
    layout="wide"
)

ARCHIVO_BASE = Path("concentrado.xlsx")
HOJA_BASE = "concentrado"
HOJA_REFACCIONES = "Refacciones"

USUARIO_CORRECTO = "Casanova"
PASSWORD_CORRECTA = "Casanova.50"

AZUL_MARINO = "#213769"
AZUL_CASANOVA = "#004876"
AMARILLO = "#FFB700"
BLANCO = "#FFFFFF"

DIAS_AMBAR = 30
DIAS_ROJO = 60

# Tiempo (segundos) que se conserva el Excel en memoria
CACHE_TTL = 300


# =========================================================
# ESTILO GLOBAL
# =========================================================

def aplicar_estilo():

    st.markdown(
        f"""
<style>

@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"], .stApp {{
    font-family: 'Inter', sans-serif;
}}

.stApp {{
    background: linear-gradient(180deg, #EEF3FA 0%, #F7F9FC 280px);
}}

.stApp p, .stApp span, .stApp label, .stMarkdown {{
    font-weight: 600;
}}

#MainMenu {{visibility: hidden;}}
footer {{visibility: hidden;}}
header {{visibility: hidden;}}

.block-container {{
    padding-top: 1.5rem;
    padding-bottom: 3rem;
    max-width: 1250px;
}}

/* ENCABEZADO */
.topbar {{
    background: linear-gradient(135deg, {AZUL_MARINO}, {AZUL_CASANOVA});
    border-radius: 18px;
    padding: 22px 28px;
    display: flex;
    align-items: center;
    gap: 16px;
    box-shadow: 0 8px 24px rgba(33,55,105,0.20);
    border-bottom: 4px solid {AMARILLO};
}}
.topbar-icono {{
    font-size: 36px;
    background: rgba(255,255,255,0.12);
    border-radius: 14px;
    padding: 8px 12px;
}}
.topbar-titulo {{
    color: white;
    font-size: 24px;
    font-weight: 800;
    letter-spacing: 0.5px;
    line-height: 1.1;
}}
.topbar-sub {{
    color: rgba(255,255,255,0.8);
    font-size: 13px;
    font-weight: 600;
    margin-top: 3px;
}}

/* SECCIONES */
.titulo-seccion {{
    color: {AZUL_MARINO};
    font-size: 18px;
    font-weight: 800;
    margin: 26px 0 12px 0;
    padding-left: 12px;
    border-left: 5px solid {AMARILLO};
    line-height: 1.2;
}}

/* TARJETAS */
.tarjeta {{
    background: {BLANCO};
    border: 1px solid #E1E8F3;
    border-top: 3px solid {AZUL_CASANOVA};
    border-radius: 12px;
    padding: 14px 16px;
    min-height: 84px;
    box-shadow: 0 2px 8px rgba(33,55,105,0.05);
    margin-bottom: 12px;
    transition: all 0.2s ease;
}}
.tarjeta:hover {{
    transform: translateY(-2px);
    box-shadow: 0 8px 18px rgba(33,55,105,0.12);
    border-top-color: {AMARILLO};
}}
.campo {{
    color: #7A879B;
    font-size: 11px;
    font-weight: 800;
    letter-spacing: 0.8px;
    text-transform: uppercase;
    margin-bottom: 6px;
}}
.valor {{
    color: #172033;
    font-size: 16px;
    font-weight: 800;
    word-break: break-word;
    line-height: 1.3;
}}
.valor-vacio {{
    color: #A7B1C2;
    font-weight: 600;
    font-style: italic;
}}

/* BADGES */
.badge {{
    display: inline-block;
    padding: 4px 12px;
    border-radius: 20px;
    font-size: 13px;
    font-weight: 800;
}}
.badge-verde {{ background: #E3F6EA; color: #167A3E; }}
.badge-amarillo {{ background: #FFF3D0; color: #8A6100; }}
.badge-rojo {{ background: #FDE5E5; color: #B42323; }}
.badge-azul {{ background: #E3ECF9; color: {AZUL_MARINO}; }}

/* KPIs */
.kpi-azul {{
    background: linear-gradient(135deg, {AZUL_MARINO}, {AZUL_CASANOVA});
    color: white;
    border-radius: 16px;
    padding: 20px 24px;
    box-shadow: 0 6px 18px rgba(33,55,105,0.20);
    height: 100%;
}}
.kpi-amarillo {{
    background: white;
    border: 2px solid {AMARILLO};
    border-radius: 16px;
    padding: 20px 24px;
    box-shadow: 0 6px 18px rgba(255,183,0,0.15);
    height: 100%;
}}
.kpi-label {{
    font-size: 12px;
    font-weight: 800;
    letter-spacing: 1px;
    text-transform: uppercase;
    opacity: 0.9;
}}
.kpi-numero {{
    font-size: 40px;
    font-weight: 800;
    margin-top: 4px;
    line-height: 1.1;
}}
.kpi-amarillo .kpi-label {{ color: #7A879B; opacity: 1; }}
.kpi-amarillo .kpi-numero {{ color: {AZUL_MARINO}; }}
.kpi-subtotal {{
    color: rgba(255,255,255,0.85);
    font-size: 13px;
    font-weight: 700;
    margin-top: 6px;
}}

/* IDENTIDAD DE LA UNIDAD */
.unidad-header {{
    background: white;
    border: 1px solid #E1E8F3;
    border-radius: 16px;
    padding: 18px 24px;
    margin: 20px 0 4px 0;
    box-shadow: 0 2px 10px rgba(33,55,105,0.06);
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 14px 30px;
}}
.unidad-placa {{
    background: {AMARILLO};
    color: {AZUL_MARINO};
    font-size: 22px;
    font-weight: 800;
    letter-spacing: 2px;
    padding: 6px 18px;
    border-radius: 8px;
    border: 2px solid {AZUL_MARINO};
}}
.unidad-serie-label {{
    color: #7A879B;
    font-size: 11px;
    font-weight: 800;
    letter-spacing: 0.8px;
}}
.unidad-serie {{
    color: {AZUL_MARINO};
    font-size: 17px;
    font-weight: 800;
    word-break: break-all;
}}

/* REFACCIONES */
.refaccion-card {{
    background: white;
    border: 1px solid #E1E8F3;
    border-left: 5px solid {AMARILLO};
    border-radius: 12px;
    padding: 16px 18px;
    margin-bottom: 12px;
    box-shadow: 0 2px 8px rgba(33,55,105,0.05);
}}
.refaccion-folio {{
    color: {AZUL_MARINO};
    font-size: 14px;
    font-weight: 800;
    margin-bottom: 8px;
    letter-spacing: 0.3px;
}}
.refaccion-actividad {{
    color: #172033;
    font-size: 14.5px;
    font-weight: 600;
    line-height: 1.7;
}}

/* SIN FOLIOS */
.sin-folios {{
    background: #E3F6EA;
    border: 1px solid #BFE6CD;
    border-radius: 12px;
    padding: 15px 18px;
    color: #167A3E;
    font-weight: 800;
    margin-top: 24px;
}}

/* SEPARADOR */
.separador {{
    height: 1px;
    background: linear-gradient(90deg, transparent, #CBD6E8, transparent);
    margin: 30px 0 6px 0;
}}

/* INPUTS Y BOTONES */
.stTextInput label,
.stTextInput label p,
.stSelectbox label,
.stSelectbox label p {{
    color: {AZUL_MARINO} !important;
    font-weight: 800 !important;
    font-size: 15px !important;
}}
.stTextInput div[data-baseweb="input"],
.stTextInput div[data-baseweb="base-input"] {{
    background: white !important;
    border-radius: 12px !important;
    color-scheme: light !important;
}}
.stTextInput input {{
    color: #172033 !important;
    -webkit-text-fill-color: #172033 !important;
    caret-color: {AZUL_MARINO} !important;
    background: white !important;
    background-color: white !important;
    border-radius: 12px !important;
    border: 2px solid #D8E1EF !important;
    padding: 12px 16px !important;
    font-size: 16px !important;
    font-weight: 700 !important;
    color-scheme: light !important;
    box-shadow: none !important;
}}
.stTextInput input:hover {{
    background: white !important;
    background-color: white !important;
}}
.stTextInput input:focus {{
    background: white !important;
    background-color: white !important;
    border-color: {AZUL_CASANOVA} !important;
    box-shadow: 0 0 0 3px rgba(0,72,118,0.12) !important;
    color: #172033 !important;
    -webkit-text-fill-color: #172033 !important;
}}
.stTextInput input::placeholder {{
    color: #64748B !important;
    -webkit-text-fill-color: #64748B !important;
    font-weight: 600 !important;
    opacity: 1 !important;
}}
.stTextInput button {{
    background: white !important;
    color: {AZUL_MARINO} !important;
    border: none !important;
}}
.stTextInput button svg {{
    fill: {AZUL_MARINO} !important;
    color: {AZUL_MARINO} !important;
}}
.stSelectbox div[data-baseweb="select"] > div {{
    background: white !important;
    color: #172033 !important;
    border-radius: 12px !important;
}}
.stSelectbox div[data-baseweb="select"] span,
.stSelectbox div[data-baseweb="select"] div {{
    color: #172033 !important;
    font-weight: 700 !important;
}}
div[data-baseweb="popover"] ul,
div[data-baseweb="popover"] li {{
    background: white !important;
    color: #172033 !important;
    font-weight: 600 !important;
}}
div[data-baseweb="popover"] li:hover {{
    background: #EEF3FA !important;
}}
.stButton > button {{
    border-radius: 12px;
    font-weight: 700;
    padding: 10px 18px;
    border: 2px solid {AZUL_MARINO};
    transition: all 0.2s ease;
}}
.stButton > button[kind="primary"] {{
    background: {AZUL_MARINO};
    color: white;
}}
.stButton > button:hover {{
    border-color: {AMARILLO};
    color: {AZUL_MARINO};
    background: {AMARILLO};
}}

/* LOGIN */
.login-box {{
    max-width: 430px;
    margin: 50px auto 24px auto;
    background: white;
    padding: 34px 30px 26px 30px;
    border-radius: 20px;
    border: 1px solid #D8E1EF;
    border-top: 6px solid {AMARILLO};
    box-shadow: 0 12px 40px rgba(33,55,105,0.15);
    text-align: center;
}}
.login-icon {{ font-size: 50px; margin-bottom: 6px; }}
.login-title {{
    color: {AZUL_MARINO};
    font-size: 28px;
    font-weight: 800;
    letter-spacing: 2px;
}}
.login-subtitle {{
    color: #64748B;
    font-size: 14px;
    font-weight: 600;
    margin-top: 4px;
}}

/* SEMÁFORO KPI */
.kpi-verde {{
    background: linear-gradient(135deg, #127A3D, #1FA855);
    color: white;
    border-radius: 16px;
    padding: 20px 24px;
    box-shadow: 0 6px 18px rgba(22,122,62,0.25);
    height: 100%;
}}
.kpi-ambar {{
    background: linear-gradient(135deg, #E59E00, {AMARILLO});
    color: #2A1D00;
    border-radius: 16px;
    padding: 20px 24px;
    box-shadow: 0 6px 18px rgba(255,183,0,0.30);
    height: 100%;
}}
.kpi-rojo {{
    background: linear-gradient(135deg, #A81E1E, #E04040);
    color: white;
    border-radius: 16px;
    padding: 20px 24px;
    box-shadow: 0 6px 18px rgba(200,40,40,0.28);
    height: 100%;
}}
.kpi-sub {{
    font-size: 13px;
    font-weight: 800;
    margin-top: 8px;
    opacity: 0.95;
}}

/* RESUMEN DE FOLIOS */
.resumen {{
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
    margin: 4px 0 8px 0;
}}
.chip {{
    display: inline-flex;
    align-items: center;
    gap: 8px;
    padding: 8px 16px;
    border-radius: 999px;
    font-size: 13px;
    font-weight: 800;
    border: 1px solid;
}}
.chip-on {{
    background: #FFF3D0;
    color: #8A6100;
    border-color: #F5D77A;
}}
.chip-off {{
    background: #F1F4F9;
    color: #94A3B8;
    border-color: #E1E8F3;
}}
.chip-punto {{
    width: 9px;
    height: 9px;
    border-radius: 50%;
    background: currentColor;
}}

/* ENCABEZADO DE SECCIÓN CON DÍAS */
.seccion-head {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    flex-wrap: wrap;
    gap: 10px;
    margin: 26px 0 12px 0;
}}
.seccion-head .titulo-seccion {{ margin: 0; }}
.dias-pill {{
    padding: 6px 16px;
    border-radius: 999px;
    font-size: 13px;
    font-weight: 800;
}}
.dias-verde {{ background: #E3F6EA; color: #167A3E; }}
.dias-ambar {{ background: #FFF3D0; color: #8A6100; }}
.dias-rojo {{ background: #FDE5E5; color: #B42323; }}

/* ANIMACIÓN */
@keyframes aparecer {{
    from {{ opacity: 0; transform: translateY(8px); }}
    to {{ opacity: 1; transform: translateY(0); }}
}}
.tarjeta, .refaccion-card, .unidad-header,
.kpi-azul, .kpi-verde, .kpi-ambar, .kpi-rojo, .kpi-amarillo {{
    animation: aparecer 0.35s ease both;
}}

/* BOTÓN DE DESCARGA */
.stDownloadButton > button {{
    border-radius: 12px;
    font-weight: 700;
    padding: 10px 18px;
    border: 2px solid {AZUL_MARINO};
    background: white;
    color: {AZUL_MARINO};
    transition: all 0.2s ease;
}}
.stDownloadButton > button:hover {{
    background: {AMARILLO};
    border-color: {AMARILLO};
    color: {AZUL_MARINO};
}}

/* CELULAR */
@media (max-width: 768px) {{
    .topbar-titulo {{ font-size: 18px; }}
    .kpi-numero {{ font-size: 30px; }}
    .unidad-placa {{ font-size: 18px; }}
}}

</style>
""",
        unsafe_allow_html=True
    )


aplicar_estilo()


# =========================================================
# LOGIN
# =========================================================

def verificar_login(usuario, password):
    return (
        usuario.strip() == USUARIO_CORRECTO
        and password == PASSWORD_CORRECTA
    )


def mostrar_login():

    st.markdown(
        """
<div class="login-box">
<div class="login-icon">🚗</div>
<div class="login-title">CASANOVA</div>
<div class="login-subtitle">Control de Unidades</div>
</div>
""",
        unsafe_allow_html=True
    )

    col1, col2, col3 = st.columns([1, 1.4, 1])

    with col2:

        usuario = st.text_input("Usuario", placeholder="Usuario")

        password = st.text_input(
            "Contraseña",
            type="password",
            placeholder="Contraseña"
        )

        iniciar = st.button(
            "🔐 Iniciar sesión",
            use_container_width=True,
            type="primary"
        )

        if iniciar:

            if verificar_login(usuario, password):
                st.session_state["autenticado"] = True
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")


if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    mostrar_login()
    st.stop()


# =========================================================
# FUNCIONES GENERALES
# =========================================================

def tiene_informacion(valor):

    if valor is None:
        return False

    try:
        if pd.isna(valor):
            return False
    except Exception:
        pass

    texto = str(valor).strip()

    if texto == "":
        return False

    if texto.lower() in ["nan", "none", "null", "nat"]:
        return False

    return True


def valor_visible(valor):

    if not tiene_informacion(valor):
        return "Sin información"

    return str(valor).strip()


def numero_entero(valor):

    if not tiene_informacion(valor):
        return "Sin información"

    try:
        return f"{int(float(valor)):,}"
    except Exception:
        return valor_visible(valor)


def modelo_visible(valor):

    if not tiene_informacion(valor):
        return "Sin información"

    try:
        return str(int(float(valor)))
    except Exception:
        return valor_visible(valor)


def fecha_visible(valor):

    if not tiene_informacion(valor):
        return "Sin información"

    try:
        fecha = pd.to_datetime(valor, errors="coerce")

        if pd.isna(fecha):
            return valor_visible(valor)

        return fecha.strftime("%d/%m/%Y")

    except Exception:
        return valor_visible(valor)


# ---------------------------------------------------------
# NORMALIZACIÓN ÚNICA PARA PLACA / SERIE
# Quita TODOS los espacios (incluidos \xa0, tabs, saltos de línea),
# guiones, pasa a mayúsculas y elimina el ".0" de números leídos como float.
# ---------------------------------------------------------

def limpiar_valor(valor):

    if not tiene_informacion(valor):
        return ""

    texto = str(valor)
    texto = texto.replace("\xa0", " ")
    texto = re.sub(r"\.0+$", "", texto.strip())
    texto = re.sub(r"[\s\-]+", "", texto)

    return texto.upper()


def limpiar_serie(serie_pandas):
    return serie_pandas.map(limpiar_valor)


def normalizar_texto(valor):
    return limpiar_valor(valor)


def normalizar_columnas(df):

    df = df.copy()

    columnas = []

    for columna in df.columns:

        nombre = (
            str(columna)
            .replace("\n", " ")
            .replace("\r", " ")
            .replace("\xa0", " ")
        )

        columnas.append(" ".join(nombre.split()).strip())

    df.columns = columnas

    return df


def buscar_columna(df, nombre):
    """Encuentra una columna sin importar mayúsculas ni espacios."""

    objetivo = nombre.strip().upper()

    for c in df.columns:
        if str(c).strip().upper() == objetivo:
            return c

    return None


def clave_columna(nombre):
    """Convierte un nombre de columna a una clave comparable:
    sin acentos, sin mayúsculas, sin espacios ni signos."""
    texto = unicodedata.normalize("NFKD", str(nombre))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]", "", texto.lower())


def obtener(fila, nombre):
    """Lee un campo de la fila ignorando acentos, mayúsculas y espacios."""
    objetivo = clave_columna(nombre)

    for c in fila.index:
        if clave_columna(c) == objetivo:
            return fila[c]

    return None


def texto_ejecutivo(valor):
    """Nombre del ejecutivo, o 'Sin ejecutivo asignado' si está en blanco."""
    if not tiene_informacion(valor):
        return "Sin ejecutivo asignado"
    return str(valor).strip()


# =========================================================
# CARGA DE DATOS
# =========================================================

@st.cache_data(ttl=CACHE_TTL)
def cargar_concentrado():

    df = pd.read_excel(
        ARCHIVO_BASE,
        sheet_name=HOJA_BASE
    )

    df = normalizar_columnas(df)

    # Evita que series/placas numéricas se pierdan o cambien de formato
    for nombre in ["SERIE", "PLACA"]:
        col = buscar_columna(df, nombre)
        if col is not None and col != nombre:
            df = df.rename(columns={col: nombre})

    return df


@st.cache_data(ttl=CACHE_TTL)
def cargar_refacciones():

    try:
        df = pd.read_excel(
            ARCHIVO_BASE,
            sheet_name=HOJA_REFACCIONES
        )
        return normalizar_columnas(df)

    except Exception:
        return pd.DataFrame()


# =========================================================
# COMPONENTES VISUALES
# =========================================================

def tarjeta(campo, valor):

    clase = "valor"

    if valor == "Sin información":
        clase = "valor valor-vacio"

    return f"""
<div class="tarjeta">
<div class="campo">{campo}</div>
<div class="{clase}">{valor}</div>
</div>
"""


def tarjeta_badge(campo, valor):

    if valor == "Sin información":
        return tarjeta(campo, valor)

    texto = valor.lower()

    if any(p in texto for p in [
        "finaliz", "complet", "cerrad", "entreg",
        "listo", "vigente", "activ"
    ]):
        clase = "badge-verde"

    elif any(p in texto for p in [
        "cancel", "venc", "rechaz", "detenid"
    ]):
        clase = "badge-rojo"

    elif any(p in texto for p in [
        "pend", "proceso", "espera", "revis",
        "abiert", "tramite", "trámite"
    ]):
        clase = "badge-amarillo"

    else:
        clase = "badge-azul"

    return f"""
<div class="tarjeta">
<div class="campo">{campo}</div>
<div class="valor"><span class="badge {clase}">{valor}</span></div>
</div>
"""


def mostrar_tarjetas(items):

    columnas = st.columns(len(items))

    for col, (tipo, campo, valor) in zip(columnas, items):

        with col:

            if tipo == "b":
                html = tarjeta_badge(campo, valor)
            else:
                html = tarjeta(campo, valor)

            st.markdown(html, unsafe_allow_html=True)


def dias_numero_o_none(valor):

    if not tiene_informacion(valor):
        return None

    try:
        return int(float(valor))
    except Exception:
        return None


def nivel_dias(dias):

    if dias is None:
        return None

    if dias >= DIAS_ROJO:
        return "rojo"

    if dias >= DIAS_AMBAR:
        return "ambar"

    return "verde"


def titulo_seccion(texto, dias=None):

    nivel = nivel_dias(dias)

    if nivel is None:
        html = f'<div class="titulo-seccion">{texto}</div>'
    else:
        html = (
            f'<div class="seccion-head">'
            f'<div class="titulo-seccion">{texto}</div>'
            f'<div class="dias-pill dias-{nivel}">⏱ {dias:,} días</div>'
            f'</div>'
        )

    st.markdown(html, unsafe_allow_html=True)


def separador():
    st.markdown('<div class="separador"></div>', unsafe_allow_html=True)


# =========================================================
# ENCABEZADO
# =========================================================

col1, col2 = st.columns([5, 1], vertical_alignment="center")

with col1:

    st.markdown(
        """
<div class="topbar">
<div class="topbar-icono">🚗</div>
<div>
<div class="topbar-titulo">CASANOVA | CONTROL DE UNIDADES</div>
<div class="topbar-sub">Consulta individual de unidades</div>
</div>
</div>
""",
        unsafe_allow_html=True
    )

with col2:

    if st.button("Cerrar sesión", use_container_width=True):
        st.session_state["autenticado"] = False
        st.rerun()

    if st.button("🔄 Actualizar datos", use_container_width=True):
        st.cache_data.clear()
        st.rerun()


# =========================================================
# CARGAR ARCHIVOS
# =========================================================

try:
    df = cargar_concentrado()
except Exception as e:
    st.error("No se pudo cargar concentrado.xlsx")
    st.code(str(e))
    st.stop()

df_refacciones = cargar_refacciones()

if "SERIE" not in df.columns or "PLACA" not in df.columns:
    st.error(
        "No se encontraron las columnas SERIE y/o PLACA en la hoja "
        f"'{HOJA_BASE}'. Columnas detectadas:"
    )
    st.code(", ".join(df.columns.astype(str)))
    st.stop()


# =========================================================
# CLAVES LIMPIAS (se calculan una sola vez)
# =========================================================

df = df.copy()
df["_PLACA_KEY"] = limpiar_serie(df["PLACA"])
df["_SERIE_KEY"] = limpiar_serie(df["SERIE"])

# Identificador de la unidad: la SERIE; si no tiene, usa la PLACA
df["_UNIDAD_KEY"] = df["_SERIE_KEY"].where(
    df["_SERIE_KEY"] != "",
    df["_PLACA_KEY"]
)


# =========================================================
# TOTAL DE UNIDADES
# =========================================================

claves_validas = df.loc[df["_UNIDAD_KEY"] != "", "_UNIDAD_KEY"]

total_unidades = claves_validas.nunique()


def mostrar_diagnostico():

    filas_totales = len(df)
    sin_serie = int((df["_SERIE_KEY"] == "").sum())
    sin_placa = int((df["_PLACA_KEY"] == "").sum())
    sin_ambas = int(
        ((df["_SERIE_KEY"] == "") & (df["_PLACA_KEY"] == "")).sum()
    )
    series_repetidas = int(
        df.loc[df["_SERIE_KEY"] != "", "_SERIE_KEY"].duplicated().sum()
    )

    with st.expander("🧪 Diagnóstico de datos"):

        st.write(f"Filas en el Excel: **{filas_totales:,}**")
        st.write(f"Unidades únicas contadas: **{total_unidades:,}**")
        st.write(f"Filas sin SERIE: **{sin_serie:,}**")
        st.write(f"Filas sin PLACA: **{sin_placa:,}**")
        st.write(f"Filas sin SERIE ni PLACA (no se cuentan): **{sin_ambas:,}**")
        st.write(f"Filas con SERIE repetida: **{series_repetidas:,}**")

        if series_repetidas:
            repetidas = df[
                (df["_SERIE_KEY"] != "")
                & df["_SERIE_KEY"].duplicated(keep=False)
            ][["PLACA", "SERIE"]].sort_values("SERIE")
            st.dataframe(repetidas, use_container_width=True)


# =========================================================
# OCR - DETECTA PLACA (letras grandes con guion) / SERIE
# =========================================================

LETRAS_A_DIGITO = {"O": "0", "Q": "0", "D": "0", "I": "1", "L": "1",
                   "Z": "2", "S": "5", "B": "8", "G": "6"}
DIGITO_A_LETRA = {"0": "O", "1": "I", "2": "Z", "5": "S", "8": "B", "6": "G"}

VIN_VALIDO = re.compile(r"^[A-HJ-NPR-Z0-9]{17}$")

# Formatos de placa (L = letra, D = dígito), sin guiones
#   R20-BSD   -> LDD-LLL      (CDMX)
#   D06-BND   -> LDD-LLL      (CDMX)
#   MKJ-121-B -> LLL-DDD-L    (Edomex)
# Si tienes otros formatos, agrégalos aquí. Ej: "LLLDDD", "LLLDDDD"
ESQUEMAS_PLACA = ["LDDLLL", "LLLDDDL"]

# Caracteres permitidos por Tesseract (ahora SÍ incluye el guion)
CONFIG_OCR = (
    "-c tessedit_char_whitelist="
    "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789- "
)


def preparar_imagen(img, angulo=0, invertir=False, lado_objetivo=2000):

    img = img.convert("L")

    if angulo:
        img = img.rotate(angulo, expand=True)

    lado = max(img.size)
    factor = lado_objetivo / lado
    img = img.resize(
        (int(img.width * factor), int(img.height * factor)),
        Image.LANCZOS
    )

    img = ImageOps.autocontrast(img, cutoff=2)

    if invertir:
        img = ImageOps.invert(img)

    return img


def corregir_con_esquema(ventana, esquema):
    """Corrige O/0, I/1, S/5, B/8... según si ahí va letra o dígito."""

    if len(ventana) != len(esquema):
        return None

    resultado = ""

    for ch, tipo in zip(ventana, esquema):

        if tipo == "D":
            ch = LETRAS_A_DIGITO.get(ch, ch)
            if not ch.isdigit():
                return None
        else:
            ch = DIGITO_A_LETRA.get(ch, ch)
            if not ch.isalpha():
                return None

        resultado += ch

    return resultado


def placas_en_linea(texto_linea):
    """Busca placas dentro de una línea. Devuelve placas SIN guiones."""

    limpio = re.sub(r"[^A-Z0-9]", "", texto_linea.upper())
    encontradas = []

    for esquema in ESQUEMAS_PLACA:
        n = len(esquema)
        for i in range(len(limpio) - n + 1):
            p = corregir_con_esquema(limpio[i:i + n], esquema)
            if p:
                encontradas.append(p)

    return encontradas


def placas_texto_grande(img_preparada, psm):
    """
    Lee la imagen con posiciones y se queda con el texto de letras GRANDES.
    Devuelve lista de (placa_sin_guiones, puntaje).
    """

    data = pytesseract.image_to_data(
        img_preparada,
        config=f"--psm {psm} {CONFIG_OCR}",
        output_type=pytesseract.Output.DICT
    )

    palabras = []

    for i, t in enumerate(data["text"]):

        t = str(t).strip().upper()

        if not t or not re.search(r"[A-Z0-9\-]", t):
            continue

        try:
            conf = float(data["conf"][i])
        except Exception:
            conf = 0

        if conf < 10:
            continue

        palabras.append({
            "t": t,
            "x": data["left"][i],
            "y": data["top"][i],
            "h": data["height"][i],
        })

    if not palabras:
        return []

    # Agrupar palabras en renglones (mismo centro vertical aprox.)
    palabras.sort(key=lambda p: p["y"] + p["h"] / 2)

    renglones = []

    for p in palabras:
        centro = p["y"] + p["h"] / 2
        if renglones:
            ult = renglones[-1]
            if abs(centro - ult["centro"]) < ult["h"] * 0.6:
                ult["items"].append(p)
                continue
        renglones.append({"centro": centro, "h": p["h"], "items": [p]})

    h_max = max(p["h"] for p in palabras)

    resultados = []

    for r in renglones:

        items = sorted(r["items"], key=lambda p: p["x"])
        texto = " ".join(p["t"] for p in items)
        altura = max(p["h"] for p in items)

        # Ignorar letras chicas (leyendas, municipio, etc.)
        if altura < h_max * 0.5:
            continue

        tiene_guion = "-" in texto

        for placa in placas_en_linea(texto):
            puntos = 10 + 20 * (altura / h_max) + (8 if tiene_guion else 0)
            resultados.append((placa, puntos))

    return resultados


def extraer_vin(texto):
    """Busca números de serie (VIN de 17 caracteres) en el texto."""

    candidatos = []

    texto = texto.upper().replace("*", " ").replace("-", " ")

    for linea in texto.splitlines():

        limpio = re.sub(r"[^A-Z0-9]", "", linea)

        if len(limpio) < 17:
            continue

        base_vin = limpio
        extra = 0

        if "VIN" in limpio:
            base_vin = limpio.split("VIN", 1)[1]
            extra = 10

        for i in range(len(base_vin) - 16):

            v = base_vin[i:i + 17]
            v = v.replace("O", "0").replace("Q", "0").replace("I", "1")

            if (
                VIN_VALIDO.match(v)
                and any(c.isdigit() for c in v)
                and any(c.isalpha() for c in v)
            ):
                candidatos.append((v, 20 + extra))

    return candidatos


def leer_texto_foto(imagen_bytes, claves_conocidas=None):

    try:

        img = Image.open(io.BytesIO(imagen_bytes))
        img = ImageOps.exif_transpose(img).convert("RGB")

        conocidas = set(claves_conocidas or [])
        puntajes = {}

        for angulo in (0, 180, 90, 270):

            for invertir in (False, True):

                for lado in (2000, 3000):

                    trabajo = preparar_imagen(img, angulo, invertir, lado)

                    for psm in (11, 6):

                        # 1) Placa: solo letras grandes, sin guiones
                        try:
                            for cand, pts in placas_texto_grande(trabajo, psm):
                                puntajes[cand] = puntajes.get(cand, 0) + pts
                        except Exception:
                            pass

                        # 2) Serie / VIN
                        texto = pytesseract.image_to_string(
                            trabajo,
                            config=f"--psm {psm} {CONFIG_OCR}"
                        )
                        for cand, pts in extraer_vin(texto):
                            puntajes[cand] = puntajes.get(cand, 0) + pts

            if puntajes and max(puntajes.values()) >= 40:
                break

        if not puntajes:
            return ""

        # 1) Lo que exista en el Excel gana
        en_base = [c for c in puntajes if c in conocidas]
        if en_base:
            return max(en_base, key=lambda c: puntajes[c])

        # 2) Coincidencia aproximada (si falla 1 carácter)
        if conocidas:

            lista = list(conocidas)

            for cand in sorted(puntajes, key=puntajes.get, reverse=True):
                corte = 0.93 if len(cand) == 17 else 0.83
                cerca = get_close_matches(cand, lista, n=1, cutoff=corte)
                if cerca:
                    return cerca[0]

        # 3) Mejor candidato por puntaje
        return max(puntajes, key=puntajes.get)

    except Exception:
        return ""


# =========================================================
# BUSCADOR
# =========================================================

titulo_seccion(
    "🔎 Buscar unidad"
)


# ---------------------------------------------------------
# SESSION STATE
# ---------------------------------------------------------

if "busqueda" not in st.session_state:
    st.session_state["busqueda"] = ""

if "foto_procesada" not in st.session_state:
    st.session_state["foto_procesada"] = ""

if "msg_foto" not in st.session_state:
    st.session_state["msg_foto"] = ("", "")


# ---------------------------------------------------------
# MODO DE BÚSQUEDA: escribir, cámara o subir foto
# ---------------------------------------------------------

modo_foto = st.radio(
    "Cómo quieres buscar",
    ["⌨️ Escribir", "📷 Cámara", "🖼️ Subir foto"],
    horizontal=True,
    label_visibility="collapsed"
)

foto = None

if modo_foto == "📷 Cámara":

    foto = st.camera_input(
        "Apunta a la placa o a la etiqueta del VIN y toma la foto"
    )

    st.caption(
        "Acércate hasta que el texto se vea grande y nítido. "
        "Evita reflejos y sombras."
    )

elif modo_foto == "🖼️ Subir foto":

    foto = st.file_uploader(
        "Selecciona una foto",
        type=["jpg", "jpeg", "png", "webp"]
    )


# ---------------------------------------------------------
# PROCESAR FOTO (una sola vez por foto)
# ---------------------------------------------------------

if foto is not None:

    datos_foto = foto.getvalue()
    hash_foto = hashlib.md5(datos_foto).hexdigest()

    if st.session_state["foto_procesada"] != hash_foto:

        st.session_state["foto_procesada"] = hash_foto

        claves_conocidas = set(df["_PLACA_KEY"]) | set(df["_SERIE_KEY"])
        claves_conocidas.discard("")

        with st.spinner("🔎 Detectando placa / serie..."):
            texto_ocr = leer_texto_foto(datos_foto, claves_conocidas)

        if texto_ocr:

            # Se escribe en el buscador ANTES de crearlo
            st.session_state["busqueda"] = texto_ocr

            st.session_state["msg_foto"] = (
                "ok",
                f"📷 Detectado: {texto_ocr} — revisa que sea correcto."
            )

        else:

            st.session_state["msg_foto"] = (
                "warn",
                "No pude detectar una placa o serie. "
                "Acércate más, con buena luz, y vuelve a tomar la foto."
            )

    tipo_msg, texto_msg = st.session_state["msg_foto"]

    if tipo_msg == "ok":
        st.success(texto_msg)
    elif tipo_msg == "warn":
        st.warning(texto_msg)


# ---------------------------------------------------------
# CAMPO DE BÚSQUEDA
# Quita guiones, asteriscos, espacios y pasa a mayúsculas
# ---------------------------------------------------------

def limpiar_busqueda():
    texto = st.session_state.get("busqueda", "")
    st.session_state["busqueda"] = re.sub(r"[^A-Za-z0-9]", "", texto).upper()


busqueda = st.text_input(
    "Busca por PLACA o SERIE",
    placeholder="Ejemplo: PDN1295 o número de serie...",
    label_visibility="collapsed",
    key="busqueda",
    on_change=limpiar_busqueda
)


# =========================================================
# SI NO HAY BÚSQUEDA
# =========================================================

if not busqueda:

    st.markdown(
        f"""
<div class="kpi-azul" style="margin-top:14px;">
<div class="kpi-label">🚗 Total de unidades</div>
<div class="kpi-numero">{total_unidades:,}</div>
<div class="kpi-subtotal">Centenario / Lala</div>
</div>
""",
        unsafe_allow_html=True
    )

    st.info("Escribe una placa o número de serie para consultar la unidad.")

    mostrar_diagnostico()

    st.stop()


# =========================================================
# BUSCAR (exacto primero, luego parcial)
# =========================================================

criterio = normalizar_texto(busqueda)

if criterio == "":
    st.warning("Escribe una placa o serie válida.")
    st.stop()

exacto = df[
    (df["_PLACA_KEY"] == criterio)
    | (df["_SERIE_KEY"] == criterio)
]

if not exacto.empty:
    resultado = exacto.copy()
else:
    resultado = df[
        df["_PLACA_KEY"].str.contains(criterio, regex=False, na=False)
        | df["_SERIE_KEY"].str.contains(criterio, regex=False, na=False)
    ].copy()


if resultado.empty:
    st.warning(f"No se encontró ninguna unidad con: {busqueda}")
    st.stop()


# =========================================================
# SELECCIONAR UNIDAD
# (se usa la posición, así funciona aunque haya filas repetidas)
# =========================================================

resultado = resultado.reset_index(drop=True)

if len(resultado) > 1:

    st.caption(f"Se encontraron {len(resultado):,} coincidencias.")

    def etiqueta(i):
        fila = resultado.iloc[i]
        return (
            f"{valor_visible(fila.get('PLACA'))} | "
            f"{valor_visible(fila.get('SERIE'))}"
        )

    posicion = st.selectbox(
        "Selecciona la unidad",
        options=list(range(len(resultado))),
        format_func=etiqueta
    )

    unidad = resultado.iloc[posicion]

else:
    unidad = resultado.iloc[0]


# =========================================================
# DATOS PRINCIPALES
# =========================================================

placa = valor_visible(unidad.get("PLACA"))
serie = valor_visible(unidad.get("SERIE"))


# =========================================================
# DÍAS EN PISO
# =========================================================

dias_piso = unidad.get("DÍAS EN PISO")

if tiene_informacion(dias_piso):

    try:
        texto_dias = f"{int(float(dias_piso)):,} días"
    except Exception:
        texto_dias = valor_visible(dias_piso)

else:
    texto_dias = "🆕 Recién ingresado"


# =========================================================
# KPIs
# =========================================================

col1, col2 = st.columns(2)

with col1:

    st.markdown(
        f"""
<div class="kpi-azul" style="margin-top:14px;">
<div class="kpi-label">🚗 Total de unidades</div>
<div class="kpi-numero">{total_unidades:,}</div>
<div class="kpi-subtotal">Centenario / Lala</div>
</div>
""",
        unsafe_allow_html=True
    )

with col2:

    nivel_piso = nivel_dias(dias_numero_o_none(dias_piso))

    clase_kpi = {
        "verde": "kpi-verde",
        "ambar": "kpi-ambar",
        "rojo": "kpi-rojo",
        None: "kpi-amarillo"
    }[nivel_piso]

    texto_nivel = {
        "verde": "✅ Rotación normal",
        "ambar": "⚠️ Requiere atención",
        "rojo": "🚨 Tiempo crítico en piso",
        None: ""
    }[nivel_piso]

    sub_html = (
        f'<div class="kpi-sub">{texto_nivel}</div>'
        if texto_nivel else ""
    )

    st.markdown(
        f"""
<div class="{clase_kpi}" style="margin-top:14px;">
<div class="kpi-label">🏭 Días en piso</div>
<div class="kpi-numero">{texto_dias}</div>
{sub_html}
</div>
""",
        unsafe_allow_html=True
    )


# =========================================================
# IDENTIDAD DE LA UNIDAD
# =========================================================

ejecutivo_ziga = obtener(unidad, "Ziga_Talleres.Ejecutivo responsable")

if tiene_informacion(ejecutivo_ziga):
    estilo_ejecutivo = ""
else:
    estilo_ejecutivo = ' style="color:#A7B1C2;font-style:italic;"'

ejecutivo_html = (
    '<div>'
    '<div class="unidad-serie-label">EJECUTIVO RESPONSABLE</div>'
    f'<div class="unidad-serie"{estilo_ejecutivo}>'
    f'{texto_ejecutivo(ejecutivo_ziga)}</div>'
    '</div>'
)

st.markdown(
    f"""
<div class="unidad-header">
<div class="unidad-placa">{placa}</div>
<div>
<div class="unidad-serie-label">NÚMERO DE SERIE</div>
<div class="unidad-serie">{serie}</div>
</div>
{ejecutivo_html}
</div>
""",
    unsafe_allow_html=True
)


# =========================================================
# REFACCIONES
# =========================================================

grupos_refacciones = []

if not df_refacciones.empty:

    columnas_ref = {
        str(c).strip().lower(): c
        for c in df_refacciones.columns
    }

    col_serie_ref = columnas_ref.get("serie")
    col_folio_ref = columnas_ref.get("folio ws")
    col_actividad_ref = columnas_ref.get("actividad")

    if (
        col_serie_ref is not None
        and col_folio_ref is not None
        and col_actividad_ref is not None
    ):

        ref = df_refacciones.copy()

        ref["_SERIE_BUSQUEDA"] = limpiar_serie(ref[col_serie_ref])

        # Usa la serie limpia de la unidad (no el texto "Sin información")
        serie_actual = unidad.get("_SERIE_KEY", "")

        if serie_actual:

            ref = ref[ref["_SERIE_BUSQUEDA"] == serie_actual].copy()

            for folio, grupo in ref.groupby(col_folio_ref, sort=False):

                if not tiene_informacion(folio):
                    continue

                actividades = []

                for actividad in grupo[col_actividad_ref]:

                    if not tiene_informacion(actividad):
                        continue

                    texto = str(actividad).strip()

                    if not any(
                        texto.casefold() == existente.casefold()
                        for existente in actividades
                    ):
                        actividades.append(texto)

                if actividades:
                    grupos_refacciones.append(
                        (str(folio).strip(), actividades)
                    )


# =========================================================
# RESUMEN DE FOLIOS ABIERTOS
# =========================================================

folios_resumen = [
    ("📋 Gestoría", tiene_informacion(unidad.get("Gestorías1.Folio"))),
    ("🔧 Mantenimiento", tiene_informacion(unidad.get("Mantenimiento.FOLIO"))),
    ("🛠️ Ziga Talleres", tiene_informacion(unidad.get("Ziga_Talleres.Folio"))),
    ("🚨 Siniestro", tiene_informacion(unidad.get("Siniestro"))),
    ("📦 Refacciones", bool(grupos_refacciones)),
]

total_abiertos = sum(1 for _, activo in folios_resumen if activo)

chips_html = "".join(
    f'<span class="chip {"chip-on" if activo else "chip-off"}">'
    f'<span class="chip-punto"></span>{nombre}</span>'
    for nombre, activo in folios_resumen
)

titulo_seccion(f"🚦 Resumen de folios abiertos ({total_abiertos})")

st.markdown(
    f'<div class="resumen">{chips_html}</div>',
    unsafe_allow_html=True
)


# =========================================================
# INFORMACIÓN GENERAL
# =========================================================

titulo_seccion("📋 Información general")

mostrar_tarjetas([
    ("t", "LÍNEA", valor_visible(unidad.get("LINEA"))),
    ("t", "MODELO", modelo_visible(unidad.get("MODELO"))),
    ("b", "ESTATUS", valor_visible(unidad.get("ESTATUS"))),
    ("t", "SUBESTATUS", valor_visible(unidad.get("SUBESTATUS"))),
])


# =========================================================
# DATOS DE LA UNIDAD
# =========================================================

titulo_seccion("🚘 Datos de la unidad")

mostrar_tarjetas([
    ("t", "NOMBRE DEL CLIENTE", valor_visible(unidad.get("NOMBRE DEL CLIENTE"))),
    ("t", "ALMACÉN", valor_visible(unidad.get("ALMACÉN"))),
    ("t", "UNIDAD DE NEGOCIO", valor_visible(unidad.get("UNIDAD NEGOCIO"))),
])


hay_folio_abierto = False


# =========================================================
# GESTORÍA
# =========================================================

if tiene_informacion(unidad.get("Gestorías1.Folio")):

    hay_folio_abierto = True

    separador()

    titulo_seccion(
        "📋 Gestoría",
        dias_numero_o_none(unidad.get("Gestorías1.Personalizado"))
    )

    mostrar_tarjetas([
        ("t", "FOLIO", valor_visible(unidad.get("Gestorías1.Folio"))),
        ("b", "ESTATUS DE PÓLIZA", valor_visible(unidad.get("Estatus poliza"))),
        ("t", "FECHA DE CREACIÓN", fecha_visible(unidad.get("Gestorías1.Fecha de creación"))),
    ])

    mostrar_tarjetas([
        ("b", "ESTADO DE ATENCIÓN", valor_visible(unidad.get("Gestorías1.Estado de atención"))),
        ("t", "DÍAS", numero_entero(unidad.get("Gestorías1.Personalizado"))),
        ("t", "CONCEPTOS", valor_visible(unidad.get("Gestorías1.Conceptos"))),
    ])


# =========================================================
# MANTENIMIENTO
# =========================================================

if tiene_informacion(unidad.get("Mantenimiento.FOLIO")):

    hay_folio_abierto = True

    separador()

    titulo_seccion(
        "🔧 Mantenimiento",
        dias_numero_o_none(unidad.get("Mantenimiento.dias"))
    )

    mostrar_tarjetas([
        ("t", "FOLIO", valor_visible(unidad.get("Mantenimiento.FOLIO"))),
        ("t", "FECHA DE CREACIÓN", fecha_visible(unidad.get("Mantenimiento.FECHA DE CREACION DEL FOLIO"))),
        ("b", "ESTATUS", valor_visible(unidad.get("Mantenimiento.ESTATUS DEL FOLIO"))),
    ])

    mostrar_tarjetas([
        ("t", "PROVEEDOR", valor_visible(unidad.get("Mantenimiento.PROVEEDOR"))),
        ("t", "DÍAS", numero_entero(unidad.get("Mantenimiento.dias"))),
    ])


# =========================================================
# ZIGA TALLERES
# =========================================================

if tiene_informacion(unidad.get("Ziga_Talleres.Folio")):

    hay_folio_abierto = True

    separador()

    titulo_seccion(
        "🛠️ Ziga Talleres",
        dias_numero_o_none(unidad.get("Ziga_Talleres.dias"))
    )

    mostrar_tarjetas([
        ("t", "FOLIO", valor_visible(unidad.get("Ziga_Talleres.Folio"))),
        ("b", "ESTATUS", valor_visible(unidad.get("Ziga_Talleres.Estatus"))),
        ("t", "ÁREA DEL TALLER ACTIVA", valor_visible(unidad.get("Ziga_Talleres.Área del taller activa"))),
    ])

    mostrar_tarjetas([
        ("t", "FECHA DE INGRESO", fecha_visible(unidad.get("Ziga_Talleres.Fecha de ingreso"))),
        ("t", "DÍAS", numero_entero(unidad.get("Ziga_Talleres.dias"))),
        ("t", "EJECUTIVO RESPONSABLE", texto_ejecutivo(obtener(unidad, "Ziga_Talleres.Ejecutivo responsable"))),
    ])

    if not any(
        clave_columna(c) == clave_columna("Ziga_Talleres.Ejecutivo responsable")
        for c in unidad.index
    ):
        st.warning(
            "No se encontró la columna 'Ziga_Talleres.Ejecutivo responsable' "
            "en la hoja concentrado. Columnas de Ziga detectadas: "
            + ", ".join(c for c in unidad.index if "ziga" in clave_columna(c))
        )


# =========================================================
# SINIESTRO
# =========================================================

if tiene_informacion(unidad.get("Siniestro")):

    hay_folio_abierto = True

    separador()

    titulo_seccion(
        "🚨 Siniestro",
        dias_numero_o_none(unidad.get("Siniestros2.dias"))
    )

    mostrar_tarjetas([
        ("t", "CATEGORÍA", valor_visible(unidad.get("Siniestros2.CATEGORIA"))),
        ("t", "ETAPA", valor_visible(unidad.get("Siniestros2.ETAPA"))),
        ("t", "DÍAS", numero_entero(unidad.get("Siniestros2.dias"))),
    ])

    mostrar_tarjetas([
        ("t", "FECHA DE APERTURA", fecha_visible(unidad.get("Siniestros2.FECHA APERTURA"))),
        ("t", "FECHA DEL SINIESTRO", fecha_visible(unidad.get("Siniestros2.FECHA SINIESTRO"))),
    ])


# =========================================================
# REFACCIONES PENDIENTES
# =========================================================

if grupos_refacciones:

    hay_folio_abierto = True

    separador()

    titulo_seccion("📦 Refacciones pendientes")

    for folio, actividades in grupos_refacciones:

        actividades_html = "<br>".join(
            f"• {actividad}" for actividad in actividades
        )

        st.markdown(
            f"""
<div class="refaccion-card">
<div class="refaccion-folio">FOLIO WS: {folio}</div>
<div class="refaccion-actividad">{actividades_html}</div>
</div>
""",
            unsafe_allow_html=True
        )


# =========================================================
# SIN FOLIOS ABIERTOS
# =========================================================

if not hay_folio_abierto:

    st.markdown(
        """
<div class="sin-folios">
✅ Esta unidad no tiene folios abiertos.
</div>
""",
        unsafe_allow_html=True
    )


# =========================================================
# DESCARGAR FICHA
# =========================================================

separador()

# Se omiten las columnas internas de búsqueda
ficha_serie = unidad.drop(
    labels=["_PLACA_KEY", "_SERIE_KEY", "_UNIDAD_KEY"],
    errors="ignore"
).dropna()

ficha_df = pd.DataFrame({
    "CAMPO": ficha_serie.index.astype(str),
    "VALOR": ficha_serie.astype(str).values
})

st.download_button(
    "⬇️ Descargar ficha de la unidad (CSV)",
    data=ficha_df.to_csv(index=False).encode("utf-8-sig"),
    file_name=f"ficha_{placa}.csv",
    mime="text/csv"
)

mostrar_diagnostico()

st.caption("Casanova | Control de Unidades")
