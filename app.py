import re
import unicodedata
import io
import hashlib

import streamlit as st
import pandas as pd
from pathlib import Path
from PIL import Image, ImageEnhance, ImageFilter
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

.badge-verde {{
    background: #E3F6EA;
    color: #167A3E;
}}

.badge-amarillo {{
    background: #FFF3D0;
    color: #8A6100;
}}

.badge-rojo {{
    background: #FDE5E5;
    color: #B42323;
}}

.badge-azul {{
    background: #E3ECF9;
    color: {AZUL_MARINO};
}}

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

.kpi-amarillo .kpi-label {{
    color: #7A879B;
    opacity: 1;
}}

.kpi-amarillo .kpi-numero {{
    color: {AZUL_MARINO};
}}

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

.login-icon {{
    font-size: 50px;
    margin-bottom: 6px;
}}

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

.seccion-head .titulo-seccion {{
    margin: 0;
}}

.dias-pill {{
    padding: 6px 16px;
    border-radius: 999px;
    font-size: 13px;
    font-weight: 800;
}}

.dias-verde {{
    background: #E3F6EA;
    color: #167A3E;
}}

.dias-ambar {{
    background: #FFF3D0;
    color: #8A6100;
}}

.dias-rojo {{
    background: #FDE5E5;
    color: #B42323;
}}

/* ANIMACIÓN */
@keyframes aparecer {{
    from {{ opacity: 0; transform: translateY(8px); }}
    to {{ opacity: 1; transform: translateY(0); }}
}}

.tarjeta,
.refaccion-card,
.unidad-header,
.kpi-azul,
.kpi-verde,
.kpi-ambar,
.kpi-rojo,
.kpi-amarillo {{
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

        usuario = st.text_input(
            "Usuario",
            placeholder="Usuario"
        )

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

        fecha = pd.to_datetime(
            valor,
            errors="coerce"
        )

        if pd.isna(fecha):
            return valor_visible(valor)

        return fecha.strftime("%d/%m/%Y")

    except Exception:
        return valor_visible(valor)


# =========================================================
# NORMALIZACIÓN PARA PLACA / SERIE
# =========================================================

def limpiar_valor(valor):

    if not tiene_informacion(valor):
        return ""

    texto = str(valor)

    texto = texto.replace("\xa0", " ")

    texto = re.sub(
        r"\.0+$",
        "",
        texto.strip()
    )

    # Quita TODOS los espacios y guiones
    texto = re.sub(
        r"[\s\-]+",
        "",
        texto
    )

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

        columnas.append(
            " ".join(nombre.split()).strip()
        )

    df.columns = columnas

    return df


def buscar_columna(df, nombre):

    objetivo = nombre.strip().upper()

    for c in df.columns:

        if str(c).strip().upper() == objetivo:
            return c

    return None


def clave_columna(nombre):

    texto = unicodedata.normalize(
        "NFKD",
        str(nombre)
    )

    texto = "".join(
        c for c in texto
        if not unicodedata.combining(c)
    )

    return re.sub(
        r"[^a-z0-9]",
        "",
        texto.lower()
    )


def obtener(fila, nombre):

    objetivo = clave_columna(nombre)

    for c in fila.index:

        if clave_columna(c) == objetivo:
            return fila[c]

    return None


def texto_ejecutivo(valor):

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

    for nombre in ["SERIE", "PLACA"]:

        col = buscar_columna(
            df,
            nombre
        )

        if col is not None and col != nombre:

            df = df.rename(
                columns={
                    col: nombre
                }
            )

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

    if any(
        p in texto
        for p in [
            "finaliz",
            "complet",
            "cerrad",
            "entreg",
            "listo",
            "vigente",
            "activ"
        ]
    ):

        clase = "badge-verde"

    elif any(
        p in texto
        for p in [
            "cancel",
            "venc",
            "rechaz",
            "detenid"
        ]
    ):

        clase = "badge-rojo"

    elif any(
        p in texto
        for p in [
            "pend",
            "proceso",
            "espera",
            "revis",
            "abiert",
            "tramite",
            "trámite"
        ]
    ):

        clase = "badge-amarillo"

    else:

        clase = "badge-azul"

    return f"""
<div class="tarjeta">
<div class="campo">{campo}</div>
<div class="valor">
<span class="badge {clase}">{valor}</span>
</div>
</div>
"""


def mostrar_tarjetas(items):

    columnas = st.columns(
        len(items)
    )

    for col, (tipo, campo, valor) in zip(
        columnas,
        items
    ):

        with col:

            if tipo == "b":

                html = tarjeta_badge(
                    campo,
                    valor
                )

            else:

                html = tarjeta(
                    campo,
                    valor
                )

            st.markdown(
                html,
                unsafe_allow_html=True
            )


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

        html = (
            f'<div class="titulo-seccion">'
            f'{texto}'
            f'</div>'
        )

    else:

        html = (
            f'<div class="seccion-head">'
            f'<div class="titulo-seccion">'
            f'{texto}'
            f'</div>'
            f'<div class="dias-pill dias-{nivel}">'
            f'⏱️ {dias:,} días'
            f'</div>'
            f'</div>'
        )

    st.markdown(
        html,
        unsafe_allow_html=True
    )


def separador():

    st.markdown(
        '<div class="separador"></div>',
        unsafe_allow_html=True
    )


# =========================================================
# ENCABEZADO
# =========================================================

col1, col2 = st.columns(
    [5, 1],
    vertical_alignment="center"
)

with col1:

    st.markdown(
        """
<div class="topbar">
<div class="topbar-icono">🚗</div>
<div>
<div class="topbar-titulo">
CASANOVA | CONTROL DE UNIDADES
</div>
<div class="topbar-sub">
Consulta individual de unidades
</div>
</div>
</div>
""",
        unsafe_allow_html=True
    )


with col2:

    if st.button(
        "Cerrar sesión",
        use_container_width=True
    ):

        st.session_state["autenticado"] = False
        st.rerun()

    if st.button(
        "🔄 Actualizar datos",
        use_container_width=True
    ):

        st.cache_data.clear()
        st.rerun()


# =========================================================
# CARGAR ARCHIVOS
# =========================================================

try:

    df = cargar_concentrado()

except Exception as e:

    st.error(
        "No se pudo cargar concentrado.xlsx"
    )

    st.code(str(e))

    st.stop()


df_refacciones = cargar_refacciones()


if (
    "SERIE" not in df.columns
    or
    "PLACA" not in df.columns
):

    st.error(
        "No se encontraron las columnas "
        "SERIE y/o PLACA en la hoja "
        f"'{HOJA_BASE}'. Columnas detectadas:"
    )

    st.code(
        ", ".join(
            df.columns.astype(str)
        )
    )

    st.stop()


# =========================================================
# CLAVES LIMPIAS
# =========================================================

df = df.copy()

df["_PLACA_KEY"] = limpiar_serie(
    df["PLACA"]
)

df["_SERIE_KEY"] = limpiar_serie(
    df["SERIE"]
)

df["_UNIDAD_KEY"] = df[
    "_SERIE_KEY"
].where(
    df["_SERIE_KEY"] != "",
    df["_PLACA_KEY"]
)


# =========================================================
# TOTAL DE UNIDADES
# =========================================================

claves_validas = df.loc[
    df["_UNIDAD_KEY"] != "",
    "_UNIDAD_KEY"
]

total_unidades = claves_validas.nunique()


def mostrar_diagnostico():

    filas_totales = len(df)

    sin_serie = int(
        (
            df["_SERIE_KEY"] == ""
        ).sum()
    )

    sin_placa = int(
        (
            df["_PLACA_KEY"] == ""
        ).sum()
    )

    sin_ambas = int(
        (
            (df["_SERIE_KEY"] == "")
            &
            (df["_PLACA_KEY"] == "")
        ).sum()
    )

    series_repetidas = int(
        df.loc[
            df["_SERIE_KEY"] != "",
            "_SERIE_KEY"
        ].duplicated().sum()
    )

    with st.expander(
        "🧪 Diagnóstico de datos"
    ):

        st.write(
            f"Filas en el Excel: "
            f"**{filas_totales:,}**"
        )

        st.write(
            f"Unidades únicas contadas: "
            f"**{total_unidades:,}**"
        )

        st.write(
            f"Filas sin SERIE: "
            f"**{sin_serie:,}**"
        )

        st.write(
            f"Filas sin PLACA: "
            f"**{sin_placa:,}**"
        )

        st.write(
            f"Filas sin SERIE ni PLACA "
            f"(no se cuentan): **{sin_ambas:,}**"
        )

        st.write(
            f"Filas con SERIE repetida: "
            f"**{series_repetidas:,}**"
        )

        if series_repetidas:

            repetidas = df[
                (
                    df["_SERIE_KEY"] != ""
                )
                &
                df["_SERIE_KEY"].duplicated(
                    keep=False
                )
            ][
                ["PLACA", "SERIE"]
            ].sort_values(
                "SERIE"
            )

            st.dataframe(
                repetidas,
                use_container_width=True
            )


# =========================================================
# OCR - NORMALIZAR TEXTO
# =========================================================

def normalizar_texto_ocr(texto):

    if not texto:
        return ""

    texto = str(texto).upper()

    reemplazos = {
        "—": "-",
        "–": "-",
        "−": "-",
        "_": "-",
        "O": "O",
    }

    for viejo, nuevo in reemplazos.items():
        texto = texto.replace(
            viejo,
            nuevo
        )

    texto = re.sub(
        r"[^A-Z0-9\-\s]",
        " ",
        texto
    )

    texto = re.sub(
        r"\s+",
        " ",
        texto
    ).strip()

    return texto


# =========================================================
# OCR - VALIDAR POSIBLE PLACA
# =========================================================

def parece_placa(texto):

    if not texto:
        return False

    limpio = limpiar_valor(texto)

    # Longitud razonable para una placa
    if len(limpio) < 5 or len(limpio) > 9:
        return False

    # Debe contener letras Y números
    tiene_letras = bool(
        re.search(
            r"[A-Z]",
            limpio
        )
    )

    tiene_numeros = bool(
        re.search(
            r"\d",
            limpio
        )
    )

    if not tiene_letras or not tiene_numeros:
        return False

    # Evitar palabras comunes que aparecen
    # en las placas / fotografías
    palabras_prohibidas = {
        "MEXICO",
        "MEXIC",
        "CIUDAD",
        "TRANSPORTE",
        "PRIVADO",
        "AUTOMOVIL",
        "AUTOMOVIL",
        "GOBIERNO",
        "ESTADO",
        "CDMX",
        "DISTRITO",
        "FEDERAL",
        "VEHICULO",
        "VEHICUL",
        "NISSAN",
        "CHEVROLET",
        "CASANOVA",
    }

    if limpio in palabras_prohibidas:
        return False

    return True


# =========================================================
# OCR - EXTRAER POSIBLES PLACAS
# =========================================================

def extraer_candidatos_placa(texto):

    if not texto:
        return []

    texto = normalizar_texto_ocr(texto)

    candidatos = []

    # -----------------------------------------------------
    # 1. Buscar bloques separados por espacios
    # -----------------------------------------------------

    tokens = re.findall(
        r"[A-Z0-9]+(?:-[A-Z0-9]+)*",
        texto
    )

    for token in tokens:

        limpio = limpiar_valor(token)

        if parece_placa(limpio):
            candidatos.append(limpio)

    # -----------------------------------------------------
    # 2. Buscar combinaciones de 2 o 3 tokens consecutivos
    #
    # Ejemplo:
    #
    # MKJ-
    # 121-B
    #
    # puede venir separado por OCR.
    # -----------------------------------------------------

    tokens_limpios = []

    for token in tokens:

        token_limpio = re.sub(
            r"[^A-Z0-9]",
            "",
            token
        )

        if token_limpio:
            tokens_limpios.append(
                token_limpio
            )

    for cantidad in [2, 3]:

        for i in range(
            len(tokens_limpios) - cantidad + 1
        ):

            combinado = "".join(
                tokens_limpios[
                    i:i + cantidad
                ]
            )

            if parece_placa(combinado):

                candidatos.append(
                    combinado
                )

    # -----------------------------------------------------
    # 3. Buscar patrones conocidos de placas
    # pero sin amarrarnos a uno solo.
    # -----------------------------------------------------

    patrones = [

        # ABC123D
        r"\b[A-Z]{3}\d{3}[A-Z]\b",

        # ABC123
        r"\b[A-Z]{3}\d{3}\b",

        # A12BCD
        r"\b[A-Z]\d{2}[A-Z]{3}\b",

        # A12BC
        r"\b[A-Z]\d{2}[A-Z]{2,4}\b",

        # AB123CD
        r"\b[A-Z]{2}\d{3}[A-Z]{1,2}\b",

        # ABC12
        r"\b[A-Z]{3}\d{2,3}\b",

        # 123ABC
        r"\b\d{2,3}[A-Z]{3}\b",

        # Combinaciones generales
        r"\b[A-Z0-9]{6,8}\b",
    ]

    for patron in patrones:

        encontrados = re.findall(
            patron,
            texto
        )

        for encontrado in encontrados:

            limpio = limpiar_valor(
                encontrado
            )

            if parece_placa(limpio):

                candidatos.append(
                    limpio
                )

    # -----------------------------------------------------
    # Quitar duplicados
    # -----------------------------------------------------

    resultado = []

    vistos = set()

    for candidato in candidatos:

        candidato = limpiar_valor(
            candidato
        )

        if not candidato:
            continue

        if candidato in vistos:
            continue

        vistos.add(candidato)

        resultado.append(
            candidato
        )

    return resultado


# =========================================================
# OCR - ELEGIR LA MEJOR PLACA
# =========================================================

def seleccionar_mejor_placa(candidatos):

    if not candidatos:
        return ""

    # Palabras que NO queremos seleccionar
    palabras_malas = {
        "MEXICO",
        "CIUDAD",
        "TRANSPORTE",
        "PRIVADO",
        "AUTOMOVIL",
        "GOBIERNO",
        "VEHICULO",
        "NISSAN",
        "CHEVROLET",
        "CASANOVA",
    }

    candidatos_validos = []

    for candidato in candidatos:

        candidato = limpiar_valor(
            candidato
        )

        if not parece_placa(candidato):
            continue

        if candidato in palabras_malas:
            continue

        # Puntaje
        puntos = 0

        # Longitud habitual
        if 6 <= len(candidato) <= 8:
            puntos += 4

        # Letras y números
        if re.search(r"[A-Z]", candidato):
            puntos += 2

        if re.search(r"\d", candidato):
            puntos += 2

        # Si mezcla letras y números de forma clara
        if (
            re.search(r"[A-Z]", candidato)
            and
            re.search(r"\d", candidato)
        ):
            puntos += 3

        # Penalizar cadenas demasiado largas
        if len(candidato) >= 9:
            puntos -= 2

        # Penalizar cadenas con demasiadas letras
        if (
            sum(c.isalpha() for c in candidato)
            >= 7
        ):
            puntos -= 2

        candidatos_validos.append(
            (
                puntos,
                candidato
            )
        )

    if not candidatos_validos:
        return ""

    candidatos_validos.sort(
        key=lambda x: (
            x[0],
            len(x[1])
        ),
        reverse=True
    )

    return candidatos_validos[0][1]


# =========================================================
# OCR - LEER PLACA DESDE FOTO
# =========================================================

def leer_texto_foto(imagen_bytes):

    try:

        img_original = Image.open(
            io.BytesIO(imagen_bytes)
        ).convert("RGB")

        resultados_ocr = []

        # =================================================
        # GENERAR DIFERENTES VERSIONES DE LA FOTO
        # =================================================

        imagenes = []

        # Original
        imagenes.append(
            img_original
        )

        # Imagen ampliada
        ancho, alto = img_original.size

        if ancho < 1800:

            imagen_ampliada = img_original.resize(
                (
                    ancho * 2,
                    alto * 2
                ),
                Image.Resampling.LANCZOS
            )

        else:

            imagen_ampliada = img_original

        imagenes.append(
            imagen_ampliada
        )

        # Escala de grises
        gris = imagen_ampliada.convert(
            "L"
        )

        imagenes.append(
            gris
        )

        # Más contraste
        contraste = ImageEnhance.Contrast(
            gris
        ).enhance(2.0)

        imagenes.append(
            contraste
        )

        # Más nitidez
        nitida = contraste.filter(
            ImageFilter.SHARPEN
        )

        imagenes.append(
            nitida
        )

        # =================================================
        # CONFIGURACIONES OCR
        # =================================================

        configuraciones = [
            "--psm 6",
            "--psm 11",
            "--psm 12",
            "--psm 7",
        ]

        # =================================================
        # EJECUTAR OCR
        # =================================================

        for imagen in imagenes:

            for config in configuraciones:

                try:

                    texto = pytesseract.image_to_string(
                        imagen,
                        config=(
                            config
                            +
                            " -c "
                            "tessedit_char_whitelist="
                            "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                            "abcdefghijklmnopqrstuvwxyz"
                            "0123456789-"
                        )
                    )

                    if texto:

                        resultados_ocr.append(
                            texto
                        )

                except Exception:
                    continue

        # =================================================
        # EXTRAER CANDIDATOS
        # =================================================

        todos_candidatos = []

        for texto in resultados_ocr:

            candidatos = extraer_candidatos_placa(
                texto
            )

            todos_candidatos.extend(
                candidatos
            )

        # =================================================
        # SELECCIONAR MEJOR RESULTADO
        # =================================================

        placa = seleccionar_mejor_placa(
            todos_candidatos
        )

        if placa:
            return placa

        # =================================================
        # ÚLTIMO INTENTO:
        # ANALIZAR CADA LÍNEA DEL OCR
        # =================================================

        for texto in resultados_ocr:

            for linea in texto.splitlines():

                linea = normalizar_texto_ocr(
                    linea
                )

                if not linea:
                    continue

                # Quitar palabras comunes
                linea_sin_texto = re.sub(
                    r"\b(MEXICO|CIUDAD|TRANSPORTE|PRIVADO|AUTOMOVIL|GOBIERNO)\b",
                    " ",
                    linea
                )

                candidatos = extraer_candidatos_placa(
                    linea_sin_texto
                )

                placa = seleccionar_mejor_placa(
                    candidatos
                )

                if placa:
                    return placa

        return ""

    except Exception:
        return ""


# =========================================================
# ENCABEZADO
# =========================================================

titulo_seccion(
    "🔎 Buscar unidad"
)


# =========================================================
# SESSION STATE
# =========================================================

if "busqueda" not in st.session_state:
    st.session_state["busqueda"] = ""

if "foto_procesada" not in st.session_state:
    st.session_state["foto_procesada"] = ""

if "msg_foto" not in st.session_state:
    st.session_state["msg_foto"] = ("", "")


# =========================================================
# BUSCAR CON FOTO
# =========================================================

foto = None

usar_foto = st.toggle(
    "📷 Buscar con foto (galería o cámara trasera)",
    value=False
)

if usar_foto:

    foto = st.file_uploader(
        "Selecciona una foto o toma una nueva",
        type=[
            "jpg",
            "jpeg",
            "png",
            "webp"
        ]
    )

    st.caption(
        "En el celular, al pulsar el botón "
        "elige 'Galería' o 'Tomar foto' "
        "(cámara trasera)."
    )


# =========================================================
# PROCESAR FOTO
# =========================================================

if foto is not None:

    datos_foto = foto.getvalue()

    hash_foto = hashlib.md5(
        datos_foto
    ).hexdigest()

    if (
        st.session_state["foto_procesada"]
        != hash_foto
    ):

        st.session_state[
            "foto_procesada"
        ] = hash_foto

        with st.spinner(
            "🔎 Leyendo placa / serie..."
        ):

            texto_ocr = leer_texto_foto(
                datos_foto
            )

        if texto_ocr:

            st.session_state[
                "busqueda"
            ] = texto_ocr

            st.session_state[
                "msg_foto"
            ] = (
                "ok",
                f"📷 Placa detectada: "
                f"{texto_ocr} "
                "— revisa que sea correcta."
            )

        else:

            st.session_state[
                "msg_foto"
            ] = (
                "warn",
                "No pude identificar la placa. "
                "Intenta con una foto más cercana "
                "a la placa, de frente y con buena "
                "iluminación."
            )

    tipo_msg, texto_msg = st.session_state[
        "msg_foto"
    ]

    if tipo_msg == "ok":

        st.success(
            texto_msg
        )

    elif tipo_msg == "warn":

        st.warning(
            texto_msg
        )


# =========================================================
# CAMPO DE BÚSQUEDA
# =========================================================

busqueda = st.text_input(
    "Busca por PLACA o SERIE",
    placeholder=(
        "Ejemplo: PDN1295 "
        "o número de serie..."
    ),
    label_visibility="collapsed",
    key="busqueda"
)


# =========================================================
# SI NO HAY BÚSQUEDA
# =========================================================

if not busqueda:

    st.markdown(
        f"""
<div class="kpi-azul" style="margin-top:14px;">
<div class="kpi-label">
🚗 Total de unidades
</div>
<div class="kpi-numero">
{total_unidades:,}
</div>
<div class="kpi-subtotal">
Centenario / Lala
</div>
</div>
""",
        unsafe_allow_html=True
    )

    st.info(
        "Escribe una placa o número de serie "
        "para consultar la unidad."
    )

    mostrar_diagnostico()

    st.stop()


# =========================================================
# BUSCAR
# =========================================================

criterio = normalizar_texto(
    busqueda
)

if criterio == "":

    st.warning(
        "Escribe una placa o serie válida."
    )

    st.stop()


# Exacto primero
exacto = df[
    (
        df["_PLACA_KEY"] == criterio
    )
    |
    (
        df["_SERIE_KEY"] == criterio
    )
]


if not exacto.empty:

    resultado = exacto.copy()

else:

    resultado = df[
        df["_PLACA_KEY"].str.contains(
            criterio,
            regex=False,
            na=False
        )
        |
        df["_SERIE_KEY"].str.contains(
            criterio,
            regex=False,
            na=False
        )
    ].copy()


if resultado.empty:

    st.warning(
        f"No se encontró ninguna unidad "
        f"con: {busqueda}"
    )

    st.stop()


# =========================================================
# SELECCIONAR UNIDAD
# =========================================================

resultado = resultado.reset_index(
    drop=True
)

if len(resultado) > 1:

    st.caption(
        f"Se encontraron "
        f"{len(resultado):,} coincidencias."
    )

    def etiqueta(i):

        fila = resultado.iloc[i]

        return (
            f"{valor_visible(fila.get('PLACA'))} | "
            f"{valor_visible(fila.get('SERIE'))}"
        )

    posicion = st.selectbox(
        "Selecciona la unidad",
        options=list(
            range(len(resultado))
        ),
        format_func=etiqueta
    )

    unidad = resultado.iloc[
        posicion
    ]

else:

    unidad = resultado.iloc[0]


# =========================================================
# DATOS PRINCIPALES
# =========================================================

placa = valor_visible(
    unidad.get("PLACA")
)

serie = valor_visible(
    unidad.get("SERIE")
)


# =========================================================
# DÍAS EN PISO
# =========================================================

dias_piso = unidad.get(
    "DÍAS EN PISO"
)

if tiene_informacion(
    dias_piso
):

    try:

        texto_dias = (
            f"{int(float(dias_piso)):,} días"
        )

    except Exception:

        texto_dias = valor_visible(
            dias_piso
        )

else:

    texto_dias = (
        "🆕 Recién ingresado"
    )


# =========================================================
# KPIs
# =========================================================

col1, col2 = st.columns(2)


with col1:

    st.markdown(
        f"""
<div class="kpi-azul" style="margin-top:14px;">
<div class="kpi-label">
🚗 Total de unidades
</div>
<div class="kpi-numero">
{total_unidades:,}
</div>
<div class="kpi-subtotal">
Centenario / Lala
</div>
</div>
""",
        unsafe_allow_html=True
    )


with col2:

    nivel_piso = nivel_dias(
        dias_numero_o_none(
            dias_piso
        )
    )

    clase_kpi = {
        "verde": "kpi-verde",
        "ambar": "kpi-ambar",
        "rojo": "kpi-rojo",
        None: "kpi-amarillo"
    }[
        nivel_piso
    ]

    texto_nivel = {
        "verde": "✅ Rotación normal",
        "ambar": "⚠️ Requiere atención",
        "rojo": "🚨 Tiempo crítico en piso",
        None: ""
    }[
        nivel_piso
    ]

    sub_html = (
        f'<div class="kpi-sub">'
        f'{texto_nivel}'
        f'</div>'
        if texto_nivel
        else ""
    )

    st.markdown(
        f"""
<div class="{clase_kpi}" style="margin-top:14px;">
<div class="kpi-label">
🏭 Días en piso
</div>
<div class="kpi-numero">
{texto_dias}
</div>
{sub_html}
</div>
""",
        unsafe_allow_html=True
    )


# =========================================================
# IDENTIDAD DE LA UNIDAD
# =========================================================

ejecutivo_ziga = obtener(
    unidad,
    "Ziga_Talleres.Ejecutivo responsable"
)


if tiene_informacion(
    ejecutivo_ziga
):

    estilo_ejecutivo = ""

else:

    estilo_ejecutivo = (
        ' style="color:#A7B1C2;'
        'font-style:italic;"'
    )


ejecutivo_html = (
    '<div>'
    '<div class="unidad-serie-label">'
    'EJECUTIVO RESPONSABLE'
    '</div>'
    f'<div class="unidad-serie"'
    f'{estilo_ejecutivo}>'
    f'{texto_ejecutivo(ejecutivo_ziga)}'
    '</div>'
    '</div>'
)


st.markdown(
    f"""
<div class="unidad-header">
<div class="unidad-placa">
{placa}
</div>

<div>
<div class="unidad-serie-label">
NÚMERO DE SERIE
</div>

<div class="unidad-serie">
{serie}
</div>
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

    col_serie_ref = columnas_ref.get(
        "serie"
    )

    col_folio_ref = columnas_ref.get(
        "folio ws"
    )

    col_actividad_ref = columnas_ref.get(
        "actividad"
    )

    if (
        col_serie_ref is not None
        and
        col_folio_ref is not None
        and
        col_actividad_ref is not None
    ):

        ref = df_refacciones.copy()

        ref["_SERIE_BUSQUEDA"] = limpiar_serie(
            ref[col_serie_ref]
        )

        serie_actual = unidad.get(
            "_SERIE_KEY",
            ""
        )

        if serie_actual:

            ref = ref[
                ref["_SERIE_BUSQUEDA"]
                == serie_actual
            ].copy()

            for folio, grupo in ref.groupby(
                col_folio_ref,
                sort=False
            ):

                if not tiene_informacion(
                    folio
                ):
                    continue

                actividades = []

                for actividad in grupo[
                    col_actividad_ref
                ]:

                    if not tiene_informacion(
                        actividad
                    ):
                        continue

                    texto = str(
                        actividad
                    ).strip()

                    if not any(
                        texto.casefold()
                        == existente.casefold()
                        for existente in actividades
                    ):

                        actividades.append(
                            texto
                        )

                if actividades:

                    grupos_refacciones.append(
                        (
                            str(folio).strip(),
                            actividades
                        )
                    )


# =========================================================
# RESUMEN DE FOLIOS ABIERTOS
# =========================================================

folios_resumen = [

    (
        "📋 Gestoría",
        tiene_informacion(
            unidad.get(
                "Gestorías1.Folio"
            )
        )
    ),

    (
        "🔧 Mantenimiento",
        tiene_informacion(
            unidad.get(
                "Mantenimiento.FOLIO"
            )
        )
    ),

    (
        "🛠️ Ziga Talleres",
        tiene_informacion(
            unidad.get(
                "Ziga_Talleres.Folio"
            )
        )
    ),

    (
        "🚨 Siniestro",
        tiene_informacion(
            unidad.get(
                "Siniestro"
            )
        )
    ),

    (
        "📦 Refacciones",
        bool(
            grupos_refacciones
        )
    ),

]


total_abiertos = sum(
    1
    for _, activo in folios_resumen
    if activo
)


chips_html = "".join(

    f'<span class="chip '
    f'{"chip-on" if activo else "chip-off"}">'
    f'<span class="chip-punto"></span>'
    f'{nombre}'
    f'</span>'

    for nombre, activo
    in folios_resumen
)


titulo_seccion(
    f"🚦 Resumen de folios abiertos "
    f"({total_abiertos})"
)


st.markdown(
    f'<div class="resumen">'
    f'{chips_html}'
    f'</div>',
    unsafe_allow_html=True
)


# =========================================================
# INFORMACIÓN GENERAL
# =========================================================

titulo_seccion(
    "📋 Información general"
)


mostrar_tarjetas([

    (
        "t",
        "LÍNEA",
        valor_visible(
            unidad.get("LINEA")
        )
    ),

    (
        "t",
        "MODELO",
        modelo_visible(
            unidad.get("MODELO")
        )
    ),

    (
        "b",
        "ESTATUS",
        valor_visible(
            unidad.get("ESTATUS")
        )
    ),

    (
        "t",
        "SUBESTATUS",
        valor_visible(
            unidad.get("SUBESTATUS")
        )
])


# =========================================================
# DATOS DE LA UNIDAD
# =========================================================

titulo_seccion(
    "🚘 Datos de la unidad"
)


mostrar_tarjetas([

    (
        "t",
        "NOMBRE DEL CLIENTE",
        valor_visible(
            unidad.get(
                "NOMBRE DEL CLIENTE"
            )
        )
    ),

    (
        "t",
        "ALMACÉN",
        valor_visible(
            unidad.get(
                "ALMACÉN"
            )
        )
    ),

    (
        "t",
        "UNIDAD DE NEGOCIO",
        valor_visible(
            unidad.get(
                "UNIDAD NEGOCIO"
            )
        )
    )

])


hay_folio_abierto = False


# =========================================================
# GESTORÍA
# =========================================================

if tiene_informacion(
    unidad.get(
        "Gestorías1.Folio"
    )
):

    hay_folio_abierto = True

    separador()

    titulo_seccion(
        "📋 Gestoría",
        dias_numero_o_none(
            unidad.get(
                "Gestorías1.Personalizado"
            )
        )
    )

    mostrar_tarjetas([

        (
            "t",
            "FOLIO",
            valor_visible(
                unidad.get(
                    "Gestorías1.Folio"
                )
            )
        ),

        (
            "b",
            "ESTATUS DE PÓLIZA",
            valor_visible(
                unidad.get(
                    "Estatus poliza"
                )
            )
        ),

        (
            "t",
            "FECHA DE CREACIÓN",
            fecha_visible(
                unidad.get(
                    "Gestorías1.Fecha de creación"
                )
            )
        )

    ])

    mostrar_tarjetas([

        (
            "b",
            "ESTADO DE ATENCIÓN",
            valor_visible(
                unidad.get(
                    "Gestorías1.Estado de atención"
                )
            )
        ),

        (
            "t",
            "DÍAS",
            numero_entero(
                unidad.get(
                    "Gestorías1.Personalizado"
                )
            )
        ),

        (
            "t",
            "CONCEPTOS",
            valor_visible(
                unidad.get(
                    "Gestorías1.Conceptos"
                )
            )
        )

    ])


# =========================================================
# MANTENIMIENTO
# =========================================================

if tiene_informacion(
    unidad.get(
        "Mantenimiento.FOLIO"
    )
):

    hay_folio_abierto = True

    separador()

    titulo_seccion(
        "🔧 Mantenimiento",
        dias_numero_o_none(
            unidad.get(
                "Mantenimiento.dias"
            )
        )
    )

    mostrar_tarjetas([

        (
            "t",
            "FOLIO",
            valor_visible(
                unidad.get(
                    "Mantenimiento.FOLIO"
                )
            )
        ),

        (
            "t",
            "FECHA DE CREACIÓN",
            fecha_visible(
                unidad.get(
                    "Mantenimiento.FECHA DE CREACION DEL FOLIO"
                )
            )
        ),

        (
            "b",
            "ESTATUS",
            valor_visible(
                unidad.get(
                    "Mantenimiento.ESTATUS DEL FOLIO"
                )
            )
        )

    ])

    mostrar_tarjetas([

        (
            "t",
            "PROVEEDOR",
            valor_visible(
                unidad.get(
                    "Mantenimiento.PROVEEDOR"
                )
            )
        ),

        (
            "t",
            "DÍAS",
            numero_entero(
                unidad.get(
                    "Mantenimiento.dias"
                )
            )
        )

    ])


# =========================================================
# ZIGA TALLERES
# =========================================================

if tiene_informacion(
    unidad.get(
        "Ziga_Talleres.Folio"
    )
):

    hay_folio_abierto = True

    separador()

    titulo_seccion(
        "🛠️ Ziga Talleres",
        dias_numero_o_none(
            unidad.get(
                "Ziga_Talleres.dias"
            )
        )
    )

    mostrar_tarjetas([

        (
            "t",
            "FOLIO",
            valor_visible(
                unidad.get(
                    "Ziga_Talleres.Folio"
                )
            )
        ),

        (
            "b",
            "ESTATUS",
            valor_visible(
                unidad.get(
                    "Ziga_Talleres.Estatus"
                )
            )
        ),

        (
            "t",
            "ÁREA DEL TALLER ACTIVA",
            valor_visible(
                unidad.get(
                    "Ziga_Talleres.Área del taller activa"
                )
            )
        )

    ])

    mostrar_tarjetas([

        (
            "t",
            "FECHA DE INGRESO",
            fecha_visible(
                unidad.get(
                    "Ziga_Talleres.Fecha de ingreso"
                )
            )
        ),

        (
            "t",
            "DÍAS",
            numero_entero(
                unidad.get(
                    "Ziga_Talleres.dias"
                )
            )
        ),

        (
            "t",
            "EJECUTIVO RESPONSABLE",
            texto_ejecutivo(
                obtener(
                    unidad,
                    "Ziga_Talleres.Ejecutivo responsable"
                )
            )
        )

    ])

    if not any(

        clave_columna(c)
        ==
        clave_columna(
            "Ziga_Talleres.Ejecutivo responsable"
        )

        for c in unidad.index

    ):

        st.warning(
            "No se encontró la columna "
            "'Ziga_Talleres.Ejecutivo responsable' "
            "en la hoja concentrado. Columnas "
            "de Ziga detectadas: "
            +
            ", ".join(
                c
                for c in unidad.index
                if "ziga" in clave_columna(c)
            )
        )


# =========================================================
# SINIESTRO
# =========================================================

if tiene_informacion(
    unidad.get(
        "Siniestro"
    )
):

    hay_folio_abierto = True

    separador()

    titulo_seccion(
        "🚨 Siniestro",
        dias_numero_o_none(
            unidad.get(
                "Siniestros2.dias"
            )
        )
    )

    mostrar_tarjetas([

        (
            "t",
            "CATEGORÍA",
            valor_visible(
                unidad.get(
                    "Siniestros2.CATEGORIA"
                )
            )
        ),

        (
            "t",
            "ETAPA",
            valor_visible(
                unidad.get(
                    "Siniestros2.ETAPA"
                )
            )
        ),

        (
            "t",
            "DÍAS",
            numero_entero(
                unidad.get(
                    "Siniestros2.dias"
                )
            )
        )

    ])

    mostrar_tarjetas([

        (
            "t",
            "FECHA DE APERTURA",
            fecha_visible(
                unidad.get(
                    "Siniestros2.FECHA APERTURA"
                )
            )
        ),

        (
            "t",
            "FECHA DEL SINIESTRO",
            fecha_visible(
                unidad.get(
                    "Siniestros2.FECHA SINIESTRO"
                )
            )
        )

    ])


# =========================================================
# REFACCIONES PENDIENTES
# =========================================================

if grupos_refacciones:

    hay_folio_abierto = True

    separador()

    titulo_seccion(
        "📦 Refacciones pendientes"
    )

    for folio, actividades in grupos_refacciones:

        actividades_html = "<br>".join(
            f"• {actividad}"
            for actividad in actividades
        )

        st.markdown(
            f"""
<div class="refaccion-card">
<div class="refaccion-folio">
FOLIO WS: {folio}
</div>

<div class="refaccion-actividad">
{actividades_html}
</div>

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


ficha_serie = unidad.drop(
    labels=[
        "_PLACA_KEY",
        "_SERIE_KEY",
        "_UNIDAD_KEY"
    ],
    errors="ignore"
).dropna()


ficha_df = pd.DataFrame({

    "CAMPO":
        ficha_serie.index.astype(str),

    "VALOR":
        ficha_serie.astype(str).values

})


st.download_button(

    "⬇️ Descargar ficha de la unidad (CSV)",

    data=ficha_df.to_csv(
        index=False
    ).encode("utf-8-sig"),

    file_name=f"ficha_{placa}.csv",

    mime="text/csv"

)


mostrar_diagnostico()


st.caption(
    "Casanova | Control de Unidades"
)
