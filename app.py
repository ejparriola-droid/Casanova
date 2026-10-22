import streamlit as st
import pandas as pd
import re
import difflib
import io
from pathlib import Path

import requests
from PIL import Image, ImageOps, ImageEnhance, ImageFilter

try:
    import pytesseract
except Exception:
    pytesseract = None


# =========================================================
# CONFIGURACIÓN
# =========================================================

st.set_page_config(
    page_title="Casanova | Control de Unidades - Centenario / Lala",
    page_icon="🚗",
    layout="wide"
)


# =========================================================
# ARCHIVO EXCEL
# =========================================================

ARCHIVO_BASE = Path("concentrado.xlsx")

HOJA_BASE = "concentrado"
HOJA_REFACCIONES = "Refacciones"

TABLA_REFACCIONES = "Pendiente_refacciones"


# =========================================================
# USUARIO Y CONTRASEÑA
# =========================================================

USUARIO_CORRECTO = "Casanova"
PASSWORD_CORRECTA = "Casanova.50"


# =========================================================
# COLORES CASANOVA
# =========================================================

AZUL_MARINO = "#213769"
AZUL_CASANOVA = "#004876"
AMARILLO = "#FFB700"
BLANCO = "#FFFFFF"


# =========================================================
# SEMÁFORO DE DÍAS (ajusta estos números a tu operación)
# =========================================================

DIAS_AMBAR = 30    # desde aquí: amarillo (requiere atención)
DIAS_ROJO = 60     # desde aquí: rojo (crítico)


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

/* Todo el texto en negritas */
.stApp p, .stApp span, .stApp label, .stMarkdown {{
    font-weight: 600;
}}

/* Ocultar elementos de Streamlit */
#MainMenu {{visibility: hidden;}}
footer {{visibility: hidden;}}
header {{visibility: hidden;}}

.block-container {{
    padding-top: 1.5rem;
    padding-bottom: 3rem;
    max-width: 1250px;
}}

/* =================================================
   ENCABEZADO
   ================================================= */

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

/* =================================================
   SECCIONES
   ================================================= */

.titulo-seccion {{
    color: {AZUL_MARINO};
    font-size: 18px;
    font-weight: 800;
    margin: 26px 0 12px 0;
    padding-left: 12px;
    border-left: 5px solid {AMARILLO};
    line-height: 1.2;
}}

/* =================================================
   TARJETAS DE DATOS
   ================================================= */

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

/* =================================================
   BADGES DE ESTATUS
   ================================================= */

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

/* =================================================
   KPIs
   ================================================= */

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

/* =================================================
   IDENTIDAD DE LA UNIDAD
   ================================================= */

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

/* =================================================
   REFACCIONES
   ================================================= */

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

/* =================================================
   SIN FOLIOS
   ================================================= */

.sin-folios {{
    background: #E3F6EA;
    border: 1px solid #BFE6CD;
    border-radius: 12px;
    padding: 15px 18px;
    color: #167A3E;
    font-weight: 800;
    margin-top: 24px;
}}

/* =================================================
   SEPARADOR
   ================================================= */

.separador {{
    height: 1px;
    background: linear-gradient(90deg, transparent, #CBD6E8, transparent);
    margin: 30px 0 6px 0;
}}

/* =================================================
   INPUTS Y BOTONES
   ================================================= */

/* Etiquetas (Usuario, Contraseña, etc.) */
.stTextInput label,
.stTextInput label p,
.stSelectbox label,
.stSelectbox label p {{
    color: {AZUL_MARINO} !important;
    font-weight: 800 !important;
    font-size: 15px !important;
}}

/* Contenedor del input (incluye el ojito) */
.stTextInput div[data-baseweb="input"],
.stTextInput div[data-baseweb="base-input"] {{
    background: white !important;
    border-radius: 12px !important;
}}

/* Texto que escribe el usuario */
.stTextInput input {{
    color: #172033 !important;
    -webkit-text-fill-color: #172033 !important;
    caret-color: {AZUL_MARINO} !important;
    background: white !important;
    border-radius: 12px !important;
    border: 2px solid #D8E1EF !important;
    padding: 12px 16px !important;
    font-size: 16px !important;
    font-weight: 700 !important;
}}

/* Placeholder */
.stTextInput input::placeholder {{
    color: #64748B !important;
    -webkit-text-fill-color: #64748B !important;
    font-weight: 600 !important;
    opacity: 1 !important;
}}

.stTextInput input:focus {{
    border-color: {AZUL_CASANOVA} !important;
    box-shadow: 0 0 0 3px rgba(0,72,118,0.12) !important;
}}

/* Botón del ojito (mostrar contraseña) */
.stTextInput button {{
    background: white !important;
    color: {AZUL_MARINO} !important;
    border: none !important;
}}

.stTextInput button svg {{
    fill: {AZUL_MARINO} !important;
    color: {AZUL_MARINO} !important;
}}

/* Selectbox */
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

/* Lista desplegable del selectbox */
div[data-baseweb="popover"] ul,
div[data-baseweb="popover"] li {{
    background: white !important;
    color: #172033 !important;
    font-weight: 600 !important;
}}

div[data-baseweb="popover"] li:hover {{
    background: #EEF3FA !important;
}}

/* Botones */
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

/* =================================================
   LOGIN
   ================================================= */

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

/* =================================================
   SEMÁFORO KPI
   ================================================= */

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

/* =================================================
   RESUMEN DE FOLIOS (CHIPS)
   ================================================= */

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

/* =================================================
   ENCABEZADO DE SECCIÓN CON DÍAS
   ================================================= */

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

.dias-verde {{ background: #E3F6EA; color: #167A3E; }}
.dias-ambar {{ background: #FFF3D0; color: #8A6100; }}
.dias-rojo {{ background: #FDE5E5; color: #B42323; }}

/* =================================================
   ANIMACIÓN SUAVE
   ================================================= */

@keyframes aparecer {{
    from {{ opacity: 0; transform: translateY(8px); }}
    to {{ opacity: 1; transform: translateY(0); }}
}}

.tarjeta, .refaccion-card, .unidad-header,
.kpi-azul, .kpi-verde, .kpi-ambar, .kpi-rojo, .kpi-amarillo {{
    animation: aparecer 0.35s ease both;
}}

/* Botón de descarga */
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

/* =================================================
   LECTOR DE PLACAS
   ================================================= */

.ocr-card {{
    background: white;
    border: 1px solid #E1E8F3;
    border-left: 5px solid {AMARILLO};
    border-radius: 14px;
    padding: 16px 18px;
    margin: 12px 0 18px 0;
    box-shadow: 0 2px 8px rgba(33,55,105,0.05);
}}

.ocr-result {{
    background: #EEF3FA;
    border-radius: 12px;
    padding: 14px 16px;
    margin-top: 10px;
    color: {AZUL_MARINO};
    font-size: 20px;
    font-weight: 800;
    letter-spacing: 2px;
    text-align: center;
}}

/* =================================================
   CELULAR
   ================================================= */

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

                st.error(
                    "Usuario o contraseña incorrectos."
                )


# =========================================================
# CONTROL DE SESIÓN
# =========================================================

if "autenticado" not in st.session_state:

    st.session_state["autenticado"] = False


if not st.session_state["autenticado"]:

    mostrar_login()

    st.stop()


# =========================================================
# NORMALIZAR COLUMNAS
# =========================================================

def normalizar_columnas(df):

    df = df.copy()

    columnas = []

    for columna in df.columns:

        nombre = str(columna)

        nombre = (
            nombre
            .replace("\n", " ")
            .replace("\r", " ")
            .replace("\xa0", " ")
        )

        nombre = " ".join(
            nombre.split()
        )

        columnas.append(
            nombre.strip()
        )

    df.columns = columnas

    return df


# =========================================================
# CARGAR CONCENTRADO
# =========================================================

@st.cache_data
def cargar_concentrado():

    df = pd.read_excel(
        ARCHIVO_BASE,
        sheet_name=HOJA_BASE
    )

    return normalizar_columnas(df)


# =========================================================
# CARGAR REFACCIONES
# =========================================================

@st.cache_data
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

    if texto.lower() in [
        "nan",
        "none",
        "null",
        "nat"
    ]:

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

        return fecha.strftime(
            "%d/%m/%Y"
        )

    except Exception:

        return valor_visible(valor)


def normalizar_texto(valor):

    if not tiene_informacion(valor):

        return ""

    return (
        str(valor)
        .strip()
        .upper()
    )


def normalizar_busqueda(valor):

    if not tiene_informacion(valor):

        return ""

    # Permite encontrar placas aunque vengan con espacios, guiones,
    # puntos u otros separadores. No depende de una placa específica.
    return re.sub(
        r"[^A-Z0-9]",
        "",
        str(valor).upper()
    )


# =========================================================
# LECTOR DE PLACAS CON PLATE RECOGNIZER
# =========================================================

PLATE_RECOGNIZER_URL = "https://api.platerecognizer.com/v1/plate-reader/"


def obtener_token_plate_recognizer():
    """Lee el token desde Streamlit Secrets. Nunca se guarda en GitHub."""
    try:
        token = st.secrets["PLATE_RECOGNIZER_TOKEN"]
        return str(token).strip()
    except Exception:
        return ""


def preparar_imagen_para_api(archivo):
    """
    Abre la foto y la comprime si es necesario.
    Plate Recognizer Cloud admite imágenes de hasta 3 MB.
    """
    try:
        imagen = Image.open(archivo)
        imagen = ImageOps.exif_transpose(imagen).convert("RGB")

        # Evitamos mandar fotografías enormes desde el celular.
        max_lado = 2200
        ancho, alto = imagen.size
        if max(ancho, alto) > max_lado:
            escala = max_lado / max(ancho, alto)
            imagen = imagen.resize(
                (int(ancho * escala), int(alto * escala)),
                Image.Resampling.LANCZOS
            )

        salida = io.BytesIO()
        calidad = 92
        imagen.save(salida, format="JPEG", quality=calidad, optimize=True)

        # Si todavía pesa demasiado, reducimos calidad/tamaño.
        while salida.tell() > 2_800_000 and calidad > 55:
            calidad -= 8
            salida = io.BytesIO()
            imagen.save(
                salida,
                format="JPEG",
                quality=calidad,
                optimize=True
            )

        salida.seek(0)
        return salida.read(), "image/jpeg"

    except Exception as e:
        raise ValueError(f"No se pudo preparar la imagen: {e}")


def distancia_edicion(a, b):
    """Distancia Levenshtein para comparar OCR contra el concentrado."""
    if a == b:
        return 0
    if not a:
        return len(b)
    if not b:
        return len(a)

    anterior = list(range(len(b) + 1))

    for i, ca in enumerate(a, 1):
        actual = [i]
        for j, cb in enumerate(b, 1):
            actual.append(min(
                actual[-1] + 1,
                anterior[j] + 1,
                anterior[j - 1] + (ca != cb)
            ))
        anterior = actual

    return anterior[-1]


def similitud_placa(a, b):
    a = normalizar_busqueda(a)
    b = normalizar_busqueda(b)

    if not a or not b:
        return 0.0

    ratio = difflib.SequenceMatcher(None, a, b).ratio()
    distancia = distancia_edicion(a, b)
    max_len = max(len(a), len(b))
    score_dist = 1 - (distancia / max_len)

    return (ratio * 0.55) + (score_dist * 0.45)


def obtener_placas_reales(df):
    """Obtiene las placas válidas que existen en concentrado.xlsx."""
    if "PLACA" not in df.columns:
        return []

    placas = []

    for valor in df["PLACA"].dropna().astype(str):
        limpia = normalizar_busqueda(valor)

        if 5 <= len(limpia) <= 8 and limpia not in placas:
            placas.append(limpia)

    return placas


def buscar_mejor_placa_concentrado(placa_ocr, df, score_api=0.0):
    """
    Después de que Plate Recognizer lee la placa,
    la valida contra las placas reales del concentrado.
    Esto ayuda con errores como B/8, D/O, I/1, etc.
    """
    placa_ocr = normalizar_busqueda(placa_ocr)

    if not placa_ocr:
        return None, 0.0, ""

    reales = obtener_placas_reales(df)

    if placa_ocr in reales:
        return placa_ocr, 1.0, "Coincidencia exacta con concentrado"

    mejores = []

    for real in reales:
        diferencia_longitud = abs(len(placa_ocr) - len(real))
        distancia = distancia_edicion(placa_ocr, real)

        if diferencia_longitud > 1 or distancia > 2:
            continue

        score = similitud_placa(placa_ocr, real)
        mejores.append((score, real, distancia))

    if not mejores:
        return placa_ocr, score_api, "Lectura de Plate Recognizer"

    mejores.sort(reverse=True)
    score, real, distancia = mejores[0]

    # Solo corregimos automáticamente cuando la coincidencia es fuerte.
    if score >= 0.80:
        return real, score, "Coincidencia corregida contra concentrado"

    return placa_ocr, score_api, "Lectura de Plate Recognizer"


def leer_placa_plate_recognizer(archivo, df):
    """Envía una fotografía a Plate Recognizer y devuelve las mejores placas."""

    token = obtener_token_plate_recognizer()

    if not token:
        return [], (
            "No está configurado PLATE_RECOGNIZER_TOKEN. "
            "Agrega el token en Streamlit Cloud → Settings → Secrets."
        )

    try:
        imagen_bytes, mime = preparar_imagen_para_api(archivo)
    except Exception as e:
        return [], str(e)

    nombre = getattr(archivo, "name", "placa.jpg")

    headers = {
        "Authorization": f"Token {token}"
    }

    files = {
        "upload": (nombre, imagen_bytes, mime)
    }

    # MX hace que el motor use el patrón de placas de México,
    # ayudando a distinguir letras/números parecidos.
    data = {
        "regions": "mx"
    }

    try:
        respuesta = requests.post(
            PLATE_RECOGNIZER_URL,
            headers=headers,
            files=files,
            data=data,
            timeout=45
        )
    except requests.RequestException as e:
        return [], f"No se pudo conectar con Plate Recognizer: {e}"

    if respuesta.status_code == 401 or respuesta.status_code == 403:
        return [], "El token de Plate Recognizer no es válido o no tiene acceso."

    if respuesta.status_code == 429:
        return [], "Se alcanzó el límite de consultas de Plate Recognizer."

    if respuesta.status_code != 200:
        detalle = respuesta.text[:500]
        return [], f"Plate Recognizer respondió {respuesta.status_code}: {detalle}"

    try:
        datos = respuesta.json()
    except Exception:
        return [], "Plate Recognizer no devolvió una respuesta JSON válida."

    resultados = datos.get("results", []) or []

    if not resultados:
        return [], "No se detectó ninguna placa en la fotografía."

    encontrados = []

    for resultado in resultados:
        placa_api = normalizar_busqueda(resultado.get("plate", ""))

        if not placa_api:
            continue

        score_api = float(resultado.get("score", 0) or 0)

        placa_final, score_final, metodo = buscar_mejor_placa_concentrado(
            placa_api,
            df,
            score_api
        )

        encontrados.append({
            "placa": placa_final,
            "placa_ocr": placa_api,
            "score_api": score_api,
            "score_final": score_final,
            "metodo": metodo,
            "candidatos": resultado.get("candidates", []) or []
        })

    # Ordenamos primero por coincidencia contra el concentrado.
    encontrados.sort(
        key=lambda x: x["score_final"],
        reverse=True
    )

    return encontrados, ""


def buscar_placa_en_concentrado(df, placa):
    objetivo = normalizar_busqueda(placa)

    if not objetivo or "PLACA" not in df.columns:
        return pd.DataFrame()

    placas_df = df["PLACA"].apply(normalizar_busqueda)

    return df[placas_df == objetivo].copy()


def obtener_series_reales(df):
    """Obtiene las series/VIN válidas que existen en concentrado.xlsx."""
    if "SERIE" not in df.columns:
        return []

    series = []
    for valor in df["SERIE"].dropna().astype(str):
        limpia = normalizar_busqueda(valor)
        if len(limpia) >= 8 and limpia not in series:
            series.append(limpia)
    return series


def buscar_mejor_serie_concentrado(texto_ocr, df):
    """
    Busca una serie/VIN leída desde una fotografía contra las series reales
    del concentrado. Primero intenta coincidencia exacta y después una
    coincidencia tolerante para errores típicos de OCR.
    """
    if not texto_ocr or "SERIE" not in df.columns:
        return None, 0.0, ""

    texto_normalizado = normalizar_busqueda(texto_ocr)
    if not texto_normalizado:
        return None, 0.0, ""

    series_reales = obtener_series_reales(df)
    if not series_reales:
        return None, 0.0, ""

    # 1) Coincidencia exacta: la más confiable.
    for serie in series_reales:
        if serie in texto_normalizado:
            return serie, 1.0, "Coincidencia exacta de VIN/serie por OCR"

    # 2) Candidatos de 17 caracteres, que es la longitud habitual de un VIN.
    candidatos = []
    if len(texto_normalizado) >= 15:
        for longitud in range(15, 19):
            if len(texto_normalizado) < longitud:
                continue
            for i in range(len(texto_normalizado) - longitud + 1):
                candidatos.append(texto_normalizado[i:i + longitud])

    # 3) También usamos tokens largos por si OCR separó el VIN/serie.
    for token in re.findall(r"[A-Z0-9]{8,25}", str(texto_ocr).upper()):
        limpio = normalizar_busqueda(token)
        if limpio:
            candidatos.append(limpio)

    # Evita comparar cientos de veces el mismo candidato.
    candidatos = list(dict.fromkeys(candidatos))

    mejor = None
    for candidato in candidatos:
        for serie_real in series_reales:
            diferencia_longitud = abs(len(candidato) - len(serie_real))
            if diferencia_longitud > 3:
                continue

            ratio = difflib.SequenceMatcher(None, candidato, serie_real).ratio()
            distancia = distancia_edicion(candidato, serie_real)
            max_len = max(len(candidato), len(serie_real))
            score_dist = 1 - (distancia / max_len) if max_len else 0
            score = (ratio * 0.55) + (score_dist * 0.45)

            if mejor is None or score > mejor[0]:
                mejor = (score, serie_real)

    if mejor and mejor[0] >= 0.78:
        return mejor[1], mejor[0], "Coincidencia corregida contra concentrado por OCR"

    return None, 0.0, ""


def leer_serie_desde_imagen(archivo, df):
    """
    OCR general para fotografías donde aparece el VIN/número de serie
    en lugar de una placa. No sustituye Plate Recognizer; funciona como
    segunda vía para que una misma foto pueda localizar la unidad por placa
    o por VIN/serie.
    """
    if pytesseract is None:
        return None, 0.0, "OCR de VIN/serie no disponible en este entorno."

    try:
        imagen = Image.open(archivo)
        imagen = ImageOps.exif_transpose(imagen).convert("RGB")

        # Reducimos imágenes gigantes para no gastar CPU innecesariamente.
        max_lado = 2600
        ancho, alto = imagen.size
        if max(ancho, alto) > max_lado:
            escala = max_lado / max(ancho, alto)
            imagen = imagen.resize(
                (int(ancho * escala), int(alto * escala)),
                Image.Resampling.LANCZOS
            )

        gris = ImageOps.grayscale(imagen)
        gris = ImageEnhance.Contrast(gris).enhance(1.8)
        gris = gris.filter(ImageFilter.SHARPEN)

        textos = []
        configuraciones = [
            "--psm 6 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789",
            "--psm 11 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789",
        ]

        for config in configuraciones:
            try:
                texto = pytesseract.image_to_string(gris, config=config)
                if texto:
                    textos.append(texto)
            except Exception:
                continue

        texto_completo = "\n".join(textos)
        serie, score, metodo = buscar_mejor_serie_concentrado(
            texto_completo,
            df
        )

        if not serie:
            return None, 0.0, "No se detectó un VIN/serie válido en la fotografía."

        return serie, score, metodo

    except Exception as e:
        return None, 0.0, f"No se pudo leer el VIN/serie: {e}"


def buscar_serie_en_concentrado(df, serie):
    objetivo = normalizar_busqueda(serie)

    if not objetivo or "SERIE" not in df.columns:
        return pd.DataFrame()

    series_df = df["SERIE"].apply(normalizar_busqueda)

    return df[series_df == objetivo].copy()


def mostrar_lector_placas(df):
    titulo_seccion("📷 Leer placa o VIN / serie desde una foto")

    st.markdown(
        """
        <div class="ocr-card">
        <b>Lectura automática de placa o VIN / serie</b><br>
        La fotografía se procesa para intentar identificar la unidad por su
        <b>placa</b> o por su <b>VIN / número de serie</b>.
        <br><small>Si la placa no se detecta, se intenta leer el VIN / serie como respaldo.</small>
        </div>
        """,
        unsafe_allow_html=True
    )

    col_foto, col_cam = st.columns(2)

    with col_foto:
        fotos = st.file_uploader(
            "Selecciona foto(s) de placa o VIN / serie",
            type=["jpg", "jpeg", "png", "webp"],
            accept_multiple_files=True,
            key="fotos_placas"
        )

    with col_cam:
        activar_camara = st.checkbox(
            "📷 Activar cámara",
            key="activar_camara_placas",
            help="La cámara permanece apagada hasta que actives esta opción."
        )

        foto_camara = None

        if activar_camara:
            foto_camara = st.camera_input(
                "Tomar foto de placa o VIN / serie",
                key="foto_placa_camara"
            )

    archivos = list(fotos) if fotos else []

    if foto_camara is not None:
        archivos.append(foto_camara)

    if not archivos:
        return

    if st.button(
        "🔎 Identificar unidad",
        key="btn_leer_placas",
        type="primary",
        use_container_width=True
    ):
        todos = []

        for archivo in archivos:
            # Primero intentamos placa con Plate Recognizer si hay token.
            resultados_placa = []
            error_placa = ""

            if obtener_token_plate_recognizer():
                with st.spinner(f"Analizando {archivo.name}..."):
                    resultados_placa, error_placa = leer_placa_plate_recognizer(
                        archivo,
                        df
                    )

            if resultados_placa:
                for resultado in resultados_placa:
                    placa = resultado["placa"]

                    if not placa:
                        continue

                    if not any(
                        x.get("tipo") == "PLACA" and x.get("valor") == placa
                        for x in todos
                    ):
                        resultado["foto"] = archivo.name
                        resultado["tipo"] = "PLACA"
                        resultado["valor"] = placa
                        todos.append(resultado)

            # También intentamos VIN/serie. Así una foto puede localizar la
            # unidad aunque no haya placa o Plate Recognizer no la detecte.
            with st.spinner(f"Buscando VIN / serie en {archivo.name}..."):
                serie, score_serie, metodo_serie = leer_serie_desde_imagen(
                    archivo,
                    df
                )

            if serie:
                if not any(
                    x.get("tipo") == "SERIE" and x.get("valor") == serie
                    for x in todos
                ):
                    todos.append({
                        "tipo": "SERIE",
                        "valor": serie,
                        "serie": serie,
                        "score_final": score_serie,
                        "metodo": metodo_serie,
                        "foto": archivo.name
                    })
            elif not resultados_placa and error_placa:
                # Solo mostramos el error de Plate Recognizer si tampoco
                # conseguimos localizar la unidad por VIN/serie.
                st.warning(f"{archivo.name}: {error_placa}")

        st.session_state["placas_ocr"] = todos

    resultados_ocr = st.session_state.get("placas_ocr", [])

    if not resultados_ocr:
        return

    st.markdown(
        '<div class="ocr-result">Unidades detectadas</div>',
        unsafe_allow_html=True
    )

    for i, resultado in enumerate(resultados_ocr):
        tipo = resultado.get("tipo", "PLACA")
        valor_detectado = normalizar_busqueda(
            resultado.get("valor", "")
        )
        nombre_foto = resultado.get("foto", "")

        if tipo == "PLACA":
            encontrados = buscar_placa_en_concentrado(
                df,
                valor_detectado
            )

            score_api = resultado.get("score_api", 0.0)
            score_final = resultado.get("score_final", 0.0)
            placa_original = resultado.get("placa_ocr", valor_detectado)

            col1, col2, col3 = st.columns([2.2, 2.2, 1.4])

            with col1:
                st.markdown(f"**🚘 PLACA: {valor_detectado}**")
                st.caption(nombre_foto)
                st.caption(
                    f"Lectura: {placa_original} · "
                    f"confianza API: {score_api:.0%}"
                )

            with col2:
                if not encontrados.empty:
                    st.success("✅ Encontrada en concentrado")
                    if placa_original != valor_detectado:
                        st.caption(
                            f"Se ajustó automáticamente contra el concentrado "
                            f"({score_final:.0%})."
                        )
                else:
                    st.warning("⚠️ No encontrada en concentrado")

            with col3:
                if not encontrados.empty and st.button(
                    "Usar placa",
                    key=f"usar_ocr_{i}",
                    use_container_width=True
                ):
                    st.session_state["busqueda_placa"] = valor_detectado
                    st.session_state["buscador_unidad"] = valor_detectado
                    st.rerun()

        else:
            encontrados = buscar_serie_en_concentrado(
                df,
                valor_detectado
            )

            score_final = resultado.get("score_final", 0.0)
            metodo = resultado.get("metodo", "OCR")

            col1, col2, col3 = st.columns([2.2, 2.2, 1.4])

            with col1:
                st.markdown(f"**🔑 VIN / SERIE: {valor_detectado}**")
                st.caption(nombre_foto)
                st.caption(f"{metodo} · confianza: {score_final:.0%}")

            with col2:
                if not encontrados.empty:
                    st.success("✅ Encontrada en concentrado")
                else:
                    st.warning("⚠️ No encontrada en concentrado")

            with col3:
                if not encontrados.empty and st.button(
                    "Usar VIN / serie",
                    key=f"usar_serie_ocr_{i}",
                    use_container_width=True
                ):
                    st.session_state["busqueda_placa"] = valor_detectado
                    st.session_state["buscador_unidad"] = valor_detectado
                    st.rerun()


# ---------------------------------------------------------
# TARJETAS
# ---------------------------------------------------------

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
        "finaliz", "complet", "cerrad",
        "entreg", "listo", "vigente", "activ"
    ]):

        clase = "badge-verde"

    elif any(p in texto for p in [
        "cancel", "venc", "rechaz", "detenid"
    ]):

        clase = "badge-rojo"

    elif any(p in texto for p in [
        "pend", "proceso", "espera",
        "revis", "abiert", "tramite", "trámite"
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
    """
    items = lista de (tipo, campo, valor)
    tipo: "t" = tarjeta normal, "b" = tarjeta con badge
    Cada llamada dibuja una fila de columnas.
    """

    columnas = st.columns(len(items))

    for col, (tipo, campo, valor) in zip(columnas, items):

        with col:

            if tipo == "b":

                html = tarjeta_badge(campo, valor)

            else:

                html = tarjeta(campo, valor)

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

        html = f'<div class="titulo-seccion">{texto}</div>'

    else:

        html = (
            f'<div class="seccion-head">'
            f'<div class="titulo-seccion">{texto}</div>'
            f'<div class="dias-pill dias-{nivel}">⏱ {dias:,} días</div>'
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
<div class="topbar-titulo">CASANOVA | CONTROL DE UNIDADES</div>
<div class="topbar-sub">Consulta individual de unidades</div>
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


# =========================================================
# LECTOR DE PLACAS
# =========================================================
mostrar_lector_placas(df)


# =========================================================
# TOTAL DE UNIDADES
# =========================================================

if "SERIE" in df.columns:

    series_validas = (
        df["SERIE"]
        .astype("string")
        .fillna("")
        .str.strip()
        .str.upper()
    )

    series_validas = series_validas[
        series_validas != ""
    ]

    total_unidades = (
        series_validas
        .nunique()
    )

else:

    total_unidades = 0


# =========================================================
# BUSCADOR
# =========================================================

if "buscador_unidad" not in st.session_state and st.session_state.get("busqueda_placa"):
    st.session_state["buscador_unidad"] = st.session_state["busqueda_placa"]

titulo_seccion("🔎 Buscar unidad")

busqueda = st.text_input(
    "Busca por PLACA o SERIE",
    placeholder="Ejemplo: PDN1295 o número de serie...",
    label_visibility="collapsed",
    key="buscador_unidad"
)

st.session_state["busqueda_placa"] = busqueda


# =========================================================
# SI NO HAY BÚSQUEDA
# =========================================================

if not busqueda:

    st.markdown(
        f"""
<div class="kpi-azul" style="margin-top:14px;">
<div class="kpi-label">🚗 Total de unidades</div>
<div class="kpi-numero">{total_unidades:,}</div>
</div>
""",
        unsafe_allow_html=True
    )

    st.info(
        "Escribe una placa o número de serie para consultar la unidad."
    )

    st.stop()


# =========================================================
# PREPARAR BÚSQUEDA
# =========================================================

criterio = normalizar_busqueda(busqueda)


placas = (
    df["PLACA"]
    .apply(normalizar_busqueda)
)


series = (
    df["SERIE"]
    .apply(normalizar_busqueda)
)


# =========================================================
# BUSCAR
# =========================================================

resultado = df[
    placas.str.contains(
        criterio,
        case=False,
        na=False,
        regex=False
    )
    |
    series.str.contains(
        criterio,
        case=False,
        na=False,
        regex=False
    )
].copy()


# =========================================================
# SIN RESULTADOS
# =========================================================

if resultado.empty:

    st.warning(
        f"No se encontró ninguna unidad con: {busqueda}"
    )

    st.stop()


# =========================================================
# SELECCIONAR UNIDAD
# =========================================================

if len(resultado) > 1:

    opciones = []

    for _, fila in resultado.iterrows():

        placa_opcion = valor_visible(
            fila.get("PLACA")
        )

        serie_opcion = valor_visible(
            fila.get("SERIE")
        )

        opciones.append(
            f"{placa_opcion} | {serie_opcion}"
        )

    seleccion = st.selectbox(
        "Selecciona la unidad",
        opciones
    )

    posicion = opciones.index(
        seleccion
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


if tiene_informacion(dias_piso):

    try:

        dias_numero = int(
            float(dias_piso)
        )

        texto_dias = (
            f"{dias_numero:,} días"
        )

    except Exception:

        texto_dias = valor_visible(
            dias_piso
        )

else:

    texto_dias = "🆕 Recién ingresado"


# =========================================================
# KPIs: TOTAL + DÍAS EN PISO
# =========================================================

col1, col2 = st.columns(2)

with col1:

    st.markdown(
        f"""
<div class="kpi-azul" style="margin-top:14px;">
<div class="kpi-label">🚗 Total de unidades</div>
<div class="kpi-numero">{total_unidades:,}</div>
</div>
""",
        unsafe_allow_html=True
    )

with col2:

    nivel_piso = nivel_dias(
        dias_numero_o_none(dias_piso)
    )

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

st.markdown(
    f"""
<div class="unidad-header">
<div class="unidad-placa">{placa}</div>
<div>
<div class="unidad-serie-label">NÚMERO DE SERIE</div>
<div class="unidad-serie">{serie}</div>
</div>
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

        ref["_SERIE_BUSQUEDA"] = (
            ref[col_serie_ref]
            .astype("string")
            .fillna("")
            .str.strip()
            .str.upper()
        )

        serie_actual = normalizar_texto(serie)

        ref = ref[
            ref["_SERIE_BUSQUEDA"] == serie_actual
        ].copy()

        for folio, grupo in ref.groupby(
            col_folio_ref,
            sort=False
        ):

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
# RESUMEN DE FOLIOS ABIERTOS (CHIPS)
# =========================================================

folios_resumen = [
    ("📋 Gestoría", tiene_informacion(unidad.get("Gestorías1.Folio"))),
    ("🔧 Mantenimiento", tiene_informacion(unidad.get("Mantenimiento.FOLIO"))),
    ("🛠️ Ziga Talleres", tiene_informacion(unidad.get("Ziga_Talleres.Folio"))),
    ("🚨 Siniestro", tiene_informacion(unidad.get("Siniestro"))),
    ("📦 Refacciones", bool(grupos_refacciones)),
]

total_abiertos = sum(
    1 for _, activo in folios_resumen if activo
)

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

mostrar_tarjetas([
    ("t", "GPS", valor_visible(unidad.get("GPS"))),
    ("t", "PLATAFORMA", valor_visible(unidad.get("PLATAFORMA"))),
    ("t", "¿REPORTA?", valor_visible(unidad.get("¿REPORTA?"))),
])


# =========================================================
# CONTROL DE FOLIOS ABIERTOS
# =========================================================

hay_folio_abierto = False


# =========================================================
# GESTORÍA
# =========================================================

if tiene_informacion(unidad.get("Gestorías1.Folio")):

    hay_folio_abierto = True

    separador()

    titulo_seccion("📋 Gestoría", dias_numero_o_none(unidad.get("Gestorías1.Personalizado")))

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

    titulo_seccion("🔧 Mantenimiento", dias_numero_o_none(unidad.get("Mantenimiento.dias")))

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

    titulo_seccion("🛠️ Ziga Talleres", dias_numero_o_none(unidad.get("Ziga_Talleres.dias")))

    mostrar_tarjetas([
        ("t", "FOLIO", valor_visible(unidad.get("Ziga_Talleres.Folio"))),
        ("b", "ESTATUS", valor_visible(unidad.get("Ziga_Talleres.Estatus"))),
        ("t", "ÁREA DEL TALLER ACTIVA", valor_visible(unidad.get("Ziga_Talleres.Área del taller activa"))),
    ])

    mostrar_tarjetas([
        ("t", "FECHA DE INGRESO", fecha_visible(unidad.get("Ziga_Talleres.Fecha de ingreso"))),
        ("t", "DÍAS", numero_entero(unidad.get("Ziga_Talleres.dias"))),
    ])


# =========================================================
# SINIESTRO
# =========================================================

if tiene_informacion(unidad.get("Siniestro")):

    hay_folio_abierto = True

    separador()

    titulo_seccion("🚨 Siniestro", dias_numero_o_none(unidad.get("Siniestros2.dias")))

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
# MOSTRAR REFACCIONES
# =========================================================

if grupos_refacciones:

    hay_folio_abierto = True

    separador()

    titulo_seccion("📦 Refacciones pendientes")

    for folio, actividades in grupos_refacciones:

        actividades_html = "<br>".join(
            f"• {actividad}"
            for actividad in actividades
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
# DESCARGAR FICHA DE LA UNIDAD
# =========================================================

separador()

ficha_serie = unidad.dropna()

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


# =========================================================
# PIE
# =========================================================

st.caption(
    "Casanova | Control de Unidades"
)
