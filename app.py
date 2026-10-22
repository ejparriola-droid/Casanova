import re
import hashlib
from collections import Counter
from pathlib import Path

import pandas as pd
import streamlit as st
from PIL import Image, ImageEnhance, ImageFilter, ImageOps
import pytesseract


# =========================================================
# CONFIGURACIÓN
# =========================================================

st.set_page_config(
    page_title="Casanova | Unidades en Centenario y Lala",
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


# =========================================================
# ESTILO
# =========================================================

def aplicar_estilo():
    st.markdown(
        f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@700;800&display=swap');

html, body, [class*="css"], .stApp {{ font-family: 'Inter', sans-serif; }}
.stApp {{ background: linear-gradient(180deg, #E6EEF9 0%, #F7F9FC 340px); }}
.stApp p, .stApp span, .stApp label, .stMarkdown {{ font-weight: 600; }}
#MainMenu, footer, header {{ visibility: hidden; }}
.block-container {{ padding-top: 1.3rem; padding-bottom: 3rem; max-width: 1250px; }}

/* ENCABEZADO */
.topbar {{
    position: relative; overflow: hidden;
    background: linear-gradient(120deg, {AZUL_MARINO} 0%, {AZUL_CASANOVA} 100%);
    border-radius: 20px; padding: 22px 28px;
    display: flex; align-items: center; gap: 18px;
    box-shadow: 0 10px 28px rgba(33,55,105,0.22);
    border-bottom: 4px solid {AMARILLO};
}}
.topbar::before {{
    content: ""; position: absolute; right: -60px; top: -80px;
    width: 260px; height: 260px; border-radius: 50%;
    background: radial-gradient(circle, rgba(255,183,0,0.28), transparent 70%);
}}
.topbar::after {{
    content: ""; position: absolute; left: 0; right: 0; bottom: 0; height: 100%;
    background: repeating-linear-gradient(115deg, transparent 0 28px, rgba(255,255,255,0.025) 28px 56px);
    pointer-events: none;
}}
.logo-c {{
    z-index: 1; width: 84px; height: 58px; border-radius: 16px; flex-shrink: 0;
    background: {AMARILLO}; color: {AZUL_MARINO};
    display: flex; align-items: center; justify-content: center;
    font-size: 34px; font-weight: 800;
    box-shadow: 0 6px 16px rgba(0,0,0,0.25), inset 0 -4px 0 rgba(0,0,0,0.10);
}}
.topbar-txt {{ z-index: 1; }}
.topbar-titulo {{ color: white; font-size: 25px; font-weight: 800; letter-spacing: .4px; line-height: 1.1; }}
.topbar-titulo b {{ color: {AMARILLO}; }}
.topbar-sub {{ color: rgba(255,255,255,.82); font-size: 13px; font-weight: 600; margin-top: 5px; }}
.sede {{
    display: inline-block; margin-right: 6px; padding: 2px 10px; border-radius: 999px;
    background: rgba(255,255,255,.14); font-size: 12px; font-weight: 800; color: white;
}}

.titulo-seccion {{
    color: {AZUL_MARINO}; font-size: 18px; font-weight: 800; margin: 26px 0 12px 0;
    padding-left: 12px; border-left: 5px solid {AMARILLO}; line-height: 1.2;
}}

/* TARJETAS */
.tarjeta {{
    background: {BLANCO}; border: 1px solid #E1E8F3; border-top: 3px solid {AZUL_CASANOVA};
    border-radius: 12px; padding: 14px 16px; min-height: 84px;
    box-shadow: 0 2px 8px rgba(33,55,105,0.05); margin-bottom: 12px; transition: all .2s ease;
}}
.tarjeta:hover {{ transform: translateY(-2px); box-shadow: 0 8px 18px rgba(33,55,105,0.12); border-top-color: {AMARILLO}; }}
.campo {{ color: #7A879B; font-size: 11px; font-weight: 800; letter-spacing: .8px; text-transform: uppercase; margin-bottom: 6px; }}
.valor {{ color: #172033; font-size: 16px; font-weight: 800; word-break: break-word; line-height: 1.3; }}
.valor-vacio {{ color: #A7B1C2; font-weight: 600; font-style: italic; }}

.badge {{ display: inline-block; padding: 4px 12px; border-radius: 20px; font-size: 13px; font-weight: 800; }}
.badge-verde {{ background: #E3F6EA; color: #167A3E; }}
.badge-amarillo {{ background: #FFF3D0; color: #8A6100; }}
.badge-rojo {{ background: #FDE5E5; color: #B42323; }}
.badge-azul {{ background: #E3ECF9; color: {AZUL_MARINO}; }}

/* KPIs */
.kpi-azul, .kpi-amarillo, .kpi-verde, .kpi-ambar, .kpi-rojo {{ border-radius: 16px; padding: 20px 24px; height: 100%; }}
.kpi-azul {{ background: linear-gradient(135deg, {AZUL_MARINO}, {AZUL_CASANOVA}); color: white; box-shadow: 0 6px 18px rgba(33,55,105,.20); }}
.kpi-amarillo {{ background: white; border: 2px solid {AMARILLO}; box-shadow: 0 6px 18px rgba(255,183,0,.15); }}
.kpi-verde {{ background: linear-gradient(135deg, #127A3D, #1FA855); color: white; box-shadow: 0 6px 18px rgba(22,122,62,.25); }}
.kpi-ambar {{ background: linear-gradient(135deg, #E59E00, {AMARILLO}); color: #2A1D00; box-shadow: 0 6px 18px rgba(255,183,0,.30); }}
.kpi-rojo {{ background: linear-gradient(135deg, #A81E1E, #E04040); color: white; box-shadow: 0 6px 18px rgba(200,40,40,.28); }}
.kpi-label {{ font-size: 12px; font-weight: 800; letter-spacing: 1px; text-transform: uppercase; opacity: .9; }}
.kpi-numero {{ font-size: 40px; font-weight: 800; margin-top: 4px; line-height: 1.1; }}
.kpi-amarillo .kpi-label {{ color: #7A879B; opacity: 1; }}
.kpi-amarillo .kpi-numero {{ color: {AZUL_MARINO}; }}
.kpi-sub {{ font-size: 13px; font-weight: 800; margin-top: 8px; opacity: .95; }}
.barra {{ height: 8px; background: rgba(255,255,255,.28); border-radius: 99px; margin-top: 12px; overflow: hidden; }}
.barra > div {{ height: 100%; background: rgba(255,255,255,.95); border-radius: 99px; }}

/* PLACA ESTILO MX */
.unidad-header {{
    background: white; border: 1px solid #E1E8F3; border-radius: 16px; padding: 18px 24px;
    margin: 20px 0 4px 0; box-shadow: 0 2px 10px rgba(33,55,105,.06);
    display: flex; flex-wrap: wrap; align-items: center; gap: 14px 30px;
}}
.placa {{
    position: relative; background: linear-gradient(180deg, #FFFFFF 0%, #F1F4F9 100%);
    color: {AZUL_MARINO}; font-family: 'JetBrains Mono', monospace;
    font-size: 26px; font-weight: 800; letter-spacing: 3px;
    padding: 12px 22px 8px 22px; border-radius: 10px; border: 3px solid {AZUL_MARINO};
    box-shadow: 0 3px 0 {AMARILLO}, 0 6px 14px rgba(33,55,105,.18); text-align: center;
}}
.placa small {{ display: block; font-family: 'Inter', sans-serif; font-size: 9px; letter-spacing: 3px; color: {AZUL_CASANOVA}; margin-top: 2px; font-weight: 800; }}
.placa.mini {{ font-size: 20px; padding: 8px 16px 6px 16px; display: inline-block; }}
.unidad-serie-label {{ color: #7A879B; font-size: 11px; font-weight: 800; letter-spacing: .8px; }}
.unidad-serie {{ color: {AZUL_MARINO}; font-size: 17px; font-weight: 800; word-break: break-all; }}

/* REFACCIONES */
.refaccion-card {{
    background: white; border: 1px solid #E1E8F3; border-left: 5px solid {AMARILLO};
    border-radius: 12px; padding: 16px 18px; margin-bottom: 12px; box-shadow: 0 2px 8px rgba(33,55,105,.05);
}}
.refaccion-folio {{ color: {AZUL_MARINO}; font-size: 14px; font-weight: 800; margin-bottom: 8px; }}
.refaccion-actividad {{ color: #172033; font-size: 14.5px; font-weight: 600; line-height: 1.7; }}

.sin-folios {{ background: #E3F6EA; border: 1px solid #BFE6CD; border-radius: 12px; padding: 15px 18px; color: #167A3E; font-weight: 800; margin-top: 24px; }}
.separador {{ height: 1px; background: linear-gradient(90deg, transparent, #CBD6E8, transparent); margin: 30px 0 6px 0; }}

/* CHIPS */
.resumen {{ display: flex; flex-wrap: wrap; gap: 10px; margin: 4px 0 8px 0; }}
.chip {{ display: inline-flex; align-items: center; gap: 8px; padding: 8px 16px; border-radius: 999px; font-size: 13px; font-weight: 800; border: 1px solid; }}
.chip-on {{ background: #FFF3D0; color: #8A6100; border-color: #F5D77A; }}
.chip-off {{ background: #F1F4F9; color: #94A3B8; border-color: #E1E8F3; }}
.chip-punto {{ width: 9px; height: 9px; border-radius: 50%; background: currentColor; }}

.seccion-head {{ display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px; margin: 26px 0 12px 0; }}
.seccion-head .titulo-seccion {{ margin: 0; }}
.dias-pill {{ padding: 6px 16px; border-radius: 999px; font-size: 13px; font-weight: 800; }}
.dias-verde {{ background: #E3F6EA; color: #167A3E; }}
.dias-ambar {{ background: #FFF3D0; color: #8A6100; }}
.dias-rojo {{ background: #FDE5E5; color: #B42323; }}

/* PANEL INICIO */
.panel {{ background: white; border: 1px solid #E1E8F3; border-radius: 16px; padding: 18px 22px; box-shadow: 0 2px 10px rgba(33,55,105,.06); height: 100%; }}
.panel-t {{ color: {AZUL_MARINO}; font-weight: 800; font-size: 14px; letter-spacing: .6px; text-transform: uppercase; margin-bottom: 12px; }}
.fila-b {{ display: flex; align-items: center; gap: 10px; margin: 9px 0; font-size: 13px; color: #172033; font-weight: 700; }}
.fila-b .n {{ width: 150px; flex-shrink: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
.fila-b .b {{ flex: 1; height: 10px; background: #EEF3FA; border-radius: 99px; overflow: hidden; }}
.fila-b .b > div {{ height: 100%; background: linear-gradient(90deg, {AZUL_CASANOVA}, {AMARILLO}); border-radius: 99px; }}
.fila-b .c {{ width: 44px; text-align: right; color: {AZUL_MARINO}; font-weight: 800; }}

/* ESCÁNER */
.scan-card {{
    background: linear-gradient(135deg, #FFFFFF, #F4F8FE); border: 1px solid #D8E1EF;
    border-left: 6px solid {AMARILLO}; border-radius: 16px; padding: 16px 20px; margin: 10px 0 14px 0;
    box-shadow: 0 4px 14px rgba(33,55,105,.08);
}}
.scan-card .t {{ color: {AZUL_MARINO}; font-weight: 800; font-size: 16px; }}
.scan-card .s {{ color: #64748B; font-size: 13px; font-weight: 600; margin-top: 3px; }}
.res-card {{
    background: white; border: 1px solid #E1E8F3; border-radius: 14px; padding: 12px 16px;
    margin-bottom: 8px; box-shadow: 0 2px 8px rgba(33,55,105,.05);
}}
.conf {{ font-size: 12px; font-weight: 800; color: #64748B; margin-top: 6px; }}

/* INPUTS */
.stTextInput label, .stTextInput label p, .stSelectbox label, .stSelectbox label p {{ color: {AZUL_MARINO} !important; font-weight: 800 !important; font-size: 15px !important; }}
.stTextInput div[data-baseweb="input"], .stTextInput div[data-baseweb="base-input"] {{ background: white !important; border-radius: 12px !important; }}
.stTextInput input {{
    color: #172033 !important; -webkit-text-fill-color: #172033 !important; caret-color: {AZUL_MARINO} !important;
    background: white !important; border-radius: 12px !important; border: 2px solid #D8E1EF !important;
    padding: 12px 16px !important; font-size: 16px !important; font-weight: 700 !important;
}}
.stTextInput input::placeholder {{ color: #64748B !important; -webkit-text-fill-color: #64748B !important; font-weight: 600 !important; opacity: 1 !important; }}
.stTextInput input:focus {{ border-color: {AZUL_CASANOVA} !important; box-shadow: 0 0 0 3px rgba(0,72,118,.12) !important; }}
.stTextInput button {{ background: white !important; color: {AZUL_MARINO} !important; border: none !important; }}
.stTextInput button svg {{ fill: {AZUL_MARINO} !important; color: {AZUL_MARINO} !important; }}
.stSelectbox div[data-baseweb="select"] > div {{ background: white !important; color: #172033 !important; border-radius: 12px !important; }}
.stSelectbox div[data-baseweb="select"] span, .stSelectbox div[data-baseweb="select"] div {{ color: #172033 !important; font-weight: 700 !important; }}
div[data-baseweb="popover"] ul, div[data-baseweb="popover"] li {{ background: white !important; color: #172033 !important; font-weight: 600 !important; }}
div[data-baseweb="popover"] li:hover {{ background: #EEF3FA !important; }}

.stButton > button, .stDownloadButton > button {{
    border-radius: 12px; font-weight: 700; padding: 10px 18px; border: 2px solid {AZUL_MARINO}; transition: all .2s ease;
}}
.stButton > button[kind="primary"] {{ background: {AZUL_MARINO}; color: white; }}
.stButton > button:hover {{ border-color: {AMARILLO}; color: {AZUL_MARINO}; background: {AMARILLO}; }}
.stDownloadButton > button {{ background: white; color: {AZUL_MARINO}; }}
.stDownloadButton > button:hover {{ background: {AMARILLO}; border-color: {AMARILLO}; color: {AZUL_MARINO}; }}

/* Toggle y radio */
.stToggle label p {{ color: {AZUL_MARINO} !important; font-weight: 800 !important; font-size: 15px !important; }}
.stRadio label p {{ color: {AZUL_MARINO} !important; font-weight: 700 !important; }}

/* LOGIN */
.login-box {{
    max-width: 430px; margin: 50px auto 24px auto; background: white; padding: 34px 30px 26px 30px;
    border-radius: 20px; border: 1px solid #D8E1EF; border-top: 6px solid {AMARILLO};
    box-shadow: 0 12px 40px rgba(33,55,105,.15); text-align: center;
}}
.login-box .logo-c {{ margin: 0 auto 10px auto; }}
.login-title {{ color: {AZUL_MARINO}; font-size: 28px; font-weight: 800; letter-spacing: 2px; }}
.login-subtitle {{ color: #64748B; font-size: 14px; font-weight: 600; margin-top: 4px; }}

@keyframes aparecer {{ from {{ opacity: 0; transform: translateY(8px); }} to {{ opacity: 1; transform: translateY(0); }} }}
.tarjeta, .refaccion-card, .unidad-header, .kpi-azul, .kpi-verde, .kpi-ambar, .kpi-rojo, .kpi-amarillo, .panel, .res-card, .scan-card {{ animation: aparecer .35s ease both; }}

@media (max-width: 768px) {{
    .topbar-titulo {{ font-size: 18px; }}
    .kpi-numero {{ font-size: 30px; }}
    .placa {{ font-size: 20px; }}
}}
</style>
""",
        unsafe_allow_html=True
    )


aplicar_estilo()


# =========================================================
# LOGIN
# =========================================================

def mostrar_login():
    st.markdown(
        """
<div class="login-box">
<div class="logo-c"><svg viewBox="0 0 64 40" width="52" height="33" xmlns="http://www.w3.org/2000/svg"><path d="M3 29C3 23 7 21 13 19.5C17 12 25 7 34 7C44 7 51 12.5 54 19.5C59 20.5 62 23.5 62 29L62 31L3 31Z" fill="#213769"/><path d="M18.5 19.5C21.500 14 27 10.500 32.500 10.500L32.500 19.500Z" fill="#FFB700"/><path d="M36 10.500C42 10.800 47 14.500 49 19.500L36 19.500Z" fill="#FFB700"/><circle cx="58.500" cy="25" r="2.200" fill="#FFB700"/><rect x="3" y="27" width="5" height="2.500" rx="1" fill="#FFB700"/><circle cx="17" cy="31" r="8" fill="#FFB700"/><circle cx="17" cy="31" r="6.300" fill="#213769"/><circle cx="17" cy="31" r="2.300" fill="#FFB700"/><circle cx="49" cy="31" r="8" fill="#FFB700"/><circle cx="49" cy="31" r="6.300" fill="#213769"/><circle cx="49" cy="31" r="2.300" fill="#FFB700"/></svg></div>
<div class="login-title">CASANOVA</div>
<div class="login-subtitle">Unidades en Centenario y Lala</div>
</div>
""",
        unsafe_allow_html=True
    )

    _, col2, _ = st.columns([1, 1.4, 1])

    with col2:
        usuario = st.text_input("Usuario", placeholder="Usuario")
        password = st.text_input("Contraseña", type="password", placeholder="Contraseña")

        if st.button("🔐 Iniciar sesión", use_container_width=True, type="primary"):
            if usuario.strip() == USUARIO_CORRECTO and password == PASSWORD_CORRECTA:
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
# DATOS
# =========================================================

def normalizar_columnas(df):
    df = df.copy()
    columnas = []
    for c in df.columns:
        n = str(c).replace("\n", " ").replace("\r", " ").replace("\xa0", " ")
        columnas.append(" ".join(n.split()).strip())
    df.columns = columnas
    return df


@st.cache_data
def cargar_concentrado():
    return normalizar_columnas(pd.read_excel(ARCHIVO_BASE, sheet_name=HOJA_BASE))


@st.cache_data
def cargar_refacciones():
    try:
        return normalizar_columnas(pd.read_excel(ARCHIVO_BASE, sheet_name=HOJA_REFACCIONES))
    except Exception:
        return pd.DataFrame()


def tiene_informacion(valor):
    if valor is None:
        return False
    try:
        if pd.isna(valor):
            return False
    except Exception:
        pass
    t = str(valor).strip()
    return t != "" and t.lower() not in ["nan", "none", "null", "nat"]


def valor_visible(valor):
    return str(valor).strip() if tiene_informacion(valor) else "Sin información"


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
        f = pd.to_datetime(valor, errors="coerce")
        return valor_visible(valor) if pd.isna(f) else f.strftime("%d/%m/%Y")
    except Exception:
        return valor_visible(valor)


def normalizar_texto(valor):
    return str(valor).strip().upper() if tiene_informacion(valor) else ""


def normalizar_busqueda(valor):
    """Solo letras y números: ignora guiones, espacios y puntos."""
    if not tiene_informacion(valor):
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(valor).upper())


# =========================================================
# LECTOR OCR DE PLACAS
# Las placas son cualquier mezcla de letras, números y guiones
# (SKD-S4S5, 55S-S-S554, DGD-DDF-5...), así que NO se valida por
# formato: se lee todo lo posible y se compara contra las placas
# reales del concentrado, tolerando confusiones típicas del OCR
# (O/0/D/Q, I/1/L, S/5, B/8, Z/2, G/6, A/4, U/V, T/7).
# =========================================================

CONFUSABLES = {
    "O": "0", "Q": "0", "D": "0",
    "I": "1", "L": "1",
    "S": "5", "B": "8", "Z": "2",
    "G": "6", "A": "4", "V": "U", "T": "7",
}


def canonica(texto):
    return "".join(CONFUSABLES.get(c, c) for c in texto)


def distancia_limitada(a, b, limite=1):
    """Levenshtein con corte temprano."""
    if abs(len(a) - len(b)) > limite:
        return limite + 1
    anterior = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        actual = [i]
        for j, cb in enumerate(b, 1):
            actual.append(min(actual[-1] + 1, anterior[j] + 1, anterior[j - 1] + (ca != cb)))
        if min(actual) > limite:
            return limite + 1
        anterior = actual
    return anterior[-1]


def cargar_indice_placas(df):
    """norm -> texto original (con guiones tal como está en el Excel)."""
    indice = {}
    if "PLACA" not in df.columns:
        return indice
    for v in df["PLACA"].dropna().astype(str):
        n = normalizar_busqueda(v)
        if 4 <= len(n) <= 10 and n not in indice:
            indice[n] = v.strip().upper()
    return indice


def preparar_base(imagen):
    imagen = ImageOps.exif_transpose(imagen).convert("RGB")
    ancho, alto = imagen.size
    objetivo = 1500
    if ancho != objetivo:
        factor = objetivo / ancho
        imagen = imagen.resize((objetivo, max(1, int(alto * factor))), Image.Resampling.LANCZOS)
    return imagen


def recortes(imagen):
    w, h = imagen.size
    return [
        imagen,
        imagen.crop((int(w * .05), int(h * .25), int(w * .95), int(h * .75))),
        imagen.crop((int(w * .15), int(h * .35), int(w * .85), int(h * .65))),
    ]


def variantes(recorte):
    gris = ImageOps.autocontrast(ImageOps.grayscale(recorte), cutoff=2)
    gris = ImageEnhance.Sharpness(gris).enhance(2.0)
    media = sum(i * c for i, c in enumerate(gris.histogram())) / max(1, gris.size[0] * gris.size[1])
    binaria = gris.point(lambda p, t=media * 0.92: 255 if p > t else 0)
    return [gris, binaria, ImageOps.invert(binaria)]


def ocr_textos(imagen_pil):
    """Devuelve todos los textos que Tesseract lee con varias configuraciones."""
    base = preparar_base(imagen_pil)
    cfg_char = "-c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-"
    textos = []
    for k, rec in enumerate(recortes(base)):
        psms = (7, 6, 11) if k == 0 else (7, 8)
        for var in variantes(rec):
            for psm in psms:
                texto = pytesseract.image_to_string(
                    var, config=f"--oem 3 --psm {psm} {cfg_char}", lang="eng"
                )
                if texto.strip():
                    textos.append(texto.upper())
    return textos


def extraer_cadenas(textos):
    """Cadenas compactas (solo A-Z 0-9) con cuántas veces aparecieron."""
    votos = Counter()
    for t in textos:
        lineas = [re.sub(r"[^A-Z0-9]", "", l) for l in t.splitlines()]
        lineas = [l for l in lineas if l]
        for l in lineas:
            if 5 <= len(l) <= 14:
                votos[l] += 1
        # placas partidas en varias líneas
        for i in range(len(lineas) - 1):
            junto = lineas[i] + lineas[i + 1]
            if 6 <= len(junto) <= 12:
                votos[junto] += 1
        # tokens separados por espacio/guion
        for tok in re.split(r"[\s]+", t):
            c = re.sub(r"[^A-Z0-9]", "", tok)
            if 5 <= len(c) <= 10:
                votos[c] += 1
    return votos


def emparejar(votos, indice):
    """Compara lo leído con las placas reales y devuelve las mejores."""
    reales_canon = {}
    for n in indice:
        reales_canon.setdefault(canonica(n), []).append(n)

    encontrados = {}

    def anotar(norma, score, metodo):
        if norma not in encontrados or score > encontrados[norma][0]:
            encontrados[norma] = (score, metodo)

    top = [c for c, _ in votos.most_common(40)]

    for cand in top:
        v = min(votos[cand], 5) * 0.01
        cc = canonica(cand)

        # 1) exacta
        if cand in indice:
            anotar(cand, 1.0, "Coincidencia exacta")
            continue

        # 2) igual tras igualar caracteres confundibles (O/0, S/5...)
        for n in reales_canon.get(cc, []):
            anotar(n, 0.93 + v, "Corregida (caracteres parecidos)")

        # 3) la placa real aparece dentro de lo leído (OCR metió ruido)
        for rc, lista in reales_canon.items():
            if len(rc) >= 6 and rc in cc and len(cc) - len(rc) <= 4:
                for n in lista:
                    anotar(n, 0.86 + v, "Encontrada dentro del texto leído")

    # 4) a un carácter de distancia (solo si aún no hay nada fuerte)
    if not any(s >= 0.9 for s, _ in encontrados.values()):
        for cand in top[:15]:
            cc = canonica(cand)
            if not 6 <= len(cc) <= 10:
                continue
            for rc, lista in reales_canon.items():
                if distancia_limitada(cc, rc, 1) <= 1:
                    for n in lista:
                        anotar(n, 0.78 + min(votos[cand], 5) * 0.01, "Parecida (1 carácter distinto)")

    resultado = [
        {"norma": n, "mostrar": indice[n], "score": s, "metodo": m, "en_concentrado": True}
        for n, (s, m) in encontrados.items()
    ]
    resultado.sort(key=lambda r: r["score"], reverse=True)
    return resultado[:5]


def leer_placa_foto(archivo, df):
    try:
        imagen = Image.open(archivo)
    except Exception:
        return [], "No se pudo abrir la imagen."

    try:
        textos = ocr_textos(imagen)
    except pytesseract.TesseractNotFoundError:
        return [], "Tesseract no está instalado en el servidor."
    except Exception as e:
        return [], f"No se pudo ejecutar OCR: {e}"

    votos = extraer_cadenas(textos)
    indice = cargar_indice_placas(df)
    resultados = emparejar(votos, indice)

    if resultados:
        return resultados, ""

    crudos = [c for c, _ in votos.most_common(6) if 6 <= len(c) <= 10][:3]
    return [
        {"norma": c, "mostrar": c, "score": 0.0, "metodo": "Leída por OCR (no está en el concentrado)", "en_concentrado": False}
        for c in crudos
    ], ""


def usar_placa(placa_texto):
    st.session_state["buscador_unidad"] = placa_texto
    st.session_state["scan_on"] = False
    st.session_state.pop("ocr_resultados", None)
    st.session_state.pop("ocr_hash", None)


def mostrar_lector_placas(df):
    st.toggle(
        "📷 Escanear placa con foto o cámara",
        key="scan_on",
        help="La cámara y la carga de fotos aparecen solo al activar."
    )

    if not st.session_state.get("scan_on"):
        return

    st.markdown(
        """
<div class="scan-card">
<div class="t">📷 Lector de placas</div>
<div class="s">Acércate a la placa, que ocupe buena parte de la imagen y esté bien enfocada.
Lee letras, números y guiones en cualquier combinación.</div>
</div>
""",
        unsafe_allow_html=True
    )

    modo = st.radio(
        "Origen de la foto",
        ["📸 Tomar foto", "🖼️ Subir foto"],
        horizontal=True,
        label_visibility="collapsed",
        key="scan_modo"
    )

    archivo = None

    if modo == "📸 Tomar foto":
        archivo = st.camera_input("Toma la foto de la placa", key="foto_placa_camara")
    else:
        archivo = st.file_uploader(
            "Selecciona la foto de la placa",
            type=["jpg", "jpeg", "png", "webp"],
            key="foto_placa_archivo"
        )

    if archivo is None:
        return

    huella = hashlib.md5(archivo.getvalue()).hexdigest()

    if st.session_state.get("ocr_hash") != huella:
        with st.spinner("Leyendo placa..."):
            resultados, error = leer_placa_foto(archivo, df)
        st.session_state["ocr_hash"] = huella
        st.session_state["ocr_resultados"] = resultados
        st.session_state["ocr_error"] = error

    if st.session_state.get("ocr_error"):
        st.warning(st.session_state["ocr_error"])
        return

    resultados = st.session_state.get("ocr_resultados", [])

    if not resultados:
        st.warning(
            "No pude identificar la placa. Intenta con más luz, sin reflejos y con la placa más grande en la foto."
        )
        return

    st.markdown('<div class="campo" style="margin-top:10px;">Placas detectadas</div>', unsafe_allow_html=True)

    for i, r in enumerate(resultados):
        c1, c2, c3 = st.columns([2.2, 2.6, 1.4], vertical_alignment="center")
        with c1:
            st.markdown(
                f'<div class="placa mini">{r["mostrar"]}</div>',
                unsafe_allow_html=True
            )
        with c2:
            if r["en_concentrado"]:
                st.success("Encontrada en concentrado")
                st.markdown(
                    f'<div class="conf">{r["metodo"]} · confianza {min(r["score"], 1):.0%}</div>',
                    unsafe_allow_html=True
                )
            else:
                st.warning("No está en el concentrado")
                st.markdown(f'<div class="conf">{r["metodo"]}</div>', unsafe_allow_html=True)
        with c3:
            st.button(
                "Usar placa",
                key=f"usar_ocr_{i}",
                use_container_width=True,
                on_click=usar_placa,
                args=(r["mostrar"],),
                disabled=not r["en_concentrado"]
            )


# =========================================================
# COMPONENTES VISUALES
# =========================================================

def tarjeta(campo, valor):
    clase = "valor valor-vacio" if valor == "Sin información" else "valor"
    return f'<div class="tarjeta"><div class="campo">{campo}</div><div class="{clase}">{valor}</div></div>'


def tarjeta_badge(campo, valor):
    if valor == "Sin información":
        return tarjeta(campo, valor)
    t = valor.lower()
    if any(p in t for p in ["finaliz", "complet", "cerrad", "entreg", "listo", "vigente", "activ"]):
        clase = "badge-verde"
    elif any(p in t for p in ["cancel", "venc", "rechaz", "detenid"]):
        clase = "badge-rojo"
    elif any(p in t for p in ["pend", "proceso", "espera", "revis", "abiert", "tramite", "trámite"]):
        clase = "badge-amarillo"
    else:
        clase = "badge-azul"
    return (
        f'<div class="tarjeta"><div class="campo">{campo}</div>'
        f'<div class="valor"><span class="badge {clase}">{valor}</span></div></div>'
    )


def mostrar_tarjetas(items):
    for col, (tipo, campo, valor) in zip(st.columns(len(items)), items):
        with col:
            st.markdown(
                tarjeta_badge(campo, valor) if tipo == "b" else tarjeta(campo, valor),
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
            f'<div class="seccion-head"><div class="titulo-seccion">{texto}</div>'
            f'<div class="dias-pill dias-{nivel}">⏱ {dias:,} días</div></div>'
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
<div class="logo-c"><svg viewBox="0 0 64 40" width="52" height="33" xmlns="http://www.w3.org/2000/svg"><path d="M3 29C3 23 7 21 13 19.5C17 12 25 7 34 7C44 7 51 12.5 54 19.5C59 20.5 62 23.5 62 29L62 31L3 31Z" fill="#213769"/><path d="M18.5 19.5C21.500 14 27 10.500 32.500 10.500L32.500 19.500Z" fill="#FFB700"/><path d="M36 10.500C42 10.800 47 14.500 49 19.500L36 19.500Z" fill="#FFB700"/><circle cx="58.500" cy="25" r="2.200" fill="#FFB700"/><rect x="3" y="27" width="5" height="2.500" rx="1" fill="#FFB700"/><circle cx="17" cy="31" r="8" fill="#FFB700"/><circle cx="17" cy="31" r="6.300" fill="#213769"/><circle cx="17" cy="31" r="2.300" fill="#FFB700"/><circle cx="49" cy="31" r="8" fill="#FFB700"/><circle cx="49" cy="31" r="6.300" fill="#213769"/><circle cx="49" cy="31" r="2.300" fill="#FFB700"/></svg></div>
<div class="topbar-txt">
<div class="topbar-titulo">CASANOVA <b>|</b> UNIDADES EN CENTENARIO Y LALA</div>
<div class="topbar-sub"><span class="sede">📍 Centenario</span><span class="sede">📍 Lala</span> Consulta individual de unidades</div>
</div>
</div>
""",
        unsafe_allow_html=True
    )

with col2:
    if st.button("Cerrar sesión", use_container_width=True):
        st.session_state["autenticado"] = False
        st.rerun()


# =========================================================
# CARGA
# =========================================================

try:
    df = cargar_concentrado()
except Exception as e:
    st.error("No se pudo cargar concentrado.xlsx")
    st.code(str(e))
    st.stop()

df_refacciones = cargar_refacciones()


# =========================================================
# BUSCADOR + ESCÁNER
# =========================================================

titulo_seccion("🔎 Buscar unidad")

busqueda = st.text_input(
    "Busca por PLACA o SERIE",
    placeholder="Ejemplo: SKD-S4S5, 55S-S-S554 o número de serie...",
    label_visibility="collapsed",
    key="buscador_unidad"
)

mostrar_lector_placas(df)


# =========================================================
# TOTALES
# =========================================================

if "SERIE" in df.columns:
    s = df["SERIE"].astype("string").fillna("").str.strip().str.upper()
    total_unidades = s[s != ""].nunique()
else:
    total_unidades = 0


# =========================================================
# INICIO (SIN BÚSQUEDA)
# =========================================================

def usar_reciente(p):
    st.session_state["buscador_unidad"] = p


if not busqueda:
    dias_col = pd.to_numeric(df.get("DÍAS EN PISO"), errors="coerce") if "DÍAS EN PISO" in df.columns else pd.Series(dtype=float)
    n_rojo = int((dias_col >= DIAS_ROJO).sum()) if not dias_col.empty else 0
    n_ambar = int(((dias_col >= DIAS_AMBAR) & (dias_col < DIAS_ROJO)).sum()) if not dias_col.empty else 0
    prom = f"{dias_col.mean():.0f}" if not dias_col.dropna().empty else "—"

    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(f'<div class="kpi-azul"><div class="kpi-label">🚗 Total de unidades</div><div class="kpi-numero">{total_unidades:,}</div></div>', unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-amarillo"><div class="kpi-label">⏱ Promedio días en piso</div><div class="kpi-numero">{prom}</div></div>', unsafe_allow_html=True)
    k3.markdown(f'<div class="kpi-ambar"><div class="kpi-label">⚠️ Requieren atención</div><div class="kpi-numero">{n_ambar:,}</div><div class="kpi-sub">{DIAS_AMBAR}+ días en piso</div></div>', unsafe_allow_html=True)
    k4.markdown(f'<div class="kpi-rojo"><div class="kpi-label">🚨 Críticas</div><div class="kpi-numero">{n_rojo:,}</div><div class="kpi-sub">{DIAS_ROJO}+ días en piso</div></div>', unsafe_allow_html=True)

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

    p1, p2 = st.columns([1.4, 1])

    with p1:
        if "ESTATUS" in df.columns:
            conteo = df["ESTATUS"].dropna().astype(str).str.strip().value_counts().head(6)
            if not conteo.empty:
                maximo = conteo.max()
                filas = "".join(
                    f'<div class="fila-b"><div class="n">{n}</div>'
                    f'<div class="b"><div style="width:{c / maximo * 100:.0f}%"></div></div>'
                    f'<div class="c">{c:,}</div></div>'
                    for n, c in conteo.items()
                )
                st.markdown(f'<div class="panel"><div class="panel-t">📊 Unidades por estatus</div>{filas}</div>', unsafe_allow_html=True)

    with p2:
        recientes = st.session_state.get("recientes", [])
        st.markdown('<div class="panel"><div class="panel-t">🕘 Consultadas recientemente</div>', unsafe_allow_html=True)
        if recientes:
            for i, r in enumerate(recientes):
                st.button(r, key=f"rec_{i}", use_container_width=True, on_click=usar_reciente, args=(r,))
        else:
            st.caption("Aquí aparecerán las últimas unidades que consultes.")
        st.markdown("</div>", unsafe_allow_html=True)

    st.info("Escribe una placa o serie, o activa el escáner para leer la placa con una foto.")
    st.stop()


# =========================================================
# BÚSQUEDA
# =========================================================

criterio = normalizar_busqueda(busqueda)

placas = df["PLACA"].apply(normalizar_busqueda)
series = df["SERIE"].apply(normalizar_busqueda)

resultado = df[
    placas.str.contains(criterio, case=False, na=False, regex=False)
    | series.str.contains(criterio, case=False, na=False, regex=False)
].copy()

if resultado.empty:
    st.warning(f"No se encontró ninguna unidad con: {busqueda}")
    st.stop()

if len(resultado) > 1:
    opciones = [
        f"{valor_visible(f.get('PLACA'))} | {valor_visible(f.get('SERIE'))}"
        for _, f in resultado.iterrows()
    ]
    seleccion = st.selectbox("Selecciona la unidad", opciones)
    unidad = resultado.iloc[opciones.index(seleccion)]
else:
    unidad = resultado.iloc[0]

placa = valor_visible(unidad.get("PLACA"))
serie = valor_visible(unidad.get("SERIE"))

# historial reciente
rec = st.session_state.get("recientes", [])
if placa != "Sin información":
    rec = [placa] + [r for r in rec if r != placa]
    st.session_state["recientes"] = rec[:5]


# =========================================================
# DÍAS EN PISO
# =========================================================

dias_piso = unidad.get("DÍAS EN PISO")
dias_n = dias_numero_o_none(dias_piso)

if tiene_informacion(dias_piso) and dias_n is not None:
    texto_dias = f"{dias_n:,} días"
elif tiene_informacion(dias_piso):
    texto_dias = valor_visible(dias_piso)
else:
    texto_dias = "🆕 Recién ingresado"

nivel_piso = nivel_dias(dias_n)
clase_kpi = {"verde": "kpi-verde", "ambar": "kpi-ambar", "rojo": "kpi-rojo", None: "kpi-amarillo"}[nivel_piso]
texto_nivel = {
    "verde": "✅ Rotación normal",
    "ambar": "⚠️ Requiere atención",
    "rojo": "🚨 Tiempo crítico en piso",
    None: ""
}[nivel_piso]

sub_html = f'<div class="kpi-sub">{texto_nivel}</div>' if texto_nivel else ""
barra_html = ""
if dias_n is not None and nivel_piso in ("verde", "ambar", "rojo"):
    pct = min(100, dias_n / DIAS_ROJO * 100)
    barra_html = f'<div class="barra"><div style="width:{pct:.0f}%"></div></div>'

c1, c2 = st.columns(2)

c1.markdown(
    f'<div class="kpi-azul" style="margin-top:14px;"><div class="kpi-label">🚗 Total de unidades</div>'
    f'<div class="kpi-numero">{total_unidades:,}</div></div>',
    unsafe_allow_html=True
)
c2.markdown(
    f'<div class="{clase_kpi}" style="margin-top:14px;"><div class="kpi-label">🏭 Días en piso</div>'
    f'<div class="kpi-numero">{texto_dias}</div>{sub_html}{barra_html}</div>',
    unsafe_allow_html=True
)


# =========================================================
# IDENTIDAD DE LA UNIDAD
# =========================================================

st.markdown(
    f"""
<div class="unidad-header">
<div class="placa">{placa}<small>CASANOVA · MÉXICO</small></div>
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
    cols_ref = {str(c).strip().lower(): c for c in df_refacciones.columns}
    c_serie = cols_ref.get("serie")
    c_folio = cols_ref.get("folio ws")
    c_act = cols_ref.get("actividad")

    if c_serie is not None and c_folio is not None and c_act is not None:
        ref = df_refacciones.copy()
        ref["_S"] = ref[c_serie].astype("string").fillna("").str.strip().str.upper()
        ref = ref[ref["_S"] == normalizar_texto(serie)].copy()

        for folio, grupo in ref.groupby(c_folio, sort=False):
            if not tiene_informacion(folio):
                continue
            actividades = []
            for a in grupo[c_act]:
                if not tiene_informacion(a):
                    continue
                t = str(a).strip()
                if not any(t.casefold() == e.casefold() for e in actividades):
                    actividades.append(t)
            if actividades:
                grupos_refacciones.append((str(folio).strip(), actividades))


# =========================================================
# RESUMEN DE FOLIOS
# =========================================================

folios_resumen = [
    ("📋 Gestoría", tiene_informacion(unidad.get("Gestorías1.Folio"))),
    ("🔧 Mantenimiento", tiene_informacion(unidad.get("Mantenimiento.FOLIO"))),
    ("🛠️ Ziga Talleres", tiene_informacion(unidad.get("Ziga_Talleres.Folio"))),
    ("🚨 Siniestro", tiene_informacion(unidad.get("Siniestro"))),
    ("📦 Refacciones", bool(grupos_refacciones)),
]

total_abiertos = sum(1 for _, a in folios_resumen if a)

chips_html = "".join(
    f'<span class="chip {"chip-on" if a else "chip-off"}"><span class="chip-punto"></span>{n}</span>'
    for n, a in folios_resumen
)

titulo_seccion(f"🚦 Resumen de folios abiertos ({total_abiertos})")
st.markdown(f'<div class="resumen">{chips_html}</div>', unsafe_allow_html=True)


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
# FOLIOS ABIERTOS
# =========================================================

hay_folio_abierto = False

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

if grupos_refacciones:
    hay_folio_abierto = True
    separador()
    titulo_seccion("📦 Refacciones pendientes")
    for folio, actividades in grupos_refacciones:
        acts = "<br>".join(f"• {a}" for a in actividades)
        st.markdown(
            f'<div class="refaccion-card"><div class="refaccion-folio">FOLIO WS: {folio}</div>'
            f'<div class="refaccion-actividad">{acts}</div></div>',
            unsafe_allow_html=True
        )

if not hay_folio_abierto:
    st.markdown('<div class="sin-folios">✅ Esta unidad no tiene folios abiertos.</div>', unsafe_allow_html=True)


# =========================================================
# DESCARGA
# =========================================================

separador()

ficha_serie = unidad.dropna()
ficha_df = pd.DataFrame({"CAMPO": ficha_serie.index.astype(str), "VALOR": ficha_serie.astype(str).values})

st.download_button(
    "⬇️ Descargar ficha de la unidad (CSV)",
    data=ficha_df.to_csv(index=False).encode("utf-8-sig"),
    file_name=f"ficha_{normalizar_busqueda(placa)}.csv",
    mime="text/csv"
)

st.caption("Casanova | Unidades en Centenario y Lala")
