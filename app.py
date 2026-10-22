import io
import re
import hashlib
from collections import Counter
from datetime import datetime
from pathlib import Path

import altair as alt
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
VERDE = "#1FA855"
ROJO = "#D63A3A"

DIAS_AMBAR = 30
DIAS_ROJO = 60

LOGO_SVG = (
    '<svg viewBox="0 0 64 40" width="52" height="33" xmlns="http://www.w3.org/2000/svg">'
    '<path d="M3 29C3 23 7 21 13 19.5C17 12 25 7 34 7C44 7 51 12.5 54 19.5C59 20.5 62 23.5 62 29L62 31L3 31Z" fill="#213769"/>'
    '<path d="M18.5 19.5C21.500 14 27 10.500 32.500 10.500L32.500 19.500Z" fill="#FFB700"/>'
    '<path d="M36 10.500C42 10.800 47 14.500 49 19.500L36 19.500Z" fill="#FFB700"/>'
    '<circle cx="58.500" cy="25" r="2.200" fill="#FFB700"/>'
    '<rect x="3" y="27" width="5" height="2.500" rx="1" fill="#FFB700"/>'
    '<circle cx="17" cy="31" r="8" fill="#FFB700"/><circle cx="17" cy="31" r="6.300" fill="#213769"/><circle cx="17" cy="31" r="2.300" fill="#FFB700"/>'
    '<circle cx="49" cy="31" r="8" fill="#FFB700"/><circle cx="49" cy="31" r="6.300" fill="#213769"/><circle cx="49" cy="31" r="2.300" fill="#FFB700"/>'
    '</svg>'
)


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

/* TARJETAS INTERACTIVAS (PANEL) */
.kc {{ min-height: 142px; margin-bottom: 8px; transition: all .2s ease; }}
.kc:hover {{ transform: translateY(-3px); }}
.kc-blanco {{
    background: white; border: 1px solid #E1E8F3; border-top: 5px solid {AZUL_CASANOVA};
    border-radius: 16px; padding: 18px 22px; box-shadow: 0 4px 14px rgba(33,55,105,.08);
}}
.kc-blanco:hover {{ border-top-color: {AMARILLO}; box-shadow: 0 10px 24px rgba(33,55,105,.16); }}
.kc-blanco .kpi-label {{ color: #7A879B; opacity: 1; }}
.kc-blanco .kpi-numero {{ color: {AZUL_MARINO}; }}
.kc-blanco .kpi-sub {{ color: #64748B; }}
.kc-blanco .barra {{ background: #EEF3FA; }}
.kc-blanco .barra > div {{ background: linear-gradient(90deg, {AZUL_CASANOVA}, {AMARILLO}); }}

.headline {{
    background: white; border: 1px solid #E1E8F3; border-left: 6px solid {AMARILLO};
    border-radius: 16px; padding: 18px 24px; box-shadow: 0 4px 16px rgba(33,55,105,.08);
    color: #172033; font-size: 16px; font-weight: 600; line-height: 1.7; margin: 4px 0 16px 0;
}}
.headline b {{ color: {AZUL_MARINO}; font-weight: 800; }}
.headline .rojo {{ color: #B42323; font-weight: 800; }}

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
.placa.mini.vin {{ font-size: 15px; letter-spacing: 1.5px; padding: 8px 12px 6px 12px; }}
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
.dias-azul {{ background: #E3ECF9; color: {AZUL_MARINO}; }}

/* PANEL */
.panel {{ background: white; border: 1px solid #E1E8F3; border-radius: 16px; padding: 18px 22px; box-shadow: 0 2px 10px rgba(33,55,105,.06); height: 100%; }}
.panel-t {{ color: {AZUL_MARINO}; font-weight: 800; font-size: 14px; letter-spacing: .6px; text-transform: uppercase; margin-bottom: 12px; }}
.fila-b {{ display: flex; align-items: center; gap: 10px; margin: 9px 0; font-size: 13px; color: #172033; font-weight: 700; }}
.fila-b .n {{ width: 170px; flex-shrink: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
.fila-b .b {{ flex: 1; height: 10px; background: #EEF3FA; border-radius: 99px; overflow: hidden; }}
.fila-b .b > div {{ height: 100%; background: linear-gradient(90deg, {AZUL_CASANOVA}, {AMARILLO}); border-radius: 99px; }}
.fila-b .c {{ width: 52px; text-align: right; color: {AZUL_MARINO}; font-weight: 800; }}

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
.tarjeta, .refaccion-card, .unidad-header, .kpi-azul, .kpi-verde, .kpi-ambar, .kpi-rojo, .kpi-amarillo, .kc-blanco, .headline, .panel, .res-card, .scan-card {{ animation: aparecer .35s ease both; }}

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
        f"""
<div class="login-box">
<div class="logo-c">{LOGO_SVG}</div>
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
        if pd.isna(f):
            return valor_visible(valor)
        return f.strftime("%d/%m/%Y")
    except Exception:
        return valor_visible(valor)


def normalizar_texto(valor):
    return str(valor).strip().upper() if tiene_informacion(valor) else ""


def normalizar_busqueda(valor):
    """Ignora espacios, guiones, puntos y símbolos. No impone formato."""
    if not tiene_informacion(valor):
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(valor).upper().strip())


def fecha_base():
    try:
        return datetime.fromtimestamp(ARCHIVO_BASE.stat().st_mtime).strftime("%d/%m/%Y")
    except Exception:
        return ""


# =========================================================
# OCR DE PLACAS (sin formato fijo; se compara contra el concentrado)
# =========================================================

CONFUSABLES = {
    "O": "0", "Q": "0", "D": "0",
    "I": "1", "L": "1",
    "S": "5", "B": "8", "Z": "2",
    "G": "6", "A": "4",
    "V": "U", "T": "7", "E": "3"
}

PARES_CONFUSION = {frozenset((k, v)) for k, v in CONFUSABLES.items()}


def canonica(texto):
    texto = normalizar_busqueda(texto)
    return "".join(CONFUSABLES.get(c, c) for c in texto)


def distancia_limitada(a, b, limite=2):
    if abs(len(a) - len(b)) > limite:
        return limite + 1

    anterior = list(range(len(b) + 1))

    for i, ca in enumerate(a, 1):
        actual = [i]
        for j, cb in enumerate(b, 1):
            costo = 0 if ca == cb else 1
            actual.append(min(actual[-1] + 1, anterior[j] + 1, anterior[j - 1] + costo))
        if min(actual) > limite:
            return limite + 1
        anterior = actual

    return anterior[-1]


def costo_caracter(a, b):
    if a == b:
        return 0
    if frozenset((a, b)) in PARES_CONFUSION:
        return 0.15
    return 1


def distancia_ocr(a, b, limite=2):
    if abs(len(a) - len(b)) > limite:
        return limite + 1

    anterior = list(range(len(b) + 1))

    for i, ca in enumerate(a, 1):
        actual = [i]
        for j, cb in enumerate(b, 1):
            actual.append(min(
                anterior[j - 1] + costo_caracter(ca, cb),
                actual[j - 1] + 1,
                anterior[j] + 1
            ))
        if min(actual) > limite:
            return limite + 1
        anterior = actual

    return anterior[-1]


def cargar_indice_placas(df):
    indice = {}

    if "PLACA" not in df.columns:
        return indice

    for valor in df["PLACA"].dropna():
        original = str(valor).strip()
        norma = normalizar_busqueda(original)

        if not norma or len(norma) < 4 or len(norma) > 14:
            continue

        if norma not in indice:
            indice[norma] = original.upper()

    return indice


def preparar_base(imagen):
    imagen = ImageOps.exif_transpose(imagen).convert("RGB")
    ancho, alto = imagen.size
    objetivo = 1800

    if ancho < objetivo:
        factor = objetivo / ancho
        imagen = imagen.resize((objetivo, max(1, int(alto * factor))), Image.Resampling.LANCZOS)
    elif ancho > 2200:
        factor = 2200 / ancho
        imagen = imagen.resize((2200, max(1, int(alto * factor))), Image.Resampling.LANCZOS)

    return imagen


def recortes(imagen):
    w, h = imagen.size
    return [
        imagen,
        imagen.crop((int(w * .03), int(h * .20), int(w * .97), int(h * .80))),
        imagen.crop((int(w * .10), int(h * .30), int(w * .90), int(h * .70))),
        imagen.crop((int(w * .02), int(h * .35), int(w * .98), int(h * .65))),
    ]


def variantes(recorte):
    gris = ImageOps.grayscale(recorte)
    gris = ImageOps.autocontrast(gris, cutoff=2)
    gris = ImageEnhance.Contrast(gris).enhance(1.7)
    gris = ImageEnhance.Sharpness(gris).enhance(2.5)

    suave = gris.filter(ImageFilter.MedianFilter(size=3))

    hist = gris.histogram()
    pixeles = sum(hist)
    media = sum(i * c for i, c in enumerate(hist)) / pixeles if pixeles > 0 else 128
    umbral = max(80, min(190, media * 0.92))

    binaria = gris.point(lambda p, t=umbral: 255 if p > t else 0)
    binaria_invertida = ImageOps.invert(binaria)

    return [gris, suave, binaria, binaria_invertida]


def ocr_textos(imagen_pil):
    base = preparar_base(imagen_pil)
    cfg_char = "-c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-"
    textos = []

    for k, rec in enumerate(recortes(base)):
        psms = (6, 7, 11, 12, 13) if k == 0 else (6, 7, 8, 11, 13)

        for var in variantes(rec):
            for psm in psms:
                try:
                    texto = pytesseract.image_to_string(
                        var,
                        config=f"--oem 3 --psm {psm} {cfg_char}",
                        lang="eng"
                    )
                    if texto and texto.strip():
                        textos.append(texto.upper())
                except pytesseract.TesseractNotFoundError:
                    raise
                except Exception:
                    continue

    return textos


def extraer_cadenas(textos):
    votos = Counter()

    for texto in textos:
        texto = str(texto).upper()
        lineas = []

        for linea in texto.splitlines():
            limpia = re.sub(r"[^A-Z0-9]", "", linea)
            if limpia:
                lineas.append(limpia)
                if 4 <= len(limpia) <= 14:
                    votos[limpia] += 2

        for i in range(len(lineas) - 1):
            junto = lineas[i] + lineas[i + 1]
            if 4 <= len(junto) <= 14:
                votos[junto] += 1

        for token in re.split(r"[\s\-]+", texto):
            limpio = re.sub(r"[^A-Z0-9]", "", token)
            if 4 <= len(limpio) <= 14:
                votos[limpio] += 1

        for grupo in re.findall(r"[A-Z0-9]+(?:[-][A-Z0-9]+)+", texto):
            limpio = normalizar_busqueda(grupo)
            if 4 <= len(limpio) <= 14:
                votos[limpio] += 3

        compacto = re.sub(r"[^A-Z0-9]", "", texto)

        if 4 <= len(compacto) <= 14:
            votos[compacto] += 2

        for longitud in range(4, min(14, len(compacto)) + 1):
            for inicio in range(0, len(compacto) - longitud + 1):
                votos[compacto[inicio:inicio + longitud]] += 0.05

    return votos


def emparejar(votos, indice):
    if not votos or not indice:
        return []

    reales_canon = {}
    for norma in indice:
        reales_canon.setdefault(canonica(norma), []).append(norma)

    encontrados = {}

    def anotar(norma, score, metodo):
        if norma not in encontrados or score > encontrados[norma][0]:
            encontrados[norma] = (score, metodo)

    top = [c for c, _ in votos.most_common(60) if 4 <= len(c) <= 14]

    for candidato in top:
        if candidato in indice:
            anotar(candidato, 1.0 + min(votos[candidato], 5) * 0.005, "Coincidencia exacta")

    if encontrados:
        resultado = [
            {"norma": n, "mostrar": indice[n], "score": min(s, 1.0), "metodo": m, "en_concentrado": True}
            for n, (s, m) in encontrados.items()
        ]
        resultado.sort(key=lambda x: x["score"], reverse=True)
        return resultado[:5]

    for candidato in top:
        cc = canonica(candidato)
        frecuencia = min(votos[candidato], 5)
        for canon_real, lista in reales_canon.items():
            if cc == canon_real:
                for norma in lista:
                    anotar(norma, 0.96 + frecuencia * 0.005, "Corregida por caracteres parecidos")

    for candidato in top:
        cc = canonica(candidato)
        frecuencia = min(votos[candidato], 5)
        for canon_real, lista in reales_canon.items():
            if len(canon_real) < 5:
                continue
            if canon_real in cc and len(cc) - len(canon_real) <= 4:
                for norma in lista:
                    anotar(norma, 0.88 + frecuencia * 0.005, "Encontrada dentro del texto leído")

    if not any(score >= 0.94 for score, _ in encontrados.values()):
        for candidato in top[:30]:
            cc = canonica(candidato)
            if not 4 <= len(cc) <= 14:
                continue
            frecuencia = min(votos[candidato], 5)
            for canon_real, lista in reales_canon.items():
                if abs(len(cc) - len(canon_real)) > 2:
                    continue
                distancia = distancia_ocr(cc, canon_real, limite=2)
                if distancia <= 2:
                    score = 0.90 - distancia * 0.09 + frecuencia * 0.004
                    for norma in lista:
                        anotar(norma, score, f"Parecida por OCR ({distancia:g} diferencia{'s' if distancia != 1 else ''})")

    resultado = [
        {"norma": n, "mostrar": indice[n], "score": min(max(s, 0), 1), "metodo": m, "en_concentrado": True}
        for n, (s, m) in encontrados.items()
    ]
    resultado.sort(key=lambda r: (r["score"], r["norma"]), reverse=True)
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

    crudos, vistos = [], set()

    for candidato, _ in votos.most_common(30):
        if not 4 <= len(candidato) <= 14 or candidato in vistos:
            continue
        vistos.add(candidato)
        crudos.append(candidato)
        if len(crudos) >= 5:
            break

    return [
        {"norma": c, "mostrar": c, "score": 0.0, "metodo": "Leída por OCR (no está en el concentrado)", "en_concentrado": False}
        for c in crudos
    ], ""


# =========================================================
# OCR DE VIN / SERIE  (NUEVO)
# =========================================================

"""
Un VIN tiene 17 caracteres y NUNCA usa las letras I, O, Q.
Se lee la etiqueta (aunque venga girada), se buscan ventanas de 17
caracteres y se comparan contra la columna SERIE del concentrado.
El dígito verificador (posición 9) solo se usa como bonus: algunas
series de China/México no lo cumplen, por eso no se exige.
"""

WL_VIN = "ABCDEFGHJKLMNPRSTUVWXYZ0123456789"
TR_VIN = str.maketrans({"O": "0", "I": "1", "Q": "0"})

_VIN_VALORES = {str(i): i for i in range(10)}
for _letras, _v in (("AJ", 1), ("BKS", 2), ("CLT", 3), ("DMU", 4), ("ENV", 5),
                    ("FW", 6), ("GPX", 7), ("HY", 8), ("RZ", 9)):
    for _l in _letras:
        _VIN_VALORES[_l] = _v

_VIN_PESOS = [8, 7, 6, 5, 4, 3, 2, 10, 0, 9, 8, 7, 6, 5, 4, 3, 2]


def vin_valido(v):
    if len(v) != 17 or any(c not in _VIN_VALORES for c in v):
        return False
    total = sum(_VIN_VALORES[c] * p for c, p in zip(v, _VIN_PESOS)) % 11
    return v[8] == ("X" if total == 10 else str(total))


def cargar_indice_series(df):
    """Devuelve ({serie_normalizada: serie_original}, {serie_normalizada: placa})."""
    indice, placas = {}, {}

    if "SERIE" not in df.columns:
        return indice, placas

    col_placa = df["PLACA"] if "PLACA" in df.columns else pd.Series([None] * len(df), index=df.index)

    for s, p in zip(df["SERIE"], col_placa):
        norma = normalizar_busqueda(s)
        if not norma or len(norma) < 6 or len(norma) > 20 or norma in indice:
            continue
        indice[norma] = str(s).strip().upper()
        placas[norma] = valor_visible(p)

    return indice, placas


def sumar_vins(texto, votos):
    for linea in str(texto).upper().splitlines():
        compacta = re.sub(r"[^A-Z0-9]", "", linea).translate(TR_VIN)

        if len(compacta) < 15:
            continue

        for largo, peso in ((17, 1.0), (16, 0.3), (18, 0.3)):
            for i in range(len(compacta) - largo + 1):
                v = compacta[i:i + largo]
                bono = 2.0 if (largo == 17 and vin_valido(v)) else 0
                votos[v] += peso + bono


def hay_exacta_vin(votos, indice):
    return any(len(c) == 17 and c in indice for c in votos)


def variantes_vin(zona):
    v = variantes(zona)
    return [v[0], v[2], v[3]]  # gris, binaria, binaria invertida


def ocr_vin(imagen_pil, indice):
    cfg = f"--oem 3 -c tessedit_char_whitelist={WL_VIN}"
    img = ImageOps.exif_transpose(imagen_pil).convert("RGB")
    w, h = img.size

    # Etiquetas verticales: probar primero los giros de 90°
    angulos = [90, 270, 0] if h > w * 1.2 else [0, 90, 270]
    votos = Counter()

    for ang in angulos:
        rot = img.rotate(ang, expand=True) if ang else img
        base = preparar_base(rot)
        bw, bh = base.size
        zonas = [base, base.crop((0, int(bh * .15), bw, int(bh * .90)))]

        for zona in zonas:
            for var in variantes_vin(zona):
                for psm in (6, 11):
                    try:
                        texto = pytesseract.image_to_string(
                            var, config=f"{cfg} --psm {psm}", lang="eng"
                        )
                    except pytesseract.TesseractNotFoundError:
                        raise
                    except Exception:
                        continue

                    if texto and texto.strip():
                        sumar_vins(texto, votos)
                        if hay_exacta_vin(votos, indice):
                            return votos

    return votos


def emparejar_vin(votos, indice, placas):
    if not votos or not indice:
        return []

    encontrados = {}

    def anotar(norma, score, metodo):
        if norma not in encontrados or score > encontrados[norma][0]:
            encontrados[norma] = (score, metodo)

    def armar():
        res = [
            {"norma": n, "mostrar": indice[n], "placa": placas.get(n, ""),
             "score": min(max(s, 0), 1), "metodo": m, "en_concentrado": True}
            for n, (s, m) in encontrados.items()
        ]
        res.sort(key=lambda r: (r["score"], votos.get(r["norma"], 0)), reverse=True)
        return res[:5]

    # 1. Exacta
    for c in votos:
        if c in indice:
            anotar(c, 1.0, "Coincidencia exacta")

    if encontrados:
        return armar()

    top = [c for c, _ in votos.most_common(40) if 14 <= len(c) <= 19]

    canon_idx = {}
    suf_idx = {}
    for n in indice:
        canon_idx.setdefault(canonica(n), []).append(n)
        if len(n) >= 12:
            suf_idx.setdefault(n[-8:], []).append(n)

    # 2. Caracteres parecidos
    for c in top:
        for n in canon_idx.get(canonica(c), []):
            anotar(n, 0.96, "Corregida por caracteres parecidos")

    # 3. Últimos 8 caracteres (útil cuando el inicio está borroso)
    for c in top:
        if len(c) >= 17:
            for n in suf_idx.get(c[-8:], []):
                anotar(n, 0.86, "Coincide con los últimos 8 caracteres")

    # 4. Diferencias pequeñas
    if not any(s >= 0.94 for s, _ in encontrados.values()):
        for c in top[:25]:
            cc = canonica(c)
            for canon_real, lista in canon_idx.items():
                if abs(len(cc) - len(canon_real)) > 2:
                    continue

                if len(cc) == len(canon_real):
                    d = sum(1 for x, y in zip(cc, canon_real) if x != y)
                else:
                    d = distancia_limitada(cc, canon_real, limite=2)

                if d <= 2:
                    for n in lista:
                        anotar(n, 0.90 - d * 0.09, f"Parecida por OCR ({d} diferencia{'s' if d != 1 else ''})")

    return armar()


def leer_vin_foto(archivo, df):
    try:
        imagen = Image.open(archivo)
    except Exception:
        return [], "No se pudo abrir la imagen."

    indice, placas = cargar_indice_series(df)

    if not indice:
        return [], "El concentrado no tiene la columna SERIE."

    try:
        votos = ocr_vin(imagen, indice)
    except pytesseract.TesseractNotFoundError:
        return [], "Tesseract no está instalado en el servidor."
    except Exception as e:
        return [], f"No se pudo ejecutar OCR: {e}"

    resultados = emparejar_vin(votos, indice, placas)

    if resultados:
        return resultados, ""

    # Sin coincidencia: mostrar lo que realmente leyó
    candidatos = [c for c in votos if len(c) == 17]
    candidatos.sort(key=lambda c: (vin_valido(c), votos[c]), reverse=True)

    crudos = []
    for c in candidatos:
        if c not in crudos:
            crudos.append(c)
        if len(crudos) >= 5:
            break

    return [
        {"norma": c, "mostrar": c, "placa": "", "score": 0.0,
         "metodo": "Leído por OCR · dígito verificador " + ("✔" if vin_valido(c) else "✘") + " (no está en el concentrado)",
         "en_concentrado": False}
        for c in crudos
    ], ""


# =========================================================
# LECTOR UNIFICADO (PLACA / VIN)
# =========================================================

def usar_placa(texto):
    st.session_state["buscador_unidad"] = texto
    st.session_state["scan_on"] = False
    st.session_state.pop("ocr_resultados", None)
    st.session_state.pop("ocr_hash", None)


def mostrar_lector(df):
    if not st.session_state.get("scan_on"):
        return

    tipo = st.radio(
        "¿Qué quieres leer?",
        ["🔢 Placa", "🔑 VIN / Serie"],
        horizontal=True,
        key="scan_tipo"
    )
    es_vin = tipo.startswith("🔑")

    if es_vin:
        titulo = "🔑 Lector de VIN / Serie"
        ayuda = ("Enfoca la etiqueta del VIN (pilar de la puerta, tablero o parabrisas). "
                 "Funciona aunque la foto esté girada. Se compara contra la columna SERIE.")
    else:
        titulo = "📷 Lector de placas"
        ayuda = ("Acércate a la placa, que ocupe buena parte de la imagen y esté bien enfocada. "
                 "Lee letras, números y guiones en cualquier combinación.")

    st.markdown(
        f'<div class="scan-card"><div class="t">{titulo}</div><div class="s">{ayuda}</div></div>',
        unsafe_allow_html=True
    )

    modo = st.radio(
        "Origen de la foto",
        ["📸 Tomar foto", "🖼️ Subir foto"],
        horizontal=True,
        label_visibility="collapsed",
        key="scan_modo"
    )

    if modo == "📸 Tomar foto":
        archivo = st.camera_input("Toma la foto", key="foto_placa_camara")
    else:
        archivo = st.file_uploader(
            "Selecciona la foto",
            type=["jpg", "jpeg", "png", "webp"],
            key="foto_placa_archivo"
        )

    if archivo is None:
        return

    huella = hashlib.md5(archivo.getvalue()).hexdigest() + ("_vin" if es_vin else "_placa")

    if st.session_state.get("ocr_hash") != huella:
        with st.spinner("🔎 Leyendo VIN..." if es_vin else "🔎 Leyendo placa..."):
            resultados, error = leer_vin_foto(archivo, df) if es_vin else leer_placa_foto(archivo, df)

        st.session_state["ocr_hash"] = huella
        st.session_state["ocr_resultados"] = resultados
        st.session_state["ocr_error"] = error

    if st.session_state.get("ocr_error"):
        st.warning(st.session_state["ocr_error"])
        return

    resultados = st.session_state.get("ocr_resultados", [])

    if not resultados:
        st.warning(
            "No pude identificar el VIN. Intenta con más luz, sin reflejos y con la etiqueta más grande en la foto."
            if es_vin else
            "No pude identificar la placa. Intenta con más luz, sin reflejos y con la placa más grande en la foto."
        )
        return

    st.markdown(
        f'<div class="campo" style="margin-top:10px;">{"VIN detectados" if es_vin else "Placas detectadas"}</div>',
        unsafe_allow_html=True
    )

    for i, r in enumerate(resultados):
        c1, c2, c3 = st.columns([2.6, 2.4, 1.3], vertical_alignment="center")

        with c1:
            clase = "placa mini vin" if es_vin else "placa mini"
            st.markdown(f'<div class="{clase}">{r["mostrar"]}</div>', unsafe_allow_html=True)
            if es_vin and r.get("placa") and r["placa"] != "Sin información":
                st.markdown(f'<div class="conf">Placa: {r["placa"]}</div>', unsafe_allow_html=True)

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
                "Usar serie" if es_vin else "Usar placa",
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
    return (
        f'<div class="tarjeta"><div class="campo">{campo}</div>'
        f'<div class="{clase}">{valor}</div></div>'
    )


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
# PANEL DE GERENCIA: DATOS
# =========================================================

NOMBRES_FOLIOS = ["Gestoría", "Mantenimiento", "Ziga Talleres", "Siniestro", "Refacciones"]

EXPORT_COLS = [
    "PLACA", "SERIE", "LÍNEA", "MODELO", "ESTATUS", "SUBESTATUS", "CLIENTE", "ALMACÉN",
    "DÍAS EN PISO", "NIVEL", "FOLIOS ABIERTOS", "N FOLIOS",
    "GESTORÍA", "DÍAS GESTORÍA", "MANTENIMIENTO", "DÍAS MANTENIMIENTO",
    "ZIGA TALLERES", "DÍAS ZIGA", "EJECUTIVO ZIGA", "SINIESTRO", "DÍAS SINIESTRO",
    "REFACCIONES (FOLIO WS)",
]


def _col(d, nombre):
    return d[nombre] if nombre in d.columns else pd.Series([None] * len(d), index=d.index)


def _texto_folio(v):
    if not tiene_informacion(v):
        return ""
    try:
        f = float(v)
        if f.is_integer():
            return str(int(f))
    except Exception:
        pass
    return str(v).strip()


def _mapa_refacciones(df_ref):
    if df_ref is None or df_ref.empty:
        return {}

    cols = {str(c).strip().lower(): c for c in df_ref.columns}
    c_serie, c_folio, c_act = cols.get("serie"), cols.get("folio ws"), cols.get("actividad")

    if c_serie is None or c_folio is None or c_act is None:
        return {}

    mapa = {}

    for s, f, a in zip(df_ref[c_serie], df_ref[c_folio], df_ref[c_act]):
        if not (tiene_informacion(f) and tiene_informacion(a)):
            continue
        s = normalizar_texto(s)
        f = _texto_folio(f)
        if s and f not in mapa.setdefault(s, []):
            mapa[s].append(f)

    return mapa


@st.cache_data(show_spinner=False)
def construir_panel(df, df_ref):
    d = df.copy()

    serie_n = _col(d, "SERIE").apply(normalizar_texto)
    d = d[~(serie_n.duplicated() & (serie_n != ""))].copy()
    serie_n = _col(d, "SERIE").apply(normalizar_texto)

    p = pd.DataFrame(index=d.index)
    p["PLACA"] = _col(d, "PLACA").apply(valor_visible)
    p["SERIE"] = _col(d, "SERIE").apply(valor_visible)
    p["LÍNEA"] = _col(d, "LINEA").apply(valor_visible)
    p["MODELO"] = _col(d, "MODELO").apply(modelo_visible)
    p["ESTATUS"] = _col(d, "ESTATUS").apply(valor_visible)
    p["SUBESTATUS"] = _col(d, "SUBESTATUS").apply(valor_visible)
    p["CLIENTE"] = _col(d, "NOMBRE DEL CLIENTE").apply(valor_visible)
    p["ALMACÉN"] = _col(d, "ALMACÉN").apply(valor_visible)
    p["DÍAS EN PISO"] = pd.to_numeric(_col(d, "DÍAS EN PISO"), errors="coerce")

    p["GESTORÍA"] = _col(d, "Gestorías1.Folio").apply(_texto_folio)
    p["MANTENIMIENTO"] = _col(d, "Mantenimiento.FOLIO").apply(_texto_folio)
    p["ZIGA TALLERES"] = _col(d, "Ziga_Talleres.Folio").apply(_texto_folio)
    p["SINIESTRO"] = _col(d, "Siniestro").apply(_texto_folio)
    p["EJECUTIVO ZIGA"] = _col(d, "Ziga_Talleres.Ejecutivo responsable").apply(
        lambda v: str(v).strip() if tiene_informacion(v) else ""
    )

    mapa_ref = _mapa_refacciones(df_ref)
    p["REFACCIONES (FOLIO WS)"] = serie_n.map(lambda s: ", ".join(mapa_ref.get(s, [])))

    p["DÍAS GESTORÍA"] = pd.to_numeric(_col(d, "Gestorías1.Personalizado"), errors="coerce")
    p["DÍAS MANTENIMIENTO"] = pd.to_numeric(_col(d, "Mantenimiento.dias"), errors="coerce")
    p["DÍAS ZIGA"] = pd.to_numeric(_col(d, "Ziga_Talleres.dias"), errors="coerce")
    p["DÍAS SINIESTRO"] = pd.to_numeric(_col(d, "Siniestros2.dias"), errors="coerce")

    p["_g"] = p["GESTORÍA"] != ""
    p["_m"] = p["MANTENIMIENTO"] != ""
    p["_z"] = p["ZIGA TALLERES"] != ""
    p["_s"] = p["SINIESTRO"] != ""
    p["_r"] = p["REFACCIONES (FOLIO WS)"] != ""

    flags = list(zip(p["_g"], p["_m"], p["_z"], p["_s"], p["_r"]))
    p["FOLIOS ABIERTOS"] = [" · ".join(n for n, f in zip(NOMBRES_FOLIOS, fl) if f) or "—" for fl in flags]
    p["N FOLIOS"] = [sum(fl) for fl in flags]

    def nivel(x):
        if pd.isna(x):
            return "Sin dato"
        if x >= DIAS_ROJO:
            return "Crítica"
        if x >= DIAS_AMBAR:
            return "Atención"
        return "Normal"

    p["NIVEL"] = p["DÍAS EN PISO"].apply(nivel)
    p["_TXT"] = (p["PLACA"] + " " + p["SERIE"] + " " + p["LÍNEA"] + " " + p["CLIENTE"] + " " + p["ESTATUS"] + " " + p["EJECUTIVO ZIGA"]).str.upper()
    p["_NORM"] = (p["PLACA"] + p["SERIE"]).apply(normalizar_busqueda)

    return p.reset_index(drop=True)


def metricas(p):
    total = len(p)
    d = p["DÍAS EN PISO"]
    con = int((p["N FOLIOS"] > 0).sum())

    return {
        "total": total,
        "con": con,
        "sin": total - con,
        "multi": int((p["N FOLIOS"] >= 2).sum()),
        "normal": int((~(d >= DIAS_AMBAR)).sum()),
        "atencion": int(((d >= DIAS_AMBAR) & (d < DIAS_ROJO)).sum()),
        "criticas": int((d >= DIAS_ROJO).sum()),
        "gestoria": int(p["_g"].sum()),
        "mantenimiento": int(p["_m"].sum()),
        "ziga": int(p["_z"].sum()),
        "siniestro": int(p["_s"].sum()),
        "refacciones": int(p["_r"].sum()),
    }


def definicion(p, clave):
    d = p["DÍAS EN PISO"]
    todo = pd.Series(True, index=p.index)

    if clave == "estatus":
        est = st.session_state.get("detalle_estatus")
        return f"📊 Estatus: {est}", p["ESTATUS"] == est

    tabla = {
        "todas": ("🚗 Todas las unidades", todo),
        "con_folio": ("📂 Unidades con al menos un folio abierto", p["N FOLIOS"] > 0),
        "sin_folio": ("✅ Unidades sin folios abiertos", p["N FOLIOS"] == 0),
        "multi": ("🔀 Unidades con 2 o más folios abiertos", p["N FOLIOS"] >= 2),
        "normal": (f"✅ Rotación normal (menos de {DIAS_AMBAR} días)", ~(d >= DIAS_AMBAR)),
        "atencion": (f"⚠️ Requieren atención ({DIAS_AMBAR} a {DIAS_ROJO - 1} días en piso)", (d >= DIAS_AMBAR) & (d < DIAS_ROJO)),
        "criticas": (f"🚨 Críticas ({DIAS_ROJO}+ días en piso)", d >= DIAS_ROJO),
        "gestoria": ("📋 Unidades con folio de Gestoría", p["_g"]),
        "mantenimiento": ("🔧 Unidades con folio de Mantenimiento", p["_m"]),
        "ziga": ("🛠️ Unidades en Ziga Talleres", p["_z"]),
        "siniestro": ("🚨 Unidades con Siniestro", p["_s"]),
        "refacciones": ("📦 Unidades con Refacciones pendientes", p["_r"]),
    }

    return tabla.get(clave, tabla["todas"])


def exportable(sub):
    cols = [c for c in EXPORT_COLS if c in sub.columns]
    return sub.sort_values("DÍAS EN PISO", ascending=False, na_position="last")[cols]


def a_excel_bytes(hojas):
    from openpyxl.styles import Alignment, Font, PatternFill

    buf = io.BytesIO()

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for nombre, d in hojas.items():
            hoja = nombre[:31]
            d.to_excel(writer, sheet_name=hoja, index=False)
            ws = writer.sheets[hoja]

            for celda in ws[1]:
                celda.fill = PatternFill("solid", fgColor="213769")
                celda.font = Font(bold=True, color="FFFFFF")
                celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            for i, col in enumerate(d.columns, 1):
                largo = max([len(str(col))] + [len(str(v)) for v in d[col].head(300)]) if len(d) else len(str(col))
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(max(12, largo + 2), 42)

            ws.freeze_panes = "A2"
            if len(d):
                ws.auto_filter.ref = ws.dimensions

    return buf.getvalue()


@st.cache_data(show_spinner=False)
def reporte_ejecutivo(p):
    m = metricas(p)
    total = max(m["total"], 1)

    filas = [
        ("Total de unidades", m["total"]),
        ("Con al menos un folio abierto", m["con"]),
        ("Sin folios abiertos", m["sin"]),
        ("Con 2 o más folios abiertos", m["multi"]),
        ("Gestoría", m["gestoria"]),
        ("Mantenimiento", m["mantenimiento"]),
        ("Ziga Talleres", m["ziga"]),
        ("Siniestro", m["siniestro"]),
        ("Refacciones pendientes", m["refacciones"]),
        (f"Rotación normal (<{DIAS_AMBAR} días)", m["normal"]),
        (f"Atención ({DIAS_AMBAR}-{DIAS_ROJO - 1} días)", m["atencion"]),
        (f"Críticas ({DIAS_ROJO}+ días)", m["criticas"]),
    ]

    resumen = pd.DataFrame({
        "INDICADOR": [f[0] for f in filas],
        "UNIDADES": [f[1] for f in filas],
        "% DEL TOTAL": [round(f[1] / total * 100, 1) for f in filas],
    })

    est = p["ESTATUS"].value_counts().rename_axis("ESTATUS").reset_index(name="UNIDADES")

    hojas = {
        "Resumen": resumen,
        "Por estatus": est,
        "Todas": exportable(p),
        "Críticas": exportable(p[p["DÍAS EN PISO"] >= DIAS_ROJO]),
        "Atención": exportable(p[(p["DÍAS EN PISO"] >= DIAS_AMBAR) & (p["DÍAS EN PISO"] < DIAS_ROJO)]),
        "Gestoría": exportable(p[p["_g"]]),
        "Mantenimiento": exportable(p[p["_m"]]),
        "Ziga Talleres": exportable(p[p["_z"]]),
        "Siniestro": exportable(p[p["_s"]]),
        "Refacciones": exportable(p[p["_r"]]),
        "Sin folios": exportable(p[p["N FOLIOS"] == 0]),
    }

    return a_excel_bytes(hojas)


# =========================================================
# PANEL DE GERENCIA: INTERFAZ
# =========================================================

def abrir_detalle(clave):
    st.session_state["detalle"] = clave


def cerrar_detalle():
    st.session_state["detalle"] = None


def abrir_estatus(mapa):
    st.session_state["detalle_estatus"] = mapa.get(st.session_state.get("sel_estatus"))
    st.session_state["detalle"] = "estatus"


def abrir_ficha(valor):
    st.session_state["detalle"] = None
    st.session_state["buscador_unidad"] = valor


def limpiar_busqueda():
    st.session_state["buscador_unidad"] = ""


def kcard(col, clave, clase, etiqueta, numero, sub="", pct=None):
    barra = f'<div class="barra"><div style="width:{max(0, min(100, pct)):.0f}%"></div></div>' if pct is not None else ""
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""

    with col:
        st.markdown(
            f'<div class="{clase} kc"><div class="kpi-label">{etiqueta}</div>'
            f'<div class="kpi-numero">{numero:,}</div>{sub_html}{barra}</div>',
            unsafe_allow_html=True
        )
        st.button(
            "Ver unidades →",
            key=f"btn_{clave}",
            use_container_width=True,
            on_click=abrir_detalle,
            args=(clave,)
        )


def barras_h(d, cat, val, colores=None, orden=None, alto=250):
    y = alt.Y(f"{cat}:N", sort=orden if orden else "-x", title=None,
              axis=alt.Axis(labelFontWeight="bold", labelLimit=240, labelColor=AZUL_MARINO))

    base = alt.Chart(d).encode(y=y)

    if colores:
        color = alt.Color(f"{cat}:N", legend=None,
                          scale=alt.Scale(domain=list(colores.keys()), range=list(colores.values())))
    else:
        color = alt.value(AZUL_CASANOVA)

    barras = base.mark_bar(cornerRadiusEnd=6, size=22).encode(
        x=alt.X(f"{val}:Q", title=None, axis=None),
        color=color,
        tooltip=[alt.Tooltip(f"{cat}:N", title="Categoría"), alt.Tooltip(f"{val}:Q", title="Unidades")]
    )

    texto = base.mark_text(align="left", dx=6, fontWeight="bold", color=AZUL_MARINO).encode(
        x=alt.X(f"{val}:Q"), text=alt.Text(f"{val}:Q", format=",")
    )

    return (barras + texto).properties(height=alto, padding={"right": 40}).configure_view(strokeWidth=0)


def mostrar_graficas(p, m):
    titulo_seccion("📈 Gráficas")

    g1, g2 = st.columns(2)

    with g1:
        st.markdown('<div class="panel-t">Folios abiertos por área</div>', unsafe_allow_html=True)
        d = pd.DataFrame({
            "Area": NOMBRES_FOLIOS,
            "Unidades": [m["gestoria"], m["mantenimiento"], m["ziga"], m["siniestro"], m["refacciones"]],
        })
        st.altair_chart(barras_h(d, "Area", "Unidades"), use_container_width=True)

    with g2:
        st.markdown('<div class="panel-t">Unidades con y sin folio abierto</div>', unsafe_allow_html=True)
        d = pd.DataFrame({"Grupo": ["Con folio abierto", "Sin folios"], "Unidades": [m["con"], m["sin"]]})
        donut = alt.Chart(d).mark_arc(innerRadius=62, outerRadius=105).encode(
            theta=alt.Theta("Unidades:Q"),
            color=alt.Color("Grupo:N", legend=alt.Legend(orient="bottom", title=None),
                            scale=alt.Scale(domain=["Con folio abierto", "Sin folios"], range=[AMARILLO, VERDE])),
            tooltip=["Grupo:N", "Unidades:Q"]
        ).properties(height=250)
        st.altair_chart(donut, use_container_width=True)

    g3, g4 = st.columns(2)

    with g3:
        st.markdown('<div class="panel-t">Antigüedad en piso</div>', unsafe_allow_html=True)
        dias = p["DÍAS EN PISO"]
        cortes = [
            (f"0-{DIAS_AMBAR - 1} días", int((dias < DIAS_AMBAR).sum())),
            (f"{DIAS_AMBAR}-{DIAS_ROJO - 1} días", int(((dias >= DIAS_AMBAR) & (dias < DIAS_ROJO)).sum())),
            (f"{DIAS_ROJO}-89 días", int(((dias >= DIAS_ROJO) & (dias < 90)).sum())),
            ("90+ días", int((dias >= 90).sum())),
            ("Sin dato", int(dias.isna().sum())),
        ]
        d = pd.DataFrame(cortes, columns=["Rango", "Unidades"])
        colores = {cortes[0][0]: VERDE, cortes[1][0]: AMARILLO, cortes[2][0]: "#E8643C",
                   cortes[3][0]: "#A81E1E", cortes[4][0]: "#A7B1C2"}
        st.altair_chart(barras_h(d, "Rango", "Unidades", colores=colores, orden=[c[0] for c in cortes]),
                        use_container_width=True)

    with g4:
        st.markdown('<div class="panel-t">Unidades por estatus (top 8)</div>', unsafe_allow_html=True)
        e = p["ESTATUS"].value_counts().head(8).rename_axis("Estatus").reset_index(name="Unidades")
        st.altair_chart(barras_h(e, "Estatus", "Unidades"), use_container_width=True)


def pagina_resumen(p):
    m = metricas(p)
    total = max(m["total"], 1)

    def pct(n):
        return n / total * 100

    titulo_seccion("📊 Panel de gerencia")

    mayor = max(
        [("Gestoría", m["gestoria"]), ("Mantenimiento", m["mantenimiento"]), ("Ziga Talleres", m["ziga"]),
         ("Siniestro", m["siniestro"]), ("Refacciones", m["refacciones"])],
        key=lambda x: x[1]
    )

    st.markdown(
        f'<div class="headline">De <b>{m["total"]:,}</b> unidades, <b>{m["con"]:,}</b> '
        f'(<b>{pct(m["con"]):.0f}%</b>) tienen al menos un folio abierto y <b>{m["sin"]:,}</b> están libres. '
        f'La mayor carga está en <b>{mayor[0]}</b> con <b>{mayor[1]:,}</b> unidades. '
        f'<span class="rojo">{m["criticas"]:,} unidades llevan {DIAS_ROJO}+ días en piso</span> '
        f'y {m["atencion"]:,} más están entre {DIAS_AMBAR} y {DIAS_ROJO - 1} días.</div>',
        unsafe_allow_html=True
    )

    b1, b2 = st.columns([1, 1], vertical_alignment="center")

    with b1:
        st.download_button(
            "⬇️ Descargar reporte ejecutivo (Excel)",
            data=reporte_ejecutivo(p),
            file_name=f"reporte_ejecutivo_{datetime.now():%Y%m%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with b2:
        st.toggle("📈 Mostrar gráficas", key="graficas_on")

    st.caption("Toca **Ver unidades →** en cualquier tarjeta para abrir el listado con placas, series y folios abiertos.")

    # --- Panorama
    titulo_seccion("🧭 Panorama general")
    c = st.columns(4)
    kcard(c[0], "todas", "kpi-azul", "🚗 Total de unidades", m["total"], "Base completa")
    kcard(c[1], "con_folio", "kc-blanco", "📂 Con folio abierto", m["con"], f'{pct(m["con"]):.0f}% del total', pct(m["con"]))
    kcard(c[2], "sin_folio", "kpi-verde", "✅ Sin folios abiertos", m["sin"], f'{pct(m["sin"]):.0f}% del total', pct(m["sin"]))
    kcard(c[3], "multi", "kc-blanco", "🔀 Con 2 o más folios", m["multi"], f'{pct(m["multi"]):.0f}% del total', pct(m["multi"]))

    # --- Tiempo en piso
    titulo_seccion("⏱ Tiempo en piso")
    c = st.columns(3)
    kcard(c[0], "normal", "kpi-verde", f"✅ Rotación normal · menos de {DIAS_AMBAR} días", m["normal"], f'{pct(m["normal"]):.0f}% del total', pct(m["normal"]))
    kcard(c[1], "atencion", "kpi-ambar", f"⚠️ Atención · {DIAS_AMBAR} a {DIAS_ROJO - 1} días", m["atencion"], f'{pct(m["atencion"]):.0f}% del total', pct(m["atencion"]))
    kcard(c[2], "criticas", "kpi-rojo", f"🚨 Críticas · {DIAS_ROJO}+ días", m["criticas"], f'{pct(m["criticas"]):.0f}% del total', pct(m["criticas"]))

    # --- Folios por área
    titulo_seccion("📂 Folios abiertos por área")
    c = st.columns(5)
    kcard(c[0], "gestoria", "kc-blanco", "📋 Gestoría", m["gestoria"], f'{pct(m["gestoria"]):.0f}% del total', pct(m["gestoria"]))
    kcard(c[1], "mantenimiento", "kc-blanco", "🔧 Mantenimiento", m["mantenimiento"], f'{pct(m["mantenimiento"]):.0f}% del total', pct(m["mantenimiento"]))
    kcard(c[2], "ziga", "kc-blanco", "🛠️ Ziga Talleres", m["ziga"], f'{pct(m["ziga"]):.0f}% del total', pct(m["ziga"]))
    kcard(c[3], "siniestro", "kc-blanco", "🚨 Siniestro", m["siniestro"], f'{pct(m["siniestro"]):.0f}% del total', pct(m["siniestro"]))
    kcard(c[4], "refacciones", "kc-blanco", "📦 Refacciones", m["refacciones"], f'{pct(m["refacciones"]):.0f}% del total', pct(m["refacciones"]))

    # --- Estatus
    titulo_seccion("📊 Unidades por estatus")
    conteo = p["ESTATUS"].value_counts()
    p1, p2 = st.columns([1.5, 1])

    with p1:
        if not conteo.empty:
            maximo = conteo.max()
            filas = "".join(
                f'<div class="fila-b"><div class="n">{n}</div>'
                f'<div class="b"><div style="width:{c_ / maximo * 100:.0f}%"></div></div>'
                f'<div class="c">{c_:,}</div></div>'
                for n, c_ in conteo.head(12).items()
            )
            st.markdown(f'<div class="panel"><div class="panel-t">Distribución actual</div>{filas}</div>', unsafe_allow_html=True)

    with p2:
        st.markdown('<div class="panel"><div class="panel-t">Consultar un estatus</div>', unsafe_allow_html=True)
        mapa = {f"{n} ({c_:,})": n for n, c_ in conteo.items()}
        st.selectbox("Elige un estatus", list(mapa.keys()), key="sel_estatus")
        st.button("Ver unidades y folios →", key="btn_estatus", use_container_width=True,
                  on_click=abrir_estatus, args=(mapa,))
        st.markdown("</div>", unsafe_allow_html=True)

    if st.session_state.get("graficas_on"):
        mostrar_graficas(p, m)


def pagina_detalle(p):
    clave = st.session_state.get("detalle")
    titulo, mascara = definicion(p, clave)
    sub = p[mascara].copy()

    st.button("← Volver al panel", key="btn_volver", on_click=cerrar_detalle)

    st.markdown(
        f'<div class="seccion-head"><div class="titulo-seccion">{titulo}</div>'
        f'<div class="dias-pill dias-azul">{len(sub):,} unidades</div></div>',
        unsafe_allow_html=True
    )

    f1, f2 = st.columns([3, 1.3], vertical_alignment="center")

    with f1:
        filtro = st.text_input(
            "Filtrar",
            placeholder="Filtra esta lista por placa, serie, línea, cliente o estatus...",
            label_visibility="collapsed",
            key=f"filtro_{clave}_{st.session_state.get('detalle_estatus', '')}"
        )

    with f2:
        ver_num = st.toggle("Ver números de folio", key="ver_folios")

    if filtro.strip():
        t = filtro.strip().upper()
        n = normalizar_busqueda(filtro)
        m = sub["_TXT"].str.contains(t, regex=False, na=False)
        if n:
            m = m | sub["_NORM"].str.contains(n, regex=False, na=False)
        sub = sub[m]

    sub = sub.sort_values("DÍAS EN PISO", ascending=False, na_position="last")

    if sub.empty:
        st.warning("No hay unidades que coincidan.")
        return

    cols = ["PLACA", "SERIE", "LÍNEA", "MODELO", "ESTATUS", "DÍAS EN PISO", "FOLIOS ABIERTOS"]
    if clave == "ziga":
        cols.insert(cols.index("FOLIOS ABIERTOS") + 1, "EJECUTIVO ZIGA")

    if ver_num:
        cols += ["GESTORÍA", "MANTENIMIENTO", "ZIGA TALLERES", "SINIESTRO", "REFACCIONES (FOLIO WS)"]
        if "EJECUTIVO ZIGA" not in cols:
            cols.append("EJECUTIVO ZIGA")

    maximo = int(max(90, sub["DÍAS EN PISO"].max())) if sub["DÍAS EN PISO"].notna().any() else 90

    st.dataframe(
        sub[cols],
        hide_index=True,
        use_container_width=True,
        height=min(640, 40 + 35 * len(sub)),
        column_config={
            "DÍAS EN PISO": st.column_config.ProgressColumn(
                "DÍAS EN PISO", min_value=0, max_value=maximo, format="%d"
            ),
            "FOLIOS ABIERTOS": st.column_config.TextColumn("FOLIOS ABIERTOS", width="large"),
        }
    )

    d1, d2 = st.columns(2)
    export = exportable(sub)
    nombre = f"{re.sub(r'[^a-z0-9]+', '_', str(clave).lower())}_{datetime.now():%Y%m%d}"

    with d1:
        st.download_button(
            "⬇️ Descargar esta lista (Excel)",
            data=a_excel_bytes({"Unidades": export}),
            file_name=f"{nombre}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with d2:
        st.download_button(
            "⬇️ Descargar esta lista (CSV)",
            data=export.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"{nombre}.csv",
            mime="text/csv",
            use_container_width=True
        )

    separador()
    titulo_seccion("🔎 Abrir la ficha de una unidad de esta lista")

    opciones = {f"{r['PLACA']} | {r['SERIE']}": r["SERIE"] for _, r in sub.head(500).iterrows()}
    o1, o2 = st.columns([3, 1.3], vertical_alignment="bottom")

    with o1:
        sel = st.selectbox("Unidad", list(opciones.keys()), key=f"sel_ficha_{clave}")

    with o2:
        st.button("Abrir ficha →", key="btn_abrir_ficha", use_container_width=True,
                  on_click=abrir_ficha, args=(opciones[sel],))


# =========================================================
# ENCABEZADO
# =========================================================

col1, col2 = st.columns([5, 1], vertical_alignment="center")

with col1:
    fb = fecha_base()
    actualizado = f' · Base actualizada: {fb}' if fb else ""

    st.markdown(
        f"""
<div class="topbar">
<div class="logo-c">{LOGO_SVG}</div>
<div class="topbar-txt">
<div class="topbar-titulo">CASANOVA <b>|</b> UNIDADES EN CENTENARIO Y LALA</div>
<div class="topbar-sub"><span class="sede">📍 Centenario</span><span class="sede">📍 Lala</span> Consulta individual de unidades{actualizado}</div>
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
# BUSCADOR + OPCIONES
# =========================================================

titulo_seccion("🔎 Buscar unidad")

busqueda = st.text_input(
    "Busca por PLACA o SERIE",
    placeholder="Ejemplo: SKD-S4S5, número de serie / VIN completo o solo los últimos dígitos...",
    label_visibility="collapsed",
    key="buscador_unidad"
)

t1, t2 = st.columns(2)

with t1:
    st.toggle(
        "📷 Escanear placa o VIN con foto o cámara",
        key="scan_on",
        help="La cámara y la carga de fotos aparecen solo al activar."
    )

with t2:
    st.toggle(
        "📊 Panel de gerencia",
        key="panel_on",
        help="Resumen general, folios abiertos, unidades críticas y gráficas."
    )

mostrar_lector(df)


# =========================================================
# TOTALES
# =========================================================

if "SERIE" in df.columns:
    s = df["SERIE"].astype("string").fillna("").str.strip().str.upper()
    total_unidades = s[s != ""].nunique()
else:
    total_unidades = 0


def usar_reciente(p):
    st.session_state["buscador_unidad"] = p


# =========================================================
# INICIO / PANEL
# =========================================================

if not busqueda:

    if not st.session_state.get("panel_on"):
        recientes = st.session_state.get("recientes", [])

        if recientes:
            st.markdown('<div class="campo" style="margin-top:14px;">🕘 Consultadas recientemente</div>', unsafe_allow_html=True)
            for col, (i, r) in zip(st.columns(len(recientes)), enumerate(recientes)):
                with col:
                    st.button(r, key=f"rec_{i}", use_container_width=True, on_click=usar_reciente, args=(r,))

        st.info(
            "Escribe una placa o serie, escanea una placa o VIN con la cámara, "
            "o activa el **Panel de gerencia** para ver el resumen general."
        )
        st.stop()

    panel_df = construir_panel(df, df_refacciones)

    if st.session_state.get("detalle"):
        pagina_detalle(panel_df)
    else:
        pagina_resumen(panel_df)

    st.stop()


# =========================================================
# BÚSQUEDA
# =========================================================

criterio = normalizar_busqueda(busqueda)

# Si pegan "VIN LSGHD52H8ND038393", quitar el prefijo
if criterio.startswith("VIN") and len(criterio) >= 20:
    criterio = criterio[3:]

placas = df["PLACA"].apply(normalizar_busqueda) if "PLACA" in df.columns else pd.Series("", index=df.index)
series = df["SERIE"].apply(normalizar_busqueda) if "SERIE" in df.columns else pd.Series("", index=df.index)

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


# =========================================================
# HISTORIAL RECIENTE
# =========================================================

rec = st.session_state.get("recientes", [])

if placa != "Sin información":
    rec = [placa] + [r for r in rec if r != placa]
    st.session_state["recientes"] = rec[:5]

st.button("✖ Limpiar búsqueda", key="btn_limpiar", on_click=limpiar_busqueda)


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
texto_nivel = {"verde": "✅ Rotación normal", "ambar": "⚠️ Requiere atención", "rojo": "🚨 Tiempo crítico en piso", None: ""}[nivel_piso]
sub_html = f'<div class="kpi-sub">{texto_nivel}</div>' if texto_nivel else ""

barra_html = ""

if dias_n is not None and nivel_piso in ("verde", "ambar", "rojo"):
    pct_barra = min(100, dias_n / DIAS_ROJO * 100)
    barra_html = f'<div class="barra"><div style="width:{pct_barra:.0f}%"></div></div>'

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

ejecutivo = valor_visible(unidad.get("Ziga_Talleres.Ejecutivo responsable"))
clase_ejecutivo = "unidad-serie" + (" valor-vacio" if ejecutivo == "Sin información" else "")

st.markdown(
    f"""
<div class="unidad-header">
<div class="placa">{placa}<small>CASANOVA · MÉXICO</small></div>
<div>
<div class="unidad-serie-label">NÚMERO DE SERIE</div>
<div class="unidad-serie">{serie}</div>
</div>
<div>
<div class="unidad-serie-label">👤 EJECUTIVO RESPONSABLE</div>
<div class="{clase_ejecutivo}">{ejecutivo}</div>
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
# FOLIOS ABIERTOS
# =========================================================

hay_folio_abierto = False

# --- Gestoría
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

# --- Mantenimiento
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

# --- Ziga Talleres
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

# --- Siniestro
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

# --- Refacciones
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

ficha_df = pd.DataFrame({
    "CAMPO": ficha_serie.index.astype(str),
    "VALOR": ficha_serie.astype(str).values
})

st.download_button(
    "⬇️ Descargar ficha de la unidad (CSV)",
    data=ficha_df.to_csv(index=False).encode("utf-8-sig"),
    file_name=f"ficha_{normalizar_busqueda(placa)}.csv",
    mime="text/csv"
)

st.caption("Casanova | Unidades en Centenario y Lala")	
