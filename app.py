import streamlit as st
import pandas as pd
from pathlib import Path


# =========================================================
# CONFIGURACIÓN
# =========================================================

st.set_page_config(
    page_title="Casanova | Control de Unidades",
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
# ESTILO GLOBAL (se aplica también en el login)
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
    color: rgba(255,255,255,0.75);
    font-size: 13px;
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
    font-weight: 700;
    letter-spacing: 0.8px;
    text-transform: uppercase;
    margin-bottom: 6px;
}}

.valor {{
    color: #172033;
    font-size: 16px;
    font-weight: 700;
    word-break: break-word;
    line-height: 1.3;
}}

.valor-vacio {{
    color: #A7B1C2;
    font-weight: 500;
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
    font-weight: 700;
}}

.badge-verde {{ background: #E3F6EA; color: #167A3E; }}
.badge-amarillo {{ background: #FFF3D0; color: #8A6100; }}
.badge-rojo {{ background: #FDE5E5; color: #B42323; }}
.badge-azul {{ background: #E3ECF9; color: {AZUL_MARINO}; }}

/* =================================================
   KPIs (TOTAL Y DÍAS EN PISO)
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
    font-weight: 700;
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
    font-weight: 700;
    letter-spacing: 0.8px;
}}

.unidad-serie {{
    color: {AZUL_MARINO};
    font-size: 17px;
    font-weight: 700;
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
    font-weight: 500;
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
    font-weight: 700;
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

.stTextInput input {{
    border-radius: 12px !important;
    border: 2px solid #D8E1EF !important;
    padding: 12px 16px !important;
    font-size: 16px !important;
    background: white !important;
}}

.stTextInput input:focus {{
    border-color: {AZUL_CASANOVA} !important;
    box-shadow: 0 0 0 3px rgba(0,72,118,0.12) !important;
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

.stSelectbox > div > div {{
    border-radius: 12px !important;
}}

/* =================================================
   LOGIN
   ================================================= */

.login-box {{
    max-width: 430px;
    margin: 50px auto 10px auto;
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
    margin-top: 4px;
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


# ---------------------------------------------------------
# TARJETA (ahora con estilo para "Sin información")
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


# ---------------------------------------------------------
# TARJETA CON BADGE (para estatus)
# ---------------------------------------------------------

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


def titulo_seccion(texto):

    st.markdown(
        f'<div class="titulo-seccion">{texto}</div>',
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

titulo_seccion("🔎 Buscar unidad")

busqueda = st.text_input(
    "Busca por PLACA o SERIE",
    placeholder="Ejemplo: PDN1295 o número de serie...",
    label_visibility="collapsed"
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

criterio = normalizar_texto(busqueda)


placas = (
    df["PLACA"]
    .astype("string")
    .fillna("")
    .str.strip()
    .str.upper()
)


series = (
    df["SERIE"]
    .astype("string")
    .fillna("")
    .str.strip()
    .str.upper()
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
# KPIs: TOTAL + DÍAS EN PISO (LADO A LADO)
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

    st.markdown(
        f"""
<div class="kpi-amarillo" style="margin-top:14px;">
<div class="kpi-label">🏭 Días en piso</div>
<div class="kpi-numero">{texto_dias}</div>
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
# INFORMACIÓN GENERAL
# =========================================================

titulo_seccion("📋 Información general")


col1, col2, col3, col4 = st.columns(4)


with col1:

    st.markdown(
        tarjeta(
            "LÍNEA",
            valor_visible(
                unidad.get("LINEA")
            )
        ),
        unsafe_allow_html=True
    )


with col2:

    st.markdown(
        tarjeta(
            "MODELO",
            modelo_visible(
                unidad.get("MODELO")
            )
        ),
        unsafe_allow_html=True
    )


with col3:

    st.markdown(
        tarjeta_badge(
            "ESTATUS",
            valor_visible(
                unidad.get("ESTATUS")
            )
        ),
        unsafe_allow_html=True
    )


with col4:

    st.markdown(
        tarjeta(
            "SUBESTATUS",
            valor_visible(
                unidad.get("SUBESTATUS")
            )
        ),
        unsafe_allow_html=True
    )


# =========================================================
# DATOS DE LA UNIDAD
# =========================================================

titulo_seccion("🚘 Datos de la unidad")


col1, col2, col3 = st.columns(3)


with col1:

    st.markdown(
        tarjeta(
            "NOMBRE DEL CLIENTE",
            valor_visible(
                unidad.get(
                    "NOMBRE DEL CLIENTE"
                )
            )
        ),
        unsafe_allow_html=True
    )


with col2:

    st.markdown(
        tarjeta(
            "ALMACÉN",
            valor_visible(
                unidad.get(
                    "ALMACÉN"
                )
            )
        ),
        unsafe_allow_html=True
    )


with col3:

    st.markdown(
        tarjeta(
            "UNIDAD DE NEGOCIO",
            valor_visible(
                unidad.get(
                    "UNIDAD NEGOCIO"
                )
            )
        ),
        unsafe_allow_html=True
    )


col1, col2, col3 = st.columns(3)


with col1:

    st.markdown(
        tarjeta(
            "GPS",
            valor_visible(
                unidad.get(
                    "GPS"
                )
            )
        ),
        unsafe_allow_html=True
    )


with col2:

    st.markdown(
        tarjeta(
            "PLATAFORMA",
            valor_visible(
                unidad.get(
                    "PLATAFORMA"
                )
            )
        ),
        unsafe_allow_html=True
    )


with col3:

    st.markdown(
        tarjeta(
            "¿REPORTA?",
            valor_visible(
                unidad.get(
                    "¿REPORTA?"
                )
            )
        ),
        unsafe_allow_html=True
    )


# =========================================================
# CONTROL DE FOLIOS ABIERTOS
# =========================================================

hay_folio_abierto = False


# =========================================================
# GESTORÍA
# =========================================================

folio_gestoria = unidad.get(
    "Gestorías1.Folio"
)


if tiene_informacion(folio_gestoria):

    hay_folio_abierto = True

    separador()

    titulo_seccion("📋 Gestoría")

    col1, col2, col3 = st.columns(3)

    with col1:

        st.markdown(
            tarjeta(
                "FOLIO",
                valor_visible(
                    unidad.get(
                        "Gestorías1.Folio"
                    )
                )
            ),
            unsafe_allow_html=True
        )

    with col2:

        st.markdown(
            tarjeta_badge(
                "ESTATUS DE PÓLIZA",
                valor_visible(
                    unidad.get(
                        "Estatus poliza"
                    )
                )
            ),
            unsafe_allow_html=True
        )

    with col3:

        st.markdown(
            tarjeta(
                "FECHA DE CREACIÓN",
                fecha_visible(
                    unidad.get(
                        "Gestorías1.Fecha de creación"
                    )
                )
            ),
            unsafe_allow_html=True
        )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.markdown(
            tarjeta_badge(
                "ESTADO DE ATENCIÓN",
                valor_visible(
                    unidad.get(
                        "Gestorías1.Estado de atención"
                    )
                )
            ),
            unsafe_allow_html=True
        )

    with col2:

        st.markdown(
            tarjeta(
                "DÍAS",
                numero_entero(
                    unidad.get(
                        "Gestorías1.Personalizado"
                    )
                )
            ),
            unsafe_allow_html=True
        )

    with col3:

        st.markdown(
            tarjeta(
                "CONCEPTOS",
                valor_visible(
                    unidad.get(
                        "Gestorías1.Conceptos"
                    )
                )
            ),
            unsafe_allow_html=True
        )


# =========================================================
# MANTENIMIENTO
# =========================================================

folio_mantenimiento = unidad.get(
    "Mantenimiento.FOLIO"
)


if tiene_informacion(folio_mantenimiento):

    hay_folio_abierto = True

    separador()

    titulo_seccion("🔧 Mantenimiento")

    col1, col2, col3 = st.columns(3)

    with col1:

        st.markdown(
            tarjeta(
                "FOLIO",
                valor_visible(
                    unidad.get(
                        "Mantenimiento.FOLIO"
                    )
                )
            ),
            unsafe_allow_html=True
        )

    with col2:

        st.markdown(
            tarjeta(
                "FECHA DE CREACIÓN",
                fecha_visible(
                    unidad.get(
                        "Mantenimiento.FECHA DE CREACION DEL FOLIO"
                    )
                )
            ),
            unsafe_allow_html=True
        )

    with col3:

        st.markdown(
            tarjeta_badge(
                "ESTATUS",
                valor_visible(
                    unidad.get(
                        "Mantenimiento.ESTATUS DEL FOLIO"
                    )
                )
            ),
            unsafe_allow_html=True
        )

    col1, col2 = st.columns(2)

    with col1:

        st.markdown(
            tarjeta(
                "PROVEEDOR",
                valor_visible(
                    unidad.get(
                        "Mantenimiento.PROVEEDOR"
                    )
                )
            ),
            unsafe_allow_html=True
        )

    with col2:

        st.markdown(
            tarjeta(
                "DÍAS",
                numero_entero(
                    unidad.get(
                        "Mantenimiento.dias"
                    )
                )
            ),
            unsafe_allow_html=True
        )


# =========================================================
# ZIGA TALLERES
# =========================================================

folio_ziga = unidad.get(
    "Ziga_Talleres.Folio"
)


if tiene_informacion(folio_ziga):

    hay_folio_abierto = True

    separador()

    titulo_seccion("🛠️ Ziga Talleres")

    col1, col2, col3 = st.columns(3)

    with col1:

        st.markdown(
            tarjeta(
                "FOLIO",
                valor_visible(
                    unidad.get(
                        "Ziga_Talleres.Folio"
                    )
                )
            ),
            unsafe_allow_html=True
        )

    with col2:

        st.markdown(
            tarjeta_badge(
                "ESTATUS",
                valor_visible(
                    unidad.get(
                        "Ziga_Talleres.Estatus"
                    )
                )
            ),
            unsafe_allow_html=True
        )

    with col3:

        st.markdown(
            tarjeta(
                "ÁREA DEL TALLER ACTIVA",
                valor_visible(
                    unidad.get(
                        "Ziga_Talleres.Área del taller activa"
                    )
                )
            ),
            unsafe_allow_html=True
        )

    col1, col2 = st.columns(2)

    with col1:

        st.markdown(
            tarjeta(
                "FECHA DE INGRESO",
                fecha_visible(
                    unidad.get(
                        "Ziga_Talleres.Fecha de ingreso"
                    )
                )
            ),
            unsafe_allow_html=True
        )

    with col2:

        st.markdown(
            tarjeta(
                "DÍAS",
                numero_entero(
                    unidad.get(
                        "Ziga_Talleres.dias"
                    )
                )
            ),
            unsafe_allow_html=True
        )


# =========================================================
# SINIESTRO
# =========================================================

folio_siniestro = unidad.get(
    "Siniestro"
)


if tiene_informacion(folio_siniestro):

    hay_folio_abierto = True

    separador()

    titulo_seccion("🚨 Siniestro")

    col1, col2, col3 = st.columns(3)

    with col1:

        st.markdown(
            tarjeta(
                "CATEGORÍA",
                valor_visible(
                    unidad.get(
                        "Siniestros2.CATEGORIA"
                    )
                )
            ),
            unsafe_allow_html=True
        )

    with col2:

        st.markdown(
            tarjeta(
                "ETAPA",
                valor_visible(
                    unidad.get(
                        "Siniestros2.ETAPA"
                    )
                )
            ),
            unsafe_allow_html=True
        )

    with col3:

        st.markdown(
            tarjeta(
                "DÍAS",
                numero_entero(
                    unidad.get(
                        "Siniestros2.dias"
                    )
                )
            ),
            unsafe_allow_html=True
        )

    col1, col2 = st.columns(2)

    with col1:

        st.markdown(
            tarjeta(
                "FECHA DE APERTURA",
                fecha_visible(
                    unidad.get(
                        "Siniestros2.FECHA APERTURA"
                    )
                )
            ),
            unsafe_allow_html=True
        )

    with col2:

        st.markdown(
            tarjeta(
                "FECHA DEL SINIESTRO",
                fecha_visible(
                    unidad.get(
                        "Siniestros2.FECHA SINIESTRO"
                    )
                )
            ),
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

        serie_actual = normalizar_texto(
            serie
        )

        ref = ref[
            ref["_SERIE_BUSQUEDA"]
            == serie_actual
        ].copy()

        for folio, grupo in ref.groupby(
            col_folio_ref,
            sort=False
        ):

            if not tiene_informacion(folio):

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
                    ==
                    existente.casefold()
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
# PIE
# =========================================================

separador()

st.caption(
    "Casanova | Control de Unidades"
)
