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
# ARCHIVO BASE
# =========================================================

ARCHIVO_BASE = Path(
    r"C:\Users\epacheco\Music\Nuevo proyecto\concentrado.xlsx"
)

HOJA_BASE = "concentrado"


# =========================================================
# COLORES CASANOVA
# =========================================================

AZUL_MARINO = "#213769"
AZUL_CASANOVA = "#004876"
AMARILLO = "#FFB700"
GRIS_NIEBLA = "#EEF2F8"
BLANCO = "#FFFFFF"


# =========================================================
# CARGAR EXCEL
# =========================================================

@st.cache_data
def cargar_datos():

    df = pd.read_excel(
        ARCHIVO_BASE,
        sheet_name=HOJA_BASE
    )

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    return df


# =========================================================
# FUNCIONES
# =========================================================

def valor_visible(valor):

    if pd.isna(valor):
        return "Sin información"

    texto = str(valor).strip()

    if texto == "":
        return "Sin información"

    return texto


def tiene_informacion(valor):

    if pd.isna(valor):
        return False

    texto = str(valor).strip()

    return texto != ""


def fecha_visible(valor):

    if pd.isna(valor):
        return "Sin información"

    try:

        return pd.to_datetime(valor).strftime("%d/%m/%Y")

    except:

        return valor_visible(valor)


def dias_visible(valor):

    if pd.isna(valor):
        return "Sin información"

    try:

        numero = float(valor)

        if numero.is_integer():

            return f"{int(numero):,} días"

        return f"{numero:,.1f} días"

    except:

        return valor_visible(valor)


# =========================================================
# ESTILO
# =========================================================

st.markdown(
    f"""
    <style>

    /* Fondo general */
    .stApp {{
        background-color: #F7F9FC;
    }}

    /* Título */
    .titulo-principal {{
        color: {AZUL_MARINO};
        font-size: 30px;
        font-weight: 800;
        margin-bottom: 0px;
    }}

    .subtitulo {{
        color: #64748B;
        font-size: 14px;
        margin-bottom: 20px;
    }}

    /* Títulos de sección */
    .titulo-seccion {{
        color: {AZUL_MARINO};
        font-size: 21px;
        font-weight: 750;
        margin-top: 22px;
        margin-bottom: 12px;
    }}

    /* Tarjetas */
    .tarjeta {{
        background: {BLANCO};
        border: 1px solid #D8E1EF;
        border-radius: 14px;
        padding: 16px 18px;
        min-height: 78px;
        box-shadow: 0 2px 8px rgba(33,55,105,0.06);
        margin-bottom: 12px;
    }}

    .tarjeta-azul {{
        background: {AZUL_CASANOVA};
        border: none;
        color: white;
        border-radius: 14px;
        padding: 18px;
        box-shadow: 0 3px 10px rgba(0,72,118,0.18);
    }}

    .tarjeta-piso {{
        background: {BLANCO};
        border: 2px solid {AMARILLO};
        border-radius: 14px;
        padding: 20px;
        text-align: center;
        box-shadow: 0 3px 10px rgba(255,183,0,0.12);
    }}

    .campo {{
        color: #64748B;
        font-size: 12px;
        font-weight: 600;
        margin-bottom: 5px;
    }}

    .valor {{
        color: #172033;
        font-size: 17px;
        font-weight: 700;
        word-break: break-word;
    }}

    .valor-blanco {{
        color: white;
        font-size: 17px;
        font-weight: 700;
    }}

    .dias-grande {{
        color: {AZUL_MARINO};
        font-size: 34px;
        font-weight: 800;
    }}

    .dias-label {{
        color: #64748B;
        font-size: 13px;
        font-weight: 600;
    }}

    /* Tarjeta total */
    .total-piso {{
        background: linear-gradient(
            135deg,
            {AZUL_MARINO},
            {AZUL_CASANOVA}
        );
        color: white;
        border-radius: 16px;
        padding: 20px 24px;
        margin-bottom: 25px;
        box-shadow: 0 4px 14px rgba(33,55,105,0.18);
    }}

    .total-piso-titulo {{
        font-size: 14px;
        font-weight: 600;
        opacity: 0.9;
    }}

    .total-piso-numero {{
        font-size: 38px;
        font-weight: 800;
        margin-top: 4px;
    }}

    /* Estado */
    .estado-si {{
        background: #EAF8F0;
        color: #198754;
        border: 1px solid #BFE8CE;
        border-radius: 10px;
        padding: 10px 14px;
        font-weight: 700;
        margin-bottom: 12px;
    }}

    .estado-no {{
        background: #FFF7E0;
        color: #A66B00;
        border: 1px solid #FFE19A;
        border-radius: 10px;
        padding: 10px 14px;
        font-weight: 700;
        margin-bottom: 12px;
    }}

    /* Línea divisoria */
    .separador {{
        height: 2px;
        background: #DCE4F0;
        margin: 28px 0;
        border-radius: 5px;
    }}

    /* Search */
    div[data-baseweb="input"] {{
        border-radius: 10px;
    }}

    </style>
    """,
    unsafe_allow_html=True
)


# =========================================================
# ENCABEZADO
# =========================================================

st.markdown(
    '<div class="titulo-principal">🚗 CASANOVA | CONTROL DE UNIDADES</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitulo">Consulta individual de unidades</div>',
    unsafe_allow_html=True
)


# =========================================================
# CARGAR INFORMACIÓN
# =========================================================

try:

    df = cargar_datos()

except Exception as e:

    st.error(
        "No se pudo cargar el archivo concentrado.xlsx"
    )

    st.code(
        r"C:\Users\epacheco\Music\Nuevo proyecto\concentrado.xlsx"
    )

    st.code(str(e))

    st.stop()


# =========================================================
# TOTAL DE UNIDADES EN PISO
# =========================================================

columna_dias = "DÍAS EN PISO"

if columna_dias in df.columns:

    unidades_en_piso = (
        df[columna_dias]
        .notna()
        .sum()
    )

else:

    unidades_en_piso = 0


st.markdown(
    f"""
    <div class="total-piso">
        <div class="total-piso-titulo">
            🏭 TOTAL DE UNIDADES EN PISO
        </div>
        <div class="total-piso-numero">
            {unidades_en_piso:,}
        </div>
    </div>
    """,
    unsafe_allow_html=True
)


# =========================================================
# BUSCADOR
# =========================================================

st.markdown(
    '<div class="titulo-seccion">🔎 Buscar unidad</div>',
    unsafe_allow_html=True
)

busqueda = st.text_input(
    "Busca por PLACA o SERIE",
    placeholder="Ejemplo: PDN1295 o número de serie..."
)


# =========================================================
# SIN BÚSQUEDA
# =========================================================

if not busqueda:

    st.info(
        "Escribe una placa o número de serie para consultar la unidad."
    )

    st.stop()


# =========================================================
# PREPARAR BÚSQUEDA
# =========================================================

criterio = busqueda.strip().upper()


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
# VARIAS UNIDADES
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

    indice_seleccionado = opciones.index(
        seleccion
    )

    unidad = resultado.iloc[
        indice_seleccionado
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


st.markdown(
    f"""
    <div class="titulo-seccion">
        🚗 {placa} | {serie}
    </div>
    """,
    unsafe_allow_html=True
)


# =========================================================
# INFORMACIÓN GENERAL
# =========================================================

st.markdown(
    '<div class="titulo-seccion">📋 Información general</div>',
    unsafe_allow_html=True
)

col1, col2, col3, col4 = st.columns(4)


with col1:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">LÍNEA</div>
            <div class="valor">
                {valor_visible(unidad.get("LINEA"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with col2:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">MODELO</div>
            <div class="valor">
                {valor_visible(unidad.get("MODELO"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with col3:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">ESTATUS</div>
            <div class="valor">
                {valor_visible(unidad.get("ESTATUS"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with col4:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">SUBESTATUS</div>
            <div class="valor">
                {valor_visible(unidad.get("SUBESTATUS"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# DATOS DE LA UNIDAD
# =========================================================

st.markdown(
    '<div class="titulo-seccion">🚘 Datos de la unidad</div>',
    unsafe_allow_html=True
)


col1, col2, col3 = st.columns(3)


with col1:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">NOMBRE DEL CLIENTE</div>
            <div class="valor">
                {valor_visible(unidad.get("NOMBRE DEL CLIENTE"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with col2:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">ALMACÉN</div>
            <div class="valor">
                {valor_visible(unidad.get("ALMACÉN"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with col3:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">UNIDAD DE NEGOCIO</div>
            <div class="valor">
                {valor_visible(unidad.get("UNIDAD NEGOCIO"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


col1, col2, col3 = st.columns(3)


with col1:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">GPS</div>
            <div class="valor">
                {valor_visible(unidad.get("GPS"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with col2:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">PLATAFORMA</div>
            <div class="valor">
                {valor_visible(unidad.get("PLATAFORMA"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with col3:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">¿REPORTA?</div>
            <div class="valor">
                {valor_visible(unidad.get("¿REPORTA?"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# PISO
# =========================================================

st.markdown(
    '<div class="separador"></div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="titulo-seccion">🏭 Piso</div>',
    unsafe_allow_html=True
)


dias_piso = dias_visible(
    unidad.get("DÍAS EN PISO")
)


st.markdown(
    f"""
    <div class="tarjeta-piso">
        <div class="dias-label">
            DÍAS EN PISO
        </div>
        <div class="dias-grande">
            {dias_piso}
        </div>
    </div>
    """,
    unsafe_allow_html=True
)


# =========================================================
# GESTORÍA
# =========================================================

st.markdown(
    '<div class="separador"></div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="titulo-seccion">📋 Gestoría</div>',
    unsafe_allow_html=True
)


folio_gestoria = unidad.get(
    "Gestorías1.Folio"
)

gestoria = unidad.get(
    "Gestoria"
)


tiene_gestoria = (
    tiene_informacion(folio_gestoria)
    or
    tiene_informacion(gestoria)
)


if tiene_gestoria:

    st.markdown(
        """
        <div class="estado-si">
            ✓ SÍ TIENE GESTORÍA
        </div>
        """,
        unsafe_allow_html=True
    )

else:

    st.markdown(
        """
        <div class="estado-no">
            — SIN GESTORÍA
        </div>
        """,
        unsafe_allow_html=True
    )


col1, col2, col3 = st.columns(3)


with col1:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">FOLIO</div>
            <div class="valor">
                {valor_visible(unidad.get("Gestorías1.Folio"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with col2:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">ESTATUS DE PÓLIZA</div>
            <div class="valor">
                {valor_visible(unidad.get("Estatus poliza"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with col3:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">FECHA DE CREACIÓN</div>
            <div class="valor">
                {fecha_visible(unidad.get("Gestorías1.Fecha de creación"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


col1, col2, col3 = st.columns(3)


with col1:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">ESTADO DE ATENCIÓN</div>
            <div class="valor">
                {valor_visible(unidad.get("Gestorías1.Estado de atención"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with col2:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">DÍAS</div>
            <div class="valor">
                {valor_visible(unidad.get("Gestorías1.Personalizado"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with col3:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">CONCEPTOS</div>
            <div class="valor">
                {valor_visible(unidad.get("Gestorías1.Conceptos"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# MANTENIMIENTO
# =========================================================

st.markdown(
    '<div class="separador"></div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="titulo-seccion">🔧 Mantenimiento</div>',
    unsafe_allow_html=True
)


folio_mantenimiento = unidad.get(
    "Mantenimiento.FOLIO"
)


if tiene_informacion(folio_mantenimiento):

    st.markdown(
        """
        <div class="estado-si">
            ✓ TIENE MANTENIMIENTO
        </div>
        """,
        unsafe_allow_html=True
    )

else:

    st.markdown(
        """
        <div class="estado-no">
            — SIN MANTENIMIENTO
        </div>
        """,
        unsafe_allow_html=True
    )


col1, col2, col3 = st.columns(3)


with col1:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">FOLIO</div>
            <div class="valor">
                {valor_visible(unidad.get("Mantenimiento.FOLIO"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with col2:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">FECHA DE CREACIÓN</div>
            <div class="valor">
                {fecha_visible(unidad.get("Mantenimiento.FECHA DE CREACION DEL FOLIO"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with col3:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">ESTATUS</div>
            <div class="valor">
                {valor_visible(unidad.get("Mantenimiento.ESTATUS DEL FOLIO"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


col1, col2 = st.columns(2)


with col1:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">PROVEEDOR</div>
            <div class="valor">
                {valor_visible(unidad.get("Mantenimiento.PROVEEDOR"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with col2:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">DÍAS</div>
            <div class="valor">
                {valor_visible(unidad.get("Mantenimiento.dias"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# ZIGA TALLERES
# =========================================================

st.markdown(
    '<div class="separador"></div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="titulo-seccion">🛠️ Ziga Talleres</div>',
    unsafe_allow_html=True
)


folio_ziga = unidad.get(
    "Ziga_Talleres.Folio"
)


if tiene_informacion(folio_ziga):

    st.markdown(
        """
        <div class="estado-si">
            ✓ TIENE REGISTRO EN ZIGA TALLERES
        </div>
        """,
        unsafe_allow_html=True
    )

else:

    st.markdown(
        """
        <div class="estado-no">
            — SIN REGISTRO EN ZIGA TALLERES
        </div>
        """,
        unsafe_allow_html=True
    )


col1, col2, col3 = st.columns(3)


with col1:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">FOLIO</div>
            <div class="valor">
                {valor_visible(unidad.get("Ziga_Talleres.Folio"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with col2:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">ESTATUS</div>
            <div class="valor">
                {valor_visible(unidad.get("Ziga_Talleres.Estatus"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with col3:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">ÁREA DEL TALLER ACTIVA</div>
            <div class="valor">
                {valor_visible(unidad.get("Ziga_Talleres.Área del taller activa"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


col1, col2 = st.columns(2)


with col1:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">FECHA DE INGRESO</div>
            <div class="valor">
                {fecha_visible(unidad.get("Ziga_Talleres.Fecha de ingreso"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with col2:

    st.markdown(
        f"""
        <div class="tarjeta">
            <div class="campo">DÍAS</div>
            <div class="valor">
                {valor_visible(unidad.get("Ziga_Talleres.dias"))}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# PIE
# =========================================================

st.markdown(
    '<div class="separador"></div>',
    unsafe_allow_html=True
)

st.caption(
    "Casanova | Control de Unidades"
)
