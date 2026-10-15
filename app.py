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

# El archivo concentrado.xlsx debe estar en el mismo
# repositorio de GitHub que este app.py

ARCHIVO_BASE = Path("concentrado.xlsx")

HOJA_BASE = "concentrado"

# Hoja nueva de refacciones
HOJA_REFACCIONES = "Refacciones"


# =========================================================
# COLORES CASANOVA
# =========================================================

AZUL_MARINO = "#213769"
AZUL_CASANOVA = "#004876"
AMARILLO = "#FFB700"
BLANCO = "#FFFFFF"


# =========================================================
# CARGAR CONCENTRADO
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
        .str.replace("\n", " ", regex=False)
        .str.replace("\r", " ", regex=False)
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
    )

    return df


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

        df.columns = (
            df.columns
            .astype(str)
            .str.replace("\n", " ", regex=False)
            .str.replace("\r", " ", regex=False)
            .str.replace("\xa0", " ", regex=False)
            .str.strip()
        )

        return df

    except Exception:

        return pd.DataFrame()


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


def numero_entero(valor):

    if not tiene_informacion(valor):
        return "Sin información"

    try:

        numero = float(valor)

        return f"{int(numero):,}"

    except Exception:

        return valor_visible(valor)


def dias_piso_visible(valor):

    if not tiene_informacion(valor):

        return "🆕 Recién ingresado"

    try:

        numero = float(valor)

        return f"{int(numero):,} días"

    except Exception:

        return valor_visible(valor)


# =========================================================
# FUNCIONES PARA TARJETAS
# =========================================================

def tarjeta(campo, valor):

    return f"""
<div class="tarjeta">
<div class="campo">{campo}</div>
<div class="valor">{valor}</div>
</div>
"""


def tarjeta_azul(campo, valor):

    return f"""
<div class="tarjeta-azul">
<div class="campo-blanco">{campo}</div>
<div class="valor-blanco">{valor}</div>
</div>
"""


# =========================================================
# ESTILO
# =========================================================

st.markdown(
f"""
<style>

.stApp {{
    background-color: #F7F9FC;
}}


/* =====================================================
   TÍTULO
   ===================================================== */

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


/* =====================================================
   TÍTULOS DE SECCIÓN
   ===================================================== */

.titulo-seccion {{
    color: {AZUL_MARINO};
    font-size: 21px;
    font-weight: 750;
    margin-top: 22px;
    margin-bottom: 12px;
}}


/* =====================================================
   TARJETAS
   ===================================================== */

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


/* =====================================================
   CAMPOS
   ===================================================== */

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

.campo-blanco {{
    color: white;
    font-size: 12px;
    font-weight: 600;
    margin-bottom: 5px;
}}

.valor-blanco {{
    color: white;
    font-size: 17px;
    font-weight: 700;
}}


/* =====================================================
   TOTAL
   ===================================================== */

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

.total-piso-desglose {{
    font-size: 14px;
    margin-top: 5px;
}}


/* =====================================================
   PISO
   ===================================================== */

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


/* =====================================================
   REFACCIONES
   ===================================================== */

.refaccion-card {{
    background: {BLANCO};
    border: 1px solid #D8E1EF;
    border-left: 5px solid {AMARILLO};
    border-radius: 14px;
    padding: 16px 18px;
    margin-bottom: 12px;
    box-shadow: 0 2px 8px rgba(33,55,105,0.06);
}}

.refaccion-folio {{
    color: {AZUL_MARINO};
    font-size: 14px;
    font-weight: 800;
    margin-bottom: 7px;
}}

.refaccion-actividad {{
    color: #172033;
    font-size: 16px;
    font-weight: 700;
    line-height: 1.5;
}}


/* =====================================================
   ESTADOS
   ===================================================== */

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


/* =====================================================
   SEPARADOR
   ===================================================== */

.separador {{
    height: 2px;
    background: #DCE4F0;
    margin: 28px 0;
    border-radius: 5px;
}}


/* =====================================================
   BUSCADOR
   ===================================================== */

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

    st.code(str(e))

    st.stop()


# =========================================================
# CARGAR REFACCIONES
# =========================================================

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

    total_unidades = len(df)


# =========================================================
# CENTENARIO / LALA
# =========================================================

centenario = 0
lala = 0


if "Piso" in df.columns and "SERIE" in df.columns:

    df_pisos = df.copy()

    df_pisos["_SERIE"] = (
        df_pisos["SERIE"]
        .astype("string")
        .fillna("")
        .str.strip()
        .str.upper()
    )

    df_pisos["_PISO"] = (
        df_pisos["Piso"]
        .astype("string")
        .fillna("")
        .str.strip()
        .str.upper()
    )

    df_pisos = df_pisos[
        df_pisos["_SERIE"] != ""
    ]

    # Una unidad solamente una vez
    df_pisos = df_pisos.drop_duplicates(
        subset="_SERIE"
    )

    centenario = int(
        (
            df_pisos["_PISO"]
            == "CENTENARIO"
        ).sum()
    )

    lala = int(
        (
            df_pisos["_PISO"]
            == "LALA"
        ).sum()
    )


# =========================================================
# TARJETA TOTAL
# =========================================================

# IMPORTANTE:
# No se dejan espacios al principio de las líneas HTML
# para evitar que Streamlit lo interprete como código.

html_total = (
    '<div class="total-piso">'
    '<div class="total-piso-titulo">🚗 TOTAL DE UNIDADES</div>'
    f'<div class="total-piso-numero">{total_unidades:,}</div>'
    '<div class="total-piso-desglose">'
    f'Centenario: <b>{centenario:,}</b>'
    '&nbsp;&nbsp; / &nbsp;&nbsp;'
    f'Lala: <b>{lala:,}</b>'
    '</div>'
    '</div>'
)

st.markdown(
    html_total,
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
# DÍAS EN PISO
# =========================================================

# Esta tarjeta aparecerá DEBAJO DEL BUSCADOR
# cuando exista una unidad seleccionada.
#
# No se muestra antes de buscar.


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

criterio = (
    busqueda
    .strip()
    .upper()
)


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


# =========================================================
# DÍAS EN PISO
# =========================================================

dias_piso = dias_piso_visible(
    unidad.get("DÍAS EN PISO")
)


# =========================================================
# TARJETA DÍAS EN PISO
# =========================================================

html_dias = (
    '<div class="tarjeta-piso">'
    '<div class="dias-label">🏭 DÍAS EN PISO</div>'
    f'<div class="dias-grande">{dias_piso}</div>'
    '</div>'
)

st.markdown(
    html_dias,
    unsafe_allow_html=True
)


# =========================================================
# TÍTULO DE UNIDAD
# =========================================================

st.markdown(
    f'<div class="titulo-seccion">🚗 {placa} | {serie}</div>',
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
            numero_entero(
                unidad.get("MODELO")
            )
        ),
        unsafe_allow_html=True
    )


with col3:

    st.markdown(
        tarjeta(
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

st.markdown(
    '<div class="titulo-seccion">🚘 Datos de la unidad</div>',
    unsafe_allow_html=True
)


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
# GESTORÍA
# =========================================================

folio_gestoria = unidad.get(
    "Gestorías1.Folio"
)

gestoria = unidad.get(
    "Gestoria"
)

fecha_gestoria = unidad.get(
    "Gestorías1.Fecha de creación"
)

estado_gestoria = unidad.get(
    "Gestorías1.Estado de atención"
)

conceptos_gestoria = unidad.get(
    "Gestorías1.Conceptos"
)

dias_gestoria = unidad.get(
    "Gestorías1.Personalizado"
)


tiene_gestoria = any(
    tiene_informacion(valor)
    for valor in [
        folio_gestoria,
        gestoria,
        fecha_gestoria,
        estado_gestoria,
        conceptos_gestoria,
        dias_gestoria
    ]
)


if tiene_gestoria:

    st.markdown(
        '<div class="separador"></div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="titulo-seccion">📋 Gestoría</div>',
        unsafe_allow_html=True
    )

    col1, col2, col3 = st.columns(3)


    with col1:

        st.markdown(
            tarjeta(
                "FOLIO",
                valor_visible(
                    folio_gestoria
                )
            ),
            unsafe_allow_html=True
        )


    with col2:

        st.markdown(
            tarjeta(
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
                    fecha_gestoria
                )
            ),
            unsafe_allow_html=True
        )


    col1, col2, col3 = st.columns(3)


    with col1:

        st.markdown(
            tarjeta(
                "ESTADO DE ATENCIÓN",
                valor_visible(
                    estado_gestoria
                )
            ),
            unsafe_allow_html=True
        )


    with col2:

        st.markdown(
            tarjeta(
                "DÍAS",
                numero_entero(
                    dias_gestoria
                )
            ),
            unsafe_allow_html=True
        )


    with col3:

        st.markdown(
            tarjeta(
                "CONCEPTOS",
                valor_visible(
                    conceptos_gestoria
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

fecha_mantenimiento = unidad.get(
    "Mantenimiento.FECHA DE CREACION DEL FOLIO"
)

estatus_mantenimiento = unidad.get(
    "Mantenimiento.ESTATUS DEL FOLIO"
)

proveedor_mantenimiento = unidad.get(
    "Mantenimiento.PROVEEDOR"
)

dias_mantenimiento = unidad.get(
    "Mantenimiento.dias"
)


tiene_mantenimiento = any(
    tiene_informacion(valor)
    for valor in [
        folio_mantenimiento,
        fecha_mantenimiento,
        estatus_mantenimiento,
        proveedor_mantenimiento,
        dias_mantenimiento
    ]
)


if tiene_mantenimiento:

    st.markdown(
        '<div class="separador"></div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="titulo-seccion">🔧 Mantenimiento</div>',
        unsafe_allow_html=True
    )


    col1, col2, col3 = st.columns(3)


    with col1:

        st.markdown(
            tarjeta(
                "FOLIO",
                valor_visible(
                    folio_mantenimiento
                )
            ),
            unsafe_allow_html=True
        )


    with col2:

        st.markdown(
            tarjeta(
                "FECHA DE CREACIÓN",
                fecha_visible(
                    fecha_mantenimiento
                )
            ),
            unsafe_allow_html=True
        )


    with col3:

        st.markdown(
            tarjeta(
                "ESTATUS",
                valor_visible(
                    estatus_mantenimiento
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
                    proveedor_mantenimiento
                )
            ),
            unsafe_allow_html=True
        )


    with col2:

        st.markdown(
            tarjeta(
                "DÍAS",
                numero_entero(
                    dias_mantenimiento
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

estatus_ziga = unidad.get(
    "Ziga_Talleres.Estatus"
)

area_ziga = unidad.get(
    "Ziga_Talleres.Área del taller activa"
)

fecha_ziga = unidad.get(
    "Ziga_Talleres.Fecha de ingreso"
)

dias_ziga = unidad.get(
    "Ziga_Talleres.dias"
)


tiene_ziga = any(
    tiene_informacion(valor)
    for valor in [
        folio_ziga,
        estatus_ziga,
        area_ziga,
        fecha_ziga,
        dias_ziga
    ]
)


if tiene_ziga:

    st.markdown(
        '<div class="separador"></div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="titulo-seccion">🛠️ Ziga Talleres</div>',
        unsafe_allow_html=True
    )


    col1, col2, col3 = st.columns(3)


    with col1:

        st.markdown(
            tarjeta(
                "FOLIO",
                valor_visible(
                    folio_ziga
                )
            ),
            unsafe_allow_html=True
        )


    with col2:

        st.markdown(
            tarjeta(
                "ESTATUS",
                valor_visible(
                    estatus_ziga
                )
            ),
            unsafe_allow_html=True
        )


    with col3:

        st.markdown(
            tarjeta(
                "ÁREA DEL TALLER ACTIVA",
                valor_visible(
                    area_ziga
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
                    fecha_ziga
                )
            ),
            unsafe_allow_html=True
        )


    with col2:

        st.markdown(
            tarjeta(
                "DÍAS",
                numero_entero(
                    dias_ziga
                )
            ),
            unsafe_allow_html=True
        )


# =========================================================
# SINIESTRO
# =========================================================

siniestro = unidad.get(
    "Siniestro"
)

categoria_siniestro = unidad.get(
    "Siniestros2.CATEGORIA"
)

dias_siniestro = unidad.get(
    "Siniestros2.dias"
)

etapa_siniestro = unidad.get(
    "Siniestros2.ETAPA"
)

fecha_apertura_siniestro = unidad.get(
    "Siniestros2.FECHA APERTURA"
)

fecha_siniestro = unidad.get(
    "Siniestros2.FECHA SINIESTRO"
)


tiene_siniestro = any(
    tiene_informacion(valor)
    for valor in [
        siniestro,
        categoria_siniestro,
        dias_siniestro,
        etapa_siniestro,
        fecha_apertura_siniestro,
        fecha_siniestro
    ]
)


if tiene_siniestro:

    st.markdown(
        '<div class="separador"></div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="titulo-seccion">🚨 Siniestro</div>',
        unsafe_allow_html=True
    )


    col1, col2, col3 = st.columns(3)


    with col1:

        st.markdown(
            tarjeta(
                "SINIESTRO",
                valor_visible(
                    siniestro
                )
            ),
            unsafe_allow_html=True
        )


    with col2:

        st.markdown(
            tarjeta(
                "CATEGORÍA",
                valor_visible(
                    categoria_siniestro
                )
            ),
            unsafe_allow_html=True
        )


    with col3:

        st.markdown(
            tarjeta(
                "ETAPA",
                valor_visible(
                    etapa_siniestro
                )
            ),
            unsafe_allow_html=True
        )


    col1, col2, col3 = st.columns(3)


    with col1:

        st.markdown(
            tarjeta(
                "DÍAS",
                numero_entero(
                    dias_siniestro
                )
            ),
            unsafe_allow_html=True
        )


    with col2:

        st.markdown(
            tarjeta(
                "FECHA DE APERTURA",
                fecha_visible(
                    fecha_apertura_siniestro
                )
            ),
            unsafe_allow_html=True
        )


    with col3:

        st.markdown(
            tarjeta(
                "FECHA DEL SINIESTRO",
                fecha_visible(
                    fecha_siniestro
                )
            ),
            unsafe_allow_html=True
        )


# =========================================================
# REFACCIONES
# =========================================================

if not df_refacciones.empty:

    # -----------------------------------------------------
    # BUSCAR COLUMNAS
    # -----------------------------------------------------

    columna_serie_ref = None
    columna_folio_ref = None
    columna_actividad_ref = None


    for columna in df_refacciones.columns:

        nombre = (
            str(columna)
            .strip()
            .lower()
        )

        if nombre == "serie":

            columna_serie_ref = columna

        elif nombre == "folio ws":

            columna_folio_ref = columna

        elif nombre == "actividad":

            columna_actividad_ref = columna


    # -----------------------------------------------------
    # SI EXISTEN LAS TRES COLUMNAS
    # -----------------------------------------------------

    if (
        columna_serie_ref is not None
        and columna_folio_ref is not None
        and columna_actividad_ref is not None
    ):

        ref = df_refacciones.copy()


        # -------------------------------------------------
        # NORMALIZAR SERIE DE REFACCIONES
        # -------------------------------------------------

        ref["_SERIE"] = (
            ref[columna_serie_ref]
            .astype("string")
            .fillna("")
            .str.strip()
            .str.upper()
        )


        serie_busqueda = (
            str(serie)
            .strip()
            .upper()
        )


        # -------------------------------------------------
        # RELACIONAR CON SERIE DEL CONCENTRADO
        # -------------------------------------------------

        ref = ref[
            ref["_SERIE"]
            == serie_busqueda
        ].copy()


        # -------------------------------------------------
        # SI ENCONTRÓ REFACCIONES
        # -------------------------------------------------

        if not ref.empty:

            grupos_refacciones = []


            # -------------------------------------------------
            # AGRUPAR POR FOLIO WS
            # -------------------------------------------------

            for folio, grupo in ref.groupby(
                columna_folio_ref,
                sort=False
            ):

                if not tiene_informacion(
                    folio
                ):

                    continue


                actividades = []


                # -------------------------------------------------
                # ACTIVIDADES ÚNICAS
                # -------------------------------------------------

                for actividad in grupo[
                    columna_actividad_ref
                ]:

                    if not tiene_informacion(
                        actividad
                    ):

                        continue


                    actividad = str(
                        actividad
                    ).strip()


                    # No repetir la misma actividad
                    # aunque aparezca varias veces

                    repetida = any(
                        actividad.casefold()
                        ==
                        existente.casefold()
                        for existente in actividades
                    )


                    if not repetida:

                        actividades.append(
                            actividad
                        )


                if actividades:

                    grupos_refacciones.append(
                        (
                            str(
                                folio
                            ).strip(),

                            " / ".join(
                                actividades
                            )
                        )
                    )


            # -------------------------------------------------
            # MOSTRAR REFACCIONES
            # -------------------------------------------------

            if grupos_refacciones:

                st.markdown(
                    '<div class="separador"></div>',
                    unsafe_allow_html=True
                )

                st.markdown(
                    '<div class="titulo-seccion">📦 Refacciones</div>',
                    unsafe_allow_html=True
                )


                for folio, actividades in grupos_refacciones:

                    html_refaccion = (
                        '<div class="refaccion-card">'
                        f'<div class="refaccion-folio">'
                        f'FOLIO WS: {folio}'
                        '</div>'
                        f'<div class="refaccion-actividad">'
                        f'{actividades}'
                        '</div>'
                        '</div>'
                    )


                    st.markdown(
                        html_refaccion,
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
