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

ARCHIVO_BASE = Path("concentrado.xlsx")

HOJA_BASE = "concentrado"
HOJA_REFACCIONES = "refacciones"


# =========================================================
# USUARIO Y CONTRASEÑA
# =========================================================

USUARIO_CORRECTO = "Casanova"
CONTRASENA_CORRECTA = "Casanova.50"


# =========================================================
# COLORES CASANOVA
# =========================================================

AZUL_MARINO = "#213769"
AZUL_CASANOVA = "#004876"
AMARILLO = "#FFB700"
BLANCO = "#FFFFFF"


# =========================================================
# SESIÓN
# =========================================================

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False


# =========================================================
# LOGIN
# =========================================================

def pantalla_login():

    st.markdown(
        f"""
        <style>

        .stApp {{
            background-color: #0D1117;
        }}

        .login-contenedor {{
            max-width: 460px;
            margin: 100px auto 0 auto;
            background: white;
            border-radius: 20px;
            padding: 40px;
            box-shadow: 0 10px 35px rgba(0,0,0,0.30);
        }}

        .login-icon {{
            text-align: center;
            font-size: 52px;
            margin-bottom: 5px;
        }}

        .login-title {{
            text-align: center;
            color: {AZUL_MARINO};
            font-size: 30px;
            font-weight: 800;
        }}

        .login-subtitle {{
            text-align: center;
            color: #64748B;
            font-size: 15px;
            margin-bottom: 25px;
        }}

        </style>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        f"""
        <div class="login-contenedor">
            <div class="login-icon">🚗</div>
            <div class="login-title">CASANOVA</div>
            <div class="login-subtitle">
                Control de Unidades
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    usuario = st.text_input(
        "Usuario",
        placeholder="Usuario"
    )

    contrasena = st.text_input(
        "Contraseña",
        type="password",
        placeholder="Contraseña"
    )

    if st.button(
        "🔐 Iniciar sesión",
        use_container_width=True
    ):

        if (
            usuario == USUARIO_CORRECTO
            and
            contrasena == CONTRASENA_CORRECTA
        ):

            st.session_state.autenticado = True
            st.rerun()

        else:

            st.error(
                "Usuario o contraseña incorrectos."
            )


# =========================================================
# MOSTRAR LOGIN
# =========================================================

if not st.session_state.autenticado:

    pantalla_login()
    st.stop()


# =========================================================
# FUNCIONES
# =========================================================

def tiene_informacion(valor):

    if valor is None:
        return False

    try:

        if pd.isna(valor):
            return False

    except:
        pass

    texto = str(valor).strip()

    if texto == "":
        return False

    if texto.lower() in [
        "nan",
        "none",
        "null",
        "sin información"
    ]:
        return False

    return True


def valor_visible(valor):

    if not tiene_informacion(valor):
        return ""

    return str(valor).strip()


def fecha_visible(valor):

    if not tiene_informacion(valor):
        return ""

    try:

        return pd.to_datetime(
            valor
        ).strftime("%d/%m/%Y")

    except:

        return valor_visible(valor)


def entero_visible(valor):

    if not tiene_informacion(valor):
        return ""

    try:

        numero = float(valor)

        return f"{int(numero):,}"

    except:

        return valor_visible(valor)


def dias_piso_visible(valor):

    if not tiene_informacion(valor):

        return "Recién ingresado"

    try:

        numero = float(valor)

        return f"{int(numero):,} días"

    except:

        return valor_visible(valor)


# =========================================================
# TARJETAS
# =========================================================

def tarjeta(campo, valor):

    if not tiene_informacion(valor):
        return ""

    return f"""
    <div class="tarjeta">
        <div class="campo">{campo}</div>
        <div class="valor">{valor}</div>
    </div>
    """


def mostrar_tarjeta(campo, valor):

    if tiene_informacion(valor):

        st.markdown(
            tarjeta(campo, valor),
            unsafe_allow_html=True
        )


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
   TÍTULOS
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


/* =====================================================
   TOTAL UNIDADES
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
    margin-top: 10px;
    margin-bottom: 25px;
    box-shadow: 0 4px 14px rgba(33,55,105,0.18);
}}

.total-titulo {{
    font-size: 14px;
    font-weight: 600;
    opacity: 0.9;
}}

.total-numero {{
    font-size: 42px;
    font-weight: 800;
    margin-top: 4px;
}}

.total-desglose {{
    font-size: 15px;
    margin-top: 8px;
}}


/* =====================================================
   PISO
   ===================================================== */

.tarjeta-piso {{
    background: {BLANCO};
    border: 2px solid {AMARILLO};
    border-radius: 14px;
    padding: 20px;
    text-align: center;
    box-shadow: 0 3px 10px rgba(255,183,0,0.12);
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


/* =====================================================
   BOTÓN LOGOUT
   ===================================================== */

.boton-salir {{
    text-align: right;
}}

</style>
""",
unsafe_allow_html=True
)


# =========================================================
# ENCABEZADO
# =========================================================

col_titulo, col_salida = st.columns(
    [5, 1]
)


with col_titulo:

    st.markdown(
        '<div class="titulo-principal">'
        '🚗 CASANOVA | CONTROL DE UNIDADES'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="subtitulo">'
        'Consulta individual de unidades'
        '</div>',
        unsafe_allow_html=True
    )


with col_salida:

    if st.button(
        "Cerrar sesión",
        use_container_width=True
    ):

        st.session_state.autenticado = False
        st.rerun()


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
        .str.strip()
    )

    return df


# =========================================================
# CARGAR REFACCIONES
# =========================================================

@st.cache_data
def cargar_refacciones():

    try:

        df_ref = pd.read_excel(
            ARCHIVO_BASE,
            sheet_name=HOJA_REFACCIONES
        )

        df_ref.columns = (
            df_ref.columns
            .astype(str)
            .str.strip()
        )

        return df_ref

    except:

        return pd.DataFrame()


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


df_refacciones = cargar_refacciones()


# =========================================================
# NORMALIZAR SERIE
# =========================================================

if "SERIE" in df.columns:

    df["SERIE_BUSQUEDA"] = (
        df["SERIE"]
        .astype("string")
        .fillna("")
        .str.strip()
        .str.upper()
    )


if (
    not df_refacciones.empty
    and
    "Serie" in df_refacciones.columns
):

    df_refacciones["SERIE_BUSQUEDA"] = (
        df_refacciones["Serie"]
        .astype("string")
        .fillna("")
        .str.strip()
        .str.upper()
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
# TOTAL UNIDADES
# =========================================================

if "Piso" in df.columns:

    piso_normalizado = (
        df["Piso"]
        .astype("string")
        .fillna("")
        .str.strip()
        .str.upper()
    )

    total_unidades = len(df)

    total_centenario = (
        piso_normalizado == "CENTENARIO"
    ).sum()

    total_lala = (
        piso_normalizado == "LALA"
    ).sum()

else:

    total_unidades = len(df)
    total_centenario = 0
    total_lala = 0


st.markdown(
f"""
<div class="total-piso">

<div class="total-titulo">
🚗 TOTAL DE UNIDADES
</div>

<div class="total-numero">
{total_unidades:,}
</div>

<div class="total-desglose">
Centenario: <b>{total_centenario:,}</b>
&nbsp;&nbsp;&nbsp; / &nbsp;&nbsp;&nbsp;
Lala: <b>{total_lala:,}</b>
</div>

</div>
""",
unsafe_allow_html=True
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

    mostrar_tarjeta(
        "LÍNEA",
        valor_visible(
            unidad.get("LINEA")
        )
    )


with col2:

    mostrar_tarjeta(
        "MODELO",
        entero_visible(
            unidad.get("MODELO")
        )
    )


with col3:

    mostrar_tarjeta(
        "ESTATUS",
        valor_visible(
            unidad.get("ESTATUS")
        )
    )


with col4:

    mostrar_tarjeta(
        "SUBESTATUS",
        valor_visible(
            unidad.get("SUBESTATUS")
        )
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

    mostrar_tarjeta(
        "NOMBRE DEL CLIENTE",
        valor_visible(
            unidad.get("NOMBRE DEL CLIENTE")
        )
    )


with col2:

    mostrar_tarjeta(
        "ALMACÉN",
        valor_visible(
            unidad.get("ALMACÉN")
        )
    )


with col3:

    mostrar_tarjeta(
        "UNIDAD DE NEGOCIO",
        valor_visible(
            unidad.get("UNIDAD NEGOCIO")
        )
    )


col1, col2, col3 = st.columns(3)


with col1:

    mostrar_tarjeta(
        "GPS",
        valor_visible(
            unidad.get("GPS")
        )
    )


with col2:

    mostrar_tarjeta(
        "PLATAFORMA",
        valor_visible(
            unidad.get("PLATAFORMA")
        )
    )


with col3:

    mostrar_tarjeta(
        "¿REPORTA?",
        valor_visible(
            unidad.get("¿REPORTA?")
        )
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


dias_piso = dias_piso_visible(
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
        '<div class="separador"></div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="titulo-seccion">📋 Gestoría</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="estado-si">✓ SÍ TIENE GESTORÍA</div>',
        unsafe_allow_html=True
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        mostrar_tarjeta(
            "FOLIO",
            valor_visible(
                unidad.get(
                    "Gestorías1.Folio"
                )
            )
        )

    with col2:

        mostrar_tarjeta(
            "ESTATUS DE PÓLIZA",
            valor_visible(
                unidad.get(
                    "Estatus poliza"
                )
            )
        )

    with col3:

        mostrar_tarjeta(
            "FECHA DE CREACIÓN",
            fecha_visible(
                unidad.get(
                    "Gestorías1.Fecha de creación"
                )
            )
        )


    col1, col2, col3 = st.columns(3)

    with col1:

        mostrar_tarjeta(
            "ESTADO DE ATENCIÓN",
            valor_visible(
                unidad.get(
                    "Gestorías1.Estado de atención"
                )
            )
        )

    with col2:

        mostrar_tarjeta(
            "DÍAS",
            entero_visible(
                unidad.get(
                    "Gestorías1.Personalizado"
                )
            )
        )

    with col3:

        mostrar_tarjeta(
            "CONCEPTOS",
            valor_visible(
                unidad.get(
                    "Gestorías1.Conceptos"
                )
            )
        )


# =========================================================
# MANTENIMIENTO
# =========================================================

folio_mantenimiento = unidad.get(
    "Mantenimiento.FOLIO"
)


if tiene_informacion(folio_mantenimiento):

    st.markdown(
        '<div class="separador"></div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="titulo-seccion">🔧 Mantenimiento</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="estado-si">✓ TIENE MANTENIMIENTO</div>',
        unsafe_allow_html=True
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        mostrar_tarjeta(
            "FOLIO",
            valor_visible(
                unidad.get(
                    "Mantenimiento.FOLIO"
                )
            )
        )

    with col2:

        mostrar_tarjeta(
            "FECHA DE CREACIÓN",
            fecha_visible(
                unidad.get(
                    "Mantenimiento.FECHA DE CREACION DEL FOLIO"
                )
            )
        )

    with col3:

        mostrar_tarjeta(
            "ESTATUS",
            valor_visible(
                unidad.get(
                    "Mantenimiento.ESTATUS DEL FOLIO"
                )
            )
        )


    col1, col2 = st.columns(2)

    with col1:

        mostrar_tarjeta(
            "PROVEEDOR",
            valor_visible(
                unidad.get(
                    "Mantenimiento.PROVEEDOR"
                )
            )
        )

    with col2:

        mostrar_tarjeta(
            "DÍAS",
            entero_visible(
                unidad.get(
                    "Mantenimiento.dias"
                )
            )
        )


# =========================================================
# ZIGA TALLERES
# =========================================================

folio_ziga = unidad.get(
    "Ziga_Talleres.Folio"
)


if tiene_informacion(folio_ziga):

    st.markdown(
        '<div class="separador"></div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="titulo-seccion">🛠️ Ziga Talleres</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="estado-si">'
        '✓ TIENE REGISTRO EN ZIGA TALLERES'
        '</div>',
        unsafe_allow_html=True
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        mostrar_tarjeta(
            "FOLIO",
            valor_visible(
                unidad.get(
                    "Ziga_Talleres.Folio"
                )
            )
        )

    with col2:

        mostrar_tarjeta(
            "ESTATUS",
            valor_visible(
                unidad.get(
                    "Ziga_Talleres.Estatus"
                )
            )
        )

    with col3:

        mostrar_tarjeta(
            "ÁREA DEL TALLER ACTIVA",
            valor_visible(
                unidad.get(
                    "Ziga_Talleres.Área del taller activa"
                )
            )
        )


    col1, col2 = st.columns(2)

    with col1:

        mostrar_tarjeta(
            "FECHA DE INGRESO",
            fecha_visible(
                unidad.get(
                    "Ziga_Talleres.Fecha de ingreso"
                )
            )
        )

    with col2:

        mostrar_tarjeta(
            "DÍAS",
            entero_visible(
                unidad.get(
                    "Ziga_Talleres.dias"
                )
            )
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

etapa_siniestro = unidad.get(
    "Siniestros2.ETAPA"
)

folio_siniestro = unidad.get(
    "Talleres y Mantenimientos"
)

tiene_siniestro = (
    tiene_informacion(siniestro)
    or
    tiene_informacion(categoria_siniestro)
    or
    tiene_informacion(etapa_siniestro)
    or
    tiene_informacion(folio_siniestro)
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

    st.markdown(
        '<div class="estado-si">'
        '✓ TIENE SINIESTRO'
        '</div>',
        unsafe_allow_html=True
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        mostrar_tarjeta(
            "SINIESTRO",
            valor_visible(
                unidad.get(
                    "Siniestro"
                )
            )
        )

    with col2:

        mostrar_tarjeta(
            "CATEGORÍA",
            valor_visible(
                unidad.get(
                    "Siniestros2.CATEGORIA"
                )
            )
        )

    with col3:

        mostrar_tarjeta(
            "ETAPA",
            valor_visible(
                unidad.get(
                    "Siniestros2.ETAPA"
                )
            )
        )


    col1, col2, col3 = st.columns(3)

    with col1:

        mostrar_tarjeta(
            "DÍAS",
            entero_visible(
                unidad.get(
                    "Siniestros2.dias"
                )
            )
        )

    with col2:

        mostrar_tarjeta(
            "FECHA DE APERTURA",
            fecha_visible(
                unidad.get(
                    "Siniestros2.FECHA APERTURA"
                )
            )
        )

    with col3:

        mostrar_tarjeta(
            "FECHA DEL SINIESTRO",
            fecha_visible(
                unidad.get(
                    "Siniestros2.FECHA SINIESTRO"
                )
            )
        )


# =========================================================
# REFACCIONES
# =========================================================

if (
    not df_refacciones.empty
    and
    tiene_informacion(serie)
    and
    "SERIE_BUSQUEDA" in df_refacciones.columns
):

    serie_busqueda = serie.strip().upper()

    ref_unidad = df_refacciones[
        df_refacciones["SERIE_BUSQUEDA"]
        == serie_busqueda
    ].copy()


    if not ref_unidad.empty:

        # ---------------------------------------------
        # LIMPIAR FOLIO WS
        # ---------------------------------------------

        if "Folio WS" in ref_unidad.columns:

            ref_unidad["FOLIO_WS_LIMPIO"] = (
                ref_unidad["Folio WS"]
                .astype("string")
                .fillna("")
                .str.strip()
            )

        else:

            ref_unidad["FOLIO_WS_LIMPIO"] = ""


        # ---------------------------------------------
        # AGRUPAR POR FOLIO
        # ---------------------------------------------

        registros_refacciones = []

        for folio, grupo in ref_unidad.groupby(
            "FOLIO_WS_LIMPIO",
            sort=False
        ):

            if not tiene_informacion(folio):
                continue


            actividades = []

            if "Actividad" in grupo.columns:

                for actividad in grupo["Actividad"]:

                    if tiene_informacion(actividad):

                        texto = str(
                            actividad
                        ).strip()

                        if texto not in actividades:

                            actividades.append(
                                texto
                            )


            refacciones = []

            if "Refacción" in grupo.columns:

                for refaccion in grupo["Refacción"]:

                    if tiene_informacion(refaccion):

                        texto = str(
                            refaccion
                        ).strip()

                        if texto not in refacciones:

                            refacciones.append(
                                texto
                            )


            registros_refacciones.append(
                {
                    "Folio WS": folio,
                    "Actividad": " / ".join(
                        actividades
                    ),
                    "Refacción": " / ".join(
                        refacciones
                    )
                }
            )


        # ---------------------------------------------
        # MOSTRAR REFACCIONES
        # ---------------------------------------------

        if registros_refacciones:

            st.markdown(
                '<div class="separador"></div>',
                unsafe_allow_html=True
            )

            st.markdown(
                '<div class="titulo-seccion">'
                '🔩 Refacciones pendientes'
                '</div>',
                unsafe_allow_html=True
            )

            st.markdown(
                '<div class="estado-si">'
                '✓ TIENE REFACCIONES'
                '</div>',
                unsafe_allow_html=True
            )


            for registro in registros_refacciones:

                col1, col2, col3 = st.columns(
                    [1, 2, 2]
                )

                with col1:

                    mostrar_tarjeta(
                        "FOLIO WS",
                        registro["Folio WS"]
                    )

                with col2:

                    mostrar_tarjeta(
                        "ACTIVIDAD",
                        registro["Actividad"]
                    )

                with col3:

                    mostrar_tarjeta(
                        "REFACCIÓN",
                        registro["Refacción"]
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
