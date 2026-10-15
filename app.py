import streamlit as st
import pandas as pd
from pathlib import Path
from textwrap import dedent


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
# LOGIN
# =========================================================

def verificar_login(usuario, password):

    return (
        usuario.strip() == USUARIO_CORRECTO
        and password == PASSWORD_CORRECTA
    )


def mostrar_login():

    st.markdown(
        f"""
        <style>
        .login-box {{
            max-width: 430px;
            margin: 70px auto 20px auto;
            background: white;
            padding: 38px;
            border-radius: 18px;
            border: 1px solid #D8E1EF;
            box-shadow: 0 8px 30px rgba(33,55,105,0.12);
            text-align: center;
        }}

        .login-icon {{
            font-size: 48px;
            margin-bottom: 8px;
        }}

        .login-title {{
            color: {AZUL_MARINO};
            font-size: 27px;
            font-weight: 800;
        }}

        .login-subtitle {{
            color: #64748B;
            font-size: 14px;
            margin-top: 5px;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

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

    col1, col2, col3 = st.columns([1, 2, 1])

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

        numero = float(valor)

        if numero.is_integer():

            return str(int(numero))

        return str(numero)

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


def normalizar_serie(valor):

    if not tiene_informacion(valor):
        return ""

    return (
        str(valor)
        .strip()
        .upper()
    )


# =========================================================
# TARJETA HTML
# =========================================================

def tarjeta(campo, valor):

    return f"""
<div class="tarjeta">
<div class="campo">{campo}</div>
<div class="valor">{valor}</div>
</div>
""".strip()


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
   TÍTULOS
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
    margin-bottom: 18px;
}}

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
    margin-bottom: 24px;
    box-shadow: 0 4px 14px rgba(33,55,105,0.18);
}}

.total-titulo {{
    font-size: 14px;
    font-weight: 600;
    opacity: 0.95;
}}

.total-numero {{
    font-size: 42px;
    font-weight: 800;
    margin-top: 2px;
}}

.total-desglose {{
    font-size: 14px;
    margin-top: 5px;
    line-height: 1.8;
}}


/* =====================================================
   DÍAS EN PISO
   ===================================================== */

.tarjeta-piso {{
    background: white;
    border: 2px solid {AMARILLO};
    border-radius: 14px;
    padding: 18px;
    text-align: center;
    box-shadow: 0 3px 10px rgba(255,183,0,0.12);
    margin-top: 12px;
    margin-bottom: 20px;
}}

.dias-label {{
    color: #64748B;
    font-size: 13px;
    font-weight: 600;
}}

.dias-grande {{
    color: {AZUL_MARINO};
    font-size: 32px;
    font-weight: 800;
    margin-top: 3px;
}}


/* =====================================================
   REFACCIONES
   ===================================================== */

.refaccion-card {{
    background: white;
    border: 1px solid #D8E1EF;
    border-left: 5px solid {AMARILLO};
    border-radius: 13px;
    padding: 16px 18px;
    margin-bottom: 12px;
    box-shadow: 0 2px 7px rgba(33,55,105,0.05);
}}

.refaccion-folio {{
    color: {AZUL_MARINO};
    font-size: 15px;
    font-weight: 800;
    margin-bottom: 8px;
}}

.refaccion-actividad {{
    color: #172033;
    font-size: 15px;
    font-weight: 600;
    line-height: 1.5;
}}


/* =====================================================
   MENSAJE SIN FOLIOS
   ===================================================== */

.sin-folios {{
    background: #F1F5F9;
    border: 1px solid #D8E1EF;
    color: #64748B;
    border-radius: 12px;
    padding: 16px 18px;
    margin-top: 18px;
    font-weight: 600;
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

</style>
""",
    unsafe_allow_html=True
)


# =========================================================
# ENCABEZADO
# =========================================================

col1, col2 = st.columns([5, 1])


with col1:

    st.markdown(
        '<div class="titulo-principal">🚗 CASANOVA | CONTROL DE UNIDADES</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="subtitulo">Consulta individual de unidades</div>',
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
# TOTAL DE SERIES ÚNICAS
# =========================================================

if "SERIE" in df.columns:

    series_totales = (
        df["SERIE"]
        .apply(normalizar_serie)
    )

    series_totales = series_totales[
        series_totales != ""
    ]

    total_unidades = (
        series_totales
        .nunique()
    )

else:

    total_unidades = 0


# =========================================================
# CONTEO POR PISO
# =========================================================
#
# IMPORTANTE:
# Cada serie cuenta solamente una vez.
#
# Ejemplo:
#
# Serie A | CENTENARIO
# Serie A | CENTENARIO
#
# Cuenta como 1.
#
# =========================================================

centenario = 0
lala = 0
otros_sin_piso = 0

if (
    "Piso" in df.columns
    and "SERIE" in df.columns
):

    temp_piso = df.copy()

    temp_piso["_SERIE"] = (
        temp_piso["SERIE"]
        .apply(normalizar_serie)
    )

    temp_piso["_PISO"] = (
        temp_piso["Piso"]
        .astype("string")
        .fillna("")
        .str.strip()
        .str.upper()
    )

    # Solo series válidas
    temp_piso = temp_piso[
        temp_piso["_SERIE"] != ""
    ].copy()

    # Una sola fila por SERIE
    temp_piso = temp_piso.drop_duplicates(
        subset="_SERIE",
        keep="first"
    )

    centenario = int(
        temp_piso["_PISO"]
        .eq("CENTENARIO")
        .sum()
    )

    lala = int(
        temp_piso["_PISO"]
        .eq("LALA")
        .sum()
    )

    otros_sin_piso = (
        total_unidades
        - centenario
        - lala
    )

    if otros_sin_piso < 0:
        otros_sin_piso = 0


# =========================================================
# TARJETA TOTAL
# =========================================================

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
Centenario: <b>{centenario:,}</b>
&nbsp;&nbsp; / &nbsp;&nbsp;
Lala: <b>{lala:,}</b>
&nbsp;&nbsp; / &nbsp;&nbsp;
Otros / Sin piso: <b>{otros_sin_piso:,}</b>
</div>

</div>
""".strip(),
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
    placeholder="Ejemplo: PDN1295 o número de serie...",
    label_visibility="collapsed"
)


# =========================================================
# DÍAS EN PISO
# =========================================================
#
# Si ya se buscó una unidad se mostrará debajo
# del buscador.
#
# =========================================================

if busqueda:

    criterio_temporal = (
        busqueda
        .strip()
        .upper()
    )

    placas_temp = (
        df["PLACA"]
        .astype("string")
        .fillna("")
        .str.strip()
        .str.upper()
    )

    series_temp = (
        df["SERIE"]
        .astype("string")
        .fillna("")
        .str.strip()
        .str.upper()
    )

    resultado_temp = df[
        placas_temp.str.contains(
            criterio_temporal,
            case=False,
            na=False,
            regex=False
        )
        |
        series_temp.str.contains(
            criterio_temporal,
            case=False,
            na=False,
            regex=False
        )
    ].copy()

    if not resultado_temp.empty:

        if len(resultado_temp) == 1:

            unidad_temp = resultado_temp.iloc[0]

            dias_piso_temp = unidad_temp.get(
                "DÍAS EN PISO"
            )

            if tiene_informacion(dias_piso_temp):

                try:

                    dias_numero_temp = int(
                        float(dias_piso_temp)
                    )

                    texto_dias_temp = (
                        f"{dias_numero_temp:,} días"
                    )

                except Exception:

                    texto_dias_temp = valor_visible(
                        dias_piso_temp
                    )

            else:

                texto_dias_temp = "🆕 Recién ingresado"

            st.markdown(
                f"""
<div class="tarjeta-piso">

<div class="dias-label">
🏭 DÍAS EN PISO
</div>

<div class="dias-grande">
{texto_dias_temp}
</div>

</div>
""".strip(),
                unsafe_allow_html=True
            )


# =========================================================
# SI NO HAY BÚSQUEDA
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
# TÍTULO UNIDAD
# =========================================================

st.markdown(
    f"""
<div class="titulo-seccion">
🚗 {placa} | {serie}
</div>
""".strip(),
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

    # MODELO SIN COMAS
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
# CONTROL DE FOLIOS ABIERTOS
# =========================================================

hay_folio_abierto = False


# =========================================================
# GESTORÍA
# =========================================================

folio_gestoria = unidad.get(
    "Gestorías1.Folio"
)

datos_gestoria = [

    unidad.get(
        "Gestorías1.Folio"
    ),

    unidad.get(
        "Gestorías1.Fecha de creación"
    ),

    unidad.get(
        "Gestorías1.Estado de atención"
    ),

    unidad.get(
        "Gestorías1.Conceptos"
    ),

    unidad.get(
        "Gestorías1.Personalizado"
    ),

    unidad.get(
        "Gestoria"
    )
]


tiene_gestoria = any(
    tiene_informacion(x)
    for x in datos_gestoria
)


if tiene_gestoria:

    hay_folio_abierto = True

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
                    unidad.get(
                        "Gestorías1.Folio"
                    )
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
            tarjeta(
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
            tarjeta(
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
                    unidad.get(
                        "Ziga_Talleres.Folio"
                    )
                )
            ),
            unsafe_allow_html=True
        )


    with col2:

        st.markdown(
            tarjeta(
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


datos_siniestro = [

    unidad.get(
        "Siniestro"
    ),

    unidad.get(
        "Siniestros2.CATEGORIA"
    ),

    unidad.get(
        "Siniestros2.dias"
    ),

    unidad.get(
        "Siniestros2.ETAPA"
    ),

    unidad.get(
        "Siniestros2.FECHA APERTURA"
    ),

    unidad.get(
        "Siniestros2.FECHA SINIESTRO"
    )
]


tiene_siniestro = any(
    tiene_informacion(x)
    for x in datos_siniestro
)


if tiene_siniestro:

    hay_folio_abierto = True

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

hay_refacciones = False


if not df_refacciones.empty:

    # -----------------------------------------------------
    # IDENTIFICAR COLUMNAS DE REFACCIONES
    # -----------------------------------------------------

    col_serie_ref = None
    col_folio_ref = None
    col_actividad_ref = None


    for columna in df_refacciones.columns:

        clave = (
            str(columna)
            .strip()
            .lower()
            .replace("_", " ")
        )

        clave = " ".join(
            clave.split()
        )


        if clave == "serie":

            col_serie_ref = columna


        elif clave in [
            "folio ws",
            "folio",
            "folio ws "
        ]:

            col_folio_ref = columna


        elif clave in [
            "actividad",
            "actividades"
        ]:

            col_actividad_ref = columna


    # -----------------------------------------------------
    # SI ENCONTRÓ COLUMNAS
    # -----------------------------------------------------

    if (
        col_serie_ref is not None
        and col_folio_ref is not None
        and col_actividad_ref is not None
    ):

        ref = df_refacciones.copy()


        # -------------------------------------------------
        # NORMALIZAR SERIES
        # -------------------------------------------------

        ref["_SERIE_BUSQUEDA"] = (
            ref[col_serie_ref]
            .apply(normalizar_serie)
        )


        serie_actual = normalizar_serie(
            serie
        )


        # -------------------------------------------------
        # CRUCE SERIE VS SERIE
        # -------------------------------------------------

        ref = ref[
            ref["_SERIE_BUSQUEDA"]
            == serie_actual
        ].copy()


        # -------------------------------------------------
        # AGRUPAR
        # -------------------------------------------------

        grupos = []


        if not ref.empty:

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

                    grupos.append(
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

        if grupos:

            hay_refacciones = True
            hay_folio_abierto = True

            st.markdown(
                '<div class="separador"></div>',
                unsafe_allow_html=True
            )

            st.markdown(
                '<div class="titulo-seccion">📦 Refacciones pendientes</div>',
                unsafe_allow_html=True
            )


            for folio, actividades in grupos:

                st.markdown(
                    f"""
<div class="refaccion-card">

<div class="refaccion-folio">
FOLIO WS: {folio}
</div>

<div class="refaccion-actividad">
{actividades}
</div>

</div>
""".strip(),
                    unsafe_allow_html=True
                )


# =========================================================
# MENSAJE SI NO TIENE FOLIOS
# =========================================================

if not hay_folio_abierto:

    st.markdown(
        """
<div class="separador"></div>

<div class="sin-folios">
✓ SIN FOLIOS ABIERTOS PARA ESTA UNIDAD
</div>
""".strip(),
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
