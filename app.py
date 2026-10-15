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

            if verificar_login(
                usuario,
                password
            ):

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


# =========================================================
# MODELO SIN COMAS
# =========================================================

def modelo_visible(valor):

    if not tiene_informacion(valor):

        return "Sin información"

    try:

        numero = float(valor)

        if numero.is_integer():

            return str(int(numero))

    except Exception:

        pass

    return valor_visible(valor)


# =========================================================
# FECHA
# =========================================================

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


# =========================================================
# TARJETA
# =========================================================

def tarjeta(campo, valor):

    return f"""
<div class="tarjeta">
<div class="campo">{campo}</div>
<div class="valor">{valor}</div>
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
   TOTAL DE UNIDADES
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
    margin-top: 18px;
    margin-bottom: 25px;
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
    margin-top: 14px;
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
    margin-bottom: 10px;
}}

.refaccion-actividad {{
    color: #172033;
    font-size: 15px;
    font-weight: 600;
    line-height: 1.5;
}}


/* =====================================================
   SIN FOLIOS
   ===================================================== */

.sin-folios {{
    background: #F8FAFC;
    border: 1px solid #D8E1EF;
    border-radius: 14px;
    padding: 18px;
    margin-top: 25px;
    color: #64748B;
    text-align: center;
    font-weight: 700;
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
# TOTAL DE UNIDADES
# =========================================================
# SE CUENTAN SERIES ÚNICAS
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
# CENTENARIO / LALA / OTROS
# =========================================================
# SE CUENTA UNA SOLA VEZ CADA SERIE
# Y SE LE ASIGNA SU VALOR DE PISO
# =========================================================

centenario = 0
lala = 0
otros_sin_piso = total_unidades


if (
    "Piso" in df.columns
    and "SERIE" in df.columns
):

    temp_piso = df.copy()

    temp_piso["_SERIE"] = (
        temp_piso["SERIE"]
        .astype("string")
        .fillna("")
        .str.strip()
        .str.upper()
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


    # =====================================================
    # UNA SOLA FILA POR SERIE
    # =====================================================

    temp_piso = temp_piso.drop_duplicates(
        subset="_SERIE",
        keep="first"
    )


    # =====================================================
    # CONTAR PISO
    # =====================================================

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


    # Todo lo demás
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
    placeholder="Ejemplo: PDN1295 o número de serie...",
    label_visibility="collapsed"
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
# DÍAS EN PISO
# =========================================================

dias_piso = unidad.get(
    "DÍAS EN PISO"
)


if tiene_informacion(
    dias_piso
):

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
# TARJETA DÍAS EN PISO
# =========================================================

st.markdown(
    f"""
<div class="tarjeta-piso">

<div class="dias-label">
🏭 DÍAS EN PISO
</div>

<div class="dias-grande">
{texto_dias}
</div>

</div>
""",
    unsafe_allow_html=True
)


# =========================================================
# TÍTULO UNIDAD
# =========================================================

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

folios_abiertos = []


# =========================================================
# GESTORÍA
# =========================================================

folio_gestoria = unidad.get(
    "Gestorías1.Folio"
)


if tiene_informacion(folio_gestoria):

    folios_abiertos.append(
        "Gestoría"
    )

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

    folios_abiertos.append(
        "Mantenimiento"
    )

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

    folios_abiertos.append(
        "Ziga Talleres"
    )

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


if not tiene_informacion(folio_siniestro):

    # En caso de que el folio venga directamente
    # en alguna de las columnas de siniestro

    posibles_folios_siniestro = [

        unidad.get(
            "Siniestros2.CATEGORIA"
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

    tiene_datos_siniestro = any(
        tiene_informacion(x)
        for x in posibles_folios_siniestro
    )

else:

    tiene_datos_siniestro = True


if tiene_datos_siniestro:

    folios_abiertos.append(
        "Siniestro"
    )

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
# CRUCE:
# SERIE DE LA UNIDAD
# VS
# SERIE DE LA HOJA REFACCIONES
#
# MISMO FOLIO WS = UNA SOLA TARJETA
# ACTIVIDADES REPETIDAS = UNA SOLA VEZ
# =========================================================

if not df_refacciones.empty:

    col_serie_ref = None
    col_folio_ref = None
    col_actividad_ref = None


    # =====================================================
    # IDENTIFICAR COLUMNAS
    # =====================================================

    for columna in df_refacciones.columns:

        clave = (
            str(columna)
            .strip()
            .lower()
        )

        if clave == "serie":

            col_serie_ref = columna

        elif clave == "folio ws":

            col_folio_ref = columna

        elif clave == "actividad":

            col_actividad_ref = columna


    # =====================================================
    # VALIDAR COLUMNAS
    # =====================================================

    if (
        col_serie_ref is not None
        and col_folio_ref is not None
        and col_actividad_ref is not None
    ):

        ref = df_refacciones.copy()


        # =================================================
        # NORMALIZAR SERIES
        # =================================================

        ref["_SERIE_BUSQUEDA"] = (
            ref[col_serie_ref]
            .astype("string")
            .fillna("")
            .str.strip()
            .str.upper()
        )


        serie_actual = (
            str(serie)
            .strip()
            .upper()
        )


        # =================================================
        # BUSCAR LA SERIE DE LA UNIDAD
        # =================================================

        ref = ref[
            ref["_SERIE_BUSQUEDA"]
            == serie_actual
        ].copy()


        # =================================================
        # AGRUPAR POR FOLIO
        # =================================================

        grupos_refacciones = []


        if not ref.empty:

            for folio, grupo in ref.groupby(
                col_folio_ref,
                sort=False
            ):

                # -----------------------------------------
                # IGNORAR FOLIO VACÍO
                # -----------------------------------------

                if not tiene_informacion(folio):

                    continue


                actividades_unicas = []


                # -----------------------------------------
                # RECORRER ACTIVIDADES
                # -----------------------------------------

                for actividad in grupo[
                    col_actividad_ref
                ]:

                    if not tiene_informacion(
                        actividad
                    ):

                        continue


                    actividad_texto = (
                        str(actividad)
                        .strip()
                    )


                    # -------------------------------------
                    # EVITAR ACTIVIDADES REPETIDAS
                    # -------------------------------------

                    existe = any(
                        actividad_texto.casefold()
                        ==
                        existente.casefold()
                        for existente
                        in actividades_unicas
                    )


                    if not existe:

                        actividades_unicas.append(
                            actividad_texto
                        )


                # -----------------------------------------
                # GUARDAR FOLIO
                # -----------------------------------------

                if actividades_unicas:

                    grupos_refacciones.append(
                        {
                            "folio": str(
                                folio
                            ).strip(),

                            "actividades":
                                actividades_unicas
                        }
                    )


        # =================================================
        # MOSTRAR REFACCIONES
        # =================================================

        if grupos_refacciones:

            st.markdown(
                '<div class="separador"></div>',
                unsafe_allow_html=True
            )

            st.markdown(
                '<div class="titulo-seccion">📦 Refacciones pendientes</div>',
                unsafe_allow_html=True
            )


            for grupo_ref in grupos_refacciones:

                folio_mostrado = (
                    grupo_ref["folio"]
                )

                actividades = (
                    grupo_ref["actividades"]
                )


                actividades_html = ""


                for actividad in actividades:

                    actividades_html += (
                        f"""
<div style="
margin-bottom:6px;
padding-left:4px;
">
• {actividad}
</div>
"""
                    )


                st.markdown(
                    f"""
<div class="refaccion-card">

<div class="refaccion-folio">
FOLIO WS: {folio_mostrado}
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

st.markdown(
    '<div class="separador"></div>',
    unsafe_allow_html=True
)


if not folios_abiertos:

    st.markdown(
        """
<div class="sin-folios">
✓ SIN FOLIOS ABIERTOS
</div>
""",
        unsafe_allow_html=True
    )


# =========================================================
# PIE
# =========================================================

st.caption(
    "Casanova | Control de Unidades"
)
