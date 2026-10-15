import streamlit as st
import pandas as pd
from pathlib import Path
import hashlib


# =========================================================
# CONFIGURACIÓN
# =========================================================

st.set_page_config(
    page_title="Casanova | Control de Unidades",
    page_icon="🚗",
    layout="wide"
)


# =========================================================
# ARCHIVOS
# =========================================================

ARCHIVO_BASE = Path("concentrado.xlsx")

HOJA_BASE = "concentrado"
HOJA_REFACCIONES = "refacciones"


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

USUARIO_CORRECTO = "Casanova"

# Hash SHA-256 de: Casanova.50
PASSWORD_HASH = hashlib.sha256(
    "Casanova.50".encode("utf-8")
).hexdigest()


def verificar_password(password):

    hash_ingresado = hashlib.sha256(
        password.encode("utf-8")
    ).hexdigest()

    return hash_ingresado == PASSWORD_HASH


def pantalla_login():

    st.markdown(
        f"""
        <style>

        .login-container {{
            max-width: 460px;
            margin: 80px auto 0 auto;
            background: white;
            padding: 38px;
            border-radius: 18px;
            border: 1px solid #D8E1EF;
            box-shadow: 0 8px 30px rgba(33,55,105,0.12);
        }}

        .login-logo {{
            text-align: center;
            font-size: 48px;
            margin-bottom: 5px;
        }}

        .login-title {{
            text-align: center;
            color: {AZUL_MARINO};
            font-size: 26px;
            font-weight: 800;
        }}

        .login-subtitle {{
            text-align: center;
            color: #64748B;
            font-size: 14px;
            margin-bottom: 25px;
        }}

        </style>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="login-container">
            <div class="login-logo">🚗</div>
            <div class="login-title">
                CASANOVA
            </div>
            <div class="login-subtitle">
                Control de Unidades
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    # Espacio visual
    st.write("")

    col1, col2, col3 = st.columns([1, 2, 1])

    with col2:

        usuario = st.text_input(
            "Usuario",
            placeholder="Ingresa tu usuario"
        )

        password = st.text_input(
            "Contraseña",
            type="password",
            placeholder="Ingresa tu contraseña"
        )

        iniciar = st.button(
            "🔐 Iniciar sesión",
            use_container_width=True,
            type="primary"
        )

        if iniciar:

            if (
                usuario.strip() == USUARIO_CORRECTO
                and verificar_password(password)
            ):

                st.session_state["autenticado"] = True

                st.rerun()

            else:

                st.error(
                    "Usuario o contraseña incorrectos."
                )

    st.markdown(
        """
        <div style="
            text-align:center;
            color:#94A3B8;
            font-size:12px;
            margin-top:25px;
        ">
        Acceso restringido
        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# VALIDAR LOGIN
# =========================================================

if "autenticado" not in st.session_state:

    st.session_state["autenticado"] = False


if not st.session_state["autenticado"]:

    pantalla_login()

    st.stop()


# =========================================================
# FUNCIONES AUXILIARES
# =========================================================

def normalizar_columnas(df):

    df = df.copy()

    nuevas_columnas = []

    for columna in df.columns:

        nombre = str(columna)

        nombre = (
            nombre
            .replace("\n", " ")
            .replace("\xa0", " ")
        )

        nombre = " ".join(
            nombre.split()
        )

        nuevas_columnas.append(
            nombre.strip()
        )

    df.columns = nuevas_columnas

    return df


def buscar_columna(df, nombres):

    columnas = {
        str(col).strip().lower(): col
        for col in df.columns
    }

    for nombre in nombres:

        clave = (
            str(nombre)
            .strip()
            .lower()
        )

        if clave in columnas:

            return columnas[clave]

    return None


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

    if pd.isna(valor):
        return False

    texto = str(valor).strip()

    if texto == "":
        return False

    if texto.lower() in [
        "nan",
        "none",
        "null",
        "sin información",
        "sin informacion"
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

        return fecha.strftime(
            "%d/%m/%Y"
        )

    except:

        return valor_visible(valor)


def dias_entero(valor):

    if not tiene_informacion(valor):

        return None

    try:

        numero = float(valor)

        return int(numero)

    except:

        try:

            numero = pd.to_numeric(
                valor,
                errors="coerce"
            )

            if pd.isna(numero):

                return None

            return int(float(numero))

        except:

            return None


def dias_visible(valor):

    numero = dias_entero(valor)

    if numero is None:

        return "🆕 Recién ingresado"

    if numero == 1:

        return "1 día"

    return f"{numero:,} días"


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
# CARGAR DATOS
# =========================================================

@st.cache_data
def cargar_datos():

    df = pd.read_excel(
        ARCHIVO_BASE,
        sheet_name=HOJA_BASE
    )

    df = normalizar_columnas(df)

    return df


@st.cache_data
def cargar_refacciones():

    try:

        df = pd.read_excel(
            ARCHIVO_BASE,
            sheet_name=HOJA_REFACCIONES
        )

        df = normalizar_columnas(df)

        return df

    except Exception:

        return pd.DataFrame()


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
   TOTAL PISO
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
    font-size: 42px;
    font-weight: 800;
    margin-top: 4px;
}}

.total-piso-desglose {{
    font-size: 14px;
    margin-top: 4px;
    opacity: 0.95;
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
   REFACCIONES
   ===================================================== */

.refaccion-card {{
    background: white;
    border: 1px solid #D8E1EF;
    border-left: 5px solid {AMARILLO};
    border-radius: 12px;
    padding: 15px 18px;
    margin-bottom: 12px;
    box-shadow: 0 2px 7px rgba(33,55,105,0.05);
}}

.refaccion-folio {{
    color: {AZUL_MARINO};
    font-size: 15px;
    font-weight: 800;
    margin-bottom: 6px;
}}

.refaccion-texto {{
    color: #172033;
    font-size: 15px;
    font-weight: 600;
}}


/* =====================================================
   SINIESTRO
   ===================================================== */

.siniestro-card {{
    background: #FFF8F0;
    border: 1px solid #FFD9A0;
    border-left: 5px solid #F57C00;
    border-radius: 14px;
    padding: 18px;
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
        '<div class="titulo-principal">🚗 CASANOVA | CONTROL DE UNIDADES</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="subtitulo">Consulta individual de unidades</div>',
        unsafe_allow_html=True
    )


with col_salida:

    if st.button(
        "Cerrar sesión",
        use_container_width=True
    ):

        st.session_state["autenticado"] = False

        st.rerun()


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
# TOTAL DE UNIDADES EN PISO
# =========================================================

columna_piso = buscar_columna(
    df,
    ["Piso"]
)

columna_dias_piso = buscar_columna(
    df,
    ["DÍAS EN PISO", "DIAS EN PISO"]
)


if columna_piso:

    mascara_piso = (
        df[columna_piso]
        .apply(tiene_informacion)
    )

    unidades_en_piso = int(
        mascara_piso.sum()
    )

else:

    if columna_dias_piso:

        unidades_en_piso = int(
            df[columna_dias_piso]
            .apply(tiene_informacion)
            .sum()
        )

    else:

        unidades_en_piso = 0


# =========================================================
# DESGLOSE CENTENARIO / LALA
# =========================================================

centenario = 0
lala = 0


if columna_piso:

    pisos = (
        df[columna_piso]
        .astype("string")
        .fillna("")
        .str.strip()
        .str.upper()
    )

    centenario = int(
        pisos.str.contains(
            "CENTENARIO",
            regex=False,
            na=False
        ).sum()
    )

    lala = int(
        pisos.str.contains(
            "LALA",
            regex=False,
            na=False
        ).sum()
    )


# =========================================================
# TARJETA TOTAL
# =========================================================

st.markdown(
f"""
<div class="total-piso">

<div class="total-piso-titulo">
🏭 TOTAL DE UNIDADES EN PISO
</div>

<div class="total-piso-numero">
{unidades_en_piso:,}
</div>

<div class="total-piso-desglose">
Centenario: <b>{centenario:,}</b>
&nbsp;&nbsp; / &nbsp;&nbsp;
Lala: <b>{lala:,}</b>
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
# IDENTIFICAR COLUMNAS PRINCIPALES
# =========================================================

columna_placa = buscar_columna(
    df,
    ["PLACA"]
)

columna_serie = buscar_columna(
    df,
    ["SERIE"]
)


if not columna_placa or not columna_serie:

    st.error(
        "No se encontraron las columnas PLACA y/o SERIE en concentrado.xlsx."
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
    df[columna_placa]
    .astype("string")
    .fillna("")
    .str.strip()
    .str.upper()
)


series = (
    df[columna_serie]
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
            fila.get(columna_placa)
        )

        serie_opcion = valor_visible(
            fila.get(columna_serie)
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
    unidad.get(columna_placa)
)

serie = valor_visible(
    unidad.get(columna_serie)
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
            valor_visible(
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
                unidad.get("NOMBRE DEL CLIENTE")
            )
        ),
        unsafe_allow_html=True
    )


with col2:

    st.markdown(
        tarjeta(
            "ALMACÉN",
            valor_visible(
                unidad.get("ALMACÉN")
            )
        ),
        unsafe_allow_html=True
    )


with col3:

    st.markdown(
        tarjeta(
            "UNIDAD DE NEGOCIO",
            valor_visible(
                unidad.get("UNIDAD NEGOCIO")
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
                unidad.get("GPS")
            )
        ),
        unsafe_allow_html=True
    )


with col2:

    st.markdown(
        tarjeta(
            "PLATAFORMA",
            valor_visible(
                unidad.get("PLATAFORMA")
            )
        ),
        unsafe_allow_html=True
    )


with col3:

    st.markdown(
        tarjeta(
            "¿REPORTA?",
            valor_visible(
                unidad.get("¿REPORTA?")
            )
        ),
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


valor_dias_piso = (
    unidad.get(columna_dias_piso)
    if columna_dias_piso
    else None
)


dias_piso = dias_visible(
    valor_dias_piso
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

fecha_gestoria = unidad.get(
    "Gestorías1.Fecha de creación"
)

estado_gestoria = unidad.get(
    "Gestorías1.Estado de atención"
)

conceptos_gestoria = unidad.get(
    "Gestorías1.Conceptos"
)


tiene_gestoria = any(
    tiene_informacion(valor)
    for valor in [
        folio_gestoria,
        gestoria,
        fecha_gestoria,
        estado_gestoria,
        conceptos_gestoria
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

    st.markdown(
        '<div class="estado-si">✓ TIENE GESTORÍA</div>',
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
                    unidad.get("Estatus poliza")
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
                valor_visible(
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

    st.markdown(
        '<div class="estado-si">✓ TIENE MANTENIMIENTO</div>',
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

        numero_dias = dias_entero(
            dias_mantenimiento
        )

        if numero_dias is not None:

            texto_dias = (
                f"{numero_dias:,} días"
            )

        else:

            texto_dias = "Sin información"

        st.markdown(
            tarjeta(
                "DÍAS",
                texto_dias
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

    st.markdown(
        '<div class="estado-si">✓ TIENE REGISTRO EN ZIGA TALLERES</div>',
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

        numero_dias = dias_entero(
            dias_ziga
        )

        if numero_dias is not None:

            texto_dias = (
                f"{numero_dias:,} días"
            )

        else:

            texto_dias = "Sin información"

        st.markdown(
            tarjeta(
                "DÍAS",
                texto_dias
            ),
            unsafe_allow_html=True
        )


# =========================================================
# SINIESTRO
# =========================================================

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

siniestro_general = unidad.get(
    "Siniestro"
)


tiene_siniestro = any(
    tiene_informacion(valor)
    for valor in [
        categoria_siniestro,
        dias_siniestro,
        etapa_siniestro,
        fecha_apertura_siniestro,
        fecha_siniestro,
        siniestro_general
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

    st.markdown(
        '<div class="siniestro-card">',
        unsafe_allow_html=True
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.markdown(
            tarjeta(
                "CATEGORÍA",
                valor_visible(
                    categoria_siniestro
                )
            ),
            unsafe_allow_html=True
        )

    with col2:

        st.markdown(
            tarjeta(
                "ETAPA",
                valor_visible(
                    etapa_siniestro
                )
            ),
            unsafe_allow_html=True
        )

    with col3:

        numero_dias = dias_entero(
            dias_siniestro
        )

        if numero_dias is not None:

            texto_dias = (
                f"{numero_dias:,} días"
            )

        else:

            texto_dias = "Sin información"

        st.markdown(
            tarjeta(
                "DÍAS",
                texto_dias
            ),
            unsafe_allow_html=True
        )

    col1, col2 = st.columns(2)

    with col1:

        st.markdown(
            tarjeta(
                "FECHA DE APERTURA",
                fecha_visible(
                    fecha_apertura_siniestro
                )
            ),
            unsafe_allow_html=True
        )

    with col2:

        st.markdown(
            tarjeta(
                "FECHA DEL SINIESTRO",
                fecha_visible(
                    fecha_siniestro
                )
            ),
            unsafe_allow_html=True
        )

    st.markdown(
        '</div>',
        unsafe_allow_html=True
    )


# =========================================================
# REFACCIONES
# =========================================================

if not df_refacciones.empty:

    columna_ref_serie = buscar_columna(
        df_refacciones,
        ["Serie", "SERIE"]
    )

    columna_folio_ws = buscar_columna(
        df_refacciones,
        ["Folio WS", "FOLIO WS"]
    )

    columna_actividad = buscar_columna(
        df_refacciones,
        ["Actividad", "ACTIVIDAD"]
    )

    columna_refaccion = buscar_columna(
        df_refacciones,
        ["Refacción", "REFACCIÓN", "Refaccion", "REFACCION"]
    )


    if columna_ref_serie and columna_folio_ws:

        serie_busqueda = str(
            serie
        ).strip().upper()


        ref = df_refacciones.copy()


        ref["_SERIE_BUSQUEDA"] = (
            ref[columna_ref_serie]
            .astype("string")
            .fillna("")
            .str.strip()
            .str.upper()
        )


        ref = ref[
            ref["_SERIE_BUSQUEDA"]
            == serie_busqueda
        ].copy()


        if not ref.empty:

            # =============================================
            # CREAR TEXTO DE REFACCIÓN
            # =============================================

            if columna_refaccion:

                ref["_DATO_REFACCION"] = (
                    ref[columna_refaccion]
                    .astype("string")
                    .fillna("")
                    .str.strip()
                )

            else:

                ref["_DATO_REFACCION"] = ""


            # =============================================
            # ACTIVIDAD COMO RESPALDO
            # =============================================

            if columna_actividad:

                actividad_limpia = (
                    ref[columna_actividad]
                    .astype("string")
                    .fillna("")
                    .str.strip()
                )

                ref["_DATO_REFACCION"] = ref.apply(
                    lambda fila:
                        fila["_DATO_REFACCION"]
                        if tiene_informacion(
                            fila["_DATO_REFACCION"]
                        )
                        else actividad_limpia.loc[
                            fila.name
                        ],
                    axis=1
                )


            # =============================================
            # LIMPIAR REGISTROS VACÍOS
            # =============================================

            ref = ref[
                ref["_DATO_REFACCION"]
                .apply(tiene_informacion)
            ].copy()


            if not ref.empty:

                # =========================================
                # AGRUPAR POR FOLIO WS
                # =========================================

                registros_refacciones = []


                for folio, grupo in ref.groupby(
                    columna_folio_ws,
                    sort=False
                ):

                    folio_texto = valor_visible(
                        folio
                    )


                    lista_refacciones = []

                    for valor in grupo[
                        "_DATO_REFACCION"
                    ]:

                        texto = str(
                            valor
                        ).strip()

                        if texto == "":
                            continue

                        # Evitar duplicados
                        if texto.lower() not in [
                            x.lower()
                            for x in lista_refacciones
                        ]:

                            lista_refacciones.append(
                                texto
                            )


                    if lista_refacciones:

                        registros_refacciones.append(
                            (
                                folio_texto,
                                " / ".join(
                                    lista_refacciones
                                )
                            )
                        )


                # =========================================
                # MOSTRAR REFACCIONES
                # =========================================

                if registros_refacciones:

                    st.markdown(
                        '<div class="separador"></div>',
                        unsafe_allow_html=True
                    )

                    st.markdown(
                        '<div class="titulo-seccion">📦 Refacciones</div>',
                        unsafe_allow_html=True
                    )


                    for folio, refacciones_texto in (
                        registros_refacciones
                    ):

                        st.markdown(
                            f"""
                            <div class="refaccion-card">

                            <div class="refaccion-folio">
                            FOLIO WS: {folio}
                            </div>

                            <div class="refaccion-texto">
                            {refacciones_texto}
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
