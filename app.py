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
# ARCHIVOS
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
# FUNCIONES GENERALES
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
        "nat",
        "sin información",
        "sin informacion"
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

        return f"{int(numero):,}"

    except:

        return valor_visible(valor)


def dias_visible(valor):

    if not tiene_informacion(valor):

        return "RECIÉN INGRESADO"

    try:

        numero = float(valor)

        return f"{int(numero):,} días"

    except:

        return valor_visible(valor)


def fecha_visible(valor):

    if not tiene_informacion(valor):

        return "Sin información"

    try:

        return pd.to_datetime(valor).strftime("%d/%m/%Y")

    except:

        return valor_visible(valor)


def limpiar_columna(texto):

    return (
        str(texto)
        .strip()
        .replace("\xa0", " ")
        .replace("\n", " ")
        .replace("  ", " ")
        .strip()
    )


# =========================================================
# TARJETAS
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
# CARGAR CONCENTRADO
# =========================================================

@st.cache_data
def cargar_datos():

    df = pd.read_excel(
        ARCHIVO_BASE,
        sheet_name=HOJA_BASE
    )

    df.columns = [
        limpiar_columna(c)
        for c in df.columns
    ]

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

        df.columns = [
            limpiar_columna(c)
            for c in df.columns
        ]

        return df

    except Exception:

        return pd.DataFrame()


# =========================================================
# LOGIN
# =========================================================

def mostrar_login():

    st.markdown(
        f"""
<style>

.login-contenedor {{
    max-width: 520px;
    margin: 90px auto 0 auto;
}}

.login-caja {{
    background: white;
    padding: 35px;
    border-radius: 18px;
    box-shadow: 0 8px 30px rgba(33,55,105,0.12);
    border: 1px solid #D8E1EF;
}}

.login-titulo {{
    text-align: center;
    color: {AZUL_MARINO};
    font-size: 30px;
    font-weight: 800;
    margin-bottom: 5px;
}}

.login-subtitulo {{
    text-align: center;
    color: #64748B;
    font-size: 15px;
    margin-bottom: 30px;
}}

</style>
""",
        unsafe_allow_html=True
    )

    st.markdown(
        """
<div class="login-contenedor">
<div class="login-caja">
<div class="login-titulo">🚗 CASANOVA</div>
<div class="login-subtitulo">Control de Unidades</div>
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

            st.session_state["autenticado"] = True

            st.rerun()

        else:

            st.error(
                "Usuario o contraseña incorrectos."
            )


# =========================================================
# VALIDAR LOGIN
# =========================================================

if "autenticado" not in st.session_state:

    st.session_state["autenticado"] = False


if not st.session_state["autenticado"]:

    mostrar_login()

    st.stop()


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
    padding: 18px;
    text-align: center;
    box-shadow: 0 3px 10px rgba(255,183,0,0.12);
}}

.tarjeta-refaccion {{
    background: white;
    border: 1px solid #D8E1EF;
    border-left: 5px solid {AZUL_CASANOVA};
    border-radius: 12px;
    padding: 16px;
    margin-bottom: 12px;
    box-shadow: 0 2px 8px rgba(33,55,105,0.06);
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
    margin: 20px 0 25px 0;
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

.total-desglose {{
    font-size: 14px;
    font-weight: 600;
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


df_refacciones = cargar_refacciones()


# =========================================================
# TOTAL DE UNIDADES
# =========================================================
#
# SE CUENTA LA COLUMNA PISO
# Y NO DÍAS EN PISO
#

if "Piso" in df.columns:

    piso_limpio = (
        df["Piso"]
        .astype("string")
        .fillna("")
        .str.strip()
    )

    piso_limpio = (
        piso_limpio
        .replace("", pd.NA)
    )

    total_unidades = int(
        piso_limpio.notna().sum()
    )

else:

    total_unidades = len(df)


# =========================================================
# CENTENARIO Y LALA
# =========================================================

if "Piso" in df.columns:

    piso_texto = (
        df["Piso"]
        .astype("string")
        .fillna("")
        .str.strip()
        .str.upper()
    )

    total_centenario = int(
        piso_texto.eq("CENTENARIO").sum()
    )

    total_lala = int(
        piso_texto.eq("LALA").sum()
    )

else:

    total_centenario = 0
    total_lala = 0


# =========================================================
# TARJETA TOTAL
# =========================================================

st.markdown(
    f"""
<div class="total-piso">
<div class="total-piso-titulo">🚗 TOTAL DE UNIDADES</div>
<div class="total-piso-numero">{total_unidades:,}</div>
<div class="total-desglose">
Centenario: <b>{total_centenario:,}</b>
&nbsp;&nbsp; / &nbsp;&nbsp;
Lala: <b>{total_lala:,}</b>
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
# RESUMEN DE PISO
# =========================================================

st.markdown(
    f"""
<div class="tarjeta">
<div class="campo">🏭 UNIDADES EN PISO</div>
<div class="valor">
Centenario: {total_centenario:,}
&nbsp;&nbsp; / &nbsp;&nbsp;
Lala: {total_lala:,}
&nbsp;&nbsp; / &nbsp;&nbsp;
Total: {total_unidades:,}
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
# PISO DE LA UNIDAD
# =========================================================

st.markdown(
    '<div class="separador"></div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="titulo-seccion">🏭 Piso</div>',
    unsafe_allow_html=True
)


piso_unidad = valor_visible(
    unidad.get("Piso")
)


dias_piso = dias_visible(
    unidad.get("DÍAS EN PISO")
)


col1, col2 = st.columns(2)


with col1:

    st.markdown(
        tarjeta(
            "PISO",
            piso_unidad
        ),
        unsafe_allow_html=True
    )


with col2:

    st.markdown(
        f"""
<div class="tarjeta-piso">
<div class="dias-label">DÍAS EN PISO</div>
<div class="dias-grande">{dias_piso}</div>
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

    campos_gestoria = [

        (
            "FOLIO",
            valor_visible(
                unidad.get(
                    "Gestorías1.Folio"
                )
            )
        ),

        (
            "ESTATUS DE PÓLIZA",
            valor_visible(
                unidad.get(
                    "Estatus poliza"
                )
            )
        ),

        (
            "FECHA DE CREACIÓN",
            fecha_visible(
                unidad.get(
                    "Gestorías1.Fecha de creación"
                )
            )
        ),

        (
            "ESTADO DE ATENCIÓN",
            valor_visible(
                unidad.get(
                    "Gestorías1.Estado de atención"
                )
            )
        ),

        (
            "DÍAS",
            numero_entero(
                unidad.get(
                    "Gestorías1.Personalizado"
                )
            )
        ),

        (
            "CONCEPTOS",
            valor_visible(
                unidad.get(
                    "Gestorías1.Conceptos"
                )
            )
        )
    ]

    campos_gestoria = [
        x for x in campos_gestoria
        if tiene_informacion(
            unidad.get(
                {
                    "FOLIO": "Gestorías1.Folio",
                    "ESTATUS DE PÓLIZA": "Estatus poliza",
                    "FECHA DE CREACIÓN": "Gestorías1.Fecha de creación",
                    "ESTADO DE ATENCIÓN": "Gestorías1.Estado de atención",
                    "DÍAS": "Gestorías1.Personalizado",
                    "CONCEPTOS": "Gestorías1.Conceptos"
                }[x[0]]
            )
        )
    ]

    columnas = st.columns(
        min(3, len(campos_gestoria))
    )

    for i, (campo, valor) in enumerate(
        campos_gestoria
    ):

        with columnas[i % len(columnas)]:

            st.markdown(
                tarjeta(
                    campo,
                    valor
                ),
                unsafe_allow_html=True
            )


# =========================================================
# MANTENIMIENTO
# =========================================================

folio_mantenimiento = unidad.get(
    "Mantenimiento.FOLIO"
)


if tiene_informacion(
    folio_mantenimiento
):

    st.markdown(
        '<div class="separador"></div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="titulo-seccion">🔧 Mantenimiento</div>',
        unsafe_allow_html=True
    )

    campos_mantenimiento = [

        (
            "FOLIO",
            "Mantenimiento.FOLIO",
            valor_visible(
                unidad.get(
                    "Mantenimiento.FOLIO"
                )
            )
        ),

        (
            "FECHA DE CREACIÓN",
            "Mantenimiento.FECHA DE CREACION   DEL FOLIO",
            fecha_visible(
                unidad.get(
                    "Mantenimiento.FECHA DE CREACION   DEL FOLIO"
                )
            )
        ),

        (
            "ESTATUS",
            "Mantenimiento.ESTATUS DEL FOLIO",
            valor_visible(
                unidad.get(
                    "Mantenimiento.ESTATUS DEL FOLIO"
                )
            )
        ),

        (
            "PROVEEDOR",
            "Mantenimiento.PROVEEDOR",
            valor_visible(
                unidad.get(
                    "Mantenimiento.PROVEEDOR"
                )
            )
        ),

        (
            "DÍAS",
            "Mantenimiento.dias",
            numero_entero(
                unidad.get(
                    "Mantenimiento.dias"
                )
            )
        )
    ]

    campos_mantenimiento = [
        x for x in campos_mantenimiento
        if tiene_informacion(
            unidad.get(x[1])
        )
    ]

    if campos_mantenimiento:

        columnas = st.columns(
            min(3, len(campos_mantenimiento))
        )

        for i, (campo, _, valor) in enumerate(
            campos_mantenimiento
        ):

            with columnas[
                i % len(columnas)
            ]:

                st.markdown(
                    tarjeta(
                        campo,
                        valor
                    ),
                    unsafe_allow_html=True
                )


# =========================================================
# ZIGA TALLERES
# =========================================================

folio_ziga = unidad.get(
    "Ziga_Talleres.Folio"
)


if tiene_informacion(
    folio_ziga
):

    st.markdown(
        '<div class="separador"></div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="titulo-seccion">🛠️ Ziga Talleres</div>',
        unsafe_allow_html=True
    )

    campos_ziga = [

        (
            "FOLIO",
            "Ziga_Talleres.Folio",
            valor_visible(
                unidad.get(
                    "Ziga_Talleres.Folio"
                )
            )
        ),

        (
            "ESTATUS",
            "Ziga_Talleres.Estatus",
            valor_visible(
                unidad.get(
                    "Ziga_Talleres.Estatus"
                )
            )
        ),

        (
            "ÁREA DEL TALLER ACTIVA",
            "Ziga_Talleres.Área del taller   activa",
            valor_visible(
                unidad.get(
                    "Ziga_Talleres.Área del taller   activa"
                )
            )
        ),

        (
            "FECHA DE INGRESO",
            "Ziga_Talleres.Fecha de ingreso",
            fecha_visible(
                unidad.get(
                    "Ziga_Talleres.Fecha de ingreso"
                )
            )
        ),

        (
            "DÍAS",
            "Ziga_Talleres.dias",
            numero_entero(
                unidad.get(
                    "Ziga_Talleres.dias"
                )
            )
        )
    ]

    campos_ziga = [
        x for x in campos_ziga
        if tiene_informacion(
            unidad.get(x[1])
        )
    ]

    if campos_ziga:

        columnas = st.columns(
            min(3, len(campos_ziga))
        )

        for i, (campo, _, valor) in enumerate(
            campos_ziga
        ):

            with columnas[
                i % len(columnas)
            ]:

                st.markdown(
                    tarjeta(
                        campo,
                        valor
                    ),
                    unsafe_allow_html=True
                )


# =========================================================
# SINIESTRO
# =========================================================

campos_siniestro = [

    (
        "SINIESTRO",
        "Siniestro",
        valor_visible(
            unidad.get("Siniestro")
        )
    ),

    (
        "CATEGORÍA",
        "Siniestros2.CATEGORIA",
        valor_visible(
            unidad.get(
                "Siniestros2.CATEGORIA"
            )
        )
    ),

    (
        "DÍAS",
        "Siniestros2.dias",
        numero_entero(
            unidad.get(
                "Siniestros2.dias"
            )
        )
    ),

    (
        "ETAPA",
        "Siniestros2.ETAPA",
        valor_visible(
            unidad.get(
                "Siniestros2.ETAPA"
            )
        )
    ),

    (
        "FECHA APERTURA",
        "Siniestros2.FECHA APERTURA",
        fecha_visible(
            unidad.get(
                "Siniestros2.FECHA APERTURA"
            )
        )
    ),

    (
        "FECHA SINIESTRO",
        "Siniestros2.FECHA SINIESTRO",
        fecha_visible(
            unidad.get(
                "Siniestros2.FECHA SINIESTRO"
            )
        )
    )
]


campos_siniestro = [
    x for x in campos_siniestro
    if tiene_informacion(
        unidad.get(x[1])
    )
]


if campos_siniestro:

    st.markdown(
        '<div class="separador"></div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="titulo-seccion">🚨 Siniestro</div>',
        unsafe_allow_html=True
    )

    columnas = st.columns(
        min(3, len(campos_siniestro))
    )

    for i, (campo, _, valor) in enumerate(
        campos_siniestro
    ):

        with columnas[
            i % len(columnas)
        ]:

            st.markdown(
                tarjeta(
                    campo,
                    valor
                ),
                unsafe_allow_html=True
            )


# =========================================================
# REFACCIONES
# =========================================================

if not df_refacciones.empty:

    if "SERIE" in df_refacciones.columns:

        serie_buscada = str(
            unidad.get("SERIE")
        ).strip().upper()

        ref = df_refacciones.copy()

        ref["SERIE_BUSQUEDA"] = (
            ref["SERIE"]
            .astype("string")
            .fillna("")
            .str.strip()
            .str.upper()
        )

        ref_unidad = ref[
            ref["SERIE_BUSQUEDA"]
            == serie_buscada
        ].copy()

    else:

        ref_unidad = pd.DataFrame()


else:

    ref_unidad = pd.DataFrame()


# =========================================================
# MOSTRAR REFACCIONES
# =========================================================

if not ref_unidad.empty:

    st.markdown(
        '<div class="separador"></div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="titulo-seccion">🔩 Refacciones pendientes</div>',
        unsafe_allow_html=True
    )


    # -----------------------------------------------------
    # AGRUPAR POR FOLIO WS
    # -----------------------------------------------------

    grupos_refacciones = []


    for folio_ws, grupo in ref_unidad.groupby(
        "Folio WS",
        dropna=False
    ):

        folio_texto = valor_visible(
            folio_ws
        )


        actividades = []

        if "Actividad" in grupo.columns:

            for actividad in grupo["Actividad"]:

                if tiene_informacion(
                    actividad
                ):

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

                if tiene_informacion(
                    refaccion
                ):

                    texto = str(
                        refaccion
                    ).strip()

                    if texto not in refacciones:

                        refacciones.append(
                            texto
                        )


        grupos_refacciones.append(
            {
                "folio": folio_texto,
                "actividades": actividades,
                "refacciones": refacciones
            }
        )


    # -----------------------------------------------------
    # MOSTRAR CADA FOLIO UNA SOLA VEZ
    # -----------------------------------------------------

    for grupo in grupos_refacciones:

        actividades_texto = (
            " / ".join(
                grupo["actividades"]
            )
            if grupo["actividades"]
            else "Sin actividad"
        )


        refacciones_texto = (
            " / ".join(
                grupo["refacciones"]
            )
            if grupo["refacciones"]
            else "Sin refacción"
        )


        st.markdown(
            f"""
<div class="tarjeta-refaccion">
<div class="campo">FOLIO WS</div>
<div class="valor">{grupo["folio"]}</div>

<div class="campo" style="margin-top:12px;">
ACTIVIDAD
</div>
<div class="valor">
{actividades_texto}
</div>

<div class="campo" style="margin-top:12px;">
REFACCIÓN
</div>
<div class="valor">
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
