import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import date

# =========================================================
# CONFIGURACIÓN
# =========================================================
st.set_page_config(
    page_title="Casanova Control de Unidades",
    page_icon="🚗",
    layout="wide"
)

ARCHIVO_MAESTRO = Path(r"Y:\Master Data\Consultas_Estatus-Subestatus.xlsx")
ARCHIVO_PISOS = Path(r"Y:\KPI´s MONITOREO ESTRATEGICO\PISOS CENTENARIO.xlsx")
HOJAS_PISOS = ["CENTENARIO", "LALA"]

ESTATUS_TALLERES_ABIERTOS = [
    "Entrada", "En taller", "Pendiente de entrega",
    "En validación", "En taller externo"
]
ESTATUS_MANTENIMIENTOS_ABIERTOS = [
    "Vehículo ingresado al taller", "Orden de Servicio en Proceso",
    "Solicitud Aprobada por el Administrador", "Vehículo listo para entrega",
    "Solicitud en Aprobación"
]
ESTATUS_GESTORIAS_ABIERTOS = [
    "Confirmación de servicio", "Orden de Servicio Cerrada por el Cliente",
    "Solicitud en Aprobación", "Orden de Servicio Finalizada",
    "Orden de Servicio en Proceso", "Carga de documentos",
    "Validación de documentos"
]


def limpiar_texto(serie):
    return (
        serie.astype("string")
        .str.strip()
        .str.upper()
        .replace("", pd.NA)
    )


def limpiar_columnas(df):
    df = df.copy()
    df.columns = df.columns.astype(str).str.replace("\t", "", regex=False).str.strip()
    return df


def valor_visible(valor):
    if pd.isna(valor):
        return "Sin información"
    texto = str(valor).strip()
    return texto if texto else "Sin información"


def dias_transcurridos(fecha):
    fecha = pd.to_datetime(fecha, errors="coerce")
    if pd.isna(fecha):
        return pd.NA
    return max((pd.Timestamp(date.today()) - fecha.normalize()).days, 0)


def preparar_claves(df, columna_serie, columna_placa):
    df = df.copy()
    if columna_serie in df.columns:
        df["_SERIE_CLAVE"] = limpiar_texto(df[columna_serie])
    else:
        df["_SERIE_CLAVE"] = pd.NA
    if columna_placa in df.columns:
        df["_PLACA_CLAVE"] = limpiar_texto(df[columna_placa])
    else:
        df["_PLACA_CLAVE"] = pd.NA
    return df


@st.cache_data
def cargar_unidades():
    df = limpiar_columnas(pd.read_excel(ARCHIVO_MAESTRO, sheet_name="Ziga"))
    return preparar_claves(df, "Serie", "Placas")


@st.cache_data
def cargar_total_pisos():
    """Cuenta todos los registros de CENTENARIO y LALA, antes de deduplicar."""
    total = 0
    for hoja in HOJAS_PISOS:
        piso = pd.read_excel(ARCHIVO_PISOS, sheet_name=hoja, header=1, usecols=[0])
        total += len(piso)
    return total


@st.cache_data
def cargar_pisos():
    """Une CENTENARIO y LALA; conserva el registro con más días por unidad para las búsquedas."""
    registros = []
    for hoja in HOJAS_PISOS:
        piso = limpiar_columnas(pd.read_excel(ARCHIVO_PISOS, sheet_name=hoja, header=1))
        piso = preparar_claves(piso, "SERIE", "PLACA")
        piso["PISO_REGISTRO"] = hoja
        if "DÍAS EN PISO" not in piso.columns:
            piso["DÍAS EN PISO"] = pd.NA
        piso["DÍAS EN PISO"] = pd.to_numeric(piso["DÍAS EN PISO"], errors="coerce")
        if "FECHA DE INGRESO" in piso.columns:
            piso["FECHA DE INGRESO"] = pd.to_datetime(piso["FECHA DE INGRESO"], errors="coerce")
        registros.append(piso)

    pisos = pd.concat(registros, ignore_index=True, sort=False)
    pisos = pisos[pisos["_SERIE_CLAVE"].notna() | pisos["_PLACA_CLAVE"].notna()].copy()
    pisos["_CLAVE_UNIDAD"] = pisos["_SERIE_CLAVE"]
    sin_serie = pisos["_CLAVE_UNIDAD"].isna()
    pisos.loc[sin_serie, "_CLAVE_UNIDAD"] = "PLACA:" + pisos.loc[sin_serie, "_PLACA_CLAVE"].astype("string")
    pisos["_ORDEN_ORIGINAL"] = range(len(pisos))
    pisos = pisos.sort_values(["DÍAS EN PISO", "_ORDEN_ORIGINAL"], ascending=[False, True], na_position="last")
    return pisos.drop_duplicates("_CLAVE_UNIDAD", keep="first")


@st.cache_data
def cargar_procesos():
    procesos = {}

    # Ziga Talleres
    df = limpiar_columnas(pd.read_excel(ARCHIVO_MAESTRO, sheet_name="Ziga Talleres"))
    df = preparar_claves(df, "No. de serie", "No. de placa")
    df["Estatus"] = df["Estatus"].astype("string").str.strip()
    df = df[df["Estatus"].isin(ESTATUS_TALLERES_ABIERTOS)].copy()
    df["DÍAS ABIERTOS"] = df["Fecha de ingreso"].apply(dias_transcurridos)
    df["ÁREA CON RESGUARDO"] = df["Área del taller activa"].astype("string").str.contains("Resguardo", case=False, na=False)
    procesos["Ziga Talleres"] = df

    # Mantenimientos
    df = limpiar_columnas(pd.read_excel(ARCHIVO_MAESTRO, sheet_name="Mantenimientos"))
    df = preparar_claves(df, "NO. SERIE", "PLACAS")
    df["ESTATUS DEL FOLIO"] = df["ESTATUS DEL FOLIO"].astype("string").str.strip()
    df = df[df["ESTATUS DEL FOLIO"].isin(ESTATUS_MANTENIMIENTOS_ABIERTOS)].copy()
    df["DÍAS ABIERTOS"] = df["FECHA DE CREACION DEL FOLIO"].apply(dias_transcurridos)
    procesos["Mantenimientos"] = df

    # Gestorías
    df = limpiar_columnas(pd.read_excel(ARCHIVO_MAESTRO, sheet_name="Gestorías"))
    df = preparar_claves(df, "Serie del Vehículo", "Placas del vehículo")
    df["Estado de atención"] = df["Estado de atención"].astype("string").str.strip()
    df = df[df["Estado de atención"].isin(ESTATUS_GESTORIAS_ABIERTOS)].copy()
    df["DÍAS ABIERTOS"] = df["Fecha de creación"].apply(dias_transcurridos)
    df["CONCEPTO GESTORÍA/TRÁMITES"] = df["Conceptos"].astype("string").str.contains(r"Gestoria|Tramites", case=False, na=False, regex=True)
    procesos["Gestorías"] = df

    # Siniestros: todos los registros cuyo STATUS sea Abierto.
    df = limpiar_columnas(pd.read_excel(ARCHIVO_MAESTRO, sheet_name="Siniestros"))
    df = preparar_claves(df, "SERIE", "PLACAS")
    df["STATUS"] = df["STATUS"].astype("string").str.strip()
    df = df[df["STATUS"].str.casefold().eq("abierto")].copy()
    df["DÍAS ABIERTOS"] = df["FECHA APERTURA"].apply(dias_transcurridos)
    procesos["Siniestros"] = df

    return procesos


def buscar_registros_unidad(df_proceso, serie, placa):
    """Busca por serie primero y recurre a placa si no encuentra coincidencias."""
    if df_proceso.empty:
        return df_proceso.copy()
    if pd.notna(serie):
        encontrados = df_proceso[df_proceso["_SERIE_CLAVE"].eq(serie)]
        if not encontrados.empty:
            return encontrados.copy()
    if pd.notna(placa):
        return df_proceso[df_proceso["_PLACA_CLAVE"].eq(placa)].copy()
    return df_proceso.iloc[0:0].copy()


def mostrar_proceso(titulo, registros, columnas, columna_resaltado=None):
    st.markdown(f"**{titulo}** — {len(registros)} registro(s) abierto(s)")
    if registros.empty:
        st.caption("Sin procesos abiertos relacionados con esta unidad.")
        return
    columnas = [c for c in columnas if c in registros.columns]
    vista = registros[columnas].copy()
    if columna_resaltado and columna_resaltado in registros.columns:
        vista["ALERTA"] = registros[columna_resaltado].map(
            lambda x: "⚠️ Revisar" if bool(x) else ""
        ).values
    st.dataframe(vista, use_container_width=True, hide_index=True)


# =========================================================
# INTERFAZ
# =========================================================
st.title("CASANOVA | CONTROL DE UNIDADES")
st.caption("CORPORATIVO / LEASING")
st.divider()

try:
    df = cargar_unidades()

    try:
        pisos = cargar_pisos()
        total_pisos = cargar_total_pisos()
        error_pisos = None
    except Exception as error:
        pisos = pd.DataFrame()
        total_pisos = None
        error_pisos = error

    try:
        procesos = cargar_procesos()
        error_procesos = None
    except Exception as error:
        procesos = {}
        error_procesos = error

    st.subheader("Consulta de unidades")
    busqueda = st.text_input("Buscar por placas o número de serie", placeholder="Escribe las placas o la serie...")

    if busqueda:
        criterio = busqueda.strip().upper()
        resultado = df[
            df["_PLACA_CLAVE"].fillna("").str.contains(criterio, case=False, na=False, regex=False)
            | df["_SERIE_CLAVE"].fillna("").str.contains(criterio, case=False, na=False, regex=False)
        ].copy()

        if resultado.empty:
            st.warning("No se encontraron unidades en el archivo maestro.")
        else:
            st.success(f"Unidades encontradas en el maestro: {len(resultado):,}")
            columnas_maestro = [
                "Marca", "Línea", "Versión", "Modelo", "Serie", "Placas",
                "Unidad de negocio", "Nombre del cliente", "Estatus", "Substatus"
            ]
            columnas_maestro = [c for c in columnas_maestro if c in resultado.columns]

            for _, unidad in resultado.iterrows():
                serie = unidad.get("_SERIE_CLAVE", pd.NA)
                placa = unidad.get("_PLACA_CLAVE", pd.NA)
                placa_titulo = valor_visible(unidad.get("Placas", ""))
                serie_titulo = valor_visible(unidad.get("Serie", ""))
                st.markdown("---")
                st.markdown(f"### 🚗 Placas: {placa_titulo} | Serie: {serie_titulo}")
                st.dataframe(pd.DataFrame([unidad[columnas_maestro]]), use_container_width=True, hide_index=True)

                # Información de pisos
                st.markdown("#### 📍 Pisos CENTENARIO / LALA")
                registro_piso = pd.DataFrame()
                if not pisos.empty:
                    if pd.notna(serie):
                        registro_piso = pisos[pisos["_SERIE_CLAVE"].eq(serie)]
                    if registro_piso.empty and pd.notna(placa):
                        registro_piso = pisos[pisos["_PLACA_CLAVE"].eq(placa)]
                if registro_piso.empty:
                    st.info("Sin registro en pisos.")
                else:
                    columnas_piso = [
                        "PLACA", "SERIE", "PISO_REGISTRO", "ALMACÉN", "FECHA DE INGRESO",
                        "DÍAS EN PISO", "OBSERVACIONES", "Siniestro", "Talleres y Mantenimientos",
                        "Gestoria", "Estatus poliza"
                    ]
                    columnas_piso = [c for c in columnas_piso if c in registro_piso.columns]
                    st.dataframe(registro_piso[columnas_piso], use_container_width=True, hide_index=True)

                # Procesos relacionados con la unidad
                st.markdown("#### 🧾 Procesos abiertos relacionados")
                if error_procesos is not None:
                    st.warning("No se pudieron cargar las pestañas de procesos.")
                    st.code(str(error_procesos))
                else:
                    con_proceso = st.container()
                    with con_proceso:
                        talleres = buscar_registros_unidad(procesos["Ziga Talleres"], serie, placa)
                        mostrar_proceso(
                            "Ziga Talleres", talleres,
                            ["Folio", "Estatus", "Área del taller activa", "Fecha de ingreso", "DÍAS ABIERTOS"],
                            "ÁREA CON RESGUARDO"
                        )
                        if not talleres.empty and talleres["ÁREA CON RESGUARDO"].any():
                            st.warning("Esta unidad tiene un área de taller que contiene ‘Resguardo’. Revisar el registro.")

                        mantenimientos = buscar_registros_unidad(procesos["Mantenimientos"], serie, placa)
                        mostrar_proceso(
                            "Mantenimientos", mantenimientos,
                            ["FOLIO", "ESTATUS DEL FOLIO", "PROVEEDOR", "TIPO DE SERVICIO",
                             "DESCRIPCIÓN DEL SERVICIO", "FECHA DE CREACION DEL FOLIO", "DÍAS ABIERTOS"]
                        )

                        gestorias = buscar_registros_unidad(procesos["Gestorías"], serie, placa)
                        mostrar_proceso(
                            "Gestorías", gestorias,
                            ["Folio", "Estado de atención", "Conceptos", "Fecha de creación", "DÍAS ABIERTOS"],
                            "CONCEPTO GESTORÍA/TRÁMITES"
                        )
                        if not gestorias.empty and gestorias["CONCEPTO GESTORÍA/TRÁMITES"].any():
                            st.info("El concepto de gestoría contiene ‘Gestoria’ o ‘Tramites’.")

                        siniestros = buscar_registros_unidad(procesos["Siniestros"], serie, placa)
                        mostrar_proceso(
                            "Siniestros", siniestros,
                            ["FOLIO", "STATUS", "ULTIMO COMENTARIO", "FECHA APERTURA", "DÍAS ABIERTOS"]
                        )

    if error_pisos is not None:
        st.warning("No se pudo cargar la información de CENTENARIO y LALA. El maestro sigue disponible.")
        st.code(str(error_pisos))

    if error_procesos is not None:
        st.warning("No se pudieron cargar las pestañas de procesos. Verifica que los nombres y el archivo sean correctos.")
        st.code(str(error_procesos))

    st.divider()
    st.subheader("Panel general")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        if total_pisos is not None:
            st.metric("Total de unidades en pisos", f"{total_pisos:,}")
        else:
            st.metric("Total de unidades en pisos", "N/D")
    with col2:
        if "Estatus" in df.columns:
            estatus = df["Estatus"].astype("string").str.strip()
            st.metric("Rentadas", f"{estatus.eq('Rentado').sum():,}")
        else:
            st.metric("Rentadas", "N/D")
    with col3:
        if "Substatus" in df.columns:
            substatus = df["Substatus"].astype("string").str.strip()
            st.metric("En reparación", f"{substatus.eq('En reparación').sum():,}")
        else:
            st.metric("En reparación", "N/D")
    with col4:
        st.metric("Con información de serie", f"{df['Serie'].notna().sum():,}" if "Serie" in df.columns else "N/D")

    if error_procesos is None:
        st.divider()
        st.subheader("Resumen de procesos abiertos")
        resumen = st.columns(4)
        nombres = ["Ziga Talleres", "Mantenimientos", "Gestorías", "Siniestros"]
        for col, nombre in zip(resumen, nombres):
            with col:
                st.metric(nombre, f"{len(procesos[nombre]):,}")

except Exception as error:
    st.error("No se pudo cargar el archivo maestro.")
    st.write("Verifica que la unidad Y: esté disponible y que el archivo no haya cambiado de ubicación.")
    st.code(str(error))
