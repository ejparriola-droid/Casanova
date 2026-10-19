# =========================================================
# BUSCADOR
# (reemplaza TODO desde aquí hasta justo antes de
#  "# SI NO HAY BÚSQUEDA")
# =========================================================

titulo_seccion(
    "🔎 Buscar unidad"
)


# ---------------------------------------------------------
# SESSION STATE
# ---------------------------------------------------------

if "busqueda" not in st.session_state:
    st.session_state["busqueda"] = ""

if "foto_procesada" not in st.session_state:
    st.session_state["foto_procesada"] = ""

if "msg_foto" not in st.session_state:
    st.session_state["msg_foto"] = ("", "")


# ---------------------------------------------------------
# ELEGIR CÓMO BUSCAR (la cámara NO se enciende sola)
# ---------------------------------------------------------

OPCION_ESCRIBIR = "⌨️ Escribir"
OPCION_ARCHIVO = "🖼️ Galería / cámara trasera"
OPCION_CAMARA = "📸 Cámara de la página"

modo = st.radio(
    "¿Cómo quieres buscar?",
    [OPCION_ESCRIBIR, OPCION_ARCHIVO, OPCION_CAMARA],
    horizontal=True,
    label_visibility="collapsed"
)

foto = None

if modo == OPCION_ARCHIVO:

    foto = st.file_uploader(
        "Selecciona una foto o toma una nueva",
        type=["jpg", "jpeg", "png", "webp"]
    )

    st.caption(
        "En el celular, al pulsar el botón puedes elegir "
        "'Galería' o 'Tomar foto' (cámara trasera)."
    )

elif modo == OPCION_CAMARA:

    foto = st.camera_input(
        "Tomar foto",
        label_visibility="collapsed"
    )


# ---------------------------------------------------------
# PROCESAR FOTO (una sola vez por foto)
# ---------------------------------------------------------

if foto is not None:

    datos_foto = foto.getvalue()

    hash_foto = hashlib.md5(datos_foto).hexdigest()

    if st.session_state["foto_procesada"] != hash_foto:

        st.session_state["foto_procesada"] = hash_foto

        with st.spinner("🔎 Leyendo placa / serie..."):
            texto_ocr = leer_texto_foto(datos_foto)

        if texto_ocr:

            # Se escribe en el buscador ANTES de crearlo
            st.session_state["busqueda"] = texto_ocr

            st.session_state["msg_foto"] = (
                "ok",
                f"📷 Texto detectado: {texto_ocr} "
                "— revisa que sea correcto."
            )

        else:

            st.session_state["msg_foto"] = (
                "warn",
                "No pude leer la placa o serie. "
                "Intenta con una foto más cercana "
                "y con buena iluminación."
            )

    tipo_msg, texto_msg = st.session_state["msg_foto"]

    if tipo_msg == "ok":
        st.success(texto_msg)
    elif tipo_msg == "warn":
        st.warning(texto_msg)


# ---------------------------------------------------------
# CAMPO DE BÚSQUEDA
# ---------------------------------------------------------

busqueda = st.text_input(
    "Busca por PLACA o SERIE",
    placeholder="Ejemplo: PDN1295 o número de serie...",
    label_visibility="collapsed",
    key="busqueda"
)
