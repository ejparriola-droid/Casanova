# =========================================================
# REFACCIONES
# =========================================================

if not df_refacciones.empty:

    # -----------------------------------------------------
    # IDENTIFICAR COLUMNAS
    # -----------------------------------------------------

    col_serie_ref = None
    col_folio_ref = None
    col_actividad_ref = None

    for columna in df_refacciones.columns:

        clave = (
            str(columna)
            .strip()
            .lower()
        )

        # SERIE
        if clave == "serie":
            col_serie_ref = columna

        # FOLIO WS
        elif clave == "folio ws":
            col_folio_ref = columna

        # ACTIVIDAD
        elif clave == "actividad":
            col_actividad_ref = columna


    # -----------------------------------------------------
    # VALIDAR QUE EXISTAN LAS COLUMNAS
    # -----------------------------------------------------

    if (
        col_serie_ref is not None
        and col_folio_ref is not None
        and col_actividad_ref is not None
    ):

        ref = df_refacciones.copy()


        # -------------------------------------------------
        # NORMALIZAR SERIE DE REFACCIONES
        # -------------------------------------------------

        ref["_SERIE_BUSQUEDA"] = (
            ref[col_serie_ref]
            .astype("string")
            .fillna("")
            .str.strip()
            .str.upper()
        )


        # -------------------------------------------------
        # NORMALIZAR SERIE DE LA UNIDAD
        # -------------------------------------------------

        serie_actual = (
            str(serie)
            .strip()
            .upper()
        )


        # -------------------------------------------------
        # BUSCAR LA SERIE
        # -------------------------------------------------

        ref = ref[
            ref["_SERIE_BUSQUEDA"] == serie_actual
        ].copy()


        # -------------------------------------------------
        # SI ENCONTRÓ INFORMACIÓN
        # -------------------------------------------------

        if not ref.empty:

            grupos_refacciones = []


            # =================================================
            # AGRUPAR POR FOLIO WS
            # =================================================

            for folio, grupo in ref.groupby(
                col_folio_ref,
                sort=False
            ):

                # ---------------------------------------------
                # IGNORAR FOLIOS VACÍOS
                # ---------------------------------------------

                if not tiene_informacion(folio):
                    continue


                actividades_unicas = []


                # ---------------------------------------------
                # RECORRER ACTIVIDADES
                # ---------------------------------------------

                for actividad in grupo[
                    col_actividad_ref
                ]:

                    if not tiene_informacion(actividad):
                        continue


                    actividad_texto = (
                        str(actividad)
                        .strip()
                    )


                    # -----------------------------------------
                    # SOLO ACTIVIDADES DIFERENTES
                    # -----------------------------------------

                    ya_existe = any(
                        actividad_texto.casefold()
                        ==
                        actividad_existente.casefold()
                        for actividad_existente
                        in actividades_unicas
                    )


                    if not ya_existe:

                        actividades_unicas.append(
                            actividad_texto
                        )


                # ---------------------------------------------
                # GUARDAR SOLO SI TIENE ACTIVIDAD
                # ---------------------------------------------

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


                # ---------------------------------------------
                # MOSTRAR CADA FOLIO UNA SOLA VEZ
                # ---------------------------------------------

                for grupo_ref in grupos_refacciones:

                    folio_mostrado = (
                        grupo_ref["folio"]
                    )

                    actividades = (
                        grupo_ref["actividades"]
                    )


                    # -----------------------------------------
                    # ACTIVIDADES EN FORMATO LISTA
                    # -----------------------------------------

                    actividades_html = ""


                    for actividad in actividades:

                        actividades_html += f"""
                        <div style="
                            margin-bottom:6px;
                            padding-left:4px;
                        ">
                            • {actividad}
                        </div>
                        """


                    # -----------------------------------------
                    # TARJETA DEL FOLIO
                    # -----------------------------------------

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

            # -------------------------------------------------
            # SI LA SERIE EXISTE PERO NO TIENE FOLIO VÁLIDO
            # -------------------------------------------------

            else:

                pass

    # ---------------------------------------------------------
    # SI NO EXISTEN LAS COLUMNAS ESPERADAS
    # ---------------------------------------------------------

    else:

        st.warning(
            "La hoja Refacciones no contiene las columnas "
            "'Serie', 'Folio WS' y 'Actividad'."
        )
