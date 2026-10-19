                        st.write(f"- {error}")

        except Exception as e:
            st.error(f"Error durante el procesamiento: {e}")


# ============================================================
# VISTA PREVIA Y DESCARGA
# ============================================================

if "df_consolidado" in st.session_state:
    df_mostrar = st.session_state["df_consolidado"]

    st.divider()
    st.subheader("4. Vista previa del consolidado")

    c1, c2, c3 = st.columns(3)

    c1.metric("Total de registros", f"{len(df_mostrar):,}")

    c2.metric(
        "Registros con costo",
        f"{df_mostrar['costo_unitario'].notna().sum():,}"
    )

    c3.metric(
        "Costo total",
        f"{df_mostrar['costo_total'].sum(skipna=True):,.2f}"
    )

    # --------------------------------------------------------
    # Copia para mostrar: la fecha se convierte a texto para
    # que la celda no cambie de tamaño al ser seleccionada.
    # --------------------------------------------------------
    df_vista = df_mostrar.head(500).copy()

    df_vista["fecha"] = df_vista["fecha"].apply(
        lambda v: (
            ""
            if pd.isna(v)
            else v.strftime("%d/%m/%Y")
            if isinstance(v, (pd.Timestamp, datetime))
            else str(v)
        )
    )

    st.dataframe(
        df_vista,
        use_container_width=True,
        hide_index=True,
        column_config=CONFIG_COLUMNAS,
    )

    salida_excel = io.BytesIO()

    with pd.ExcelWriter(
        salida_excel,
        engine="openpyxl"
    ) as writer:
        df_mostrar.to_excel(
            writer,
            index=False,
            sheet_name="Consolidado"
        )

    st.download_button(
        label="📥 Descargar consolidado en Excel",
        data=salida_excel.getvalue(),
        file_name="CONSOLIDADO_GENERAL.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

    if st.button(
        "☁️ EXPORTAR A GOOGLE SHEETS",
        type="primary",
        use_container_width=True
    ):
        try:
            with st.spinner(
                "Enviando registros a Google Sheets..."
            ):
                mensaje = exportar_google_sheets(
                    df_mostrar,
                    modo_exportacion
                )

            st.success(mensaje)

            st.link_button(
                "Abrir Google Sheets",
                (
                    "https://docs.google.com/spreadsheets/d/"
                    + SPREADSHEET_ID
                    + "/edit"
                ),
                use_container_width=True
            )

        except Exception as e:
            st.error(
                "No se pudo exportar a Google Sheets. "
                f"Detalle: {e}"
            )
