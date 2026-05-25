app.py
import io

        summary = results["summary"]
        summary_dict = dict(zip(summary["VALIDACION"], summary["RESULTADO"]))

        metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)
        metric_col1.metric("Registros SCI 1.0", summary_dict.get("Total registros SCI 1.0", 0))
        metric_col2.metric("Registros SCI 2.0", summary_dict.get("Total registros SCI 2.0", 0))
        metric_col3.metric("Registros con match", summary_dict.get("Registros con match", 0))
        metric_col4.metric("Diferencias numéricas", summary_dict.get("Diferencias numéricas", 0))

        metric_col5, metric_col6, metric_col7, metric_col8 = st.columns(4)
        metric_col5.metric("Solo SCI 1.0", summary_dict.get("Registros solo en SCI 1.0", 0))
        metric_col6.metric("Solo SCI 2.0", summary_dict.get("Registros solo en SCI 2.0", 0))
        metric_col7.metric("Duplicados", summary_dict.get("Duplicados SCI 1.0", 0) + summary_dict.get("Duplicados SCI 2.0", 0))
        metric_col8.metric("Claves inválidas", summary_dict.get("Claves inválidas SCI 1.0", 0) + summary_dict.get("Claves inválidas SCI 2.0", 0))

        excel_report = create_excel_report(results)
        st.download_button(
            label="📥 Descargar reporte comparativo Excel",
            data=excel_report,
            file_name="Reporte_Comparativo_SCI_Automatico_1_0_vs_2_0.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "Resumen",
            "Estructura",
            "Conteo",
            "Diferencias numéricas",
            "Solo en un archivo",
            "Errores de datos"
        ])

        with tab1:
            st.subheader("Resumen General")
            st.dataframe(results["summary"], use_container_width=True)

        with tab2:
            st.subheader("Validación de estructura")
            st.dataframe(results["structure"], use_container_width=True)

        with tab3:
            st.subheader("Conteo de registros")
            st.dataframe(results["counts"], use_container_width=True)

        with tab4:
            st.subheader("Diferencias numéricas")
            if results["numeric_differences"].empty:
                st.success("No se detectaron diferencias numéricas sobre la tolerancia configurada.")
            else:
                st.dataframe(results["numeric_differences"], use_container_width=True)

        with tab5:
            st.subheader("Registros solo en SCI 1.0")
            st.dataframe(results["only_v1"], use_container_width=True)
            st.subheader("Registros solo en SCI 2.0")
            st.dataframe(results["only_v2"], use_container_width=True)

        with tab6:
            st.subheader("Duplicados SCI 1.0")
            st.dataframe(results["duplicates_v1"], use_container_width=True)
            st.subheader("Duplicados SCI 2.0")
            st.dataframe(results["duplicates_v2"], use_container_width=True)
            st.subheader("Claves inválidas SCI 1.0")
            st.dataframe(results["invalid_v1"], use_container_width=True)
            st.subheader("Claves inválidas SCI 2.0")
            st.dataframe(results["invalid_v2"], use_container_width=True)

    except Exception as e:
        st.error("Ocurrió un error al procesar los archivos.")
        st.exception(e)
else:
    st.info("Carga ambos archivos para iniciar la comparación.")

st.markdown("---")
st.caption("Validación QA: estructura, conteo, match por clave, diferencias numéricas, duplicados y claves inválidas.")
