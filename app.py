import io
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Comparador SCI", layout="wide")

st.title("📊 Comparador SCI Automático 1.0 vs 2.0")
st.write("Carga ambos archivos para comparar estructura, registros, claves y valores numéricos.")

KEY_CONFIG = {
    3: ["ID_INDICADOR", "ID_FLOTA", "COD_PESQUERIA", "COD_ESPECIE", "MES", "ANIO", "ID_ZONA", "SERIE"],
    4: ["ID_INDICADOR", "ID_FLOTA", "COD_PESQUERIA", "COD_ESPECIE", "MES", "ANIO", "ID_ZONA", "SERIE", "LONGITUD"],
    5: ["ID_INDICADOR", "ID_FLOTA", "COD_PESQUERIA", "COD_ESPECIE", "MES", "ANIO", "ID_ZONA", "SERIE"],
    6: ["ID_INDICADOR", "ID_FLOTA", "COD_PESQUERIA", "COD_ESPECIE", "MES", "ANIO", "ID_ZONA", "SERIE"],
    7: ["ID_INDICADOR", "ID_FLOTA", "COD_PESQUERIA", "COD_ESPECIE", "MES", "ANIO", "ID_ZONA", "SERIE"],
}

VALUE_CONFIG = {
    3: ["PROPORCION"],
    4: ["PROPORCION", "TAMANIO_DE_MUESTRA"],
    5: ["PROPORCION", "TAMANIO_DE_MUESTRA"],
    6: ["PROPORCION", "TAMANIO_DE_MUESTRA"],
    7: ["PROPORCION"],
}


def safe_display(df, max_rows=5000):
    if df is None or df.empty:
        st.info("Sin registros")
        return

    safe_df = df.head(max_rows).copy()
    safe_df = safe_df.fillna("").astype(str)
    st.dataframe(safe_df, use_container_width=True)

    if len(df) > max_rows:
        st.caption(f"Se muestran {max_rows} registros de {len(df)}. El Excel descarga el detalle completo.")


def read_file(file):
    if file.name.lower().endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    df.columns = [str(c).strip().upper() for c in df.columns]
    return df


def normalize_value(value):
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def get_indicator(value):
    try:
        return int(float(value))
    except Exception:
        return None


def build_key(row):
    indicator = get_indicator(row.get("ID_INDICADOR"))

    if indicator not in KEY_CONFIG:
        return None

    values = []
    for field in KEY_CONFIG[indicator]:
        if field not in row.index:
            return None
        values.append(normalize_value(row[field]))

    if "" in values:
        return None

    return "|".join(values)


def to_number(value):
    if pd.isna(value) or value == "":
        return None
    try:
        return float(str(value).replace(",", "."))
    except Exception:
        return None


def prepare_dataframe(df, source_name):
    valid_rows = []
    invalid_keys = []

    for index, row in df.iterrows():
        key = build_key(row)
        indicator = get_indicator(row.get("ID_INDICADOR"))

        if key is None:
            invalid_keys.append({
                "ORIGEN": source_name,
                "FILA_EXCEL": index + 2,
                "ID_INDICADOR": indicator,
                "OBSERVACION": "No se pudo generar clave: indicador no configurado, columna faltante o campo clave vacío"
            })
        else:
            row_dict = row.to_dict()
            row_dict["CLAVE_COMPARACION"] = key
            valid_rows.append(row_dict)

    return pd.DataFrame(valid_rows), pd.DataFrame(invalid_keys)


def compare_files(df1, df2, tolerance):
    columns_v1 = set(df1.columns)
    columns_v2 = set(df2.columns)

    structure = pd.DataFrame([
        {
            "VALIDACION": "Columnas faltantes en SCI 2.0",
            "DETALLE": ", ".join(sorted(columns_v1 - columns_v2)) if columns_v1 - columns_v2 else "Sin diferencias"
        },
        {
            "VALIDACION": "Columnas faltantes en SCI 1.0",
            "DETALLE": ", ".join(sorted(columns_v2 - columns_v1)) if columns_v2 - columns_v1 else "Sin diferencias"
        }
    ])

    valid1, invalid1 = prepare_dataframe(df1, "SCI 1.0")
    valid2, invalid2 = prepare_dataframe(df2, "SCI 2.0")

    if valid1.empty:
        valid1 = pd.DataFrame(columns=["CLAVE_COMPARACION"])
    if valid2.empty:
        valid2 = pd.DataFrame(columns=["CLAVE_COMPARACION"])

    duplicates_1 = valid1[valid1.duplicated("CLAVE_COMPARACION", keep=False)].copy()
    duplicates_2 = valid2[valid2.duplicated("CLAVE_COMPARACION", keep=False)].copy()

    valid1_unique = valid1.drop_duplicates("CLAVE_COMPARACION", keep="first")
    valid2_unique = valid2.drop_duplicates("CLAVE_COMPARACION", keep="first")

    keys_1 = set(valid1_unique["CLAVE_COMPARACION"])
    keys_2 = set(valid2_unique["CLAVE_COMPARACION"])

    only_1 = valid1_unique[valid1_unique["CLAVE_COMPARACION"].isin(keys_1 - keys_2)].copy()
    only_2 = valid2_unique[valid2_unique["CLAVE_COMPARACION"].isin(keys_2 - keys_1)].copy()

    matched = pd.merge(
        valid1_unique,
        valid2_unique,
        on="CLAVE_COMPARACION",
        how="inner",
        suffixes=("_SCI_1_0", "_SCI_2_0")
    )

    value_differences = []

    for _, row in matched.iterrows():
        indicator = get_indicator(row.get("ID_INDICADOR_SCI_1_0"))
        fields_to_compare = VALUE_CONFIG.get(indicator, [])

        for field in fields_to_compare:
            col1 = f"{field}_SCI_1_0"
            col2 = f"{field}_SCI_2_0"

            original_value_1 = row.get(col1)
            original_value_2 = row.get(col2)

            value_1 = to_number(original_value_1)
            value_2 = to_number(original_value_2)

            if value_1 is None or value_2 is None:
                value_differences.append({
                    "CLAVE_COMPARACION": row["CLAVE_COMPARACION"],
                    "ID_INDICADOR": indicator,
                    "CAMPO": field,
                    "VALOR_SCI_1_0": str(original_value_1),
                    "VALOR_SCI_2_0": str(original_value_2),
                    "DIFERENCIA_ABSOLUTA": "N/A",
                    "DIFERENCIA_PORCENTUAL": "N/A",
                    "ESTADO": "NO NUMERICO O VACIO"
                })
                continue

            absolute_difference = abs(value_1 - value_2)
            percentage_difference = "N/A" if value_1 == 0 else (absolute_difference / abs(value_1)) * 100

            if absolute_difference > tolerance:
                value_differences.append({
                    "CLAVE_COMPARACION": row["CLAVE_COMPARACION"],
                    "ID_INDICADOR": indicator,
                    "CAMPO": field,
                    "VALOR_SCI_1_0": value_1,
                    "VALOR_SCI_2_0": value_2,
                    "DIFERENCIA_ABSOLUTA": absolute_difference,
                    "DIFERENCIA_PORCENTUAL": percentage_difference,
                    "ESTADO": "DIFERENCIA"
                })

    value_differences_df = pd.DataFrame(value_differences)

    if "ID_INDICADOR" in df1.columns and "ID_INDICADOR" in df2.columns:
        count_1 = df1.groupby("ID_INDICADOR").size().reset_index(name="SCI_1_0")
        count_2 = df2.groupby("ID_INDICADOR").size().reset_index(name="SCI_2_0")
        count_by_indicator = pd.merge(count_1, count_2, on="ID_INDICADOR", how="outer").fillna(0)
        count_by_indicator["DIFERENCIA"] = count_by_indicator["SCI_2_0"] - count_by_indicator["SCI_1_0"]
    else:
        count_by_indicator = pd.DataFrame()

    duplicate_unique_keys_1 = duplicates_1["CLAVE_COMPARACION"].nunique() if not duplicates_1.empty else 0
    duplicate_unique_keys_2 = duplicates_2["CLAVE_COMPARACION"].nunique() if not duplicates_2.empty else 0

    summary = pd.DataFrame([
        {"VALIDACION": "Total registros SCI 1.0", "RESULTADO": len(df1)},
        {"VALIDACION": "Total registros SCI 2.0", "RESULTADO": len(df2)},
        {"VALIDACION": "Registros con clave válida SCI 1.0", "RESULTADO": len(valid1)},
        {"VALIDACION": "Registros con clave válida SCI 2.0", "RESULTADO": len(valid2)},
        {"VALIDACION": "Registros con match por clave", "RESULTADO": len(matched)},
        {"VALIDACION": "Registros solo en SCI 1.0", "RESULTADO": len(only_1)},
        {"VALIDACION": "Registros solo en SCI 2.0", "RESULTADO": len(only_2)},
        {"VALIDACION": "Diferencias en PROPORCION / TAMANIO_DE_MUESTRA", "RESULTADO": len(value_differences_df)},
        {"VALIDACION": "Filas con clave duplicada SCI 1.0", "RESULTADO": len(duplicates_1)},
        {"VALIDACION": "Filas con clave duplicada SCI 2.0", "RESULTADO": len(duplicates_2)},
        {"VALIDACION": "Claves únicas repetidas SCI 1.0", "RESULTADO": duplicate_unique_keys_1},
        {"VALIDACION": "Claves únicas repetidas SCI 2.0", "RESULTADO": duplicate_unique_keys_2},
        {"VALIDACION": "Claves inválidas SCI 1.0", "RESULTADO": len(invalid1)},
        {"VALIDACION": "Claves inválidas SCI 2.0", "RESULTADO": len(invalid2)},
    ])

    return {
        "Resumen": summary,
        "Estructura": structure,
        "Conteo_ID_INDICADOR": count_by_indicator,
        "Match_por_clave": matched,
        "Solo_SCI_1_0": only_1,
        "Solo_SCI_2_0": only_2,
        "Diferencias_valores": value_differences_df,
        "Duplicados_SCI_1_0": duplicates_1,
        "Duplicados_SCI_2_0": duplicates_2,
        "Claves_invalidas_1_0": invalid1,
        "Claves_invalidas_2_0": invalid2,
    }


def create_excel(results):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet_name, df in results.items():
            if isinstance(df, pd.DataFrame) and not df.empty:
                export_df = df.copy()
                for col in export_df.columns:
                    if export_df[col].dtype == "object":
                        export_df[col] = export_df[col].astype(str)
                export_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            else:
                pd.DataFrame([{"RESULTADO": "Sin registros"}]).to_excel(
                    writer, sheet_name=sheet_name[:31], index=False
                )

    return output.getvalue()


with st.sidebar:
    st.header("Configuración")
    tolerance = st.number_input(
        "Tolerancia decimal",
        min_value=0.0,
        value=0.0001,
        step=0.0001,
        format="%.6f"
    )

file_v1 = st.file_uploader("Subir archivo SCI Automático 1.0", type=["xlsx", "csv"], key="v1")
file_v2 = st.file_uploader("Subir archivo SCI Automático 2.0", type=["xlsx", "csv"], key="v2")

if file_v1 and file_v2:
    with st.spinner("Procesando comparación..."):
        df1 = read_file(file_v1)
        df2 = read_file(file_v2)
        results = compare_files(df1, df2, tolerance)

    summary = results["Resumen"]
    summary_dict = dict(zip(summary["VALIDACION"], summary["RESULTADO"]))

    st.success("Comparación ejecutada correctamente")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Registros SCI 1.0", summary_dict["Total registros SCI 1.0"])
    col2.metric("Registros SCI 2.0", summary_dict["Total registros SCI 2.0"])
    col3.metric("Match por clave", summary_dict["Registros con match por clave"])
    col4.metric("Diferencias valores", summary_dict["Diferencias en PROPORCION / TAMANIO_DE_MUESTRA"])

    col5, col6, col7, col8 = st.columns(4)
    col5.metric("Solo SCI 1.0", summary_dict["Registros solo en SCI 1.0"])
    col6.metric("Solo SCI 2.0", summary_dict["Registros solo en SCI 2.0"])
    col7.metric(
        "Filas duplicadas",
        summary_dict["Filas con clave duplicada SCI 1.0"] + summary_dict["Filas con clave duplicada SCI 2.0"]
    )
    col8.metric(
        "Claves inválidas",
        summary_dict["Claves inválidas SCI 1.0"] + summary_dict["Claves inválidas SCI 2.0"]
    )

    st.caption(
        "Nota: 'Filas duplicadas' cuenta todas las filas que pertenecen a una clave repetida. "
        "En el resumen también se informa la cantidad de claves únicas repetidas."
    )

    excel_report = create_excel(results)

    st.download_button(
        "📥 Descargar reporte Excel",
        data=excel_report,
        file_name="Reporte_Comparativo_SCI_Automatico.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    tabs = st.tabs([
        "Resumen",
        "Estructura",
        "Conteo ID_INDICADOR",
        "Match por clave",
        "Solo en un archivo",
        "Diferencias valores",
        "Duplicados / inválidos"
    ])

    with tabs[0]:
        st.subheader("Resumen general")
        safe_display(results["Resumen"])

    with tabs[1]:
        st.subheader("Validación de estructura")
        safe_display(results["Estructura"])

    with tabs[2]:
        st.subheader("Conteo por ID_INDICADOR")
        safe_display(results["Conteo_ID_INDICADOR"])

    with tabs[3]:
        st.subheader("Registros que coinciden por clave")
        safe_display(results["Match_por_clave"])

    with tabs[4]:
        st.subheader("Registros solo en SCI 1.0")
        safe_display(results["Solo_SCI_1_0"])
        st.subheader("Registros solo en SCI 2.0")
        safe_display(results["Solo_SCI_2_0"])

    with tabs[5]:
        st.subheader("Diferencias en PROPORCION / TAMANIO_DE_MUESTRA")
        if results["Diferencias_valores"].empty:
            st.success("No se detectaron diferencias en PROPORCION / TAMANIO_DE_MUESTRA.")
        else:
            safe_display(results["Diferencias_valores"])

    with tabs[6]:
        st.subheader("Duplicados SCI 1.0")
        safe_display(results["Duplicados_SCI_1_0"])
        st.subheader("Duplicados SCI 2.0")
        safe_display(results["Duplicados_SCI_2_0"])
        st.subheader("Claves inválidas SCI 1.0")
        safe_display(results["Claves_invalidas_1_0"])
        st.subheader("Claves inválidas SCI 2.0")
        safe_display(results["Claves_invalidas_2_0"])

else:
    st.info("Carga ambos archivos para iniciar la comparación.")
