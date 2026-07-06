import io
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Certificación y Comparador SCI", layout="wide")
st.title("📊 Certificación y Comparador SCI Automático")
st.write("Herramienta QA para validar el archivo SCI 2.0 o comparar SCI 1.0 vs SCI 2.0.")

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

COLUMNAS_OBLIGATORIAS_2_0 = [
    "ID_PROCESO", "ID_INDICADOR", "ID_FLOTA", "COD_PESQUERIA", "COD_ESPECIE",
    "MES", "ANIO", "ID_ZONA", "SERIE", "PROPORCION", "LONGITUD",
    "TAMANIO_DE_MUESTRA", "ID_ESCALA", "FECHA_GENERACION"
]


def safe_display(df, max_rows=5000):
    if df is None or df.empty:
        st.info("Sin registros")
        return
    safe_df = df.head(max_rows).copy().fillna("").astype(str)
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


def values_are_equal(value_1, value_2, tolerance):
    num_1 = to_number(value_1)
    num_2 = to_number(value_2)
    if num_1 is None and num_2 is None:
        return normalize_value(value_1) == normalize_value(value_2)
    if num_1 is None or num_2 is None:
        return False
    return abs(num_1 - num_2) <= tolerance


def get_difference(value_1, value_2):
    num_1 = to_number(value_1)
    num_2 = to_number(value_2)
    if num_1 is None or num_2 is None:
        return "N/A", "N/A"
    absolute_difference = abs(num_1 - num_2)
    percentage_difference = "N/A" if num_1 == 0 else (absolute_difference / abs(num_1)) * 100
    return absolute_difference, percentage_difference


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
                pd.DataFrame([{"RESULTADO": "Sin registros"}]).to_excel(writer, sheet_name=sheet_name[:31], index=False)
    return output.getvalue()


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


# =========================================================
# MODO VALIDAR SOLO SCI 2.0
# =========================================================
def validar_estructura_2_0(df):
    columnas_archivo = set(df.columns)
    columnas_esperadas = set(COLUMNAS_OBLIGATORIAS_2_0)
    faltantes = sorted(list(columnas_esperadas - columnas_archivo))
    adicionales = sorted(list(columnas_archivo - columnas_esperadas))
    return pd.DataFrame([
        {"VALIDACION": "Columnas obligatorias esperadas", "RESULTADO": len(COLUMNAS_OBLIGATORIAS_2_0)},
        {"VALIDACION": "Columnas presentes en archivo", "RESULTADO": len(df.columns)},
        {"VALIDACION": "Columnas faltantes", "RESULTADO": ", ".join(faltantes) if faltantes else "Sin columnas faltantes"},
        {"VALIDACION": "Columnas adicionales", "RESULTADO": ", ".join(adicionales) if adicionales else "Sin columnas adicionales"},
        {"VALIDACION": "Estado estructura", "RESULTADO": "APROBADA" if not faltantes else "CON OBSERVACIONES"},
    ])


def validar_tipos_datos_2_0(df):
    reglas = {
        "ID_PROCESO": "entero", "ID_INDICADOR": "entero", "ID_FLOTA": "entero",
        "COD_PESQUERIA": "entero", "COD_ESPECIE": "entero", "MES": "entero",
        "ANIO": "entero", "ID_ZONA": "entero", "SERIE": "entero",
        "PROPORCION": "decimal", "LONGITUD": "decimal", "TAMANIO_DE_MUESTRA": "decimal",
    }
    resultados = []
    for campo, tipo in reglas.items():
        if campo not in df.columns:
            resultados.append({"CAMPO": campo, "TIPO_ESPERADO": tipo, "REGISTROS_INVALIDOS": "No aplica", "ESTADO": "COLUMNA NO EXISTE"})
            continue
        serie = df[campo]
        valores = pd.to_numeric(serie, errors="coerce")
        invalidos = valores.isna() & serie.notna() & (serie.astype(str).str.strip() != "")
        if tipo == "entero":
            no_enteros = valores.notna() & (valores % 1 != 0)
            total_invalidos = int(invalidos.sum() + no_enteros.sum())
        else:
            total_invalidos = int(invalidos.sum())
        resultados.append({"CAMPO": campo, "TIPO_ESPERADO": tipo, "REGISTROS_INVALIDOS": total_invalidos, "ESTADO": "OK" if total_invalidos == 0 else "CON OBSERVACIONES"})
    return pd.DataFrame(resultados)


def validar_nulos_2_0(df):
    registros = []
    for columna in df.columns:
        vacios = df[columna].isna().sum() + (df[columna].astype(str).str.strip() == "").sum()
        registros.append({"CAMPO": columna, "VALORES_NULOS_O_VACIOS": int(vacios), "ESTADO": "OK" if vacios == 0 else "CON OBSERVACIONES"})
    return pd.DataFrame(registros)


def validar_rangos_2_0(df):
    validaciones = []
    reglas = [
        ("MES", "Debe estar entre 1 y 12", lambda s: (pd.to_numeric(s, errors="coerce") < 1) | (pd.to_numeric(s, errors="coerce") > 12)),
        ("ANIO", "Debe estar entre 2000 y 2100", lambda s: (pd.to_numeric(s, errors="coerce") < 2000) | (pd.to_numeric(s, errors="coerce") > 2100)),
        ("ID_INDICADOR", "Debe ser 3, 4, 5, 6 o 7", lambda s: ~pd.to_numeric(s, errors="coerce").isin([3, 4, 5, 6, 7])),
        ("PROPORCION", "Debe estar entre 0 y 1", lambda s: (pd.to_numeric(s, errors="coerce") < 0) | (pd.to_numeric(s, errors="coerce") > 1)),
        ("LONGITUD", "Debe ser mayor que 0 cuando exista valor", lambda s: (pd.to_numeric(s, errors="coerce") <= 0) & s.notna()),
        ("TAMANIO_DE_MUESTRA", "Debe ser mayor o igual que 0 cuando exista valor", lambda s: (pd.to_numeric(s, errors="coerce") < 0) & s.notna()),
    ]
    for campo, regla, funcion in reglas:
        if campo not in df.columns:
            validaciones.append({"CAMPO": campo, "REGLA": regla, "REGISTROS_FUERA_DE_RANGO": "No aplica", "ESTADO": "COLUMNA NO EXISTE"})
            continue
        cantidad = int(funcion(df[campo]).sum())
        validaciones.append({"CAMPO": campo, "REGLA": regla, "REGISTROS_FUERA_DE_RANGO": cantidad, "ESTADO": "OK" if cantidad == 0 else "CON OBSERVACIONES"})
    return pd.DataFrame(validaciones)


def validar_claves_duplicados_2_0(df):
    valid_df, invalid_df = prepare_dataframe(df, "SCI 2.0")
    if valid_df.empty:
        valid_df = pd.DataFrame(columns=["CLAVE_COMPARACION"])
    duplicados = valid_df[valid_df.duplicated("CLAVE_COMPARACION", keep=False)].copy()
    resumen = pd.DataFrame([
        {"VALIDACION": "Total registros archivo SCI 2.0", "RESULTADO": len(df)},
        {"VALIDACION": "Registros con clave válida", "RESULTADO": len(valid_df)},
        {"VALIDACION": "Registros con clave inválida", "RESULTADO": len(invalid_df)},
        {"VALIDACION": "Filas con clave duplicada", "RESULTADO": len(duplicados)},
        {"VALIDACION": "Claves únicas repetidas", "RESULTADO": duplicados["CLAVE_COMPARACION"].nunique() if not duplicados.empty else 0},
        {"VALIDACION": "Estado claves", "RESULTADO": "OK" if len(invalid_df) == 0 else "CON OBSERVACIONES"},
    ])
    return resumen, valid_df, invalid_df, duplicados


def conteos_distribucion_2_0(df):
    resultados = {}
    campos = ["ID_INDICADOR", "COD_PESQUERIA", "COD_ESPECIE", "MES", "ANIO"]
    for campo in campos:
        if campo in df.columns:
            resultados[f"Conteo_{campo}"] = df.groupby(campo).size().reset_index(name="CANTIDAD")
    return resultados


def generar_validacion_2_0(df):
    estructura = validar_estructura_2_0(df)
    tipos = validar_tipos_datos_2_0(df)
    nulos = validar_nulos_2_0(df)
    rangos = validar_rangos_2_0(df)
    resumen_claves, valid_df, invalid_df, duplicados = validar_claves_duplicados_2_0(df)
    distribuciones = conteos_distribucion_2_0(df)

    observaciones = 0
    observaciones += int((estructura["RESULTADO"].astype(str).str.contains("CON OBSERVACIONES")).sum())
    observaciones += int((tipos["ESTADO"] != "OK").sum())
    observaciones += int((nulos["ESTADO"] != "OK").sum())
    observaciones += int((rangos["ESTADO"] != "OK").sum())
    claves_invalidas = len(invalid_df)
    duplicadas = len(duplicados)
    estado_general = "APROBADO" if observaciones == 0 and claves_invalidas == 0 and duplicadas == 0 else "APROBADO CON OBSERVACIONES"

    resumen_general = pd.DataFrame([
        {"VALIDACION": "Total registros SCI 2.0", "RESULTADO": len(df)},
        {"VALIDACION": "Total columnas SCI 2.0", "RESULTADO": len(df.columns)},
        {"VALIDACION": "Registros con clave válida", "RESULTADO": len(valid_df)},
        {"VALIDACION": "Registros con clave inválida", "RESULTADO": len(invalid_df)},
        {"VALIDACION": "Filas con clave duplicada", "RESULTADO": len(duplicados)},
        {"VALIDACION": "Estado general validación", "RESULTADO": estado_general},
    ])

    resultados = {
        "Resumen_Validacion_2_0": resumen_general,
        "Estructura_2_0": estructura,
        "Tipos_Datos_2_0": tipos,
        "Nulos_Vacios_2_0": nulos,
        "Rangos_2_0": rangos,
        "Resumen_Claves_2_0": resumen_claves,
        "Registros_Clave_Valida": valid_df,
        "Claves_Invalidas_2_0": invalid_df,
        "Duplicados_2_0": duplicados,
    }
    resultados.update(distribuciones)
    return resultados


# =========================================================
# MODO COMPARAR SCI 1.0 VS SCI 2.0
# =========================================================
def compare_files(df1, df2, tolerance):
    columns_v1 = set(df1.columns)
    columns_v2 = set(df2.columns)
    structure = pd.DataFrame([
        {"VALIDACION": "Columnas faltantes en SCI 2.0", "DETALLE": ", ".join(sorted(columns_v1 - columns_v2)) if columns_v1 - columns_v2 else "Sin diferencias"},
        {"VALIDACION": "Columnas faltantes en SCI 1.0", "DETALLE": ", ".join(sorted(columns_v2 - columns_v1)) if columns_v2 - columns_v1 else "Sin diferencias"}
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
    matched = pd.merge(valid1_unique, valid2_unique, on="CLAVE_COMPARACION", how="inner", suffixes=("_SCI_1_0", "_SCI_2_0"))

    value_differences = []
    matched_value_summary = []
    keys_with_identical_values = 0
    keys_with_any_difference = 0
    keys_with_only_proportion_difference = 0
    keys_with_only_sample_size_difference = 0
    keys_with_both_differences = 0
    keys_with_proportion_difference = 0
    keys_with_sample_size_difference = 0

    for _, row in matched.iterrows():
        indicator = get_indicator(row.get("ID_INDICADOR_SCI_1_0"))
        fields_to_compare = VALUE_CONFIG.get(indicator, [])
        has_proportion_difference = False
        has_sample_size_difference = False
        compared_fields = 0
        equal_fields = 0
        for field in fields_to_compare:
            col1 = f"{field}_SCI_1_0"
            col2 = f"{field}_SCI_2_0"
            original_value_1 = row.get(col1)
            original_value_2 = row.get(col2)
            compared_fields += 1
            if values_are_equal(original_value_1, original_value_2, tolerance):
                equal_fields += 1
                continue
            absolute_difference, percentage_difference = get_difference(original_value_1, original_value_2)
            if field == "PROPORCION":
                has_proportion_difference = True
            if field == "TAMANIO_DE_MUESTRA":
                has_sample_size_difference = True
            value_differences.append({
                "CLAVE_COMPARACION": row["CLAVE_COMPARACION"],
                "ID_INDICADOR": indicator,
                "CAMPO": field,
                "VALOR_SCI_1_0": original_value_1,
                "VALOR_SCI_2_0": original_value_2,
                "DIFERENCIA_ABSOLUTA": absolute_difference,
                "DIFERENCIA_PORCENTUAL": percentage_difference,
                "ESTADO": "DIFERENCIA"
            })

        if compared_fields > 0 and compared_fields == equal_fields:
            keys_with_identical_values += 1
            estado_clave = "VALORES IDENTICOS"
        else:
            estado_clave = "CON DIFERENCIAS"
        if has_proportion_difference:
            keys_with_proportion_difference += 1
        if has_sample_size_difference:
            keys_with_sample_size_difference += 1
        if has_proportion_difference or has_sample_size_difference:
            keys_with_any_difference += 1
        if has_proportion_difference and not has_sample_size_difference:
            keys_with_only_proportion_difference += 1
        if has_sample_size_difference and not has_proportion_difference:
            keys_with_only_sample_size_difference += 1
        if has_proportion_difference and has_sample_size_difference:
            keys_with_both_differences += 1
        matched_value_summary.append({
            "CLAVE_COMPARACION": row["CLAVE_COMPARACION"],
            "ID_INDICADOR": indicator,
            "CAMPOS_COMPARADOS": compared_fields,
            "CAMPOS_IGUALES": equal_fields,
            "DIFERENCIA_PROPORCION": "SI" if has_proportion_difference else "NO",
            "DIFERENCIA_TAMANIO_DE_MUESTRA": "SI" if has_sample_size_difference else "NO",
            "ESTADO_CLAVE": estado_clave
        })

    value_differences_df = pd.DataFrame(value_differences)
    matched_value_summary_df = pd.DataFrame(matched_value_summary)

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
        {"VALIDACION": "Claves con valores numéricos idénticos", "RESULTADO": keys_with_identical_values},
        {"VALIDACION": "Claves con alguna diferencia numérica", "RESULTADO": keys_with_any_difference},
        {"VALIDACION": "Claves con diferencia en PROPORCION", "RESULTADO": keys_with_proportion_difference},
        {"VALIDACION": "Claves con diferencia en TAMANIO_DE_MUESTRA", "RESULTADO": keys_with_sample_size_difference},
        {"VALIDACION": "Claves con diferencia solo en PROPORCION", "RESULTADO": keys_with_only_proportion_difference},
        {"VALIDACION": "Claves con diferencia solo en TAMANIO_DE_MUESTRA", "RESULTADO": keys_with_only_sample_size_difference},
        {"VALIDACION": "Claves con diferencia en ambos campos", "RESULTADO": keys_with_both_differences},
        {"VALIDACION": "Diferencias registradas por campo", "RESULTADO": len(value_differences_df)},
        {"VALIDACION": "Registros solo en SCI 1.0", "RESULTADO": len(only_1)},
        {"VALIDACION": "Registros solo en SCI 2.0", "RESULTADO": len(only_2)},
        {"VALIDACION": "Filas con clave duplicada SCI 1.0", "RESULTADO": len(duplicates_1)},
        {"VALIDACION": "Filas con clave duplicada SCI 2.0", "RESULTADO": len(duplicates_2)},
        {"VALIDACION": "Claves únicas repetidas SCI 1.0", "RESULTADO": duplicate_unique_keys_1},
        {"VALIDACION": "Claves únicas repetidas SCI 2.0", "RESULTADO": duplicate_unique_keys_2},
        {"VALIDACION": "Claves inválidas SCI 1.0", "RESULTADO": len(invalid1)},
        {"VALIDACION": "Claves inválidas SCI 2.0", "RESULTADO": len(invalid2)},
    ])
    value_summary_chart = pd.DataFrame([
        {"CATEGORIA": "Claves con valores idénticos", "CANTIDAD": keys_with_identical_values},
        {"CATEGORIA": "Diferencia solo PROPORCION", "CANTIDAD": keys_with_only_proportion_difference},
        {"CATEGORIA": "Diferencia solo TAMANIO_DE_MUESTRA", "CANTIDAD": keys_with_only_sample_size_difference},
        {"CATEGORIA": "Diferencia en ambos campos", "CANTIDAD": keys_with_both_differences},
    ])
    return {
        "Resumen": summary,
        "Estructura": structure,
        "Conteo_ID_INDICADOR": count_by_indicator,
        "Match_por_clave": matched,
        "Resumen_valores_match": matched_value_summary_df,
        "Grafico_resumen_valores": value_summary_chart,
        "Solo_SCI_1_0": only_1,
        "Solo_SCI_2_0": only_2,
        "Diferencias_valores": value_differences_df,
        "Duplicados_SCI_1_0": duplicates_1,
        "Duplicados_SCI_2_0": duplicates_2,
        "Claves_invalidas_1_0": invalid1,
        "Claves_invalidas_2_0": invalid2,
    }


# =========================================================
# INTERFAZ
# =========================================================
with st.sidebar:
    st.header("Configuración")
    modo = st.radio("Modo de trabajo", ["Validar archivo SCI 2.0", "Comparar SCI 1.0 vs SCI 2.0"])
    tolerance = st.number_input("Tolerancia decimal", min_value=0.0, value=0.0001, step=0.0001, format="%.6f")

if modo == "Validar archivo SCI 2.0":
    st.header("✅ Validación individual archivo SCI 2.0")
    file_2 = st.file_uploader("Subir archivo SCI Automático 2.0", type=["xlsx", "csv"], key="validacion_2")
    if file_2:
        with st.spinner("Validando archivo SCI 2.0..."):
            df_2 = read_file(file_2)
            validation_results = generar_validacion_2_0(df_2)
        resumen_validacion = validation_results["Resumen_Validacion_2_0"]
        resumen_dict = dict(zip(resumen_validacion["VALIDACION"], resumen_validacion["RESULTADO"]))
        st.success("Validación ejecutada correctamente")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total registros", resumen_dict["Total registros SCI 2.0"])
        col2.metric("Total columnas", resumen_dict["Total columnas SCI 2.0"])
        col3.metric("Claves válidas", resumen_dict["Registros con clave válida"])
        col4.metric("Claves inválidas", resumen_dict["Registros con clave inválida"])
        col5, col6 = st.columns(2)
        col5.metric("Filas duplicadas", resumen_dict["Filas con clave duplicada"])
        col6.metric("Estado general", resumen_dict["Estado general validación"])
        excel_validation = create_excel(validation_results)
        st.download_button("📥 Descargar reporte validación SCI 2.0", data=excel_validation, file_name="Reporte_Validacion_SCI_2_0.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        tabs = st.tabs(["Resumen", "Estructura", "Tipos de datos", "Nulos/Vacíos", "Rangos", "Claves y duplicados", "Distribuciones"])
        with tabs[0]:
            st.subheader("Resumen validación SCI 2.0")
            safe_display(validation_results["Resumen_Validacion_2_0"])
        with tabs[1]:
            st.subheader("Validación de estructura")
            safe_display(validation_results["Estructura_2_0"])
        with tabs[2]:
            st.subheader("Validación de tipos de datos")
            safe_display(validation_results["Tipos_Datos_2_0"])
        with tabs[3]:
            st.subheader("Validación de valores nulos o vacíos")
            safe_display(validation_results["Nulos_Vacios_2_0"])
        with tabs[4]:
            st.subheader("Validación de rangos")
            safe_display(validation_results["Rangos_2_0"])
        with tabs[5]:
            st.subheader("Resumen de claves")
            safe_display(validation_results["Resumen_Claves_2_0"])
            st.subheader("Claves inválidas")
            safe_display(validation_results["Claves_Invalidas_2_0"])
            st.subheader("Duplicados")
            safe_display(validation_results["Duplicados_2_0"])
        with tabs[6]:
            for key in ["Conteo_ID_INDICADOR", "Conteo_COD_PESQUERIA", "Conteo_COD_ESPECIE", "Conteo_MES", "Conteo_ANIO"]:
                if key in validation_results:
                    st.subheader(key.replace("_", " "))
                    safe_display(validation_results[key])
                    if key == "Conteo_ID_INDICADOR":
                        st.bar_chart(validation_results[key].set_index("ID_INDICADOR")["CANTIDAD"])
    else:
        st.info("Carga el archivo SCI 2.0 para iniciar la validación.")

elif modo == "Comparar SCI 1.0 vs SCI 2.0":
    st.header("🔄 Comparación SCI 1.0 vs SCI 2.0")
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
        col4.metric("Valores idénticos", summary_dict["Claves con valores numéricos idénticos"])
        col5, col6, col7, col8 = st.columns(4)
        col5.metric("Con diferencia numérica", summary_dict["Claves con alguna diferencia numérica"])
        col6.metric("Solo SCI 1.0", summary_dict["Registros solo en SCI 1.0"])
        col7.metric("Solo SCI 2.0", summary_dict["Registros solo en SCI 2.0"])
        col8.metric("Filas duplicadas", summary_dict["Filas con clave duplicada SCI 1.0"] + summary_dict["Filas con clave duplicada SCI 2.0"])
        st.caption("Nota: 'Valores idénticos' cuenta claves que hicieron match y cuyos valores comparados son iguales entre SCI 1.0 y SCI 2.0 según la tolerancia configurada.")
        excel_report = create_excel(results)
        st.download_button("📥 Descargar reporte comparativo Excel", data=excel_report, file_name="Reporte_Comparativo_SCI_Automatico.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        tabs = st.tabs(["Resumen", "Estructura", "Conteo ID_INDICADOR", "Match por clave", "Resumen valores match", "Solo en un archivo", "Diferencias valores", "Duplicados / inválidos"])
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
            st.subheader("Resumen de valores numéricos para claves con match")
            chart_df = results["Grafico_resumen_valores"]
            if chart_df.empty:
                st.info("Sin datos para graficar.")
            else:
                st.bar_chart(chart_df.set_index("CATEGORIA")["CANTIDAD"])
            safe_display(results["Resumen_valores_match"])
        with tabs[5]:
            st.subheader("Registros solo en SCI 1.0")
            safe_display(results["Solo_SCI_1_0"])
            st.subheader("Registros solo en SCI 2.0")
            safe_display(results["Solo_SCI_2_0"])
        with tabs[6]:
            st.subheader("Diferencias en PROPORCION / TAMANIO_DE_MUESTRA")
            if results["Diferencias_valores"].empty:
                st.success("No se detectaron diferencias en PROPORCION / TAMANIO_DE_MUESTRA.")
            else:
                safe_display(results["Diferencias_valores"])
        with tabs[7]:
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
