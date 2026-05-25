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
    for campo in KEY_CONFIG[indicator]:
        if campo not in row.index:
            return None
        values.append(normalize_value(row[campo]))
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

def prepare_dataframe(df, origen):
    registros_validos = []
    claves_invalidas = []
    for index, row in df.iterrows():
        key = build_key(row)
        indicador = get_indicator(row.get("ID_INDICADOR"))
        if key is None:
            claves_invalidas.append({
                "ORIGEN": origen,
                "FILA_EXCEL": index + 2,
                "ID_INDICADOR": indicador,
                "OBSERVACION": "No se pudo generar clave: indicador no configurado, columna faltante o campo clave vacío"
            })
        else:
            row_dict = row.to_dict()
            row_dict["CLAVE_COMPARACION"] = key
            registros_validos.append(row_dict)
    return pd.DataFrame(registros_validos), pd.DataFrame(claves_invalidas)

def compare_files(df1, df2, tolerance):
    columnas_v1 = set(df1.columns)
    columnas_v2 = set(df2.columns)
    estructura = pd.DataFrame([
        {"VALIDACION": "Columnas faltantes en SCI 2.0", "DETALLE": ", ".join(sorted(columnas_v1 - columnas_v2)) if columnas_v1 - columnas_v2 else "Sin diferencias"},
        {"VALIDACION": "Columnas faltantes en SCI 1.0", "DETALLE": ", ".join(sorted(columnas_v2 - columnas_v1)) if columnas_v2 - columnas_v1 else "Sin diferencias"},
    ])

    valid1, invalid1 = prepare_dataframe(df1, "SCI 1.0")
    valid2, invalid2 = prepare_dataframe(df2, "SCI 2.0")

    if valid1.empty:
        valid1 = pd.DataFrame(columns=["CLAVE_COMPARACION"])
    if valid2.empty:
        valid2 = pd.DataFrame(columns=["CLAVE_COMPARACION"])

    dup1 = valid1[valid1.duplicated("CLAVE_COMPARACION", keep=False)].copy()
    dup2 = valid2[valid2.duplicated("CLAVE_COMPARACION", keep=False)].copy()

    valid1_unique = valid1.drop_duplicates("CLAVE_COMPARACION", keep="first")
    valid2_unique = valid2.drop_duplicates("CLAVE_COMPARACION", keep="first")

    claves1 = set(valid1_unique["CLAVE_COMPARACION"])
    claves2 = set(valid2_unique["CLAVE_COMPARACION"])

    solo_1 = valid1_unique[valid1_unique["CLAVE_COMPARACION"].isin(claves1 - claves2)].copy()
    solo_2 = valid2_unique[valid2_unique["CLAVE_COMPARACION"].isin(claves2 - claves1)].copy()

    match = pd.merge(
        valid1_unique,
        valid2_unique,
        on="CLAVE_COMPARACION",
        how="inner",
        suffixes=("_SCI_1_0", "_SCI_2_0")
    )

    diferencias_valores = []
    for _, row in match.iterrows():
        indicador = get_indicator(row.get("ID_INDICADOR_SCI_1_0"))
        for campo in VALUE_CONFIG.get(indicador, []):
            col1 = f"{campo}_SCI_1_0"
            col2 = f"{campo}_SCI_2_0"
            valor1_original = row.get(col1)
            valor2_original = row.get(col2)
            valor1 = to_number(valor1_original)
            valor2 = to_number(valor2_original)
            if valor1 is None or valor2 is None:
                diferencias_valores.append({
                    "CLAVE_COMPARACION": row["CLAVE_COMPARACION"],
                    "ID_INDICADOR": indicador,
                    "CAMPO": campo,
                    "VALOR_SCI_1_0": valor1_original,
                    "VALOR_SCI_2_0": valor2_original,
                    "DIFERENCIA_ABSOLUTA": "",
                    "DIFERENCIA_PORCENTUAL": "",
                    "ESTADO": "NO NUMERICO O VACIO"
                })
                continue
            diferencia = abs(valor1 - valor2)
            diferencia_pct = "" if valor1 == 0 else (diferencia / abs(valor1)) * 100
            if diferencia > tolerance:
                diferencias_valores.append({
                    "CLAVE_COMPARACION": row["CLAVE_COMPARACION"],
                    "ID_INDICADOR": indicador,
                    "CAMPO": campo,
                    "VALOR_SCI_1_0": valor1,
                    "VALOR_SCI_2_0": valor2,
                    "DIFERENCIA_ABSOLUTA": diferencia,
                    "DIFERENCIA_PORCENTUAL": diferencia_pct,
                    "ESTADO": "DIFERENCIA"
                })
    diferencias_valores_df = pd.DataFrame(diferencias_valores)

    if "ID_INDICADOR" in df1.columns and "ID_INDICADOR" in df2.columns:
        conteo_id_1 = df1.groupby("ID_INDICADOR").size().reset_index(name="SCI_1_0")
        conteo_id_2 = df2.groupby("ID_INDICADOR").size().reset_index(name="SCI_2_0")
        conteo_indicador = pd.merge(conteo_id_1, conteo_id_2, on="ID_INDICADOR", how="outer").fillna(0)
        conteo_indicador["DIFERENCIA"] = conteo_indicador["SCI_2_0"] - conteo_indicador["SCI_1_0"]
    else:
        conteo_indicador = pd.DataFrame()

    resumen = pd.DataFrame([
        {"VALIDACION": "Total registros SCI 1.0", "RESULTADO": len(df1)},
        {"VALIDACION": "Total registros SCI 2.0", "RESULTADO": len(df2)},
        {"VALIDACION": "Registros con clave válida SCI 1.0", "RESULTADO": len(valid1)},
        {"VALIDACION": "Registros con clave válida SCI 2.0", "RESULTADO": len(valid2)},
        {"VALIDACION": "Registros con match por clave", "RESULTADO": len(match)},
        {"VALIDACION": "Registros solo en SCI 1.0", "RESULTADO": len(solo_1)},
        {"VALIDACION": "Registros solo en SCI 2.0", "RESULTADO": len(solo_2)},
        {"VALIDACION": "Diferencias en PROPORCION / TAMANIO_DE_MUESTRA", "RESULTADO": len(diferencias_valores_df)},
        {"VALIDACION": "Claves duplicadas SCI 1.0", "RESULTADO": len(dup1)},
        {"VALIDACION": "Claves duplicadas SCI 2.0", "RESULTADO": len(dup2)},
        {"VALIDACION": "Claves inválidas SCI 1.0", "RESULTADO": len(invalid1)},
        {"VALIDACION": "Claves inválidas SCI 2.0", "RESULTADO": len(invalid2)},
    ])

    return {
        "Resumen": resumen,
        "Estructura": estructura,
        "Conteo_ID_INDICADOR": conteo_indicador,
        "Match_por_clave": match,
        "Solo_SCI_1_0": solo_1,
        "Solo_SCI_2_0": solo_2,
        "Diferencias_valores": diferencias_valores_df,
        "Duplicados_SCI_1_0": dup1,
        "Duplicados_SCI_2_0": dup2,
        "Claves_invalidas_1_0": invalid1,
        "Claves_invalidas_2_0": invalid2,
    }

def crear_excel(resultados):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for nombre, df in resultados.items():
            if isinstance(df, pd.DataFrame) and not df.empty:
                df.to_excel(writer, sheet_name=nombre[:31], index=False)
            else:
                pd.DataFrame([{"RESULTADO": "Sin registros"}]).to_excel(writer, sheet_name=nombre[:31], index=False)
    return output.getvalue()

with st.sidebar:
    st.header("Configuración")
    tolerance = st.number_input("Tolerancia decimal", min_value=0.0, value=0.0001, step=0.0001, format="%.6f")

archivo_v1 = st.file_uploader("Subir archivo SCI Automático 1.0", type=["xlsx", "csv"], key="v1")
archivo_v2 = st.file_uploader("Subir archivo SCI Automático 2.0", type=["xlsx", "csv"], key="v2")

if archivo_v1 and archivo_v2:
    df1 = read_file(archivo_v1)
    df2 = read_file(archivo_v2)
    resultados = compare_files(df1, df2, tolerance)

    resumen = resultados["Resumen"]
    resumen_dict = dict(zip(resumen["VALIDACION"], resumen["RESULTADO"]))

    st.success("Comparación ejecutada correctamente")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Registros SCI 1.0", resumen_dict["Total registros SCI 1.0"])
    c2.metric("Registros SCI 2.0", resumen_dict["Total registros SCI 2.0"])
    c3.metric("Match por clave", resumen_dict["Registros con match por clave"])
    c4.metric("Diferencias valores", resumen_dict["Diferencias en PROPORCION / TAMANIO_DE_MUESTRA"])

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Solo SCI 1.0", resumen_dict["Registros solo en SCI 1.0"])
    c6.metric("Solo SCI 2.0", resumen_dict["Registros solo en SCI 2.0"])
    c7.metric("Duplicados", resumen_dict["Claves duplicadas SCI 1.0"] + resumen_dict["Claves duplicadas SCI 2.0"])
    c8.metric("Claves inválidas", resumen_dict["Claves inválidas SCI 1.0"] + resumen_dict["Claves inválidas SCI 2.0"])

    excel = crear_excel(resultados)
    st.download_button(
        "📥 Descargar reporte Excel",
        data=excel,
        file_name="Reporte_Comparativo_SCI_Automatico.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    tabs = st.tabs(["Resumen", "Estructura", "Conteo ID_INDICADOR", "Match por clave", "Solo en un archivo", "Diferencias valores", "Duplicados / inválidos"])

    with tabs[0]:
        st.dataframe(resultados["Resumen"], use_container_width=True)
    with tabs[1]:
        st.dataframe(resultados["Estructura"], use_container_width=True)
    with tabs[2]:
        st.dataframe(resultados["Conteo_ID_INDICADOR"], use_container_width=True)
    with tabs[3]:
        st.dataframe(resultados["Match_por_clave"], use_container_width=True)
    with tabs[4]:
        st.subheader("Registros solo en SCI 1.0")
        st.dataframe(resultados["Solo_SCI_1_0"], use_container_width=True)
        st.subheader("Registros solo en SCI 2.0")
        st.dataframe(resultados["Solo_SCI_2_0"], use_container_width=True)
    with tabs[5]:
        if resultados["Diferencias_valores"].empty:
            st.success("No se detectaron diferencias en PROPORCION / TAMANIO_DE_MUESTRA.")
        else:
            st.dataframe(resultados["Diferencias_valores"], use_container_width=True)
    with tabs[6]:
        st.subheader("Duplicados SCI 1.0")
        st.dataframe(resultados["Duplicados_SCI_1_0"], use_container_width=True)
        st.subheader("Duplicados SCI 2.0")
        st.dataframe(resultados["Duplicados_SCI_2_0"], use_container_width=True)
        st.subheader("Claves inválidas SCI 1.0")
        st.dataframe(resultados["Claves_invalidas_1_0"], use_container_width=True)
        st.subheader("Claves inválidas SCI 2.0")
        st.dataframe(resultados["Claves_invalidas_2_0"], use_container_width=True)
else:
    st.info("Carga ambos archivos para iniciar la comparación.")
