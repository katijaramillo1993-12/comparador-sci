import streamlit as st
import pandas as pd

st.set_page_config(page_title="Comparador SCI", layout="wide")

st.title("📊 Comparador SCI Automático 1.0 vs 2.0")
st.write("Carga ambos archivos para comparar diferencias.")

archivo_v1 = st.file_uploader(
    "Subir archivo SCI 1.0",
    type=["xlsx", "csv"],
    key="v1"
)

archivo_v2 = st.file_uploader(
    "Subir archivo SCI 2.0",
    type=["xlsx", "csv"],
    key="v2"
)

if archivo_v1 and archivo_v2:
    if archivo_v1.name.endswith(".csv"):
        df1 = pd.read_csv(archivo_v1)
    else:
        df1 = pd.read_excel(archivo_v1)

    if archivo_v2.name.endswith(".csv"):
        df2 = pd.read_csv(archivo_v2)
    else:
        df2 = pd.read_excel(archivo_v2)

    st.success("Archivos cargados correctamente")

    df1.columns = [str(c).strip().upper() for c in df1.columns]
    df2.columns = [str(c).strip().upper() for c in df2.columns]

    columnas_v1 = set(df1.columns)
    columnas_v2 = set(df2.columns)

    faltan_v2 = columnas_v1 - columnas_v2
    faltan_v1 = columnas_v2 - columnas_v1

    st.subheader("🔎 Validación estructura")

    col1, col2 = st.columns(2)

    with col1:
        st.write("Columnas faltantes en SCI 2.0")
        st.write(list(faltan_v2))

    with col2:
        st.write("Columnas faltantes en SCI 1.0")
        st.write(list(faltan_v1))

    st.subheader("📈 Totales")

    c1, c2 = st.columns(2)
    c1.metric("Registros SCI 1.0", len(df1))
    c2.metric("Registros SCI 2.0", len(df2))

    if "ID_INDICADOR" in df1.columns and "ID_INDICADOR" in df2.columns:
        resumen1 = df1.groupby("ID_INDICADOR").size().reset_index(name="SCI_1_0")
        resumen2 = df2.groupby("ID_INDICADOR").size().reset_index(name="SCI_2_0")

        comparacion = pd.merge(
            resumen1,
            resumen2,
            on="ID_INDICADOR",
            how="outer"
        ).fillna(0)

        comparacion["DIFERENCIA"] = comparacion["SCI_2_0"] - comparacion["SCI_1_0"]

        st.subheader("📋 Comparación por ID_INDICADOR")
        st.dataframe(comparacion, use_container_width=True)

    with st.expander("Ver SCI 1.0"):
        st.dataframe(df1.head(100), use_container_width=True)

    with st.expander("Ver SCI 2.0"):
        st.dataframe(df2.head(100), use_container_width=True)

else:
    st.info("Debes cargar ambos archivos.")
