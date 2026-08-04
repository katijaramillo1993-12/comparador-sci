import io
import streamlit as st
import pandas as pd
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

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
    "TAMANIO_DE_MUESTRA", "ID_ESCALA", "FECHA_GENERACION", "USUARIO"
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


def to_decimal(value):
    """Convierte un valor numérico con coma o punto a Decimal."""
    if pd.isna(value) or str(value).strip() == "":
        return None

    try:
        normalized = str(value).strip().replace(",", ".")
        return Decimal(normalized)
    except (InvalidOperation, ValueError, TypeError):
        return None


def round_decimal(value, decimal_places):
    """Redondea usando ROUND_HALF_UP a la cantidad de decimales seleccionada."""
    number = to_decimal(value)

    if number is None:
        return None

    quantizer = (
        Decimal("1")
        if decimal_places == 0
        else Decimal("1." + ("0" * decimal_places))
    )

    return number.quantize(quantizer, rounding=ROUND_HALF_UP)


def values_are_equal(
    value_1,
    value_2,
    comparison_method,
    decimal_places,
    tolerance,
):
    """
    Compara los valores según el método seleccionado:

    - Solo redondeo por cantidad de decimales.
    - Solo tolerancia absoluta.
    - Redondeo + tolerancia.
    """
    number_1 = to_decimal(value_1)
    number_2 = to_decimal(value_2)

    if number_1 is None and number_2 is None:
        return normalize_value(value_1) == normalize_value(value_2)

    if number_1 is None or number_2 is None:
        return False

    tolerance_decimal = Decimal(str(tolerance))

    if comparison_method == "Solo redondeo por cantidad de decimales":
        rounded_1 = round_decimal(value_1, decimal_places)
        rounded_2 = round_decimal(value_2, decimal_places)
        return rounded_1 == rounded_2

    if comparison_method == "Solo tolerancia absoluta":
        return abs(number_1 - number_2) <= tolerance_decimal

    # Método combinado:
    # primero se redondean ambos valores y luego se aplica la tolerancia.
    rounded_1 = round_decimal(value_1, decimal_places)
    rounded_2 = round_decimal(value_2, decimal_places)

    return abs(rounded_1 - rounded_2) <= tolerance_decimal

def get_difference(value_1, value_2):
    """Calcula la diferencia sobre los valores originales, sin redondear."""
    num_1 = to_decimal(value_1)
    num_2 = to_decimal(value_2)

    if num_1 is None or num_2 is None:
        return "N/A", "N/A"

    absolute_difference = abs(num_1 - num_2)

    if num_1 == 0:
        percentage_difference = "N/A"
    else:
        percentage_difference = float(
            (absolute_difference / abs(num_1)) * Decimal("100")
        )

    return float(absolute_difference), percentage_difference


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
def compare_files(df1, df2, comparison_method, decimal_places, tolerance):
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
            if values_are_equal(
                original_value_1,
                original_value_2,
                comparison_method,
                decimal_places,
                tolerance,
            ):
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
                "VALOR_FOXPRO_ORIGINAL": original_value_1,
                "VALOR_SCI_2_0_ORIGINAL": original_value_2,
                "VALOR_FOXPRO_REDONDEADO": str(
                    round_decimal(original_value_1, decimal_places)
                ),
                "VALOR_SCI_2_0_REDONDEADO": str(
                    round_decimal(original_value_2, decimal_places)
                ),
                "METODO_COMPARACION": comparison_method,
                "DECIMALES_CONSIDERADOS": (
                    decimal_places
                    if comparison_method in [
                        "Solo redondeo por cantidad de decimales",
                        "Redondeo + tolerancia",
                    ]
                    else "No aplica"
                ),
                "TOLERANCIA_APLICADA": (
                    tolerance
                    if comparison_method in [
                        "Solo tolerancia absoluta",
                        "Redondeo + tolerancia",
                    ]
                    else "No aplica"
                ),
                "DIFERENCIA_ABSOLUTA_ORIGINAL": absolute_difference,
                "DIFERENCIA_PORCENTUAL_ORIGINAL": percentage_difference,
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

    # =====================================================
    # Resumen consolidado por ID_INDICADOR
    # =====================================================
    consolidated_rows = []

    for indicator in sorted(VALUE_CONFIG.keys()):
        indicator_summary = matched_value_summary_df[
            matched_value_summary_df["ID_INDICADOR"] == indicator
        ].copy()

        compared_records = len(indicator_summary)

        # PROPORCION aplica para los indicadores 3, 4, 5, 6 y 7.
        if "PROPORCION" in VALUE_CONFIG.get(indicator, []):
            proportion_differences = int(
                (indicator_summary["DIFERENCIA_PROPORCION"] == "SI").sum()
            )
            proportion_matches = compared_records - proportion_differences
            proportion_match_pct = (
                (proportion_matches / compared_records) * 100
                if compared_records > 0 else 0
            )
            proportion_diff_pct = (
                (proportion_differences / compared_records) * 100
                if compared_records > 0 else 0
            )
        else:
            proportion_matches = "N/A"
            proportion_match_pct = "N/A"
            proportion_differences = "N/A"
            proportion_diff_pct = "N/A"

        # TAMANIO_DE_MUESTRA solo aplica para 4, 5 y 6.
        if "TAMANIO_DE_MUESTRA" in VALUE_CONFIG.get(indicator, []):
            sample_differences = int(
                (indicator_summary["DIFERENCIA_TAMANIO_DE_MUESTRA"] == "SI").sum()
            )
            sample_matches = compared_records - sample_differences
            sample_match_pct = (
                (sample_matches / compared_records) * 100
                if compared_records > 0 else 0
            )
            sample_diff_pct = (
                (sample_differences / compared_records) * 100
                if compared_records > 0 else 0
            )
        else:
            sample_matches = "N/A"
            sample_match_pct = "N/A"
            sample_differences = "N/A"
            sample_diff_pct = "N/A"

        consolidated_rows.append({
            "ID_INDICADOR": indicator,
            "REGISTROS_COMPARADOS": compared_records,
            "PROPORCION_COINCIDENCIAS": proportion_matches,
            "PROPORCION_%_COINCIDENCIA": proportion_match_pct,
            "PROPORCION_DIFERENCIAS": proportion_differences,
            "PROPORCION_%_DIFERENCIA": proportion_diff_pct,
            "TAMANIO_MUESTRA_COINCIDENCIAS": sample_matches,
            "TAMANIO_MUESTRA_%_COINCIDENCIA": sample_match_pct,
            "TAMANIO_MUESTRA_DIFERENCIAS": sample_differences,
            "TAMANIO_MUESTRA_%_DIFERENCIA": sample_diff_pct,
        })

    # Totales de PROPORCION sobre todas las claves con match.
    total_compared = len(matched_value_summary_df)
    total_proportion_differences = int(
        (matched_value_summary_df["DIFERENCIA_PROPORCION"] == "SI").sum()
    )
    total_proportion_matches = total_compared - total_proportion_differences

    # Totales de TAMANIO_DE_MUESTRA solo sobre indicadores 4, 5 y 6.
    sample_applicable = matched_value_summary_df[
        matched_value_summary_df["ID_INDICADOR"].isin([4, 5, 6])
    ]
    total_sample_compared = len(sample_applicable)
    total_sample_differences = int(
        (sample_applicable["DIFERENCIA_TAMANIO_DE_MUESTRA"] == "SI").sum()
    )
    total_sample_matches = total_sample_compared - total_sample_differences

    consolidated_rows.append({
        "ID_INDICADOR": "TOTAL",
        "REGISTROS_COMPARADOS": total_compared,
        "PROPORCION_COINCIDENCIAS": total_proportion_matches,
        "PROPORCION_%_COINCIDENCIA": (
            (total_proportion_matches / total_compared) * 100
            if total_compared > 0 else 0
        ),
        "PROPORCION_DIFERENCIAS": total_proportion_differences,
        "PROPORCION_%_DIFERENCIA": (
            (total_proportion_differences / total_compared) * 100
            if total_compared > 0 else 0
        ),
        "TAMANIO_MUESTRA_COINCIDENCIAS": total_sample_matches,
        "TAMANIO_MUESTRA_%_COINCIDENCIA": (
            (total_sample_matches / total_sample_compared) * 100
            if total_sample_compared > 0 else 0
        ),
        "TAMANIO_MUESTRA_DIFERENCIAS": total_sample_differences,
        "TAMANIO_MUESTRA_%_DIFERENCIA": (
            (total_sample_differences / total_sample_compared) * 100
            if total_sample_compared > 0 else 0
        ),
    })

    consolidated_by_indicator = pd.DataFrame(consolidated_rows)

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
    comparison_configuration = pd.DataFrame([
        {
            "PARAMETRO": "Método de comparación",
            "VALOR": comparison_method
        },
        {
            "PARAMETRO": "Cantidad de decimales",
            "VALOR": (
                decimal_places
                if comparison_method in [
                    "Solo redondeo por cantidad de decimales",
                    "Redondeo + tolerancia",
                ]
                else "No aplica"
            )
        },
        {
            "PARAMETRO": "Tolerancia absoluta",
            "VALOR": (
                tolerance
                if comparison_method in [
                    "Solo tolerancia absoluta",
                    "Redondeo + tolerancia",
                ]
                else "No aplica"
            )
        },
        {
            "PARAMETRO": "Criterio aplicado",
            "VALOR": (
                "Los valores se redondean y deben ser exactamente iguales."
                if comparison_method == "Solo redondeo por cantidad de decimales"
                else (
                    "Se comparan los valores originales aplicando tolerancia absoluta."
                    if comparison_method == "Solo tolerancia absoluta"
                    else (
                        "Los valores se redondean primero y luego se aplica "
                        "la tolerancia absoluta sobre los valores redondeados."
                    )
                )
            )
        }
    ])

    return {
        "Configuracion_comparacion": comparison_configuration,
        "Resumen": summary,
        "Resumen_por_ID_INDICADOR": consolidated_by_indicator,
        "Resumen_PROPORCION": build_summary_tables(consolidated_by_indicator)[0],
        "Resumen_TAM_MUESTRA": build_summary_tables(consolidated_by_indicator)[1],
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

def format_percentage(value):
    """Da formato de porcentaje con coma decimal."""
    if value == "N/A" or pd.isna(value):
        return "N/A"

    return f"{float(value):.2f}%".replace(".", ",")


def build_summary_tables(consolidated_df):
    """
    Separa el resumen consolidado en dos tablas:
    1. PROPORCION
    2. TAMANIO_DE_MUESTRA
    """
    detail_df = consolidated_df[
        consolidated_df["ID_INDICADOR"].astype(str) != "TOTAL"
    ].copy()

    proportion_df = pd.DataFrame({
        "ID_INDICADOR": detail_df["ID_INDICADOR"],
        "Registros comparados": detail_df["REGISTROS_COMPARADOS"],
        "Coincidencias": detail_df["PROPORCION_COINCIDENCIAS"],
        "% Coincidencia": detail_df["PROPORCION_%_COINCIDENCIA"].apply(
            format_percentage
        ),
        "Diferencias": detail_df["PROPORCION_DIFERENCIAS"],
        "% Diferencia": detail_df["PROPORCION_%_DIFERENCIA"].apply(
            format_percentage
        ),
    })

    sample_source = detail_df[
        detail_df["ID_INDICADOR"].isin([4, 5, 6])
    ].copy()

    sample_df = pd.DataFrame({
        "ID_INDICADOR": sample_source["ID_INDICADOR"],
        "Registros comparados": sample_source["REGISTROS_COMPARADOS"],
        "Coincidencias": sample_source["TAMANIO_MUESTRA_COINCIDENCIAS"],
        "% Coincidencia": sample_source[
            "TAMANIO_MUESTRA_%_COINCIDENCIA"
        ].apply(format_percentage),
        "Diferencias": sample_source["TAMANIO_MUESTRA_DIFERENCIAS"],
        "% Diferencia": sample_source[
            "TAMANIO_MUESTRA_%_DIFERENCIA"
        ].apply(format_percentage),
    })

    return proportion_df, sample_df


def render_summary_tables(consolidated_df):
    """
    Muestra las tablas separadas como en el formato solicitado.
    """
    proportion_df, sample_df = build_summary_tables(consolidated_df)

    st.markdown("### PROPORCIÓN")
    st.dataframe(
        proportion_df,
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("### TAMAÑO_DE_MUESTRA")
    st.dataframe(
        sample_df,
        use_container_width=True,
        hide_index=True,
    )


# =========================================================
with st.sidebar:
    st.header("Configuración")
    modo = st.radio("Modo de trabajo", ["Validar archivo SCI 2.0", "Comparar SCI 1.0 vs SCI 2.0"])
    comparison_method = st.selectbox(
        "Método de comparación para el punto 4",
        [
            "Solo redondeo por cantidad de decimales",
            "Solo tolerancia absoluta",
            "Redondeo + tolerancia",
        ],
        index=2,
        help=(
            "Puedes aplicar solo redondeo, solo tolerancia o ambos criterios."
        )
    )

    decimal_places = st.number_input(
        "Cantidad de decimales para comparar valores",
        min_value=0,
        max_value=10,
        value=1,
        step=1,
        disabled=comparison_method == "Solo tolerancia absoluta",
        help=(
            "Los valores de PROPORCION y TAMANIO_DE_MUESTRA se redondean "
            "a esta cantidad de decimales."
        )
    )

    tolerance = st.number_input(
        "Nivel de tolerancia absoluta",
        min_value=0.0,
        value=0.1,
        step=0.01,
        format="%.6f",
        disabled=comparison_method == "Solo redondeo por cantidad de decimales",
        help=(
            "Después del redondeo, o directamente sobre los valores originales, "
            "se consideran coincidentes cuando la diferencia absoluta es menor "
            "o igual a esta tolerancia."
        )
    )

    with st.expander("Ver ejemplo del criterio combinado"):
        example_df = pd.DataFrame([
            {
                "FoxPro": 0.142,
                "SCI 2.0": 0.149,
                "Decimales": 1,
                "Resultado del redondeo": "0.1 vs 0.1",
                "Tolerancia": 0.1,
                "¿Coinciden?": "✅ Sí",
            },
            {
                "FoxPro": 0.142,
                "SCI 2.0": 0.151,
                "Decimales": 1,
                "Resultado del redondeo": "0.1 vs 0.2",
                "Tolerancia": 0.1,
                "¿Coinciden?": "✅ Sí (diferencia=0.1)",
            },
            {
                "FoxPro": 0.142,
                "SCI 2.0": 0.151,
                "Decimales": 1,
                "Resultado del redondeo": "0.1 vs 0.2",
                "Tolerancia": 0.05,
                "¿Coinciden?": "❌ No",
            },
        ])
        st.dataframe(example_df, use_container_width=True, hide_index=True)

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
            results = compare_files(
                df1,
                df2,
                comparison_method,
                decimal_places,
                tolerance,
            )
        summary = results["Resumen"]
        summary_dict = dict(zip(summary["VALIDACION"], summary["RESULTADO"]))
        st.success("Comparación ejecutada correctamente")
        if comparison_method == "Solo redondeo por cantidad de decimales":
            st.info(
                f"**Criterio del punto 4:** los valores de PROPORCION y "
                f"TAMANIO_DE_MUESTRA se redondearon a **{decimal_places} "
                f"decimal(es)** y se consideraron coincidentes cuando los "
                f"valores redondeados fueron exactamente iguales."
            )
        elif comparison_method == "Solo tolerancia absoluta":
            st.info(
                f"**Criterio del punto 4:** se compararon los valores originales "
                f"utilizando una tolerancia absoluta de **±{tolerance:.6f}**."
            )
        else:
            st.info(
                f"**Criterio del punto 4:** primero se redondearon los valores "
                f"de FoxPro y SCI 2.0 a **{decimal_places} decimal(es)** y luego "
                f"se aplicó una tolerancia absoluta de **±{tolerance:.6f}** "
                f"sobre los valores redondeados."
            )
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
        if comparison_method == "Solo redondeo por cantidad de decimales":
            st.caption(
                f"Nota: 'Valores idénticos' considera las claves que hicieron "
                f"match y cuyos valores resultaron iguales después de redondear "
                f"ambos sistemas a {decimal_places} decimal(es)."
            )
        elif comparison_method == "Solo tolerancia absoluta":
            st.caption(
                f"Nota: 'Valores idénticos' considera las claves que hicieron "
                f"match y cuya diferencia absoluta original fue menor o igual "
                f"a {tolerance:.6f}."
            )
        else:
            st.caption(
                f"Nota: 'Valores idénticos' considera las claves que hicieron "
                f"match, cuyos valores fueron redondeados a {decimal_places} "
                f"decimal(es) y cuya diferencia absoluta posterior fue menor "
                f"o igual a {tolerance:.6f}."
            )
        excel_report = create_excel(results)
        st.download_button("📥 Descargar reporte comparativo Excel", data=excel_report, file_name="Reporte_Comparativo_SCI_Automatico.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        tabs = st.tabs([
            "Resumen",
            "Resumen por ID_INDICADOR",
            "Estructura",
            "Conteo ID_INDICADOR",
            "Match por clave",
            "Resumen valores match",
            "Solo en un archivo",
            "Diferencias valores",
            "Duplicados / inválidos"
        ])
        with tabs[0]:
            st.subheader("Resumen general")
            safe_display(results["Resumen"])

        with tabs[1]:
            st.subheader("Resumen consolidado por ID_INDICADOR")
            st.write(
                "Las siguientes tablas presentan, por indicador, los registros "
                "comparados, las coincidencias, las diferencias y sus porcentajes "
                "para las claves que hicieron match."
            )

            render_summary_tables(results["Resumen_por_ID_INDICADOR"])

            st.caption(
                "TAMANIO_DE_MUESTRA solo aplica para los indicadores 4, 5 y 6. "
                "Por esta razón, los indicadores 3 y 7 no se muestran en esa tabla."
            )

        with tabs[2]:
            st.subheader("Validación de estructura")
            safe_display(results["Estructura"])

        with tabs[3]:
            st.subheader("Conteo por ID_INDICADOR")
            safe_display(results["Conteo_ID_INDICADOR"])

        with tabs[4]:
            st.subheader("Registros que coinciden por clave")
            safe_display(results["Match_por_clave"])

        with tabs[5]:
            st.subheader("Resumen de valores numéricos para claves con match")
            chart_df = results["Grafico_resumen_valores"]
            if chart_df.empty:
                st.info("Sin datos para graficar.")
            else:
                st.bar_chart(chart_df.set_index("CATEGORIA")["CANTIDAD"])
            safe_display(results["Resumen_valores_match"])
        with tabs[6]:
            st.subheader("Registros solo en SCI 1.0")
            safe_display(results["Solo_SCI_1_0"])
            st.subheader("Registros solo en SCI 2.0")
            safe_display(results["Solo_SCI_2_0"])
        with tabs[7]:
            st.subheader("Diferencias en PROPORCION / TAMANIO_DE_MUESTRA")
            if results["Diferencias_valores"].empty:
                st.success("No se detectaron diferencias en PROPORCION / TAMANIO_DE_MUESTRA.")
            else:
                safe_display(results["Diferencias_valores"])
        with tabs[8]:
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
