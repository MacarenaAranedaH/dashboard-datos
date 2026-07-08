"""
Dashboard de análisis de datos con Streamlit
--------------------------------------------
Permite:
  - Cargar datos desde archivos Excel (.xlsx/.xls) o CSV.
  - Manejar hojas de Excel con varias tablas y formatos distintos
    (selección de hoja, fila de encabezado, filas/columnas a omitir,
    y auto-detección de tablas separadas por filas vacías).
  - Guardar y reutilizar los datasets más comunes (persistencia local).
  - Aplicar filtros dinámicos según el tipo de cada columna.
  - Calcular totales, promedios y tasas configurables.
  - Generar distintos gráficos (barras, líneas, dispersión, pastel, etc.).

Ejecutar con:
    streamlit run app.py
"""

import io
import json
import warnings
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

try:
    import plotly.express as px
except ImportError:  # Mensaje claro si falta la dependencia
    px = None


# ---------------------------------------------------------------------------
# Configuración general
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Dashboard de Análisis de Datos",
    page_icon="📊",
    layout="wide",
)

SAVE_DIR = Path("saved_datasets")
SAVE_DIR.mkdir(exist_ok=True)
REGISTRY_PATH = SAVE_DIR / "registry.json"


# ---------------------------------------------------------------------------
# Persistencia de datasets (los "más comunes")
# ---------------------------------------------------------------------------
def load_registry() -> dict:
    """Carga el registro de datasets guardados."""
    if REGISTRY_PATH.exists():
        try:
            return json.loads(REGISTRY_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_registry(registry: dict) -> None:
    REGISTRY_PATH.write_text(
        json.dumps(registry, indent=2, ensure_ascii=False), encoding="utf-8"
    )


def save_dataset(name: str, df: pd.DataFrame) -> None:
    """Guarda un DataFrame en disco (pickle preserva los tipos de datos)."""
    safe_name = "".join(c for c in name if c.isalnum() or c in (" ", "_", "-")).strip()
    if not safe_name:
        raise ValueError("El nombre no es válido.")
    filename = f"{safe_name}.pkl"
    df.to_pickle(SAVE_DIR / filename)

    registry = load_registry()
    registry[safe_name] = {
        "filename": filename,
        "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "rows": int(df.shape[0]),
        "cols": int(df.shape[1]),
    }
    save_registry(registry)


def load_saved_dataset(name: str) -> pd.DataFrame:
    registry = load_registry()
    info = registry.get(name)
    if not info:
        raise FileNotFoundError(f"No existe el dataset '{name}'.")
    return pd.read_pickle(SAVE_DIR / info["filename"])


def delete_saved_dataset(name: str) -> None:
    registry = load_registry()
    info = registry.pop(name, None)
    if info:
        fpath = SAVE_DIR / info["filename"]
        if fpath.exists():
            fpath.unlink()
        save_registry(registry)


# ---------------------------------------------------------------------------
# Lectura de archivos
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def get_excel_sheets(file_bytes: bytes) -> list:
    """Devuelve la lista de hojas de un Excel."""
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    return xls.sheet_names


@st.cache_data(show_spinner=False)
def read_raw_sheet(file_bytes: bytes, sheet_name) -> pd.DataFrame:
    """Lee una hoja completa sin encabezado, para inspección/auto-detección."""
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)


def split_tables_by_blank_rows(raw: pd.DataFrame, min_rows: int = 2) -> list:
    """
    Divide una hoja en bloques (tablas) separados por filas totalmente vacías.
    Devuelve una lista de DataFrames (con índices originales) para que el
    usuario elija con cuál trabajar.
    """
    blocks = []
    current = []
    for idx, row in raw.iterrows():
        if row.isna().all():
            if current:
                blocks.append(raw.loc[current])
                current = []
        else:
            current.append(idx)
    if current:
        blocks.append(raw.loc[current])
    # Filtra bloques muy pequeños (probable ruido)
    return [b for b in blocks if len(b) >= min_rows]


def promote_header(df: pd.DataFrame, header_row_pos: int) -> pd.DataFrame:
    """Usa una fila (por posición) como encabezado y limpia el resto."""
    df = df.reset_index(drop=True)
    header = df.iloc[header_row_pos].tolist()
    body = df.iloc[header_row_pos + 1:].copy()
    body.columns = _dedupe_columns(header)
    body = body.reset_index(drop=True)
    return infer_types(body)


def _dedupe_columns(cols: list) -> list:
    """Evita nombres de columna repetidos o vacíos.

    Convierte de forma segura cualquier tipo (float nan, pd.NA, None, números)
    a un nombre de columna válido.
    """
    seen = {}
    out = []
    for i, c in enumerate(cols):
        try:
            is_na = pd.isna(c)
        except (TypeError, ValueError):
            is_na = False
        if is_na:
            name = f"col_{i}"
        else:
            name = str(c).strip()
            if not name or name.lower() == "nan":
                name = f"col_{i}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        out.append(name)
    return out


def infer_types(df: pd.DataFrame) -> pd.DataFrame:
    """Intenta convertir columnas a numérico o fecha cuando tiene sentido.

    Compatible con pandas >= 2 y con pandas 3.0 (donde las columnas de texto
    pasan a ser de tipo 'str' en lugar de 'object').
    """
    df = df.copy()
    for col in df.columns:
        serie = df[col]
        # Saltar columnas que ya tienen un tipo "final"
        if (
            pd.api.types.is_numeric_dtype(serie)
            or pd.api.types.is_datetime64_any_dtype(serie)
            or pd.api.types.is_bool_dtype(serie)
        ):
            continue

        # Solo evaluamos valores no nulos para decidir el umbral
        no_nulos = serie.dropna()
        if no_nulos.empty:
            continue

        # Numérico
        converted = pd.to_numeric(serie, errors="coerce")
        if converted.loc[no_nulos.index].notna().mean() >= 0.8:
            df[col] = converted
            continue
        # Fecha
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                dt = pd.to_datetime(serie, errors="coerce", dayfirst=True)
            if dt.loc[no_nulos.index].notna().mean() >= 0.8:
                df[col] = dt
        except Exception:
            pass
    return df


# ---------------------------------------------------------------------------
# Filtros dinámicos
# ---------------------------------------------------------------------------
def build_filters(df: pd.DataFrame) -> pd.DataFrame:
    """Renderiza filtros en la barra lateral y devuelve el DataFrame filtrado."""
    st.sidebar.markdown("### 🔍 Filtros")
    if df.empty:
        return df

    cols_to_filter = st.sidebar.multiselect(
        "Columnas a filtrar",
        options=list(df.columns),
        help="Elige las columnas por las que quieres filtrar.",
    )

    filtered = df.copy()
    for col in cols_to_filter:
        serie = df[col]
        st.sidebar.markdown(f"**{col}**")

        if pd.api.types.is_numeric_dtype(serie):
            lo, hi = float(np.nanmin(serie)), float(np.nanmax(serie))
            if lo == hi:
                st.sidebar.caption(f"Valor único: {lo}")
                continue
            rango = st.sidebar.slider(
                f"Rango de {col}", lo, hi, (lo, hi), key=f"num_{col}"
            )
            filtered = filtered[
                filtered[col].between(rango[0], rango[1]) | filtered[col].isna()
            ]

        elif pd.api.types.is_datetime64_any_dtype(serie):
            min_d, max_d = serie.min(), serie.max()
            if pd.isna(min_d) or pd.isna(max_d):
                continue
            rango = st.sidebar.date_input(
                f"Rango de fechas de {col}",
                value=(min_d.date(), max_d.date()),
                key=f"date_{col}",
            )
            if isinstance(rango, tuple) and len(rango) == 2:
                mask = filtered[col].between(
                    pd.to_datetime(rango[0]), pd.to_datetime(rango[1])
                )
                filtered = filtered[mask]

        else:  # Categórico / texto
            opciones = serie.dropna().astype(str).unique().tolist()
            if len(opciones) <= 200:
                sel = st.sidebar.multiselect(
                    f"Valores de {col}",
                    options=sorted(opciones),
                    default=sorted(opciones),
                    key=f"cat_{col}",
                )
                filtered = filtered[filtered[col].astype(str).isin(sel)]
            else:
                texto = st.sidebar.text_input(
                    f"Buscar en {col} (contiene)", key=f"txt_{col}"
                )
                if texto:
                    filtered = filtered[
                        filtered[col].astype(str).str.contains(texto, case=False, na=False)
                    ]

    return filtered


# ---------------------------------------------------------------------------
# Sección: carga de datos (barra lateral)
# ---------------------------------------------------------------------------
def data_source_section():
    st.sidebar.title("📥 Fuente de datos")
    fuente = st.sidebar.radio(
        "Origen", ["Subir archivo", "Datasets guardados"], key="fuente"
    )

    if fuente == "Subir archivo":
        upload_flow()
    else:
        saved_flow()


def upload_flow():
    archivo = st.sidebar.file_uploader(
        "Sube un Excel o CSV",
        type=["xlsx", "xls", "csv"],
        help="Formatos soportados: .xlsx, .xls, .csv",
    )
    if archivo is None:
        return

    file_bytes = archivo.getvalue()
    name = archivo.name.lower()

    if name.endswith(".csv"):
        sep = st.sidebar.selectbox(
            "Separador", [",", ";", "\\t", "|"], index=0, key="csv_sep"
        )
        sep = "\t" if sep == "\\t" else sep
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), sep=sep)
            df = infer_types(df)
            _set_active_df(df, archivo.name)
        except Exception as e:
            st.sidebar.error(f"No se pudo leer el CSV: {e}")
        return

    # Excel
    try:
        sheets = get_excel_sheets(file_bytes)
    except Exception as e:
        st.sidebar.error(f"No se pudo abrir el Excel: {e}")
        return

    hoja = st.sidebar.selectbox("Hoja", sheets, key="hoja_excel")
    modo = st.sidebar.radio(
        "Modo de lectura",
        ["Encabezado en fila específica", "Auto-detectar tablas"],
        key="modo_excel",
    )

    raw = read_raw_sheet(file_bytes, hoja)

    if modo == "Encabezado en fila específica":
        header_row = st.sidebar.number_input(
            "Fila del encabezado (1 = primera fila)",
            min_value=1,
            max_value=max(1, len(raw)),
            value=1,
            key="header_row",
        )
        with st.sidebar.expander("Vista previa de la hoja (sin procesar)"):
            st.dataframe(raw.head(15), use_container_width=True)
        if st.sidebar.button("Cargar tabla", key="load_excel_single"):
            df = promote_header(raw, header_row - 1)
            _set_active_df(df, f"{archivo.name} · {hoja}")

    else:  # Auto-detectar
        bloques = split_tables_by_blank_rows(raw)
        if not bloques:
            st.sidebar.warning("No se detectaron tablas separadas por filas vacías.")
            return
        st.sidebar.caption(f"Se detectaron {len(bloques)} bloque(s).")
        idx = st.sidebar.selectbox(
            "Selecciona el bloque",
            range(len(bloques)),
            format_func=lambda i: f"Tabla {i + 1} ({len(bloques[i])} filas)",
            key="bloque_idx",
        )
        bloque = bloques[idx].reset_index(drop=True)
        with st.sidebar.expander("Vista previa del bloque"):
            st.dataframe(bloque.head(15), use_container_width=True)
        header_row = st.sidebar.number_input(
            "Fila del encabezado dentro del bloque",
            min_value=1,
            max_value=len(bloque),
            value=1,
            key="header_row_block",
        )
        if st.sidebar.button("Cargar bloque", key="load_excel_block"):
            df = promote_header(bloque, header_row - 1)
            _set_active_df(df, f"{archivo.name} · {hoja} · Tabla {idx + 1}")


def saved_flow():
    registry = load_registry()
    if not registry:
        st.sidebar.info("Aún no hay datasets guardados.")
        return

    nombre = st.sidebar.selectbox(
        "Dataset guardado",
        list(registry.keys()),
        format_func=lambda n: f"{n} ({registry[n]['rows']}×{registry[n]['cols']})",
        key="saved_select",
    )
    info = registry[nombre]
    st.sidebar.caption(f"Guardado: {info['saved_at']}")

    c1, c2 = st.sidebar.columns(2)
    if c1.button("Cargar", key="load_saved"):
        try:
            df = load_saved_dataset(nombre)
            _set_active_df(df, f"[guardado] {nombre}")
        except Exception as e:
            st.sidebar.error(f"Error al cargar: {e}")
    if c2.button("Eliminar", key="del_saved"):
        delete_saved_dataset(nombre)
        st.sidebar.success(f"'{nombre}' eliminado.")
        st.rerun()


def _set_active_df(df: pd.DataFrame, origen: str):
    st.session_state["df"] = df
    st.session_state["origen"] = origen


def save_current_section():
    """Permite guardar el dataset activo para reutilizarlo luego."""
    if "df" not in st.session_state:
        return
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 💾 Guardar dataset actual")
    nombre = st.sidebar.text_input("Nombre para guardar", key="save_name")
    if st.sidebar.button("Guardar", key="save_btn"):
        if nombre.strip():
            try:
                save_dataset(nombre.strip(), st.session_state["df"])
                st.sidebar.success(f"Guardado como '{nombre.strip()}'.")
            except Exception as e:
                st.sidebar.error(f"No se pudo guardar: {e}")
        else:
            st.sidebar.warning("Escribe un nombre.")


# ---------------------------------------------------------------------------
# Pestañas principales
# ---------------------------------------------------------------------------
def tab_datos(df: pd.DataFrame, df_filtrado: pd.DataFrame):
    st.subheader("Vista de datos")
    c1, c2, c3 = st.columns(3)
    c1.metric("Filas (total)", f"{len(df):,}")
    c2.metric("Filas (filtradas)", f"{len(df_filtrado):,}")
    c3.metric("Columnas", f"{df.shape[1]:,}")

    # Vista previa acotada: enviar todo el DataFrame al navegador puede superar
    # el límite de tamaño de mensaje de Streamlit (MessageSizeError) con datos grandes.
    total = len(df_filtrado)
    if total > 2000:
        n_prev = st.slider(
            "Filas a mostrar (vista previa)",
            min_value=500,
            max_value=min(20000, total),
            value=2000,
            step=500,
        )
        st.caption(
            f"Mostrando {n_prev:,} de {total:,} filas. Usa los filtros para acotar el análisis."
        )
        st.dataframe(df_filtrado.head(n_prev), use_container_width=True, height=420)
    else:
        st.dataframe(df_filtrado, use_container_width=True, height=420)

    # Descarga bajo demanda: el CSV completo solo se genera al marcar la casilla,
    # evitando codificar todo el dataset en cada recarga.
    if st.checkbox("Preparar descarga del CSV filtrado"):
        csv = df_filtrado.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Descargar datos filtrados (CSV)",
            data=csv,
            file_name="datos_filtrados.csv",
            mime="text/csv",
        )

    with st.expander("Tipos de datos detectados"):
        tipos = pd.DataFrame(
            {"columna": df.columns, "tipo": [str(t) for t in df.dtypes]}
        )
        st.dataframe(tipos, use_container_width=True)


def tab_metricas(df: pd.DataFrame):
    st.subheader("Totales y tasas")
    num_cols = df.select_dtypes(include=np.number).columns.tolist()

    # --- Totales / resúmenes ---
    st.markdown("#### 📌 Totales y resúmenes")
    if num_cols:
        cols_sel = st.multiselect(
            "Columnas numéricas", num_cols, default=num_cols[: min(3, len(num_cols))]
        )
        if cols_sel:
            resumen = df[cols_sel].agg(["sum", "mean", "min", "max", "count"]).T
            resumen.columns = ["Suma", "Promedio", "Mínimo", "Máximo", "Conteo"]
            st.dataframe(resumen.style.format("{:,.2f}"), use_container_width=True)
    else:
        st.info("No hay columnas numéricas para resumir.")

    st.markdown("---")

    # --- Tasas configurables ---
    st.markdown("#### 📐 Cálculo de tasas")
    tipo_tasa = st.radio(
        "Tipo de tasa",
        ["Razón entre dos columnas", "Proporción por condición"],
        horizontal=True,
    )

    if tipo_tasa == "Razón entre dos columnas":
        if len(num_cols) < 2:
            st.info("Necesitas al menos dos columnas numéricas.")
            return
        c1, c2, c3 = st.columns(3)
        numerador = c1.selectbox("Numerador", num_cols, key="num_tasa")
        denominador = c2.selectbox(
            "Denominador", num_cols, index=min(1, len(num_cols) - 1), key="den_tasa"
        )
        factor = c3.number_input(
            "Multiplicar por (ej. 100 para %)", value=100.0, key="factor_tasa"
        )
        suma_num = df[numerador].sum()
        suma_den = df[denominador].sum()
        if suma_den == 0:
            st.warning("El denominador suma 0; no se puede calcular la tasa.")
        else:
            tasa = suma_num / suma_den * factor
            m1, m2, m3 = st.columns(3)
            m1.metric(f"Σ {numerador}", f"{suma_num:,.2f}")
            m2.metric(f"Σ {denominador}", f"{suma_den:,.2f}")
            m3.metric("Tasa resultante", f"{tasa:,.2f}")

    else:  # Proporción por condición
        col = st.selectbox("Columna", df.columns, key="col_prop")
        serie = df[col]
        if pd.api.types.is_numeric_dtype(serie):
            operador = st.selectbox("Condición", [">", ">=", "<", "<=", "==", "!="])
            umbral = st.number_input("Valor", value=float(serie.median()))
            mask = {
                ">": serie > umbral,
                ">=": serie >= umbral,
                "<": serie < umbral,
                "<=": serie <= umbral,
                "==": serie == umbral,
                "!=": serie != umbral,
            }[operador]
        else:
            valores = st.multiselect(
                "Valores que cuentan como 'positivos'",
                sorted(serie.dropna().astype(str).unique().tolist()),
            )
            mask = serie.astype(str).isin(valores)

        total = len(df)
        positivos = int(mask.sum())
        prop = positivos / total * 100 if total else 0
        m1, m2, m3 = st.columns(3)
        m1.metric("Casos positivos", f"{positivos:,}")
        m2.metric("Total", f"{total:,}")
        m3.metric("Proporción", f"{prop:,.2f}%")


def tab_graficos(df: pd.DataFrame):
    st.subheader("Gráficos")
    if px is None:
        st.error("Falta la librería 'plotly'. Instálala con: pip install plotly")
        return
    if df.empty:
        st.info("No hay datos que graficar con los filtros actuales.")
        return

    tipo = st.selectbox(
        "Tipo de gráfico",
        ["Barras", "Líneas", "Dispersión", "Pastel", "Histograma", "Caja (boxplot)", "Área"],
    )
    cols = list(df.columns)
    num_cols = df.select_dtypes(include=np.number).columns.tolist()

    agg_funcs = {"suma": "sum", "promedio": "mean", "conteo": "count", "máximo": "max", "mínimo": "min"}

    fig = None
    try:
        if tipo in ("Barras", "Líneas", "Área"):
            c1, c2, c3 = st.columns(3)
            x = c1.selectbox("Eje X", cols, key="x_axis")
            y = c2.selectbox("Eje Y (numérico)", num_cols or cols, key="y_axis")
            color = c3.selectbox("Color (opcional)", ["(ninguno)"] + cols, key="color_axis")
            agg = st.selectbox("Agregación", list(agg_funcs.keys()), index=0)

            group_cols = [x] + ([color] if color != "(ninguno)" else [])
            data = (
                df.groupby(group_cols, dropna=False)[y]
                .agg(agg_funcs[agg])
                .reset_index()
            )
            color_arg = None if color == "(ninguno)" else color
            if tipo == "Barras":
                fig = px.bar(data, x=x, y=y, color=color_arg, barmode="group")
            elif tipo == "Líneas":
                fig = px.line(data, x=x, y=y, color=color_arg, markers=True)
            else:
                fig = px.area(data, x=x, y=y, color=color_arg)

        elif tipo == "Dispersión":
            c1, c2, c3 = st.columns(3)
            x = c1.selectbox("Eje X", num_cols or cols, key="sx")
            y = c2.selectbox("Eje Y", num_cols or cols, key="sy")
            color = c3.selectbox("Color (opcional)", ["(ninguno)"] + cols, key="sc")
            color_arg = None if color == "(ninguno)" else color
            plot_df = df
            if len(df) > 20000:
                plot_df = df.sample(20000, random_state=0)
                st.caption(
                    f"Muestra de 20.000 de {len(df):,} puntos para no saturar el navegador."
                )
            fig = px.scatter(plot_df, x=x, y=y, color=color_arg, opacity=0.7)

        elif tipo == "Pastel":
            c1, c2 = st.columns(2)
            nombres = c1.selectbox("Categorías", cols, key="pie_names")
            valores = c2.selectbox("Valores (numérico)", num_cols or cols, key="pie_vals")
            data = df.groupby(nombres, dropna=False)[valores].sum().reset_index()
            fig = px.pie(data, names=nombres, values=valores)

        elif tipo == "Histograma":
            c1, c2 = st.columns(2)
            x = c1.selectbox("Columna", cols, key="hist_x")
            bins = c2.slider("Número de bins", 5, 100, 20)
            fig = px.histogram(df, x=x, nbins=bins)

        elif tipo == "Caja (boxplot)":
            c1, c2 = st.columns(2)
            y = c1.selectbox("Valores (numérico)", num_cols or cols, key="box_y")
            x = c2.selectbox("Agrupar por (opcional)", ["(ninguno)"] + cols, key="box_x")
            x_arg = None if x == "(ninguno)" else x
            fig = px.box(df, x=x_arg, y=y)

        if fig is not None:
            fig.update_layout(margin=dict(l=10, r=10, t=30, b=10))
            st.plotly_chart(fig, use_container_width=True)
    except Exception as e:
        st.error(f"No se pudo generar el gráfico con esta configuración: {e}")


# ---------------------------------------------------------------------------
# App principal
# ---------------------------------------------------------------------------
def main():
    st.title("📊 Dashboard de Análisis de Datos")

    data_source_section()
    save_current_section()

    if "df" not in st.session_state:
        st.info(
            "👈 Sube un archivo Excel/CSV o carga un dataset guardado para comenzar."
        )
        st.markdown(
            "**Qué puedes hacer aquí:**\n"
            "- Leer hojas de Excel con varias tablas (auto-detección o encabezado manual).\n"
            "- Filtrar por cualquier columna (numérica, categórica o de fecha).\n"
            "- Calcular totales, promedios y tasas configurables.\n"
            "- Generar gráficos de barras, líneas, dispersión, pastel, etc.\n"
            "- Guardar tus datasets más usados para reutilizarlos."
        )
        return

    df = st.session_state["df"]
    st.caption(f"Origen: {st.session_state.get('origen', '—')}")

    df_filtrado = build_filters(df)

    t1, t2, t3 = st.tabs(["📄 Datos", "📊 Métricas", "📈 Gráficos"])
    with t1:
        tab_datos(df, df_filtrado)
    with t2:
        tab_metricas(df_filtrado)
    with t3:
        tab_graficos(df_filtrado)


if __name__ == "__main__":
    main()
