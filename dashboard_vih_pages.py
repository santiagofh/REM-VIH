from __future__ import annotations

import re
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from dashboard_indicadores_esenciales import render_indicadores_page


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "2025" / "salida"
WORKBOOKS = {
    "REM A05": DATA_DIR / "A05_2025.xlsx",
    "REM A11": DATA_DIR / "A11_2025.xlsx",
    "REM P1":  DATA_DIR / "P1_2025.xlsx",
    "REM P11": DATA_DIR / "P11_2025.xlsx",
}
SOURCE_NOTES = {
    "REM A05": "Acumulado enero a diciembre 2025.",
    "REM A11": "Acumulado enero a diciembre 2025.",
    "REM P1":  "Corte de diciembre 2025.",
    "REM P11": "Corte de diciembre 2025.",
}
EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
CSV_MIME = "text/csv"
PRIMARY_COLUMNS = [
    "Región",
    "Servicio de Salud",
    "Comuna",
    "Establecimiento",
    "Codigo DEIS",
    "Prestación",
    "Detalle de prestación",
    "Mes",
    "Año",
    "Codigo Servicio REM",
    "Codigo Comuna REM",
]
IDENTIFIER_COLUMNS = [
    "Codigo DEIS",
    "Prestación",
    "Mes",
    "Año",
    "Codigo Servicio REM",
    "Codigo Comuna REM",
]

# Jerarquía geográfica/organizacional para desagregación
GEO_HIERARCHY = [
    "Región",
    "Servicio de Salud",
    "Comuna",
    "Establecimiento",
]

# Columnas no-numéricas que siempre van en el GROUP BY al agregar
BASE_GROUP_COLUMNS = [
    "Prestación",
    "Detalle de prestación",
    "Mes",
    "Año",
]

DISAGGREGATION_LEVELS = {
    "Nacional":          [],
    "Región":            ["Región"],
    "Servicio de Salud": ["Región", "Servicio de Salud"],
    "Comuna":            ["Región", "Servicio de Salud", "Comuna"],
    "Establecimiento":   ["Región", "Servicio de Salud", "Comuna", "Establecimiento", "Codigo DEIS"],
    "Sin agregar":       None,   # None = devuelve el df tal cual
}


def render_year_badge() -> None:
    st.markdown(
        '<div class="year-badge">Datos Provisorios</div>',
        unsafe_allow_html=True,
    )


def file_safe_name(text: str) -> str:
    clean = re.sub(r"[^\w\-]+", "_", str(text).strip(), flags=re.UNICODE)
    clean = re.sub(r"_+", "_", clean).strip("_")
    return clean or "datos"


def excel_sheet_name(text: str) -> str:
    clean = re.sub(r"[\[\]\:\*\?\/\\]", "_", str(text).strip())
    return clean[:31] or "Datos"


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def ensure_workbook(rem_name: str) -> Path:
    path = WORKBOOKS[rem_name]
    if not path.exists():
        raise FileNotFoundError(f"No se encontró el archivo de datos: {path}")
    return path


def sorted_unique_values(series: pd.Series) -> list[str]:
    values = [
        normalize_text(value)
        for value in series.dropna().tolist()
        if normalize_text(value)
    ]
    return sorted(set(values))


def default_visible_columns(columns: list[str]) -> list[str]:
    preferred = [column for column in PRIMARY_COLUMNS if column in columns]
    extras = [column for column in columns if column not in preferred]
    return preferred + extras[: min(10, len(extras))]


def build_prestation_labels(df: pd.DataFrame) -> dict[str, str]:
    if "Prestación" not in df.columns:
        return {}

    work = df[["Prestación"]].copy()
    if "Detalle de prestación" in df.columns:
        work["Detalle de prestación"] = df["Detalle de prestación"]
    else:
        work["Detalle de prestación"] = ""

    work = work.fillna("").drop_duplicates().sort_values(
        ["Detalle de prestación", "Prestación"], kind="stable"
    )

    labels: dict[str, str] = {}
    for _, row in work.iterrows():
        code = normalize_text(row["Prestación"])
        detail = normalize_text(row["Detalle de prestación"])
        label = f"{code} | {detail}" if detail else code
        labels[label] = code
    return labels


def aggregate_dataframe(df: pd.DataFrame, level: str) -> pd.DataFrame:
    """
    Agrupa df según el nivel de desagregación elegido y suma columnas numéricas.
    Si level == 'Sin agregar' devuelve el df intacto.
    """
    geo_cols = DISAGGREGATION_LEVELS.get(level)
    if geo_cols is None:
        return df

    # Columnas de agrupación: geo del nivel + base (prestación, mes, año…)
    group_cols = [
        col for col in geo_cols + BASE_GROUP_COLUMNS
        if col in df.columns
    ]

    if not group_cols:
        return df

    # Columnas numéricas que se suman (excluir las de agrupación y códigos)
    exclude = set(group_cols) | {"Codigo Servicio REM", "Codigo Comuna REM"}
    num_cols = [
        col for col in df.select_dtypes(include="number").columns
        if col not in exclude
    ]

    if not num_cols:
        return df.groupby(group_cols, sort=True, as_index=False).first()

    agg_dict = {col: "sum" for col in num_cols}
    result = df.groupby(group_cols, sort=True, as_index=False).agg(agg_dict)
    return result.reset_index(drop=True)


@st.cache_data(show_spinner=False)
def load_sheet_names(rem_name: str) -> list[str]:
    path = ensure_workbook(rem_name)
    return pd.ExcelFile(path).sheet_names


@st.cache_data(show_spinner=False)
def load_sheet(rem_name: str, sheet_name: str) -> pd.DataFrame:
    path = ensure_workbook(rem_name)
    df = pd.read_excel(path, sheet_name=sheet_name)
    df.columns = [str(column).strip() for column in df.columns]

    for column in IDENTIFIER_COLUMNS:
        if column in df.columns:
            df[column] = df[column].map(normalize_text)

    for column in ["Región", "Servicio de Salud", "Comuna", "Establecimiento", "Detalle de prestación"]:
        if column in df.columns:
            df[column] = df[column].fillna("").map(str).str.strip()

    return df


@st.cache_data(show_spinner=False)
def load_workbook_bytes(rem_name: str) -> bytes:
    path = ensure_workbook(rem_name)
    return path.read_bytes()


@st.cache_data(show_spinner=False)
def build_data_catalog() -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for rem_name, path in WORKBOOKS.items():
        ensure_workbook(rem_name)
        workbook = pd.ExcelFile(path)
        for sheet_name in workbook.sheet_names:
            df = pd.read_excel(path, sheet_name=sheet_name)
            prestaciones = (
                int(df["Prestación"].nunique(dropna=True))
                if "Prestación" in df.columns
                else pd.NA
            )
            rows.append(
                {
                    "REM": rem_name,
                    "Archivo": path.name,
                    "Sección": sheet_name.replace("Seccion", "Sección"),
                    "Filas": int(len(df)),
                    "Columnas": int(len(df.columns)),
                    "Prestaciones": prestaciones,
                    "Corte": SOURCE_NOTES.get(rem_name, ""),
                }
            )
    return pd.DataFrame(rows)


def filter_dataframe(
    df: pd.DataFrame,
    selected_services: list[str],
    selected_communes: list[str],
    selected_establishments: list[str],
    selected_prestations: list[str],
    search_text: str,
) -> pd.DataFrame:
    filtered = df.copy()

    if selected_services and "Servicio de Salud" in filtered.columns:
        filtered = filtered[filtered["Servicio de Salud"].isin(selected_services)]
    if selected_communes and "Comuna" in filtered.columns:
        filtered = filtered[filtered["Comuna"].isin(selected_communes)]
    if selected_establishments and "Establecimiento" in filtered.columns:
        filtered = filtered[filtered["Establecimiento"].isin(selected_establishments)]
    if selected_prestations and "Prestación" in filtered.columns:
        filtered = filtered[filtered["Prestación"].isin(selected_prestations)]

    text = search_text.strip()
    if text:
        search_columns = [
            column
            for column in [
                "Prestación",
                "Detalle de prestación",
                "Establecimiento",
                "Comuna",
            ]
            if column in filtered.columns
        ]
        if search_columns:
            mask = pd.Series(False, index=filtered.index)
            for column in search_columns:
                mask = mask | filtered[column].astype(str).str.contains(
                    text,
                    case=False,
                    na=False,
                    regex=False,
                )
            filtered = filtered[mask]

    return filtered.reset_index(drop=True)


def dataframe_to_excel_bytes(
    df: pd.DataFrame,
    data_sheet_name: str,
    extra_sheets: dict[str, pd.DataFrame] | None = None,
) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        if extra_sheets:
            for sheet_name, extra_df in extra_sheets.items():
                extra_df.to_excel(
                    writer,
                    sheet_name=excel_sheet_name(sheet_name),
                    index=False,
                )
        df.to_excel(
            writer,
            sheet_name=excel_sheet_name(data_sheet_name),
            index=False,
        )
    return buffer.getvalue()


def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def build_export_context(
    rem_name: str,
    section_name: str,
    total_rows: int,
    filtered_rows: int,
    aggregated_rows: int,
    disagg_level: str,
    selected_services: list[str],
    selected_communes: list[str],
    selected_establishments: list[str],
    selected_prestations: list[str],
    search_text: str,
) -> pd.DataFrame:
    rows = [
        {"Campo": "REM", "Valor": rem_name},
        {"Campo": "Sección", "Valor": section_name.replace("Seccion", "Sección")},
        {"Campo": "Corte", "Valor": SOURCE_NOTES.get(rem_name, "")},
        {"Campo": "Nivel de desagregación", "Valor": disagg_level},
        {"Campo": "Filas totales en sección", "Valor": total_rows},
        {"Campo": "Filas filtradas", "Valor": filtered_rows},
        {"Campo": "Filas exportadas (agregadas)", "Valor": aggregated_rows},
        {"Campo": "Servicio de Salud", "Valor": ", ".join(selected_services) or "Todos"},
        {"Campo": "Comuna", "Valor": ", ".join(selected_communes) or "Todas"},
        {"Campo": "Establecimiento", "Valor": ", ".join(selected_establishments) or "Todos"},
        {"Campo": "Prestación", "Valor": ", ".join(selected_prestations) or "Todas"},
        {"Campo": "Búsqueda libre", "Valor": search_text.strip() or "Sin búsqueda"},
    ]
    return pd.DataFrame(rows)


def render_home_page() -> None:
    build_data_catalog()

    st.title("Dashboard REM VIH 2025")
    render_year_badge()

    st.markdown("### Instrucciones")
    st.markdown(
        """
1. Entra a **Explorador** desde el menú lateral.
2. Selecciona la serie REM y la sección que quieras revisar.
3. Aplica filtros por servicio, comuna, establecimiento o prestación.
4. Elige el **nivel de desagregación** para ver totales por región, servicio, comuna o establecimiento.
5. Descarga la tabla filtrada en Excel o CSV cuando la dejes lista.
        """
    )


def render_explorer_page() -> None:
    st.title("Explorador de tablas")
    render_year_badge()
    st.caption("Filtra y revisa las tablas REM VIH 2025 sin alterar la estructura de los datos.")

    rem_col, section_col = st.columns([1, 1.2])
    with rem_col:
        rem_name = st.selectbox("Serie REM", list(WORKBOOKS.keys()), index=0)
    with section_col:
        section_name = st.selectbox("Sección", load_sheet_names(rem_name), index=0)

    df = load_sheet(rem_name, section_name)

    st.markdown(
        f"""
        <div class="info-card">
            <div class="info-card-title">{section_name.replace("Seccion", "Sección")}</div>
            <p class="info-card-copy">{SOURCE_NOTES.get(rem_name, "")} Usa los filtros para acotar la tabla y luego descarga el resultado.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.expander("Filtros de consulta", expanded=True):
        filter_df = df.copy()

        service_options = (
            sorted_unique_values(filter_df["Servicio de Salud"])
            if "Servicio de Salud" in filter_df.columns
            else []
        )
        selected_services = st.multiselect("Servicio de Salud", service_options)
        if selected_services and "Servicio de Salud" in filter_df.columns:
            filter_df = filter_df[filter_df["Servicio de Salud"].isin(selected_services)]

        commune_options = (
            sorted_unique_values(filter_df["Comuna"])
            if "Comuna" in filter_df.columns
            else []
        )
        selected_communes = st.multiselect("Comuna", commune_options)
        if selected_communes and "Comuna" in filter_df.columns:
            filter_df = filter_df[filter_df["Comuna"].isin(selected_communes)]

        establishment_options = (
            sorted_unique_values(filter_df["Establecimiento"])
            if "Establecimiento" in filter_df.columns
            else []
        )
        selected_establishments = st.multiselect("Establecimiento", establishment_options)
        if selected_establishments and "Establecimiento" in filter_df.columns:
            filter_df = filter_df[filter_df["Establecimiento"].isin(selected_establishments)]

        prestation_labels = build_prestation_labels(filter_df)
        selected_prestation_labels = st.multiselect(
            "Prestación",
            list(prestation_labels.keys()),
        )
        selected_prestations = [
            prestation_labels[label] for label in selected_prestation_labels
        ]

        search_text = st.text_input(
            "Búsqueda libre",
            placeholder="Busca por código, detalle, comuna o establecimiento",
        )

        show_all_columns = st.toggle("Mostrar todas las columnas", value=True)
        selected_columns = list(df.columns)
        if not show_all_columns:
            selected_columns = st.multiselect(
                "Columnas visibles",
                options=list(df.columns),
                default=default_visible_columns(list(df.columns)),
            )
            if not selected_columns:
                selected_columns = default_visible_columns(list(df.columns))

    # ── Nivel de desagregación ────────────────────────────────────────────
    st.markdown("### Nivel de desagregación")
    st.caption(
        "Elige hasta qué nivel quieres ver los totales. "
        "Las columnas numéricas se suman dentro de cada grupo. "
        "'Sin agregar' muestra la tabla original fila a fila."
    )
    disagg_level = st.radio(
        "Nivel",
        options=list(DISAGGREGATION_LEVELS.keys()),
        index=len(DISAGGREGATION_LEVELS) - 1,  # Por defecto: Sin agregar
        horizontal=True,
        label_visibility="collapsed",
    )

    # Aplicar filtros primero, luego agregar
    filtered_df = filter_dataframe(
        df,
        selected_services=selected_services,
        selected_communes=selected_communes,
        selected_establishments=selected_establishments,
        selected_prestations=selected_prestations,
        search_text=search_text,
    )

    aggregated_df = aggregate_dataframe(filtered_df, disagg_level)

    # Columnas visibles (sobre el df ya agregado)
    available_cols = [c for c in selected_columns if c in aggregated_df.columns]
    display_df = aggregated_df[available_cols].copy()
    display_df = display_df.where(pd.notna(display_df), "")

    st.markdown("### Tabla")
    st.caption(
        f"Mostrando **{len(aggregated_df):,} filas** "
        f"(de {len(filtered_df):,} registros filtrados, {len(df):,} totales) "
        f"— nivel: **{disagg_level}**."
    )
    st.dataframe(
        display_df,
        width="stretch",
        hide_index=True,
        height=620,
    )

    context_df = build_export_context(
        rem_name=rem_name,
        section_name=section_name,
        total_rows=len(df),
        filtered_rows=len(filtered_df),
        aggregated_rows=len(aggregated_df),
        disagg_level=disagg_level,
        selected_services=selected_services,
        selected_communes=selected_communes,
        selected_establishments=selected_establishments,
        selected_prestations=selected_prestation_labels,
        search_text=search_text,
    )

    filtered_excel = dataframe_to_excel_bytes(
        aggregated_df,
        data_sheet_name=section_name,
        extra_sheets={"Contexto": context_df},
    )
    filtered_csv = dataframe_to_csv_bytes(aggregated_df)
    section_excel = dataframe_to_excel_bytes(df, data_sheet_name=section_name)

    st.markdown("### Descarga")
    button_col1, button_col2, button_col3 = st.columns(3)
    with button_col1:
        st.download_button(
            "Descargar filtrado en Excel",
            data=filtered_excel,
            file_name=f"{file_safe_name(rem_name)}_{file_safe_name(section_name)}_filtrado.xlsx",
            mime=EXCEL_MIME,
            use_container_width=True,
        )
    with button_col2:
        st.download_button(
            "Descargar filtrado en CSV",
            data=filtered_csv,
            file_name=f"{file_safe_name(rem_name)}_{file_safe_name(section_name)}_filtrado.csv",
            mime=CSV_MIME,
            use_container_width=True,
        )
    with button_col3:
        st.download_button(
            "Descargar sección completa",
            data=section_excel,
            file_name=f"{file_safe_name(rem_name)}_{file_safe_name(section_name)}.xlsx",
            mime=EXCEL_MIME,
            use_container_width=True,
        )

    if aggregated_df.empty:
        st.warning("Los filtros actuales no devuelven registros. Puedes ajustar la selección y volver a descargar.")


def render_downloads_page() -> None:
    catalog = build_data_catalog()

    st.title("Descargas")
    render_year_badge()
    st.caption("Acceso rápido a los archivos fuente y exportaciones por sección.")

    rem_name = st.selectbox("Serie REM para descargar", list(WORKBOOKS.keys()), index=0, key="download-rem")
    section_name = st.selectbox(
        "Sección para exportar",
        load_sheet_names(rem_name),
        index=0,
        key="download-section",
    )

    section_df = load_sheet(rem_name, section_name)
    section_excel = dataframe_to_excel_bytes(section_df, section_name)
    section_csv = dataframe_to_csv_bytes(section_df)
    catalog_excel = dataframe_to_excel_bytes(catalog, "Catalogo")

    top_left, top_right = st.columns(2)
    with top_left:
        st.download_button(
            "Descargar archivo REM completo",
            data=load_workbook_bytes(rem_name),
            file_name=ensure_workbook(rem_name).name,
            mime=EXCEL_MIME,
            use_container_width=True,
        )
    with top_right:
        st.download_button(
            "Descargar catálogo de secciones",
            data=catalog_excel,
            file_name="REM_VIH_2025_catalogo.xlsx",
            mime=EXCEL_MIME,
            use_container_width=True,
        )

    bottom_left, bottom_right = st.columns(2)
    with bottom_left:
        st.download_button(
            "Descargar sección en Excel",
            data=section_excel,
            file_name=f"{file_safe_name(rem_name)}_{file_safe_name(section_name)}.xlsx",
            mime=EXCEL_MIME,
            use_container_width=True,
        )
    with bottom_right:
        st.download_button(
            "Descargar sección en CSV",
            data=section_csv,
            file_name=f"{file_safe_name(rem_name)}_{file_safe_name(section_name)}.csv",
            mime=CSV_MIME,
            use_container_width=True,
        )

    st.markdown("### Vista rápida")
    st.dataframe(
        section_df.head(100).where(pd.notna(section_df.head(100)), ""),
        width="stretch",
        hide_index=True,
        height=420,
    )
    st.caption("La vista rápida muestra las primeras 100 filas de la sección seleccionada.")


def get_navigation_pages():
    return [
        st.Page(render_home_page, title="Inicio", icon=":material/home:", default=True),
        st.Page(
            render_explorer_page,
            title="Explorador",
            icon=":material/table_view:",
            url_path="explorador",
        ),
        st.Page(
            render_indicadores_page,
            title="Indicadores",
            icon=":material/bar_chart:",
            url_path="indicadores",
        ),
        st.Page(
            render_downloads_page,
            title="Descargas",
            icon=":material/download:",
            url_path="descargas",
        ),
    ]