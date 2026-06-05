from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st




def render_year_badge() -> None:
    st.markdown(
        '<div class="year-badge">Datos Provisorios</div>',
        unsafe_allow_html=True,
    )


BASE_DIR = Path(__file__).resolve().parent

ANOS_DISPONIBLES = ["2025"]

GEO_COLS = ["Servicio de Salud", "Comuna", "Establecimiento"]

SIFILIS_GESTANTE_CODES = {
    "Seccion A.1": [11053700, 11053800, 11053900, 11054000],
    "Seccion A.2": [11054202, 11054203, 11054204, 11054205],
    "Seccion A.3": [11056400, 11056401, 11056402, 11056403],
    "Seccion A.4": [11056418, 11056419, 11056420, 11056421],
}


def _find_col(df: pd.DataFrame, partial: str) -> str:
    for c in df.columns:
        if partial in c:
            return c
    raise KeyError(f"Columna con '{partial}' no encontrada")


def _col01_name(rem: str) -> str:
    if rem == "REM A05":
        return "COL01 - TOTAL"
    return "COL01 - TOTAL - Procesados"


@st.cache_data(show_spinner=False)
def load_sheet(anio: str, rem: str, sheet: str) -> pd.DataFrame:
    rem_file = {
        "REM A05": f"A05_{anio}.xlsx",
        "REM A11": f"A11_{anio}.xlsx",
    }[rem]
    path = BASE_DIR / anio / "salida" / rem_file
    df = pd.read_excel(path, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]
    return df


@st.cache_data(show_spinner=False)
def compute_indicators(
    anio: str,
    level: str | None = None,
    selected_entity: str | None = None,
) -> pd.DataFrame:
    raw: dict[str, float] = {}

    def _sum_codes(rem: str, sheet: str, codes: list[int]) -> float:
        df = load_sheet(anio, rem, sheet)
        prest_col = _find_col(df, "Prestaci")
        col = _find_col(df, _col01_name(rem))

        if level and selected_entity and level != "Nacional":
            geo_col = _find_col(df, level)
            df = df[df[geo_col] == selected_entity]

        mask = df[prest_col].isin(codes)
        return float(df.loc[mask, col].sum())

    raw["ingresos_prenatal"] = _sum_codes(
        "REM A05", "Seccion A", [1080008]
    )

    raw["gestantes_1er_vih"] = _sum_codes(
        "REM A11", "Seccion C.1", [11052500]
    )

    raw["gestantes_2do_vih"] = _sum_codes(
        "REM A11", "Seccion C.1", [11052600]
    )

    raw["gestantes_preparto_vih"] = _sum_codes(
        "REM A11", "Seccion C.1", [11052700]
    )

    sifilis_sum = 0.0
    for sheet, codes in SIFILIS_GESTANTE_CODES.items():
        sifilis_sum += _sum_codes("REM A11", sheet, codes)
    raw["gestantes_sifilis"] = sifilis_sum

    rows: list[dict] = []

    rows.append({
        "Indicador": "Ingresos a control prenatal (sistema público)",
        "Dato": raw["ingresos_prenatal"],
        "Tipo": "Número",
        "Observaciones": f"REM A05. Sección A. Dato disponible al {anio}.",
    })

    rows.append({
        "Indicador": "Número de gestantes con 1er examen VIH",
        "Dato": raw["gestantes_1er_vih"],
        "Tipo": "Número",
        "Observaciones": f"REM A11. Sección C.1. Dato disponible al {anio}.",
    })

    cobertura_vih = 0.0
    if raw["ingresos_prenatal"] > 0:
        cobertura_vih = raw["gestantes_1er_vih"] / raw["ingresos_prenatal"] * 100
    rows.append({
        "Indicador": "% Cobertura del 1er examen VIH en gestantes",
        "Dato": round(cobertura_vih, 1),
        "Tipo": "Porcentaje",
        "Observaciones": (
            "(Número de gestantes con 1er examen VIH / "
            "Ingresos a control prenatal) × 100"
        ),
    })

    rows.append({
        "Indicador": "Número de gestantes con 2º examen VIH",
        "Dato": raw["gestantes_2do_vih"],
        "Tipo": "Número",
        "Observaciones": f"REM A11. Sección C.1. Dato disponible al {anio}.",
    })

    rows.append({
        "Indicador": "Número de gestantes con examen VIH en prepartos",
        "Dato": raw["gestantes_preparto_vih"],
        "Tipo": "Número",
        "Observaciones": (
            "Suma de filas Gestante tamizada durante el parto cuyo resultado "
            "es reactivo (+) y no reactivo (-) del indicador 7b. "
            f"REM A11. Sección C.1."
        ),
    })

    cobertura_sifilis = 0.0
    if raw["ingresos_prenatal"] > 0:
        cobertura_sifilis = raw["gestantes_sifilis"] / raw["ingresos_prenatal"] * 100
    rows.append({
        "Indicador": "% Cobertura del 1er examen de Sífilis en gestantes",
        "Dato": round(cobertura_sifilis, 1),
        "Tipo": "Porcentaje",
        "Observaciones": (
            "(REM A11 suma secciones A.1 + A.2 + A.3 + A.4 / "
            "REM A05 sección A) × 100"
        ),
    })

    return pd.DataFrame(rows)


def format_dato(val: object, tipo: str) -> str:
    if pd.isna(val):
        return ""
    if tipo == "Porcentaje":
        return f"{float(val):.1f}%"
    return f"{int(float(val)):,}".replace(",", ".")


def render_indicadores_page() -> None:
    st.title("Indicadores Esenciales")
    render_year_badge()
    st.caption("Indicadores priorizados de VIH y sífilis en gestantes de la Región Metropolitana.")

    col_anio, col_nivel = st.columns([1, 2])
    with col_anio:
        anio = st.selectbox("Año", ANOS_DISPONIBLES, index=0)

    df_full_a05 = load_sheet(anio, "REM A05", "Seccion A")
    geo_col = _find_col(df_full_a05, "Servicio")

    niveles = ["Nacional", "Servicio de Salud"]
    if "Comuna" in [str(c).strip() for c in df_full_a05.columns]:
        niveles.append("Comuna")
    niveles.append("Establecimiento")

    with col_nivel:
        level = st.selectbox(
            "Nivel de desagregación",
            niveles,
            index=0,
            help="Selecciona el nivel geográfico para agregar los datos.",
        )

    selected_entity = None
    if level != "Nacional":
        df_sheet = load_sheet(anio, "REM A05", "Seccion A")
        geo_col_name = level
        options = sorted(
            str(v) for v in df_sheet[geo_col_name].dropna().unique() if str(v).strip()
        )
        selected_entity = st.selectbox(
            level,
            options,
            index=None,
            placeholder=f"Selecciona {level.lower()}...",
        )

    if selected_entity is None and level != "Nacional":
        st.info(f"Selecciona un {level.lower()} para ver los indicadores.")
        return

    df_indicadores = compute_indicators(anio, level, selected_entity)

    show_df = df_indicadores[["Indicador", "Dato", "Observaciones"]].copy()
    show_df["Dato"] = show_df.apply(
        lambda r: format_dato(r["Dato"], df_indicadores.loc[r.name, "Tipo"]),
        axis=1,
    )

    st.markdown("### Tabla de indicadores")

    st.dataframe(
        show_df,
        column_config={
            "Indicador": st.column_config.TextColumn("Indicador", width="large"),
            "Dato": st.column_config.TextColumn("Dato", width="small"),
            "Observaciones": st.column_config.TextColumn("Observaciones", width="large"),
        },
        hide_index=True,
        use_container_width=True,
    )

    st.info(
        "**Nota sobre las coberturas superiores al 100%:**\n\n"
        "- **VIH:** el numerador cuenta exámenes de laboratorio procesados "
        "(`COL01 - Procesados`), no mujeres únicas. Una gestante puede tener "
        "exámenes contabilizados sin estar registrada como ingreso a control "
        "prenatal en el mismo establecimiento (referencias externas), o "
        "generar más de un registro si se repite el examen.\n\n"
        "- **Sífilis:** el numerador suma 16 códigos distintos que combinan "
        "tipos de test (VDRL y RPR), trimestres (1º, 2º, 3º e ignorado) y "
        "modalidades (laboratorio propio y compra de servicio). Una misma "
        "gestante puede generar múltiples registros en el numerador, "
        "inflando la razón respecto al denominador de ingresos a prenatal."
    )

    st.markdown("### Descarga")
    buf_excel = BytesIO()
    with pd.ExcelWriter(buf_excel, engine="xlsxwriter") as writer:
        df_indicadores.to_excel(writer, sheet_name="Indicadores", index=False)
    buf_csv = df_indicadores.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

    b1, b2 = st.columns(2)
    with b1:
        st.download_button(
            "Descargar en Excel",
            data=buf_excel.getvalue(),
            file_name=f"indicadores_esenciales_{anio}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with b2:
        st.download_button(
            "Descargar en CSV",
            data=buf_csv,
            file_name=f"indicadores_esenciales_{anio}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with st.expander("Metodología"):
        st.markdown(f"""
**Indicadores Esenciales VIH / Sífilis en gestantes — Año {anio}**

| # | Indicador | Fuente | Fórmula |
|---|-----------|--------|---------|
| 1 | Ingresos a control prenatal | REM A05, Sección A, código 01080008 | Suma directa de COL01 (TOTAL) |
| 2 | Gestantes con 1er examen VIH | REM A11, Sección C.1, código 11052500 | Suma directa de COL01 (TOTAL - Procesados) |
| 3 | % Cobertura 1er examen VIH | — | (Indicador 2 / Indicador 1) × 100 |
| 4 | Gestantes con 2º examen VIH | REM A11, Sección C.1, código 11052600 | Suma directa de COL01 (TOTAL - Procesados) |
| 5 | Gestantes con examen VIH en prepartos | REM A11, Sección C.1, código 11052700 | Suma directa de COL01 (TOTAL - Procesados) |
| 6 | % Cobertura 1er examen Sífilis | REM A11, Secciones A.1–A.4 (suma de 16 códigos de gestantes por trimestre) | (Suma sífilis / Indicador 1) × 100 |

**Notas:**
- Los datos corresponden a la Región Metropolitana (RM), año {anio}.
- Las coberturas pueden superar el 100% si el número de exámenes realizados supera el número de ingresos a control prenatal (una gestante puede recibir más de un examen o ser derivada desde otro establecimiento).
- Los datos son provisorios y están sujetos a revisión por parte de la SEREMI de Salud RM.
        """)
