from __future__ import annotations

import json
import re
import shutil
from pathlib import Path

import pandas as pd


DICCIONARIO_DIR = Path(
    r"C:\Users\fariass\OneDrive - SUBSECRETARIA DE SALUD PUBLICA\Escritorio\REM\REM-VIH\2025"
)
DATOS_REM_DIR = Path(
    r"D:\DATA\REM\REM_2025\Datos"
)
DEIS_PATH = Path(
    r"C:\Users\fariass\OneDrive - SUBSECRETARIA DE SALUD PUBLICA\Escritorio\DATA\ESTABLECIMIENTOS\establecimientos_deis_actual.xlsx"
)

OUTPUT_DIR = DICCIONARIO_DIR / "salida_comparativo_p11"

REGION_RM = "13"
ANIO = "2025"
MESES_COMPARAR = ("6", "12")
CHUNKSIZE = 200_000
CSV_ENCODING = "utf-8-sig"
CSV_SEP = ";"

REM_CONFIG = {
    "csv_path": DATOS_REM_DIR / "SerieP2025.csv",
    "secciones_path": DICCIONARIO_DIR / "REM_P11_2025_seccion_a_codigos.json",
    "detalle_path": DICCIONARIO_DIR / "REM_P11_2025_codigo_a_detalle.json",
    "columnas_path": DICCIONARIO_DIR / "REM_P11_2025_columnas_por_codigo.json",
    "comparativo_output_name": "P11_2025_comparativo_mes06_vs_mes12.xlsx",
    "alertas_output_name": "P11_2025_informe_alertas_mes06_vs_mes12.xlsx",
}


def cargar_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def normalizar_texto(valor: object) -> str:
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    return texto


def normalizar_codigo_prestacion(valor: object) -> str:
    return normalizar_texto(valor).upper()


def nombre_hoja(seccion: str, sufijo: str = "") -> str:
    limpio = re.sub(r"[\[\]\:\*\?\/\\]", "_", str(seccion))
    return f"Seccion {limpio}{sufijo}"[:31]


def cargar_deis() -> pd.DataFrame:
    columnas_deis = {
        "EstablecimientoCodigo": "Codigo DEIS",
        "EstablecimientoGlosa": "Establecimiento",
        "RegionCodigo": "Codigo Region DEIS",
        "RegionGlosa": "Region",
        "SeremiSaludCodigo_ServicioDeSaludCodigo": "Codigo Servicio DEIS",
        "SeremiSaludGlosa_ServicioDeSaludGlosa": "Servicio de Salud",
        "ComunaCodigo": "Codigo Comuna DEIS",
        "ComunaGlosa": "Comuna",
    }

    try:
        deis = pd.read_excel(
            DEIS_PATH,
            sheet_name="establecimientos",
            dtype=str,
            usecols=list(columnas_deis.keys()),
        )
    except PermissionError:
        copia_temporal = OUTPUT_DIR / "establecimientos_deis_actual_tmp.xlsx"
        copia_temporal.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(DEIS_PATH, copia_temporal)
        deis = pd.read_excel(
            copia_temporal,
            sheet_name="establecimientos",
            dtype=str,
            usecols=list(columnas_deis.keys()),
        )

    deis = deis.rename(columns=columnas_deis)
    for col in deis.columns:
        deis[col] = deis[col].map(normalizar_texto)

    deis["Codigo DEIS"] = deis["Codigo DEIS"].map(normalizar_texto)
    deis = deis.drop_duplicates(subset=["Codigo DEIS"], keep="first")
    return deis


def cargar_diccionarios() -> tuple[dict[str, list[str]], dict[str, str], dict[str, dict[str, str]]]:
    seccion_a_codigos = cargar_json(REM_CONFIG["secciones_path"])
    codigo_a_detalle = cargar_json(REM_CONFIG["detalle_path"])
    columnas_data = cargar_json(REM_CONFIG["columnas_path"])
    columnas_por_seccion = columnas_data["columnas_por_seccion"]

    seccion_a_codigos = {
        seccion: [normalizar_codigo_prestacion(codigo) for codigo in codigos]
        for seccion, codigos in seccion_a_codigos.items()
    }
    codigo_a_detalle = {
        normalizar_codigo_prestacion(codigo): detalle
        for codigo, detalle in codigo_a_detalle.items()
    }

    return seccion_a_codigos, codigo_a_detalle, columnas_por_seccion


def columnas_csv_necesarias(columnas_por_seccion: dict[str, dict[str, str]]) -> list[str]:
    codigos_columna = set()
    for columnas in columnas_por_seccion.values():
        codigos_columna.update(columnas.keys())

    columnas_rem = [f"Col{codigo[-2:]}" for codigo in sorted(codigos_columna)]
    columnas_base = [
        "Mes",
        "IdServicio",
        "Ano",
        "IdEstablecimiento",
        "CodigoPrestacion",
        "IdRegion",
        "IdComuna",
    ]
    return columnas_base + columnas_rem


def leer_rem_filtrado(
    csv_path: Path,
    codigos_solicitados: set[str],
    columnas_uso: list[str],
) -> pd.DataFrame:
    partes = []

    for chunk in pd.read_csv(
        csv_path,
        sep=CSV_SEP,
        encoding=CSV_ENCODING,
        dtype=str,
        keep_default_na=False,
        usecols=columnas_uso,
        chunksize=CHUNKSIZE,
    ):
        chunk["IdRegion"] = chunk["IdRegion"].map(normalizar_texto)
        chunk["Ano"] = chunk["Ano"].map(normalizar_texto)
        chunk["Mes"] = chunk["Mes"].map(normalizar_texto)
        chunk["CodigoPrestacion"] = chunk["CodigoPrestacion"].map(normalizar_codigo_prestacion)

        filtro = (
            chunk["IdRegion"].eq(REGION_RM)
            & chunk["Ano"].eq(ANIO)
            & chunk["Mes"].isin(MESES_COMPARAR)
            & chunk["CodigoPrestacion"].isin(codigos_solicitados)
        )

        if filtro.any():
            partes.append(chunk.loc[filtro].copy())

    if not partes:
        return pd.DataFrame(columns=columnas_uso)

    return pd.concat(partes, ignore_index=True)


def preparar_base(
    df: pd.DataFrame,
    deis: pd.DataFrame,
    codigo_a_detalle: dict[str, str],
) -> pd.DataFrame:
    df = df.copy()
    df["Codigo DEIS"] = df["IdEstablecimiento"].map(normalizar_texto)
    df["Codigo Prestacion"] = df["CodigoPrestacion"].map(normalizar_codigo_prestacion)
    df["Detalle de prestacion"] = df["Codigo Prestacion"].map(codigo_a_detalle).fillna("")

    df = df.merge(deis, how="left", on="Codigo DEIS")

    for col in ["Region", "Servicio de Salud", "Comuna", "Establecimiento"]:
        df[col] = df[col].fillna("")

    df["Prestacion"] = df["Codigo Prestacion"]
    df["Mes"] = df["Mes"].map(normalizar_texto)
    df["Ano"] = df["Ano"].map(normalizar_texto)
    df["Codigo Servicio REM"] = df["IdServicio"].map(normalizar_texto)
    df["Codigo Comuna REM"] = df["IdComuna"].map(normalizar_texto)

    return df


def clasificar_alerta(total_mes_6: float, total_mes_12: float) -> str:
    if total_mes_6 > 0 and total_mes_12 == 0:
        return "Pasa a 0 en mes 12"
    if total_mes_12 < total_mes_6:
        return "Baja de mes 6 a mes 12"
    if total_mes_6 == 0 and total_mes_12 > 0:
        return "Aparece solo en mes 12"
    if total_mes_6 == total_mes_12:
        return "Sin variacion"
    return "Aumento esperado"


def construir_comparativo_seccion(
    df_base: pd.DataFrame,
    seccion: str,
    codigos_seccion: list[str],
    columnas_seccion: dict[str, str],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    hoja = df_base[df_base["Codigo Prestacion"].isin(set(codigos_seccion))].copy()

    columnas_valor = []
    renombrar = {}
    for codigo_col, detalle_col in columnas_seccion.items():
        col_csv = f"Col{codigo_col[-2:]}"
        if col_csv in hoja.columns:
            columnas_valor.append(col_csv)
            renombrar[col_csv] = f"{codigo_col} - {detalle_col}"
            hoja[col_csv] = pd.to_numeric(hoja[col_csv].replace("", pd.NA), errors="coerce").fillna(0)

    columnas_id = [
        "Region",
        "Servicio de Salud",
        "Comuna",
        "Establecimiento",
        "Codigo DEIS",
        "Prestacion",
        "Detalle de prestacion",
        "Ano",
        "Codigo Servicio REM",
        "Codigo Comuna REM",
        "Mes",
    ]

    if hoja.empty:
        comparativo_vacio = pd.DataFrame(columns=columnas_id[:-1] + ["Total mes 6", "Total mes 12", "Diferencia", "Variacion %", "Alerta"])
        return comparativo_vacio, comparativo_vacio.copy()

    agrupado = hoja.groupby(columnas_id, dropna=False, as_index=False)[columnas_valor].sum()

    pivot = agrupado.pivot_table(
        index=columnas_id[:-1],
        columns="Mes",
        values=columnas_valor,
        aggfunc="sum",
        fill_value=0,
    )

    pivot.columns = [f"{renombrar[col]} | Mes {mes}" for col, mes in pivot.columns]
    comparativo = pivot.reset_index()

    columnas_mes_6 = []
    columnas_mes_12 = []
    for col_csv in columnas_valor:
        nombre = renombrar[col_csv]
        col_6 = f"{nombre} | Mes 6"
        col_12 = f"{nombre} | Mes 12"

        if col_6 not in comparativo.columns:
            comparativo[col_6] = 0
        if col_12 not in comparativo.columns:
            comparativo[col_12] = 0

        delta_col = f"{nombre} | Diferencia"
        comparativo[delta_col] = comparativo[col_12] - comparativo[col_6]

        columnas_mes_6.append(col_6)
        columnas_mes_12.append(col_12)

    comparativo["Total mes 6"] = comparativo[columnas_mes_6].sum(axis=1)
    comparativo["Total mes 12"] = comparativo[columnas_mes_12].sum(axis=1)
    comparativo["Diferencia"] = comparativo["Total mes 12"] - comparativo["Total mes 6"]
    comparativo["Variacion %"] = 0.0
    mask_base = comparativo["Total mes 6"] != 0
    comparativo.loc[mask_base, "Variacion %"] = (
        comparativo.loc[mask_base, "Diferencia"] / comparativo.loc[mask_base, "Total mes 6"] * 100
    ).round(2)
    comparativo["Alerta"] = comparativo.apply(
        lambda fila: clasificar_alerta(fila["Total mes 6"], fila["Total mes 12"]),
        axis=1,
    )

    orden_fijo = [
        "Region",
        "Servicio de Salud",
        "Comuna",
        "Establecimiento",
        "Codigo DEIS",
        "Prestacion",
        "Detalle de prestacion",
        "Ano",
        "Codigo Servicio REM",
        "Codigo Comuna REM",
        "Total mes 6",
        "Total mes 12",
        "Diferencia",
        "Variacion %",
        "Alerta",
    ]
    orden_detalle = []
    for col_csv in columnas_valor:
        nombre = renombrar[col_csv]
        orden_detalle.extend(
            [
                f"{nombre} | Mes 6",
                f"{nombre} | Mes 12",
                f"{nombre} | Diferencia",
            ]
        )

    comparativo = comparativo[orden_fijo + orden_detalle].sort_values(
        ["Servicio de Salud", "Comuna", "Establecimiento", "Prestacion"],
        kind="stable",
    )

    alertas = comparativo[
        comparativo["Alerta"].isin(["Pasa a 0 en mes 12", "Baja de mes 6 a mes 12"])
    ].copy()

    return comparativo, alertas


def construir_resumen_alertas(alertas_por_seccion: dict[str, pd.DataFrame]) -> pd.DataFrame:
    filas = []
    for seccion, alertas in alertas_por_seccion.items():
        if alertas.empty:
            filas.append(
                {
                    "Seccion": seccion,
                    "Total alertas": 0,
                    "Pasa a 0 en mes 12": 0,
                    "Baja de mes 6 a mes 12": 0,
                }
            )
            continue

        conteo = alertas["Alerta"].value_counts()
        filas.append(
            {
                "Seccion": seccion,
                "Total alertas": len(alertas),
                "Pasa a 0 en mes 12": int(conteo.get("Pasa a 0 en mes 12", 0)),
                "Baja de mes 6 a mes 12": int(conteo.get("Baja de mes 6 a mes 12", 0)),
            }
        )

    return pd.DataFrame(filas).sort_values("Seccion", kind="stable")


def ajustar_formato_excel(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:
    ws = writer.book[sheet_name]
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, col_name in enumerate(df.columns, start=1):
        max_len = min(max(12, len(str(col_name)) + 2), 55)
        ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = max_len


def generar_salidas() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    deis = cargar_deis()
    seccion_a_codigos, codigo_a_detalle, columnas_por_seccion = cargar_diccionarios()

    codigos_solicitados = {
        codigo
        for codigos in seccion_a_codigos.values()
        for codigo in codigos
    }
    columnas_uso = columnas_csv_necesarias(columnas_por_seccion)
    datos = leer_rem_filtrado(REM_CONFIG["csv_path"], codigos_solicitados, columnas_uso)
    datos = preparar_base(datos, deis, codigo_a_detalle)

    comparativo_path = OUTPUT_DIR / REM_CONFIG["comparativo_output_name"]
    alertas_path = OUTPUT_DIR / REM_CONFIG["alertas_output_name"]

    alertas_por_seccion: dict[str, pd.DataFrame] = {}

    with pd.ExcelWriter(comparativo_path, engine="openpyxl") as writer_comparativo:
        for seccion, codigos_seccion in seccion_a_codigos.items():
            comparativo, alertas = construir_comparativo_seccion(
                datos,
                seccion,
                codigos_seccion,
                columnas_por_seccion.get(seccion, {}),
            )
            alertas_por_seccion[seccion] = alertas

            sheet_name = nombre_hoja(seccion)
            comparativo.to_excel(writer_comparativo, sheet_name=sheet_name, index=False)
            ajustar_formato_excel(writer_comparativo, sheet_name, comparativo)

        resumen = construir_resumen_alertas(alertas_por_seccion)
        resumen.to_excel(writer_comparativo, sheet_name="Resumen alertas", index=False)
        ajustar_formato_excel(writer_comparativo, "Resumen alertas", resumen)

    with pd.ExcelWriter(alertas_path, engine="openpyxl") as writer_alertas:
        resumen = construir_resumen_alertas(alertas_por_seccion)
        resumen.to_excel(writer_alertas, sheet_name="Resumen alertas", index=False)
        ajustar_formato_excel(writer_alertas, "Resumen alertas", resumen)

        for seccion, alertas in alertas_por_seccion.items():
            sheet_name = nombre_hoja(seccion, " alertas")
            alertas.to_excel(writer_alertas, sheet_name=sheet_name, index=False)
            ajustar_formato_excel(writer_alertas, sheet_name, alertas)

    print(f"Archivo comparativo creado: {comparativo_path}")
    print(f"Archivo de alertas creado: {alertas_path}")


if __name__ == "__main__":
    generar_salidas()
