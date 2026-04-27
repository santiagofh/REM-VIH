from __future__ import annotations

import json
import re
import shutil
from pathlib import Path

import pandas as pd


# Rutas de entrada
DICCIONARIO_DIR = Path(
    r"C:\Users\fariass\OneDrive - SUBSECRETARIA DE SALUD PUBLICA\Escritorio\REM\REM-VIH\2025"
)
DATOS_REM_DIR = Path(
    r"C:\Users\fariass\OneDrive - SUBSECRETARIA DE SALUD PUBLICA\Escritorio\DATA\REM\REM_2025\Datos"
)
DEIS_PATH = Path(
    r"C:\Users\fariass\OneDrive - SUBSECRETARIA DE SALUD PUBLICA\Escritorio\DATA\ESTABLECIMIENTOS\establecimientos_deis_actual.xlsx"
)

# Carpeta de salida
OUTPUT_DIR = DICCIONARIO_DIR / "salida"

# Parametros generales
REGION_RM = "13"
ANIO = "2025"
CHUNKSIZE = 200_000
CSV_ENCODING = "utf-8-sig"
CSV_SEP = ";"


REM_CONFIG = {
    "A05": {
        "csv_path": DATOS_REM_DIR / "SerieA2025.csv",
        "output_name": "A05_2025.xlsx",
        "secciones_path": DICCIONARIO_DIR / "REM_A05_2025_seccion_a_codigos.json",
        "detalle_path": DICCIONARIO_DIR / "REM_A05_2025_codigo_a_detalle.json",
        "columnas_path": DICCIONARIO_DIR / "REM_A05_2025_columnas_por_codigo.json",
        "acumular_meses": True,
        "mes_corte": None,
    },
    "P11": {
        "csv_path": DATOS_REM_DIR / "SerieP2025.csv",
        "output_name": "P11_2025.xlsx",
        "secciones_path": DICCIONARIO_DIR / "REM_P11_2025_seccion_a_codigos.json",
        "detalle_path": DICCIONARIO_DIR / "REM_P11_2025_codigo_a_detalle.json",
        "columnas_path": DICCIONARIO_DIR / "REM_P11_2025_columnas_por_codigo.json",
        "acumular_meses": False,
        "mes_corte": "12",
    },
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


def nombre_hoja(seccion: str) -> str:
    limpio = re.sub(r"[\[\]\:\*\?\/\\]", "_", str(seccion))
    return f"Seccion {limpio}"[:31]


def cargar_deis() -> pd.DataFrame:
    columnas_deis = {
        "EstablecimientoCodigo": "Codigo DEIS",
        "EstablecimientoGlosa": "Establecimiento",
        "RegionCodigo": "Codigo Region DEIS",
        "RegionGlosa": "Región",
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


def cargar_diccionarios(config: dict) -> tuple[dict[str, list[str]], dict[str, str], dict[str, dict[str, str]]]:
    seccion_a_codigos = cargar_json(config["secciones_path"])
    codigo_a_detalle = cargar_json(config["detalle_path"])
    columnas_data = cargar_json(config["columnas_path"])
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

    # En el CSV vienen como Col01, Col02, etc.; en el diccionario como COL01.
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
    mes_corte: str | None = None,
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
            & chunk["CodigoPrestacion"].isin(codigos_solicitados)
        )
        if mes_corte is not None:
            filtro = filtro & chunk["Mes"].eq(mes_corte)

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
    df["Detalle de prestación"] = df["Codigo Prestacion"].map(codigo_a_detalle).fillna("")

    df = df.merge(deis, how="left", on="Codigo DEIS")

    # Si el establecimiento no aparece en DEIS, se mantiene la fila y quedan vacios
    # los detalles descriptivos. Los codigos fuente quedan disponibles para trazabilidad.
    for col in ["Región", "Servicio de Salud", "Comuna", "Establecimiento"]:
        df[col] = df[col].fillna("")

    df["Prestación"] = df["Codigo Prestacion"]
    df["Mes"] = df["Mes"].map(normalizar_texto)
    df["Año"] = df["Ano"].map(normalizar_texto)
    df["Codigo Servicio REM"] = df["IdServicio"].map(normalizar_texto)
    df["Codigo Comuna REM"] = df["IdComuna"].map(normalizar_texto)

    return df


def construir_hoja(
    df_base: pd.DataFrame,
    seccion: str,
    codigos_seccion: list[str],
    columnas_seccion: dict[str, str],
    acumular_meses: bool,
) -> pd.DataFrame:
    codigos_set = set(codigos_seccion)
    hoja = df_base[df_base["Codigo Prestacion"].isin(codigos_set)].copy()

    columnas_valor = []
    renombrar = {}
    for codigo_col, detalle_col in columnas_seccion.items():
        col_csv = f"Col{codigo_col[-2:]}"
        if col_csv in hoja.columns:
            nombre = f"{codigo_col} - {detalle_col}"
            columnas_valor.append(col_csv)
            renombrar[col_csv] = nombre
            hoja[col_csv] = pd.to_numeric(hoja[col_csv].replace("", pd.NA), errors="coerce")

    columnas_identificacion = [
        "Región",
        "Servicio de Salud",
        "Comuna",
        "Establecimiento",
        "Codigo DEIS",
        "Prestación",
        "Detalle de prestación",
        "Año",
        "Codigo Servicio REM",
        "Codigo Comuna REM",
    ]

    if acumular_meses:
        if not hoja.empty:
            hoja = (
                hoja.groupby(columnas_identificacion, dropna=False, as_index=False)[columnas_valor]
                .sum(min_count=1)
            )
        hoja["Mes"] = "Enero-Diciembre"
    else:
        columnas_identificacion = columnas_identificacion[:7] + ["Mes"] + columnas_identificacion[7:]

    columnas_finales = columnas_identificacion + (["Mes"] if acumular_meses else []) + columnas_valor

    hoja = hoja[columnas_finales].rename(columns=renombrar)

    orden = [
        "Servicio de Salud",
        "Comuna",
        "Establecimiento",
        "Codigo DEIS",
        "Prestación",
        "Mes",
    ]
    if not hoja.empty:
        hoja = hoja.sort_values(orden, kind="stable")

    return hoja


def ajustar_formato_excel(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:
    ws = writer.book[sheet_name]
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, col_name in enumerate(df.columns, start=1):
        max_len = min(max(12, len(str(col_name)) + 2), 55)
        ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = max_len


def generar_excel(rem: str, config: dict, deis: pd.DataFrame) -> None:
    print(f"Preparando {rem}...")
    seccion_a_codigos, codigo_a_detalle, columnas_por_seccion = cargar_diccionarios(config)

    codigos_solicitados = {
        codigo
        for codigos in seccion_a_codigos.values()
        for codigo in codigos
    }
    columnas_uso = columnas_csv_necesarias(columnas_por_seccion)

    datos = leer_rem_filtrado(
        config["csv_path"],
        codigos_solicitados,
        columnas_uso,
        mes_corte=config["mes_corte"],
    )
    datos = preparar_base(datos, deis, codigo_a_detalle)

    output_path = OUTPUT_DIR / config["output_name"]
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for seccion, codigos_seccion in seccion_a_codigos.items():
            columnas_seccion = columnas_por_seccion.get(seccion, {})
            hoja = construir_hoja(
                datos,
                seccion,
                codigos_seccion,
                columnas_seccion,
                acumular_meses=config["acumular_meses"],
            )
            sheet_name = nombre_hoja(seccion)
            hoja.to_excel(writer, sheet_name=sheet_name, index=False)
            ajustar_formato_excel(writer, sheet_name, hoja)
            print(f"  Hoja {sheet_name}: {len(hoja):,} filas")

    print(f"Archivo creado: {output_path}")


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    deis = cargar_deis()

    for rem, config in REM_CONFIG.items():
        generar_excel(rem, config, deis)


if __name__ == "__main__":
    main()
