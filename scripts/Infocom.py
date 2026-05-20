import pandas as pd
from pathlib import Path
import re
from datetime import datetime

# -------------------------------------------------
# CONFIGURACIÓN
# -------------------------------------------------
BASE_PATH = Path(
    r"C:\Users\c31352e\OneDrive - EXPERIAN SERVICES CORP\Archivos de Cardenas, Oscar - Calidad de Datos\Diego\Validacion post carga"
)

RUTA_LOGS = BASE_PATH / "LOGS"
RUTA_EXCEL = BASE_PATH / "Consolidado validaciones de carga 2026.xlsx"
SHEET_NAME = "Infocom"

# -------------------------------------------------
# REGEX
# -------------------------------------------------
BLOQUE_VALIDACION = re.compile(
    r"Estadistica Final del Proceso de Validacion\s*-+\s*(.*?)(?:\n\s*\n|$)",
    re.DOTALL | re.IGNORECASE
)

BLOQUE_CARGA = re.compile(
    r"Estadistica Final del Proceso de Carga\s*-+\s*(.*?)(?:\n\s*\n|$)",
    re.DOTALL | re.IGNORECASE
)

CAMPOS = {
    "validacion": {
        "Total Registros Procesados (VAL)": r"Total\s+Registros\s+Procesados\s*:\s*(\d+)",
        "Total Registros Correctos": r"Total\s+Registros\s+Correctos\s*:\s*(\d+)",
        "Total Registros Erroneos": r"Total\s+Registros\s+Erroneos\s*:\s*(\d+)",
    },
    "opc4": {
        "Total Registros Procesados (CARGA OPC 4)": r"Total\s+Registros\s+Procesados\s*:\s*(\d+)",
        "Total Registros Cargados (CARGA OPC 4)": r"Total\s+Registros\s+Cargados\s*:\s*(\d+)",
        "Total Registros Modificados (CARGA OPC 4)": r"Total\s+Registros\s+Modificados\s*:\s*(\d+)",
    },
    "opc5": {
        "Total Registros Procesados (CARGA OPC 5)": r"Total\s+Registros\s+Procesados\s*:\s*(\d+)",
        "Total Registros Cargados (CARGA OPC 5)": r"Total\s+Registros\s+Cargados\s*:\s*(\d+)",
        "Total Registros Modificados (CARGA OPC 5)": r"Total\s+Registros\s+Modificados\s*:\s*(\d+)",
        "Total Registros Eliminados (CARGA OPC 5)": r"Total\s+Registros\s+Eliminados\s*:\s*(\d+)",
    },
}

# -------------------------------------------------
# UTILIDADES
# -------------------------------------------------
def extraer_fecha(nombre_archivo: str):
    try:
        fecha_raw = nombre_archivo.replace("Infocom_carga_", "").replace(".txt", "")
        return pd.to_datetime(fecha_raw, format="%Y%m%d")
    except Exception:
        return None


def extraer_campos(bloque: str, patrones: dict):
    return {
        campo: int(m.group(1))
        for campo, patron in patrones.items()
        if (m := re.search(patron, bloque))
    }


# -------------------------------------------------
# PARSEO DE LOG
# -------------------------------------------------
def leer_log(path: Path):
    try:
        contenido = path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return None

    fecha = extraer_fecha(path.name)
    if fecha is None:
        return None

    datos = {}

    if m := BLOQUE_VALIDACION.search(contenido):
        datos.update(extraer_campos(m.group(1), CAMPOS["validacion"]))

    for opcion in ("opc4", "opc5"):
        idx = contenido.lower().find(f"opcion : {opcion[-1]}")
        if idx == -1:
            continue

        if m := BLOQUE_CARGA.search(contenido[idx:]):
            datos.update(extraer_campos(m.group(1), CAMPOS[opcion]))

    return (fecha, datos) if datos else None


# -------------------------------------------------
# PROCESAMIENTO
# -------------------------------------------------
def procesar_logs() -> pd.DataFrame:
    if not RUTA_LOGS.exists():
        raise FileNotFoundError(f"No existe la carpeta de logs: {RUTA_LOGS}")

    registros = []

    archivos = list(RUTA_LOGS.glob("Infocom_carga_*.txt"))

    if not archivos:
        print("⚠️ No hay archivos de log")
        return pd.DataFrame()

    for archivo in archivos:
        resultado = leer_log(archivo)
        if not resultado:
            continue

        fecha, datos = resultado
        registros.extend(
            {"Metrica": m, "Fecha": fecha, "Valor": v}
            for m, v in datos.items()
        )

    return pd.DataFrame(registros)


def pivotear(df: pd.DataFrame) -> pd.DataFrame:
    return (
        df.pivot_table(
            index="Metrica",
            columns="Fecha",
            values="Valor",
            aggfunc="first",
        )
        .rename_axis(None, axis=1)
        .reset_index()
    )


# -------------------------------------------------
# EXCEL
# -------------------------------------------------
def cargar_excel() -> pd.DataFrame:
    if not RUTA_EXCEL.exists():
        return pd.DataFrame(columns=["Metrica"])

    try:
        df = pd.read_excel(RUTA_EXCEL, sheet_name=SHEET_NAME)

        df.columns = [
            pd.to_datetime(c, errors="ignore") if c != "Metrica" else c
            for c in df.columns
        ]

        return df
    except Exception:
        return pd.DataFrame(columns=["Metrica"])


def actualizar(existente: pd.DataFrame, nuevo: pd.DataFrame) -> pd.DataFrame:
    existente = existente.set_index("Metrica")
    nuevo = nuevo.set_index("Metrica")

    resultado = existente.combine_first(nuevo)
    resultado.update(nuevo)

    return resultado.reset_index()


def ordenar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    columnas_fecha = sorted(
        [c for c in df.columns if isinstance(c, (pd.Timestamp, datetime))]
    )
    return df[["Metrica", *columnas_fecha]]


# -------------------------------------------------
# MAIN
# -------------------------------------------------
def main():
    df_logs = procesar_logs()

    if df_logs.empty:
        print("⚠️ Sin datos para procesar")
        return

    nuevo = pivotear(df_logs)
    existente = cargar_excel()

    final = nuevo if existente.empty else actualizar(existente, nuevo)
    final = ordenar_columnas(final)

    with pd.ExcelWriter(
        RUTA_EXCEL,
        engine="openpyxl",
        mode="a" if RUTA_EXCEL.exists() else "w",
        if_sheet_exists="replace",
    ) as writer:
        final.to_excel(writer, sheet_name=SHEET_NAME, index=False)

    print("✅ Infocom actualizado correctamente")


if __name__ == "__main__":
    main()