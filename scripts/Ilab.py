import pandas as pd
from pathlib import Path
import re
from datetime import datetime

# =================================================
# CONFIGURACIÓN
# =================================================
BASE_PATH = Path(
    r"C:\Users\c31352e\OneDrive - EXPERIAN SERVICES CORP\Archivos de Cardenas, Oscar - Calidad de Datos\Diego\Validacion post carga"
)

RUTA_LOGS = BASE_PATH / "LOGS"
RUTA_EXCEL = BASE_PATH / "Consolidado validaciones de carga 2026.xlsx"
SHEET_NAME = "ILab"

# =================================================
# DEFINICIONES
# =================================================
BLOQUES = {
    "Estadistica del Archivo Codigos": "ACU",
    "Estadistica del Archivo Resumen": "ARE",
    "Estadistica del Archivo Detalle": "ADE",
    "Estadistica del Archivo Trabajadores": "ATR",
}

PATRONES_BASE = {
    "Total_Procesados": re.compile(r"Total Registros Procesados\s*:\s*([\d\s]+)"),
    "Total_Existentes": re.compile(r"Total Registros Existentes\s*:\s*([\d\s]+)"),
    "Total_Cargados": re.compile(r"Total Registros Cargados\s*:\s*([\d\s]+)"),
    "Total_Reemplazados": re.compile(r"Total Registros Reemplazados\s*:\s*([\d\s]+)"),
    "Total_Eliminados": re.compile(r"Total Registros Eliminados\s*:\s*([\d\s]+)"),
    "Total_Erroneos": re.compile(r"Total Registros Erroneos\s*:\s*([\d\s]+)"),
}

REGEX_FECHA = re.compile(r"(\d{8})")

# =================================================
# FUNCIONES
# =================================================
def extraer_fecha(nombre: str):
    if m := REGEX_FECHA.search(nombre):
        try:
            return pd.to_datetime(m.group(1), format="%Y%m%d")
        except Exception:
            return None
    return None


def leer_txt(path: Path):
    datos = {}
    prefijo_actual = None

    try:
        with path.open(encoding="utf-8", errors="ignore") as f:
            for linea in f:
                # Detectar bloque activo
                for titulo, prefijo in BLOQUES.items():
                    if titulo in linea:
                        prefijo_actual = prefijo
                        break

                if not prefijo_actual:
                    continue

                # Extraer métricas
                for campo, patron in PATRONES_BASE.items():
                    if m := patron.search(linea):
                        datos[f"{prefijo_actual}_{campo}"] = int(
                            m.group(1).replace(" ", "")
                        )
    except Exception:
        return None

    return datos or None


def procesar_logs() -> pd.DataFrame:
    if not RUTA_LOGS.exists():
        raise FileNotFoundError(f"No existe la carpeta: {RUTA_LOGS}")

    registros = []
    archivos = list(RUTA_LOGS.glob("laborales_diarios_*.txt"))

    if not archivos:
        print("⚠️ No hay archivos de logs ILab")
        return pd.DataFrame()

    for archivo in archivos:
        fecha = extraer_fecha(archivo.name)
        if fecha is None:
            continue

        datos = leer_txt(archivo)
        if not datos:
            continue

        registros.extend(
            {
                "Metrica": metrica,
                "Fecha": fecha,
                "Valor": valor,
            }
            for metrica, valor in datos.items()
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


def cargar_excel() -> pd.DataFrame:
    if not RUTA_EXCEL.exists():
        return pd.DataFrame(columns=["Metrica"])

    try:
        df = pd.read_excel(RUTA_EXCEL, sheet_name=SHEET_NAME)

        # Normalizar columnas de fecha sin warnings
        df.columns = [
            pd.to_datetime(col, errors="ignore") if col != "Metrica" else col
            for col in df.columns
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


# =================================================
# MAIN
# =================================================
def main():
    df_logs = procesar_logs()

    if df_logs.empty:
        print("⚠️ Sin datos para procesar en ILab")
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

    print("✅ ILab actualizado correctamente")


if __name__ == "__main__":
    main()