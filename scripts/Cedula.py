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

HOJA = "Cedulas"

METRICA_TOTAL = "Total de Registros Cargados"
METRICA_ERROR = "Porcentaje de Erroneos"
METRICAS = {METRICA_TOTAL, METRICA_ERROR}

ARCHIVO_RE = re.compile(r"^Cedulas_(\d{8})\.txt$", re.IGNORECASE)

# -------------------------------------------------
# FUNCIONES
# -------------------------------------------------
def normalizar(linea: str) -> str:
    return re.sub(r"\s+", " ", linea.upper()).strip()


def leer_txt(ruta: Path) -> dict | None:
    datos = {}
    en_bloque = False
    paso_prod = False

    try:
        with ruta.open(encoding="latin-1") as f:
            for linea in map(normalizar, f):

                if "ESTADISTICA DE CARGADOR DE CEDULAS BLOQUEADAS" in linea:
                    en_bloque = True
                    paso_prod = False
                    continue

                if not en_bloque:
                    continue

                if "PROCESO: PASO A PRODUCCION" in linea:
                    paso_prod = True
                    continue

                if paso_prod and METRICA_TOTAL.upper() in linea:
                    datos[METRICA_TOTAL] = int(
                        re.search(r"\d+", linea).group()
                    )

                elif not paso_prod and METRICA_ERROR.upper() in linea:
                    valor = re.search(r"[\d.,]+", linea).group()
                    datos[METRICA_ERROR] = float(
                        valor.replace(",", ".")
                    )

    except Exception as e:
        print(f"❌ Error leyendo {ruta.name}: {e}")
        return None

    return datos if METRICAS.issubset(datos) else None


def procesar_logs() -> pd.DataFrame:
    registros = []

    # ✅ SOLO archivos Cedulas
    for archivo in RUTA_LOGS.glob("Cedulas_*.txt"):
        match = ARCHIVO_RE.match(archivo.name)
        if not match:
            continue

        # ✅ FECHA UNIFICADA A Timestamp
        fecha = pd.to_datetime(match.group(1), format="%Y%m%d")

        datos = leer_txt(archivo)

        if not datos:
            print(f"⚠️ Datos incompletos en {archivo.name}")
            continue

        registros.extend(
            {"Metrica": k, "Fecha": fecha, "Valor": v}
            for k, v in datos.items()
        )

    return pd.DataFrame(registros, columns=["Metrica", "Fecha", "Valor"])


def cargar_excel() -> pd.DataFrame:
    if not RUTA_EXCEL.exists():
        return pd.DataFrame(columns=["Metrica"])

    try:
        return pd.read_excel(RUTA_EXCEL, sheet_name=HOJA)
    except Exception:
        return pd.DataFrame(columns=["Metrica"])


def transformar(df: pd.DataFrame) -> pd.DataFrame:
    return (
        df.pivot(index="Metrica", columns="Fecha", values="Valor")
        .reset_index()
    )


def actualizar(existente: pd.DataFrame, nuevo: pd.DataFrame) -> pd.DataFrame:
    existente = existente.set_index("Metrica")
    nuevo = nuevo.set_index("Metrica")

    combinado = existente.combine_first(nuevo)
    combinado.update(nuevo)

    return combinado.reset_index()


def ordenar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    columnas = ["Metrica"] + sorted(
        c for c in df.columns if c != "Metrica"
    )
    return df[columnas]


# -------------------------------------------------
# MAIN
# -------------------------------------------------
df_logs = procesar_logs()

if df_logs.empty:
    print("❌ No hay datos válidos para procesar")
    exit()

nuevo = transformar(df_logs)
existente = cargar_excel()

final = nuevo if existente.empty else actualizar(existente, nuevo)
final = ordenar_columnas(final)

# -------------------------------------------------
# GUARDAR
# -------------------------------------------------
with pd.ExcelWriter(
    RUTA_EXCEL,
    engine="openpyxl",
    mode="a" if RUTA_EXCEL.exists() else "w",
    if_sheet_exists="replace"
) as writer:
    final.to_excel(writer, sheet_name=HOJA, index=False)

print("✅ Hoja Cedulas actualizada correctamente")