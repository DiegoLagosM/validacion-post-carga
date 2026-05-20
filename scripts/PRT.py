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

HOJA = "PRT"

CAMPOS = {
    "Total Registros Procesados": "Correctos",
    "Total Registros Erroneos": "Erroneos",
    "Total Registros Tipo NP": "Tipo NP",
    "Total Registros Tipo AN": "Tipo AN",
    "ACL": "ACL",
    "CFP": "CFP",
    "CPA-PPD": "CPA-PPD",
    "Total Registros Tipo AE": "Tipo AE",
    "Total Registros Tipo REC": "Tipo REC",
}

# -------------------------------------------------
# REGEX
# -------------------------------------------------
RE_FECHA = re.compile(r"PRT_(\d{8})")
RE_NUMERO = re.compile(r":\s*(\d+)")
RE_INFORMADOS = re.compile(r"BC5070\s*[:\-]?\s*(\d+)")

# -------------------------------------------------
# FUNCIONES
# -------------------------------------------------
def leer_txt(ruta: Path) -> dict | None:
    datos = {}

    try:
        with ruta.open(encoding="utf-8") as f:
            for linea in f:
                if "BC5070" in linea:
                    m = RE_INFORMADOS.search(linea)
                    if m:
                        datos["Informados"] = int(m.group(1))

                for texto, concepto in CAMPOS.items():
                    if texto in linea:
                        m = RE_NUMERO.search(linea)
                        if m:
                            datos[concepto] = int(m.group(1))

    except Exception as e:
        print(f"❌ Error leyendo {ruta.name}: {e}")
        return None

    return datos or None


def extraer_fecha(nombre_archivo: str) -> pd.Timestamp | None:
    """Extrae fecha desde el nombre del archivo como Timestamp."""
    m = RE_FECHA.search(nombre_archivo)
    if not m:
        return None
    return pd.to_datetime(m.group(1), format="%Y%m%d")


def procesar_logs() -> pd.DataFrame:
    registros = []

    for archivo in RUTA_LOGS.glob("validacion_y_carga_PRT_*.txt"):
        fecha = extraer_fecha(archivo.name)
        if fecha is None:
            continue

        datos = leer_txt(archivo)
        if not datos:
            print(f"⚠️ Sin datos válidos en {archivo.name}")
            continue

        registros.extend(
            {
                "Concepto": concepto,
                "Fecha": fecha,
                "Valor": valor,
            }
            for concepto, valor in datos.items()
        )

    return pd.DataFrame.from_records(registros)


def transformar(df: pd.DataFrame) -> pd.DataFrame:
    return (
        df.pivot(index="Concepto", columns="Fecha", values="Valor")
        .reset_index()
    )


def cargar_excel() -> pd.DataFrame:
    if not RUTA_EXCEL.exists():
        return pd.DataFrame(columns=["Concepto"])

    try:
        df = pd.read_excel(RUTA_EXCEL, sheet_name=HOJA)

        # ✅ Blindaje: normalizar columnas de fecha que vengan como texto
        df.columns = [
            pd.to_datetime(c) if isinstance(c, str) and c != "Concepto" else c
            for c in df.columns
        ]

        return df
    except Exception as e:
        print(f"❌ Error leyendo Excel: {e}")
        return pd.DataFrame(columns=["Concepto"])


def actualizar(existente: pd.DataFrame, nuevo: pd.DataFrame) -> pd.DataFrame:
    existente = existente.set_index("Concepto")
    nuevo = nuevo.set_index("Concepto")

    resultado = existente.combine_first(nuevo)
    resultado.update(nuevo)

    return resultado.reset_index()


def ordenar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    # ✅ Acepta Timestamp y datetime.datetime
    columnas_fecha = sorted(
        c for c in df.columns
        if isinstance(c, (pd.Timestamp, datetime))
    )
    return df[["Concepto", *columnas_fecha]]


# -------------------------------------------------
# MAIN
# -------------------------------------------------
if not RUTA_LOGS.exists():
    raise FileNotFoundError(f"No existe la ruta: {RUTA_LOGS}")

df_logs = procesar_logs()

if df_logs.empty:
    raise SystemExit("❌ No hay datos válidos para procesar")

nuevo = transformar(df_logs)
existente = cargar_excel()

final = nuevo if existente.empty else actualizar(existente, nuevo)
final = ordenar_columnas(final)

# -------------------------------------------------
# GUARDAR EXCEL
# -------------------------------------------------
with pd.ExcelWriter(
    RUTA_EXCEL,
    engine="openpyxl",
    mode="a" if RUTA_EXCEL.exists() else "w",
    if_sheet_exists="replace",
) as writer:
    final.to_excel(writer, sheet_name=HOJA, index=False)

print("✅ PRT actualizado correctamente")