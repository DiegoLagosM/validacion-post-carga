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

HOJA = "CCV"

METRICAS = {
    "REGISTROS PROCESADOS :": "Registros Procesados (Definicion)",
    "REGISTROS PAREADOS :": "Registros Pareados",
    "REGISTROS NO PAREADOS :": "Registros NO Pareados",
    "REGISTROS CARGADOS :": "Registros Cargados (Paso a Produccion)",
}

PATRON_FECHA = re.compile(r"(\d{8,})")

# -------------------------------------------------
# FUNCIONES
# -------------------------------------------------

def leer_txt(ruta_txt: Path) -> dict | None:
    datos = {}
    seccion = None

    try:
        with ruta_txt.open(encoding="latin-1") as f:
            for linea in map(str.upper, f):
                linea = re.sub(r"\s+", " ", linea.strip())

                if "ESTADISTICAS DE DEFINICION DE ACCIONES" in linea:
                    seccion = "definicion"
                    continue

                if "ESTADISTICAS DE PASO A PRODUCCION" in linea:
                    seccion = "produccion"
                    continue

                for patron, nombre in METRICAS.items():
                    if patron in linea:
                        if seccion == "definicion" or patron == "REGISTROS CARGADOS :":
                            datos[nombre] = int(re.search(r"\d+", linea).group())
    except Exception:
        return None

    return datos if len(datos) == len(METRICAS) else None


def procesar_logs() -> pd.DataFrame:
    registros = []
    archivos_procesados = 0

    # ✅ Leer .txt y .log
    archivos = list(RUTA_LOGS.glob("CCV*.txt")) + list(RUTA_LOGS.glob("CCV*.log"))

    for archivo in archivos:
        match = PATRON_FECHA.search(archivo.name)

        if not match:
            print(f"⚠️ Sin fecha en nombre: {archivo.name}")
            continue

        fecha_raw = match.group(1)[:8]

        try:
            fecha = pd.to_datetime(fecha_raw, format="%Y%m%d")
        except ValueError:
            print(f"⚠️ Fecha inválida en {archivo.name}")
            continue

        print(f"📄 Procesando: {archivo.name} → {fecha.date()}")

        datos = leer_txt(archivo)

        if not datos:
            print(f"⚠️ Datos incompletos en {archivo.name}")
            continue

        archivos_procesados += 1

        for metrica, valor in datos.items():
            registros.append({
                "Metrica": metrica,
                "Fecha": fecha,
                "Valor": valor,
            })

    print(f"✅ Total archivos procesados: {archivos_procesados}")

    df = pd.DataFrame(registros)

    if not df.empty:
        # ✅ Suma cuando hay misma fecha
        df = (
            df.groupby(["Metrica", "Fecha"], as_index=False)["Valor"]
            .sum()
        )

    return df


def cargar_excel() -> pd.DataFrame:
    if not RUTA_EXCEL.exists():
        return pd.DataFrame({"Metrica": []})

    try:
        return pd.read_excel(RUTA_EXCEL, sheet_name=HOJA)
    except Exception:
        return pd.DataFrame({"Metrica": []})


def transformar(df: pd.DataFrame) -> pd.DataFrame:
    return (
        df.pivot(index="Metrica", columns="Fecha", values="Valor")
        .reset_index()
    )


def actualizar(existente: pd.DataFrame, nuevo: pd.DataFrame) -> pd.DataFrame:
    existente = existente.set_index("Metrica")
    nuevo = nuevo.set_index("Metrica")

    resultado = existente.combine_first(nuevo)
    resultado.update(nuevo)

    return resultado.reset_index()


def ordenar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    columnas_fecha = sorted(
        c for c in df.columns if isinstance(c, (pd.Timestamp, datetime))
    )
    return df[["Metrica", *columnas_fecha]]


# -------------------------------------------------
# MAIN
# -------------------------------------------------

df_logs = procesar_logs()

if df_logs.empty:
    raise SystemExit("❌ No hay datos CCV válidos")

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
    if_sheet_exists="replace",
) as writer:
    final.to_excel(writer, sheet_name=HOJA, index=False)

print("✅ Proceso CCV completado correctamente")