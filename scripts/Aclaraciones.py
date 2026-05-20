import pandas as pd
from pathlib import Path
import re

# -------------------------------------------------
# CONFIGURACIÓN
# -------------------------------------------------

BASE_PATH = Path(
    r"C:\Users\c31352e\OneDrive - EXPERIAN SERVICES CORP\Archivos de Cardenas, Oscar - Calidad de Datos\Diego\Validacion post carga"
)

RUTA_LOGS = BASE_PATH / "LOGS"
RUTA_EXCEL = BASE_PATH / "Consolidado validaciones de carga 2026.xlsx"
HOJA = "Aclaraciones"

PATRON_FECHA = re.compile(r"(\d{8})")

METRICAS = [
    "TOTAL PROCESADOS",
    "NP",
    "AN",
    "ACL",
    "CFP",
    "AE",
    "REC"
]

# -------------------------------------------------
# ✅ LECTURA ROBUSTA Y CORRECTA
# -------------------------------------------------

def leer_txt(ruta_txt: Path):
    try:
        with ruta_txt.open(encoding="latin-1") as f:
            contenido = f.read().upper()

        # ✅ Buscar bloque correcto (evita capturar fechas u otros números)
        match = re.search(
            r"TOTAL PROCESADOS.*?NP:.*?AN:.*?ACL:.*?CFP:.*?AE:.*?REC:",
            contenido,
            re.DOTALL
        )

        if not match:
            print(f"⚠️ No se encontró bloque válido en {ruta_txt.name}")
            return None

        linea = re.sub(r"\s+", " ", match.group())

        print(f"\n📄 {ruta_txt.name}")
        print(f"LINEA CORRECTA → {linea}")

        # ✅ extracción precisa (por etiqueta)
        patrones = {
            "TOTAL PROCESADOS": r"TOTAL PROCESADOS.*?(\d+)",
            "NP": r"NP:\s*(\d+)",
            "AN": r"AN:\s*(\d+)",
            "ACL": r"ACL:\s*(\d+)",
            "CFP": r"CFP:\s*(\d+)",
            "AE": r"AE:\s*(\d+)",
            "REC": r"REC:\s*(\d+)",
        }

        datos = {}

        for clave, patron in patrones.items():
            m = re.search(patron, linea)
            if m:
                datos[clave] = int(m.group(1))

        print(f"✔ EXTRAÍDO → {datos}")

        return datos if datos else None

    except Exception as e:
        print(f"❌ Error en {ruta_txt.name}: {e}")
        return None


# -------------------------------------------------
# PROCESAR LOGS
# -------------------------------------------------

def procesar_logs():
    registros = []

    for archivo in RUTA_LOGS.glob("AclaracionesDiarias*.txt"):

        match = PATRON_FECHA.search(archivo.name)
        if not match:
            continue

        fecha = pd.to_datetime(match.group(1), format="%Y%m%d") - pd.Timedelta(days=1)

        datos = leer_txt(archivo)

        if not datos:
            continue

        for metrica, valor in datos.items():
            registros.append({
                "Metrica": metrica,
                "Fecha": fecha,
                "Valor": valor
            })

    df = pd.DataFrame(registros)

    print("\n📊 DATAFRAME LOGS:")
    print(df)

    return df


# -------------------------------------------------
# PIVOT
# -------------------------------------------------

def transformar(df):
    return (
        df.pivot_table(
            index="Metrica",
            columns="Fecha",
            values="Valor",
            aggfunc="first"
        ).reset_index()
    )


# -------------------------------------------------
# CARGAR EXCEL
# -------------------------------------------------

def cargar_excel():
    if not RUTA_EXCEL.exists():
        return pd.DataFrame()

    try:
        df = pd.read_excel(RUTA_EXCEL, sheet_name=HOJA)

        # normalizar columnas fecha
        nuevas_cols = []
        for c in df.columns:
            if c != "Metrica":
                try:
                    nuevas_cols.append(pd.to_datetime(c))
                except:
                    nuevas_cols.append(c)
            else:
                nuevas_cols.append(c)

        df.columns = nuevas_cols

        # eliminar métricas antiguas
        df = df[df["Metrica"].isin(METRICAS)]

        return df

    except:
        return pd.DataFrame()


# -------------------------------------------------
# ACTUALIZAR
# -------------------------------------------------

def actualizar(existente, nuevo):
    if existente.empty:
        return nuevo

    df = pd.concat([existente, nuevo])
    df = df.groupby("Metrica", as_index=False).first()

    return df


# -------------------------------------------------
# ORDENAR COLUMNAS
# -------------------------------------------------

def ordenar_columnas(df):
    cols = [c for c in df.columns if c != "Metrica"]

    cols = sorted([
        pd.to_datetime(c) for c in cols if not isinstance(c, str)
    ])

    return df[["Metrica", *cols]]


# -------------------------------------------------
# MAIN
# -------------------------------------------------

df_logs = procesar_logs()

if df_logs.empty:
    raise SystemExit("❌ No hay datos válidos")

nuevo = transformar(df_logs)
existente = cargar_excel()

final = actualizar(existente, nuevo)
final = ordenar_columnas(final)

print("\n📊 RESULTADO FINAL:")
print(final)

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

print("\n✅ Proceso completado correctamente")