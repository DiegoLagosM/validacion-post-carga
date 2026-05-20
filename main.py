import subprocess
import sys
import os
import time

# ✅ Detecta el ejecutable correcto
if getattr(sys, 'frozen', False):
    PYTHON_EXE = "python"
else:
    PYTHON_EXE = sys.executable

# ✅ Carpeta base del proyecto
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ✅ Lista de scripts (en carpeta scripts)
SCRIPTS = [
    os.path.join(BASE_DIR, "scripts", "CCC-ONP.py"),
    os.path.join(BASE_DIR, "scripts", "CCV.py"),
    os.path.join(BASE_DIR, "scripts", "Cedula.py"),
    os.path.join(BASE_DIR, "scripts", "PRT.py"),
    os.path.join(BASE_DIR, "scripts", "Aclaraciones.py"),
    os.path.join(BASE_DIR, "scripts", "Infocom.py"),
    os.path.join(BASE_DIR, "scripts", "ILab.py"),
    os.path.join(BASE_DIR, "scripts", "IG_CA.py"),
    os.path.join(BASE_DIR, "scripts", "R04.py"),
    os.path.join(BASE_DIR, "scripts", "SIR.py"),
]

def main():
    start_time = time.time()
    errores = []

    print("🚀 Inicio de ejecución secuencial\n")

    # ✅ Mostrar rutas (para debug en GitHub)
    print("📂 Scripts detectados:")
    for s in SCRIPTS:
        print(s)
    print()

    for script in SCRIPTS:
        nombre = os.path.basename(script)

        # ✅ Validación clave (evita tu error actual)
        if not os.path.exists(script):
            print(f"❌ No existe el archivo: {script}")
            errores.append(nombre)
            break

        print(f"▶ Ejecutando {nombre}")

        try:
            result = subprocess.run(
                [PYTHON_EXE, script],
                text=True,
                capture_output=True
            )

            # ✅ Logs del script
            if result.stdout:
                print(result.stdout)
            if result.stderr:
                print(result.stderr)

            if result.returncode != 0:
                print(f"❌ Error en {nombre}")
                errores.append(nombre)
                break

            print(f"✅ {nombre} terminado correctamente\n")

        except Exception as e:
            print(f"⚠️ Excepción ejecutando {nombre}: {e}")
            errores.append(nombre)
            break

    total_time = time.time() - start_time

    print("\n📊 RESUMEN FINAL")
    print(f"⏱ Tiempo total: {total_time:.2f} segundos")

    if errores:
        print("❌ El proceso terminó con errores en:")
        for e in errores:
            print(f"   - {e}")
        print("⛔ Estado final: FALLIDO")
        return 1
    else:
        print("✅ Todos los scripts terminaron correctamente")
        print("🎉 Estado final: EXITOSO")
        return 0


if __name__ == "__main__":
    sys.exit(main())
