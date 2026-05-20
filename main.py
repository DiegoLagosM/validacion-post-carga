import subprocess
import sys
import os
import time

# ✅ Corrección clave: evitar recursión en .exe
if getattr(sys, 'frozen', False):
    PYTHON_EXE = "py"
else:
    PYTHON_EXE = sys.executable

SCRIPTS = [
    r"C:\Users\c31352e\OneDrive - EXPERIAN SERVICES CORP\Escritorio\Python\CCC-ONP.py",
    r"C:\Users\c31352e\OneDrive - EXPERIAN SERVICES CORP\Escritorio\Python\CCV.py",
    r"C:\Users\c31352e\OneDrive - EXPERIAN SERVICES CORP\Escritorio\Python\Cedula.py",
    r"C:\Users\c31352e\OneDrive - EXPERIAN SERVICES CORP\Escritorio\Python\PRT.py",
    r"C:\Users\c31352e\OneDrive - EXPERIAN SERVICES CORP\Escritorio\Python\Aclaraciones.py",
    r"C:\Users\c31352e\OneDrive - EXPERIAN SERVICES CORP\Escritorio\Python\Infocom.py",
    r"C:\Users\c31352e\OneDrive - EXPERIAN SERVICES CORP\Escritorio\Python\ILab.py",
    r"C:\Users\c31352e\OneDrive - EXPERIAN SERVICES CORP\Escritorio\Python\IG - CA.py",
    r"C:\Users\c31352e\OneDrive - EXPERIAN SERVICES CORP\Escritorio\Python\R04.py",
    r"C:\Users\c31352e\OneDrive - EXPERIAN SERVICES CORP\Escritorio\Python\SIR.py"
]

def main():
    start_time = time.time()
    errores = []

    print("🚀 Inicio de ejecución secuencial\n")

    for script in SCRIPTS:
        nombre = os.path.basename(script)
        print(f"▶ Ejecutando {nombre}")

        try:
            result = subprocess.run(
                [PYTHON_EXE, script],
                text=True
            )

            if result.returncode != 0:
                print(f"❌ Error en {nombre}")
                errores.append(nombre)
                break
            else:
                print(f"✅ {nombre} terminado correctamente\n")

        except Exception as e:
            print(f"⚠️ Excepción ejecutando {nombre}: {e}")
            errores.append(nombre)
            break

    # ✅ Resumen final
    end_time = time.time()
    duration = end_time - start_time

    print("\n📊 RESUMEN FINAL")
    print(f"⏱ Tiempo total: {duration:.2f} segundos")

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