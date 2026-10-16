"""Lanzador de Librería MM (.exe).

Debe vivir en la raíz del proyecto, junto a app.py y venv/.
Al ejecutarlo: verifica app.py y el entorno virtual, levanta
`streamlit run app.py` y abre el navegador. Ctrl+C para detener.
"""
import subprocess
import sys
import time
import webbrowser
from pathlib import Path


def base_dir() -> Path:
    if getattr(sys, "frozen", False):  # corriendo como .exe compilado
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def main() -> int:
    root = base_dir()
    app_py = root / "app.py"
    py = root / "venv" / "Scripts" / "python.exe"
    print("=" * 52)
    print("  LIBRERÍA MM — iniciando aplicación...")
    print("=" * 52)
    if not app_py.exists():
        print(f"[ERROR] No se encontró app.py en: {root}")
        input("Pulsa Enter para salir...")
        return 1
    if not py.exists():
        print(f"[ERROR] No se encontró el entorno virtual en: {py.parent}")
        print("Créalo con:  python -m venv venv  +  pip install -r requirements.txt")
        input("Pulsa Enter para salir...")
        return 1
    url = "http://localhost:8501"
    print(f"Abriendo {url} ...")
    print("Para detener la app: pulsa Ctrl+C en esta ventana.")
    print("-" * 52)
    proc = subprocess.Popen(
        [str(py), "-m", "streamlit", "run", str(app_py),
         "--server.headless", "true", "--browser.gatherUsageStats", "false"],
        cwd=str(root),
    )
    time.sleep(4)
    try:
        webbrowser.open(url)
    except Exception:
        pass
    try:
        proc.wait()
    except KeyboardInterrupt:
        print("\nDeteniendo Librería MM...")
        proc.terminate()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
