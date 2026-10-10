@echo off
REM ============================================================
REM  Libreria MM - Inicio rapido
REM  Activa el entorno virtual y ejecuta: streamlit run app.py
REM  Uso: doble clic en este archivo o desde el acceso directo
REM         del Escritorio (ver crear_acceso_directo.bat).
REM ============================================================
chcp 65001 >nul
title Libreria MM

REM Ir a la carpeta del proyecto (donde esta este .bat)
cd /d "%~dp0"

REM Verificar archivos clave
if not exist "app.py" (
    echo [ERROR] No se encontro app.py en "%CD%".
    echo Ejecuta este archivo desde la raiz del proyecto C:\LibreriaMM.
    pause
    exit /b 1
)
if not exist "venv\Scripts\activate.bat" (
    echo [ERROR] No se encontro venv\Scripts\activate.bat.
    echo Crea el entorno con: python -m venv venv
    pause
    exit /b 1
)

REM Activar entorno virtual
call "venv\Scripts\activate.bat"

REM Ejecutar Streamlit (python -m es mas robusto que streamlit.exe)
echo Iniciando Libreria MM en http://localhost:8501 ...
echo Para detener: pulsa Ctrl+C en esta ventana.
python -m streamlit run app.py

REM Si Streamlit termina con error, no cerrar de golpe
if errorlevel 1 (
    echo.
    echo [ERROR] Streamlit termino con codigo %errorlevel%.
    pause
)
