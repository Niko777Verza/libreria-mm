@echo off
REM ============================================================
REM  Libreria MM - Generador de acceso directo en el Escritorio
REM  Crea "Libreria MM.lnk" en el Escritorio apuntando a run_app.bat
REM  Uso: doble clic en este archivo (una sola vez).
REM ============================================================
setlocal
chcp 65001 >nul
title Libreria MM - Crear acceso directo

REM Rutas (este .bat vive en la raiz del proyecto)
set "TARGET=%~dp0run_app.bat"
set "WORKDIR=%~dp0"
REM Quitar la barra final para WorkingDirectory (queda C:\LibreriaMM)
set "WORKDIR=%WORKDIR:~0,-1%"
set "SHORTCUT=%USERPROFILE%\Desktop\Libreria MM.lnk"

if not exist "%TARGET%" (
    echo [ERROR] No se encontro "%TARGET%".
    echo Asegurate de que run_app.bat este junto a este archivo.
    pause
    exit /b 1
)

echo Creando acceso directo...
echo   Destino : "%TARGET%"
echo   Carpeta : "%WORKDIR%"
echo   Acceso  : "%SHORTCUT%"
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command "$s = (New-Object -ComObject WScript.Shell).CreateShortcut('%SHORTCUT%'); $s.TargetPath = '%TARGET%'; $s.WorkingDirectory = '%WORKDIR%'; $s.Description = 'Libreria MM - Sistema de gestion'; $s.IconLocation = '%WORKDIR%\venv\Scripts\python.exe,0'; $s.Save()"

if exist "%SHORTCUT%" (
    echo.
    echo [OK] Acceso directo creado: "%SHORTCUT%"
    echo Ya puedes abrir Libreria MM con doble clic desde el Escritorio.
) else (
    echo.
    echo [ERROR] No se pudo crear el acceso directo.
    echo Revisa permisos de escritura en el Escritorio.
    pause
    exit /b 1
)

endlocal
pause
