@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
title Monitor de inundaciones - Actualizar desde Excel

set "CARPETA=E:\Usuarios\ecacifuentes\OneDrive - CETCO S.A\Operaciones\Distribución\Panificación Fenómeno del Niño\files excel"

cd /d "%CARPETA%" 2>nul || (
  echo [ERROR] No se pudo acceder a la carpeta:
  echo   %CARPETA%
  echo.
  pause & exit /b 1
)

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd"') do set "HOY=%%i"
if "%HOY%"=="" set "HOY=salida"

set "PY="
where python >nul 2>&1 && set "PY=python"
if not defined PY ( where py >nul 2>&1 && set "PY=py" )
if not defined PY (
  echo [ERROR] No se encontro Python en el PATH.
  echo Instala Python 3 desde https://www.python.org/downloads/ y marca "Add Python to PATH".
  echo.
  pause & exit /b 1
)

set "PARTE=%~1"
if "%PARTE%"=="" set "PARTE=parte_diario_template.xlsx"
set "BASE=%~2"
if "%BASE%"=="" set "BASE=monitor_base.json"
set "SALIDA=%~3"
if "%SALIDA%"=="" set "SALIDA=monitor_%HOY%.json"

if not exist "actualizar_desde_excel.py" ( echo [ERROR] Falta actualizar_desde_excel.py en la carpeta. & echo. & pause & exit /b 1 )
if not exist "%PARTE%" ( echo [ERROR] No existe el parte: "%PARTE%" & echo. & pause & exit /b 1 )
if not exist "%BASE%"  ( echo [ERROR] No existe el JSON base: "%BASE%" & echo. & pause & exit /b 1 )

%PY% -c "import openpyxl" 1>nul 2>nul
if errorlevel 1 (
  echo Instalando dependencia openpyxl...
  %PY% -m pip install openpyxl
  if errorlevel 1 ( echo [ERROR] No se pudo instalar openpyxl. & echo. & pause & exit /b 1 )
)

echo ============================================================
echo  Carpeta: %CD%
echo  Parte  : %PARTE%
echo  Base   : %BASE%
echo  Salida : %SALIDA%
echo ============================================================
echo.

%PY% "actualizar_desde_excel.py" "%PARTE%" "%BASE%" "%SALIDA%"
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
  echo [OK] Listo. Espera a que OneDrive sincronice "%SALIDA%" e importalo en el dashboard ^(modo Estado operativo^).
) else (
  echo [ERROR] El script termino con codigo %RC%.
)
echo.
pause
endlocal