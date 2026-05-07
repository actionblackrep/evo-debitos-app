@echo off
REM Reporte de Debitos - Windows
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo Python no esta instalado. Descarga desde https://www.python.org/downloads/
  pause
  exit /b 1
)

if not exist ".venv" (
  echo [setup] Creando entorno virtual...
  python -m venv .venv
)

call .venv\Scripts\activate.bat

echo [setup] Instalando dependencias...
python -m pip install --upgrade pip >nul
pip install -r requirements.txt
if errorlevel 1 goto :err

echo [run] Ejecutando agente...
python generate_report.py %*
if errorlevel 1 goto :err

echo.
echo Listo. PDF disponible en .\output\
pause
exit /b 0

:err
echo.
echo Hubo un error. Revisa el mensaje arriba.
pause
exit /b 1
