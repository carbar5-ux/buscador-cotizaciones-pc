@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo === Instalando Buscador de cotizaciones ===
where py >nul 2>nul && (set PY=py -3) || (set PY=python)
%PY% --version >nul 2>nul || (echo No se encontro Python. Instalalo desde https://www.python.org/downloads/ marcando "Add python.exe to PATH" y vuelve a ejecutar este archivo. & pause & exit /b 1)
if not exist .venv %PY% -m venv .venv
.venv\Scripts\python -m pip install --upgrade pip
.venv\Scripts\python -m pip install -r requirements.txt
call crear_acceso_directo.bat /sinpausa
echo.
echo Listo. Ahora ejecuta "procesar.bat" (una vez) o "vigilar.bat" (queda atento a cotizaciones nuevas).
pause
