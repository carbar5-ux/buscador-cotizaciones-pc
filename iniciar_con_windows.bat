@echo off
chcp 65001 >nul
cd /d "%~dp0"
set INICIO=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
(echo @start "" /d "%~dp0" "%~dp0.venv\Scripts\pythonw.exe" "%~dp0buscador_cotizaciones.py" --vigilar) > "%INICIO%\Buscador cotizaciones.cmd"
echo Listo: el buscador se iniciara solo (en segundo plano) cada vez que enciendas el PC.
echo Para quitarlo, borra "Buscador cotizaciones.cmd" de la carpeta: %INICIO%
echo El registro de lo que hace queda en registro.log
pause
