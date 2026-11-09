@echo off
chcp 65001 >nul
cd /d "%~dp0"
set INICIO=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
(echo @start "" /d "%~dp0" "%~dp0.venv\Scripts\pythonw.exe" "%~dp0ventana.py" --minimizado) > "%INICIO%\Buscador cotizaciones.cmd"
echo Listo: el buscador se abrira solo (minimizado en la barra de tareas) cada vez que enciendas el PC.
echo Para quitarlo, borra "Buscador cotizaciones.cmd" de la carpeta: %INICIO%
echo El registro de lo que hace queda en registro.log
pause
