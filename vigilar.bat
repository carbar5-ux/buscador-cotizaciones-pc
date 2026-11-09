@echo off
chcp 65001 >nul & cd /d "%~dp0" & .venv\Scripts\python -c "import tkinter" 2>nul && (start "" ".venv\Scripts\pythonw.exe" ventana.py) || (title Buscador de cotizaciones & .venv\Scripts\python buscador_cotizaciones.py --vigilar || pause) & exit /b
