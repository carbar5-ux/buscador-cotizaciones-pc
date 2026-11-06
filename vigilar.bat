@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Buscador de cotizaciones - vigilando
.venv\Scripts\python buscador_cotizaciones.py --vigilar
pause
