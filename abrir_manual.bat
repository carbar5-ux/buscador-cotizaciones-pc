@echo off
cd /d "%~dp0" & if exist Manual_de_uso.pdf (start "" "Manual_de_uso.pdf") else (echo Todavia no esta el manual: ejecuta actualizar.bat & pause) & exit /b
