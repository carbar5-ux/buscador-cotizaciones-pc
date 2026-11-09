@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "CARPETA=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$c=$env:CARPETA; $d=[Environment]::GetFolderPath('Desktop'); $w=New-Object -ComObject WScript.Shell; $s=$w.CreateShortcut((Join-Path $d 'Buscador Cotizaciones.lnk')); $s.TargetPath=(Join-Path $c 'vigilar.bat'); $s.WorkingDirectory=$c; $s.IconLocation=(Join-Path $c 'icono_cotizaciones.ico')+',0'; $s.Description='Lee las cotizaciones PDF y actualiza el Excel'; $s.Save(); $s=$w.CreateShortcut((Join-Path $d 'Excel Cotizaciones.lnk')); $s.TargetPath=(Join-Path $c 'abrir_excel.bat'); $s.WorkingDirectory=$c; $s.IconLocation=(Join-Path $c 'icono_excel.ico')+',0'; $s.Description='Abre el Excel base de cotizaciones'; $s.Save(); $s=$w.CreateShortcut((Join-Path $d 'Manual Cotizaciones.lnk')); $s.TargetPath=(Join-Path $c 'Manual_de_uso.pdf'); $s.WorkingDirectory=$c; $s.Description='Manual de uso del Buscador de cotizaciones'; $s.Save()"
if errorlevel 1 (echo No se pudieron crear los accesos directos. & pause & exit /b 1)
echo Listo: en tu Escritorio quedaron tres accesos directos:
echo   - Buscador Cotizaciones  (verde $)   : lee las cotizaciones y actualiza el Excel
echo   - Excel Cotizaciones     (naranjo $) : abre el Excel base
echo   - Manual Cotizaciones              : abre el manual de uso (PDF)
echo Si el icono no aparece de inmediato, presiona F5 en el Escritorio.
if /i "%~1"=="/sinpausa" exit /b 0
pause
