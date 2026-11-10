# Prueba la carpeta armada por armar.ps1 ANTES de crear el instalador: que el Python incluido tenga la ventana
# (tkinter) y las librerías, y que el programa lea cotizaciones de ejemplo y arme el Excel.
$ErrorActionPreference = "Stop"
$raiz = Split-Path $PSScriptRoot -Parent

# Se prueba en una copia, para no tocar el config.toml que va en el instalador
$prueba = Join-Path $env:RUNNER_TEMP "prueba-app"
Remove-Item $prueba -Recurse -Force -ErrorAction SilentlyContinue
Copy-Item (Join-Path $PSScriptRoot "app") $prueba -Recurse
$py = Join-Path $prueba "python\python.exe"

# 1) Python incluido: ventana y librerías
& $py -c "import tkinter, pdfplumber, openpyxl, pypdfium2, pip; r = tkinter.Tk(); r.withdraw(); print('Tcl/Tk', r.tk.call('info', 'patchlevel')); r.destroy()"
if ($LASTEXITCODE) { throw "El Python incluido no tiene la ventana o las librerías" }
& $py -c "import ventana, buscador_cotizaciones, actualizador; print('módulos del programa OK, versión', actualizador.version_local())"
if ($LASTEXITCODE) { throw "No se pueden importar los módulos del programa" }

# 2) Lectura real de cotizaciones de ejemplo (con todos los núcleos, como en el PC)
python -m pip install --disable-pip-version-check -q reportlab
$ejemplos = Join-Path $env:RUNNER_TEMP "ejemplos"
python (Join-Path $raiz "pruebas\generar_ejemplos.py") $ejemplos
Get-ChildItem $ejemplos -Recurse -Filter *.pdf | ForEach-Object { $_.LastWriteTime = (Get-Date).AddMinutes(-5) }
$cfg = Join-Path $prueba "config.toml"
(Get-Content $cfg) -replace "^actualizar_automatico = .*", "actualizar_automatico = false" | Set-Content $cfg
& $py (Join-Path $prueba "buscador_cotizaciones.py") --carpeta $ejemplos
if ($LASTEXITCODE) { throw "El programa terminó con error" }
$excel = Join-Path $ejemplos "Base_Cotizaciones.xlsx"
if (-not (Test-Path $excel)) { throw "No se creó el Excel" }
& $py -c "import openpyxl, sys; wb = openpyxl.load_workbook(sys.argv[1]); n = wb['Precios'].max_row - 1; print('Excel OK:', wb.sheetnames, n, 'líneas'); sys.exit(0 if n > 0 else 1)" $excel
if ($LASTEXITCODE) { throw "El Excel no tiene líneas de precio" }
Write-Host "Prueba del programa empaquetado: OK"
