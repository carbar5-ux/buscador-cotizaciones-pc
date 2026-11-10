# Arma la carpeta del programa con un Python propio (embebido oficial de python.org), la ventana (tkinter)
# y todas las librerías, lista para que Inno Setup la empaquete. Se ejecuta en GitHub Actions (Windows).
#   ./instalador/armar.ps1 -PyVersion 3.12.10
param([Parameter(Mandatory = $true)][string]$PyVersion)
$ErrorActionPreference = "Stop"

$raiz = Split-Path $PSScriptRoot -Parent
$app = Join-Path $PSScriptRoot "app"
$py = Join-Path $app "python"
Remove-Item $app -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory $py | Out-Null

# 1) Python embebido oficial (el mismo número de versión que el Python completo del runner)
$zip = Join-Path $env:RUNNER_TEMP "python-embed.zip"
Invoke-WebRequest "https://www.python.org/ftp/python/$PyVersion/python-$PyVersion-embed-amd64.zip" -OutFile $zip
Expand-Archive $zip -DestinationPath $py
$mm = ($PyVersion.Split(".")[0..1] -join "")          # "312"
# Rutas de búsqueda: la biblioteca, Lib (tkinter), las librerías y la carpeta del programa (..)
Set-Content (Join-Path $py "python$mm._pth") "python$mm.zip`r`n.`r`nLib`r`nLib\site-packages`r`n..`r`nimport site" -Encoding ascii

# 2) La ventana (tkinter + Tcl/Tk): el embebido no la trae; se copia del Python completo del runner
$base = (python -c "import sys; print(sys.base_prefix)").Trim()
New-Item -ItemType Directory (Join-Path $py "Lib") | Out-Null
Copy-Item (Join-Path $base "Lib\tkinter") (Join-Path $py "Lib\tkinter") -Recurse
foreach ($patron in "_tkinter.pyd", "tcl*.dll", "tk*.dll", "zlib*.dll") {
    Get-ChildItem (Join-Path $base "DLLs") -Filter $patron | Copy-Item -Destination $py
}
Copy-Item (Join-Path $base "tcl") (Join-Path $py "tcl") -Recurse

# 3) Librerías del programa (+ pip, para que el actualizador pueda instalar librerías nuevas)
python -m pip install --no-cache-dir --disable-pip-version-check --target (Join-Path $py "Lib\site-packages") `
    -r (Join-Path $raiz "requirements.txt") pip
if ($LASTEXITCODE) { throw "No se pudieron instalar las librerías" }

# 4) Archivos del programa (sin los .bat del modo con Python aparte: el instalador crea sus propios accesos)
Get-ChildItem $raiz -File | Where-Object {
    ($_.Extension -in ".py", ".ico", ".pdf", ".md", ".txt" -or $_.Name -in "VERSION", "config.toml") -and
    $_.Name -notin "requirements-dev.txt"
} | Copy-Item -Destination $app

Write-Host "Carpeta del programa lista: $app"
Get-ChildItem $app | Format-Table Name, Length
