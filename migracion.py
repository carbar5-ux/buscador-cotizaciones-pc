"""Trae la configuración del programa antiguo (el que se instalaba con instalar.bat) al programa instalado.

La primera vez que se abre el programa instalado, si todavía no tiene carpeta de cotizaciones, busca la
instalación anterior y copia config.toml, categorias.txt y sinonimos.txt. Si el programa antiguo se abría
solo al encender el PC, ese inicio pasa a abrir el programa nuevo (así no quedan los dos funcionando).
"""

from __future__ import annotations

import logging
import os
import re
import shutil
import sys
import tomllib
from pathlib import Path

log = logging.getLogger("cotizaciones")

BASE = Path(__file__).resolve().parent
ARCHIVOS = ("config.toml", "categorias.txt", "sinonimos.txt")
INICIO_ANTIGUO = "Buscador cotizaciones.cmd"  # lo creaba iniciar_con_windows.bat
PROFUNDIDAD = 3


def _es_programa(carpeta: Path) -> bool:
    return (carpeta / "buscador_cotizaciones.py").is_file() and (carpeta / "config.toml").is_file()


def _carpeta_configurada(carpeta: Path) -> str:
    try:
        cfg = tomllib.loads((carpeta / "config.toml").read_text(encoding="utf-8-sig"))
    except (OSError, tomllib.TOMLDecodeError):
        return ""
    valor = str(cfg.get("carpeta_cotizaciones") or "")
    return valor if valor and Path(valor).is_dir() else ""


def _carpeta_inicio() -> Path:
    return Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"


def _rutas_en_archivo(ruta: Path) -> list[Path]:
    """Carpetas del programa mencionadas en un acceso directo (.lnk) o en el .cmd de inicio."""
    try:
        datos = ruta.read_bytes()
    except OSError:
        return []
    textos = [datos.decode("utf-16-le", "ignore"), datos.decode("mbcs" if sys.platform == "win32" else "latin-1", "ignore")]
    rutas = []
    for texto in textos:
        for m in re.finditer(r"([A-Za-z]:\\[^\x00\"<>|*?\r\n]*?)\\(?:vigilar\.bat|ventana\.py|\.venv|"
                             r"abrir_excel\.bat|icono_cotizaciones\.ico)", texto):
            rutas.append(Path(m.group(1)))
    return rutas


def candidatos(inicio: Path | None = None, escritorio: Path | None = None, casa: Path | None = None) -> list[Path]:
    """Carpetas donde puede estar el programa antiguo (sin repetir y sin la del programa actual)."""
    casa = casa or Path.home()
    inicio = inicio or _carpeta_inicio()
    escritorios = [escritorio] if escritorio else [casa / "Desktop", casa / "Escritorio",
                                                   *casa.glob("OneDrive*/Desktop"), *casa.glob("OneDrive*/Escritorio")]
    vistas: list[Path] = []

    def agregar(p: Path) -> None:
        try:
            p = p.resolve()
        except OSError:
            return
        if p != BASE and p not in vistas and _es_programa(p):
            vistas.append(p)

    for p in _rutas_en_archivo(inicio / INICIO_ANTIGUO):
        agregar(p)
    for esc in escritorios:
        if esc.is_dir():
            for lnk in esc.glob("*otizaciones*.lnk"):
                for p in _rutas_en_archivo(lnk):
                    agregar(p)
    # buscar en las carpetas de siempre (Descargas, Documentos, Escritorio...), sin bajar demasiado
    raices = [casa / n for n in ("Downloads", "Descargas", "Documents", "Documentos", "Desktop", "Escritorio")]
    raices += [*casa.glob("OneDrive*")] + [casa]
    for raiz in raices:
        if not raiz.is_dir():
            continue
        base_prof = len(raiz.parts)
        for actual, dirs, archivos in os.walk(raiz):
            if len(Path(actual).parts) - base_prof >= PROFUNDIDAD or actual.endswith(("AppData", "Library")):
                dirs[:] = []
            dirs[:] = [d for d in dirs if not d.startswith((".", "$")) and d not in ("AppData", "node_modules")]
            if "buscador_cotizaciones.py" in archivos:
                agregar(Path(actual))
    return vistas


def traer_configuracion(destino: Path = BASE, **donde) -> Path | None:
    """Copia la configuración del programa antiguo a `destino`. Devuelve la carpeta de donde la trajo."""
    validos = [c for c in candidatos(**donde) if _carpeta_configurada(c)]
    if not validos:
        return None
    # el que se usó más recientemente (su Excel/registro se actualizó último)
    def uso(c: Path) -> float:
        try:
            return max(f.stat().st_mtime for f in (c / "registro.log", c / "config.toml") if f.exists())
        except ValueError:
            return 0.0
    origen = max(validos, key=uso)
    for nombre in ARCHIVOS:
        if (origen / nombre).is_file():
            shutil.copy2(origen / nombre, destino / nombre)
    log.info("Se trajo la configuración del programa anterior: %s", origen)
    _cambiar_inicio_antiguo(origen)
    return origen


def _cambiar_inicio_antiguo(origen: Path) -> None:
    """Si el programa antiguo se abría al encender el PC, ahora se abre el instalado."""
    cmd = _carpeta_inicio() / INICIO_ANTIGUO
    pythonw = BASE / "python" / "pythonw.exe"
    if not cmd.is_file() or not pythonw.exists() or str(origen).lower() not in cmd.read_text(errors="ignore").lower():
        return
    try:
        if (_carpeta_inicio() / "Buscador de cotizaciones.lnk").exists():  # el instalador ya puso su propio inicio
            cmd.unlink()
            log.info("Se quitó el inicio automático del programa anterior (el instalado ya tiene el suyo).")
            return
        cmd.write_text(f'@start "" /d "{BASE}" "{pythonw}" "{BASE / "ventana.py"}" --minimizado\r\n', encoding="mbcs"
                       if sys.platform == "win32" else "utf-8")
        log.info("El inicio automático ahora abre el programa instalado.")
    except OSError as e:
        log.warning("No se pudo cambiar el inicio automático antiguo: %s", e)
