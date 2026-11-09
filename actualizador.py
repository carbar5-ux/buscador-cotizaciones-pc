"""Actualización automática desde GitHub.

Compara el archivo VERSION local con el del repositorio. Si hay una versión
nueva, descarga el repositorio y reemplaza solo los archivos del programa
(nunca config.toml, el Excel ni los datos).
"""

from __future__ import annotations

import io
import logging
import os
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
import zipfile
from pathlib import Path

REPO = "carbar5-ux/buscador-cotizaciones-pc"
RAMA = "main"
BASE = Path(__file__).resolve().parent
ARCHIVO_VERSION = BASE / "VERSION"

# Solo estos archivos se actualizan; config.toml y los datos del usuario nunca se tocan.
EXTENSIONES = {".py", ".bat", ".ico", ".md"}
EXTRA = {"VERSION", "requirements.txt"}
NUNCA = {"config.toml"}

log = logging.getLogger("cotizaciones")


def version_local() -> str:
    try:
        return ARCHIVO_VERSION.read_text(encoding="utf-8").strip()
    except OSError:
        return "0"


def _pedir(url: str, token: str, crudo: bool = False, timeout: int = 15) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": "buscador-cotizaciones-pc"})
    req.add_header("Accept", "application/vnd.github.raw" if crudo else "application/vnd.github+json")
    if token:
        req.add_header("Authorization", f"Bearer {token}")
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def version_remota(token: str) -> str:
    url = f"https://api.github.com/repos/{REPO}/contents/VERSION?ref={RAMA}"
    return _pedir(url, token, crudo=True).decode("utf-8").strip()


def _clave_version(v: str) -> tuple:
    return tuple(int(p) if p.isdigit() else 0 for p in v.replace("-", ".").split("."))


def buscar_actualizacion(token: str) -> str | None:
    """Devuelve la versión nueva disponible, o None si no hay (o no se pudo revisar)."""
    try:
        remota = version_remota(token)
    except urllib.error.HTTPError as e:
        if e.code in (401, 403, 404):
            log.warning("No se pudo buscar actualizaciones (el repositorio es privado o la clave no sirve). "
                        "Revisa 'token_github' en config.toml o el README.")
        else:
            log.info("No se pudo buscar actualizaciones (GitHub respondió %s).", e.code)
        return None
    except (urllib.error.URLError, TimeoutError, OSError):
        log.info("Sin conexión: no se buscaron actualizaciones.")
        return None
    return remota if _clave_version(remota) > _clave_version(version_local()) else None


def instalar(token: str) -> bool:
    """Descarga la versión nueva y reemplaza los archivos del programa. True si actualizó."""
    try:
        datos = _pedir(f"https://api.github.com/repos/{REPO}/zipball/{RAMA}", token, timeout=120)
    except Exception as e:
        log.warning("No se pudo descargar la actualización: %s", e)
        return False

    req_antes = (BASE / "requirements.txt").read_bytes() if (BASE / "requirements.txt").exists() else b""
    cambiados = []
    with zipfile.ZipFile(io.BytesIO(datos)) as z:
        for info in z.infolist():
            if info.is_dir():
                continue
            partes = Path(info.filename).parts[1:]  # quita la carpeta "dueño-repo-commit/"
            if not partes or ".." in partes or partes[0].startswith("."):
                continue
            rel = Path(*partes)
            if rel.name in NUNCA or not (rel.suffix.lower() in EXTENSIONES or rel.name in EXTRA):
                continue
            destino = BASE / rel
            contenido = z.read(info)
            if destino.exists() and destino.read_bytes() == contenido:
                continue
            destino.parent.mkdir(parents=True, exist_ok=True)
            fd, tmp = tempfile.mkstemp(dir=destino.parent, prefix="~act_")
            with os.fdopen(fd, "wb") as f:
                f.write(contenido)
            os.replace(tmp, destino)
            cambiados.append(str(rel))
    if not cambiados:
        return False
    log.info("Archivos actualizados: %s", ", ".join(cambiados))

    req_ahora = (BASE / "requirements.txt").read_bytes() if (BASE / "requirements.txt").exists() else b""
    if req_ahora != req_antes:
        log.info("Instalando librerías nuevas...")
        subprocess.run([sys.executable, "-m", "pip", "install", "-q", "-r", str(BASE / "requirements.txt")],
                       check=False)
    return True


def reiniciar() -> None:
    """Vuelve a abrir el programa (con la versión nueva) en la misma ventana y cierra este."""
    script = str(BASE / "buscador_cotizaciones.py")
    subprocess.Popen([sys.executable, script, *sys.argv[1:]], cwd=str(BASE))
    sys.exit(0)


def actualizar_si_hay(cfg: dict, reiniciar_al_terminar: bool = True) -> bool:
    if not cfg.get("actualizar_automatico", True):
        return False
    token = str(cfg.get("token_github") or "").strip()
    nueva = buscar_actualizacion(token)
    if not nueva:
        return False
    log.info("Hay una versión nueva (%s → %s). Actualizando...", version_local(), nueva)
    if instalar(token):
        log.info("Actualizado a la versión %s.", version_local())
        if reiniciar_al_terminar:
            log.info("Reiniciando el programa...")
            reiniciar()
        return True
    return False


if __name__ == "__main__":  # uso manual: python actualizador.py
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    sys.path.insert(0, str(BASE))
    from buscador_cotizaciones import leer_config

    cfg = leer_config()
    cfg["actualizar_automatico"] = True
    if not actualizar_si_hay(cfg, reiniciar_al_terminar=False):
        print(f"Versión instalada: {version_local()}")
