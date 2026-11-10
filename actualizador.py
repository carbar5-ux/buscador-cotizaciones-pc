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
EXTENSIONES = {".py", ".bat", ".ico", ".md", ".pdf"}
EXTRA = {"VERSION", "requirements.txt"}
NUNCA = {"config.toml"}
MAX_ZIP = 50 * 1024 * 1024  # tope de bytes al descargar/descomprimir una actualización

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
        datos = r.read(MAX_ZIP + 1)
    if len(datos) > MAX_ZIP:
        raise ValueError("la respuesta es demasiado grande")
    return datos


def version_remota(token: str) -> str:
    url = f"https://api.github.com/repos/{REPO}/contents/VERSION?ref={RAMA}"
    return _pedir(url, token, crudo=True).decode("utf-8").strip()


def _clave_version(v: str) -> tuple:
    return tuple(int(p) if p.isdigit() else 0 for p in v.replace("-", ".").split("."))


def consultar(token: str) -> tuple[str, str]:
    """("nueva", versión) si hay una versión nueva, ("al_dia", versión instalada) si no, o ("error", motivo)."""
    try:
        remota = version_remota(token)
    except urllib.error.HTTPError as e:
        if e.code in (401, 403, 404):
            return "error", ("No se pudo buscar actualizaciones (el repositorio es privado o la clave no sirve). "
                             "Revisa 'token_github' en config.toml o el README.")
        return "error", f"No se pudo buscar actualizaciones (GitHub respondió {e.code})."
    except (urllib.error.URLError, TimeoutError, OSError):
        return "error", "Sin conexión a internet: no se pudo buscar actualizaciones."
    if _clave_version(remota) > _clave_version(version_local()):
        return "nueva", remota
    return "al_dia", version_local()


def buscar_actualizacion(token: str) -> str | None:
    """Devuelve la versión nueva disponible, o None si no hay (o no se pudo revisar)."""
    estado, dato = consultar(token)
    if estado == "error":
        (log.warning if "privado" in dato else log.info)(dato)
        return None
    return dato if estado == "nueva" else None


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
            if not partes or ".." in partes or partes[0].startswith(".") or info.file_size > MAX_ZIP:
                continue
            rel = Path(*partes)
            if rel.is_absolute() or not (BASE / rel).resolve().is_relative_to(BASE):
                continue
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
    """Vuelve a abrir el programa (con la versión nueva) y cierra este.

    Sirve tanto para la consola como para la ventana: relanza el mismo script con
    los mismos argumentos y termina el proceso actual de inmediato.
    """
    script = Path(sys.argv[0]).resolve() if sys.argv and sys.argv[0].endswith(".py") \
        else BASE / "buscador_cotizaciones.py"
    subprocess.Popen([sys.executable, str(script), *sys.argv[1:]], cwd=str(BASE))
    logging.shutdown()
    os._exit(0)


# Archivos que una versión anterior del actualizador no sabía descargar: si faltan, se traen.
REQUERIDOS = ["Manual_de_uso.pdf"]


def asegurar_archivos(cfg: dict) -> None:
    token = str(cfg.get("token_github") or "").strip()
    for nombre in REQUERIDOS:
        destino = BASE / nombre
        if destino.exists():
            continue
        try:
            datos = _pedir(f"https://api.github.com/repos/{REPO}/contents/{nombre}?ref={RAMA}", token,
                           crudo=True, timeout=60)
        except Exception:
            continue  # sin internet: se intenta la próxima vez
        tmp = destino.with_name("~act_" + nombre)
        tmp.write_bytes(datos)
        os.replace(tmp, destino)
        log.info("Descargado: %s", nombre)


def actualizar_si_hay(cfg: dict, reiniciar_al_terminar: bool = True) -> bool:
    if not cfg.get("actualizar_automatico", True):
        return False
    asegurar_archivos(cfg)
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
