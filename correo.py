"""Baja a la carpeta de cotizaciones los PDF que llegan adjuntos por correo.

Se conecta por IMAP (Gmail, Outlook, el correo de la empresa...) y revisa los correos de los últimos días.
Solo baja los PDF adjuntos cuyo asunto, remitente o nombre de archivo contiene alguna de las palabras
de `correo_palabras` (cotización, presupuesto...). Los guarda en <carpeta de cotizaciones>/Correo/ y desde
ahí el buscador los lee como cualquier otra cotización. Nunca borra ni marca como leído ningún correo.

Para Gmail se usa una "contraseña de aplicación" (no la clave normal): ver el README.
"""

from __future__ import annotations

import email
import email.policy
import imaplib
import json
import logging
import re
import ssl
import threading
import time
from datetime import date, timedelta
from email.message import EmailMessage
from email.utils import parseaddr, parsedate_to_datetime
from pathlib import Path

from extractor import normalizar

log = logging.getLogger("cotizaciones")

CARPETA_CORREO = "Correo"
SERVIDORES = {  # el servidor se adivina por el correo si no se indica
    "gmail.com": "imap.gmail.com", "googlemail.com": "imap.gmail.com",
    "outlook.com": "outlook.office365.com", "hotmail.com": "outlook.office365.com",
    "live.com": "outlook.office365.com", "office365.com": "outlook.office365.com",
    "yahoo.com": "imap.mail.yahoo.com", "icloud.com": "imap.mail.me.com",
}
MAX_CORREOS = 400  # por revisión (la primera vez puede haber muchos)

_en_curso = threading.Lock()


class ErrorCorreo(Exception):
    """Problema para entrar al correo (clave, servidor...), con un mensaje para mostrar."""


def servidor_de(cfg: dict) -> str:
    if cfg.get("correo_servidor"):
        return str(cfg["correo_servidor"]).strip()
    dominio = str(cfg.get("correo_usuario", "")).rpartition("@")[2].lower()
    return SERVIDORES.get(dominio, f"imap.{dominio}" if dominio else "")


def configurado(cfg: dict) -> bool:
    return bool(cfg.get("correo_usuario") and cfg.get("correo_clave"))


def _nombre_seguro(texto: str, largo: int = 80) -> str:
    texto = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", texto or "")
    return re.sub(r"\s+", " ", texto).strip(" .")[:largo] or "adjunto"


def _calza(texto: str, palabras: list[str]) -> bool:
    if not palabras:
        return True
    t = normalizar(texto)
    return any(normalizar(p) and normalizar(p) in t for p in palabras)


def pdfs_del_mensaje(crudo: bytes, palabras: list[str]) -> list[tuple[str, bytes]]:
    """[(nombre de archivo sugerido, contenido)] de los PDF adjuntos que interesan."""
    msg = email.message_from_bytes(crudo, policy=email.policy.default)
    assert isinstance(msg, EmailMessage)
    asunto = str(msg.get("Subject", "") or "")
    nombre_rem, dir_rem = parseaddr(str(msg.get("From", "") or ""))
    try:
        fecha = parsedate_to_datetime(str(msg.get("Date"))).date().isoformat()
    except (TypeError, ValueError):
        fecha = date.today().isoformat()
    remitente = nombre_rem or dir_rem.split("@")[0]
    resultado = []
    for parte in msg.walk():
        if parte.is_multipart():
            continue
        nombre = parte.get_filename() or ""
        es_pdf = parte.get_content_type() == "application/pdf" or nombre.lower().endswith(".pdf")
        if not es_pdf:
            continue
        if not _calza(f"{asunto} {nombre} {remitente} {dir_rem}", palabras):
            continue
        contenido = parte.get_payload(decode=True)
        if not contenido or not contenido.startswith(b"%PDF"):
            continue
        base = nombre[:-4] if nombre.lower().endswith(".pdf") else (nombre or asunto)
        resultado.append((f"{fecha} - {_nombre_seguro(remitente, 40)} - {_nombre_seguro(base)}.pdf", contenido))
    return resultado


def guardar_pdf(destino_dir: Path, nombre: str, contenido: bytes) -> Path | None:
    """Guarda el PDF; si ya hay uno igual con ese nombre no hace nada (None)."""
    destino_dir.mkdir(parents=True, exist_ok=True)
    destino = destino_dir / nombre
    n = 2
    while destino.exists():
        if destino.read_bytes() == contenido:
            return None
        destino = destino_dir / f"{Path(nombre).stem} ({n}).pdf"
        n += 1
    tmp = destino.with_name("~" + destino.name)  # el buscador ignora los que empiezan con ~ mientras se escribe
    tmp.write_bytes(contenido)
    tmp.replace(destino)
    return destino


def conectar(cfg: dict) -> imaplib.IMAP4_SSL:
    servidor = servidor_de(cfg)
    if not servidor:
        raise ErrorCorreo("Falta el correo (correo_usuario).")
    try:
        imap = imaplib.IMAP4_SSL(servidor, 993, ssl_context=ssl.create_default_context(), timeout=40)
    except OSError as e:
        raise ErrorCorreo(f"No se pudo conectar con {servidor}: {e}") from e
    try:
        imap.login(str(cfg["correo_usuario"]).strip(), str(cfg["correo_clave"]).replace(" ", ""))
    except imaplib.IMAP4.error as e:
        try:
            imap.logout()
        except Exception:
            pass
        raise ErrorCorreo("El correo o la clave no son correctos. En Gmail se usa una \"contraseña de aplicación\" "
                          f"de 16 letras, no la clave normal. ({e})") from e
    return imap


def probar(cfg: dict) -> str:
    """Prueba la conexión. Devuelve un texto para mostrar; lanza ErrorCorreo si no se puede entrar."""
    imap = conectar(cfg)
    try:
        estado, datos = imap.select(str(cfg.get("correo_carpeta") or "INBOX"), readonly=True)
        if estado != "OK":
            raise ErrorCorreo(f"No existe la carpeta de correo {cfg.get('correo_carpeta')!r}.")
        return f"Conexión correcta con {servidor_de(cfg)}: {int(datos[0] or 0):,} correos en la bandeja.".replace(",", ".")
    finally:
        try:
            imap.logout()
        except Exception:
            pass


def _cargar_estado(ruta: Path) -> dict:
    try:
        return json.loads(ruta.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"vistos": {}}


def revisar(cfg: dict, carpeta: Path, archivo_estado: Path) -> int:
    """Baja los PDF nuevos. Devuelve cuántos guardó. Lanza ErrorCorreo si no puede entrar."""
    if not configurado(cfg):
        return 0
    if not _en_curso.acquire(blocking=False):
        return 0  # ya hay una revisión andando
    try:
        return _revisar(cfg, carpeta, archivo_estado)
    finally:
        _en_curso.release()


def _revisar(cfg: dict, carpeta: Path, archivo_estado: Path) -> int:
    estado = _cargar_estado(archivo_estado)
    vistos: dict[str, str] = estado.setdefault("vistos", {})
    dias = max(1, int(cfg.get("correo_dias", 30) or 30))
    palabras = [str(p) for p in cfg.get("correo_palabras", []) or []]
    destino = carpeta / CARPETA_CORREO if cfg.get("incluir_subcarpetas", True) else carpeta
    imap = conectar(cfg)
    guardados = 0
    try:
        if imap.select(str(cfg.get("correo_carpeta") or "INBOX"), readonly=True)[0] != "OK":
            raise ErrorCorreo(f"No existe la carpeta de correo {cfg.get('correo_carpeta')!r}.")
        if "gmail" in servidor_de(cfg):  # Gmail filtra en el servidor: solo correos con PDF adjunto
            r, datos = imap.uid("SEARCH", "X-GM-RAW", f'"has:attachment filename:pdf newer_than:{dias}d"')
        else:
            desde = (date.today() - timedelta(days=dias)).strftime("%d-%b-%Y")
            r, datos = imap.uid("SEARCH", None, "SINCE", desde)
        uids = (datos[0] or b"").split() if r == "OK" else []
        clave_buzon = f"{servidor_de(cfg)}|{cfg.get('correo_usuario')}|{cfg.get('correo_carpeta') or 'INBOX'}"
        validez = (imap.response("UIDVALIDITY")[1] or [b""])[0]
        validez = validez.decode() if isinstance(validez, bytes) else str(validez or "")
        prefijo = f"{clave_buzon}|{validez}|"
        nuevos = [u for u in uids if prefijo + u.decode() not in vistos][-MAX_CORREOS:]
        for uid in nuevos:
            r, partes = imap.uid("FETCH", uid, "(BODY.PEEK[])")  # PEEK: no lo marca como leído
            crudo = next((p[1] for p in partes or [] if isinstance(p, tuple) and len(p) > 1), None)
            if r != "OK" or not crudo:
                continue
            for nombre, contenido in pdfs_del_mensaje(crudo, palabras):
                try:
                    ruta = guardar_pdf(destino, nombre, contenido)
                except OSError as e:
                    log.warning("No se pudo guardar el adjunto %s: %s", nombre, e)
                    continue
                if ruta:
                    guardados += 1
                    log.info("Correo: cotización recibida -> %s", ruta.name)
            vistos[prefijo + uid.decode()] = date.today().isoformat()
    finally:
        try:
            imap.logout()
        except Exception:
            pass
        # se olvidan los correos más viejos que la ventana revisada (el archivo no crece sin fin)
        limite = (date.today() - timedelta(days=dias + 30)).isoformat()
        estado["vistos"] = {k: v for k, v in vistos.items() if v >= limite}
        estado["ultima_revision"] = time.strftime("%Y-%m-%d %H:%M:%S")
        try:
            archivo_estado.write_text(json.dumps(estado, ensure_ascii=False, indent=0), encoding="utf-8")
        except OSError:
            pass
    return guardados
