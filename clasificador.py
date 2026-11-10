"""Clasifica las cotizaciones por categoría (Cables, Alumbrado, ...) y guarda copias en carpetas.

Las categorías están en `categorias.txt` (junto al programa), una por línea:

    Cables: cable, conductor, thhn, awg

Una cotización va a una categoría si alguna de sus líneas (o el nombre del archivo)
contiene alguna de esas palabras. Puede ir a varias. Los PDF originales no se mueven:
se copian a  <carpeta de cotizaciones>/Clasificadas/<Categoría>/.
"""

from __future__ import annotations

import logging
import re
import shutil
from pathlib import Path

from extractor import normalizar

BASE = Path(__file__).resolve().parent
ARCHIVO = BASE / "categorias.txt"
CARPETA_CLASIFICADAS = "Clasificadas"

log = logging.getLogger("cotizaciones")

EJEMPLO = """\
# CATEGORÍAS DE COTIZACIONES
# Una categoría por línea:   Nombre: palabra1, palabra2, palabra3
# Una cotización se copia a la carpeta de la categoría si alguna de sus líneas
# (descripción o código) o el nombre del archivo contiene alguna de esas palabras.
# Mayúsculas y tildes dan igual. Puede quedar en varias categorías.
# Agrega, borra o cambia líneas cuando quieras, guarda (Ctrl + S) y el programa
# reclasifica todo solo. Las líneas que empiezan con # no cuentan.

Cables: cable, conductor, alambre, thhn, thwn, nyy, xtu, superflex, awg, mcm, multiconductor
Alumbrado: luminaria, foco, ampolleta, led, reflector, proyector, lampara, panel led, tubo fluorescente, aplique, iluminacion
Canalizacion: conduit, emt, ducto, bandeja, escalerilla, canaleta, cajas de paso, coplas, tubo pvc, tuberia
Tableros y protecciones: tablero, interruptor, automatico, diferencial, contactor, breaker, disyuntor, guardamotor, fusible
"""


def asegurar_archivo() -> Path:
    if not ARCHIVO.exists():
        ARCHIVO.write_text(EJEMPLO, encoding="utf-8")
    return ARCHIVO


def leer_categorias() -> dict[str, list[re.Pattern]]:
    """{'Cables': [patrones...], ...} desde categorias.txt."""
    asegurar_archivo()
    categorias: dict[str, list[re.Pattern]] = {}
    for linea in ARCHIVO.read_text(encoding="utf-8-sig", errors="replace").splitlines():
        linea = linea.strip()
        if not linea or linea.startswith("#") or ":" not in linea:
            continue
        nombre, palabras = linea.split(":", 1)
        nombre = re.sub(r'[<>:"/\\|?*]', "", nombre).strip()  # sirve como nombre de carpeta
        claves = [normalizar(p) for p in palabras.split(",") if normalizar(p)]
        if nombre and claves:
            # la palabra debe empezar ahí (\b): "led" no calza dentro de "soldadura"
            categorias[nombre] = [re.compile(r"\b" + re.escape(k)) for k in claves]
    return categorias


def _calza(texto: str, patrones: list[re.Pattern]) -> bool:
    return any(p.search(texto) for p in patrones)


def clasificar(datos: dict, clave: str, categorias: dict[str, list[re.Pattern]]) -> list[str]:
    """Marca la categoría de cada ítem y devuelve las categorías de la cotización."""
    encontradas: list[str] = []
    for it in datos.get("items", []):
        texto = normalizar(f"{it.get('descripcion', '')} {it.get('codigo', '')}")
        cats = [n for n, pats in categorias.items() if _calza(texto, pats)]
        it["categoria"] = ", ".join(cats)
        for c in cats:
            if c not in encontradas:
                encontradas.append(c)
    nombre = normalizar(Path(clave).stem)
    for n, pats in categorias.items():
        if n not in encontradas and _calza(nombre, pats):
            encontradas.append(n)
    datos["categorias"] = encontradas
    return encontradas


def _copiar(origen: Path, carpeta_cat: Path) -> tuple[Path, bool]:
    """Copia el PDF a la carpeta de la categoría. Devuelve (copia, True si se hizo ahora / False si ya estaba)."""
    carpeta_cat.mkdir(parents=True, exist_ok=True)
    destino = carpeta_cat / origen.name
    tam = origen.stat().st_size
    n = 2
    while destino.exists():
        if destino.stat().st_size == tam:
            return destino, False  # ya copiado
        destino = carpeta_cat / f"{origen.stem} ({n}){origen.suffix}"  # otro PDF con el mismo nombre
        n += 1
    shutil.copy2(origen, destino)
    return destino, True


def _relativa(carpeta: Path, ruta: Path) -> str:
    return ruta.relative_to(carpeta).as_posix()


def copiar_a_categorias(origen: Path, carpeta: Path, cats: list[str], clave: str = "") -> tuple[int, list[str]]:
    """Copia el PDF a cada categoría. Devuelve (copias nuevas, rutas de todas sus copias relativas a la carpeta)."""
    nuevas, rutas = 0, []
    if not origen.exists():
        return 0, rutas
    for cat in cats:
        try:
            destino, nueva = _copiar(origen, carpeta / CARPETA_CLASIFICADAS / cat)
            rutas.append(_relativa(carpeta, destino))
            nuevas += nueva
        except OSError as e:
            log.warning("No se pudo copiar %s a %s: %s", clave or origen.name, cat, e)
    return nuevas, rutas


def copias_existentes(origen: Path, carpeta: Path) -> list[str]:
    """Copias de este PDF que ya hay en Clasificadas (mismo nombre, o "nombre (2)", y mismo tamaño).

    Sirve para las cotizaciones clasificadas por versiones anteriores, que no anotaban sus copias."""
    raiz = carpeta / CARPETA_CLASIFICADAS
    if not origen.exists() or not raiz.is_dir():
        return []
    tam = origen.stat().st_size
    patron = re.compile(re.escape(origen.stem) + r"(?: \(\d+\))?" + re.escape(origen.suffix) + "$", re.I)
    rutas = []
    for sub in raiz.iterdir():
        if not sub.is_dir():
            continue
        for f in sub.iterdir():
            try:
                if patron.match(f.name) and f.is_file() and f.stat().st_size == tam:
                    rutas.append(_relativa(carpeta, f))
            except OSError:
                pass
    return rutas


def borrar_copias(carpeta: Path, rutas: list[str], conservar: set[str] | None = None) -> int:
    """Borra copias que ya no corresponden (solo dentro de Clasificadas). Quita las carpetas que queden vacías."""
    borradas = 0
    for r in rutas:
        if (conservar and r in conservar) or not es_clasificada(r):
            continue
        f = carpeta / r
        try:
            if f.is_file():
                f.unlink()
                borradas += 1
            if f.parent != carpeta / CARPETA_CLASIFICADAS and not any(f.parent.iterdir()):
                f.parent.rmdir()
        except OSError as e:
            log.warning("No se pudo quitar la copia %s: %s", r, e)
    return borradas


def clasificar_una(datos: dict, clave: str, carpeta: Path, categorias: dict[str, list[re.Pattern]]) -> tuple[int, list[str]]:
    """Clasifica una cotización recién leída y copia su PDF a sus carpetas.

    Devuelve (copias hechas ahora, rutas de sus copias) — las rutas se guardan para poder quitarlas después."""
    return copiar_a_categorias(carpeta / clave, carpeta, clasificar(datos, clave, categorias), clave)


def fecha_archivo() -> float:
    try:
        return ARCHIVO.stat().st_mtime
    except OSError:
        return 0.0


def clasificar_y_copiar(db: dict, carpeta: Path, cfg: dict, todo: bool = False) -> bool:
    """Clasifica las cotizaciones (todas si `todo`, si no solo las sin clasificar), copia los PDF y quita las
    copias que ya no corresponden (la cotización cambió de categoría).

    Devuelve True si cambió la categoría de alguna cotización (hay que rehacer el Excel).
    """
    if not cfg.get("clasificar", True):
        return False
    categorias = leer_categorias()
    copias, cambio = 0, False
    sobrantes: list[str] = []
    for clave, reg in db.items():
        datos = reg.get("datos", {})
        if not todo and "categorias" in datos and "copias" in reg:
            continue
        antes = datos.get("categorias")
        cats = clasificar(datos, clave, categorias)
        cambio = cambio or antes != cats
        origen = carpeta / clave
        anteriores = reg.get("copias")
        if anteriores is None:
            anteriores = copias_existentes(origen, carpeta)
        if not origen.exists():
            reg["copias"] = anteriores
            continue
        nuevas, rutas = copiar_a_categorias(origen, carpeta, cats, clave)
        copias += nuevas
        sobrantes += [r for r in anteriores if r not in rutas]
        reg["copias"] = rutas
    if copias:
        log.info("Copias nuevas en %s: %d", carpeta / CARPETA_CLASIFICADAS, copias)
    if sobrantes:
        en_uso = {r for reg in db.values() for r in reg.get("copias", [])}
        n = borrar_copias(carpeta, sobrantes, en_uso)
        if n:
            log.info("Copias quitadas de %s (la cotización ya no es de esa categoría): %d",
                     CARPETA_CLASIFICADAS, n)
    return cambio


def es_clasificada(clave: str) -> bool:
    """Las copias de Clasificadas/ no se leen como cotizaciones nuevas."""
    return clave.split("/", 1)[0] == CARPETA_CLASIFICADAS
