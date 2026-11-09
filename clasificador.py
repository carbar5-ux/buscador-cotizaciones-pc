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


def _copiar(origen: Path, carpeta_cat: Path) -> bool:
    """Copia el PDF a la carpeta de la categoría (si ya está igual, no hace nada)."""
    carpeta_cat.mkdir(parents=True, exist_ok=True)
    destino = carpeta_cat / origen.name
    tam = origen.stat().st_size
    n = 2
    while destino.exists():
        if destino.stat().st_size == tam:
            return False  # ya copiado
        destino = carpeta_cat / f"{origen.stem} ({n}){origen.suffix}"  # otro PDF con el mismo nombre
        n += 1
    shutil.copy2(origen, destino)
    return True


def fecha_archivo() -> float:
    try:
        return ARCHIVO.stat().st_mtime
    except OSError:
        return 0.0


def clasificar_y_copiar(db: dict, carpeta: Path, cfg: dict, todo: bool = False) -> bool:
    """Clasifica las cotizaciones (todas si `todo`, si no solo las sin clasificar) y copia los PDF.

    Devuelve True si cambió la categoría de alguna cotización (hay que rehacer el Excel).
    """
    if not cfg.get("clasificar", True):
        return False
    categorias = leer_categorias()
    raiz = carpeta / CARPETA_CLASIFICADAS
    copias, cambio = 0, False
    for clave, reg in db.items():
        datos = reg.get("datos", {})
        if not todo and "categorias" in datos:
            continue
        antes = datos.get("categorias")
        cats = clasificar(datos, clave, categorias)
        cambio = cambio or antes != cats
        origen = carpeta / clave
        if not origen.exists():
            continue
        for cat in cats:
            try:
                if _copiar(origen, raiz / cat):
                    copias += 1
            except OSError as e:
                log.warning("No se pudo copiar %s a %s: %s", clave, cat, e)
    if copias:
        log.info("Copias nuevas en %s: %d", raiz, copias)
    return cambio


def es_clasificada(clave: str) -> bool:
    """Las copias de Clasificadas/ no se leen como cotizaciones nuevas."""
    return clave.split("/", 1)[0] == CARPETA_CLASIFICADAS
