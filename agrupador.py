"""Reconoce el mismo producto escrito de distintas formas, para comparar precios.

"Cable THHN 12 AWG rojo" y "Conductor THHN N°12 AWG negro" son el mismo producto;
"Cable THHN 12 AWG" y "Cable THHN 14 AWG" NO (las medidas deben coincidir exactas).

Cómo:
  1. Firma de cada descripción: sus medidas (números) + sus palabras importantes, sin
     tildes, sin colores ni palabras de relleno, y con sinónimos unificados.
  2. Mismo producto = mismas medidas, misma unidad, misma moneda y palabras casi iguales
     (al menos 60 % en común).

Los sinónimos se editan en `sinonimos.txt` (junto al programa).
"""

from __future__ import annotations

import re
from collections import Counter, defaultdict
from pathlib import Path

from extractor import sin_tildes

BASE = Path(__file__).resolve().parent
ARCHIVO_SINONIMOS = BASE / "sinonimos.txt"
SIMILITUD_MINIMA = 0.6

EJEMPLO_SINONIMOS = """\
# SINÓNIMOS PARA COMPARAR PRECIOS
# Una línea por palabra:   palabra_principal: sinónimo1, sinónimo2
# Así "conductor" y "cable" cuentan como la misma palabra al comparar productos.
# Mayúsculas y tildes dan igual. Guarda (Ctrl + S) y el Excel se rehace en la próxima revisión.

cable: conductor, cables, conductores
lampara: ampolleta, ampolletas, lamparas, bombilla, foco, focos
luminaria: luminarias, equipo de iluminacion
canaleta: canaletas, moldura, molduras
tubo: tubos, tuberia, caneria
interruptor: interruptores, automatico, automaticos, breaker
enchufe: enchufes, tomacorriente, tomacorrientes
caja: cajas
saco: sacos, bolsa
plancha: planchas, placa, placas
"""

# Palabras que no cambian el producto (colores, relleno, marcas genéricas)
RELLENO = {
    "de", "del", "la", "el", "los", "las", "y", "o", "con", "sin", "para", "por", "en", "a", "al", "tipo", "modelo",
    "marca", "color", "un", "una", "x", "n", "no", "nro", "num", "cod", "codigo", "ref", "aprox", "incluye", "inc",
    "rojo", "roja", "negro", "negra", "azul", "blanco", "blanca", "verde", "amarillo", "amarilla", "gris", "cafe",
    "naranjo", "naranja", "celeste", "violeta", "transparente", "rollo", "rollos",
}
UNIDADES = {
    "un": "un", "und": "un", "uds": "un", "ud": "un", "unid": "un", "unidad": "un", "unidades": "un", "c/u": "un",
    "cu": "un", "pza": "un", "pzas": "un", "pc": "un", "pcs": "un", "ea": "un",
    "m": "m", "mt": "m", "mts": "m", "ml": "m", "metro": "m", "metros": "m", "mtr": "m",
    "kg": "kg", "kgs": "kg", "kilo": "kg", "kilos": "kg",
    "gl": "gl", "glb": "gl", "global": "gl",
    "m2": "m2", "m²": "m2", "m3": "m3", "m³": "m3",
    "lt": "lt", "lts": "lt", "l": "lt", "litro": "lt", "litros": "lt",
    "hr": "hr", "hrs": "hr", "h": "hr", "hh": "hh",
}

_NUM = re.compile(r"\d+(?:[.,]\d+)?")
_PALABRA = re.compile(r"[a-z]+")

_sinonimos_cache: tuple[float, dict[str, str]] | None = None


def _simple(s: str) -> str:
    s = sin_tildes(str(s or "")).lower().replace("²", "2").replace("³", "3")
    return s


def asegurar_sinonimos() -> Path:
    if not ARCHIVO_SINONIMOS.exists():
        ARCHIVO_SINONIMOS.write_text(EJEMPLO_SINONIMOS, encoding="utf-8")
    return ARCHIVO_SINONIMOS


def sinonimos() -> dict[str, str]:
    """{'conductor': 'cable', ...} desde sinonimos.txt (se relee solo si cambió)."""
    global _sinonimos_cache
    asegurar_sinonimos()
    try:
        fecha = ARCHIVO_SINONIMOS.stat().st_mtime
    except OSError:
        fecha = 0.0
    if _sinonimos_cache and _sinonimos_cache[0] == fecha:
        return _sinonimos_cache[1]
    tabla: dict[str, str] = {}
    for linea in ARCHIVO_SINONIMOS.read_text(encoding="utf-8-sig", errors="replace").splitlines():
        linea = linea.strip()
        if not linea or linea.startswith("#") or ":" not in linea:
            continue
        principal, otros = linea.split(":", 1)
        principal = " ".join(_PALABRA.findall(_simple(principal)))
        if not principal:
            continue
        for o in otros.split(","):
            o = " ".join(_PALABRA.findall(_simple(o)))
            if o:
                tabla[o] = principal
    _sinonimos_cache = (fecha, tabla)
    return tabla


def numeros(texto: str) -> frozenset[str]:
    """Medidas de la descripción: '3x2,5 mm2 12 AWG' -> {'3', '2.5', '12'}."""
    t = _simple(texto)
    t = re.sub(r"\bm[23]\b", " ", t)  # 'm2'/'m3' son unidades, no medidas
    salida = set()
    for n in _NUM.findall(t):
        n = n.replace(",", ".")
        if "." in n:
            n = n.rstrip("0").rstrip(".")
        salida.add(n.lstrip("0") or "0")
    return frozenset(salida)


def palabras(texto: str, sin: dict[str, str] | None = None) -> frozenset[str]:
    """Palabras importantes, con sinónimos unificados (también de dos palabras: 'equipo de iluminacion')."""
    sin = sinonimos() if sin is None else sin
    t = " ".join(_PALABRA.findall(_simple(texto)))
    for frase, principal in sin.items():
        if " " in frase and frase in t:
            t = t.replace(frase, principal)
    salida = set()
    for p in t.split():
        p = sin.get(p, p)
        if len(p) >= 2 and p not in RELLENO:
            salida.add(p)
    return frozenset(salida)


def unidad_normal(u: str) -> str:
    u = _simple(u).strip(". ")
    return UNIDADES.get(u, u)


def similitud(a: frozenset[str], b: frozenset[str]) -> float:
    if not a and not b:
        return 1.0
    return len(a & b) / len(a | b)


def agrupar(elementos: list[tuple[str, str, str, object]]) -> list[list[object]]:
    """Agrupa elementos (descripción, unidad, moneda, dato) que son el mismo producto.

    Devuelve listas de `dato`. Solo se comparan entre sí los que tienen las mismas medidas,
    unidad y moneda; dentro de eso, se juntan si sus palabras se parecen lo suficiente.
    """
    sin = sinonimos()
    cubetas: dict[tuple, list[tuple[frozenset[str], object]]] = defaultdict(list)
    for desc, unidad, moneda, dato in elementos:
        cubetas[(numeros(desc), unidad_normal(unidad), moneda)].append((palabras(desc, sin), dato))
    grupos: list[list[object]] = []
    for lista in cubetas.values():
        locales: list[tuple[Counter, list[object]]] = []  # (palabras del grupo, datos)
        for pals, dato in lista:
            mejor, mejor_sim = None, 0.0
            for i, (cont, _) in enumerate(locales):
                rep = frozenset(p for p, n in cont.items() if n * 2 >= max(cont.values()))  # palabras frecuentes
                s = similitud(pals, rep)
                if s > mejor_sim:
                    mejor, mejor_sim = i, s
            if mejor is not None and mejor_sim >= SIMILITUD_MINIMA:
                locales[mejor][0].update(pals)
                locales[mejor][1].append(dato)
            else:
                locales.append((Counter(pals), [dato]))
        grupos.extend(datos for _, datos in locales)
    return grupos


def descripcion_representativa(descripciones: list[str]) -> str:
    """La descripción más repetida del grupo (si empatan, la más corta)."""
    cuenta = Counter(d.strip() for d in descripciones if d and d.strip())
    if not cuenta:
        return ""
    return sorted(cuenta.items(), key=lambda t: (-t[1], len(t[0])))[0][0]


def calza_busqueda(consulta: str, descripcion: str, sin: dict[str, str] | None = None) -> bool:
    """¿La descripción tiene todas las palabras y medidas buscadas? ('cable thhn 12')."""
    sin = sinonimos() if sin is None else sin
    nums_q = numeros(consulta)
    if nums_q and not nums_q <= numeros(descripcion):
        return False
    pals_d = palabras(descripcion, sin) | {p for p in _PALABRA.findall(_simple(descripcion))}
    for q in palabras(consulta, sin) | ({p for p in _PALABRA.findall(_simple(consulta))} - RELLENO):
        q = sin.get(q, q)
        if len(q) < 2:
            continue
        if not any(p.startswith(q) or sin.get(p, p).startswith(q) for p in pals_d):
            return False
    return True


def preparar(descripcion: str, codigo: str = "", sin: dict[str, str] | None = None) -> tuple[frozenset, frozenset]:
    """Medidas y palabras de una línea, calculadas una vez para buscar rápido."""
    sin = sinonimos() if sin is None else sin
    texto = f"{descripcion} {codigo}"
    crudas = set(_PALABRA.findall(_simple(texto)))
    return numeros(texto), frozenset(palabras(texto, sin) | crudas | {sin.get(p, p) for p in crudas})


def preparar_consulta(consulta: str, sin: dict[str, str] | None = None) -> tuple[frozenset, list[str]]:
    sin = sinonimos() if sin is None else sin
    crudas = [p for p in _PALABRA.findall(_simple(consulta)) if len(p) >= 2 and p not in RELLENO]
    return numeros(consulta), [sin.get(p, p) for p in crudas]


def calza_preparado(consulta: tuple[frozenset, list[str]], linea: tuple[frozenset, frozenset]) -> bool:
    nums_q, pals_q = consulta
    nums_d, pals_d = linea
    if nums_q and not nums_q <= nums_d:
        return False
    return all(any(p.startswith(q) for p in pals_d) for q in pals_q)
