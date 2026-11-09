"""Consola con colores: barra de progreso, línea de estado y aviso TERMINADO."""

from __future__ import annotations

import logging
import sys
import time
from pathlib import Path

ROJO, VERDE, AMARILLO, CELESTE, GRIS, BLANCO = "91", "92", "93", "96", "90", "97"
NEGRITA = "1"


def _activar_colores() -> bool:
    """Activa los colores ANSI (en Windows 10/11 hay que pedirlo a la consola)."""
    if sys.stdout is None or not hasattr(sys.stdout, "isatty") or not sys.stdout.isatty():
        return False
    if sys.platform != "win32":
        return True
    try:
        import ctypes

        k = ctypes.windll.kernel32
        h = k.GetStdHandle(-11)  # STD_OUTPUT_HANDLE
        modo = ctypes.c_uint32()
        if not k.GetConsoleMode(h, ctypes.byref(modo)):
            return False
        return bool(k.SetConsoleMode(h, modo.value | 0x0004))  # ENABLE_VIRTUAL_TERMINAL_PROCESSING
    except Exception:
        return False


COLOR = _activar_colores()
HAY_CONSOLA = sys.stdout is not None


def _bloques_ok() -> bool:
    try:
        "█░⟳✔".encode(sys.stdout.encoding or "utf-8")
        return True
    except (UnicodeEncodeError, LookupError, AttributeError):
        return False


BLOQUES = HAY_CONSOLA and _bloques_ok()
LLENO, VACIO = ("█", "░") if BLOQUES else ("#", "-")


def c(texto: str, *codigos: str) -> str:
    return f"\033[{';'.join(codigos)}m{texto}\033[0m" if COLOR and codigos else texto


def limpiar_linea() -> str:
    return "\r\033[K" if COLOR else "\r" + " " * 118 + "\r"


def escribir(texto: str, salto: bool = False) -> None:
    if not HAY_CONSOLA:
        return
    try:
        sys.stdout.write(limpiar_linea() + texto + ("\n" if salto else ""))
        sys.stdout.flush()
    except Exception:
        pass


class ManejadorConsola(logging.StreamHandler):
    """Escribe los mensajes sin pisar la barra ni la línea de estado (y en color)."""

    def emit(self, record: logging.LogRecord) -> None:
        if getattr(record, "solo_archivo", False):
            return
        try:
            msg = self.format(record)
            if record.levelno >= logging.ERROR:
                msg = c(msg, ROJO, NEGRITA)
            elif record.levelno >= logging.WARNING:
                msg = c(msg, AMARILLO)
            escribir(msg, salto=True)
        except Exception:
            self.handleError(record)


def _duracion(seg: float) -> str:
    if seg < 1.5:
        return "1 segundo"
    if seg < 90:
        return f"{seg:.0f} segundos"
    if seg < 5400:
        return f"{seg / 60:.0f} minutos"
    return f"{seg / 3600:.1f} horas"


class Barra:
    """[██████░░░░]  60%  12/20  quedan ~1 min   archivo.pdf   (rojo → amarillo → verde)"""

    ANCHO = 34

    def __init__(self, total: int):
        self.total, self.hechos, self.inicio = total, 0, time.monotonic()
        self.activa = HAY_CONSOLA and total > 0
        self.dibujar("")

    def avanzar(self, nombre: str) -> None:
        self.hechos += 1
        self.dibujar(nombre)

    def dibujar(self, nombre: str) -> None:
        if not self.activa:
            return
        pct = self.hechos / self.total
        lleno = int(self.ANCHO * pct)
        color = ROJO if pct < 0.34 else AMARILLO if pct < 0.67 else VERDE
        resto = ""
        if self.hechos and self.hechos < self.total:
            seg = (time.monotonic() - self.inicio) / self.hechos * (self.total - self.hechos)
            resto = f"quedan ~{seg:.0f} s" if seg < 90 else f"quedan ~{seg / 60:.0f} min"
        elif self.hechos == self.total:
            resto = "listo"
        texto = (
            c("[", GRIS) + c(LLENO * lleno, color) + c(VACIO * (self.ANCHO - lleno), GRIS) + c("]", GRIS)
            + " " + c(f"{pct:4.0%}", BLANCO, NEGRITA)
            + "  " + c(f"{self.hechos}/{self.total}", color)
            + "  " + c(f"{resto:<15}", CELESTE)
            + " " + c(Path(nombre).name[:32], GRIS)
        )
        escribir(texto)

    def cerrar(self) -> float:
        """Deja la barra en pantalla (100 %) y pasa a la línea siguiente."""
        seg = time.monotonic() - self.inicio
        if self.activa:
            self.hechos = self.total
            self.dibujar("")
            try:
                sys.stdout.write("\n")  # baja de línea sin borrar la barra
                sys.stdout.flush()
            except Exception:
                pass
        return seg


_LETRAS = {
    "T": ["█████", "  █  ", "  █  ", "  █  ", "  █  "],
    "E": ["█████", "█    ", "████ ", "█    ", "█████"],
    "R": ["████ ", "█   █", "████ ", "█  █ ", "█   █"],
    "M": ["█   █", "██ ██", "█ █ █", "█   █", "█   █"],
    "I": ["███", " █ ", " █ ", " █ ", "███"],
    "N": ["█   █", "██  █", "█ █ █", "█  ██", "█   █"],
    "A": [" ███ ", "█   █", "█████", "█   █", "█   █"],
    "D": ["████ ", "█   █", "█   █", "█   █", "████ "],
    "O": [" ███ ", "█   █", "█   █", "█   █", " ███ "],
}


def terminado(leidas: int, seg: float, total_excel: int, excel: Path) -> None:
    """TERMINADO en letras grandes rojas, con la cantidad de cotizaciones."""
    if not HAY_CONSOLA:
        return
    filas = ["  ".join(_LETRAS[l][i] for l in "TERMINADO") for i in range(5)]
    if not BLOQUES:
        filas = [f.replace("█", "#") for f in filas]
    marco = "═" * (len(filas[0]) + 6) if BLOQUES else "=" * (len(filas[0]) + 6)
    if leidas:
        detalle = f"{leidas} cotizaciones leídas en {_duracion(seg)}  ·  {total_excel} en total en el Excel"
    else:
        detalle = f"Sin cotizaciones nuevas  ·  {total_excel} cotizaciones en el Excel"
    lineas = ["", c(marco, ROJO, NEGRITA)] + [c("   " + f, ROJO, NEGRITA) for f in filas] + [
        "",
        c("   " + detalle, ROJO, NEGRITA),
        c(f"   Excel: {excel}", GRIS),
        c(marco, ROJO, NEGRITA),
        "",
    ]
    for l in lineas:
        escribir(l, salto=True)


_GIRO = "⠋⠙⠹⠸⠼⠴⠦⠧⠇⠏"


def estado_vigilancia(total_excel: int, ultima: str, tick: int) -> None:
    """Línea fija abajo mientras vigila: siempre muestra que el programa sigue trabajando."""
    giro = _GIRO[tick % len(_GIRO)] if BLOQUES else "|/-\\"[tick % 4]
    texto = (
        c("[", GRIS) + c(LLENO * Barra.ANCHO, VERDE) + c("]", GRIS) + " " + c("100%", BLANCO, NEGRITA)
        + "  " + c(f"{giro} Vigilando", CELESTE, NEGRITA)
        + c(f"  ·  {total_excel} cotizaciones en el Excel  ·  revisado {ultima}  ·  Ctrl+C para salir", GRIS)
    )
    escribir(texto)
