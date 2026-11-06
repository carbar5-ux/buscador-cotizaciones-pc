"""Buscador de cotizaciones PDF -> Excel base.

Uso:
    python buscador_cotizaciones.py              # procesa la carpeta una vez
    python buscador_cotizaciones.py --vigilar    # queda atento y actualiza el Excel al llegar PDFs nuevos
    python buscador_cotizaciones.py --reprocesar # vuelve a leer todos los PDF
    python buscador_cotizaciones.py --configurar # elegir otra carpeta

La configuración vive en config.toml (al lado de este archivo).
"""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import re
import sys
import time
from datetime import datetime
from pathlib import Path

if sys.version_info < (3, 11):
    sys.exit("Se necesita Python 3.11 o superior (https://www.python.org/downloads/).")

import tomllib  # noqa: E402

from excel_writer import construir_excel
from extractor import extraer_cotizacion

BASE = Path(__file__).resolve().parent
CONFIG = BASE / "config.toml"
VERSION_EXTRACTOR = 1  # subir si cambian las reglas de lectura: fuerza relectura

log = logging.getLogger("cotizaciones")

CONFIG_DEFECTO = {
    "carpeta_cotizaciones": "",
    "archivo_excel": "",
    "incluir_subcarpetas": True,
    "intervalo_segundos": 10,
    "formato_numeros": "chile",
    "dias_validez_por_defecto": 30,
    "dias_aviso_vencimiento": 7,
    "mis_ruts": [],
    "ocr": True,
    "quitar_borrados": True,
}


# --------------------------------------------------------------------------
# Configuración
# --------------------------------------------------------------------------


def _toml_valor(v) -> str:
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, (int, float)):
        return str(v)
    if isinstance(v, list):
        return "[" + ", ".join(_toml_valor(x) for x in v) + "]"
    return "'" + str(v).replace("'", "") + "'"  # literal: respeta las \ de Windows


_LINEA_RUTA = re.compile(r"^(\s*(?:carpeta_cotizaciones|archivo_excel)\s*=\s*)(.*?)\s*$")


def _reparar_rutas(texto: str) -> str:
    """Arregla las rutas mal escritas: sin comillas, comillas dobles con \\, comillas de más."""
    lineas = []
    for linea in texto.splitlines():
        m = _LINEA_RUTA.match(linea)
        if m:
            valor = m.group(2).strip().strip("'\"‘’“”").strip()
            linea = f"{m.group(1)}'{valor}'"
        lineas.append(linea)
    return "\n".join(lineas) + "\n"


def leer_config() -> dict:
    cfg = dict(CONFIG_DEFECTO)
    if not CONFIG.exists():
        return cfg
    texto = CONFIG.read_text(encoding="utf-8-sig")
    try:
        cfg.update(tomllib.loads(texto))
        return cfg
    except tomllib.TOMLDecodeError:
        pass
    reparado = _reparar_rutas(texto)
    try:
        cfg.update(tomllib.loads(reparado))
    except tomllib.TOMLDecodeError as e:
        sys.exit(
            f"\nHay un error en config.toml: {e}\n"
            "Revisa esa línea. Las rutas van entre comillas simples, por ejemplo:\n"
            "    carpeta_cotizaciones = 'C:\\Users\\Carlos\\Documents\\Cotizaciones'\n"
        )
    CONFIG.write_text(reparado, encoding="utf-8")
    print("Se corrigió el formato de la ruta en config.toml.")
    return cfg


def guardar_carpeta(cfg: dict, carpeta: Path) -> None:
    """Escribe la carpeta en config.toml conservando el resto del archivo."""
    cfg["carpeta_cotizaciones"] = str(carpeta)
    linea = f"carpeta_cotizaciones = {_toml_valor(str(carpeta))}"
    if CONFIG.exists():
        lineas = CONFIG.read_text(encoding="utf-8").splitlines()
        lineas = [linea if l.strip().startswith("carpeta_cotizaciones") else l for l in lineas]
        if linea not in lineas:
            lineas.insert(0, linea)
        CONFIG.write_text("\n".join(lineas) + "\n", encoding="utf-8")
    else:
        CONFIG.write_text("\n".join(f"{k} = {_toml_valor(v)}" for k, v in cfg.items()) + "\n", encoding="utf-8")


def elegir_carpeta() -> Path | None:
    try:
        import tkinter as tk
        from tkinter import filedialog

        raiz = tk.Tk()
        raiz.withdraw()
        raiz.attributes("-topmost", True)
        elegida = filedialog.askdirectory(title="Elige la carpeta donde guardas las cotizaciones PDF")
        raiz.destroy()
        if elegida:
            return Path(elegida)
    except Exception:
        pass
    try:
        texto = input("Ruta de la carpeta de cotizaciones: ").strip().strip('"')
    except (EOFError, RuntimeError):
        return None
    return Path(texto) if texto else None


# --------------------------------------------------------------------------
# Base de datos incremental (JSON junto al Excel)
# --------------------------------------------------------------------------


def hash_archivo(ruta: Path) -> str:
    h = hashlib.sha256()
    with ruta.open("rb") as f:
        for bloque in iter(lambda: f.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest()


def cargar_db(ruta: Path) -> dict:
    if ruta.exists():
        try:
            data = json.loads(ruta.read_text(encoding="utf-8"))
            if data.get("version") == VERSION_EXTRACTOR:
                return data["cotizaciones"]
            log.info("Las reglas de lectura cambiaron: se releerán todos los PDF.")
        except (json.JSONDecodeError, KeyError):
            log.warning("Base %s dañada: se reconstruye.", ruta.name)
    return {}


def guardar_db(ruta: Path, db: dict) -> None:
    tmp = ruta.with_suffix(".tmp")
    tmp.write_text(json.dumps({"version": VERSION_EXTRACTOR, "cotizaciones": db}, ensure_ascii=False, indent=1),
                   encoding="utf-8")
    tmp.replace(ruta)


def buscar_pdfs(carpeta: Path, recursivo: bool) -> list[Path]:
    patron = "**/*" if recursivo else "*"
    return sorted(
        p for p in carpeta.glob(patron)
        if p.is_file() and p.suffix.lower() == ".pdf" and not p.name.startswith(("~", "."))
    )


def escanear(cfg: dict, carpeta: Path, db: dict, reprocesar: bool = False) -> bool:
    """Lee los PDF nuevos o modificados. Devuelve True si algo cambió."""
    cambios = False
    pdfs = buscar_pdfs(carpeta, cfg["incluir_subcarpetas"])
    vigentes = set()
    por_hash = {reg["hash"]: clave for clave, reg in db.items()}
    ahora = time.time()

    for pdf in pdfs:
        clave = pdf.relative_to(carpeta).as_posix()
        vigentes.add(clave)
        try:
            st = pdf.stat()
        except OSError:
            continue
        if ahora - st.st_mtime < 3:  # se está copiando/guardando todavía: siguiente pasada
            continue
        reg = db.get(clave)
        if reg and not reprocesar and reg["tamano"] == st.st_size and reg["mtime"] == st.st_mtime:
            continue
        try:
            h = hash_archivo(pdf)
        except OSError as e:
            log.warning("No se pudo abrir %s (%s); se reintentará.", clave, e)
            continue
        if reg and reg["hash"] == h and not reprocesar:
            reg["tamano"], reg["mtime"] = st.st_size, st.st_mtime
            continue
        # archivo renombrado/movido: reutilizar lo ya leído
        anterior = por_hash.get(h)
        if anterior and anterior in db and not reprocesar and anterior != clave:
            datos = dict(db[anterior]["datos"], archivo=str(pdf))
            log.info("Renombrado/movido: %s -> %s", anterior, clave)
        else:
            log.info("Leyendo %s ...", clave)
            try:
                cot = extraer_cotizacion(
                    pdf, mis_ruts=set(cfg["mis_ruts"]), formato=cfg["formato_numeros"],
                    dias_validez=int(cfg["dias_validez_por_defecto"]), usar_ocr=bool(cfg["ocr"]),
                )
                datos = cot.to_dict()
                log.info("   %s | %s | N° %s | %s ítems | total %s %s",
                         cot.proveedor or "¿proveedor?", cot.fecha or "¿fecha?", cot.numero or "-",
                         len(cot.items), cot.moneda, f"{cot.total:,.0f}" if cot.total is not None else "-")
            except Exception as e:  # PDF dañado, protegido, etc.
                log.error("   Error leyendo %s: %s", clave, e)
                datos = {"archivo": str(pdf), "metodo": "error", "observaciones": [f"No se pudo leer el PDF: {e}"], "items": []}
        db[clave] = {"hash": h, "tamano": st.st_size, "mtime": st.st_mtime,
                     "procesado": datetime.now().isoformat(timespec="seconds"), "datos": datos}
        por_hash[h] = clave
        cambios = True

    if cfg["quitar_borrados"]:
        for clave in [c for c in db if c not in vigentes]:
            log.info("Ya no está en la carpeta: %s (se quita del Excel)", clave)
            del db[clave]
            cambios = True
    return cambios


# --------------------------------------------------------------------------
# Programa
# --------------------------------------------------------------------------


def actualizar_excel(cfg: dict, carpeta: Path, db: dict) -> bool:
    excel = Path(cfg["_excel"])
    try:
        construir_excel(db, carpeta, excel, int(cfg["dias_aviso_vencimiento"]))
        log.info("Excel actualizado: %s (%d cotizaciones)", excel, len(db))
        return True
    except PermissionError:
        log.warning("El Excel está abierto. Ciérralo y se actualizará en la próxima pasada.")
        return False


def main() -> int:
    ap = argparse.ArgumentParser(description="Lee cotizaciones PDF de una carpeta y arma un Excel base.")
    ap.add_argument("--vigilar", action="store_true", help="queda atento a PDFs nuevos y actualiza el Excel")
    ap.add_argument("--reprocesar", action="store_true", help="vuelve a leer todos los PDF")
    ap.add_argument("--configurar", action="store_true", help="elegir la carpeta de cotizaciones")
    ap.add_argument("--carpeta", type=Path, help="carpeta de cotizaciones (no la guarda en config)")
    ap.add_argument("--excel", type=Path, help="ruta del Excel de salida")
    args = ap.parse_args()

    manejadores: list[logging.Handler] = [logging.FileHandler(BASE / "registro.log", encoding="utf-8")]
    if sys.stdout is not None:  # con pythonw (segundo plano) no hay consola
        manejadores.append(logging.StreamHandler(sys.stdout))
    logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(message)s", datefmt="%d-%m %H:%M:%S",
                        handlers=manejadores)

    cfg = leer_config()
    carpeta = args.carpeta or (Path(cfg["carpeta_cotizaciones"]) if cfg["carpeta_cotizaciones"] else None)
    if args.configurar or carpeta is None or not carpeta.is_dir():
        if carpeta is not None and not args.configurar:
            log.warning("No existe la carpeta configurada: %s", carpeta)
        carpeta = elegir_carpeta()
        if not carpeta or not carpeta.is_dir():
            log.error("No se eligió una carpeta válida.")
            return 1
        guardar_carpeta(cfg, carpeta)
        log.info("Carpeta guardada en config.toml: %s", carpeta)

    carpeta = carpeta.resolve()
    excel = args.excel or (Path(cfg["archivo_excel"]) if cfg["archivo_excel"] else carpeta / "Base_Cotizaciones.xlsx")
    cfg["_excel"] = str(excel.resolve())
    archivo_db = excel.with_name(excel.stem + "_datos.json")

    db = cargar_db(archivo_db)
    log.info("Carpeta: %s", carpeta)
    cambios = escanear(cfg, carpeta, db, args.reprocesar) or not excel.exists()
    guardar_db(archivo_db, db)
    pendiente = cambios and not actualizar_excel(cfg, carpeta, db)
    if not cambios:
        log.info("Sin cotizaciones nuevas. Excel al día: %s", excel)

    if not args.vigilar:
        return 0

    intervalo = max(2, int(cfg["intervalo_segundos"]))
    log.info("Vigilando la carpeta cada %d s. Guarda cotizaciones y el Excel se actualizará solo. (Ctrl+C para salir)",
             intervalo)
    try:
        while True:
            time.sleep(intervalo)
            if not carpeta.is_dir():
                log.warning("La carpeta no está disponible; reintentando...")
                continue
            if escanear(cfg, carpeta, db):
                guardar_db(archivo_db, db)
                pendiente = True
            if pendiente:
                pendiente = not actualizar_excel(cfg, carpeta, db)
    except KeyboardInterrupt:
        log.info("Vigilancia detenida.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
