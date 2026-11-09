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
import os
import re
import sys
import threading
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

if sys.version_info < (3, 11):
    sys.exit("Se necesita Python 3.11 o superior (https://www.python.org/downloads/).")

import tomllib  # noqa: E402

from actualizador import actualizar_si_hay, version_local
from consola import Barra, ManejadorConsola, _duracion, estado_vigilancia, terminado
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
    "procesos": 0,
    "aviso_ventana": True,
    "aviso_sonido": True,
    "avisar_desde": 5,
    "actualizar_automatico": True,
    "horas_entre_actualizaciones": 6,
    "token_github": "",
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


def _prioridad_baja() -> None:
    """Los procesos lectores usan todo el procesador libre sin trabar el PC."""
    try:
        if sys.platform == "win32":
            import ctypes

            BELOW_NORMAL_PRIORITY_CLASS = 0x4000
            ctypes.windll.kernel32.SetPriorityClass(ctypes.windll.kernel32.GetCurrentProcess(),
                                                    BELOW_NORMAL_PRIORITY_CLASS)
        else:
            os.nice(5)
    except Exception:
        pass


def _leer_pdf(ruta: str, opciones: dict) -> dict:
    """Lee un PDF (se ejecuta en un proceso aparte para aprovechar todos los núcleos)."""
    try:
        return extraer_cotizacion(Path(ruta), **opciones).to_dict()
    except Exception as e:  # PDF dañado, protegido, etc.
        return {"archivo": ruta, "metodo": "error", "observaciones": [f"No se pudo leer el PDF: {e}"], "items": []}


def _resumen(datos: dict) -> str:
    total = datos.get("total")
    return " | ".join([
        datos.get("proveedor") or "¿proveedor?", datos.get("fecha") or "¿fecha?", f"N° {datos.get('numero') or '-'}",
        f"{len(datos.get('items', []))} ítems",
        f"total {datos.get('moneda', '')} {total:,.0f}" if total is not None else "total -",
    ])


def _titulo_consola(texto: str) -> None:
    if sys.platform == "win32":
        try:
            import ctypes

            ctypes.windll.kernel32.SetConsoleTitleW(texto)
        except Exception:
            pass


def avisar_fin(cfg: dict, excel: Path, minimo: int, esperar: bool) -> None:
    """Sonido y ventana emergente al terminar de leer cotizaciones."""
    leidas, seg = cfg.pop("_ultima_lectura", (0, 0.0))
    if not leidas:
        return
    _titulo_consola(f"Buscador de cotizaciones - Terminado: {leidas} leídas")
    if sys.platform != "win32":
        return
    if cfg.get("aviso_sonido", True):
        try:
            import winsound

            winsound.MessageBeep(winsound.MB_ICONASTERISK)
        except Exception:
            pass
    if not cfg.get("aviso_ventana", True) or leidas < minimo:
        return

    def ventana() -> None:
        try:
            import ctypes

            mensaje = (f"Terminó la lectura de cotizaciones.\n\n"
                       f"Cotizaciones leídas: {leidas}\nTiempo: {_duracion(seg)}\n\n"
                       f"Excel: {excel}\n\n¿Abrir el Excel ahora?")
            # MB_YESNO | MB_ICONINFORMATION | MB_SETFOREGROUND | MB_TOPMOST
            r = ctypes.windll.user32.MessageBoxW(0, mensaje, "Buscador de cotizaciones", 0x4 | 0x40 | 0x10000 | 0x40000)
            if r == 6 and excel.exists():  # Sí
                os.startfile(excel)
        except Exception as e:
            log.warning("No se pudo mostrar el aviso: %s", e)

    if esperar:
        ventana()
    else:  # en modo vigilancia el programa sigue trabajando mientras el aviso está abierto
        threading.Thread(target=ventana, daemon=True).start()


def escanear(cfg: dict, carpeta: Path, db: dict, reprocesar: bool = False, guardar=None) -> bool:
    """Lee los PDF nuevos o modificados. Devuelve True si algo cambió."""
    cambios = False
    pdfs = buscar_pdfs(carpeta, cfg["incluir_subcarpetas"])
    vigentes = set()
    por_hash = {reg["hash"]: clave for clave, reg in db.items()}
    ahora = time.time()
    pendientes: list[tuple[str, Path, os.stat_result, str]] = []

    # 1) qué hay que leer (rápido: solo fecha y tamaño; huella solo si cambió)
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
        anterior = por_hash.get(h)
        if anterior and anterior in db and not reprocesar and anterior != clave:
            # archivo renombrado/movido: reutilizar lo ya leído
            db[clave] = dict(db[anterior], tamano=st.st_size, mtime=st.st_mtime,
                             datos=dict(db[anterior]["datos"], archivo=str(pdf)))
            log.info("Renombrado/movido: %s -> %s", anterior, clave)
            cambios = True
            continue
        pendientes.append((clave, pdf, st, h))

    # 2) leer los PDF pendientes, varios a la vez, con barra de progreso
    if pendientes:
        log.info("Cotizaciones por leer: %d", len(pendientes))
        opciones = {"mis_ruts": set(cfg["mis_ruts"]), "formato": cfg["formato_numeros"],
                    "dias_validez": int(cfg["dias_validez_por_defecto"]), "usar_ocr": bool(cfg["ocr"])}
        nucleos = int(cfg.get("procesos") or 0) or (os.cpu_count() or 2)  # todos los núcleos
        nucleos = min(nucleos, len(pendientes))
        log.info("Usando %d núcleo(s) del procesador.", nucleos)
        barra = Barra(len(pendientes))
        ultimo_guardado = time.monotonic()

        def registrar(clave: str, st: os.stat_result, h: str, datos: dict) -> None:
            db[clave] = {"hash": h, "tamano": st.st_size, "mtime": st.st_mtime,
                         "procesado": datetime.now().isoformat(timespec="seconds"), "datos": datos}
            por_hash[h] = clave
            detalle = f"Error en {clave}: {datos['observaciones'][0]}" if datos.get("metodo") == "error" \
                else f"{clave}  ->  {_resumen(datos)}"
            log.info(detalle, extra={"solo_archivo": barra.activa})
            barra.avanzar(clave)
            nonlocal ultimo_guardado
            if guardar and time.monotonic() - ultimo_guardado > 60:
                guardar()  # avance guardado cada minuto: si se corta, no parte de cero
                ultimo_guardado = time.monotonic()

        if nucleos == 1:
            for clave, pdf, st, h in pendientes:
                registrar(clave, st, h, _leer_pdf(str(pdf), opciones))
        else:
            with ProcessPoolExecutor(max_workers=nucleos, initializer=_prioridad_baja) as pool:
                tareas = {pool.submit(_leer_pdf, str(pdf), opciones): (clave, st, h) for clave, pdf, st, h in pendientes}
                for fut in as_completed(tareas):
                    clave, st, h = tareas[fut]
                    try:
                        datos = fut.result()
                    except Exception as e:  # el proceso que leía ese PDF se cayó
                        datos = {"archivo": clave, "metodo": "error", "observaciones": [f"No se pudo leer el PDF: {e}"], "items": []}
                    registrar(clave, st, h, datos)
        seg = barra.cerrar()
        log.info("Leídas %d cotizaciones en %s.", len(pendientes), _duracion(seg), extra={"solo_archivo": barra.activa})
        cfg["_ultima_lectura"] = (len(pendientes), seg)
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


def actualizar_excel(cfg: dict, carpeta: Path, db: dict, en_barra: bool = False) -> bool:
    excel = Path(cfg["_excel"])
    try:
        construir_excel(db, carpeta, excel, int(cfg["dias_aviso_vencimiento"]))
        log.info("Excel actualizado: %s (%d cotizaciones)", excel, len(db), extra={"solo_archivo": en_barra})
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
    ap.add_argument("--abrir", action="store_true", help="abre el Excel base")
    args = ap.parse_args()

    manejadores: list[logging.Handler] = [logging.FileHandler(BASE / "registro.log", encoding="utf-8")]
    if sys.stdout is not None:  # con pythonw (segundo plano) no hay consola
        # en color y sin pisar la barra; el detalle de cada PDF va solo al registro
        manejadores.append(ManejadorConsola(sys.stdout))
    logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(message)s", datefmt="%d-%m %H:%M:%S",
                        handlers=manejadores)

    cfg = leer_config()
    log.info("Buscador de cotizaciones, versión %s", version_local())
    if not args.abrir:
        actualizar_si_hay(cfg)  # si hay versión nueva, se instala y el programa se reinicia solo
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

    if args.abrir:
        if not excel.exists():
            log.error("Todavía no existe el Excel: %s\nEjecuta primero vigilar.bat o procesar.bat.", excel)
            return 1
        log.info("Abriendo %s", excel)
        if hasattr(os, "startfile"):
            os.startfile(excel)  # Windows
        return 0

    db = cargar_db(archivo_db)
    log.info("Carpeta de cotizaciones: %s", carpeta)
    log.info("=" * 70)
    log.info("EL EXCEL SE GUARDA EN:  %s", excel)
    log.info("(para abrirlo directo usa  abrir_excel.bat)")
    log.info("=" * 70)
    if not excel.exists():
        actualizar_excel(cfg, carpeta, db)  # se crea de inmediato y se va llenando

    def guardar_avance() -> None:
        guardar_db(archivo_db, db)
        actualizar_excel(cfg, carpeta, db, en_barra=True)

    cambios = escanear(cfg, carpeta, db, args.reprocesar, guardar_avance)
    guardar_db(archivo_db, db)
    pendiente = cambios and not actualizar_excel(cfg, carpeta, db)
    if not cambios:
        log.info("Sin cotizaciones nuevas. Excel al día: %s", excel)
    _titulo_consola("Buscador de cotizaciones - " + ("vigilando" if args.vigilar else "terminado"))
    leidas, seg = cfg.get("_ultima_lectura", (0, 0.0))
    terminado(leidas, seg, len(db), excel)
    avisar_fin(cfg, excel, minimo=1, esperar=not args.vigilar)  # primera lectura: siempre avisa

    if not args.vigilar:
        return 0

    intervalo = max(2, int(cfg["intervalo_segundos"]))
    log.info("Vigilando la carpeta cada %d s. Guarda cotizaciones y el Excel se actualizará solo. (Ctrl+C para salir)",
             intervalo)
    tick, proxima, ultima = 0, time.monotonic() + intervalo, datetime.now().strftime("%H:%M:%S")
    cada_act = max(1.0, float(cfg["horas_entre_actualizaciones"])) * 3600
    proxima_act = time.monotonic() + cada_act
    try:
        while True:
            estado_vigilancia(len(db), ultima, tick)  # línea siempre presente
            time.sleep(1)
            tick += 1
            if time.monotonic() > proxima_act:
                proxima_act = time.monotonic() + cada_act
                actualizar_si_hay(cfg)  # se reinicia solo si encontró una versión nueva
            if time.monotonic() < proxima:
                continue
            proxima = time.monotonic() + intervalo
            ultima = datetime.now().strftime("%H:%M:%S")
            if not carpeta.is_dir():
                log.warning("La carpeta no está disponible; reintentando...")
                continue
            if escanear(cfg, carpeta, db, guardar=guardar_avance):
                guardar_db(archivo_db, db)
                pendiente = True
            if pendiente:
                pendiente = not actualizar_excel(cfg, carpeta, db)
            if "_ultima_lectura" in cfg:
                leidas, seg = cfg["_ultima_lectura"]
                terminado(leidas, seg, len(db), excel)
            # cotizaciones sueltas: solo sonido; tandas grandes: también la ventana
            avisar_fin(cfg, excel, minimo=int(cfg["avisar_desde"]), esperar=False)
    except KeyboardInterrupt:
        log.info("Vigilancia detenida.")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        import traceback

        traceback.print_exc()
        logging.getLogger("cotizaciones").exception("Error inesperado")
        if sys.stdin is not None and sys.stdin.isatty():
            input("\nOcurrió un error (queda anotado en registro.log). Presiona Enter para cerrar...")
        sys.exit(1)
