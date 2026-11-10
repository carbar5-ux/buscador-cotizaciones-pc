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
import socket
import sys
import threading
import time
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

if sys.version_info < (3, 11):
    sys.exit("Se necesita Python 3.11 o superior (https://www.python.org/downloads/).")

import shutil  # noqa: E402
import tomllib  # noqa: E402
from datetime import date  # noqa: E402

from actualizador import actualizar_si_hay, version_local
from agrupador import agrupar, asegurar_sinonimos, variaciones
from clasificador import (borrar_copias, clasificar_una, clasificar_y_copiar, es_clasificada, fecha_archivo,
                          leer_categorias)
import correo  # noqa: E402
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
    "max_paginas_tablas": 10,
    "clasificar": True,
    "aviso_ventana": True,
    "aviso_sonido": True,
    "avisar_desde": 5,
    "actualizar_automatico": True,
    "horas_entre_actualizaciones": 6,
    "token_github": "",
    "respaldos_dias": 14,
    "correo_usuario": "",
    "correo_clave": "",
    "correo_servidor": "",
    "correo_carpeta": "INBOX",
    "correo_palabras": ["cotiz", "presupuesto", "proforma", "oferta"],
    "correo_dias": 30,
    "correo_minutos": 10,
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


def guardar_valores(cfg: dict, valores: dict) -> None:
    """Escribe valores en config.toml conservando el resto del archivo (comentarios incluidos)."""
    cfg.update(valores)
    lineas = CONFIG.read_text(encoding="utf-8-sig").splitlines() if CONFIG.exists() else []
    for clave, valor in valores.items():
        linea = f"{clave} = {_toml_valor(valor)}"
        patron = re.compile(rf"^\s*{re.escape(clave)}\s*=")
        for i, l in enumerate(lineas):
            if patron.match(l):
                lineas[i] = linea
                break
        else:
            lineas.append(linea)
    CONFIG.write_text("\n".join(lineas) + "\n", encoding="utf-8")


def guardar_carpeta(cfg: dict, carpeta: Path) -> None:
    """Escribe la carpeta en config.toml conservando el resto del archivo."""
    guardar_valores(cfg, {"carpeta_cotizaciones": str(carpeta)})


def problema_carpeta(carpeta: Path) -> str:
    """Por qué una carpeta NO sirve como carpeta de cotizaciones ("" si sirve).

    Un disco completo (C:\\) o la carpeta de usuario harían leer todos los PDF del PC, y en la raíz del disco
    Windows no deja guardar el Excel."""
    try:
        c = carpeta.resolve()
    except OSError:
        return ""
    if c == Path(c.anchor) or len(c.parts) <= 1:
        return (f"Elegiste el disco completo ({c}). Elige la carpeta donde están tus cotizaciones, por ejemplo "
                "Documentos\\Cotizaciones (entra a la carpeta y pulsa Seleccionar carpeta).")
    if c == Path.home().resolve():
        return ("Elegiste tu carpeta de usuario completa. Elige la carpeta donde están tus cotizaciones, por ejemplo "
                "Documentos\\Cotizaciones.")
    if c == BASE or (len(c.parts) == 2 and c.name.lower() in ("windows", "program files", "program files (x86)",
                                                              "programdata", "users", "usuarios")):
        return f"La carpeta {c} no es una carpeta de cotizaciones. Elige la carpeta donde guardas tus PDF."
    return ""


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


def _version_guardada(ruta: Path) -> str:
    """Versión del programa que escribió por última vez los datos (y el Excel)."""
    try:
        with ruta.open(encoding="utf-8") as f:
            inicio = f.read(300)
        m = re.search(r'"programa":\s*"([^"]+)"', inicio)
        return m.group(1) if m else ""
    except OSError:
        return ""


def guardar_db(ruta: Path, db: dict, programa: str | None = None) -> None:
    """programa: versión que hizo el Excel actual. Solo pasa a la versión nueva cuando el Excel se rehízo de
    verdad (si estaba abierto y no se pudo, la próxima vez se vuelve a intentar)."""
    tmp = ruta.with_suffix(".tmp")
    programa = version_local() if programa is None else programa
    tmp.write_text(json.dumps({"version": VERSION_EXTRACTOR, "programa": programa, "cotizaciones": db},
                              ensure_ascii=False, indent=1),
                   encoding="utf-8")
    tmp.replace(ruta)


def buscar_pdfs(carpeta: Path, recursivo: bool) -> list[Path]:
    patron = "**/*" if recursivo else "*"
    return sorted(
        p for p in carpeta.glob(patron)
        if p.is_file() and p.suffix.lower() == ".pdf" and not p.name.startswith(("~", "."))
    )


def _prioridad_baja() -> None:
    """Prepara cada proceso lector: sin mensajes en pantalla y con prioridad baja.

    En Windows cada proceso lector parte de cero; sin esto, cualquier aviso de las
    librerías de PDF saldría directo a la consola y taparía la barra.
    """
    logging.getLogger().addHandler(logging.NullHandler())
    logging.getLogger().setLevel(logging.ERROR)
    import warnings

    warnings.simplefilter("ignore")
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
    inicio = time.monotonic()
    try:
        datos = extraer_cotizacion(Path(ruta), **opciones).to_dict()
    except Exception as e:  # PDF dañado, protegido, etc.
        datos = {"archivo": ruta, "metodo": "error", "observaciones": [f"No se pudo leer el PDF: {e}"], "items": []}
    datos["_segundos"] = round(time.monotonic() - inicio, 1)
    return datos


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
    if not cfg.get("aviso_ventana", True) or leidas < minimo or cfg.get("_ventana"):
        return  # con la ventana abierta, el aviso se muestra en ella

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
    por_revisar: list[tuple[str, Path, os.stat_result]] = []
    for pdf in pdfs:
        clave = pdf.relative_to(carpeta).as_posix()
        if es_clasificada(clave):  # copias hechas por el clasificador: no son cotizaciones nuevas
            continue
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
        por_revisar.append((clave, pdf, st))

    # huellas de varios archivos a la vez (con cientos de PDF nuevos esto ahorra bastante)
    def _huella(pdf: Path):
        try:
            return hash_archivo(pdf)
        except OSError as e:
            return e

    with ThreadPoolExecutor(max_workers=8) as hilos:
        huellas = list(hilos.map(_huella, [p for _, p, _ in por_revisar]))

    for (clave, pdf, st), h in zip(por_revisar, huellas):
        reg = db.get(clave)
        if isinstance(h, Exception):
            log.warning("No se pudo abrir %s (%s); se reintentará.", clave, h)
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
                    "dias_validez": int(cfg["dias_validez_por_defecto"]), "usar_ocr": bool(cfg["ocr"]),
                    "max_paginas_tablas": int(cfg["max_paginas_tablas"])}
        nucleos = int(cfg.get("procesos") or 0) or (os.cpu_count() or 2)  # todos los núcleos
        nucleos = min(nucleos, len(pendientes))
        log.info("Usando %d núcleo(s) del procesador.", nucleos)
        categorias = leer_categorias() if cfg.get("clasificar", True) else None
        copias = 0
        barra = Barra(len(pendientes))
        ultimo_guardado = time.monotonic()

        def registrar(clave: str, st: os.stat_result, h: str, datos: dict) -> None:
            nonlocal copias
            anteriores = db.get(clave, {}).get("copias")
            db[clave] = {"hash": h, "tamano": st.st_size, "mtime": st.st_mtime,
                         "procesado": datetime.now().isoformat(timespec="seconds"), "datos": datos}
            if categorias is not None:  # se copia a su carpeta apenas se lee
                nuevas, rutas = clasificar_una(datos, clave, carpeta, categorias)
                copias += nuevas
                db[clave]["copias"] = rutas
                if anteriores:  # el PDF cambió: quitar copias que ya no corresponden
                    borrar_copias(carpeta, [r for r in anteriores if r not in rutas])
            por_hash[h] = clave
            cfg.setdefault("_nuevas", []).append(clave)  # para avisar alzas de precio
            detalle = f"Error en {clave}: {datos['observaciones'][0]}" if datos.get("metodo") == "error" \
                else f"{clave}  ->  {_resumen(datos)}"
            seg_pdf = datos.pop("_segundos", 0)
            if seg_pdf >= 20:
                detalle += f"   (lento: {seg_pdf:.0f} s, {datos.get('paginas', '?')} páginas)"
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
        if copias:
            log.info("Copias a las carpetas de categoría en esta lectura: %d", copias)
        log.info("Leídas %d cotizaciones en %s.", len(pendientes), _duracion(seg), extra={"solo_archivo": barra.activa})
        cfg["_ultima_lectura"] = (len(pendientes), seg)
        cambios = True

    if cfg["quitar_borrados"]:
        sin_original: list[str] = []
        for clave in [c for c in db if c not in vigentes]:
            log.info("Ya no está en la carpeta: %s (se quita del Excel)", clave)
            sin_original += db[clave].get("copias", [])
            del db[clave]
            cambios = True
        if sin_original:  # sus copias en Clasificadas también sobran (salvo que otra cotización las use)
            en_uso = {r for reg in db.values() for r in reg.get("copias", [])}
            n = borrar_copias(carpeta, sin_original, en_uso)
            if n:
                log.info("Copias quitadas de Clasificadas (se borró la cotización original): %d", n)
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
    except PermissionError as e:
        log.warning("No se pudo guardar el Excel (%s). Puede estar abierto, bloqueado por OneDrive/antivirus, o la "
                    "carpeta no permite escribir. Se reintentará en la próxima pasada. Detalle: %s", excel, e)
        return False
    except Exception:  # nunca detener la lectura por un problema al escribir el Excel
        log.exception("No se pudo escribir el Excel (se reintentará). Los datos leídos quedan guardados.")
        return False


def fecha_sinonimos() -> float:
    try:
        return asegurar_sinonimos().stat().st_mtime
    except OSError:
        return 0.0


def _limpiar_temporales(*carpetas: Path) -> None:
    """Borra temporales que quedan si el programa se cerró a mitad de un guardado o una actualización
    (~tmp_*.xlsx junto al Excel, *_datos.tmp, ~act_* del actualizador). Solo los de más de 10 minutos."""
    limite = time.time() - 600
    for carpeta in carpetas:
        for patron in ("~tmp_*", "*_datos.tmp", "~act_*"):
            try:
                for f in carpeta.glob(patron):
                    try:
                        if f.is_file() and f.stat().st_mtime < limite:
                            f.unlink()
                    except OSError:
                        pass
            except OSError:
                pass


def respaldar(excel: Path, archivo_db: Path, guardar: int) -> bool:
    """Una copia diaria del Excel y de los datos leídos en <carpeta del Excel>/Respaldos (se guardan los últimos
    `guardar` días). Si algo se daña o se borra, basta copiar el respaldo de vuelta. True si respaldó hoy."""
    if guardar <= 0 or not archivo_db.exists():
        return False
    carpeta = excel.parent / "Respaldos"
    hoy = date.today().isoformat()
    destino_db = carpeta / f"{archivo_db.stem}_{hoy}.json"
    if destino_db.exists():
        return False
    try:
        carpeta.mkdir(exist_ok=True)
        if excel.exists():
            shutil.copy2(excel, carpeta / f"{excel.stem}_{hoy}{excel.suffix}")
        shutil.copy2(archivo_db, destino_db)
    except OSError as e:  # Excel abierto, disco lleno...: se reintenta en la próxima pasada
        log.warning("No se pudo hacer el respaldo diario: %s", e)
        destino_db.unlink(missing_ok=True)
        return False
    for base, ext in ((excel.stem, excel.suffix), (archivo_db.stem, ".json")):
        viejos = sorted(carpeta.glob(f"{base}_????-??-??{ext}"))[:-guardar]
        for f in viejos:
            try:
                f.unlink()
            except OSError:
                pass
    log.info("Respaldo del día guardado en %s", carpeta)
    return True


ALZA_AVISO = 0.05  # desde 5 % de alza se avisa


def _miles(v: float) -> str:
    return f"{v:,.0f}".replace(",", ".") if v >= 100 else f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def avisar_alzas(cfg: dict, db: dict, maximo: int = 15) -> list[str]:
    """Avisa (en el registro y en la ventana) los productos que subieron de precio en las cotizaciones recién
    leídas, frente a la cotización anterior del mismo producto. Devuelve los avisos."""
    nuevas = set(cfg.pop("_nuevas", []))
    if not nuevas:
        return []
    elementos = []
    for clave, reg in list(db.items()):
        d = reg.get("datos", {})
        for it in d.get("items", []):
            if it.get("precio_unitario") and it.get("descripcion"):
                elementos.append((it["descripcion"], it.get("unidad", ""), d.get("moneda", ""), (it, d, clave)))
    var = variaciones(agrupar(elementos), fecha=lambda t: t[1].get("fecha") or "",
                      precio=lambda t: t[0]["precio_unitario"], cotizacion=lambda t: t[2])
    avisos = []
    for _, _, _, t in elementos:
        v = var.get(id(t))
        if t[2] not in nuevas or not v or v[0] < ALZA_AVISO:
            continue
        (it, d, _), (it_a, d_a, _) = t, v[1]
        mon = d.get("moneda", "")
        avisos.append(f"ALZA DE PRECIO {v[0]:+.0%}: {it['descripcion'][:60]} — {d.get('proveedor') or '?'} "
                      f"{mon} {_miles(it['precio_unitario'])} (antes {_miles(it_a['precio_unitario'])}, "
                      f"{d_a.get('proveedor') or '?'}, {d_a.get('fecha') or '?'})")
    if len(nuevas) > 30:  # carga grande (primera vez): un resumen en vez de una lista larga
        if avisos:
            log.info("Hay %d alza(s) de precio entre las cotizaciones leídas: mira la columna \"Var. vs anterior\" "
                     "de la hoja Precios (en rojo).", len(avisos))
        return avisos
    for a in avisos[:maximo]:
        log.warning(a)
    if len(avisos) > maximo:
        log.warning("... y %d alza(s) más: mira la columna \"Var. vs anterior\" de la hoja Precios.", len(avisos) - maximo)
    return avisos


def revisar_correo(cfg: dict, carpeta: Path, archivo_estado: Path) -> None:
    """Baja las cotizaciones del correo en segundo plano (no detiene la lectura de la carpeta)."""
    if not correo.configurado(cfg) or cfg.get("_correo_error"):
        return

    def tarea() -> None:
        try:
            n = correo.revisar(cfg, carpeta, archivo_estado)
            if n:
                log.info("Correo: %d cotización(es) nueva(s) guardada(s) en la carpeta %s.", n, correo.CARPETA_CORREO)
        except correo.ErrorCorreo as e:
            cfg["_correo_error"] = str(e)  # no insistir cada 10 minutos con una clave mala
            log.warning("Correo: %s  (corrígelo con el botón Correo de la ventana)", e)
        except Exception as e:
            log.warning("Correo: no se pudo revisar (%s); se reintentará.", e)

    threading.Thread(target=tarea, daemon=True).start()


def candado_instancia():
    """Evita dos copias vigilando a la vez (escribirían el mismo Excel). None si ya hay otra abierta."""
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", 47653))
        s.listen(1)
        s.setblocking(False)
        return s
    except OSError:
        s.close()
        return None


def _piden_la_ventana(candado: socket.socket) -> bool:
    """La ventana pidió el control (otra copia quiere abrirse)."""
    try:
        conn, _ = candado.accept()
    except (BlockingIOError, OSError):
        return False
    try:
        conn.sendall(b"consola")
        conn.close()
    except OSError:
        pass
    candado.close()
    return True


def main(argv: list[str] | None = None, detener: threading.Event | None = None, ventana: bool = False,
         compartido: dict | None = None) -> int:
    ap = argparse.ArgumentParser(description="Lee cotizaciones PDF de una carpeta y arma un Excel base.")
    ap.add_argument("--vigilar", action="store_true", help="queda atento a PDFs nuevos y actualiza el Excel")
    ap.add_argument("--reprocesar", action="store_true", help="vuelve a leer todos los PDF")
    ap.add_argument("--configurar", action="store_true", help="elegir la carpeta de cotizaciones")
    ap.add_argument("--carpeta", type=Path, help="carpeta de cotizaciones (no la guarda en config)")
    ap.add_argument("--excel", type=Path, help="ruta del Excel de salida")
    ap.add_argument("--abrir", action="store_true", help="abre el Excel base")
    args = ap.parse_args(argv)

    manejadores: list[logging.Handler] = [logging.FileHandler(BASE / "registro.log", encoding="utf-8")]
    if sys.stdout is not None and not ventana:  # con pythonw (segundo plano) no hay consola
        # en color y sin pisar la barra; el detalle de cada PDF va solo al registro
        manejadores.append(ManejadorConsola(sys.stdout))
    logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(message)s", datefmt="%d-%m %H:%M:%S",
                        handlers=manejadores)

    cfg = leer_config()
    cfg["_ventana"] = ventana
    log.info("Buscador de cotizaciones, versión %s", version_local())
    if not args.abrir:
        actualizar_si_hay(cfg)  # si hay versión nueva, se instala y el programa se reinicia solo
    carpeta = args.carpeta or (Path(cfg["carpeta_cotizaciones"]) if cfg["carpeta_cotizaciones"] else None)
    if args.configurar or carpeta is None or not carpeta.is_dir() or problema_carpeta(carpeta):
        if carpeta is not None and not args.configurar:
            log.warning(problema_carpeta(carpeta) if carpeta.is_dir() else f"No existe la carpeta configurada: {carpeta}")
        carpeta = elegir_carpeta()
        if not carpeta or not carpeta.is_dir() or problema_carpeta(carpeta):
            log.error(problema_carpeta(carpeta) if carpeta and carpeta.is_dir() else "No se eligió una carpeta válida.")
            return 1
        guardar_carpeta(cfg, carpeta)
        log.info("Carpeta guardada en config.toml: %s", carpeta)

    carpeta = carpeta.resolve()
    excel = args.excel or (Path(cfg["archivo_excel"]) if cfg["archivo_excel"] else carpeta / "Base_Cotizaciones.xlsx")
    cfg["_excel"] = str(excel.resolve())
    archivo_db = excel.with_name(excel.stem + "_datos.json")
    archivo_correo = excel.with_name(excel.stem + "_correo.json")
    _limpiar_temporales(excel.parent, BASE)

    if args.abrir:
        if not excel.exists():
            log.error("Todavía no existe el Excel: %s\nEjecuta primero vigilar.bat o procesar.bat.", excel)
            return 1
        log.info("Abriendo %s", excel)
        if hasattr(os, "startfile"):
            os.startfile(excel)  # Windows
        return 0

    _guardar_db, _actualizar_excel = globals()["guardar_db"], globals()["actualizar_excel"]
    db = cargar_db(archivo_db)
    version_excel = [_version_guardada(archivo_db)]  # versión que hizo el Excel que hay ahora
    excel_de_otra_version = version_excel[0] != version_local()  # nuevo diseño: rehacer el Excel

    def guardar_db(ruta: Path, datos: dict) -> None:
        _guardar_db(ruta, datos, version_excel[0])

    def actualizar_excel(*a, **k) -> bool:
        hecho = _actualizar_excel(*a, **k)
        if hecho and version_excel[0] != version_local():
            version_excel[0] = version_local()
            _guardar_db(archivo_db, db, version_excel[0])
        return hecho
    if compartido is not None:  # la ventana usa estos datos para el buscador de precios
        compartido.update(db=db, carpeta=carpeta)
    log.info("Carpeta de cotizaciones: %s", carpeta)
    log.info("=" * 70)
    log.info("EL EXCEL SE GUARDA EN:  %s", excel)
    log.info("(para abrirlo directo usa  abrir_excel.bat)")
    log.info("=" * 70)
    if not excel.exists():
        actualizar_excel(cfg, carpeta, db)  # se crea de inmediato y se va llenando

    ultimo_excel = [time.monotonic()]

    def guardar_avance() -> None:
        """Cada minuto se guardan los datos leídos; el Excel cada 3 minutos (rehacerlo pausa la lectura)."""
        guardar_db(archivo_db, db)
        if time.monotonic() - ultimo_excel[0] >= 180:
            actualizar_excel(cfg, carpeta, db, en_barra=True)
            ultimo_excel[0] = time.monotonic()

    if args.vigilar:
        revisar_correo(cfg, carpeta, archivo_correo)  # en paralelo: lo que llegue se lee en la siguiente pasada
    cambios = escanear(cfg, carpeta, db, args.reprocesar, guardar_avance) or excel_de_otra_version
    fecha_cat = fecha_archivo()
    fecha_sin = fecha_sinonimos()
    cambios = clasificar_y_copiar(db, carpeta, cfg, todo=True) or cambios  # por si cambiaron las categorías
    guardar_db(archivo_db, db)
    pendiente = cambios and not actualizar_excel(cfg, carpeta, db)
    avisar_alzas(cfg, db)
    respaldar(excel, archivo_db, int(cfg["respaldos_dias"]))
    if not cambios:
        log.info("Sin cotizaciones nuevas. Excel al día: %s", excel)
    _titulo_consola("Buscador de cotizaciones - " + ("vigilando" if args.vigilar else "terminado"))
    leidas, seg = cfg.get("_ultima_lectura", (0, 0.0))
    terminado(leidas, seg, len(db), excel)
    avisar_fin(cfg, excel, minimo=1, esperar=not args.vigilar)  # primera lectura: siempre avisa

    if not args.vigilar:
        return 0

    intervalo = max(2, int(cfg["intervalo_segundos"]))
    log.info("Vigilando la carpeta cada %d s. Guarda cotizaciones y el Excel se actualizará solo.%s",
             intervalo, "" if ventana else " (Ctrl+C para salir)")
    tick, proxima, ultima = 0, time.monotonic() + intervalo, datetime.now().strftime("%H:%M:%S")
    candado = candado_instancia() if not ventana else True  # la ventana toma su propio candado
    if candado is None:
        log.error("El buscador ya está abierto en este PC (quizás en segundo plano). Se cierra esta copia.")
        return 0
    cada_act = max(1.0, float(cfg["horas_entre_actualizaciones"])) * 3600
    proxima_act = time.monotonic() + cada_act
    cada_correo = max(2.0, float(cfg["correo_minutos"])) * 60
    proximo_correo = time.monotonic() + cada_correo
    try:
        while not (detener and detener.is_set()):
            estado_vigilancia(len(db), ultima, tick)  # línea siempre presente
            time.sleep(1)
            tick += 1
            if isinstance(candado, socket.socket) and _piden_la_ventana(candado):
                log.info("Se abrió la ventana del buscador: esta copia en segundo plano se cierra.")
                return 0
            if detener is not None and getattr(detener, "revisar_ya", False):
                detener.revisar_ya = False
                proxima = 0  # botón "Revisar ahora" de la ventana
            if time.monotonic() > proxima_act:
                proxima_act = time.monotonic() + cada_act
                actualizar_si_hay(cfg)  # se reinicia solo si encontró una versión nueva
            if time.monotonic() > proximo_correo:
                proximo_correo = time.monotonic() + cada_correo
                revisar_correo(cfg, carpeta, archivo_correo)
            if detener is not None and getattr(detener, "correo_ya", False):
                detener.correo_ya = False  # se acaba de configurar el correo en la ventana
                cfg.update({k: v for k, v in leer_config().items() if k.startswith("correo_")})
                cfg.pop("_correo_error", None)
                revisar_correo(cfg, carpeta, archivo_correo)
            if time.monotonic() < proxima:
                continue
            proxima = time.monotonic() + intervalo
            ultima = datetime.now().strftime("%H:%M:%S")
            if not carpeta.is_dir():
                log.warning("La carpeta no está disponible; reintentando...")
                continue
            hubo = escanear(cfg, carpeta, db, guardar=guardar_avance)
            nueva_fecha = fecha_archivo()  # ¿editaron categorias.txt? -> reclasificar todo
            reclasificar = nueva_fecha != fecha_cat
            fecha_cat = nueva_fecha
            if reclasificar:
                log.info("Cambiaron las categorías: reclasificando todas las cotizaciones...")
            if clasificar_y_copiar(db, carpeta, cfg, todo=reclasificar) or hubo:
                guardar_db(archivo_db, db)
                pendiente = True
            if fecha_sinonimos() != fecha_sin:  # cambiaron los sinónimos: rehacer el Comparativo
                fecha_sin = fecha_sinonimos()
                log.info("Cambiaron los sinónimos: se rehace el Comparativo.")
                pendiente = True
            if pendiente:
                pendiente = not actualizar_excel(cfg, carpeta, db)
            avisar_alzas(cfg, db)
            respaldar(excel, archivo_db, int(cfg["respaldos_dias"]))  # una vez al día (si ya está, no hace nada)
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
