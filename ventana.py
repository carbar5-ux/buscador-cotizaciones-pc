"""Ventana de Windows del Buscador de cotizaciones (creado por Carlos).

La lectura de los PDF sigue igual que en la consola (todos los núcleos, en
procesos aparte); esta ventana solo muestra el avance, por eso no la hace más lenta.
"""

from __future__ import annotations

import datetime
import logging
import os
import queue
import re
import socket
import subprocess
import sys
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))

import consola  # noqa: E402
from actualizador import reiniciar, version_local  # noqa: E402
from agrupador import (agrupar, calza_preparado, descripcion_representativa, preparar,  # noqa: E402
                       preparar_consulta, sinonimos, unidad_normal)
from clasificador import CARPETA_CLASIFICADAS, asegurar_archivo  # noqa: E402
from excel_writer import limpiar_texto  # noqa: E402
from buscador_cotizaciones import candado_instancia, guardar_carpeta, leer_config, main  # noqa: E402
from consola import _duracion  # noqa: E402

# Colores
VERDE_OSC, VERDE, NARANJO = "#0B6E3A", "#16A34A", "#E8700A"
ROJO, AMARILLO = "#D62828", "#E9A800"
FONDO, TARJETA, TEXTO, GRIS, BORDE = "#F3F5F7", "#FFFFFF", "#1F2933", "#6B7280", "#E5E7EB"
FUENTE = "Segoe UI"

log = logging.getLogger("cotizaciones")


class ManejadorVentana(logging.Handler):
    def __init__(self, cola: queue.Queue):
        super().__init__()
        self.cola = cola

    def emit(self, record: logging.LogRecord) -> None:
        try:
            self.cola.put(("log", {"texto": self.format(record), "nivel": record.levelno}))
        except Exception:
            pass


def abrir(ruta: Path) -> None:
    if not ruta.exists():
        messagebox.showinfo("Buscador de cotizaciones", f"Todavía no existe:\n{ruta}")
        return
    if hasattr(os, "startfile"):
        os.startfile(ruta)  # Windows
    else:
        subprocess.Popen(["xdg-open", str(ruta)])


def tomar_candado():
    """Una sola copia del buscador a la vez. Si hay una ventana abierta, la muestra y sale."""
    s = candado_instancia()
    if s is not None:
        return s
    try:
        c = socket.create_connection(("127.0.0.1", 47653), timeout=3)
        c.sendall(b"hola")
        respuesta = c.recv(16)
        c.close()
    except OSError:
        respuesta = b""
    if respuesta == b"gui":
        sys.exit(0)  # ya hay una ventana: se mostró al frente
    for _ in range(20):  # había una copia en segundo plano: se está cerrando para dejar la ventana
        time.sleep(0.5)
        s = candado_instancia()
        if s is not None:
            return s
    return None


class Ventana:
    def __init__(self, raiz: tk.Tk, candado: socket.socket, minimizado: bool):
        self.raiz = raiz
        self.candado = candado
        self.cola: queue.Queue = queue.Queue()
        self.detener = threading.Event()
        self.compartido: dict = {}  # el motor deja aquí los datos (para el buscador de precios)
        self._indice: tuple[int, list] | None = None
        self.leidas_sesion = 0
        self.ultimo_avance = 0.0

        self.cfg = leer_config()
        self.carpeta = Path(self.cfg["carpeta_cotizaciones"]) if self.cfg["carpeta_cotizaciones"] else None
        self.excel = None

        raiz.title("Buscador de cotizaciones — Carlos")
        raiz.configure(bg=FONDO)
        try:
            raiz.iconbitmap(str(BASE / "icono_cotizaciones.ico"))
        except Exception:
            pass
        raiz.protocol("WM_DELETE_WINDOW", self.cerrar)
        self._estilos()
        self._armar()
        self._ajustar_tamano()

        if self.carpeta is None or not self.carpeta.is_dir():
            raiz.after(300, self.primera_carpeta)
        else:
            self.arrancar()
        if minimizado:
            raiz.iconify()
        raiz.after(100, self.procesar_cola)
        raiz.after(500, self.atender_candado)

    def _ajustar_tamano(self) -> None:
        """Tamaño según lo que ocupa el contenido y la pantalla (con letra agrandada en Windows no se corta)."""
        r = self.raiz
        r.update_idletasks()
        ancho_p, alto_p = r.winfo_screenwidth(), r.winfo_screenheight()
        ancho = min(max(r.winfo_reqwidth(), int(ancho_p * 0.55)), int(ancho_p * 0.92))
        alto = min(r.winfo_reqheight() + 20, int(alto_p * 0.88))
        r.geometry(f"{ancho}x{alto}+{(ancho_p - ancho) // 2}+{max(0, (alto_p - alto) // 3)}")
        r.minsize(min(ancho, int(ancho_p * 0.5)), min(alto, int(alto_p * 0.6)))

    # ---------------- aspecto ----------------
    def _estilos(self) -> None:
        st = ttk.Style(self.raiz)
        st.theme_use("clam")
        st.configure("Barra.Horizontal.TProgressbar", troughcolor=BORDE, background=ROJO, lightcolor=ROJO,
                     darkcolor=ROJO, bordercolor=BORDE, thickness=28)
        for nombre, color in (("Verde", VERDE), ("Naranjo", NARANJO), ("Gris", "#4B5563")):
            st.configure(f"{nombre}.TButton", font=(FUENTE, 10, "bold"), foreground="white", background=color,
                         bordercolor=color, focusthickness=0, padding=(14, 8))
            st.map(f"{nombre}.TButton", background=[("active", color), ("pressed", color)])
        self.estilo = st

    def _tarjeta(self, padre, **kw) -> tk.Frame:
        return tk.Frame(padre, bg=TARJETA, highlightbackground=BORDE, highlightthickness=1, **kw)

    def _armar(self) -> None:
        r = self.raiz
        # Encabezado
        enc = tk.Frame(r, bg=VERDE_OSC)
        enc.pack(fill="x")
        moneda = tk.Canvas(enc, width=54, height=54, bg=VERDE_OSC, highlightthickness=0)
        moneda.create_oval(3, 3, 51, 51, fill=VERDE, outline="white", width=3)
        moneda.create_text(27, 28, text="$", fill="white", font=(FUENTE, 22, "bold"))
        moneda.pack(side="left", padx=(18, 10), pady=12)
        tit = tk.Frame(enc, bg=VERDE_OSC)
        tit.pack(side="left")
        tk.Label(tit, text="Buscador de cotizaciones", bg=VERDE_OSC, fg="white",
                 font=(FUENTE, 18, "bold")).pack(anchor="w")
        tk.Label(tit, text=f"por Carlos  ·  versión {version_local()}", bg=VERDE_OSC, fg="#CDEBD8",
                 font=(FUENTE, 10)).pack(anchor="w")
        busca = tk.Frame(enc, bg=VERDE_OSC)
        busca.pack(side="right", padx=18)
        tk.Label(busca, text="Buscar precio", bg=VERDE_OSC, fg="white", font=(FUENTE, 10, "bold")).pack(anchor="w")
        fila_b = tk.Frame(busca, bg=VERDE_OSC)
        fila_b.pack()
        self.txt_buscar = tk.Entry(fila_b, font=(FUENTE, 12), width=24, relief="flat")
        self.txt_buscar.pack(side="left", ipady=5)
        self.txt_buscar.bind("<Return>", lambda e: self.buscar_precio())
        tk.Button(fila_b, text="Buscar", command=self.buscar_precio, bg=NARANJO, fg="white", relief="flat",
                  font=(FUENTE, 10, "bold"), cursor="hand2", activebackground=NARANJO).pack(side="left", ipady=3)
        tk.Button(fila_b, text="Planilla", command=self.abrir_planilla, bg="white", fg=VERDE_OSC, relief="flat",
                  font=(FUENTE, 10, "bold"), cursor="hand2").pack(side="left", padx=(8, 0), ipady=3)
        r.bind_all("<Control-f>", lambda e: self.txt_buscar.focus_set())

        cuerpo = tk.Frame(r, bg=FONDO)
        cuerpo.pack(fill="both", expand=True, padx=18, pady=14)

        # Carpeta y Excel
        info = self._tarjeta(cuerpo)
        info.pack(fill="x")
        fila = tk.Frame(info, bg=TARJETA)
        fila.pack(fill="x", padx=14, pady=(10, 2))
        tk.Label(fila, text="Carpeta", bg=TARJETA, fg=GRIS, font=(FUENTE, 9, "bold"), width=8,
                 anchor="w").pack(side="left")
        self.lbl_carpeta = tk.Label(fila, text="—", bg=TARJETA, fg=TEXTO, font=(FUENTE, 10), anchor="w")
        self.lbl_carpeta.pack(side="left", fill="x", expand=True)
        tk.Button(fila, text="Cambiar…", command=self.cambiar_carpeta, relief="flat", bg=BORDE, fg=TEXTO,
                  font=(FUENTE, 9), cursor="hand2").pack(side="right")
        fila2 = tk.Frame(info, bg=TARJETA)
        fila2.pack(fill="x", padx=14, pady=(2, 10))
        tk.Label(fila2, text="Excel", bg=TARJETA, fg=GRIS, font=(FUENTE, 9, "bold"), width=8,
                 anchor="w").pack(side="left")
        self.lbl_excel = tk.Label(fila2, text="—", bg=TARJETA, fg=TEXTO, font=(FUENTE, 10), anchor="w")
        self.lbl_excel.pack(side="left", fill="x", expand=True)

        # Estado y barra
        est = self._tarjeta(cuerpo)
        est.pack(fill="x", pady=12)
        self.lbl_estado = tk.Label(est, text="Preparando…", bg=TARJETA, fg=TEXTO, font=(FUENTE, 20, "bold"),
                                   anchor="w")
        self.lbl_estado.pack(fill="x", padx=16, pady=(12, 0))
        self.lbl_detalle = tk.Label(est, text="", bg=TARJETA, fg=GRIS, font=(FUENTE, 10), anchor="w")
        self.lbl_detalle.pack(fill="x", padx=16)
        self.barra = ttk.Progressbar(est, style="Barra.Horizontal.TProgressbar", maximum=100)
        self.barra.pack(fill="x", padx=16, pady=(10, 4))
        pie = tk.Frame(est, bg=TARJETA)
        pie.pack(fill="x", padx=16, pady=(0, 12))
        self.lbl_pct = tk.Label(pie, text="0 %", bg=TARJETA, fg=TEXTO, font=(FUENTE, 12, "bold"))
        self.lbl_pct.pack(side="left")
        self.lbl_cuenta = tk.Label(pie, text="", bg=TARJETA, fg=TEXTO, font=(FUENTE, 11))
        self.lbl_cuenta.pack(side="left", padx=14)
        self.lbl_falta = tk.Label(pie, text="", bg=TARJETA, fg="#0E7490", font=(FUENTE, 11))
        self.lbl_falta.pack(side="left")
        self.lbl_archivo = tk.Label(pie, text="", bg=TARJETA, fg=GRIS, font=(FUENTE, 9), anchor="e")
        self.lbl_archivo.pack(side="right")

        # Indicadores
        ind = tk.Frame(cuerpo, bg=FONDO)
        ind.pack(fill="x")
        self.kpi = {}
        for i, (clave, titulo) in enumerate((("excel", "Cotizaciones en el Excel"),
                                             ("sesion", "Leídas en esta sesión"),
                                             ("revision", "Última revisión"))):
            t = self._tarjeta(ind)
            t.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 10, 0))
            ind.columnconfigure(i, weight=1, uniform="kpi")
            tk.Label(t, text=titulo, bg=TARJETA, fg=GRIS, font=(FUENTE, 9, "bold")).pack(anchor="w", padx=14,
                                                                                       pady=(10, 0))
            v = tk.Label(t, text="—", bg=TARJETA, fg=TEXTO, font=(FUENTE, 22, "bold"))
            v.pack(anchor="w", padx=14, pady=(0, 10))
            self.kpi[clave] = v

        # Botones
        bot = tk.Frame(cuerpo, bg=FONDO)
        bot.pack(fill="x", pady=12)
        ttk.Button(bot, text="Abrir Excel", style="Naranjo.TButton", command=self.abrir_excel,
                   cursor="hand2").pack(side="left")
        ttk.Button(bot, text="Abrir carpeta", style="Verde.TButton",
                   command=lambda: self.carpeta and abrir(self.carpeta), cursor="hand2").pack(side="left", padx=8)
        ttk.Button(bot, text="Revisar ahora", style="Gris.TButton", command=self.revisar_ahora,
                   cursor="hand2").pack(side="left")
        ttk.Button(bot, text="Manual", style="Gris.TButton", command=lambda: abrir(BASE / "Manual_de_uso.pdf"),
                   cursor="hand2").pack(side="right")
        ttk.Button(bot, text="Categorías", style="Gris.TButton", command=self.editar_categorias,
                   cursor="hand2").pack(side="right", padx=8)
        ttk.Button(bot, text="Clasificadas", style="Verde.TButton", command=self.abrir_clasificadas,
                   cursor="hand2").pack(side="right", padx=(8, 0))

        # Registro
        reg = self._tarjeta(cuerpo)
        reg.pack(fill="both", expand=True)
        self.txt = tk.Text(reg, height=4, bg=TARJETA, fg=GRIS, font=("Consolas", 9), relief="flat",
                           wrap="none", state="disabled", padx=10, pady=8)
        self.txt.pack(fill="both", expand=True)
        self.txt.tag_configure("aviso", foreground=AMARILLO)
        self.txt.tag_configure("error", foreground=ROJO)

        tk.Label(r, text="Hecho por Carlos", bg=FONDO, fg=GRIS, font=(FUENTE, 8)).pack(side="bottom", pady=(0, 6))

    # ---------------- motor ----------------
    def primera_carpeta(self) -> None:
        messagebox.showinfo("Buscador de cotizaciones",
                            "Elige la carpeta donde guardas tus cotizaciones PDF.")
        elegida = filedialog.askdirectory(parent=self.raiz, title="Carpeta de cotizaciones PDF")
        if not elegida:
            self.raiz.destroy()
            return
        guardar_carpeta(self.cfg, Path(elegida))
        self.carpeta = Path(elegida)
        self.arrancar()

    def arrancar(self) -> None:
        self.cfg = leer_config()
        self.excel = Path(self.cfg["archivo_excel"]) if self.cfg["archivo_excel"] else \
            self.carpeta / "Base_Cotizaciones.xlsx"
        self.lbl_carpeta.config(text=str(self.carpeta))
        self.lbl_excel.config(text=str(self.excel))
        threading.Thread(target=self._motor, daemon=True).start()

    def _motor(self) -> None:
        try:
            main(["--vigilar"], detener=self.detener, ventana=True, compartido=self.compartido)
        except SystemExit:
            pass
        except Exception as e:
            log.exception("Error inesperado")
            self.cola.put(("error", {"texto": str(e)}))

    # ---------------- eventos ----------------
    def procesar_cola(self) -> None:
        try:
            for _ in range(500):
                tipo, d = self.cola.get_nowait()
                getattr(self, f"en_{tipo}", lambda d: None)(d)
        except queue.Empty:
            pass
        self.raiz.after(100, self.procesar_cola)

    def en_log(self, d) -> None:
        etiqueta = "error" if d["nivel"] >= logging.ERROR else "aviso" if d["nivel"] >= logging.WARNING else None
        self.txt.config(state="normal")
        self.txt.insert("end", d["texto"] + "\n", etiqueta or ())
        lineas = int(self.txt.index("end-1c").split(".")[0])
        if lineas > 400:
            self.txt.delete("1.0", f"{lineas - 400}.0")
        self.txt.see("end")
        self.txt.config(state="disabled")

    def _color_barra(self, pct: float) -> None:
        color = ROJO if pct < 34 else AMARILLO if pct < 67 else VERDE
        self.estilo.configure("Barra.Horizontal.TProgressbar", background=color, lightcolor=color, darkcolor=color)

    def en_inicio(self, d) -> None:
        self.lbl_estado.config(text="Leyendo cotizaciones…", fg=TEXTO)
        self.lbl_detalle.config(text=f"{d['total']} cotizaciones nuevas o modificadas")
        self.barra["value"] = 0
        self._color_barra(0)
        self.lbl_pct.config(text="0 %")
        self.lbl_cuenta.config(text=f"0 / {d['total']}")
        self.lbl_falta.config(text="")

    def en_avance(self, d) -> None:
        pct = 100 * d["hechos"] / d["total"]
        self.barra["value"] = pct
        self._color_barra(pct)
        self.lbl_pct.config(text=f"{pct:.0f} %")
        self.lbl_cuenta.config(text=f"{d['hechos']} / {d['total']}")
        f = d["falta"]
        self.lbl_falta.config(text=("" if d["hechos"] == d["total"] else
                                    f"quedan ~{f:.0f} s" if f < 90 else f"quedan ~{f / 60:.0f} min"))
        self.lbl_archivo.config(text=d["nombre"][:50])
        if d["hechos"] == d["total"]:
            self.lbl_estado.config(text="Guardando el Excel…", fg=TEXTO)

    def en_terminado(self, d) -> None:
        self.barra["value"] = 100
        self._color_barra(100)
        self.lbl_pct.config(text="100 %")
        self.lbl_falta.config(text="")
        self.lbl_archivo.config(text="")
        self.kpi["excel"].config(text=f"{d['total_excel']:,}".replace(",", "."))
        if d["leidas"]:
            self.leidas_sesion += d["leidas"]
            self.kpi["sesion"].config(text=f"{self.leidas_sesion:,}".replace(",", "."))
            self.lbl_estado.config(text="TERMINADO", fg=ROJO)
            self.lbl_detalle.config(text=f"{d['leidas']} cotizaciones leídas en {_duracion(d['seg'])}  ·  "
                                         f"{d['total_excel']} en total en el Excel")
            self.lbl_cuenta.config(text=f"{d['leidas']} / {d['leidas']}")
            self._llamar_atencion()
        else:
            self.lbl_estado.config(text="TERMINADO", fg=ROJO)
            self.lbl_detalle.config(text=f"Sin cotizaciones nuevas  ·  {d['total_excel']} en el Excel")
            self.lbl_cuenta.config(text="")
            self.kpi["sesion"].config(text=str(self.leidas_sesion))

    def en_vigilando(self, d) -> None:
        self.kpi["excel"].config(text=f"{d['total_excel']:,}".replace(",", "."))
        self.kpi["revision"].config(text=d["ultima"])
        if self.lbl_estado.cget("text") != "TERMINADO":
            self.lbl_estado.config(text="Vigilando la carpeta", fg=VERDE)
            self.lbl_detalle.config(text="Guarda cotizaciones en la carpeta y entran solas al Excel")
            self.barra["value"] = 100
            self._color_barra(100)
            self.lbl_pct.config(text="100 %")

    def en_error(self, d) -> None:
        self.lbl_estado.config(text="Ocurrió un error", fg=ROJO)
        self.lbl_detalle.config(text=d["texto"][:120] + "  (detalle en registro.log)")

    def _llamar_atencion(self) -> None:
        if self.cfg.get("aviso_sonido", True):
            try:
                import winsound

                winsound.MessageBeep(winsound.MB_ICONASTERISK)
            except Exception:
                self.raiz.bell()
        if self.cfg.get("aviso_ventana", True):
            self.raiz.deiconify()
            self.raiz.lift()
            self.raiz.attributes("-topmost", True)
            self.raiz.after(1500, lambda: self.raiz.attributes("-topmost", False))

    # ---------------- botones ----------------
    def abrir_excel(self) -> None:
        if self.excel:
            abrir(self.excel)

    # ---------------- buscador de precios ----------------
    def _indice_precios(self) -> list:
        """Todas las líneas con precio, preparadas una vez (se rehace si cambian los datos)."""
        db = self.compartido.get("db")
        if not db:
            return []
        try:
            registros = list(db.items())
        except RuntimeError:  # el motor está agregando cotizaciones justo ahora
            registros = list(db.items())
        firma = (len(registros), sum(len(r.get("datos", {}).get("items", [])) for _, r in registros))
        if self._indice and self._indice[0] == firma:
            return self._indice[1]
        sin = sinonimos()
        filas = []
        for clave, reg in registros:
            d = reg.get("datos", {})
            for it in d.get("items", []):
                pu = it.get("precio_unitario")
                if pu is None or not it.get("descripcion"):
                    continue
                filas.append((preparar(it["descripcion"], it.get("codigo", ""), sin), it, d, clave))
        self._indice = (firma, filas)
        return filas

    def abrir_planilla(self) -> None:
        if not self.compartido.get("db"):
            messagebox.showinfo("Planilla", "Todavía se están cargando las cotizaciones. Intenta en unos segundos.")
            return
        Planilla(self)

    def buscar_precio(self) -> None:
        consulta = self.txt_buscar.get().strip()
        if not consulta:
            self.txt_buscar.focus_set()
            return
        if not self.compartido.get("db"):
            messagebox.showinfo("Buscar precio", "Todavía se están cargando las cotizaciones. Intenta en unos segundos.")
            return
        Planilla(self, consulta)

    def editar_categorias(self) -> None:
        """Abre categorias.txt en el Bloc de notas; al guardarlo, el programa reclasifica solo."""
        ruta = asegurar_archivo()
        if sys.platform == "win32":
            subprocess.Popen(["notepad.exe", str(ruta)])
        else:
            abrir(ruta)

    def abrir_clasificadas(self) -> None:
        if not self.carpeta:
            return
        destino = self.carpeta / CARPETA_CLASIFICADAS
        destino.mkdir(exist_ok=True)
        abrir(destino)

    def revisar_ahora(self) -> None:
        self.detener.revisar_ya = True
        self.lbl_estado.config(text="Revisando la carpeta…", fg=TEXTO)

    def cambiar_carpeta(self) -> None:
        elegida = filedialog.askdirectory(parent=self.raiz, title="Carpeta de cotizaciones PDF",
                                          initialdir=str(self.carpeta or Path.home()))
        if not elegida or (self.carpeta and Path(elegida) == self.carpeta):
            return
        guardar_carpeta(self.cfg, Path(elegida))
        reiniciar()  # vuelve a abrir la ventana con la carpeta nueva

    def atender_candado(self) -> None:
        """Si alguien abre el buscador otra vez, se muestra esta ventana en vez de abrir otra."""
        try:
            conn, _ = self.candado.accept()
            conn.sendall(b"gui")
            conn.close()
            self.raiz.deiconify()
            self.raiz.lift()
            self.raiz.focus_force()
        except (BlockingIOError, OSError):
            pass
        self.raiz.after(500, self.atender_candado)

    def cerrar(self) -> None:
        if self.lbl_estado.cget("text").startswith("Leyendo") and not messagebox.askyesno(
                "Buscador de cotizaciones", "Todavía está leyendo cotizaciones.\n¿Cerrar de todas formas?\n"
                                            "(lo leído hasta ahora queda guardado)"):
            return
        self.detener.set()
        self.raiz.destroy()
        logging.shutdown()
        os._exit(0)


def _dinero(v, moneda: str) -> str:
    if v is None:
        return ""
    if moneda in ("CLP", ""):
        return "$ " + f"{v:,.0f}".replace(",", ".")
    return f"{moneda} " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _fecha_txt(iso: str) -> str:
    return f"{iso[8:10]}-{iso[5:7]}-{iso[:4]}" if iso and len(iso) >= 10 else ""


class Planilla:
    """Todas las líneas de todas las cotizaciones, con filtros, como una planilla dentro del programa.

    Doble clic en una fila abre la cotización. Clic en un título de columna ordena por ella.
    """

    COLUMNAS = (("descripcion", "Descripción", 320, "w"), ("precio", "Precio unitario", 115, "e"),
                ("cantidad", "Cantidad", 75, "e"), ("unidad", "Unidad", 60, "center"), ("total", "Total", 115, "e"),
                ("proveedor", "Empresa", 190, "w"), ("fecha", "Fecha", 95, "center"), ("numero", "N° cot.", 80, "w"),
                ("categoria", "Categoría", 130, "w"))
    MAXIMO = 3000  # filas dibujadas a la vez (con más, la tabla se pone lenta: se pide filtrar)
    TODAS = "(todas)"

    def __init__(self, app: "Ventana", texto: str = ""):
        self.app = app
        self.orden = ("precio", False) if texto else ("fecha", True)
        self.top = tk.Toplevel(app.raiz)
        self.top.title("Planilla de cotizaciones — Carlos")
        self.top.configure(bg=FONDO)
        try:
            self.top.iconbitmap(str(BASE / "icono_cotizaciones.ico"))
        except Exception:
            pass
        filas_idx = app._indice_precios()

        # ---- filtros ----
        f = tk.Frame(self.top, bg=FONDO)
        f.pack(fill="x", padx=14, pady=(8, 2))

        fb = tk.Frame(self.top, bg=FONDO)  # buscador en su propia línea (siempre visible)
        fb.pack(fill="x", padx=14, pady=(12, 0), before=f)
        tk.Label(fb, text="Buscar suministro (descripción o código):", bg=FONDO, fg=TEXTO,
                 font=(FUENTE, 11, "bold")).pack(side="left")
        self.txt = tk.Entry(fb, font=(FUENTE, 13), width=40)
        self.txt.insert(0, texto)
        self.txt.pack(side="left", padx=8, ipady=4, fill="x", expand=True)
        self.txt.bind("<KeyRelease>", lambda e: self._programar())
        self.txt.focus_set()
        empresas = sorted({d.get("proveedor", "") for _, _, d, _ in filas_idx if d.get("proveedor")}, key=str.lower)
        cats = sorted({c.strip() for _, it, _, _ in filas_idx for c in (it.get("categoria") or "").split(",") if c.strip()})
        monedas = sorted({d.get("moneda", "") for _, _, d, _ in filas_idx if d.get("moneda")})
        self.cb_empresa = self._combo(f, "Empresa", [self.TODAS] + empresas, 0, 26)
        self.cb_cat = self._combo(f, "Categoría", [self.TODAS] + cats + ["(sin categoría)"], 1, 18)
        self.cb_moneda = self._combo(f, "Moneda", [self.TODAS] + monedas, 2, 8)
        etiqueta_d = tk.Label(f, text="Fecha desde / hasta (dd-mm-aaaa)", bg=FONDO, fg=GRIS, font=(FUENTE, 9, "bold"))
        etiqueta_d.grid(row=0, column=3, columnspan=2, sticky="w")
        self.txt_desde = tk.Entry(f, font=(FUENTE, 11), width=11)
        self.txt_hasta = tk.Entry(f, font=(FUENTE, 11), width=11)
        self.txt_desde.grid(row=1, column=3, padx=(0, 4), ipady=3)
        self.txt_hasta.grid(row=1, column=4, padx=(0, 10), ipady=3)
        for e in (self.txt_desde, self.txt_hasta):
            e.bind("<KeyRelease>", lambda ev: self._programar())
        ttk.Button(f, text="Limpiar filtros", style="Gris.TButton", command=self.limpiar).grid(row=1, column=5, padx=4)
        ttk.Button(f, text="Exportar a Excel", style="Verde.TButton", command=self.exportar).grid(row=1, column=6)

        self.lbl_resumen = tk.Label(self.top, text="", bg=FONDO, fg=TEXTO, font=(FUENTE, 11, "bold"), anchor="w",
                                    justify="left", wraplength=1150)
        self.lbl_resumen.pack(fill="x", padx=14, pady=(8, 0))
        tk.Label(self.top, text="Doble clic en una fila: abre la cotización  ·  clic en un título: ordena por esa columna",
                 bg=FONDO, fg=GRIS, font=(FUENTE, 9), anchor="w").pack(fill="x", padx=14)

        # ---- tabla ----
        marco = tk.Frame(self.top, bg=FONDO)
        marco.pack(fill="both", expand=True, padx=14, pady=8)
        self.tabla = ttk.Treeview(marco, columns=[c[0] for c in self.COLUMNAS], show="headings")
        for clave, titulo, ancho, alin in self.COLUMNAS:
            self.tabla.heading(clave, text=titulo, command=lambda k=clave: self.ordenar(k))
            self.tabla.column(clave, width=ancho, anchor=alin, stretch=clave == "descripcion")
        sy = ttk.Scrollbar(marco, orient="vertical", command=self.tabla.yview)
        sx = ttk.Scrollbar(marco, orient="horizontal", command=self.tabla.xview)
        self.tabla.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tabla.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        marco.rowconfigure(0, weight=1)
        marco.columnconfigure(0, weight=1)
        self.tabla.tag_configure("barato", background="#DCFCE7")
        self.tabla.tag_configure("par", background="#F9FAFB")
        self.tabla.bind("<Double-1>", self.abrir_pdf)
        st = ttk.Style(self.top)
        st.configure("Treeview", font=(FUENTE, 10), rowheight=26)
        st.configure("Treeview.Heading", font=(FUENTE, 10, "bold"))

        ancho = min(1200, self.top.winfo_screenwidth() - 40)
        alto = min(720, self.top.winfo_screenheight() - 80)
        self.top.geometry(f"{ancho}x{alto}+{max(0, (self.top.winfo_screenwidth() - ancho) // 2)}+20")
        self.claves: dict[str, str] = {}
        self.visibles: list = []
        self._tarea = None
        self.aplicar()

    def _combo(self, padre, titulo, valores, col, ancho):
        tk.Label(padre, text=titulo, bg=FONDO, fg=GRIS, font=(FUENTE, 9, "bold")).grid(row=0, column=col, sticky="w")
        cb = ttk.Combobox(padre, values=valores, state="readonly", width=ancho, font=(FUENTE, 10))
        cb.set(self.TODAS)
        cb.grid(row=1, column=col, padx=(0, 10), ipady=2)
        cb.bind("<<ComboboxSelected>>", lambda e: self.aplicar())
        return cb

    def _programar(self) -> None:
        """Filtra medio segundo después de dejar de escribir (no en cada tecla)."""
        if self._tarea:
            self.top.after_cancel(self._tarea)
        self._tarea = self.top.after(450, self.aplicar)

    @staticmethod
    def _iso(texto: str) -> str:
        m = re.fullmatch(r"\s*(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})\s*", texto or "")
        if not m:
            return ""
        d, mes, a = m.groups()
        a = ("20" + a) if len(a) == 2 else a
        return f"{a}-{int(mes):02d}-{int(d):02d}"

    def limpiar(self) -> None:
        for e in (self.txt, self.txt_desde, self.txt_hasta):
            e.delete(0, "end")
        for cb in (self.cb_empresa, self.cb_cat, self.cb_moneda):
            cb.set(self.TODAS)
        self.orden = ("fecha", True)
        self.aplicar()

    def ordenar(self, columna: str) -> None:
        col, rev = self.orden
        self.orden = (columna, not rev if col == columna else columna in ("fecha", "total", "precio", "cantidad"))
        self.aplicar()

    def aplicar(self) -> None:
        self._tarea = None
        texto = self.txt.get().strip()
        q = preparar_consulta(texto) if texto else None
        empresa, cat, moneda = self.cb_empresa.get(), self.cb_cat.get(), self.cb_moneda.get()
        desde, hasta = self._iso(self.txt_desde.get()), self._iso(self.txt_hasta.get())
        filas = []
        for prep, it, d, clave in self.app._indice_precios():
            if q and not calza_preparado(q, prep):
                continue
            if empresa != self.TODAS and d.get("proveedor") != empresa:
                continue
            if moneda != self.TODAS and d.get("moneda") != moneda:
                continue
            if cat != self.TODAS:
                cats_it = it.get("categoria") or ""
                if (cat == "(sin categoría)" and cats_it) or (cat != "(sin categoría)" and cat not in cats_it):
                    continue
            f = d.get("fecha") or ""
            if (desde and f < desde) or (hasta and (not f or f > hasta)):
                continue
            filas.append((it, d, clave))

        col, rev = self.orden
        claves_orden = {
            "descripcion": lambda t: (t[0].get("descripcion") or "").lower(),
            "precio": lambda t: t[0].get("precio_unitario") or 0,
            "cantidad": lambda t: t[0].get("cantidad") or 0,
            "unidad": lambda t: (t[0].get("unidad") or "").lower(),
            "total": lambda t: t[0].get("total") or 0,
            "proveedor": lambda t: (t[1].get("proveedor") or "").lower(),
            "fecha": lambda t: t[1].get("fecha") or "",
            "numero": lambda t: str(t[1].get("numero") or ""),
            "categoria": lambda t: t[0].get("categoria") or "",
        }
        filas.sort(key=claves_orden[col], reverse=rev)
        for clave_c, titulo, *_ in self.COLUMNAS:
            flecha = (" ▼" if rev else " ▲") if clave_c == col else ""
            self.tabla.heading(clave_c, text=titulo + flecha)

        self.tabla.delete(*self.tabla.get_children())
        self.claves.clear()
        self.visibles = filas
        # con búsqueda: el producto más repetido entre los resultados; se marca su precio más bajo
        self._grupo_principal = None
        minimo = None
        if q and filas:
            grupos = agrupar([(t[0].get("descripcion", ""), t[0].get("unidad", ""), t[1].get("moneda", ""), t)
                              for t in filas[:5000]])
            principal = max(grupos, key=len)
            self._grupo_principal = (principal, len(grupos))
            minimo = min(t[0]["precio_unitario"] for t in principal)
            ids_principal = {id(t) for t in principal}
        for n, (it, d, clave) in enumerate(filas[: self.MAXIMO]):
            mon = d.get("moneda", "")
            es_barato = (minimo is not None and it["precio_unitario"] == minimo
                         and id(filas[n]) in ids_principal)
            tag = ("barato",) if es_barato else (("par",) if n % 2 else ())
            iid = self.tabla.insert("", "end", tags=tag, values=(
                it.get("descripcion", ""), _dinero(it.get("precio_unitario"), mon),
                f"{it['cantidad']:g}" if it.get("cantidad") is not None else "", it.get("unidad", ""),
                _dinero(it.get("total"), mon), d.get("proveedor", ""), _fecha_txt(d.get("fecha", "")),
                d.get("numero", ""), it.get("categoria", "")))
            self.claves[iid] = clave
        self._resumen(filas, bool(q))

    def _resumen(self, filas: list, con_busqueda: bool) -> None:
        if not filas:
            self.lbl_resumen.config(text="No hay líneas con esos filtros.", fg=ROJO)
            return
        n_cot = len({t[2] for t in filas})
        totales: dict[str, float] = {}
        for it, d, _ in filas:
            if it.get("total") is not None:
                totales[d.get("moneda", "")] = totales.get(d.get("moneda", ""), 0) + it["total"]
        partes = [f"{len(filas):,} líneas de {n_cot:,} cotizaciones".replace(",", ".")]
        partes.append("suma de totales " + "  +  ".join(_dinero(v, m) for m, v in totales.items()))
        if len(filas) > self.MAXIMO:
            partes.append(f"se muestran las primeras {self.MAXIMO:,}: usa los filtros para acotar".replace(",", "."))
        texto = "   ·   ".join(partes)
        if con_busqueda and self._grupo_principal:
            # mínimo / último / promedio del producto más repetido entre los resultados (mismo producto,
            # mismas medidas, unidad y moneda): no se promedian productos distintos que contienen la palabra
            g, total_grupos = self._grupo_principal
            mon = g[0][1].get("moneda", "")
            uni = unidad_normal(g[0][0].get("unidad", ""))
            precios = [t[0]["precio_unitario"] for t in g]
            barato = min(g, key=lambda t: t[0]["precio_unitario"])
            ultimo = max(g, key=lambda t: t[1].get("fecha") or "")
            nombre = descripcion_representativa([t[0].get("descripcion", "") for t in g])[:60]
            otros = f"   (hay {total_grupos - 1} producto(s) más: acota la búsqueda)" if total_grupos > 1 else ""
            texto += (f"\n{nombre} — por {uni or 'unidad'} en {mon}:   mínimo {_dinero(min(precios), mon)} "
                      f"({barato[1].get('proveedor', '')[:30]}, {_fecha_txt(barato[1].get('fecha', ''))})   ·   "
                      f"último {_dinero(ultimo[0]['precio_unitario'], mon)} ({ultimo[1].get('proveedor', '')[:30]}, "
                      f"{_fecha_txt(ultimo[1].get('fecha', ''))})   ·   promedio {_dinero(sum(precios) / len(precios), mon)}"
                      f"{otros}")
        self.lbl_resumen.config(text=texto, fg=TEXTO)

    def abrir_pdf(self, _evento=None) -> None:
        sel = self.tabla.selection()
        carpeta = self.app.compartido.get("carpeta")
        if sel and sel[0] in self.claves and carpeta:
            abrir(Path(carpeta) / self.claves[sel[0]])

    def exportar(self) -> None:
        """Guarda en un Excel solo las líneas filtradas."""
        if not self.visibles:
            messagebox.showinfo("Exportar", "No hay líneas para exportar.", parent=self.top)
            return
        ruta = filedialog.asksaveasfilename(parent=self.top, defaultextension=".xlsx", initialfile="Precios filtrados.xlsx",
                                            filetypes=[("Excel", "*.xlsx")], title="Guardar líneas filtradas")
        if not ruta:
            return
        from openpyxl import Workbook
        from openpyxl.worksheet.table import Table, TableStyleInfo

        carpeta = Path(self.app.compartido.get("carpeta", "."))
        wb = Workbook()
        ws = wb.active
        ws.title = "Precios"
        ws.append(["Descripción", "Precio unitario", "Cantidad", "Unidad", "Total", "Moneda", "Empresa", "Fecha",
                   "N° cotización", "Categoría", "Archivo"])
        for it, d, clave in self.visibles:
            f = d.get("fecha") or ""
            ws.append([limpiar_texto(v) for v in [it.get("descripcion"), it.get("precio_unitario"), it.get("cantidad"), it.get("unidad"),
                       it.get("total"), d.get("moneda"), d.get("proveedor"),
                       datetime.date.fromisoformat(f) if f else None, d.get("numero"), it.get("categoria"),
                       str(carpeta / clave)]])
        for col, ancho in zip("ABCDEFGHIJK", (60, 15, 10, 8, 15, 8, 30, 12, 14, 20, 50)):
            ws.column_dimensions[col].width = ancho
        for fila in ws.iter_rows(min_row=2):
            fila[1].number_format = fila[4].number_format = "#,##0.##"
            fila[7].number_format = "dd-mm-yyyy"
        tabla = Table(displayName="Filtrados", ref=f"A1:K{len(self.visibles) + 1}")
        tabla.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tabla)
        ws.freeze_panes = "A2"
        try:
            wb.save(ruta)
        except PermissionError:
            messagebox.showerror("Exportar", "No se pudo guardar: ¿el archivo está abierto en Excel?", parent=self.top)
            return
        abrir(Path(ruta))


def iniciar() -> None:
    if sys.platform == "win32":
        try:
            import ctypes

            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # letras nítidas en pantallas modernas
        except Exception:
            pass
    raiz = tk.Tk()
    raiz.withdraw()
    candado = tomar_candado()
    if candado is None:
        messagebox.showerror("Buscador de cotizaciones",
                             "El buscador ya está funcionando en este PC y no se pudo abrir otra ventana.")
        return
    raiz.deiconify()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(message)s", datefmt="%d-%m %H:%M:%S",
                        handlers=[logging.FileHandler(BASE / "registro.log", encoding="utf-8")])
    app = Ventana(raiz, candado, minimizado="--minimizado" in sys.argv)
    manejador = ManejadorVentana(app.cola)
    manejador.setFormatter(logging.Formatter("%(asctime)s  %(message)s", "%H:%M:%S"))
    logging.getLogger().addHandler(manejador)
    consola.OYENTE = lambda tipo, datos: app.cola.put((tipo, datos))
    consola.HAY_CONSOLA = False  # la ventana reemplaza a la consola
    raiz.mainloop()


if __name__ == "__main__":
    iniciar()
