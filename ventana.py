"""Ventana de Windows del Buscador de cotizaciones (creado por Carlos).

La lectura de los PDF sigue igual que en la consola (todos los núcleos, en
procesos aparte); esta ventana solo muestra el avance, por eso no la hace más lenta.
"""

from __future__ import annotations

import logging
import os
import queue
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
            main(["--vigilar"], detener=self.detener, ventana=True)
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
