"""Pruebas de: alzas de precio, limpieza de Clasificadas, respaldo diario, configuración, correo,
traspaso desde el programa anterior y exportación a PDF."""

from __future__ import annotations

import json
import os
import time
from datetime import date, timedelta
from email.message import EmailMessage
from pathlib import Path

import pytest

import buscador_cotizaciones as bc
import clasificador
import correo
import migracion
from agrupador import agrupar, variaciones

PDF = b"%PDF-1.4\n1 0 obj<<>>endobj\ntrailer<<>>\n%%EOF\n"


def _item(desc, precio, unidad="m"):
    return {"descripcion": desc, "precio_unitario": precio, "unidad": unidad, "cantidad": 10, "total": precio * 10,
            "codigo": "", "item": "1", "control": "", "pagina": 1, "descuento": None, "categoria": ""}


def _reg(proveedor, fecha, items, numero="1"):
    return {"hash": f"{proveedor}{fecha}", "tamano": 1, "mtime": 0, "procesado": "2026-01-01T00:00:00",
            "datos": {"proveedor": proveedor, "fecha": fecha, "moneda": "CLP", "numero": numero, "items": items}}


# --------------------------------------------------------------------------- alzas de precio

def test_variaciones_compara_con_la_cotizacion_anterior():
    datos = [("Cable THHN 12 AWG", "m", "CLP", ("A", "2026-01-10", 450.0)),
             ("Cable THHN 12AWG", "m", "CLP", ("B", "2026-03-01", 495.0)),
             ("cable thhn 12 awg", "m", "CLP", ("C", "2026-02-01", 400.0)),
             ("Cable NYY 3x2,5", "m", "CLP", ("D", "2026-03-05", 1900.0))]
    grupos = agrupar(datos)
    var = variaciones(grupos, fecha=lambda t: t[1], precio=lambda t: t[2], cotizacion=lambda t: t[0])
    por = {t[0]: var.get(id(t)) for _, _, _, t in datos}
    assert por["A"] is None  # la primera no tiene con qué compararse
    assert por["C"][0] == pytest.approx(400 / 450 - 1)  # bajó
    assert por["B"][0] == pytest.approx(495 / 400 - 1)  # subió frente a la anterior (C), no a la primera
    assert por["B"][1][0] == "C"
    assert por["D"] is None  # otro producto


def test_avisar_alzas_solo_de_las_nuevas(caplog):
    db = {"a.pdf": _reg("Prov A", "2026-01-10", [_item("Cable THHN 12 AWG", 450)]),
          "b.pdf": _reg("Prov B", "2026-03-01", [_item("Cable THHN 12 AWG", 520)])}
    cfg = {"_nuevas": ["b.pdf"]}
    avisos = bc.avisar_alzas(cfg, db)
    assert len(avisos) == 1 and "+16%" in avisos[0] and "Prov A" in avisos[0]
    assert "_nuevas" not in cfg
    assert bc.avisar_alzas({"_nuevas": ["a.pdf"]}, db) == []


def test_excel_marca_variacion(tmp_path):
    from openpyxl import load_workbook

    from excel_writer import construir_excel

    db = {"a.pdf": _reg("Prov A", "2026-01-10", [_item("Cable THHN 12 AWG", 450)]),
          "b.pdf": _reg("Prov B", "2026-03-01", [_item("Cable THHN 12 AWG", 540)], numero="2")}
    destino = tmp_path / "Base.xlsx"
    construir_excel(db, tmp_path, destino)
    wb = load_workbook(destino)
    ws = wb["Precios"]
    cab = [c.value for c in ws[1]]
    col = cab.index("Var. vs anterior")
    valores = {ws.cell(r, 1).value: ws.cell(r, col + 1).value for r in range(2, ws.max_row + 1)}
    assert valores["Prov B"] == pytest.approx(0.2)
    assert valores["Prov A"] is None
    revisar = [ws_r.value for ws_r in wb["Revisar"]["D"]]
    assert any("Alza de precio +20%" in (v or "") for v in revisar)


# --------------------------------------------------------------------------- Clasificadas

@pytest.fixture
def categorias(tmp_path, monkeypatch):
    archivo = tmp_path / "categorias.txt"
    archivo.write_text("Cables: cable\nAlumbrado: luminaria\n", encoding="utf-8")
    monkeypatch.setattr(clasificador, "ARCHIVO", archivo)
    return archivo


def test_copias_se_quitan_al_cambiar_categoria(tmp_path, categorias):
    carpeta = tmp_path / "cot"
    carpeta.mkdir()
    (carpeta / "c1.pdf").write_bytes(PDF)
    db = {"c1.pdf": _reg("P", "2026-01-01", [_item("Cable THHN 12", 1)])}
    clasificador.clasificar_y_copiar(db, carpeta, {}, todo=True)
    assert db["c1.pdf"]["copias"] == ["Clasificadas/Cables/c1.pdf"]
    assert (carpeta / "Clasificadas/Cables/c1.pdf").exists()

    categorias.write_text("Conductores: cable\n", encoding="utf-8")  # se renombra la categoría
    clasificador.clasificar_y_copiar(db, carpeta, {}, todo=True)
    assert (carpeta / "Clasificadas/Conductores/c1.pdf").exists()
    assert not (carpeta / "Clasificadas/Cables").exists()  # copia y carpeta vacía quitadas
    assert db["c1.pdf"]["copias"] == ["Clasificadas/Conductores/c1.pdf"]


def test_copias_antiguas_sin_registro_se_reconocen(tmp_path, categorias):
    carpeta = tmp_path / "cot"
    (carpeta / "Clasificadas/Alumbrado").mkdir(parents=True)
    (carpeta / "c1.pdf").write_bytes(PDF)
    (carpeta / "Clasificadas/Alumbrado/c1.pdf").write_bytes(PDF)  # hecha por una versión anterior
    (carpeta / "Clasificadas/Alumbrado/otro.pdf").write_bytes(PDF)  # de otra cotización: no se toca
    db = {"c1.pdf": _reg("P", "2026-01-01", [_item("Cable THHN 12", 1)])}
    clasificador.clasificar_y_copiar(db, carpeta, {}, todo=True)
    assert not (carpeta / "Clasificadas/Alumbrado/c1.pdf").exists()
    assert (carpeta / "Clasificadas/Alumbrado/otro.pdf").exists()
    assert (carpeta / "Clasificadas/Cables/c1.pdf").exists()


def test_copia_compartida_no_se_borra(tmp_path, categorias):
    carpeta = tmp_path / "cot"
    carpeta.mkdir()
    for n in ("c1.pdf", "c2.pdf"):
        (carpeta / n).write_bytes(PDF)
    db = {n: _reg("P", "2026-01-01", [_item("Cable THHN 12", 1)]) for n in ("c1.pdf", "c2.pdf")}
    clasificador.clasificar_y_copiar(db, carpeta, {}, todo=True)
    (carpeta / "c1.pdf").unlink()
    sobra = db.pop("c1.pdf")["copias"]
    en_uso = {r for reg in db.values() for r in reg.get("copias", [])}
    assert clasificador.borrar_copias(carpeta, sobra, en_uso) == 1
    assert not (carpeta / "Clasificadas/Cables/c1.pdf").exists()
    assert (carpeta / "Clasificadas/Cables/c2.pdf").exists()


def test_escanear_quita_copias_de_pdf_borrado(tmp_path, categorias, monkeypatch):
    carpeta = tmp_path / "cot"
    carpeta.mkdir()
    (carpeta / "c1.pdf").write_bytes(PDF)
    st = (carpeta / "c1.pdf").stat()
    db = {"c1.pdf": dict(_reg("P", "2026-01-01", [_item("Cable THHN 12", 1)]), tamano=st.st_size, mtime=st.st_mtime)}
    clasificador.clasificar_y_copiar(db, carpeta, {}, todo=True)
    assert (carpeta / "Clasificadas/Cables/c1.pdf").exists()
    (carpeta / "c1.pdf").unlink()
    bc.escanear(dict(bc.CONFIG_DEFECTO), carpeta, db)
    assert "c1.pdf" not in db
    assert not (carpeta / "Clasificadas/Cables/c1.pdf").exists()


def test_borrar_copias_nunca_sale_de_clasificadas(tmp_path):
    (tmp_path / "importante.pdf").write_bytes(PDF)
    assert clasificador.borrar_copias(tmp_path, ["importante.pdf", "../x.pdf"]) == 0
    assert (tmp_path / "importante.pdf").exists()


# --------------------------------------------------------------------------- respaldo

def test_respaldo_diario_y_limite(tmp_path):
    excel = tmp_path / "Base_Cotizaciones.xlsx"
    datos = tmp_path / "Base_Cotizaciones_datos.json"
    excel.write_bytes(b"x")
    datos.write_text("{}")
    resp = tmp_path / "Respaldos"
    resp.mkdir()
    for i in range(1, 6):  # respaldos de días anteriores
        d = (date.today() - timedelta(days=i)).isoformat()
        (resp / f"Base_Cotizaciones_{d}.xlsx").write_bytes(b"v")
        (resp / f"Base_Cotizaciones_datos_{d}.json").write_text("{}")
    assert bc.respaldar(excel, datos, 3) is True
    assert bc.respaldar(excel, datos, 3) is False  # una vez al día
    assert len(list(resp.glob("Base_Cotizaciones_????-??-??.xlsx"))) == 3
    assert len(list(resp.glob("Base_Cotizaciones_datos_????-??-??.json"))) == 3
    assert (resp / f"Base_Cotizaciones_{date.today().isoformat()}.xlsx").exists()
    assert bc.respaldar(excel, datos, 0) is False


# --------------------------------------------------------------------------- configuración

def test_guardar_valores_conserva_comentarios(tmp_path, monkeypatch):
    cfg_arch = tmp_path / "config.toml"
    cfg_arch.write_text("# comentario\ncarpeta_cotizaciones = 'C:\\x'\ncorreo_usuario = ''\n", encoding="utf-8")
    monkeypatch.setattr(bc, "CONFIG", cfg_arch)
    cfg = {}
    bc.guardar_valores(cfg, {"correo_usuario": "yo@gmail.com", "correo_palabras": ["cotiz", "oferta"]})
    texto = cfg_arch.read_text(encoding="utf-8")
    assert "# comentario" in texto and "correo_usuario = 'yo@gmail.com'" in texto
    assert "correo_palabras = ['cotiz', 'oferta']" in texto
    import tomllib
    assert tomllib.loads(texto)["carpeta_cotizaciones"] == "C:\\x"
    assert cfg["correo_usuario"] == "yo@gmail.com"


# --------------------------------------------------------------------------- correo

def _correo(asunto, adjuntos, remitente="Ventas Eléctrica Sur <ventas@sur.cl>"):
    m = EmailMessage()
    m["Subject"], m["From"], m["Date"] = asunto, remitente, "Tue, 10 Nov 2026 10:00:00 -0300"
    m.set_content("Adjunto cotización")
    for nombre, contenido in adjuntos:
        m.add_attachment(contenido, maintype="application", subtype="pdf", filename=nombre)
    return m.as_bytes()


def test_pdfs_del_correo_filtra_por_palabras():
    crudo = _correo("Cotización 1234 cables", [("COT-1234.pdf", PDF)])
    r = correo.pdfs_del_mensaje(crudo, ["cotiz"])
    assert len(r) == 1 and r[0][0] == "2026-11-10 - Ventas Eléctrica Sur - COT-1234.pdf" and r[0][1] == PDF
    assert correo.pdfs_del_mensaje(_correo("Factura 99", [("F99.pdf", PDF)]), ["cotiz"]) == []
    assert len(correo.pdfs_del_mensaje(_correo("Factura 99", [("F99.pdf", PDF)]), [])) == 1  # [] = todos
    assert correo.pdfs_del_mensaje(_correo("Cotización", [("x.pdf", b"no es pdf")]), ["cotiz"]) == []


def test_guardar_pdf_no_duplica(tmp_path):
    assert correo.guardar_pdf(tmp_path, "a.pdf", PDF) == tmp_path / "a.pdf"
    assert correo.guardar_pdf(tmp_path, "a.pdf", PDF) is None
    assert correo.guardar_pdf(tmp_path, "a.pdf", PDF + b"x") == tmp_path / "a (2).pdf"


def test_servidor_automatico():
    assert correo.servidor_de({"correo_usuario": "yo@gmail.com"}) == "imap.gmail.com"
    assert correo.servidor_de({"correo_usuario": "yo@empresa.cl"}) == "imap.empresa.cl"
    assert correo.servidor_de({"correo_usuario": "yo@gmail.com", "correo_servidor": "mail.x.cl"}) == "mail.x.cl"


class _ImapFalso:
    def __init__(self, mensajes):
        self.mensajes = mensajes
        self.buscadas = []

    def select(self, carpeta, readonly):
        assert readonly
        return "OK", [str(len(self.mensajes)).encode()]

    def response(self, nombre):
        return nombre, [b"77"]

    def uid(self, comando, *args):
        if comando == "SEARCH":
            self.buscadas.append(args)
            return "OK", [b" ".join(self.mensajes)]
        assert comando == "FETCH" and args[1] == "(BODY.PEEK[])"  # nunca lo marca como leído
        return "OK", [(b"1 (BODY[] {1}", self.mensajes[args[0]])]

    def logout(self):
        pass


def test_revisar_correo_baja_solo_una_vez(tmp_path, monkeypatch):
    mensajes = {b"1": _correo("Cotización cables", [("COT-1.pdf", PDF)]),
                b"2": _correo("Hola", [("foto.pdf", PDF + b"2")])}
    imap = _ImapFalso(mensajes)
    monkeypatch.setattr(correo, "conectar", lambda cfg: imap)
    cfg = {"correo_usuario": "yo@gmail.com", "correo_clave": "x", "correo_palabras": ["cotiz"]}
    estado = tmp_path / "estado.json"
    assert correo.revisar(cfg, tmp_path, estado) == 1
    assert (tmp_path / "Correo" / "2026-11-10 - Ventas Eléctrica Sur - COT-1.pdf").exists()
    assert "X-GM-RAW" in imap.buscadas[0]
    assert correo.revisar(cfg, tmp_path, estado) == 0  # ya revisados
    assert len(json.loads(estado.read_text())["vistos"]) == 2


# --------------------------------------------------------------------------- programa anterior

def test_trae_configuracion_del_programa_anterior(tmp_path, monkeypatch):
    casa = tmp_path / "casa"
    cotiz = tmp_path / "Cotizaciones"
    cotiz.mkdir()
    antiguo = casa / "Downloads" / "buscador-cotizaciones-pc-main"
    antiguo.mkdir(parents=True)
    (antiguo / "buscador_cotizaciones.py").write_text("")
    (antiguo / "config.toml").write_text(f"carpeta_cotizaciones = '{cotiz}'\n", encoding="utf-8")
    (antiguo / "categorias.txt").write_text("Cables: cable, thhn, mio\n", encoding="utf-8")
    vacio = casa / "Documents" / "otro"  # otro programa sin carpeta configurada: no sirve
    vacio.mkdir(parents=True)
    (vacio / "buscador_cotizaciones.py").write_text("")
    (vacio / "config.toml").write_text("carpeta_cotizaciones = ''\n")
    inicio = tmp_path / "Startup"
    inicio.mkdir()
    destino = tmp_path / "instalado"
    destino.mkdir()
    origen = migracion.traer_configuracion(destino, inicio=inicio, escritorio=casa / "Desktop", casa=casa)
    assert origen == antiguo.resolve()
    assert str(cotiz) in (destino / "config.toml").read_text(encoding="utf-8")
    assert "mio" in (destino / "categorias.txt").read_text(encoding="utf-8")


def test_lee_ruta_del_inicio_antiguo(tmp_path):
    cmd = tmp_path / migracion.INICIO_ANTIGUO
    cmd.write_text('@start "" /d "C:\\Users\\Carlos\\Buscador\\" "C:\\Users\\Carlos\\Buscador\\.venv\\Scripts\\pythonw.exe" '
                   '"C:\\Users\\Carlos\\Buscador\\ventana.py" --minimizado', encoding="utf-8")
    rutas = {str(p).replace("/", "\\") for p in migracion._rutas_en_archivo(cmd)}
    assert "C:\\Users\\Carlos\\Buscador" in rutas


# --------------------------------------------------------------------------- PDF

def test_exportar_pdf(tmp_path):
    exportar_pdf = pytest.importorskip("exportar_pdf")
    it = _item("Cable THHN 12 AWG <rojo> & más", 450)
    filas = [(it, {"proveedor": "Prov A & Cía", "fecha": "2026-01-10", "moneda": "CLP", "numero": "1"}, "a.pdf")] * 120
    ruta = tmp_path / "x.pdf"
    exportar_pdf.precios_pdf(ruta, filas, "búsqueda \"cable\"", {id(it): (0.1, None)})
    assert ruta.read_bytes().startswith(b"%PDF") and ruta.stat().st_size > 2000


def test_temporal_del_correo_no_se_lee(tmp_path):
    (tmp_path / "~a.pdf").write_bytes(PDF)
    os.utime(tmp_path / "~a.pdf", (time.time() - 60,) * 2)
    assert bc.buscar_pdfs(tmp_path, True) == []


def test_no_acepta_disco_completo_ni_carpeta_de_usuario(tmp_path):
    raiz = Path(tmp_path.anchor)
    assert "disco completo" in bc.problema_carpeta(raiz)
    assert "usuario" in bc.problema_carpeta(Path.home())
    assert bc.problema_carpeta(bc.BASE)
    assert bc.problema_carpeta(tmp_path) == ""


# --------------------------------------------------------------------------- servidor de correo de empresa

SALIDA_NSLOOKUP = """Servidor:  router
Address:  192.168.1.1

empresa-ejemplo.cl\tMX preference = 1, mail exchanger = aspmx.l.google.com
empresa-ejemplo.cl\tMX preference = 5, mail exchanger = alt1.aspmx.l.google.com
"""


def test_mx_y_candidatos(monkeypatch):
    import subprocess

    correo._MX.clear()
    monkeypatch.setattr(subprocess, "run", lambda *a, **k: subprocess.CompletedProcess(a, 0, SALIDA_NSLOOKUP, ""))
    cfg = {"correo_usuario": "yo@empresa-ejemplo.cl"}
    assert correo._mx("empresa-ejemplo.cl") == ["aspmx.l.google.com", "alt1.aspmx.l.google.com"]
    lista = correo.candidatos(cfg)
    assert lista[0] == "imap.gmail.com"  # Google Workspace
    assert "imap.empresa-ejemplo.cl" in lista and "mail.empresa-ejemplo.cl" in lista
    assert correo.candidatos({"correo_usuario": "yo@gmail.com"}) == ["imap.gmail.com"]
    assert correo.candidatos({"correo_usuario": "yo@x.cl", "correo_servidor": "srv.x.cl"}) == ["srv.x.cl"]
    correo._MX.clear()


def test_conectar_prueba_hasta_encontrar(monkeypatch):
    import socket

    correo._MX.clear()
    monkeypatch.setattr(correo, "_mx", lambda d: [])
    probados = []

    class Falso:
        def __init__(self, host, *a, **k):
            probados.append(host)
            if host != "mail.empresa-ejemplo.cl":
                raise socket.gaierror(11001, "getaddrinfo failed")

        def login(self, u, c):
            assert c == "abcdabcdabcdabcd"

    monkeypatch.setattr(correo.imaplib, "IMAP4_SSL", Falso)
    cfg = {"correo_usuario": "yo@empresa-ejemplo.cl", "correo_clave": "abcd abcd abcd abcd"}
    correo.conectar(cfg)
    assert probados == ["imap.empresa-ejemplo.cl", "mail.empresa-ejemplo.cl"]
    assert cfg["_servidor_detectado"] == "mail.empresa-ejemplo.cl"
    assert correo.servidor_de(cfg) == "mail.empresa-ejemplo.cl"


def test_conectar_sin_servidor_explica_que_hacer(monkeypatch):
    import socket

    monkeypatch.setattr(correo, "_mx", lambda d: [])

    def falla(*a, **k):
        raise socket.gaierror(11001, "getaddrinfo failed")

    monkeypatch.setattr(correo.imaplib, "IMAP4_SSL", falla)
    with pytest.raises(correo.ErrorCorreo, match="servidor IMAP"):
        correo.conectar({"correo_usuario": "yo@empresa-ejemplo.cl", "correo_clave": "x"})


# --------------------------------------------------------------------------- origen: correo o PC

def test_origen_correo_o_pc():
    assert correo.origen("Correo/2026-11-10 - Ventas - COT-1.pdf") == "Correo"
    assert correo.origen("2026-11-10 - Ventas Ejemplo - COT-1.pdf") == "Correo"  # sin subcarpetas
    assert correo.origen("2026/Octubre/Cotizacion 4581.pdf") == "PC"
    assert correo.origen("Cotizacion 4581.pdf") == "PC"


def test_excel_columna_origen(tmp_path):
    from openpyxl import load_workbook

    from excel_writer import construir_excel

    db = {"a.pdf": _reg("Prov A", "2026-01-10", [_item("Cable THHN 12 AWG", 450)]),
          "Correo/2026-03-01 - Ventas - b.pdf": _reg("Prov B", "2026-03-01", [_item("Cable THHN 12 AWG", 540)], "2")}
    construir_excel(db, tmp_path, tmp_path / "Base.xlsx")
    wb = load_workbook(tmp_path / "Base.xlsx")
    for hoja in ("Precios", "Cotizaciones"):
        ws = wb[hoja]
        cab = [c.value for c in ws[1]]
        col = cab.index("Origen") + 1
        valores = {ws.cell(r, cab.index("Empresa" if hoja == "Precios" else "Proveedor") + 1).value: ws.cell(r, col).value
                   for r in range(2, ws.max_row + 1)}
        assert valores == {"Prov A": "PC", "Prov B": "Correo"}, hoja
    assert wb["Precios"].cell(2, [c.value for c in wb["Precios"][1]].index("PDF") + 1).hyperlink is not None
