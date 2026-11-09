"""Genera el Excel base de cotizaciones a partir de la base de datos JSON."""

from __future__ import annotations

import math
import os
import re
import statistics
from copy import copy
import tempfile
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.worksheet.table import Table, TableStyleInfo

from agrupador import agrupar, descripcion_representativa
from extractor import Cotizacion, normalizar

AZUL = "0B6E3A"  # verde de la ventana del buscador (encabezados y títulos)
VERDE_OSC, GRIS_TXT, GRIS_FONDO, BORDE_SUAVE = "0B6E3A", "6B7280", "F3F5F7", "E5E7EB"
FUENTE_XL = "Calibri"
FMT_FECHA = "dd-mm-yyyy"
FMT_ENTERO = "#,##0"
FMT_DECIMAL = "#,##0.00"
FMT_PCT = "0.0%"

ROJO = PatternFill("solid", fgColor="F8CBAD")
AMARILLO = PatternFill("solid", fgColor="FFE699")
VERDE = PatternFill("solid", fgColor="C6EFCE")


def _fecha(iso: str) -> date | None:
    try:
        return date.fromisoformat(iso) if iso else None
    except ValueError:
        return None


def _fmt_moneda(moneda: str) -> str:
    """Formato de moneda real de Excel: $ 1.234.567 / US$ 1.234,50 / UF 12,35."""
    return {
        "CLP": '"$" #,##0', "": '"$" #,##0', "USD": '"US$" #,##0.00', "UF": '"UF" #,##0.00', "EUR": '"€" #,##0.00',
    }.get(moneda, FMT_DECIMAL)


CENTRADAS = {"Fecha", "Unidad", "Moneda", "Estado", "Vence", "Ítem", "Página", "Páginas", "N° ítems", "ID", "ID cot.",
             "PDF", "Lectura", "N° cotizaciones", "N° proveedores", "N° variantes", "Fecha mínimo", "Fecha último",
             "Última cotización", "Cantidad", "N° cotización"}


# Caché de estilos: aplicar un estilo en openpyxl es lento (compara objetos); cada combinación se calcula
# una vez y luego se copia (miles de veces más rápido). Se vacía con cada libro nuevo.
_CACHE_ESTILOS: dict[tuple, object] = {}
LETRA_CUERPO = Font(name=FUENTE_XL, size=10)
CENTRO = Alignment(horizontal="center", vertical="center")
DERECHA = Alignment(horizontal="right", vertical="center")
MEDIO = Alignment(vertical="center")
_FUENTES_VINCULO: dict[bool, Font] = {
    False: Font(name=FUENTE_XL, size=10, color="0563C1", underline="single"),
    True: Font(name=FUENTE_XL, size=10, color="0563C1", underline="single", bold=True),
}


def _estilo(celda, **atributos) -> None:
    base = tuple(celda._style) if celda._style is not None else ()
    # los objetos de estilo son constantes del módulo: basta su identidad (mucho más rápido que compararlos)
    clave = (base, tuple((k, v if isinstance(v, str) else id(v)) for k, v in atributos.items()))
    guardado = _CACHE_ESTILOS.get(clave)
    if guardado is None:
        for k, v in atributos.items():
            setattr(celda, k, v)
        _CACHE_ESTILOS[clave] = copy(celda._style)
    else:
        celda._style = copy(guardado)


def _hoja(wb: Workbook, titulo: str, columnas: list[tuple[str, int, str | None]], filas: list[list], nombre_tabla: str):
    """Crea una hoja con tabla de Excel (filtros, bandas), anchos y formatos."""
    ws = wb.create_sheet(titulo)
    ws.append([c[0] for c in columnas])
    for fila in filas:
        # textos limpios (algunos PDF traen caracteres de control que Excel no acepta) y un texto que
        # empieza con "=" Excel lo tomaría como fórmula
        fila = [limpiar_texto(v) for v in fila]
        ws.append([" " + v if isinstance(v, str) and v.startswith(("=", "+", "@")) and not v[1:2].isdigit() else v
                   for v in fila])
    for idx, (_, ancho, _fmt) in enumerate(columnas, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = ancho
    # aspecto: letra limpia, formato de número por columna, centradas/derecha donde corresponde
    centradas = {i for i, c in enumerate(columnas, start=1) if c[0] in CENTRADAS}
    numericas = {i for i, c in enumerate(columnas, start=1) if c[2] in ("moneda", FMT_DECIMAL, FMT_ENTERO, FMT_PCT)}
    formatos = {i: c[2] for i, c in enumerate(columnas, start=1) if c[2] and c[2] != "moneda"}
    centro, derecha, medio = CENTRO, DERECHA, MEDIO
    for fila_celdas in ws.iter_rows(min_row=2):
        for celda in fila_celdas:
            col = celda.column
            fmt = formatos.get(col)
            # cantidades enteras sin ",00"
            if fmt == FMT_DECIMAL and isinstance(celda.value, float) and celda.value.is_integer():
                fmt = FMT_ENTERO
            if col in centradas or col in numericas or fmt:  # el resto usa el estilo por defecto del libro
                alin = centro if col in centradas else derecha if col in numericas else medio
                if fmt:
                    _estilo(celda, alignment=alin, number_format=fmt)
                else:
                    _estilo(celda, alignment=alin)
    for celda in ws[1]:
        celda.font = Font(name=FUENTE_XL, bold=True, color="FFFFFF", size=10)
        celda.fill = PatternFill("solid", fgColor=VERDE_OSC)
        lado = "center" if celda.column in centradas else "right" if celda.column in numericas else "left"
        celda.alignment = Alignment(wrap_text=True, vertical="center", horizontal=lado)
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "B2"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 100
    # impresión: horizontal, todas las columnas en el ancho de la hoja, encabezado repetido
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:1"
    if filas:
        ref = f"A1:{get_column_letter(len(columnas))}{len(filas) + 1}"
        tabla = Table(displayName=nombre_tabla, ref=ref)
        tabla.tableStyleInfo = TableStyleInfo(name="TableStyleLight15", showRowStripes=True)
        ws.add_table(tabla)
    return ws


AZUL_LINK = Font(color="0563C1", underline="single")
AZUL_BOTON = Font(color="0563C1", underline="single", bold=True)


# Todo lo que XML 1.0 no admite (Excel da "Carácter xml no válido" y quita la hoja entera): controles,
# controles C1 (\x7f-\x9f), mitades de emoji sueltas (sustitutos) y \ufffe / \uffff. El filtro de openpyxl
# solo quita los controles ASCII.
_NO_XML = re.compile("[^\t\n\r\u0020-\u007e\u00a0-\ud7ff\ue000-\ufffd\U00010000-\U0010ffff]")
MAX_TEXTO_CELDA = 32767  # límite de caracteres de una celda en Excel


def limpiar_texto(v):
    """Deja el valor listo para una celda: sin caracteres que Excel no acepta y sin números NaN/infinito."""
    if isinstance(v, str):
        return _NO_XML.sub("", v)[:MAX_TEXTO_CELDA]
    if isinstance(v, float) and not math.isfinite(v):
        return None
    return v


# Carpeta del Excel (los vínculos son relativos a ella) y conteo por hoja: Excel admite hasta 65.530 por hoja.
_BASE_VINCULOS: list[Path] = [Path(".")]
_VINCULOS_POR_HOJA: dict[str, int] = defaultdict(int)
MAX_VINCULOS_HOJA = 65000


def _destino_vinculo(ruta: Path) -> str:
    """Dirección del PDF como la escribe Excel: relativa a la carpeta del Excel si se puede (sigue
    funcionando si se mueve la carpeta o cambia la letra del disco); si no, absoluta file:///."""
    try:
        rel = os.path.relpath(ruta, _BASE_VINCULOS[0])
        texto = rel.replace("\\", "/")
        prefijo = ""
    except ValueError:  # otro disco
        texto = str(ruta).replace("\\", "/")
        prefijo = "file:///"
    return prefijo + texto.replace("%", "%25").replace(" ", "%20").replace("#", "%23")


def _vincular(ws, fila: int, col: int, ruta: Path, texto: str, fuente: Font = AZUL_LINK) -> None:
    """Celda con el texto de siempre que, con un clic, abre el PDF (vínculo real de Excel, no fórmula:
    funciona aunque Excel tenga el cálculo en manual)."""
    celda = ws.cell(fila, col)
    celda.value = limpiar_texto(texto)
    if _VINCULOS_POR_HOJA[ws.title] >= MAX_VINCULOS_HOJA:
        return
    _VINCULOS_POR_HOJA[ws.title] += 1
    celda.hyperlink = _destino_vinculo(ruta)
    _estilo(celda, font=_FUENTES_VINCULO[bool(fuente.bold)])


def construir_excel(db: dict[str, dict], carpeta: Path, destino: Path, dias_aviso: int = 7) -> None:
    hoy = date.today()
    cotizaciones: list[tuple[str, Cotizacion, dict]] = []
    for clave, reg in db.items():
        cotizaciones.append((clave, Cotizacion.from_dict(reg["datos"]), reg))
    cotizaciones.sort(key=lambda t: (t[1].fecha or "0000", t[1].proveedor), reverse=True)

    wb = Workbook()
    wb.remove(wb.active)
    wb._fonts = IndexedList([LETRA_CUERPO])  # letra por defecto de todo el libro: Calibri 10
    _BASE_VINCULOS[0] = destino.parent.resolve() if destino.parent.exists() else destino.parent
    _CACHE_ESTILOS.clear()
    _VINCULOS_POR_HOJA.clear()

    # ---------------- Hoja Cotizaciones ----------------
    col_cot = [
        ("ID", 8, None), ("PDF", 11, None), ("Archivo", 34, None), ("Carpeta", 22, None), ("Proveedor", 32, None),
        ("RUT proveedor", 14, None), ("Cliente", 26, None), ("N° cotización", 14, None),
        ("Fecha", 12, FMT_FECHA), ("Validez", 18, None), ("Vence", 12, FMT_FECHA), ("Estado", 12, None),
        ("Moneda", 8, None), ("Neto", 15, "moneda"), ("Descuento", 13, "moneda"), ("IVA", 13, "moneda"),
        ("Total", 15, "moneda"), ("Suma ítems", 15, "moneda"), ("N° ítems", 9, FMT_ENTERO),
        ("Forma de pago", 22, None), ("Plazo de entrega", 22, None), ("Lugar de entrega", 22, None),
        ("Contacto", 22, None), ("Email", 26, None), ("Teléfono", 16, None), ("Páginas", 8, None),
        ("Lectura", 10, None), ("Observaciones", 50, None), ("Procesado", 17, "dd-mm-yyyy hh:mm"),
        ("Categorías", 30, None),
    ]
    filas_cot, ids = [], {}
    for n, (clave, c, reg) in enumerate(cotizaciones, start=1):
        ids[clave] = n
        vence = _fecha(c.vence)
        estado = "" if not vence else ("Vencida" if vence < hoy else "Por vencer" if (vence - hoy).days <= dias_aviso else "Vigente")
        suma = sum(i.total for i in c.items if i.total is not None) if c.items else None
        rel = Path(clave)
        filas_cot.append([
            n, "Abrir PDF", rel.name, str(rel.parent) if str(rel.parent) != "." else "", c.proveedor, c.rut_proveedor,
            c.cliente,
            c.numero, _fecha(c.fecha), c.validez, vence, estado, c.moneda, c.neto, c.descuento, c.iva, c.total,
            suma, len(c.items), c.forma_pago, c.plazo_entrega, c.lugar_entrega, c.contacto, c.email, c.telefono,
            c.paginas, c.metodo, "; ".join(c.observaciones), datetime.fromisoformat(reg["procesado"]),
            ", ".join(c.categorias),
        ])
    ws = _hoja(wb, "Cotizaciones", col_cot, filas_cot, "TablaCotizaciones")
    for r, (clave, c, _) in enumerate(cotizaciones, start=2):
        _vincular(ws, r, 2, carpeta / clave, "Abrir PDF", AZUL_BOTON)
        _vincular(ws, r, 3, carpeta / clave, Path(clave).name)
        fmt = _fmt_moneda(c.moneda)
        for col in (14, 15, 16, 17, 18):
            _estilo(ws.cell(r, col), number_format=fmt)
        estado = ws.cell(r, 12)
        estado.fill = {"Vencida": ROJO, "Por vencer": AMARILLO, "Vigente": VERDE}.get(estado.value, PatternFill())
        if c.observaciones:
            ws.cell(r, 28).fill = AMARILLO

    # ---------------- Hoja Ítems ----------------
    col_it = [
        ("ID cot.", 8, None), ("PDF", 11, None), ("Proveedor", 28, None), ("RUT proveedor", 14, None),
        ("N° cotización", 13, None), ("Fecha", 12, FMT_FECHA), ("Moneda", 8, None), ("Ítem", 6, None), ("Código", 14, None),
        ("Descripción", 55, None), ("Cantidad", 11, FMT_DECIMAL), ("Unidad", 8, None),
        ("Precio unitario", 15, "moneda"), ("Descuento", 10, FMT_DECIMAL), ("Total línea", 15, "moneda"),
        ("Control", 26, None), ("Página", 7, None), ("Archivo", 34, None), ("Categoría", 24, None),
    ]
    filas_it, fmts_it, claves_it = [], [], []
    for clave, c, _ in cotizaciones:
        for it in c.items:
            claves_it.append(clave)
            filas_it.append([
                ids[clave], "Abrir PDF", c.proveedor, c.rut_proveedor, c.numero, _fecha(c.fecha), c.moneda, it.item, it.codigo,
                it.descripcion, it.cantidad, it.unidad, it.precio_unitario, it.descuento, it.total, it.control,
                it.pagina, Path(clave).name, it.categoria,
            ])
            fmts_it.append(_fmt_moneda(c.moneda))
    ws = _hoja(wb, "Ítems", col_it, filas_it, "TablaItems")
    for r, (fmt, clave) in enumerate(zip(fmts_it, claves_it), start=2):
        _vincular(ws, r, 2, carpeta / clave, "Abrir PDF", AZUL_BOTON)
        desc = str(ws.cell(r, 10).value or "").strip() or "(sin descripción)"
        _vincular(ws, r, 10, carpeta / clave, desc)  # clic en el suministro = abre su cotización
        _estilo(ws.cell(r, 13), number_format=fmt)
        _estilo(ws.cell(r, 15), number_format=fmt)
    if filas_it:
        ws.conditional_formatting.add(
            f"P2:P{len(filas_it) + 1}", FormulaRule(formula=['LEFT(P2,7)="Revisar"'], fill=AMARILLO)
        )

    # ---------------- Hoja Comparativo de precios ----------------
    # mismo producto aunque esté escrito distinto (medidas exactas, palabras parecidas, sinónimos)
    elementos = [(it.descripcion, it.unidad, c.moneda, (it, c, clave))
                 for clave, c, _ in cotizaciones for it in c.items
                 if it.precio_unitario is not None and it.descripcion]
    grupos_lista = agrupar(elementos)
    col_cmp = [
        ("Descripción", 55, None), ("Códigos", 16, None), ("Unidad", 8, None), ("Moneda", 8, None),
        ("N° cotizaciones", 11, FMT_ENTERO), ("N° proveedores", 11, FMT_ENTERO),
        ("Precio mínimo", 15, "moneda"), ("Proveedor más barato", 28, None), ("Fecha mínimo", 12, FMT_FECHA),
        ("Precio máximo", 15, "moneda"), ("Proveedor más caro", 28, None), ("Promedio", 15, "moneda"),
        ("Mediana", 15, "moneda"), ("Dif. máx/mín", 11, FMT_PCT), ("Último precio", 15, "moneda"),
        ("Último proveedor", 28, None), ("Fecha último", 12, FMT_FECHA), ("Ahorro vs último", 11, FMT_PCT),
        ("N° variantes", 10, FMT_ENTERO), ("Variantes (cómo lo escribe cada proveedor)", 70, None),
    ]
    filas_cmp, fmts_cmp, links_cmp = [], [], []
    for lista3 in grupos_lista:
        moneda = lista3[0][1].moneda
        lista = [(it, c) for it, c, _ in lista3]
        variantes = list(dict.fromkeys(it.descripcion.strip() for it, _ in lista))
        precios = [it.precio_unitario for it, _ in lista]
        mn = min(lista, key=lambda t: t[0].precio_unitario)
        mx = max(lista, key=lambda t: t[0].precio_unitario)
        ult = max(lista, key=lambda t: t[1].fecha or "")
        provs = {c.rut_proveedor or c.proveedor for _, c in lista}
        p_min, p_max, p_ult = mn[0].precio_unitario, mx[0].precio_unitario, ult[0].precio_unitario
        filas_cmp.append([
            descripcion_representativa([it.descripcion for it, _ in lista]), ", ".join(sorted({it.codigo for it, _ in lista if it.codigo})), mn[0].unidad, moneda, len(lista), len(provs),
            p_min, mn[1].proveedor, _fecha(mn[1].fecha), p_max, mx[1].proveedor, statistics.fmean(precios),
            statistics.median(precios), (p_max / p_min - 1) if p_min else None, p_ult, ult[1].proveedor,
            _fecha(ult[1].fecha), (1 - p_min / p_ult) if p_ult else None,
            len(variantes), " | ".join(variantes[:8]) + (" | …" if len(variantes) > 8 else ""),
        ])
        fmts_cmp.append(_fmt_moneda(moneda))
        clave_de = {id(t[1]): t[2] for t in lista3}
        links_cmp.append((clave_de[id(mn[1])], clave_de[id(mx[1])], clave_de[id(ult[1])]))
    orden = sorted(range(len(filas_cmp)), key=lambda i: (-filas_cmp[i][5], -filas_cmp[i][4], filas_cmp[i][0]))
    filas_cmp = [filas_cmp[i] for i in orden]
    fmts_cmp = [fmts_cmp[i] for i in orden]
    links_cmp = [links_cmp[i] for i in orden]
    ws = _hoja(wb, "Comparativo", col_cmp, filas_cmp, "TablaComparativo")
    for r, (fmt, (k_min, k_max, k_ult)) in enumerate(zip(fmts_cmp, links_cmp), start=2):
        for col in (7, 10, 12, 13, 15):
            _estilo(ws.cell(r, col), number_format=fmt)
        # clic en el proveedor = abre su cotización
        for col, k in ((8, k_min), (11, k_max), (16, k_ult)):
            _vincular(ws, r, col, carpeta / k, str(ws.cell(r, col).value or Path(k).name))
    if filas_cmp:
        ws.conditional_formatting.add(
            f"N2:N{len(filas_cmp) + 1}", CellIsRule(operator="greaterThan", formula=["0.2"], fill=AMARILLO)
        )

    # ---------------- Hoja Proveedores ----------------
    prov: dict[str, dict] = {}
    for clave, c, _ in cotizaciones:
        k = c.rut_proveedor or normalizar(c.proveedor) or "(sin identificar)"
        p = prov.setdefault(k, {"nombre": c.proveedor, "rut": c.rut_proveedor, "n": 0, "items": 0, "montos": defaultdict(float),
                                "ultima": "", "email": "", "fono": "", "contacto": ""})
        p["n"] += 1
        p["items"] += len(c.items)
        if c.total is not None:
            p["montos"][c.moneda] += c.total
        if (c.fecha or "") >= p["ultima"]:
            p["ultima"] = c.fecha or p["ultima"]
            p["nombre"] = c.proveedor or p["nombre"]
        p["email"] = p["email"] or c.email
        p["fono"] = p["fono"] or c.telefono
        p["contacto"] = p["contacto"] or c.contacto
    col_prov = [
        ("Proveedor", 34, None), ("RUT", 14, None), ("N° cotizaciones", 12, FMT_ENTERO), ("N° ítems", 10, FMT_ENTERO),
        ("Total cotizado CLP", 17, FMT_ENTERO), ("Total cotizado otras monedas", 26, None), ("Última cotización", 14, FMT_FECHA),
        ("Contacto", 24, None), ("Email", 28, None), ("Teléfono", 16, None),
    ]
    filas_prov = [
        [p["nombre"], p["rut"], p["n"], p["items"], p["montos"].get("CLP") or None,
         ", ".join(f"{m} {v:,.2f}" for m, v in p["montos"].items() if m != "CLP"), _fecha(p["ultima"]),
         p["contacto"], p["email"], p["fono"]]
        for p in sorted(prov.values(), key=lambda p: -p["n"])
    ]
    _hoja(wb, "Proveedores", col_prov, filas_prov, "TablaProveedores")

    # ---------------- Hoja Revisar ----------------
    filas_rev, claves_rev = [], []
    for clave, c, _ in cotizaciones:
        for obs in c.observaciones:
            filas_rev.append([ids[clave], Path(clave).name, c.proveedor, obs])
            claves_rev.append(clave)
        n_rev = sum(1 for i in c.items if i.control.startswith("Revisar"))
        if n_rev:
            filas_rev.append([ids[clave], Path(clave).name, c.proveedor, f"{n_rev} línea(s) donde cantidad × precio ≠ total"])
            claves_rev.append(clave)
    vistos: dict[str, str] = {}
    for clave, _, reg in cotizaciones:
        if reg["hash"] in vistos:
            filas_rev.append([ids[clave], Path(clave).name, "", f"Archivo duplicado de {vistos[reg['hash']]}"])
            claves_rev.append(clave)
        else:
            vistos[reg["hash"]] = clave
    ws = _hoja(wb, "Revisar", [("ID cot.", 8, None), ("Archivo", 36, None), ("Proveedor", 30, None),
                               ("Observación", 70, None)], filas_rev, "TablaRevisar")
    for r, clave in enumerate(claves_rev, start=2):
        _vincular(ws, r, 2, carpeta / clave, Path(clave).name)

    # ---------------- Hoja Resumen (tarjetas) ----------------
    ws = wb.create_sheet("Resumen", 0)
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.column_dimensions["A"].width = 3
    for col in "BCDEFGHI":
        ws.column_dimensions[col].width = 17
    ws["B2"] = "Base de cotizaciones"
    ws["B2"].font = Font(name=FUENTE_XL, bold=True, size=22, color=VERDE_OSC)
    ws["B3"] = limpiar_texto(f"por Carlos  ·  actualizado {datetime.now():%d-%m-%Y %H:%M}  ·  {carpeta}")
    ws["B3"].font = Font(name=FUENTE_XL, size=10, color=GRIS_TXT)
    estados = [f[11] for f in filas_cot]
    total_clp = sum(c.total or 0 for _, c, _ in cotizaciones if c.moneda == "CLP")
    tarjetas = [
        ("Cotizaciones", len(cotizaciones), FMT_ENTERO, VERDE_OSC),
        ("Proveedores", len(prov), FMT_ENTERO, VERDE_OSC),
        ("Líneas de precio", len(filas_it), FMT_ENTERO, VERDE_OSC),
        ("Monto cotizado (CLP)", total_clp, '"$" #,##0', VERDE_OSC),
        ("Vigentes", estados.count("Vigente"), FMT_ENTERO, "16A34A"),
        ("Por vencer", estados.count("Por vencer"), FMT_ENTERO, "D97706"),
        ("Vencidas", estados.count("Vencida"), FMT_ENTERO, "DC2626"),
        ("Por revisar", len(filas_rev), FMT_ENTERO, "DC2626"),
    ]
    fondo = PatternFill("solid", fgColor=GRIS_FONDO)
    linea = Side(style="thin", color=BORDE_SUAVE)
    for i, (titulo, valor, fmt, color) in enumerate(tarjetas):
        fila = 5 + (i // 4) * 4
        c1 = 2 + (i % 4) * 2  # cada tarjeta ocupa 2 columnas
        ws.merge_cells(start_row=fila, start_column=c1, end_row=fila, end_column=c1 + 1)
        ws.merge_cells(start_row=fila + 1, start_column=c1, end_row=fila + 1, end_column=c1 + 1)
        t = ws.cell(fila, c1, titulo)
        t.font = Font(name=FUENTE_XL, size=10, bold=True, color=GRIS_TXT)
        v = ws.cell(fila + 1, c1, valor)
        v.font = Font(name=FUENTE_XL, size=20, bold=True, color=color)
        v.number_format = fmt
        for r in (fila, fila + 1):
            for c in (c1, c1 + 1):
                celda = ws.cell(r, c)
                celda.fill = fondo
                celda.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                celda.border = Border(top=linea if r == fila else None, bottom=linea if r == fila + 1 else None,
                                      left=linea if c == c1 else None, right=linea if c == c1 + 1 else None)
        ws.row_dimensions[fila].height = 20
        ws.row_dimensions[fila + 1].height = 34

    por_cat: dict[str, int] = defaultdict(int)
    for _, c, _ in cotizaciones:
        for cat in c.categorias:
            por_cat[cat] += 1
    fila = 14
    if por_cat:
        ws.cell(fila, 2, "Cotizaciones por categoría").font = Font(name=FUENTE_XL, bold=True, size=13, color=VERDE_OSC)
        for i, (cat, n) in enumerate(sorted(por_cat.items(), key=lambda t: -t[1]), start=fila + 1):
            ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=4)
            ws.cell(i, 2, limpiar_texto(cat)).font = Font(name=FUENTE_XL, size=11)
            celda = ws.cell(i, 5, n)
            celda.font = Font(name=FUENTE_XL, size=11, bold=True)
            for c in range(2, 6):
                ws.cell(i, c).border = Border(bottom=linea)
        fila = fila + len(por_cat) + 2
    ws.cell(fila, 2, "Hojas").font = Font(name=FUENTE_XL, bold=True, size=13, color=VERDE_OSC)
    guia = [("Precios", "una fila por suministro: empresa, fecha, descripción, cantidad y precio (clic = abre el PDF)"),
            ("Cotizaciones", "una fila por PDF, con condiciones, totales y vigencia"),
            ("Ítems", "todas las líneas con todos sus datos"),
            ("Proveedores", "contacto y montos por proveedor"),
            ("Revisar", "lo que conviene verificar a mano"),
            ("Comparativo", "el mismo producto en distintas cotizaciones")]
    for i, (h, d) in enumerate(guia, start=fila + 1):
        ws.cell(i, 2, h).font = Font(name=FUENTE_XL, bold=True, size=10)
        ws.cell(i, 3, d).font = Font(name=FUENTE_XL, size=10, color=GRIS_TXT)

    # ---------------- Hoja Precios (la primera: lo esencial de cada línea) ----------------
    col_pre = [
        ("Empresa", 32, None), ("Fecha", 12, FMT_FECHA), ("Descripción", 60, None), ("Cantidad", 11, FMT_DECIMAL),
        ("Unidad", 9, None), ("Precio unitario", 16, "moneda"), ("Total", 16, "moneda"), ("Moneda", 8, None),
        ("N° cotización", 14, None), ("PDF", 11, None),
    ]
    filas_pre, extra_pre = [], []
    for clave, c, _ in cotizaciones:
        for it in c.items:
            filas_pre.append([c.proveedor, _fecha(c.fecha), it.descripcion, it.cantidad, it.unidad, it.precio_unitario,
                              it.total, c.moneda, c.numero, "Abrir PDF"])
            extra_pre.append((clave, _fmt_moneda(c.moneda)))
    ws = _hoja(wb, "Precios", col_pre, filas_pre, "TablaPrecios")
    wb.move_sheet(ws, offset=-(len(wb.sheetnames) - 1))  # primera hoja
    for r, (clave, fmt) in enumerate(extra_pre, start=2):
        _estilo(ws.cell(r, 6), number_format=fmt)
        _estilo(ws.cell(r, 7), number_format=fmt)
        # clic en el suministro (descripción) = abre su cotización
        desc = str(ws.cell(r, 3).value or "").strip() or "(sin descripción)"
        _vincular(ws, r, 3, carpeta / clave, desc)
        _vincular(ws, r, 10, carpeta / clave, "Abrir PDF", AZUL_BOTON)
    ws.freeze_panes = "A2"

    # orden de las hojas y color de cada pestaña
    orden = ["Precios", "Resumen", "Cotizaciones", "Ítems", "Proveedores", "Revisar", "Comparativo"]
    colores = {"Precios": "16A34A", "Resumen": "0B6E3A", "Cotizaciones": "E8700A", "Ítems": "0E7490",
               "Proveedores": "6B7280", "Revisar": "DC2626", "Comparativo": "9CA3AF"}
    wb._sheets.sort(key=lambda h: orden.index(h.title) if h.title in orden else len(orden))
    for h in wb.worksheets:
        h.sheet_properties.tabColor = colores.get(h.title, "9CA3AF")
    wb.active = 0

    _guardar_atomico(wb, destino)


def _guardar_atomico(wb: Workbook, destino: Path) -> None:
    """Escribe en un temporal y reemplaza. Lanza PermissionError si el Excel está abierto."""
    destino.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=destino.parent, prefix="~tmp_")
    os.close(fd)
    try:
        wb.save(tmp)
        os.replace(tmp, destino)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
