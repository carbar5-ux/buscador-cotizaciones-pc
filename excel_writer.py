"""Genera el Excel base de cotizaciones a partir de la base de datos JSON."""

from __future__ import annotations

import os
import statistics
import tempfile
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from extractor import Cotizacion, normalizar

AZUL = "1F4E78"
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
    return FMT_ENTERO if moneda in ("CLP", "") else FMT_DECIMAL


def _hoja(wb: Workbook, titulo: str, columnas: list[tuple[str, int, str | None]], filas: list[list], nombre_tabla: str):
    """Crea una hoja con tabla de Excel (filtros, bandas), anchos y formatos."""
    ws = wb.create_sheet(titulo)
    ws.append([c[0] for c in columnas])
    for fila in filas:
        # un texto que empieza con "=" Excel lo tomaría como fórmula
        ws.append([" " + v if isinstance(v, str) and v.startswith(("=", "+", "@")) and not v[1:2].isdigit() else v
                   for v in fila])
    for idx, (_, ancho, fmt) in enumerate(columnas, start=1):
        letra = get_column_letter(idx)
        ws.column_dimensions[letra].width = ancho
        if fmt:
            for celda in ws[letra][1:]:
                if fmt != "moneda":
                    celda.number_format = fmt
    for celda in ws[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor=AZUL)
        celda.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "B2" if titulo != "Resumen" else None
    if filas:
        ref = f"A1:{get_column_letter(len(columnas))}{len(filas) + 1}"
        tabla = Table(displayName=nombre_tabla, ref=ref)
        tabla.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tabla)
    return ws


AZUL_LINK = Font(color="0563C1", underline="single")
AZUL_BOTON = Font(color="0563C1", underline="single", bold=True)


def _vincular(ws, fila: int, col: int, ruta: Path, texto: str, fuente: Font = AZUL_LINK) -> None:
    """Celda que abre el PDF con un clic (fórmula HIPERVINCULO: funciona con espacios, tildes y subcarpetas)."""
    celda = ws.cell(fila, col)
    destino = str(ruta)
    if len(destino) <= 250 and len(texto) <= 250:
        celda.value = '=HYPERLINK("{}","{}")'.format(destino.replace('"', ""), texto.replace('"', '""'))
    else:  # ruta muy larga para una fórmula: vínculo directo
        celda.value = texto
        celda.hyperlink = destino
    celda.font = fuente


def construir_excel(db: dict[str, dict], carpeta: Path, destino: Path, dias_aviso: int = 7) -> None:
    hoy = date.today()
    cotizaciones: list[tuple[str, Cotizacion, dict]] = []
    for clave, reg in db.items():
        cotizaciones.append((clave, Cotizacion.from_dict(reg["datos"]), reg))
    cotizaciones.sort(key=lambda t: (t[1].fecha or "0000", t[1].proveedor), reverse=True)

    wb = Workbook()
    wb.remove(wb.active)

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
            ws.cell(r, col).number_format = fmt
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
        _vincular(ws, r, 18, carpeta / clave, Path(clave).name)
        ws.cell(r, 13).number_format = fmt
        ws.cell(r, 15).number_format = fmt
        ws.cell(r, 10).alignment = Alignment(wrap_text=True, vertical="top")
    if filas_it:
        ws.conditional_formatting.add(
            f"P2:P{len(filas_it) + 1}", FormulaRule(formula=['LEFT(P2,7)="Revisar"'], fill=AMARILLO)
        )

    # ---------------- Hoja Comparativo de precios ----------------
    grupos: dict[tuple, list] = defaultdict(list)
    for clave, c, _ in cotizaciones:
        for it in c.items:
            if it.precio_unitario is None or not it.descripcion:
                continue
            # por descripción: los códigos cambian de un proveedor a otro
            clave_g = (normalizar(it.descripcion), normalizar(it.unidad), c.moneda)
            grupos[clave_g].append((it, c, clave))
    col_cmp = [
        ("Descripción", 55, None), ("Códigos", 16, None), ("Unidad", 8, None), ("Moneda", 8, None),
        ("N° cotizaciones", 11, FMT_ENTERO), ("N° proveedores", 11, FMT_ENTERO),
        ("Precio mínimo", 15, "moneda"), ("Proveedor más barato", 28, None), ("Fecha mínimo", 12, FMT_FECHA),
        ("Precio máximo", 15, "moneda"), ("Proveedor más caro", 28, None), ("Promedio", 15, "moneda"),
        ("Mediana", 15, "moneda"), ("Dif. máx/mín", 11, FMT_PCT), ("Último precio", 15, "moneda"),
        ("Último proveedor", 28, None), ("Fecha último", 12, FMT_FECHA), ("Ahorro vs último", 11, FMT_PCT),
    ]
    filas_cmp, fmts_cmp, links_cmp = [], [], []
    for (_, unidad, moneda), lista3 in grupos.items():
        lista = [(it, c) for it, c, _ in lista3]
        precios = [it.precio_unitario for it, _ in lista]
        mn = min(lista, key=lambda t: t[0].precio_unitario)
        mx = max(lista, key=lambda t: t[0].precio_unitario)
        ult = max(lista, key=lambda t: t[1].fecha or "")
        provs = {c.rut_proveedor or c.proveedor for _, c in lista}
        p_min, p_max, p_ult = mn[0].precio_unitario, mx[0].precio_unitario, ult[0].precio_unitario
        filas_cmp.append([
            mn[0].descripcion, ", ".join(sorted({it.codigo for it, _ in lista if it.codigo})), mn[0].unidad, moneda, len(lista), len(provs),
            p_min, mn[1].proveedor, _fecha(mn[1].fecha), p_max, mx[1].proveedor, statistics.fmean(precios),
            statistics.median(precios), (p_max / p_min - 1) if p_min else None, p_ult, ult[1].proveedor,
            _fecha(ult[1].fecha), (1 - p_min / p_ult) if p_ult else None,
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
            ws.cell(r, col).number_format = fmt
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

    # ---------------- Hoja Resumen (primera) ----------------
    ws = wb.create_sheet("Resumen", 0)
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 26
    ws["A1"] = "Base de cotizaciones"
    ws["A1"].font = Font(bold=True, size=16, color=AZUL)
    ws["A2"] = f"Carpeta: {carpeta}"
    ws["A3"] = f"Actualizado: {datetime.now():%d-%m-%Y %H:%M}"
    ws["A2"].font = ws["A3"].font = Font(italic=True, color="666666")
    estados = [f[11] for f in filas_cot]
    total_clp = sum(c.total or 0 for _, c, _ in cotizaciones if c.moneda == "CLP")
    indicadores = [
        ("Cotizaciones leídas", len(cotizaciones), FMT_ENTERO),
        ("Proveedores distintos", len(prov), FMT_ENTERO),
        ("Líneas de detalle (ítems)", len(filas_it), FMT_ENTERO),
        ("Productos/servicios comparables (2+ cotizaciones)", sum(1 for f in filas_cmp if f[4] > 1), FMT_ENTERO),
        ("Monto total cotizado (CLP, con IVA)", total_clp, FMT_ENTERO),
        ("Cotizaciones vigentes", estados.count("Vigente"), FMT_ENTERO),
        ("Cotizaciones por vencer", estados.count("Por vencer"), FMT_ENTERO),
        ("Cotizaciones vencidas", estados.count("Vencida"), FMT_ENTERO),
        ("Observaciones por revisar", len(filas_rev), FMT_ENTERO),
    ]
    borde = Border(bottom=Side(style="thin", color="DDDDDD"))
    for i, (nombre, valor, fmt) in enumerate(indicadores, start=5):
        ws.cell(i, 1, nombre).border = borde
        c = ws.cell(i, 2, valor)
        c.number_format = fmt
        c.font = Font(bold=True, size=12)
        c.border = borde
    por_cat: dict[str, int] = defaultdict(int)
    for _, c, _ in cotizaciones:
        for cat in c.categorias:
            por_cat[cat] += 1
    if por_cat:
        ws.cell(17, 1, "Cotizaciones por categoría").font = Font(bold=True, size=12, color=AZUL)
        for i, (cat, n) in enumerate(sorted(por_cat.items(), key=lambda t: -t[1]), start=18):
            ws.cell(i, 1, cat).border = borde
            celda = ws.cell(i, 2, n)
            celda.font = Font(bold=True)
            celda.border = borde
    ws.cell(15, 1, "Hojas: Cotizaciones (1 fila por PDF) · Ítems (1 fila por línea) · Comparativo (precios por producto) "
                   "· Proveedores · Revisar (lo que conviene verificar a mano)").font = Font(color="666666")

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
