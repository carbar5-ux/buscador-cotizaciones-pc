"""Exporta a PDF las líneas filtradas de la Planilla (para enviar por correo o imprimir)."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import LongTable, Paragraph, SimpleDocTemplate, Spacer, TableStyle

from correo import origen

VERDE_OSC = colors.HexColor("#0B6E3A")
GRIS = colors.HexColor("#6B7280")
ROJO_FONDO = colors.HexColor("#FEE2E2")
ROJO_TXT = colors.HexColor("#9C0006")


def _dinero(v, moneda: str) -> str:
    if v is None:
        return ""
    if moneda in ("CLP", ""):
        return "$ " + f"{v:,.0f}".replace(",", ".")
    return f"{moneda} " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _fecha(iso: str) -> str:
    return f"{iso[8:10]}-{iso[5:7]}-{iso[:4]}" if iso and len(iso) >= 10 else ""


def precios_pdf(ruta: Path, filas: list, filtros: str, var_precio: dict | None = None) -> None:
    """filas: [(ítem, datos de la cotización, clave)] como las muestra la Planilla."""
    var_precio = var_precio or {}
    chico = ParagraphStyle("chico", fontName="Helvetica", fontSize=7.5, leading=9)
    titulo = ParagraphStyle("titulo", fontName="Helvetica-Bold", fontSize=16, leading=19, textColor=VERDE_OSC)
    sub = ParagraphStyle("sub", fontName="Helvetica", fontSize=8.5, leading=11, textColor=GRIS)

    encabezado = ["Descripción", "Precio unit.", "Var.", "Cantidad", "Unidad", "Total", "Empresa", "Fecha", "N° cot.", "Origen"]
    datos = [encabezado]
    estilo = [
        ("BACKGROUND", (0, 0), (-1, 0), VERDE_OSC), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ALIGN", (1, 0), (5, -1), "RIGHT"), ("ALIGN", (7, 0), (7, -1), "CENTER"),
        ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.HexColor("#E5E7EB")),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
    ]
    totales: dict[str, float] = {}
    for n, (it, d, _clave) in enumerate(filas, start=1):
        mon = d.get("moneda", "")
        var = var_precio.get(id(it))
        datos.append([
            Paragraph(escape(str(it.get("descripcion") or "")), chico), _dinero(it.get("precio_unitario"), mon),
            f"{var[0]:+.0%}" if var else "", f"{it['cantidad']:g}" if it.get("cantidad") is not None else "",
            str(it.get("unidad") or ""), _dinero(it.get("total"), mon),
            Paragraph(escape(str(d.get("proveedor") or "")), chico), _fecha(d.get("fecha") or ""),
            str(d.get("numero") or ""), origen(_clave),
        ])
        if it.get("total") is not None:
            totales[mon] = totales.get(mon, 0) + it["total"]
        if origen(_clave) == "Correo":
            estilo.append(("TEXTCOLOR", (9, n), (9, n), colors.HexColor("#1D4ED8")))
        if var and var[0] >= 0.005:
            estilo += [("BACKGROUND", (2, n), (2, n), ROJO_FONDO), ("TEXTCOLOR", (2, n), (2, n), ROJO_TXT)]
        elif n % 2 == 0:
            estilo.append(("BACKGROUND", (0, n), (-1, n), colors.HexColor("#F9FAFB")))

    anchos = [8.0 * cm, 2.3 * cm, 1.2 * cm, 1.6 * cm, 1.3 * cm, 2.5 * cm, 4.9 * cm, 1.9 * cm, 1.8 * cm, 1.3 * cm]
    tabla = LongTable(datos, colWidths=anchos, repeatRows=1)
    tabla.setStyle(TableStyle(estilo))

    def pie(canvas, doc) -> None:
        canvas.saveState()
        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(GRIS)
        canvas.drawString(doc.leftMargin, 0.8 * cm, "Buscador de cotizaciones — Carlos")
        canvas.drawRightString(doc.pagesize[0] - doc.rightMargin, 0.8 * cm, f"Página {doc.page}")
        canvas.restoreState()

    doc = SimpleDocTemplate(str(ruta), pagesize=landscape(A4), leftMargin=1.2 * cm, rightMargin=1.2 * cm,
                            topMargin=1.2 * cm, bottomMargin=1.4 * cm, title="Precios filtrados", author="Carlos")
    resumen = (f"{len(filas):,} líneas de {len({f[2] for f in filas}):,} cotizaciones".replace(",", ".")
               + ("  ·  suma de totales " + "  +  ".join(_dinero(v, m) for m, v in totales.items()) if totales else ""))
    historia = [
        Paragraph("Precios de cotizaciones", titulo),
        Paragraph(escape(f"Filtros: {filtros}  ·  {resumen}  ·  generado el {datetime.now():%d-%m-%Y %H:%M}"), sub),
        Spacer(1, 0.35 * cm), tabla,
    ]
    doc.build(historia, onFirstPage=pie, onLaterPages=pie)
