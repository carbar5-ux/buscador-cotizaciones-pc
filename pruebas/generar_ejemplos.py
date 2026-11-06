"""Crea cotizaciones PDF de ejemplo (formatos distintos) para probar el programa.

    python pruebas/generar_ejemplos.py CARPETA_DESTINO
"""

import sys
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


def miles(v: float) -> str:
    return f"{v:,.0f}".replace(",", ".")


def con_tabla(ruta: Path) -> None:
    """Cotización con tabla con bordes (la más común)."""
    est = getSampleStyleSheet()
    doc = SimpleDocTemplate(str(ruta), pagesize=letter)
    filas = [
        ("1", "CEM-25", "Cemento especial saco 25 kg", 120, "UN", 4990),
        ("2", "FIE-12", "Fierro estriado A630 12 mm x 6 m", 80, "UN", 8450),
        ("3", "PLA-OSB", "Plancha OSB 9,5 mm 1,22 x 2,44", 40, "UN", 13990),
        ("4", "", "Despacho a obra", 1, "GL", 45000),
    ]
    neto = sum(c * p for *_, c, _, p in filas)
    iva = round(neto * 0.19)
    datos = [["Ítem", "Código", "Descripción", "Cantidad", "Unidad", "Precio Unitario", "Total"]]
    datos += [[i, cod, d, str(c), u, f"$ {miles(p)}", f"$ {miles(c * p)}"] for i, cod, d, c, u, p in filas]
    t = Table(datos, colWidths=[30, 55, 190, 50, 40, 80, 80])
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.grey), ("FONTSIZE", (0, 0), (-1, -1), 8)]))
    elementos = [
        Paragraph("<b>Ferretería Los Andes SpA</b>", est["Title"]),
        Paragraph("RUT: 76.123.456-7 · Av. Matta 1234, Santiago · Fono: +56 2 2345 6789", est["Normal"]),
        Paragraph("ventas@ferreandes.cl", est["Normal"]),
        Spacer(1, 10),
        Paragraph("<b>COTIZACIÓN N° 4581</b>", est["Heading2"]),
        Paragraph("Fecha: 15/10/2026", est["Normal"]),
        Paragraph("Señores: Constructora Ejemplo Ltda. RUT 77.888.999-K", est["Normal"]),
        Spacer(1, 10), t, Spacer(1, 10),
        Paragraph(f"Subtotal: $ {miles(neto)}", est["Normal"]),
        Paragraph(f"IVA 19%: $ {miles(iva)}", est["Normal"]),
        Paragraph(f"Total: $ {miles(neto + iva)}", est["Normal"]),
        Spacer(1, 10),
        Paragraph("Forma de pago: 30 días factura", est["Normal"]),
        Paragraph("Plazo de entrega: 3 días hábiles", est["Normal"]),
        Paragraph("Validez de la oferta: 15 días", est["Normal"]),
        Paragraph("Vendedor: Juan Pérez", est["Normal"]),
    ]
    doc.build(elementos)


def sin_bordes(ruta: Path) -> None:
    """Cotización escrita como texto en columnas, sin líneas de tabla."""
    c = canvas.Canvas(str(ruta), pagesize=letter)
    y = 750

    def l(txt, x=40, size=10):
        nonlocal y
        c.setFont("Helvetica", size)
        c.drawString(x, y, txt)
        y -= 15

    l("SODIMAC CONSTRUCTOR S.A.", size=14)
    l("R.U.T. 96.792.430-K")
    l("Presupuesto Nro: P-2026-0912")
    l("Santiago, 2 de noviembre de 2026")
    l("Cliente: Constructora Ejemplo Ltda.")
    y -= 10
    l("Cod.       Descripción                                Cant.  U.    P. Unit.      Total")
    filas = [
        ("CEM-25", "Cemento especial saco 25 kg", 120, "UN", 5290),
        ("FIE-12", "Fierro estriado A630 12 mm x 6 m", 80, "UN", 7990),
        ("YES-15", "Plancha yeso cartón 15 mm ST", 60, "UN", 9490),
    ]
    for cod, d, q, u, p in filas:
        c.setFont("Helvetica", 10)
        c.drawString(40, y, cod)
        c.drawString(100, y, d)
        c.drawRightString(345, y, str(q))
        c.drawString(355, y, u)
        c.drawRightString(450, y, miles(p))
        c.drawRightString(540, y, miles(q * p))
        y -= 15
    neto = sum(q * p for *_, q, _, p in filas)
    y -= 10
    l(f"Neto                                                                   {miles(neto)}", x=300)
    l(f"IVA                                                                     {miles(round(neto * .19))}", x=300)
    l(f"Total                                                                  {miles(neto + round(neto * .19))}", x=300)
    l("Condiciones de pago: contado")
    l("Válida hasta: 30/11/2026")
    l("Ejecutivo: María Soto  - maria.soto@sodimac.cl")
    c.save()


def en_dolares(ruta: Path) -> None:
    est = getSampleStyleSheet()
    doc = SimpleDocTemplate(str(ruta), pagesize=letter)
    datos = [["Item", "Part Number", "Description", "Qty", "Unit Price", "Amount"],
             ["1", "PMP-3HP", "Bomba centrífuga 3 HP 380V", "2", "US$ 1,250.00", "US$ 2,500.00"],
             ["2", "VAL-2IN", "Válvula de bola acero inox 2\"", "10", "US$ 85.50", "US$ 855.00"]]
    t = Table(datos)
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black)]))
    doc.build([
        Paragraph("<b>Importadora Hidráulica Ltda.</b>", est["Title"]),
        Paragraph("RUT 78.555.444-3", est["Normal"]),
        Paragraph("Quotation # Q-7781", est["Normal"]),
        Paragraph("Fecha: 2026-09-20", est["Normal"]),
        Paragraph("Moneda: USD", est["Normal"]),
        t,
        Paragraph("Subtotal: US$ 3,355.00", est["Normal"]),
        Paragraph("IVA 19%: US$ 637.45", est["Normal"]),
        Paragraph("Total: US$ 3,992.45", est["Normal"]),
        Paragraph("Plazo de entrega: 6 semanas", est["Normal"]),
        Paragraph("Validez: 30 días", est["Normal"]),
    ])


if __name__ == "__main__":
    destino = Path(sys.argv[1] if len(sys.argv) > 1 else "ejemplos")
    (destino / "2026" / "Octubre").mkdir(parents=True, exist_ok=True)
    con_tabla(destino / "2026" / "Octubre" / "Cotizacion Ferreteria Los Andes 4581.pdf")
    sin_bordes(destino / "Presupuesto Sodimac.pdf")
    en_dolares(destino / "Quotation Q-7781.pdf")
    print(f"Ejemplos creados en {destino.resolve()}")
