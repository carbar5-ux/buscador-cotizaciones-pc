"""Extracción de datos de cotizaciones en PDF.

Lee el texto y las tablas de cada PDF (pdfplumber) y, con reglas pensadas para
cotizaciones chilenas y latinoamericanas, obtiene:

* Encabezado: proveedor, RUT, cliente, N° de cotización, fecha, validez,
  moneda, forma de pago, plazo y lugar de entrega, contacto, correo y teléfono.
* Totales: neto/subtotal, descuento, IVA y total.
* Líneas: ítem, código, descripción, cantidad, unidad, precio unitario,
  descuento y total de cada línea.

Si el PDF es escaneado (sin texto) intenta OCR con Tesseract, si está instalado.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import asdict, dataclass, field
from datetime import date, timedelta
from pathlib import Path

import logging
import warnings

import pdfplumber

# Las librerías de PDF avisan de detalles internos sin importancia (p. ej. "Cannot set
# non-stroke color..."). No afectan la lectura y llenan la pantalla: se silencian.
for _nombre in ("pdfminer", "pdfplumber", "pypdfium2", "PIL"):
    logging.getLogger(_nombre).setLevel(logging.ERROR)
warnings.filterwarnings("ignore", module=r"(pdfminer|pdfplumber|pypdfium2)")

# --------------------------------------------------------------------------
# Modelos
# --------------------------------------------------------------------------


@dataclass
class Item:
    item: str = ""
    codigo: str = ""
    descripcion: str = ""
    cantidad: float | None = None
    unidad: str = ""
    precio_unitario: float | None = None
    descuento: float | None = None
    total: float | None = None
    pagina: int = 0
    control: str = ""
    categoria: str = ""


@dataclass
class Cotizacion:
    archivo: str = ""
    proveedor: str = ""
    rut_proveedor: str = ""
    cliente: str = ""
    numero: str = ""
    fecha: str = ""  # ISO yyyy-mm-dd
    validez: str = ""
    vence: str = ""  # ISO yyyy-mm-dd
    moneda: str = ""
    neto: float | None = None
    descuento: float | None = None
    iva: float | None = None
    total: float | None = None
    forma_pago: str = ""
    plazo_entrega: str = ""
    lugar_entrega: str = ""
    contacto: str = ""
    email: str = ""
    telefono: str = ""
    paginas: int = 0
    metodo: str = ""
    observaciones: list[str] = field(default_factory=list)
    categorias: list[str] = field(default_factory=list)
    items: list[Item] = field(default_factory=list)

    def to_dict(self) -> dict:
        return asdict(self)

    @staticmethod
    def from_dict(d: dict) -> "Cotizacion":
        d = dict(d)
        items = [Item(**{k: v for k, v in i.items() if k in Item.__dataclass_fields__}) for i in d.pop("items", [])]
        c = Cotizacion(**{k: v for k, v in d.items() if k in Cotizacion.__dataclass_fields__})
        c.items = items
        return c


# --------------------------------------------------------------------------
# Utilidades de texto y números
# --------------------------------------------------------------------------


def sin_tildes(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def normalizar(s: str) -> str:
    """Minúsculas, sin tildes ni signos, espacios simples (para comparar)."""
    s = sin_tildes(str(s or "")).lower()
    s = re.sub(r"[^a-z0-9%/]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def limpiar(s) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip()


_NUM_RE = re.compile(r"^[-(]?\d[\d.,]*\)?$")


def parse_num(texto, formato: str = "chile") -> float | None:
    """Convierte '1.234.567', '1.234,5', '$ 12.000', 'US$ 1,234.50' a número.

    formato='chile': punto = miles, coma = decimal (salvo que el texto use
    claramente el formato 1,234.56). formato='us': coma = miles, punto = decimal.
    """
    if texto is None:
        return None
    if isinstance(texto, (int, float)):
        return float(texto)
    s = str(texto).strip()
    s = re.sub(r"(?i)(us\$|usd|clp|uf|\$|eur|€|ch\$)", "", s).replace(" ", "").replace(" ", "")
    s = s.rstrip("%")
    if not s or not _NUM_RE.match(s):
        return None
    negativo = s.startswith("-") or (s.startswith("(") and s.endswith(")"))
    s = s.strip("-()")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):  # 1.234,56
            s = s.replace(".", "").replace(",", ".")
        else:  # 1,234.56
            s = s.replace(",", "")
    elif "," in s:
        if formato == "us" and re.fullmatch(r"\d{1,3}(,\d{3})+", s):
            s = s.replace(",", "")
        elif re.fullmatch(r"\d{1,3}(,\d{3}){2,}", s):  # 1,234,567 (miles inequívoco)
            s = s.replace(",", "")
        else:
            s = s.replace(",", ".")
    elif "." in s:
        if formato == "us":
            if re.fullmatch(r"\d{1,3}(\.\d{3}){2,}", s):
                s = s.replace(".", "")
        elif re.fullmatch(r"\d{1,3}(\.\d{3})+", s):  # 12.000 / 1.234.567
            s = s.replace(".", "")
        elif s.count(".") > 1:
            s = s.replace(".", "")
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if negativo else v


def es_numero(tok: str, formato: str = "chile") -> bool:
    return parse_num(tok, formato) is not None


# --------------------------------------------------------------------------
# Fechas
# --------------------------------------------------------------------------

MESES = {
    "ene": 1, "enero": 1, "jan": 1, "january": 1,
    "feb": 2, "febrero": 2, "february": 2,
    "mar": 3, "marzo": 3, "march": 3,
    "abr": 4, "abril": 4, "apr": 4, "april": 4,
    "may": 5, "mayo": 5,
    "jun": 6, "junio": 6, "june": 6,
    "jul": 7, "julio": 7, "july": 7,
    "ago": 8, "agosto": 8, "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "septiembre": 9, "setiembre": 9, "september": 9,
    "oct": 10, "octubre": 10, "october": 10,
    "nov": 11, "noviembre": 11, "november": 11,
    "dic": 12, "diciembre": 12, "dec": 12, "december": 12,
}
_MES_ALT = "|".join(sorted(MESES, key=len, reverse=True))

_FECHA_PATRONES = [
    # 06/11/2026, 6-11-26, 06.11.2026
    (re.compile(r"\b(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4}|\d{2})\b"), "dmy"),
    # 2026-11-06
    (re.compile(r"\b(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})\b"), "ymd"),
    # 6 de noviembre de 2026 / 6 noviembre 2026 / 06-nov-2026
    (re.compile(rf"\b(\d{{1,2}})(?:\s+de\s+|\s+|-|/)({_MES_ALT})\.?(?:\s+de(?:l)?\s+|\s+|-|/|,\s*)(\d{{4}})\b", re.I), "dMy"),
    # noviembre 6, 2026
    (re.compile(rf"\b({_MES_ALT})\.?\s+(\d{{1,2}}),?\s+(\d{{4}})\b", re.I), "Mdy"),
]


def _armar_fecha(a: str, b: str, c: str, orden: str) -> date | None:
    try:
        if orden == "dmy":
            d, m, y = int(a), int(b), int(c)
        elif orden == "ymd":
            y, m, d = int(a), int(b), int(c)
        elif orden == "dMy":
            d, m, y = int(a), MESES[sin_tildes(b).lower()], int(c)
        else:  # Mdy
            m, d, y = MESES[sin_tildes(a).lower()], int(b), int(c)
        if y < 100:
            y += 2000
        if not 1990 <= y <= 2100:
            return None
        return date(y, m, d)
    except (ValueError, KeyError):
        return None


def buscar_fechas(texto: str) -> list[tuple[int, date]]:
    """Todas las fechas del texto con su posición."""
    encontradas: list[tuple[int, date]] = []
    for patron, orden in _FECHA_PATRONES:
        for m in patron.finditer(texto):
            f = _armar_fecha(*m.groups(), orden)
            if f:
                encontradas.append((m.start(), f))
    return sorted(encontradas)


# --------------------------------------------------------------------------
# Encabezado
# --------------------------------------------------------------------------

RUT_RE = re.compile(r"\b(\d{1,2}\.?\d{3}\.?\d{3}\s?-\s?[\dkK])\b")
EMAIL_RE = re.compile(r"[\w.+\-]+@[\w\-]+(?:\.[\w\-]+)+")
SOCIEDAD_RE = re.compile(r"\b(s\.?\s?a\.?|spa|ltda\.?|limitada|e\.?i\.?r\.?l\.?|s\.?p\.?a\.?|inc\.?|llc|sac|s\.?a\.?c\.?|s\.? de r\.?l\.?)(?=\s|$|,)", re.I)


def _valor_etiqueta(lineas: list[str], etiquetas: str, max_len: int = 120) -> str:
    """Valor que sigue a una etiqueta tipo 'Forma de pago: 30 días'.

    La etiqueta debe ir seguida de ':' o estar al inicio de la línea, para no
    confundir textos como 'Despacho a obra' (un ítem) con un plazo de despacho.
    """
    patron = re.compile(rf"(?i)(?:^|\b)(?:{etiquetas})\b\s*(?::\s*(.*)|[.\-]?\s+(.*)$|$)")
    for i, linea in enumerate(lineas):
        m = patron.search(linea)
        if not m or (m.group(1) is None and m.start() > 0):
            continue
        valor = (m.group(1) if m.group(1) is not None else m.group(2) or "").strip(" :.-\t")
        if m.group(1) is None and re.search(r"\d[\d.,]{3,}\s*$", valor):
            continue  # línea de detalle con montos, no una condición comercial
        # si en la misma línea hay otra etiqueta "xxx:", cortar ahí
        valor = re.split(r"\s{2,}|\s+\b[A-ZÁÉÍÓÚa-záéíóúñÑ°º. ]{3,25}:\s", valor)[0].strip()
        if not valor and i + 1 < len(lineas):
            valor = lineas[i + 1].strip()
        if valor:
            return valor[:max_len]
    return ""


def _monto_etiqueta(lineas: list[str], etiquetas: str, formato: str, ultima: bool = True) -> float | None:
    """Número que acompaña a una etiqueta de totales ('IVA 19%  $ 19.000')."""
    patron = re.compile(rf"(?i)^\s*(?:{etiquetas})\b(.*)$")
    candidatos: list[float] = []
    for linea in lineas:
        m = patron.search(linea)
        if not m:
            continue
        resto = re.sub(r"\d+(?:[.,]\d+)?\s*%", " ", m.group(1))  # quita '19%'
        nums = [parse_num(t, formato) for t in re.findall(r"-?\(?[\d][\d.,]*\)?", resto)]
        nums = [n for n in nums if n is not None]
        if nums:
            candidatos.append(nums[-1])
    if not candidatos:
        return None
    return candidatos[-1] if ultima else candidatos[0]


def _normaliza_rut(r: str) -> str:
    r = re.sub(r"[\s.]", "", r).upper()
    cuerpo, dv = r.split("-")
    return f"{int(cuerpo):,}".replace(",", ".") + "-" + dv


def extraer_encabezado(cot: Cotizacion, texto: str, mis_ruts: set[str], formato: str, dias_validez: int) -> None:
    lineas = [limpiar(l) for l in texto.splitlines() if limpiar(l)]
    bajo = sin_tildes(texto).lower()

    # --- RUT del proveedor (el primero que no sea mío) ---
    ruts = [_normaliza_rut(m.group(1)) for m in RUT_RE.finditer(texto)]
    mios = {_normaliza_rut(r) for r in mis_ruts if RUT_RE.search(r)}
    pos_cliente = re.search(r"(?i)\b(se[ñn]or(?:es)?|sres\.?|cliente|raz[oó]n social cliente|atenci[oó]n)\b", texto)
    rut_prov = ""
    for m in RUT_RE.finditer(texto):
        r = _normaliza_rut(m.group(1))
        if r in mios:
            continue
        # el RUT inmediatamente tras "Señores/Cliente" suele ser del cliente
        if pos_cliente and 0 <= m.start() - pos_cliente.start() < 250 and len(ruts) > 1:
            continue
        rut_prov = r
        break
    if not rut_prov:
        rut_prov = next((r for r in ruts if r not in mios), "")
    cot.rut_proveedor = rut_prov

    # --- Proveedor ---
    prov = _valor_etiqueta(lineas, r"raz[oó]n social|proveedor|empresa|emisor")
    if prov and RUT_RE.search(prov):
        prov = RUT_RE.split(prov)[0].strip(" -:,")
    if not prov or len(prov) < 3:
        lim = pos_cliente.start() if pos_cliente else len(texto)
        cabecera = [l for l in texto[:lim].splitlines()[:20] if limpiar(l)]
        con_soc = [limpiar(l) for l in cabecera if SOCIEDAD_RE.search(l)]
        if con_soc:
            prov = con_soc[0]
        elif rut_prov:
            for i, l in enumerate(lineas):
                if rut_prov.replace(".", "") in l.replace(".", ""):
                    previo = RUT_RE.sub("", l).replace("RUT", "").replace("R.U.T.", "").strip(" :-,")
                    prov = previo if len(previo) > 3 else (lineas[i - 1] if i else "")
                    break
        if not prov:
            prov = next(
                (l for l in lineas[:6] if not re.search(r"(?i)cotizaci|presupuesto|oferta|fecha|p[aá]gina|quot", l)),
                "",
            )
    prov = RUT_RE.sub("", prov)
    prov = re.sub(r"(?i)\b(r\.?u\.?t\.?)\s*[:.]?\s*$", "", prov).strip(" :-,")
    cot.proveedor = prov[:100]

    # --- Cliente ---
    cot.cliente = _valor_etiqueta(lineas, r"se[ñn]or(?:es)?|sres\.?|cliente|raz[oó]n social cliente", 100)
    cot.cliente = re.sub(r"(?i)\s*\b(r\.?u\.?t\.?)\s*[:.]?\s*$", "", RUT_RE.sub("", cot.cliente)).strip(" :-,")

    # --- N° de cotización ---
    m = re.search(
        r"(?i)(?:cotizaci[oó]n|presupuesto|oferta|propuesta|quotation|quote|proforma)\s*"
        r"(?:n[°ºo]?\.?|nro\.?|n[uú]m(?:ero)?\.?|#|folio)?\s*[:\-]?\s*([A-Z]{0,5}[\-/]?\d[\w\-/.]*)",
        texto,
    )
    if not m:
        m = re.search(r"(?i)\b(?:n[°º]|nro\.?|folio|n[uú]mero)\s*[:\-]?\s*([A-Z]{0,5}[\-/]?\d[\w\-/]*)", texto)
    if m:
        num = m.group(1).strip(".-/")
        if not RUT_RE.fullmatch(num) and not buscar_fechas(num):
            cot.numero = num

    # --- Fecha (preferir la que sigue a la etiqueta "Fecha") ---
    fechas = buscar_fechas(texto)
    f_doc: date | None = None
    for etiqueta in (r"fecha\s+(?:de\s+)?(?:emisi[oó]n|cotizaci[oó]n)", r"\bfecha\b", r"\bdate\b", r"emitid[ao]"):
        me = re.search(rf"(?i){etiqueta}", texto)
        if me:
            prox = [f for p, f in fechas if 0 <= p - me.end() < 80]
            if prox:
                f_doc = prox[0]
                break
    if not f_doc and fechas:
        # ciudad, 6 de noviembre de 2026  /  o la primera fecha del documento
        f_doc = fechas[0][1]
    if f_doc:
        cot.fecha = f_doc.isoformat()

    # --- Validez / vencimiento ---
    validez = _valor_etiqueta(lineas, r"validez(?:\s+de\s+la\s+oferta)?|v[aá]lid[ao]\s+(?:hasta|por)|vigencia|oferta\s+v[aá]lida|vencimiento", 80)
    cot.validez = validez
    vence: date | None = None
    if validez:
        fv = buscar_fechas(validez)
        if fv:
            vence = fv[0][1]
        else:
            md = re.search(r"(\d+)\s*(d[ií]as?|days?|semanas?|mes(?:es)?)", sin_tildes(validez).lower())
            if md and f_doc:
                n = int(md.group(1))
                unidad = md.group(2)
                dias = n * 7 if unidad.startswith("seman") else n * 30 if unidad.startswith("mes") else n
                vence = f_doc + timedelta(days=dias)
    if not vence and f_doc and dias_validez:
        vence = f_doc + timedelta(days=dias_validez)
        if not cot.validez:
            cot.validez = f"(supuesto {dias_validez} días)"
    if vence:
        cot.vence = vence.isoformat()

    # --- Moneda ---
    if re.search(r"\b(usd|us\$|d[oó]lar(?:es)?)\b", bajo, re.I) or "us$" in bajo:
        cot.moneda = "USD"
    elif re.search(r"\b(eur|euros?)\b|€", bajo):
        cot.moneda = "EUR"
    elif re.search(r"\b(uf|u\.f\.|unidades? de fomento)\b", bajo) and not re.search(r"\bclp\b|\$\s?\d", bajo):
        cot.moneda = "UF"
    else:
        cot.moneda = "CLP"

    # --- Condiciones comerciales y contacto ---
    cot.forma_pago = _valor_etiqueta(lineas, r"forma\s+de\s+pago|condici[oó]n(?:es)?\s+de\s+pago|condiciones?\s+comerciales|pago", 80)
    cot.plazo_entrega = _valor_etiqueta(lineas, r"plazo\s+de\s+entrega|tiempo\s+de\s+entrega|plazo\s+de\s+despacho|plazo|disponibilidad|entrega|delivery", 80)
    cot.lugar_entrega = _valor_etiqueta(lineas, r"lugar\s+de\s+entrega|direcci[oó]n\s+de\s+(?:entrega|despacho)|entrega\s+en\s+obra|puesto\s+en", 80)
    cot.contacto = _valor_etiqueta(lineas, r"vendedor|ejecutiv[oa](?:\s+comercial)?|contacto|atentamente|cotizado\s+por|emitido\s+por", 80)
    cot.contacto = EMAIL_RE.sub("", cot.contacto).strip(" ,-")
    emails = EMAIL_RE.findall(texto)
    cot.email = emails[0] if emails else ""
    mf = re.search(r"(?i)(?:tel[eé]fono|fono|tel\.?|cel(?:ular)?\.?|m[oó]vil|whatsapp)\s*[:.]?\s*(\+?[\d\s()\-]{8,20}\d)", texto)
    mt = mf or re.search(r"(\+56[\s\-]?(?:\(?\d{1,2}\)?[\s\-]?)?\d{4}[\s\-]?\d{4})", texto)
    cot.telefono = limpiar(mt.group(1)) if mt else ""

    # --- Totales ---
    cot.neto = _monto_etiqueta(lineas, r"(?:sub\s*-?\s*total|total\s+neto|neto|valor\s+neto|monto\s+neto|afecto)", formato)
    cot.descuento = _monto_etiqueta(lineas, r"(?:descuento|dcto\.?|desc\.)(?:\s+(?:global|total|general))?", formato)
    cot.iva = _monto_etiqueta(lineas, r"(?:i\.?v\.?a\.?|impuesto|19\s*%)", formato)
    cot.total = _monto_etiqueta(
        lineas, r"(?:total\s+(?:general|a\s+pagar|cotizaci[oó]n|final|con\s+iva|bruto|oferta)|total)(?!\s+neto)", formato
    )
    if cot.total is not None and cot.neto is not None and cot.total < cot.neto:
        cot.total = None  # tomó un total de línea, no el general
    if cot.neto is None and cot.total is not None and cot.iva:
        cot.neto = round(cot.total - cot.iva, 2)


# --------------------------------------------------------------------------
# Líneas de detalle: tablas
# --------------------------------------------------------------------------

# Orden importa: se revisan de arriba hacia abajo y gana la primera coincidencia.
COLUMNAS: list[tuple[str, re.Pattern]] = [
    ("descuento", re.compile(r"^(desc(uento)?|dcto|dscto|% ?desc|desc ?%|descto)\b")),
    ("precio_unitario", re.compile(r"(unitari|p ?unit|\bp/?u\b|\bpu\b|precio|valor u|v ?unit|unit price|price|tarifa)")),
    ("total", re.compile(r"(total|importe|monto|subtotal|valor neto|neto|amount|valor)")),
    ("cantidad", re.compile(r"(^cant|cantidad|\bqty\b|quantity|\bcant\b|\bunidades\b|\bnro unid)")),
    ("unidad", re.compile(r"(^u$|^un$|^und$|^uds?$|unidad|\bu/m\b|^um$|medida|^unit$|^ud$)")),
    ("codigo", re.compile(r"(codigo|^cod\b|\bsku\b|part|referencia|^ref\b|n ?parte|modelo)")),
    ("item", re.compile(r"^(item|it|n|no|nro|pos|posicion|linea|#|n°|nº)$")),
    ("descripcion", re.compile(r"(descrip|detalle|producto|glosa|concepto|material|servicio|articulo|description|nombre|partida)")),
]

PALABRAS_TOTALES = re.compile(r"(?i)^\s*(sub\s*-?total|total|neto|i\.?v\.?a|impuesto|descuento|dcto|afecto|exento|flete)\b")


def mapear_encabezado(fila: list) -> dict[str, int] | None:
    mapa: dict[str, int] = {}
    for idx, celda in enumerate(fila):
        n = normalizar(celda).replace("°", "").strip()
        if not n:
            continue
        for campo, patron in COLUMNAS:
            if campo in mapa:
                continue
            if patron.search(n):
                mapa[campo] = idx
                break
    # Un encabezado válido tiene descripción y al menos un dato numérico.
    if "descripcion" in mapa and ({"precio_unitario", "total", "cantidad"} & mapa.keys()):
        return mapa
    return None


def _fila_a_item(fila: list, mapa: dict[str, int], formato: str, pagina: int) -> Item | None:
    def celda(campo: str) -> str:
        i = mapa.get(campo)
        return limpiar(fila[i]) if i is not None and i < len(fila) else ""

    desc = celda("descripcion")
    if PALABRAS_TOTALES.match(desc) or (not desc and PALABRAS_TOTALES.match(" ".join(limpiar(c) for c in fila))):
        return None
    cant = parse_num(celda("cantidad"), formato)
    pu = parse_num(celda("precio_unitario"), formato)
    tot = parse_num(celda("total"), formato)
    dcto = parse_num(celda("descuento"), formato)
    unidad = celda("unidad")
    # cantidad con unidad pegada: "10 m2", "5 un"
    if cant is None and celda("cantidad"):
        mc = re.match(r"^([\d.,]+)\s*([a-zA-Z²³0-9/]{1,6})$", celda("cantidad"))
        if mc:
            cant = parse_num(mc.group(1), formato)
            unidad = unidad or mc.group(2)
    if not desc and cant is None and pu is None and tot is None:
        return None
    return Item(
        item=celda("item"), codigo=celda("codigo"), descripcion=desc, cantidad=cant, unidad=unidad,
        precio_unitario=pu, descuento=dcto, total=tot, pagina=pagina,
    )


def items_desde_tablas(pdf, formato: str, paginas: set[int] | None = None) -> list[Item]:
    """Ítems de las tablas con bordes. `paginas` (1, 2, ...) limita qué páginas revisar."""
    items: list[Item] = []
    mapa: dict[str, int] | None = None
    ncols = 0
    for npag, pagina in enumerate(pdf.pages, start=1):
        if paginas is not None and npag not in paginas:
            continue
        if not (pagina.lines or pagina.rects):  # sin líneas no hay tabla con bordes: ahorra tiempo
            continue
        try:
            tablas = pagina.extract_tables()
        except Exception:
            continue
        for tabla in tablas:
            if not tabla:
                continue
            inicio = None
            for i, fila in enumerate(tabla[:6]):
                m = mapear_encabezado(fila)
                if m:
                    mapa, ncols, inicio = m, len(fila), i + 1
                    break
            if inicio is None:
                # tabla que continúa de la página anterior (mismo n° de columnas)
                if mapa and len(tabla[0]) == ncols:
                    inicio = 0
                else:
                    continue
            for fila in tabla[inicio:]:
                fila = [c if c is not None else "" for c in fila]
                # una celda puede traer varias líneas (varios ítems apilados)
                partes = [str(c).split("\n") for c in fila]
                nlin = max(len(p) for p in partes)
                desc_i = mapa.get("descripcion", -1)
                if nlin > 1 and all(len(p) in (1, nlin) for p in partes) and sum(len(p) == nlin for p in partes) >= 3:
                    filas = [[p[j] if len(p) == nlin else (p[0] if j == 0 else "") for p in partes] for j in range(nlin)]
                else:
                    filas = [[" ".join(p) if k == desc_i else " ".join(p) for k, p in enumerate(partes)]]
                for f in filas:
                    it = _fila_a_item(f, mapa, formato, npag)
                    if not it:
                        continue
                    sin_numeros = it.cantidad is None and it.precio_unitario is None and it.total is None
                    if sin_numeros and items and it.descripcion and not it.codigo:
                        items[-1].descripcion = limpiar(items[-1].descripcion + " " + it.descripcion)
                        continue
                    if sin_numeros:
                        continue
                    items.append(it)
    return items


# --------------------------------------------------------------------------
# Líneas de detalle: texto (cuando el PDF no trae tablas con bordes)
# --------------------------------------------------------------------------

UNIDAD_RE = re.compile(
    r"(?i)^(un|und|uds?|unid|unidad(?:es)?|c/u|cu|pza|pzas?|kg|kgs|gr|g|ton|tn|lt|lts|l|m|mt|mts|ml|m2|m²|m3|m³|"
    r"cm|mm|gl|gal|glb|global|hr|hrs|h|hh|dia|días?|mes|sem|caja|cj|rollo|rl|par|jgo|juego|set|kit|bolsa|saco|tambor|"
    r"tira|barra|plancha|pl|vara|pie|ft|pulg|lote|ea|pc|pcs|bid[oó]n|galon|tineta|metros?|mtrs?|kilos?|litros?|"
    r"piezas?|rollos?|cajas?|sacos?|bolsas?|pares|juegos|unidades|und\.|uni)\.?$"
)


def _cuadra(cant: float | None, pu: float | None, tot: float | None, dcto: float | None = None) -> bool:
    if cant is None or pu is None or tot is None or tot == 0:
        return False
    bruto = cant * pu
    if abs(bruto - tot) <= max(1.0, abs(tot) * 0.01):
        return True
    if dcto is not None:
        for neto in (bruto * (1 - dcto / 100), bruto - dcto):
            if abs(neto - tot) <= max(1.0, abs(tot) * 0.01):
                return True
    return False


def _linea_a_item(linea: str, formato: str, pagina: int, en_detalle: bool) -> Item | None:
    toks = linea.split()
    if len(toks) < 3 or PALABRAS_TOTALES.match(linea):
        return None
    # unir "$ 12.000" -> "$12.000"
    t2: list[str] = []
    for t in toks:
        if t2 and t2[-1] in ("$", "US$", "USD", "CLP", "UF"):
            t2[-1] = t2[-1] + t
        else:
            t2.append(t)
    toks = t2

    nums: list[tuple[int, float]] = []
    unidad = ""
    i = len(toks) - 1
    pct = None
    while i >= 0:
        t = toks[i]
        if t.endswith("%") and parse_num(t[:-1], formato) is not None and pct is None and len(nums) <= 2:
            pct = parse_num(t[:-1], formato)
        elif es_numero(t, formato):
            nums.append((i, parse_num(t, formato)))
        elif UNIDAD_RE.match(t) and not unidad and len(nums) >= 1:
            unidad = t
        else:
            break
        i -= 1
        if len(nums) >= 4:
            break
    if len(nums) < 2:
        return None
    nums.reverse()
    resto = toks[: i + 1]
    # unidad antes de la cantidad: "... UN 10 5.000 50.000"
    if not unidad and resto and UNIDAD_RE.match(resto[-1]) and len(resto) > 1:
        unidad = resto.pop()
    valores = [v for _, v in nums]
    cant = pu = tot = None
    dcto = pct
    if len(valores) >= 4:
        # cant, pu, dcto, total   o   item?, cant, pu, total
        if _cuadra(valores[-4], valores[-3], valores[-1], valores[-2]):
            cant, pu, dcto, tot = valores[-4], valores[-3], valores[-2], valores[-1]
        else:
            cant, pu, tot = valores[-3], valores[-2], valores[-1]
            if not resto:
                resto = [str(int(valores[0]))]
    elif len(valores) == 3:
        cant, pu, tot = valores
    else:  # 2 números: cantidad + total, o precio + total
        if not en_detalle:
            return None
        a, b = valores
        cant, pu, tot = (a, b / a, b) if a and a < b and a == int(a) and a < 10000 else (1.0, a, b)
    ok = _cuadra(cant, pu, tot, dcto)
    if not ok and not en_detalle:
        return None
    # ítem y código al inicio
    item = codigo = ""
    if resto and re.fullmatch(r"\d{1,3}(\.\d{1,2})?", resto[0]) and len(resto) > 1:
        item = resto.pop(0)
    if resto and len(resto) > 1 and re.fullmatch(r"(?=.*\d)[A-Z0-9][A-Z0-9\-./]{2,}", resto[0]):
        codigo = resto.pop(0)
    desc = " ".join(resto).strip(" -:|")
    if not desc or not re.search(r"[A-Za-zÁÉÍÓÚáéíóúñÑ]{2,}", desc):
        return None
    return Item(item=item, codigo=codigo, descripcion=desc, cantidad=cant, unidad=unidad,
                precio_unitario=pu, descuento=dcto, total=tot, pagina=pagina)


def items_desde_texto(paginas_texto: list[str], formato: str) -> list[Item]:
    items: list[Item] = []
    for npag, texto in enumerate(paginas_texto, start=1):
        en_detalle = False
        for linea in texto.splitlines():
            linea = limpiar(linea)
            if not linea:
                continue
            n = normalizar(linea)
            if not en_detalle and re.search(r"(descrip|detalle|producto|glosa|concepto)", n) and re.search(
                r"(cant|precio|valor|total|unit|qty)", n
            ):
                en_detalle = True
                continue
            if en_detalle and PALABRAS_TOTALES.match(linea):
                en_detalle = False
                continue
            it = _linea_a_item(linea, formato, npag, en_detalle)
            if it:
                items.append(it)
            elif en_detalle and items and items[-1].pagina == npag and re.search(r"[A-Za-z]{3,}", linea) \
                    and not re.search(r"\d[\d.,]{3,}", linea) and len(linea) < 120:
                items[-1].descripcion = limpiar(items[-1].descripcion + " " + linea)
    return items


# --------------------------------------------------------------------------
# OCR opcional
# --------------------------------------------------------------------------


def ocr_pdf(ruta: Path, max_paginas: int = 0) -> list[str] | None:
    """Texto por página vía Tesseract. Devuelve None si no hay OCR disponible."""
    try:
        import pypdfium2 as pdfium
        import pytesseract
    except ImportError:
        return None
    try:
        doc = pdfium.PdfDocument(str(ruta))
        textos = []
        for n, pagina in enumerate(doc, start=1):
            if max_paginas and n > max_paginas:
                break
            imagen = pagina.render(scale=300 / 72).to_pil()
            try:
                textos.append(pytesseract.image_to_string(imagen, lang="spa", config="--psm 6"))
            except pytesseract.TesseractError:
                textos.append(pytesseract.image_to_string(imagen, config="--psm 6"))
        return textos
    except Exception:
        return None


# --------------------------------------------------------------------------
# Controles
# --------------------------------------------------------------------------


def controlar(cot: Cotizacion) -> None:
    for it in cot.items:
        if it.total is None and it.cantidad is not None and it.precio_unitario is not None:
            bruto = it.cantidad * it.precio_unitario
            if it.descuento:
                bruto = bruto * (1 - it.descuento / 100) if it.descuento <= 100 else bruto - it.descuento
            it.total = round(bruto, 4)
            it.control = "Total calculado"
        elif it.precio_unitario is None and it.cantidad and it.total is not None:
            it.precio_unitario = round(it.total / it.cantidad, 4)
            it.control = "P. unitario calculado"
        elif it.cantidad is None and it.precio_unitario and it.total is not None:
            c = it.total / it.precio_unitario
            it.cantidad = round(c, 4)
            it.control = "Cantidad calculada"
        elif _cuadra(it.cantidad, it.precio_unitario, it.total, it.descuento):
            it.control = "OK"
        else:
            it.control = "Revisar: cantidad × precio ≠ total"

    suma = sum(i.total for i in cot.items if i.total is not None)
    if not cot.items:
        cot.observaciones.append("No se detectaron líneas de detalle")
    if cot.neto is None and cot.items:
        cot.neto = round(suma, 2)
        cot.observaciones.append("Neto calculado como suma de ítems")
    if cot.neto is not None and cot.items:
        base = cot.neto + (cot.descuento or 0)
        if abs(suma - cot.neto) > max(2, abs(cot.neto) * 0.005) and abs(suma - base) > max(2, abs(base) * 0.005):
            cot.observaciones.append(f"Suma de ítems ({suma:,.0f}) distinta del neto ({cot.neto:,.0f})")
    if cot.total is None and cot.neto is not None:
        if cot.iva is not None:
            cot.total = round(cot.neto + cot.iva, 2)
        cot.observaciones.append("Total no encontrado en el PDF" + (" (calculado neto + IVA)" if cot.iva else ""))
    if not cot.proveedor:
        cot.observaciones.append("Proveedor no identificado")
    if not cot.fecha:
        cot.observaciones.append("Fecha no encontrada")


# --------------------------------------------------------------------------
# Punto de entrada
# --------------------------------------------------------------------------


def leer_rapido(ruta: Path) -> tuple[list[str], set[int]] | None:
    """Texto de cada página con pdfium (muy rápido) y qué páginas tienen líneas dibujadas.

    Devuelve None si pdfium no puede abrir el archivo (se usa entonces pdfplumber).
    """
    try:
        import pypdfium2 as pdfium
        import pypdfium2.raw as pdfium_c
    except ImportError:
        return None
    try:
        doc = pdfium.PdfDocument(str(ruta))
    except Exception:
        return None
    textos: list[str] = []
    con_lineas: set[int] = set()
    try:
        for i in range(len(doc)):
            pagina = doc[i]
            tp = pagina.get_textpage()
            textos.append(tp.get_text_range().replace("\r\n", "\n").replace("\r", "\n"))
            tp.close()
            if any(o.type == pdfium_c.FPDF_PAGEOBJ_PATH for o in pagina.get_objects(max_depth=3)):
                con_lineas.add(i + 1)
            pagina.close()
    except Exception:
        return None
    finally:
        doc.close()
    return textos, con_lineas


def extraer_cotizacion(
    ruta: Path,
    *,
    mis_ruts: set[str] | None = None,
    formato: str = "chile",
    dias_validez: int = 30,
    usar_ocr: bool = True,
    max_paginas_tablas: int = 10,
) -> Cotizacion:
    """Lee una cotización.

    `max_paginas_tablas`: el lector de tablas (lento) solo revisa las primeras N páginas;
    el texto de todas las páginas se lee siempre (rápido). 0 = sin límite. Evita que un
    catálogo o ficha técnica de cientos de páginas tarde minutos.
    """
    cot = Cotizacion(archivo=str(ruta))
    rapido = leer_rapido(ruta)
    if rapido is not None:
        paginas_texto, con_lineas = rapido
    else:  # respaldo: todo con pdfplumber
        with pdfplumber.open(ruta) as pdf:
            paginas_texto = [(p.extract_text(layout=False) or "") for p in pdf.pages]
        con_lineas = set(range(1, len(paginas_texto) + 1))
    cot.paginas = len(paginas_texto)
    texto = "\n".join(paginas_texto)
    if max_paginas_tablas and len(con_lineas) > 0:
        con_lineas = {n for n in con_lineas if n <= max_paginas_tablas}

    if len(texto.strip()) < 40:
        textos_ocr = ocr_pdf(ruta, max_paginas_tablas) if usar_ocr else None
        if not textos_ocr:
            cot.metodo = "sin texto"
            cot.observaciones.append(
                "PDF escaneado (imagen) sin texto: instale Tesseract OCR para leerlo" if usar_ocr
                else "PDF escaneado (imagen) sin texto"
            )
            return cot
        paginas_texto = textos_ocr
        texto = "\n".join(textos_ocr)
        cot.metodo = "OCR"
        cot.items = items_desde_texto(paginas_texto, formato)
    else:
        cot.items, cot.metodo = [], "tablas"
        items_txt = items_desde_texto(paginas_texto, formato)
        # Documento largo (catálogo, ficha técnica) cuyo texto ya trajo las líneas y casi todas cuadran
        # (cantidad × precio = total): se omite el lector de tablas, que es lo más lento. En cotizaciones
        # normales (hasta max_paginas_tablas páginas) siempre se usan las tablas: leen mejor las descripciones.
        cuadran = sum(1 for i in items_txt if _cuadra(i.cantidad, i.precio_unitario, i.total, i.descuento))
        largo = bool(max_paginas_tablas) and cot.paginas > max_paginas_tablas
        texto_basta = largo and len(items_txt) >= 3 and cuadran >= 0.95 * len(items_txt)
        if con_lineas and not texto_basta:  # el lector de tablas (lento) solo en páginas con líneas dibujadas
            with pdfplumber.open(ruta) as pdf:
                cot.items = items_desde_tablas(pdf, formato, con_lineas)
        # usar texto si encontró más líneas consistentes que las tablas
        if len(items_txt) > len(cot.items):
            cot.items, cot.metodo = items_txt, "texto"
    # encabezado y totales: primeras 3 y últimas 2 páginas (en documentos largos no hace falta leer el resto)
    texto_enc = texto if len(paginas_texto) <= 5 else "\n".join(paginas_texto[:3] + paginas_texto[-2:])
    extraer_encabezado(cot, texto_enc, mis_ruts or set(), formato, dias_validez)
    controlar(cot)
    return cot
