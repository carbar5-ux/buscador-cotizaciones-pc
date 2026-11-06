"""Pruebas rápidas:  python -m pytest pruebas   (o  python pruebas/test_extractor.py)"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from extractor import _linea_a_item, buscar_fechas, mapear_encabezado, parse_num  # noqa: E402


def test_numeros_chile():
    assert parse_num("$ 1.234.567") == 1234567
    assert parse_num("12.000") == 12000
    assert parse_num("1.234,5") == 1234.5
    assert parse_num("4,5") == 4.5
    assert parse_num("US$ 1,250.00") == 1250
    assert parse_num("(1.000)") == -1000
    assert parse_num("abc") is None


def test_numeros_us():
    assert parse_num("1,250", "us") == 1250
    assert parse_num("12.5", "us") == 12.5


def test_fechas():
    assert str(buscar_fechas("Fecha: 15/10/2026")[0][1]) == "2026-10-15"
    assert str(buscar_fechas("Santiago, 2 de noviembre de 2026")[0][1]) == "2026-11-02"
    assert str(buscar_fechas("2026-09-20")[0][1]) == "2026-09-20"
    assert str(buscar_fechas("06-nov-2026")[0][1]) == "2026-11-06"


def test_encabezado_tabla():
    m = mapear_encabezado(["Ítem", "Código", "Descripción", "Cant.", "Unidad", "Precio Unitario", "Total"])
    assert m == {"item": 0, "codigo": 1, "descripcion": 2, "cantidad": 3, "unidad": 4, "precio_unitario": 5, "total": 6}
    m = mapear_encabezado(["Item", "Part Number", "Description", "Qty", "Unit Price", "Amount"])
    assert m["codigo"] == 1 and m["cantidad"] == 3 and m["precio_unitario"] == 4 and m["total"] == 5


def test_linea_texto():
    it = _linea_a_item("1 CEM-25 Cemento especial saco 25 kg 120 UN 4.990 598.800", "chile", 1, False)
    assert (it.item, it.codigo, it.cantidad, it.unidad, it.precio_unitario, it.total) == ("1", "CEM-25", 120, "UN", 4990, 598800)
    assert _linea_a_item("Total 1.234.567", "chile", 1, False) is None


if __name__ == "__main__":
    for nombre, f in list(globals().items()):
        if nombre.startswith("test_"):
            f()
            print("OK", nombre)
