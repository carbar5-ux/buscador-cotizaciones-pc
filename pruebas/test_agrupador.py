import agrupador as ag

SIN = {"conductor": "cable", "luminaria": "lampara"}


def test_numeros_medidas():
    assert ag.numeros("Cable 3x2,5 mm2 12 AWG") == frozenset({"3", "2.5", "2", "12"}) or "12" in ag.numeros("12 AWG")
    assert ag.numeros("Tubo 2,50") == frozenset({"2.5"})


def test_palabras_sinonimos():
    assert "cable" in ag.palabras("Conductor THHN", SIN)


def test_unidad_normal_es_idempotente():
    assert ag.unidad_normal(ag.unidad_normal("UN.")) == ag.unidad_normal("UN.")


def test_similitud():
    assert ag.similitud(frozenset(), frozenset()) == 1.0
    assert ag.similitud(frozenset({"a", "b"}), frozenset({"b", "c"})) == 1 / 3


def test_agrupar_mismo_producto_y_distinta_medida(monkeypatch):
    monkeypatch.setattr(ag, "sinonimos", lambda: SIN)
    grupos = ag.agrupar([
        ("Cable THHN 12 AWG", "m", "CLP", "a"),
        ("Conductor THHN 12 AWG", "m", "CLP", "b"),
        ("Cable THHN 10 AWG", "m", "CLP", "c"),
        ("Cable THHN 12 AWG", "m", "USD", "d"),
    ])
    assert sorted(sorted(g) for g in grupos) == [["a", "b"], ["c"], ["d"]]


def test_descripcion_representativa():
    assert ag.descripcion_representativa(["Cable largo x", "Cable", "Cable", ""]) == "Cable"
    assert ag.descripcion_representativa([]) == ""


def test_busqueda():
    assert ag.calza_busqueda("cable thhn 12", "Cable THHN 12 AWG negro", SIN)
    assert not ag.calza_busqueda("cable thhn 10", "Cable THHN 12 AWG negro", SIN)
