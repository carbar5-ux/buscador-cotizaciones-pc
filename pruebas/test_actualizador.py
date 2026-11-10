import io
import zipfile

import actualizador as act


def test_clave_version():
    assert act._clave_version("1.10.0") > act._clave_version("1.9.7")
    assert act._clave_version("2") > act._clave_version("1.9")


def test_instalar_ignora_rutas_peligrosas(tmp_path, monkeypatch):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("repo-abc/nuevo.py", "x = 1")
        z.writestr("repo-abc/../../fuera.py", "x = 2")
        z.writestr("repo-abc/config.toml", "malo")
    monkeypatch.setattr(act, "BASE", tmp_path.resolve())
    monkeypatch.setattr(act, "_pedir", lambda *a, **k: buf.getvalue())
    assert act.instalar("") is True
    assert (tmp_path / "nuevo.py").exists()
    assert not (tmp_path / "config.toml").exists()
    assert not (tmp_path.parent / "fuera.py").exists()


def test_consultar_estados(monkeypatch):
    import urllib.error

    import actualizador as a

    monkeypatch.setattr(a, "version_local", lambda: "1.10.1")
    monkeypatch.setattr(a, "version_remota", lambda t: "1.10.2")
    assert a.consultar("") == ("nueva", "1.10.2")
    monkeypatch.setattr(a, "version_remota", lambda t: "1.10.1")
    assert a.consultar("") == ("al_dia", "1.10.1")

    def sin_red(t):
        raise urllib.error.URLError("sin red")

    monkeypatch.setattr(a, "version_remota", sin_red)
    estado, motivo = a.consultar("")
    assert estado == "error" and "conexión" in motivo
    assert a.buscar_actualizacion("") is None
