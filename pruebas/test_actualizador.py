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
