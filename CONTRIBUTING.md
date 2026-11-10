# Cómo contribuir

```bash
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements-dev.txt
ruff check . && python -m pytest
```

- Código y mensajes en español, como el resto del proyecto.
- Agrega pruebas en `pruebas/` para cada cambio de comportamiento.
- No subas PDF reales ni datos de clientes (`pruebas/ejemplos/` está ignorado; usa `pruebas/generar_ejemplos.py`).
- Sube la versión en `VERSION` y anota el cambio en `CHANGELOG.md`.
- Para publicar el instalador de Windows: `git tag vX.Y.Z && git push --tags` (workflow `instalador.yml`).
