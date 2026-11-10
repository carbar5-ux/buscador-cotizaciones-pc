# Cambios

## Sin publicar (repositorio)
- Se elimina el workflow que generaba un `.exe`: no abría la ventana, ignoraba `config.toml` y el
  actualizador se reinstalaba en cada inicio. El programa se instala con `instalar.bat`.

## 1.9.8
- Buscador/Planilla: el resumen (mínimo, último, promedio) se calcula para el producto más repetido entre los
  resultados (mismas medidas, unidad y moneda), no mezclando productos distintos que contienen la palabra
  buscada. Avisa cuántos productos más hay. El verde marca el más barato de ese producto.
- Al abrir, se borran los temporales que quedan si el programa se cerró a mitad de un guardado
  (`~tmp_*.xlsx`, `*_datos.tmp`, `~act_*`, de más de 10 minutos).

## Sin publicar
- Licencia MIT, CI (pruebas + ruff), `pyproject.toml`, plantillas de issues/PR y Dependabot.
- Más pruebas (agrupador, actualizador).
- Descarga de actualizaciones con límite de tamaño y protección contra rutas peligrosas.
- Scripts `instalar.sh`, `procesar.sh` y `vigilar.sh` para Linux/macOS.

## 1.9.7
- Versión previa a este historial.
