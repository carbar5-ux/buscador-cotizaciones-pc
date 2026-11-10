# Cambios

## 1.10.4
- Excel ordenado **de la más nueva a la más antigua** en Precios, Cotizaciones, Ítems y Revisar: por la fecha
  de la cotización; las que no traen fecha se ubican por la fecha del archivo (cuándo llegó) en vez de quedar al
  final; con la misma fecha, primero la última leída. La hoja Proveedores, del que cotizó más recientemente.

## 1.10.3
- **Origen de cada cotización: Correo o PC.** La Planilla tiene la columna *Origen* ("✉ Correo" en azul o "PC"),
  un filtro *Origen* y el resumen dice cuántas llegaron por correo; también el Historial. En el Excel, columna
  *Origen* en Precios y Cotizaciones (Correo en azul) y tarjeta "Llegaron por correo" en el Resumen. También en
  Exportar a Excel y a PDF.
- Planilla: los botones Exportar quedan junto al buscador.

## 1.10.2
- Botón **Actualizar** en la ventana: busca una versión nueva en GitHub en el momento; si hay, pregunta,
  la instala y el programa se vuelve a abrir solo. Si no, avisa "Ya tienes la última versión".

## 1.10.1
- No se acepta un disco completo (C:\, D:\), la carpeta de usuario ni carpetas del sistema como carpeta de
  cotizaciones: en la raíz del disco Windows no deja guardar el Excel y se leerían todos los PDF del PC. Si ya
  estaba configurada así, la ventana avisa y pide elegir la carpeta correcta (tampoco se trae del programa anterior).
- Correo de empresa: si no se indica el servidor IMAP, se busca solo (registro MX del dominio: Google Workspace,
  Microsoft 365 o hosting; prueba imap. y mail.) y el que funciona queda guardado. Mensajes de error que dicen
  qué hacer (servidor no encontrado, clave de Google, Microsoft sin acceso por contraseña).

## 1.10.0
- **Alzas de precio**: la hoja Precios tiene la columna *Var. vs anterior* (rojo si subió, verde si bajó) y
  *Precio anterior*; Revisar lista las alzas de 5 % o más; el Resumen las cuenta. La ventana avisa las alzas
  de las cotizaciones recién leídas y la Planilla muestra la columna *Var.* con las alzas en rojo.
- **Historial de precios** (botón en la Planilla): gráfico del precio de un producto en el tiempo, un color por
  proveedor, con la lista de cotizaciones.
- **Exportar a PDF** lo filtrado en la Planilla (nueva librería: reportlab; se instala sola al actualizar).
- **Cotizaciones por correo** (botón Correo): baja por IMAP los PDF adjuntos con "cotiz", "presupuesto"...
  a `<carpeta>/Correo/`. Solo lee: no borra ni marca correos.
- **Respaldo diario** del Excel y los datos en `Respaldos/` (últimos 14 días).
- **Clasificadas al día**: se quitan las copias de cotizaciones que cambiaron de categoría o cuyo PDF original
  se borró (solo copias hechas por el programa; también reconoce las de versiones anteriores).
- **Programa instalado**: la primera vez trae la configuración del programa anterior (carpeta, categorías y
  sinónimos) y cambia su inicio automático al instalado.
- **Manual de uso** nuevo (10 páginas, con imágenes): instalador, ventana, Planilla, historial y alzas de
  precio, Excel, categorías, correo (con los pasos de Gmail), respaldo, configuración y problemas frecuentes.
- GitHub Actions: checkout v7 y setup-python v7.

## Sin publicar (repositorio)
- Se reemplaza el workflow que generaba un `.exe` (no abría la ventana, ignoraba `config.toml` y el
  actualizador se reinstalaba en cada inicio) por un **instalador de Windows** (Inno Setup) con Python
  embebido, la ventana y las librerías incluidas: `Instalar-Buscador-Cotizaciones-X.Y.Z.exe`. Se arma y se
  prueba en GitHub Actions (`instalador/`), y se publica en Releases al crear una etiqueta `vX.Y.Z`.

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
