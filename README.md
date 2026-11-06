# Buscador de cotizaciones PC

Programa en **Python** que busca todas las cotizaciones en PDF de una carpeta de tu computador (y sus subcarpetas), lee su contenido y arma un **Excel base** con toda la información. Si lo dejas en **modo vigilancia**, cada vez que guardas una cotización nueva en la carpeta el Excel se actualiza solo.

## Qué saca de cada cotización

| Encabezado | Totales | Cada línea de detalle |
| --- | --- | --- |
| Proveedor y RUT | Neto / subtotal | Ítem y código |
| Cliente | Descuento | Descripción (aunque ocupe varias líneas) |
| N° de cotización | IVA | Cantidad y unidad |
| Fecha | Total | Precio unitario |
| Validez y fecha de vencimiento | Suma de ítems (control) | Descuento de la línea |
| Moneda (CLP, UF, USD, EUR) | | Total de la línea |
| Forma de pago, plazo y lugar de entrega | | Control: cantidad × precio = total |
| Contacto, email y teléfono | | Página del PDF |

## El Excel (`Base_Cotizaciones.xlsx`)

- **Resumen**: n° de cotizaciones, proveedores, ítems, monto cotizado, vigentes, por vencer y vencidas.
- **Cotizaciones**: una fila por PDF, con vínculo para abrir el PDF, y el estado (Vigente / Por vencer / Vencida) en colores.
- **Ítems**: una fila por cada línea de cada cotización, lista para filtrar o hacer tablas dinámicas.
- **Comparativo**: agrupa el mismo producto en distintas cotizaciones y muestra el precio mínimo y el proveedor más barato, el máximo, promedio, mediana, diferencia % y el último precio cotizado.
- **Proveedores**: cotizaciones, ítems y montos por proveedor, con sus datos de contacto.
- **Revisar**: lo que conviene verificar a mano (líneas que no cuadran, totales no encontrados, PDF escaneados o dañados, archivos duplicados).

## Instalación en Windows (una sola vez)

1. Instala **Python 3.11 o más nuevo** desde <https://www.python.org/downloads/>. En el instalador marca **“Add python.exe to PATH”**.
2. Descarga este repositorio (**Code → Download ZIP**) y descomprímelo en una carpeta fija, por ejemplo `C:\BuscadorCotizaciones`.
3. Haz doble clic en **`instalar.bat`**.

## Uso

| Archivo | Qué hace |
| --- | --- |
| `procesar.bat` | Lee la carpeta una vez y actualiza el Excel. |
| `vigilar.bat` | Queda abierto revisando la carpeta cada 10 segundos y actualiza el Excel cuando llegan, cambian o se borran cotizaciones. |
| `iniciar_con_windows.bat` | Hace que la vigilancia parta sola y en segundo plano cada vez que enciendes el PC. |
| `abrir_excel.bat` | Abre el Excel base directamente. |

Mientras lee, muestra una barra con el porcentaje y el tiempo que falta, y lee varios PDF a la vez (uno por núcleo del procesador). La primera lectura de una carpeta grande es la que más demora; después solo lee las cotizaciones nuevas. El Excel se crea al partir y se va actualizando cada 25 cotizaciones leídas.

La primera vez te pregunta cuál es tu carpeta de cotizaciones y la guarda en `config.toml`. Para cambiarla: `procesar.bat --configurar`.

El Excel queda en tu carpeta de cotizaciones como `Base_Cotizaciones.xlsx`. Junto a él queda `Base_Cotizaciones_datos.json`, la memoria del programa: solo lee los PDF nuevos o modificados, y si renombras o mueves un PDF no lo vuelve a leer.

> Si el Excel está abierto cuando llega una cotización, el programa espera: al cerrarlo se actualiza en la siguiente pasada. El Excel se regenera completo cada vez, así que **no escribas a mano en él** (haz tus análisis en otro libro que lo consulte).

## Configuración (`config.toml`)

| Opción | Para qué sirve |
| --- | --- |
| `carpeta_cotizaciones` | Carpeta donde guardas los PDF. |
| `archivo_excel` | Dónde dejar el Excel (vacío = dentro de la carpeta de cotizaciones). |
| `incluir_subcarpetas` | Buscar también en subcarpetas. |
| `intervalo_segundos` | Cada cuánto revisa en modo vigilancia. |
| `formato_numeros` | `'chile'` (1.234.567,89) o `'us'` (1,234,567.89). |
| `dias_validez_por_defecto` | Validez que se asume si la cotización no la indica. |
| `dias_aviso_vencimiento` | Días antes del vencimiento para marcar “Por vencer”. |
| `mis_ruts` | RUT de tu empresa, ej. `['77.888.999-K']`, para no confundirlo con el del proveedor. |
| `quitar_borrados` | Si borras un PDF, también se quita del Excel. |
| `procesos` | Cuántos PDF leer a la vez (0 = automático). |

## PDF escaneados (opcional)

Si una cotización es una foto o un escaneo, el PDF no tiene texto. Para leerlos instala **Tesseract OCR** (<https://github.com/UB-Mannheim/tesseract/wiki>, incluye el idioma español) y luego ejecuta `procesar.bat --reprocesar`. Sin Tesseract, esos PDF aparecen en la hoja **Revisar**.

## Cómo lee los PDF y sus límites

Cada proveedor arma su cotización distinta, así que el programa usa reglas generales: busca la tabla de detalle por sus encabezados (Descripción, Cantidad, Precio unitario, Total…) y, si el PDF no tiene tabla con bordes, reconoce las líneas que terminan en cantidad, precio y total y comprueba que cuadren. Los datos del encabezado se buscan por sus etiquetas habituales (“Fecha:”, “Forma de pago:”, “Validez:”, “RUT”…).

Funciona bien con cotizaciones generadas por sistemas (ERP, facturadores, Excel o Word exportados a PDF). Con formatos poco comunes puede faltar algún dato: siempre revisa la hoja **Revisar** y la columna **Control** de Ítems. El comparativo agrupa productos con la misma descripción y unidad; si dos proveedores escriben distinto el mismo producto, aparecerán en filas separadas.

## Uso avanzado

```bash
python buscador_cotizaciones.py --carpeta "D:\Cotizaciones" --excel "D:\Base.xlsx" --vigilar
```

Pruebas:

```bash
python pruebas/test_extractor.py
```

```bash
python pruebas/generar_ejemplos.py pruebas/ejemplos
```
