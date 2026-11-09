# Buscador de cotizaciones PC

Creado por **Carlos**. Manual de uso: [`Manual_de_uso.pdf`](Manual_de_uso.pdf).

Programa en **Python** que busca todas las cotizaciones en PDF de una carpeta de tu computador (y sus subcarpetas), lee su contenido y arma un **Excel base** con toda la información. Se usa desde una **ventana de Windows** con barra de progreso en colores, contadores y botones para abrir el Excel, la carpeta y el manual. Si lo dejas en **modo vigilancia**, cada vez que guardas una cotización nueva en la carpeta el Excel se actualiza solo.

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
| `vigilar.bat` | Abre la **ventana del buscador** (lo abre el ícono verde): lee las cotizaciones y queda vigilando la carpeta. |
| `vigilar_consola.bat` | Lo mismo, pero en texto (consola), por si la ventana no abre. |
| `iniciar_con_windows.bat` | Hace que la vigilancia parta sola y en segundo plano cada vez que enciendes el PC. |
| `abrir_excel.bat` | Abre el Excel base directamente. |
| `actualizar.bat` | Busca e instala la última versión ahora. |
| `abrir_manual.bat` | Abre el manual de uso en PDF (`Manual_de_uso.pdf`). |
| `crear_acceso_directo.bat` | Pone en el Escritorio accesos directos: **Buscador Cotizaciones** (verde $), **Excel Cotizaciones** (naranjo $) y **Manual Cotizaciones** (PDF). `instalar.bat` ya los crea. |

Mientras lee, muestra una barra con el porcentaje y el tiempo que falta. Usa **todos los núcleos del procesador** (un PDF por núcleo) con prioridad baja, para no trabar el PC mientras trabajas. El texto se lee con pdfium, que es muy rápido, y el lector de tablas solo se usa en las páginas que tienen líneas dibujadas. La primera lectura de una carpeta grande es la que más demora; después solo lee las cotizaciones nuevas. El Excel se crea al partir y se actualiza cada minuto mientras lee.

**En pantalla:** la barra de progreso tiene colores (rojo, amarillo y verde según avanza) y queda siempre visible. Al terminar aparece **TERMINADO** en letras grandes rojas con la cantidad de cotizaciones leídas. Mientras vigila, la última línea muestra que el programa sigue trabajando y cuántas cotizaciones hay en el Excel.

**Aviso al terminar:** cuando termina de leer, suena un aviso y aparece una ventana con la cantidad de cotizaciones leídas, el tiempo que tomó y un botón para abrir el Excel. En modo vigilancia, si llegan pocas cotizaciones (menos de `avisar_desde`), solo suena.

La primera vez te pregunta cuál es tu carpeta de cotizaciones y la guarda en `config.toml`. Para cambiarla: `procesar.bat --configurar`.

El Excel queda en tu carpeta de cotizaciones como `Base_Cotizaciones.xlsx`. Junto a él queda `Base_Cotizaciones_datos.json`, la memoria del programa: solo lee los PDF nuevos o modificados, y si renombras o mueves un PDF no lo vuelve a leer.

> Si el Excel está abierto cuando llega una cotización, el programa espera: al cerrarlo se actualiza en la siguiente pasada. El Excel se regenera completo cada vez, así que **no escribas a mano en él** (haz tus análisis en otro libro que lo consulte).

## Clasificar por categoría (Cables, Alumbrado, ...)

El programa copia cada cotización a carpetas por categoría: `<carpeta de cotizaciones>\Clasificadas\Cables\`, `...\Alumbrado\`, etc. Los originales no se mueven y el buscador no vuelve a leer las copias. Una cotización con ítems de varias categorías se copia a cada una.

Las categorías se definen en `categorias.txt` (botón **Categorías** de la ventana), una por línea:

```
Cables: cable, conductor, alambre, thhn, awg
Alumbrado: luminaria, foco, led, reflector, proyector
```

Una cotización entra a una categoría si alguna de sus líneas (descripción o código) o el nombre del archivo contiene una de esas palabras (sin importar mayúsculas ni tildes). Al guardar el archivo, el programa reclasifica todo solo. En el Excel, la columna **Categorías** (Cotizaciones) y **Categoría** (Ítems) permiten filtrar, y el Resumen cuenta cotizaciones por categoría. El botón **Clasificadas** abre la carpeta de copias.

## Abrir el PDF desde el Excel

En las hojas **Cotizaciones** e **Ítems**, la columna **PDF** tiene "Abrir PDF": un clic abre la cotización. También abren el PDF el nombre del archivo, los proveedores del **Comparativo** y los archivos de **Revisar**.

## Actualización automática

Al abrirse, y cada 6 horas mientras vigila, el programa revisa si hay una versión nueva en GitHub (compara el archivo `VERSION`). Si la hay, descarga solo los archivos del programa, instala las librerías nuevas si hacen falta y se reinicia solo. **Nunca toca `config.toml`, el Excel ni los datos.** Sin internet, sigue funcionando normal. Para buscar una actualización a mano: `actualizar.bat`.

Si el repositorio de GitHub es **privado**, el PC necesita una clave de solo lectura:

1. En GitHub: foto de perfil → **Settings → Developer settings → Personal access tokens → Fine-grained tokens → Generate new token**.
2. En *Repository access* elige **Only select repositories → buscador-cotizaciones-pc**.
3. En *Permissions → Repository permissions → Contents* elige **Read-only**. No le des ningún otro permiso.
4. Copia la clave y pégala en `config.toml`: `token_github = 'github_pat_...'`.

Si el repositorio es **público**, no hace falta clave.

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
| `aviso_sonido` / `aviso_ventana` | Sonido y ventana al terminar de leer. |
| `avisar_desde` | En modo vigilancia, la ventana sale solo si llegan al menos estas cotizaciones juntas. |
| `actualizar_automatico` | Buscar e instalar versiones nuevas solo. |
| `horas_entre_actualizaciones` | Cada cuánto revisa mientras vigila. |
| `token_github` | Clave de solo lectura, solo si el repositorio es privado. |
| `max_paginas_tablas` | El lector de tablas revisa solo las primeras N páginas de cada PDF (10); el texto se lee completo. Acelera mucho los catálogos. |
| `clasificar` | Copiar las cotizaciones a carpetas por categoría (ver `categorias.txt`). |
| `procesos` | Cuántos PDF leer a la vez (0 = todos los núcleos del procesador). |

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
