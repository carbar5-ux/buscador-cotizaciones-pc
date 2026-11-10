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

- **Precios** (la primera hoja): una fila por suministro con empresa, fecha, descripción, cantidad, unidad, precio unitario, total, moneda y N° de cotización. **Clic en el suministro abre su cotización.** La columna **Var. vs anterior** muestra cuánto cambió el precio frente a la cotización anterior del mismo producto: **en rojo si subió**, en verde si bajó (y **Precio anterior** dice cuánto era, de quién y de cuándo).
- **Resumen**: n° de cotizaciones, proveedores, ítems, monto cotizado, vigentes, por vencer y vencidas.
- **Cotizaciones**: una fila por PDF, con vínculo para abrir el PDF, y el estado (Vigente / Por vencer / Vencida) en colores.
- **Ítems**: una fila por cada línea de cada cotización, lista para filtrar o hacer tablas dinámicas.
- **Comparativo**: agrupa el mismo producto aunque cada proveedor lo escriba distinto (columna Variantes) y muestra el precio mínimo y el proveedor más barato, el máximo, promedio, mediana, diferencia % y el último precio cotizado.
- **Proveedores**: cotizaciones, ítems y montos por proveedor, con sus datos de contacto.
- **Revisar**: lo que conviene verificar a mano (líneas que no cuadran, totales no encontrados, PDF escaneados o dañados, archivos duplicados, y **alzas de precio** de 5 % o más en la última cotización de cada producto).

## Instalación en Windows con el instalador (recomendada)

1. Entra a **Releases** del repositorio (<https://github.com/carbar5-ux/buscador-cotizaciones-pc/releases>) y descarga **`Instalar-Buscador-Cotizaciones-X.Y.Z.exe`**.
2. Doble clic → **Siguiente** → **Instalar**. No pide permisos de administrador ni hace falta instalar Python: el instalador trae todo.
3. Quedan los accesos **Buscador Cotizaciones**, **Excel Cotizaciones** y **Manual** en el Escritorio y en el menú Inicio (opcional: abrir al encender el PC).

Se instala en `%LOCALAPPDATA%\Programs\Buscador de cotizaciones`, sigue actualizándose solo desde GitHub y se desinstala en *Configuración → Aplicaciones*. Tu `config.toml`, el Excel y las cotizaciones no se tocan al reinstalar ni al desinstalar.

**Si ya usabas el programa instalado con `instalar.bat`:** la primera vez que abres el programa instalado, busca el anterior (en Descargas, Documentos, Escritorio, OneDrive o donde apunten sus accesos directos) y **trae solo tu configuración**: la carpeta de cotizaciones, `categorias.txt` y `sinonimos.txt`. Si el anterior se abría al encender el PC, ese inicio pasa a abrir el instalado. Después puedes borrar la carpeta antigua y sus accesos directos.

## Instalación en Windows con Python aparte (manual)

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

## Planilla y buscador de precios (en la ventana)

- **Buscar precio** (arriba a la derecha): escribe un suministro, por ejemplo `cable thhn 12`, y presiona Enter. Se abre la planilla con todos sus precios, del más barato al más caro, y el resumen de precio mínimo, último y promedio. Reconoce el mismo producto escrito distinto ("Conductor N°12" = "Cable 12"). Las medidas deben coincidir: 12 AWG no trae 14 AWG.
- **Planilla**: todas las líneas de todas las cotizaciones (descripción, precio, cantidad, unidad, total, empresa, fecha, N° y categoría), con filtros por texto, empresa, categoría, moneda y rango de fechas. Clic en un título ordena, doble clic en una fila abre el PDF, y **Exportar a Excel** guarda solo lo filtrado.
- **Var.**: cuánto cambió el precio frente a la cotización anterior del mismo producto. Las filas que subieron se ven en rojo.
- **Historial de precios**: elige una fila (un clic) o busca un producto y presiona el botón. Muestra un gráfico de cómo cambió el precio en el tiempo, con un color por proveedor, y la lista de cotizaciones (doble clic abre el PDF).
- **Exportar a PDF**: guarda lo filtrado en un PDF listo para enviar o imprimir (con los filtros usados y la suma de totales).
- Los sinónimos (conductor = cable, ampolleta = lámpara...) se editan en `sinonimos.txt` (botón **Sinónimos…** de la planilla).

Cuando llegan cotizaciones nuevas con un producto **más caro** que la vez anterior (5 % o más), la ventana lo avisa en su registro: `ALZA DE PRECIO +12%: ... (antes $ ..., proveedor, fecha)`.

## Cotizaciones que llegan por correo

El botón **Correo** de la ventana conecta el programa a tu correo. Cada 10 minutos revisa los correos de los últimos 30 días y guarda en `<carpeta de cotizaciones>\Correo\` los **PDF adjuntos** cuyo asunto, remitente o nombre de archivo contiene alguna de las palabras (`cotiz`, `presupuesto`, `proforma`, `oferta`; se pueden cambiar). Desde ahí se leen como cualquier cotización. **No borra ni marca como leído ningún correo**, y cada correo se revisa una sola vez.

- **Gmail**: necesitas una **contraseña de aplicación** (no tu clave normal). Activa la verificación en 2 pasos en tu cuenta de Google y crea la contraseña en <https://myaccount.google.com/apppasswords> (16 letras). Pégala en el botón **Correo** y presiona **Probar y guardar**.
- **Yahoo, iCloud**: el servidor se elige solo. Algunas cuentas también piden contraseña de aplicación.
- **Correo de la empresa** (tu@empresa.cl): deja *Servidor IMAP* vacío y el programa lo busca solo (mira quién atiende el correo del dominio: Google Workspace, Microsoft 365 o el hosting, y prueba `imap.` y `mail.`); el que funciona queda guardado. Si no lo encuentra, pregunta el "servidor IMAP" a quien administra el correo (también aparece en la configuración de la cuenta en Outlook o en el celular, como correo entrante).
- **Microsoft 365 / Outlook.com**: Microsoft ya no deja entrar con contraseña desde otros programas. Alternativa: una regla en Outlook que reenvíe los correos con cotizaciones a una cuenta Gmail, y conectar esa cuenta.

La contraseña queda guardada en `config.toml` de tu PC (por eso conviene la de aplicación: se puede borrar en cualquier momento sin cambiar tu clave). **Desactivar** la borra.

## Respaldo diario

Una vez al día el programa guarda una copia del Excel y de los datos leídos en `<carpeta del Excel>\Respaldos\` (con la fecha en el nombre) y conserva los últimos 14 días (`respaldos_dias`). Si algo se daña o se borra, copia el respaldo de vuelta junto al Excel quitándole la fecha del nombre (`Base_Cotizaciones_datos_2026-11-10.json` → `Base_Cotizaciones_datos.json`).

## Clasificar por categoría (Cables, Alumbrado, ...)

El programa copia cada cotización a carpetas por categoría: `<carpeta de cotizaciones>\Clasificadas\Cables\`, `...\Alumbrado\`, etc. Los originales no se mueven y el buscador no vuelve a leer las copias. Una cotización con ítems de varias categorías se copia a cada una.

Las categorías se definen en `categorias.txt` (botón **Categorías** de la ventana), una por línea:

```
Cables: cable, conductor, alambre, thhn, awg
Alumbrado: luminaria, foco, led, reflector, proyector
```

Cada PDF se copia a su carpeta **apenas se lee** (no al final). Si una cotización deja de pertenecer a una categoría (cambiaste las palabras) o borras el PDF original, su copia se quita sola de `Clasificadas` (y la carpeta, si queda vacía). El programa solo borra copias que hizo él. Una cotización entra a una categoría si alguna de sus líneas (descripción o código) o el nombre del archivo contiene una de esas palabras (sin importar mayúsculas ni tildes). Al guardar el archivo, el programa reclasifica todo solo. En el Excel, la columna **Categorías** (Cotizaciones) y **Categoría** (Ítems) permiten filtrar, y el Resumen cuenta cotizaciones por categoría. El botón **Clasificadas** abre la carpeta de copias.

## Abrir el PDF desde el Excel

Un clic en el **suministro** (la descripción) en las hojas **Precios** e **Ítems**, o en **Abrir PDF**, abre la cotización. Son vínculos normales de Excel (no fórmulas), así que funcionan aunque Excel tenga el cálculo en manual, y son relativos a la carpeta del Excel: siguen funcionando si se mueve la carpeta. También abren el PDF el nombre del archivo en Cotizaciones, los proveedores del Comparativo y los archivos de Revisar.

El Excel usa encabezados verdes, sin cuadrícula, montos con formato `$` (o `US$` / `UF`), fechas y unidades centradas, pestañas de color, un Resumen con tarjetas y está listo para imprimir (horizontal, ajustado al ancho). Se rehace solo cuando se actualiza el programa.

## Actualización automática

Al abrirse, y cada 6 horas mientras vigila, el programa revisa si hay una versión nueva en GitHub (compara el archivo `VERSION`). Si la hay, descarga solo los archivos del programa, instala las librerías nuevas si hacen falta y se reinicia solo. **Nunca toca `config.toml`, el Excel ni los datos.** Sin internet, sigue funcionando normal. Para buscar una actualización en el momento: botón **Actualizar** de la ventana (o `actualizar.bat`).

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
| `respaldos_dias` | Cuántos días de respaldo guardar en `Respaldos` (0 = sin respaldo). |
| `correo_usuario` / `correo_clave` | Correo y contraseña de aplicación para bajar cotizaciones (botón **Correo**). |
| `correo_servidor` / `correo_carpeta` | Servidor IMAP (vacío = automático) y bandeja (`INBOX`). |
| `correo_palabras` | Palabras que debe tener el asunto, remitente o archivo (`[]` = todos los PDF). |
| `correo_dias` / `correo_minutos` | Revisa los correos de los últimos N días, cada N minutos. |

## PDF escaneados (opcional)

Si una cotización es una foto o un escaneo, el PDF no tiene texto. Para leerlos instala **Tesseract OCR** (<https://github.com/UB-Mannheim/tesseract/wiki>, incluye el idioma español) y luego ejecuta `procesar.bat --reprocesar`. Sin Tesseract, esos PDF aparecen en la hoja **Revisar**.

## Cómo lee los PDF y sus límites

Los **catálogos, fichas técnicas y planos** que estén en la carpeta se detectan solos (no tienen palabras como precio, total, IVA o $): se leen en menos de un segundo, sin buscar tablas, y quedan en **Revisar** como "No parece una cotización". Las páginas con miles de trazos (planos, diagramas) tampoco se revisan buscando tablas.

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

## Linux / macOS

`./instalar.sh`, luego `./procesar.sh` o `./vigilar.sh` (requiere Python 3.11+ y tkinter para la ventana).

## Desarrollo

Ver [`CONTRIBUTING.md`](CONTRIBUTING.md). Pruebas: `pip install -r requirements-dev.txt && python -m pytest`. Cambios: [`CHANGELOG.md`](CHANGELOG.md). Licencia: [MIT](LICENSE).
