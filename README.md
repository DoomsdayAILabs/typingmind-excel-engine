# 📊 TypingMind Excel Engine

**Motor SQL local (DuckDB‑WASM) para TypingMind.** Convierte la pestaña del navegador en un
motor de datos real: carga archivos **CSV, Parquet y Excel**, explóralos con **SQL completo** y
deja que la **IA del chat los consulte por sí misma** con *function calling*, **sin que tus datos
salgan nunca de tu equipo**.

![Versión](https://img.shields.io/badge/versi%C3%B3n-1.0-10b981)
![DuckDB-WASM](https://img.shields.io/badge/DuckDB--WASM-1.29.0-yellow)
![SheetJS](https://img.shields.io/badge/SheetJS-xlsx.full.min.js-2ea44f)
![Privacidad](https://img.shields.io/badge/datos-100%25%20locales-3b82f6)
![Instalación](https://img.shields.io/badge/instalaci%C3%B3n-extensi%C3%B3n%20%2B%20plugin-8b5cf6)

---

## 🚀 Descripción

TypingMind Excel Engine es una **extensión + plugin** para [TypingMind](https://www.typingmind.com/)
que añade un motor de base de datos analítico dentro de la propia interfaz:

1. El usuario arrastra un **CSV / Parquet / XLSX / XLS** sobre un widget flotante inyectado en la página.
2. El archivo se registra en un **sistema de archivos virtual de DuckDB‑WASM** que vive dentro de un
   **Web Worker** (procesamiento pesado aislado de la interfaz, sin bloquear el chat).
3. DuckDB crea la tabla `excel_data` y la deja consultable con **SQL estándar** (filtros, agregaciones,
   CTEs, `JOIN`, funciones de ventana, `PIVOT`, `read_csv_auto`, tipos automáticos…).
4. Un **plugin de IA** expone la función `query_excel_data(sql_query)` al modelo, de modo que el LLM
   **escribe el SQL, recibe el resultado en Markdown y responde** con datos reales del archivo.

> **Nada se sube a ningún servidor.** No hay backend, no hay API keys de datos, no hay uploads:
> el archivo, el motor SQL y los resultados viven exclusivamente en memoria del navegador.

### ¿Qué problema resuelve?

- **Privacidad:** hojas de cálculo con datos sensibles que no se pueden pegar en un chat externo.
- **Tamaño:** archivos grandes (decenas de miles de filas) que no caben en el contexto de un LLM.
- **Precisión:** el modelo no "adivina" sumas ni promedios: **consulta el dato con SQL** y devuelve
  el resultado exacto que calcula el motor.
- **Agilidad:** explorar, verificar y exportar sin salir de TypingMind.

---

## ✨ Características principales

### 🗂️ Datos y formatos
- **Multi‑formato:** `.csv`, `.parquet`, `.xlsx` y `.xls` (Excel vía **SheetJS** dentro del Worker).
- **Archivos masivos:** CSV/Parquet se registran en el VFS de DuckDB y se leen con
  `read_csv_auto(..., sample_size=-1, ignore_errors=true)`, es decir, **escaneo completo** para
  inferir tipos correctamente en archivos largos; Parquet y Excel viajan al Worker con
  **transferables** (`ArrayBuffer` sin copia).
- **SQL completo:** JOINs, `GROUP BY`, subconsultas, CTEs (`WITH`), funciones de ventana,
  `REGEXP_*`, `PIVOT`, `EXTRACT`, casts y todo el motor DuckDB 1.29.0.
- **Tipos normalizados:** `BIGINT` fuera del rango seguro se devuelve como texto, `HUGENUM`/`Decimal`
  se normaliza, fechas → ISO 8601, binarios → arrays, `NULL` preservados (sin romper `JSON`).
- **Exportación CSV:** botón que descarga los resultados actuales como `tmee_resultados.csv`
  con escapado correcto de comillas.
- **Autocomprobación de esquema:** tras cargar, el motor devuelve filas totales, columnas, tipos y
  una previsualización de 10 filas (`information_schema.columns`).

### 🔒 Privacidad
- **100 % local:** el archivo nunca se sube; el VFS de DuckDB se **elimina** (`dropFile`) al terminar
  cada carga.
- **Sin persistencia:** al recargar la página los datos desaparecen (nada queda en disco).
- **Sin credenciales:** el plugin usa `AUTH_TYPE_NONE` y no requiere API keys de datos.

### 🧭 Interfaz (UI v1.0)
- **Widget flotante inyectado** en la página de TypingMind: arrastrar y soltar, consola SQL,
  tabla de resultados, metadatos y estado del motor.
- **Multitáctil y arrastrable:** se mueve libremente por la pantalla desde la cabecera
  (ratón: `mousedown/mousemove/mouseup` · táctil: `touchstart/touchmove/touchend`), con umbral de
  5 px para distinguir *clic* de *arrastre* y clamping al viewport.
- **Modo icono flotante (FAB táctil):** al minimizar deja de ser una barra ancha y se convierte en un
  **botón circular de 50×50 px** con el emoji 📊; ideal para móvil. Un clic en el círculo lo expande.
- **Inyección al chat:** el botón **💬 Enviar a TM** escribe los resultados como **tabla Markdown**
  en el input de TypingMind (dispara el evento `input` nativo para que React sincronice su estado),
  coloca el cursor al final y **minimiza el widget automáticamente**.
- **Indicador de estado** (badge `pending`/`success`/`error`) en todo momento.

### 🤖 IA autónoma (plugin)
- **Function calling** con la herramienta `query_excel_data`.
- **Solo lectura:** el plugin sanea el SQL (elimina comentarios) y bloquea `DROP`, `ALTER`, `INSERT`,
  `UPDATE`, `DELETE`, `CREATE`, `COPY`, `ATTACH`, `PRAGMA FORCE`… Solo `SELECT`, `WITH` y `EXPLAIN`.
- **Puente `postMessage`** en el canal `tm-excel-engine`: el iframe *sandbox* del plugin pide la
  consulta a la extensión y recibe las filas serializadas.
- **Respuestas autolimitadas:** el plugin recorta a **50 filas** y devuelve Markdown escapado, para
  no saturar el contexto del modelo; reintegra errores claros (timeout, motor ausente, SQL inválido)
  como texto que el propio LLM puede corregir.
- **Descubrimiento de esquema:** la descripción de la herramienta instruye al modelo a hacer primero
  `SELECT * FROM excel_data LIMIT 1;` si no conoce las columnas.

---

## 🏗️ Arquitectura del sistema

Tres piezas independientes, todas **dentro del navegador del usuario**:

```text
+--------------------------------------------------------------------------------------+
|                   NAVEGADOR: todo el procesamiento ocurre en local                   |
+--------------------------------------------------------------------------------------+
| TypingMind (ventana principal)                                                       |
|                                                                                      |
| [1] EXTENSION UI   typingmind-excel-engine-v1.0.js                                   |
|       - Widget flotante: drag & drop, consola SQL, tabla, CSV, FAB tactil            |
|       - Crea el Web Worker desde una Blob URL (bypass de CORS)                       |
|       - Puente postMessage del canal "tm-excel-engine"                               |
|             |  { type, requestId, ...payload }                                       |
|             v                                                                        |
| [2] WEB WORKER   duckdb-worker.js                                                    |
|       - DuckDB-WASM 1.29.0 (import dinamico) + SheetJS (XLSX)                        |
|       - Blob Worker interno para el binario de DuckDB (CORS del CDN)                 |
|       - VFS virtual: registrar archivo -> CREATE TABLE excel_data                    |
|             ^                                                                        |
| [3] PLUGIN IA   plugin/implementation.js                                             |
|       - iframe sandbox: sin acceso directo a DuckDB                                  |
|       - query_excel_data({ sql_query }) -> window.parent.postMessage                 |
|       - Recibe filas -> Markdown (max. 50) -> el LLM redacta la respuesta            |
|                                                                                      |
+--------------------------------------------------------------------------------------+
```

### Piezas y responsabilidades

| # | Pieza | Archivo | Se instala en | Responsabilidad |
|---|---|---|---|---|
| ① | **Extensión UI** | `typingmind-excel-engine-v1.0.js` | TypingMind → *Settings* → **Extensions** | Interfaz, lectura de archivos, orquestación de peticiones, puente `postMessage`, exportación CSV y Markdown |
| ② | **Motor (Worker)** | `duckdb-worker.js` | GitHub Pages (lo descarga la extensión) | DuckDB‑WASM, parsing de Excel con SheetJS, SQL, normalización de tipos, ciclo de vida del VFS |
| ③ | **Plugin IA** | `plugin/implementation.js` + `plugin/plugin.json` | TypingMind → **Plugins** | Expone `query_excel_data` al modelo, valida el SQL, puentea la consulta y formatea la respuesta |

### Flujo de datos (secuencia típica)

```text
[Usuario]          carga el archivo (CSV / Parquet / Excel) en el widget
     |
     v
[Extension UI]     lee el archivo y envia loadCSV | loadParquet | loadExcel al Worker
     |
     v
[Worker DuckDB]    read_csv_auto / read_parquet / SheetJS -> CREATE TABLE excel_data
     |             responde: registros, columnas, esquema y preview (10 filas)
     v
[Extension UI]     muestra metadatos y previsualiza los resultados
     |
     v
[Usuario]          pregunta en lenguaje natural en el chat
     |
     v
[LLM]              elige la herramienta query_excel_data y genera el SQL
     |
     v
[Plugin iframe]    window.parent.postMessage -> { channel, action: "execute_sql", sql }
     |
     v
[Worker DuckDB]    executeQuery -> filas normalizadas (BigInt, fechas y binarios listos)
     |
     v
[Plugin iframe]    tabla Markdown (max. 50 filas) -> el LLM responde con el dato real

Via alternativa (sin IA): ejecutar el SQL a mano en el widget y pulsar
"Enviar a TM" para pegar los resultados como Markdown en el chat.
```

### Protocolo interno (referencia rápida)

**① ↔ ② (Web Worker).** Sobre de petición: `{ type, requestId, ...payload }` → sobre de respuesta:
`{ success, requestId, data }` o `{ success: false, requestId, error }`.

| `type` | Payload | Devuelve |
|---|---|---|
| `initDuckDB` | `{}` | `{ version: "v1.0-phase-1b", status: "ready", bundle }` |
| `loadCSV` | `{ csvData, tableName }` | `{ procesamiento: "LOCAL_NATIVO_WORKER", formato, tabla, registros, columnas, esquema, preview }` |
| `loadParquet` | `{ parquetData, tableName }` | idem (`formato: "PARQUET"`) |
| `loadExcel` | `{ excelBuffer, tableName }` | idem (`formato: "EXCEL"`) |
| `executeQuery` | `{ sql }` | Array de filas normalizadas |

Timeouts de la extensión: `initDuckDB` y cargas **120 s**, `executeQuery` **60 s**.
La tabla destino siempre se sanea (`sanitizeTableName`) y se cita (`quoteIdentifier`).

**③ ↔ ① (puente del plugin).** El iframe del plugin (`window.parent`, origen `*`) envía:

```js
{ channel: "tm-excel-engine", action: "execute_sql", sql, requestId }
```

y la extensión responde al `event.source` (o hace *broadcast* a los iframes si no hay `source`):

```js
{ channel: "tm-excel-engine", requestId, success: true, data: rows }
{ channel: "tm-excel-engine", requestId, success: false, error: "…" }
```

### Detalles de ingeniería destacables

- **Bypass de CORS con Blob URL:** GitHub Pages y los CDN no permiten `new Worker()` cross‑origin.
  La extensión hace `fetch()` del worker, lo envuelve en un `Blob` y crea el Worker desde una
  `blob:` URL; la URL temporal se revoca (`revokeObjectURL`) en cuanto DuckDB termina de iniciarse.
- **Worker anidado:** DuckDB‑WASM necesita su propio `mainWorker`; el motor crea un **segundo Blob
  Worker** que solo contiene `importScripts("…bundle.mainWorker")`, replicando la misma técnica
  dentro del Worker para esquivar CORS.
- **Normalización numérica:** `BIGINT` se convierte a `Number` si está en rango seguro y a `string`
  si no; se detectan `HugeInt`/`Decimal` (patrón de 4 elementos `[a,0,0,0]` / `[a,-1,-1,-1]`) para
  evitar `TypeError: Do not know how to serialize a BigInt` al cruzar `postMessage`.
- **Aislamiento del iframe:** el plugin se ejecuta en un *sandbox* sin acceso a DuckDB, por lo que
  **todo** el cálculo ocurre en la ventana principal; el plugin solo transporta texto.

---

## ⚡ Guía de instalación rápida en TypingMind

> Requisitos: TypingMind (web o desktop) y conexión a Internet la primera vez
> (los binarios de DuckDB‑WASM y SheetJS se cargan desde jsDelivr).
> Todo lo demás ocurre en local.

### Paso 1 · La Extensión (el motor + la interfaz)

1. Abre TypingMind → **Settings → Extensions**.
2. Pega la etiqueta del script apuntando al archivo publicado en GitHub Pages:

```html
<script src="https://doomsdayailabs.github.io/typingmind-excel-engine/typingmind-excel-engine-v1.0.js"></script>
```

3. Guarda y **recarga TypingMind**.
4. Debe aparecer el widget flotante abajo a la izquierda, con el badge de estado en
   *“Inicializando Worker…”* y, unos segundos después, **“Motor Listo”** en verde.

> **¿Usas un fork propio?** Sustituye `doomsdayailabs/typingmind-excel-engine` por tu
> `usuario/repositorio` de GitHub Pages **y actualiza también la constante `WORKER_PATH`**
> dentro de `typingmind-excel-engine-v1.0.js` (línea 16), porque el worker se descarga de una URL
> absoluta. Con GitHub Pages activado en `main` y el archivo `.nojekyll` ya incluido, los dos
> archivos quedan servidos en la misma ruta base.

### Paso 2 · El Plugin (el cerebro que da herramientas a la IA)

1. Abre TypingMind → **Plugins → New Plugin / Create Plugin** (o **Import** pegando el JSON).
2. Rellena los campos tal y como vienen en `plugin/plugin.json`:

| Campo de TypingMind | Valor |
|---|---|
| Emoji | `📊` |
| Title | `Excel Data Engine (consulta local)` |
| Authentication | **No authentication** (`AUTH_TYPE_NONE`) |
| Implementation type | **JavaScript** |
| Output | **Give plugin output to the AI** (`respond_to_ai`) |
| OpenAI Spec | el bloque `openaiSpec` de `plugin/plugin.json` (función `query_excel_data`) |
| User settings | *(vacío)* |

3. En el **Interactive JavaScript Editor**, pega **íntegro** el contenido de
   `plugin/implementation.js` (la función debe llamarse exactamente `query_excel_data`, igual que
   el `name` del spec: es un requisito de TypingMind y no debe renombrarse).
4. Guarda el plugin y **actívalo en la conversación** (icono de plugins del chat).
5. Carga un archivo en el widget del motor **antes** de preguntar.

### Paso 3 · Verificación en 30 segundos

1. Widget visible y estado **Motor Listo**.
2. Arrastra `test-data/TEUs.xlsx` (o cualquier CSV) al panel → aparece *Tabla: `excel_data`*,
   el número de registros y la tabla de resultados.
3. Pregunta al chat: *“¿Cuántas filas tiene la tabla excel_data?”* con el plugin activo → el modelo
   llamará a `query_excel_data` y responderá con el dato real.
4. Opcional: pulsa **⬇️ CSV** para descargar los resultados y **💬 Enviar a TM** para pegarlos como
   Markdown en el chat.

### Alternativa sin GitHub Pages (uso local)

Puedes probar todo el motor sin publicar nada:

```bash
# clona el repositorio y sirve la raíz como archivos estáticos
python -m http.server 8080
# o: npx serve .
```

- `test-widget.html` → simulador de TypingMind con el widget real.
- `test-duckdb-worker.html` → banco de pruebas del Worker (carga y consulta sin UI).
- `tests/sql-test-runner.html` → batería de consultas SQL contra los datos de ejemplo.

> Con `file://` los Workers y los módulos ES quedan bloqueados por el navegador: usa siempre un
> servidor local (`http://localhost`) para las pruebas.

---

## 🎛️ Guía de uso del widget

| Acción | Cómo |
|---|---|
| **Cargar datos** | Arrastra el archivo sobre la zona punteada o pulsa *Seleccionar Archivo*. Acepta `.csv`, `.parquet`, `.xlsx`, `.xls`. |
| **Consultar** | Edita el SQL en la consola y pulsa **▶ Ejecutar Consulta**. |
| **Ver resultados** | La tabla muestra hasta **100 filas** con los tipos ya convertidos; cabecera fija al hacer scroll. |
| **Exportar** | **⬇️ CSV** descarga los resultados actuales (`tmee_resultados.csv`, UTF‑8 con comillas escapadas). |
| **Enviar a la IA** | **💬 Enviar a TM** escribe la tabla Markdown en el input del chat (máx. 50 filas) y minimiza el widget. |
| **Minimizar** | Un **clic** (o *tap*) en la cabecera → el widget se convierte en el **círculo flotante 📊**. |
| **Expandir** | Un clic en el círculo → vuelve al panel completo. |
| **Mover** | Arrastra desde la **cabecera** (ratón o dedo). El widget nunca se sale de la pantalla. |

### Ejemplo de conversación con IA autónoma

```text
Tú:  He cargado el Excel de TEUs. ¿Cuáles son los 5 clientes con más contenedores?

IA:  [llama a query_excel_data({ sql_query: "SELECT * FROM excel_data LIMIT 1;" })]  ← descubre columnas
     [llama a query_excel_data({ sql_query:
       'SELECT "Cliente", SUM("TEUs") AS total FROM excel_data
        GROUP BY 1 ORDER BY total DESC LIMIT 5;' })]
     → Respuesta con la tabla y el comentario de negocio.
```

### Buenas prácticas del prompt (para el usuario)

- Menciona el archivo y el objetivo: *“Usa la tabla `excel_data` del motor local para…”*.
- Pide agregados en vez de volcados: *“agrupa por mes y suma TEUs”* es más rápido y cabe en contexto.
- Si el modelo se equivoca de columna, responde: *“primero ejecuta `SELECT * FROM excel_data LIMIT 1`
  y usa los nombres exactos”*.
- Para el modo manual, ejecuta el SQL en el widget y usa **💬 Enviar a TM**.

### Notas de SQL útiles

- Las columnas se citan con **comillas dobles**: `SELECT "Cliente", SUM("TEUs") FROM excel_data`.
- La tabla es siempre **`excel_data`** (se recrea con cada carga: `CREATE OR REPLACE TABLE`).
- DuckDB acepta `SELECT * EXCLUDE (...)`, `PIVOT`, `QUALIFY`, `GROUP BY ALL`, `USING SAMPLE`,
  `READ_*` y funciones de ventana; cualquier consulta válida de DuckDB 1.29.0 funciona.
- Antes de responder preguntas de la IA, un `SELECT * FROM excel_data LIMIT 1;` evita el 90 % de los
  errores de nombres de columna.

### Uso en móvil

El widget está pensado para pantallas pequeñas: la cabecera es el asa de arrastre y el interruptor
de minimizar, y en modo icono ocupa solo **50×50 px** en la esquina que elijas (se recuerda mientras
no recargues la página).

---

## 🔐 Seguridad y privacidad

| Aspecto | Estado |
|---|---|
| Ubicación de los datos | Memoria del navegador (VFS de DuckDB dentro del Worker). Nunca se suben. |
| Persistencia | Ninguna: al recargar la página se pierde todo (el VFS se limpia con `dropFile`). |
| Permisos de red | Solo para descargar los binarios de DuckDB‑WASM y SheetJS desde jsDelivr. |
| Superficie de escritura de la IA | Bloqueada: el plugin rechaza cualquier sentencia que no sea `SELECT`/`WITH`/`EXPLAIN`. |
| Credenciales | El plugin es `AUTH_TYPE_NONE`; no hay claves ni tokens. |
| Alcance del plugin | Corre en un iframe *sandbox* sin acceso al motor: solo intercambia texto por `postMessage`. |
| Qué ve el LLM | Únicamente lo que llega al chat: tu pregunta y las tablas Markdown (≤ 50 filas) que inyectes o que el plugin devuelva por una consulta concreta. |

> La consola SQL del **widget** permite al *usuario* ejecutar cualquier sentencia (es su propio
> motor local). La restricción de solo lectura se aplica únicamente a las consultas que **la IA**
> puede lanzar por sí misma.

---

## ⚠️ Límites conocidos

- **Una tabla a la vez:** cada carga recrea `excel_data` (`CREATE OR REPLACE TABLE`). Para trabajar con
  dos archivos, usa `UNION` o reexporta la tabla desde la consola SQL antes de cargar el segundo.
- **Excel de una sola hoja:** SheetJS procesa la **primera hoja** del libro.
- **Excel en memoria:** `.xlsx/.xls` se parsean completos con SheetJS antes de crear la tabla; para
  archivos muy grandes conviene convertir antes a **CSV o Parquet**, que se leen con el escaneo
  optimizado de DuckDB.
- **Vista previa acotada a propósito:** 100 filas en pantalla, 50 al inyectar en el chat y 50 en la
  respuesta del plugin (para no saturar el contexto del modelo).
- **Extensión y plugin deben convivir en la misma pestaña:** el plugin habla con el motor por
  `postMessage`; si la extensión no está cargada o no hay archivo procesado, devuelve un error legible.
- **Sin persistencia ni historial:** recargar la página descarta el motor y los datos.
- **Dependencia de CDN:** sin acceso a jsDelivr la primera vez no se puede iniciar DuckDB‑WASM/SheetJS
  (se puede autoalojar copiando los bundles y cambiando `DUCKDB_PACKAGE` en `duckdb-worker.js`).
- **Timeout del plugin:** 30 s por consulta (configurable en `plugin/implementation.js`).

---

## 📁 Estructura del repositorio

```text
typingmind-excel-engine/
├── typingmind-excel-engine-v1.0.js   ← ① PRODUCTO: Extensión (motor UI + puente postMessage)
├── duckdb-worker.js                  ← ② PRODUCTO: Web Worker (DuckDB-WASM + SheetJS)
├── plugin/
│   ├── plugin.json                   ← ③ PRODUCTO: spec del plugin (formato nativo TypingMind)
│   ├── implementation.js             ← ③ PRODUCTO: código del plugin (query_excel_data)
│   └── README.md                     ← Documentación de la integración del plugin
├── README.md                         ← Este archivo
├── index.html                        ← GitHub Pages: historial de versiones descargables
├── .nojekyll                         ← Necesario para servir el repo tal cual en Pages
│
├── test-widget.html                  ← DEV: simulador de TypingMind (widget real, sin IA)
├── test-duckdb-worker.html           ← DEV: banco de pruebas del Worker
├── tests/
│   ├── fase5-verificacion.html       ← DEV: 44 checks (UI móvil, FAB, arrastre, inyección)
│   ├── fase3b-verificacion.html      ← DEV: 25 checks de regresión (Markdown, CSV, puente)
│   ├── fase3b-visual.html            ← DEV: capturas visuales (#min / #drag / #send)
│   ├── sql-test-runner.html          ← DEV: batería de consultas SQL
│   └── run-headless.js               ← DEV: runner de los tests en Chrome/Edge headless
├── test-data/                        ← Datos de ejemplo (TEUs.xlsx, RequerimientoPrueba.xlsx)
│
├── backup/duckdb-worker.v0.4.23-contaminado.js  ← LEGADO (histórico, no usar)
└── duckdb/typingmind-excel-engine-v0.3-test.js  ← LEGADO (prototipo v0.3, no usar)
```

---

## 🧪 Desarrollo y pruebas

Los tests son **páginas HTML autocontenidas** que publican su resultado en `document.title`
(prefijo `RES:`), de modo que se pueden ejecutar sin dependencias ni frameworks:

```bash
# Verificación de la Fase 5 (UI móvil, FAB, arrastre, inyección) → 44 checks
node tests/run-headless.js tests/fase5-verificacion.html

# Regresión Fase 3B (Markdown, CSV, puente postMessage) → 25 checks
node tests/run-headless.js tests/fase3b-verificacion.html

# Ver todos los checks, no solo los fallos
node tests/run-headless.js tests/fase5-verificacion.html 900,760 --all
```

- `tests/run-headless.js` localiza Chrome o Edge automáticamente (o `CHROME_PATH`).
- Capturas visuales: `tests/fase3b-visual.html`, `#min`, `#drag`, `#send`.
- En headless las transiciones CSS no avanzan: el test de Fase 5 las desactiva para medir la
  geometría final del icono flotante (la declaración de la transición se verifica sobre el CSS).
- Comprobación rápida de sintaxis: `node --check typingmind-excel-engine-v1.0.js`.

Estado verificado de esta entrega: `44/44` checks de Fase 5 y `25/25` de regresión Fase 3B.

---

## 🛠️ Solución de problemas

| Síntoma | Causa probable | Solución |
|---|---|---|
| No aparece el widget | Extensión no guardada o no recargada / bloqueador de scripts | Guarda la extensión, recarga TypingMind y revisa la consola del navegador |
| *“Error de inicialización”* o *“Fallo al iniciar: Failed to fetch”* | `WORKER_PATH` no apunta a un `duckdb-worker.js` accesible | Verifica la URL de GitHub Pages (y la constante `WORKER_PATH` si usas un fork) |
| *“Error en Worker”* | CDN bloqueado (red corporativa, sin Internet) | Permite `cdn.jsdelivr.net` o autoaloja los bundles de DuckDB‑WASM/SheetJS |
| *“Error CSV/Excel”* al cargar | Archivo vacío, corrupto o formato inesperado | Prueba con `test-data/TEUs.xlsx`; para CSV confirma separador y codificación UTF‑8 |
| *“Error SQL”* | Nombre de columna inexistente o sintaxis | `SELECT * FROM excel_data LIMIT 1;` para ver columnas y tipos reales |
| La IA dice que el motor no responde (timeout) | Extensión no cargada, plugin desactivado o ningún archivo procesado | Carga el archivo en el widget y activa el plugin en la conversación |
| La IA intenta escribir datos | Comportamiento esperado del *guardrail* | Solo se permiten `SELECT`/`WITH`/`EXPLAIN`; pídele una consulta de lectura |
| Los resultados se ven recortados | Límites deliberados (100 en pantalla, 50 en chat) | Añade `LIMIT`/agregaciones; exporta por **⬇️ CSV** si necesitas todo |

---

## 🗺️ Historial de hitos

| Fase | Entregable |
|---|---|
| Prototipos v0.3 / v0.4.x | Primeras aproximaciones (ver carpetas `duckdb/` y `backup/`, no usar) |
| Fase 1B | Motor DuckDB‑WASM en Web Worker + proxy Blob para evitar CORS (`v1.0-phase-1b`) |
| Fase 1C | Soporte **Parquet** con `read_parquet` |
| Fase 1D | Soporte **Excel** con SheetJS dentro del Worker |
| Puente IA | Plugin `query_excel_data` + exportación **CSV** |
| Fase 3B | Minimizar/maximizar, inyección de resultados al chat como **Markdown** |
| Fase 5 | **UI móvil**: icono flotante de 50 px + widget **arrastrable** (ratón y táctil) |
| Fase 6 | **Documentación y empaquetado**: README, guía de instalación y carpeta `plugin/` definitiva |

---

## 🤝 Contribuir

1. No modifiques la lógica de `typingmind-excel-engine-v1.0.js` ni de `duckdb-worker.js` sin
   ejecutar antes las suites de `tests/` (deben seguir en verde).
2. Añade o amplía los checks en `tests/` cuando cambies el comportamiento del widget o del plugin.
3. Mantén la documentación sincronizada: `README.md` (producto) y `plugin/README.md` (integración IA).

## 📄 Licencia

Este repositorio **aún no incluye un archivo `LICENSE`**: los derechos quedan reservados por el
autor hasta que se defina una licencia explícita. Si quieres reutilizarlo o redistribuirlo, abre un
*issue* en el repositorio para acordar la licencia.

---

<p align="center">
  Hecho para TypingMind · DuckDB‑WASM 1.29.0 · SheetJS · 100 % en tu navegador
</p>
