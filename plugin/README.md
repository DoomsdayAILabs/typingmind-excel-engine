# Plugin TypingMind — Excel Data Engine (Opción A)

El motor sigue siendo `typingmind-excel-engine-v1.0.js`, instalado como **Extensión**.
Este plugin (function calling) permite que la IA del chat consulte por sí misma los
datos cargados, escribiendo y ejecutando SQL localmente con DuckDB-WASM.

## Piezas

| Archivo | Qué es | Dónde se instala |
|---|---|---|
| `../typingmind-excel-engine-v1.0.js` | Motor + UI + puente postMessage | TypingMind → Settings → **Extensions** |
| `plugin.json` | Metadatos + especificación de la herramienta | TypingMind → **Plugins** → Create / Import |
| `implementation.js` | Código del plugin (iframe) | Campo JavaScript del plugin |

> Ya no hace falta un script de puente aparte: el listener `postMessage` está integrado
> en el propio motor (en la ventana principal).

## Referencia de campos de `plugin.json`

`plugin.json` es **JSON puro y válido**: el formato JSON no admite comentarios, así que la
documentación de cada campo vive aquí (no añadas claves nuevas al archivo: TypingMind importa el
objeto tal cual).

| Campo | Valor en este plugin | Para qué sirve |
|---|---|---|
| `uuid` | `7f3c9e21-4b8a-4d6e-9f11-2c8a1b0e4d77` | Identificador único del plugin. **No lo cambies** si reimportas: evita duplicados. |
| `id` | `query_excel_data` | Id interno del plugin. |
| `emoji` | `📊` | Icono en la lista de plugins. |
| `title` | `Excel Data Engine (consulta local)` | Nombre visible para el usuario. |
| `version` | `1` | Versión del plugin (entero). |
| `authenticationType` | `AUTH_TYPE_NONE` | Sin credenciales: los datos no salen del navegador. |
| `implementationType` | `javascript` | El código se ejecuta en el iframe del plugin. |
| `outputType` | `respond_to_ai` | La salida se entrega al modelo como resultado de herramienta. |
| `userSettings` | `[]` | Sin campos configurables por el usuario. |
| `openaiSpec.name` | `query_excel_data` | **Debe coincidir exactamente** con el nombre de la función en `implementation.js`. |
| `openaiSpec.description` | *(texto en español)* | Prompt de la herramienta: explica que la tabla es `excel_data`, que solo hay lectura y que la respuesta llega como tabla Markdown. |
| `openaiSpec.parameters` | `sql_query` (string, **requerido**) | Contrato de entrada: la consulta SQL que genera el modelo. |

## Flujo

1. El usuario carga el Excel en el panel del motor (Extensión).
2. El usuario hace una pregunta en lenguaje natural.
3. El LLM llama `query_excel_data` con un `sql_query` (solo SELECT/WITH).
4. El plugin (iframe) envía la consulta por `window.parent.postMessage`.
5. El motor ejecuta `executeQuery` en DuckDB y responde las filas.
6. El plugin devuelve a la IA un Markdown tabular (máx. 50 filas).

El Excel **no** se envía al modelo.

## Contrato del puente `postMessage`

Sobre de petición (plugin → motor, `window.parent`):

```js
{ channel: "tm-excel-engine", action: "execute_sql", sql: "<SELECT …>", requestId: "req_…" }
```

Sobre de respuesta (motor → plugin, al `event.source`):

```js
{ channel: "tm-excel-engine", requestId: "req_…", success: true,  data: [ { col: valor }, … ] }
{ channel: "tm-excel-engine", requestId: "req_…", success: false, error: "mensaje" }
```

Notas de implementación:

- El `requestId` lo genera el plugin (`req_<timestamp>_<aleatorio>`) y el motor lo devuelve intacto:
  así se emparejan respuestas aunque haya varias consultas en vuelo.
- Cada llamada registra y **retira** su propio listener (`onMessage`), por lo que no se acumulan
  handlers entre invocaciones.
- La función **debe** llamarse `query_excel_data` en el nivel superior del código: es un requisito
  de TypingMind (el nombre de la función tiene que coincidir con `openaiSpec.name`). Poner un
  `return` suelto fuera de la función provoca un `SyntaxError` al guardar el plugin.
- Timeout de 30 s: si el motor no responde (extensión no cargada, sin datos), el plugin devuelve un
  mensaje de error legible para que el LLM lo comunique o lo reintente.

## Instalación

1. Instala `typingmind-excel-engine-v1.0.js` como Extension.
2. Recarga TypingMind; debe verse el panel del motor.
3. Crea un Plugin:
   - Implementation type: **JavaScript**
   - Output: **Give plugin output to the AI** (`respond_to_ai`)
   - Pega el contenido de `plugin.json` (spec) y `implementation.js` (código).
4. Activa el plugin en la conversación.
5. Carga un Excel en el panel del motor **antes** de preguntar.

## Límites

- Solo lectura (`SELECT` / `WITH` / `EXPLAIN`) contra la tabla `excel_data`: el plugin limpia
  comentarios y bloquea `DROP`, `ALTER`, `INSERT`, `UPDATE`, `DELETE`, `CREATE`, `COPY`, `ATTACH`,
  `EXPORT`, `PRAGMA FORCE`…
- El plugin recorta el resultado a **50 filas** y escapa `|` y saltos de línea para no romper el Markdown.
- Si el motor no está cargado o no hay Excel, el plugin devuelve un error claro al LLM.
- Requiere que la **Extensión** esté activa en la misma pestaña del chat y que se haya cargado un
  archivo en el widget (la tabla `excel_data` solo existe tras una carga).
- Es un plugin **sin estado**: no guarda historial entre consultas; cada llamada abre y cierra su
  propio listener.

## Verificación de la carpeta

Esta carpeta contiene exactamente los tres archivos necesarios y nada más:

```text
plugin/
├── plugin.json         spec importable (metadatos + openaiSpec)
├── implementation.js   código de la función query_excel_data
└── README.md           esta guía
```

- `plugin.json` valida como JSON estricto (`node -e "require('./plugin.json')"`).
- `implementation.js` declara `async function query_excel_data(params)` en el nivel superior y no
  contiene código muerto ni dependencias externas (solo `postMessage` y `setTimeout`).
