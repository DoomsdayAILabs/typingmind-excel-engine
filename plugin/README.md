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
| `openaiSpec.description` | *(texto en español)* | Prompt de la herramienta: fija la hoja 0 como `excel_data`, define las tablas secundarias `excel_data_<hoja>`, ordena descubrirlas con `information_schema.tables` / `SHOW TABLES;`, limita a lectura y anuncia la respuesta como tabla Markdown. |
| `openaiSpec.parameters` | `sql_query` (string, **requerido**) | Contrato de entrada: la consulta SQL que genera el modelo. |

## Flujo

1. El usuario carga el Excel en el panel del motor (Extensión).
2. El usuario hace una pregunta en lenguaje natural.
3. El LLM llama `query_excel_data` con un `sql_query` de solo lectura (`SELECT` / `WITH` / `EXPLAIN` / `SHOW` / `DESCRIBE`).
4. El plugin (iframe) envía la consulta por `window.parent.postMessage`.
5. El motor ejecuta `executeQuery` en DuckDB y responde las filas.
6. El plugin devuelve a la IA un Markdown tabular (máx. 50 filas).

El Excel **no** se envía al modelo.

## Multi-hoja y descubrimiento de tablas

El modelo **no ve las pestañas del libro**: solo ve tablas dentro de DuckDB. Por eso el spec le
exige descubrirlas antes de consultarlas.

| Pestaña del libro | Tabla en DuckDB | Notas |
|---|---|---|
| hoja 0 (primera) | `excel_data` | nombre retrocompatible con el widget y las suites SQL |
| hojas 1..N | `excel_data_<hoja saneada>` | una tabla por pestaña, creada por `loadExcel` en el Worker |

Saneado del sufijo (`sanitizeSheetSuffix` + `buildSheetTableName` del Worker):

- no alfanuméricos → `_` y se recortan los `_` de los extremos (`Ventas 2024` → `Ventas_2024`,
  `Otra Hoja!` → `Otra_Hoja`);
- prefijo `t_` si el nombre empieza por dígito; `hoja_<n>` si no aporta caracteres útiles;
- sufijo `_2`, `_3`… cuando dos pestañas colisionan tras el saneado;
- las hojas vacías se **omiten** (no crean tabla) y el fallo de una pestaña no aborta el libro;
- **si la hoja 0 está vacía, `excel_data` no existe**: la tabla primaria pasa a ser la primera
  pestaña con datos, de ahí que el spec ordene listar tablas en lugar de asumir el nombre.

SQL que el spec enseña al modelo:

```sql
-- 1) ¿Qué tablas hay?
SELECT table_name FROM information_schema.tables WHERE table_schema = 'main' ORDER BY table_name;
SHOW TABLES;                                    -- alternativa corta (columna "name")

-- 2) ¿Qué columnas tiene cada tabla?
SELECT column_name, data_type
FROM information_schema.columns
WHERE lower(table_name) = lower('excel_data')
ORDER BY ordinal_position;
DESCRIBE excel_data_ventas_2024;                -- alternativa corta

-- 3) Cruzar pestañas
SELECT d.nombre, v.importe
FROM excel_data d JOIN excel_data_ventas_2024 v ON d.id = v.id
WHERE v.importe > 150;
```

`LIKE` distingue mayúsculas (los identificadores de DuckDB no), de ahí el `lower()` al filtrar por
`table_name`. Todas estas consultas son de solo lectura y **pasan el guardrail** del plugin
(verificado en runtime con `SHOW TABLES;`, `information_schema.tables/.columns`, `DESCRIBE` y un
`JOIN` entre pestañas).

## Guardrail de solo lectura (Fase 8)

`implementation.js` valida el SQL con **cuatro comprobaciones en cadena** antes de enviarlo al motor
(bloque de guardrail, líneas 35-103):

1. **Saneado en una pasada (`saneaSQL`)** — un único recorrido devuelve el texto con los **comentarios**
   (de bloque y de línea) eliminados y los **literales** (`'...'` y `"..."`) vaciados, conservando sus
   delimitadores para no alterar los límites de palabra. Consecuencia doble: comparar o citar una palabra
   reservada ya **no** es un falso positivo (`WHERE estado = 'DELETE'`, `SELECT "Update" FROM
   excel_data`) y un verbo destructivo oculto tras un comentario o un literal deja de escaparse. Si el
   literal queda **sin cerrar**, el escaneo no es fiable y la consulta se rechaza.
2. **Lista blanca (`LIMITADO`)** — `^(select|with|explain|show|describe)\b` sobre el texto saneado: la
   sentencia debe empezar por una de las cinco formas de lectura.
3. **Lista negra (`PROHIBIDO`)** — sobre el texto saneado bloquea escritura, DDL, DCL y administración del
   motor: `drop`, `alter`, `insert`, `update`, `delete`, `truncate`, `create`, `merge`, `attach`,
   `detach`, `copy`, `export`, `import`, `install`, `load`, `call`, `set`, `reset`, `vacuum`,
   `checkpoint`, `pragma`, `grant`, `revoke`, `use`, `begin`, `commit`, `rollback`, `transaction`,
   `prepare`, `execute`, `deallocate`.
4. **Sentencias apiladas** — si queda algún punto y coma fuera de los literales (más allá del separador
   final opcional), la consulta se rechaza: **una llamada = una sentencia**.

Decisiones asumidas: `PRAGMA` queda bloqueado (para el esquema, `DESCRIBE` o `information_schema`) y la
consulta debe empezar por palabra clave formal (no se admite `(SELECT ...)`). `replace()` y
`EXPLAIN ANALYZE` siguen permitidos (no son sentencias) y el mensaje de rechazo es siempre el documentado
en § Contrato del puente. El guardrail es textual; la consola SQL del **widget** ejecuta cualquier
sentencia a propósito, porque es el motor local del propio usuario.

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

- Solo lectura (`SELECT` / `WITH` / `EXPLAIN` / `SHOW` / `DESCRIBE`) contra las tablas locales
  (`excel_data` + `excel_data_<hoja>`): ver § Guardrail de solo lectura (Fase 8) para las cuatro
  comprobaciones en cadena (saneado de comentarios y literales, lista blanca `LIMITADO`, lista negra
  `PROHIBIDO` y rechazo de sentencias apiladas y de literales sin cerrar).
- El plugin recorta el resultado a **50 filas** y escapa `|` y saltos de línea para no romper el Markdown.
- Si el motor no está cargado o no hay Excel, el plugin devuelve un error claro al LLM.
- Requiere que la **Extensión** esté activa en la misma pestaña del chat y que se haya cargado un
  archivo en el widget (las tablas `excel_data*` solo existen tras una carga).
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
