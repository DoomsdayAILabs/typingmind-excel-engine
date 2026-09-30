# Plugin Bridge Protocol — TypingMind Excel Engine

> Especificación de integración (puente IA ↔ motor) derivada de `plugin/implementation.js`,
> `plugin/plugin.json`, `plugin/README.md` y el listener de la extensión (líneas 100‑126).

## Piezas y superficies de instalación

| Archivo | Qué es | Dónde se instala |
|---|---|---|
| `typingmind-excel-engine-v1.0.js` | Motor + UI + puente `postMessage` | TypingMind → Settings → **Extensions** |
| `plugin/plugin.json` | Metadatos + especificación de la herramienta | TypingMind → **Plugins** → Create / Import |
| `plugin/implementation.js` | Código del plugin (se ejecuta en iframe sandbox) | Campo JavaScript del plugin |

No hace falta script puente aparte: el listener vive en el motor (ventana principal).

## Restricción de TypingMind que condiciona el diseño

El código del plugin **debe declarar en el nivel superior una función con exactamente el mismo nombre
que `openaiSpec.name`** (`query_excel_data`). Un `return` suelto en el nivel superior es un
`SyntaxError` al guardar. Por eso el cuerpo va envuelto en `async function query_excel_data(params)`.

## Sobre de petición (plugin → motor)

```js
window.parent.postMessage(
  { channel: "tm-excel-engine", action: "execute_sql", sql: "<SELECT …>", requestId: "req_…" },
  "*"
);
```

- `requestId` lo genera el plugin: `"req_" + Date.now() + "_" + aleatorio(0..999)`.
- Se emite desde `window.parent` porque el plugin corre en un iframe sandbox **sin acceso a DuckDB**.

## Sobre de respuesta (motor → plugin)

```js
// éxito
{ channel: "tm-excel-engine", requestId, success: true,  data: [ { col: valor }, … ] }
// error
{ channel: "tm-excel-engine", requestId, success: false, error: "mensaje" }
```

Comportamiento del motor (extensión, líneas 100‑126): verifica `channel` + `action`, ejecuta
`executeQuery` vía `sendRequest(..., 60000)`, guarda `ultimosResultados`, repinta la tabla del widget,
pone el estado en **"Consulta IA OK"** y responde a `event.source`; si no hay `source`, hace
**broadcast a todos los `iframe`** de la página. En error: estado **"Error SQL IA"**, limpia
`ultimosResultados` y devuelve `success: false` con el mensaje.

## Contrato de la función del plugin

```js
query_excel_data({ sql_query: string })   // sql_query REQUERIDO
```

- Acepta por compatibilidad `params.sql` o `params.query` como alias.
- **Timeout: 30 s.** Si no hay respuesta: "Error: Timeout. El motor local no respondió…".
- Resuelve **siempre como texto** (`outputType: respond_to_ai`); nunca lanza al modelo.
- Salidas posibles: tabla Markdown, "Consulta ejecutada con éxito (0 resultados).",
  "Error SQL del motor local: …", "Error: Solo se permiten consultas de lectura…",
  "Error interno: Datos no procesables.", "Error: Consulta SQL no válida o vacía…".

## Formato de la respuesta al modelo

```
Resultado (<n> de <total> filas):

| col1 | col2 |
| --- | --- |
| v1 | v2 |
```

- Máximo **50 filas** (`Math.min(rows.length, 50)`).
- Las columnas se toman de la **primera fila** devuelta (`Object.keys(displayRows[0])`).
- Se escapan `|` (→ `\|`) y los saltos de línea (→ espacio) para no romper el Markdown.
- `null`/`undefined` → celda vacía.

## Guardrail de solo lectura

1. Se eliminan comentarios `/* … */` y `-- …` antes de evaluar.
2. Se bloquea por regex con límites de palabra:
   `drop|alter|insert|update|delete|attach|copy|export|create|pragma force`.
3. **Es textual**: también salta con esas palabras dentro de identificadores o literales.

Permitidas de facto: `SELECT`, `WITH`, `EXPLAIN`, `SHOW`, `DESCRIBE` (y cualquier sentencia que no
contenga las palabras bloqueadas — ver el riesgo documentado en `progress.md`).

El spec instruye al modelo a construir las palabras prohibidas por concatenación (`'DEL' || 'ETE'`)
si necesita comparar con esos valores.

## Tablas que el modelo debe asumir

| Pestaña del libro | Tabla en DuckDB |
|---|---|
| hoja 0 (primera) | `excel_data` (nombre retrocompatible) |
| hojas 1..N | `excel_data_<hoja saneada>` |

Saneado (`sanitizeSheetSuffix` + `buildSheetTableName` del Worker): no alfanuméricos → `_`, recorte de
`_` extremos, prefijo `t_` si empieza por dígito, `hoja_<n>` si el nombre no aporta caracteres
útiles, y sufijos `_2`, `_3`… ante colisiones de saneado (case‑insensitive).

**Si la hoja 0 está vacía, `excel_data` NO existe** (la tabla primaria pasa a ser la primera pestaña
con datos). Por eso el spec ordena **listar tablas antes de consultar** en lugar de asumir el nombre.

SQL que el spec enseña al modelo:

```sql
-- 1) ¿Qué tablas hay?
SELECT table_name FROM information_schema.tables WHERE table_schema = 'main' ORDER BY table_name;
SHOW TABLES;                                    -- alternativa corta

-- 2) ¿Qué columnas tiene cada tabla?
SELECT column_name, data_type FROM information_schema.columns
WHERE lower(table_name) = lower('excel_data') ORDER BY ordinal_position;
DESCRIBE excel_data_ventas_2024;                -- alternativa corta

-- 3) Cruzar pestañas
SELECT d.nombre, v.importe FROM excel_data d
JOIN excel_data_ventas_2024 v ON d.id = v.id WHERE v.importe > 150;
```

`LIKE` distingue mayúsculas (los identificadores de DuckDB no) → de ahí el `lower()` al filtrar por
`table_name`.

## Campos de `plugin.json`

| Campo | Valor |
|---|---|
| `uuid` | `7f3c9e21-4b8a-4d6e-9f11-2c8a1b0e4d77` (**no cambiar** si se reimporta: evita duplicados) |
| `id` / `openaiSpec.name` | `query_excel_data` |
| `emoji` / `title` | `📊` / `Excel Data Engine (consulta local)` |
| `version` | `1` (entero) |
| `authenticationType` | `AUTH_TYPE_NONE` (sin credenciales) |
| `implementationType` | `javascript` |
| `outputType` | `respond_to_ai` |
| `userSettings` | `[]` |
| `openaiSpec.parameters` | `sql_query` (string, requerido) |

`plugin.json` es **JSON puro**: no admite comentarios; su documentación vive en `plugin/README.md`.

## Detalles de implementación obligatorios

- Cada llamada registra su **propio** listener y lo retira al resolver → no se acumulan handlers.
- El plugin es **sin estado**: no guarda historial entre consultas.
- Requiere que la Extensión esté activa en la misma pestaña y que haya un archivo cargado
  (las tablas `excel_data*` solo existen tras una carga).
- El Excel **no** se envía al modelo; solo viajan las filas resultantes.
- **El plugin no se despliega por GitHub Pages**: se pega a mano en TypingMind; al cambiar
  `plugin.json`/`implementation.js` hay que **volver a pegarlos** en la interfaz del plugin.

## Por confirmar

- Si el endurecimiento del guardrail (bloquear también `TRUNCATE`, `INSTALL`, `LOAD`, `SET`,
  `VACUUM`, `CHECKPOINT`…) es deseable o si la lista actual es una decisión asumida.
- Si la integración real de TypingMind acepta el *broadcast* a iframes como camino alternativo
  (código de la extensión, líneas 113‑117): no hay evidencia de prueba en el repo de ese camino,
  solo del `event.source` (la suite de Fase 3B emite eventos en la propia ventana).

## Instalación (orden correcto)

1. Instalar `typingmind-excel-engine-v1.0.js` como **Extension** y recargar TypingMind.
2. Crear un **Plugin**: Implementation type *JavaScript*, Output *Give plugin output to the AI*
   (`respond_to_ai`), pegar el spec de `plugin.json` y el código de `implementation.js`.
3. **Activar el plugin en la conversación.**
4. Cargar un archivo en el widget **antes** de preguntar.

