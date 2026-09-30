# System Patterns — TypingMind Excel Engine

> Arquitectura, decisiones técnicas y rutas críticas de implementación.
> Todo verificado leyendo el código (líneas citadas del commit `5067b16`).

## Arquitectura: 3 piezas, 2 protocolos, 1 navegador

```text
[1] EXTENSIÓN  typingmind-excel-engine-v1.0.js   (ventana principal de TypingMind)
      · fetch(WORKER_PATH) -> Blob -> new Worker()        (bypass de CORS)
      · Widget UI + consola SQL + CSV + Markdown
      · Puente postMessage del canal "tm-excel-engine"
              │  { type, requestId, ...payload }   ▲ { success, requestId, data|error }
              ▼                                     │
[2] WORKER     duckdb-worker.js                  (DuckDB-WASM 1.29.0 + SheetJS)
      · Blob Worker anidado para el binario de DuckDB
      · VFS virtual: registrar archivo -> CREATE OR REPLACE TABLE -> dropFile
              ▲  { channel, action:"execute_sql", sql, requestId }
              │  { channel, requestId, success, data|error }
[3] PLUGIN IA  plugin/implementation.js          (iframe sandbox de TypingMind)
      · query_excel_data({ sql_query }) -> window.parent.postMessage
      · Sin acceso a DuckDB: solo transporta texto
```

## Decisiones técnicas clave

### 1. Bypass de CORS con Blob URL (doble)
GitHub Pages y jsDelivr no permiten `new Worker()` cross‑origin.
- **Extensión:** `fetch(WORKER_PATH)` → `Blob` → `Worker(blobURL)`; la URL se revoca
  (`revokeObjectURL`) en cuanto `initDuckDB` termina (líneas 46‑77).
- **Worker:** DuckDB‑WASM necesita su propio `mainWorker`; se crea un **segundo Blob Worker** cuyo
  único contenido es `importScripts(bundle.mainWorker)` (worker, líneas 31‑43).

### 2. Petición/respuesta con `requestId` y `Map` de pendientes
`sendRequest(type, payload, timeoutMs, transferables)` (extensión, líneas 86‑97) genera ids
incrementales y guarda `{ resolve, reject, timer }` en `pending`. La respuesta llega por
`worker.onmessage` y se empareja por `requestId`. Timeouts: **120 s** en `initDuckDB` y cargas,
**60 s** en `executeQuery`.

### 3. Transferables para binarios
Parquet y Excel se envían al Worker como `ArrayBuffer` con lista de transferencia
(`sendRequest(..., 120000, [e.target.result])`, extensión líneas 436‑446) → sin copia.

### 4. Normalización de tipos en el Worker (frontera `postMessage`)
`normalizeValue` / `normalizeRows` (worker, líneas 61‑149):
- `BIGINT` → `Number` si está en rango seguro; si no, `string`.
- `HugeInt`/`Decimal` detectados por patrón de 4 elementos `[a,0,0,0]` / `[a,-1,-1,-1]` → escalar.
- `Date` → ISO 8601 (o `null` si inválida), `ArrayBuffer`/`DataView`/typed arrays → arrays, `null`
  preservado. Objetivo: evitar `TypeError: Do not know how to serialize a BigInt`.

### 5. Ciclo de vida del VFS (privacidad por diseño)
Patrón idéntico en `loadCSV`, `loadParquet`, `loadExcel`: registrar → `CREATE OR REPLACE TABLE` →
leer esquema/conteo/preview → **`dropFile` en el `finally`**. Los errores se guardan en `mainError`
para no enmascarar el fallo real con el de la limpieza. Las **tablas** sobreviven; los archivos no.

### 6. Excel multi‑hoja = "Variante A": una tabla por pestaña
- **Hoja 0** → tabla base recibida (`excel_data`), por retrocompatibilidad con el plugin y las suites.
- **Hojas 1..N** → `excel_data_<hoja saneada>` (`buildSheetTableName`, worker líneas 329‑359).
- Saneado: no alfanumérico → `_`, recorte de `_` extremos, prefijo `t_` si empieza por dígito,
  `hoja_<n>` si el nombre no aporta caracteres útiles.
- **Deduplicación case‑insensitive** (DuckDB compara identificadores sin distinguir mayúsculas):
  `_2`, `_3`… ante colisiones de saneado.
- Cada hoja se materializa como **CSV temporal en el VFS**
  (`upload_<requestId>_<índice>.csv` → `read_csv_auto(sample_size=-1, ignore_errors=true)`).
- **Fallo aislado por hoja:** `materializarHojaExcel` nunca lanza; devuelve
  `estado: "ok" | "omitida" | "error"` con `motivo` (`hoja_vacia`, `hoja_sin_datos`,
  `hoja_error_tabla`, `hoja_error_vfs`). Solo hay error global si el libro es ilegible, no tiene
  hojas o **ninguna** tiene datos. Si la hoja 0 está vacía, `excel_data` no existe y la tabla
  primaria pasa a ser la primera con datos.
- Los temporales se borran todos en el `finally`; una tabla parcial se descarta con
  `DROP TABLE IF EXISTS` best‑effort (`descartarTablaParcial`).

### 7. Guardrail de solo lectura en el plugin (lista negra textual)
`implementation.js` (líneas 22‑25): elimina comentarios `/* */` y `--`, y bloquea por regex
`\b(drop|alter|insert|update|delete|attach|copy|export|create|pragma force)\b`.
Es **textual**: también salta con esas palabras dentro de identificadores o literales
(`SELECT "Update" …`), y el spec instruye al modelo a construir esas palabras por concatenación
(`'DEL' || 'ETE'`). El bloqueo vive **solo en el plugin**; el widget no restringe al usuario.

### 8. Puente postMessage sin estado
Cada llamada del plugin registra su propio listener y lo **retira** al resolver (evita acumulación de
handlers). El motor responde a `event.source` o, si no hay, hace *broadcast* a todos los iframes.
Canal: `"tm-excel-engine"`.

### 9. Interacción de arrastre con discriminación clic vs. arrastre
Constantes (extensión líneas 20‑22): `UMBRAL_ARRASTRE_PX = 5`, `VENTANA_SUPRESION_MOUSE_MS = 400`,
`OPCIONES_TOUCH_MOVE = { passive: false }` (permite `preventDefault` y evitar scroll durante el arrastre).
Los `mousedown`/`mouseup` sintéticos que el navegador emite tras un toque se descartan por ventana
temporal; `touchend` se cancela para no duplicar la alternancia. Posición con `limitarDentroDeViewport`
(clamping) y reajuste en `resize` (más `setTimeout` de 320 ms para esperar la transición de 0.3 s).

## Rutas críticas de implementación

| Necesidad | Punto de entrada |
|---|---|
| Arranque (DOM listo) | extensión líneas 640‑644 → `injectWidget()` + `initWorker()` |
| Inyectar el widget | `injectWidget()` líneas 129‑230 (CSS id `tmee-widget-styles`, nodo `#tmee-widget-root`) |
| Eventos de UI | `setupEvents()` líneas 233‑256 + `setupDrag()` líneas 297‑303 |
| Cargar archivo | `procesarArchivo(file)` líneas 419‑448 (CSV texto / Parquet‑Excel binario) |
| Ejecutar SQL manual | `ejecutarSQL()` líneas 461‑477 |
| Pintar tabla | `renderizarTablaSQL(rows)` líneas 479‑515 (máx. 100 filas) |
| Enviar al chat | `inyectarEnChat()` líneas 566‑615 (`buscarInputChat` → `construirTablaMarkdown`) |
| Exportar CSV | `exportarCSV()` líneas 617‑638 (`tmee_resultados.csv`) |
| Escapado Markdown | `escaparCeldaMarkdown` líneas 523‑529 (`\` → `\\`, `|` → `\|`, saltos → `<br>`) |
| Puente IA | listener global líneas 100‑126 |
| Tipos de mensaje del Worker | `self.onmessage` worker líneas 561‑599 |

## Contrato de datos devuelto por las cargas

```js
{ procesamiento: "LOCAL_NATIVO_WORKER", formato: "CSV"|"PARQUET"|"EXCEL",
  tabla, registros, columnas[], esquema[], preview[] /* 10 filas */ }
// solo EXCEL añade:
{ hojas: [{ nombre, tabla, registros, columnas, estado, motivo?, error?, advertencia? }],
  hojas_totales, hojas_cargadas, hojas_omitidas, hojas_con_error }
```

Detalle del puente de la IA en `plugin-bridge-protocol.md`.

## Convenciones del código

- **Idioma:** comentarios, mensajes de estado y alertas en **español**; identificadores de API en
  inglés (`loadCSV`, `sendRequest`, `executeQuery`).
- **Prefijo DOM:** todo id/clase del widget usa `tmee-` (`#tmee-widget-root`, `.tmee-header`…).
- **Sin dependencias de estructura:** ambos archivos de producto son scripts autónomos
  (IIFE en la extensión; worker con `importScripts` + `import()` dinámico). No hay bundler.
- **SQL defensivo en el Worker:** identificadores siempre citados (`quoteIdentifier`) y literales
  escapados (`quoteStringLiteral`); el saneado (`sanitizeTableName`) cae a `excel_data`.

## Por confirmar

- El empaquetado/publicación no está automatizado: no hay `.github/`, ni `package.json`, ni CI.
  La única "publicación" es servir la rama `main` por GitHub Pages.
- `#tmee-dropzone` anuncia "`.csv, .parquet, .xlsx`" (línea 192) mientras el `accept` incluye `.xls`
  (línea 194): divergencia menor de copy, sin impacto funcional.
- `sanitizeTableName` (worker líneas 164‑170) **no** recorta `_` de los extremos, mientras
  `sanitizeSheetSuffix` (líneas 329‑339) sí: comportamiento intencional o descuido, por confirmar.

