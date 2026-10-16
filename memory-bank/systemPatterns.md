# System Patterns — TypingMind Excel Engine

> Arquitectura, decisiones técnicas y rutas críticas de implementación.
> Todo verificado leyendo el código. Los números de línea de la extensión corresponden a la versión con
> **Fase 6 · UI** (`d876db4`, 849 líneas) y los del worker a `duckdb-worker.js` (**661 líneas**,
> `v1.0-phase-9+tables-lifecycle`, tras `88c99e6`); la sincronización de este documento se hace en el
> commit de docs posterior.

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
  (`revokeObjectURL`) en cuanto `initDuckDB` termina (líneas 51‑93).
- **Worker:** DuckDB‑WASM necesita su propio `mainWorker`; se crea un **segundo Blob Worker** cuyo
  único contenido es `importScripts(bundle.mainWorker)` (worker, líneas 37-49).

### 2. Petición/respuesta con `requestId` y `Map` de pendientes
`sendRequest(type, payload, timeoutMs, transferables)` (extensión, líneas 95‑106) genera ids
incrementales y guarda `{ resolve, reject, timer }` en `pending`. La respuesta llega por
`worker.onmessage` y se empareja por `requestId`. Timeouts: **120 s** en `initDuckDB` y cargas,
**60 s** en `executeQuery`.

### 3. Transferables para binarios
Parquet y Excel se envían al Worker como `ArrayBuffer` con lista de transferencia
(`sendRequest(..., 120000, [e.target.result])`, extensión líneas 466‑477) → sin copia.

### 4. Normalización de tipos en el Worker (frontera `postMessage`)
`normalizeValue` / `normalizeRows` (worker, líneas 67-155):
- `BIGINT` → `Number` si está en rango seguro; si no, `string`.
- `HugeInt`/`Decimal` detectados por patrón de 4 elementos `[a,0,0,0]` / `[a,-1,-1,-1]` → escalar.
- `Date` → ISO 8601 (o `null` si inválida), `ArrayBuffer`/`DataView`/typed arrays → arrays, `null`
  preservado. Objetivo: evitar `TypeError: Do not know how to serialize a BigInt`.

### 5. Ciclo de vida del VFS (privacidad por diseño)
Patrón idéntico en `loadCSV`, `loadParquet`, `loadExcel`: registrar → `CREATE OR REPLACE TABLE` →
leer esquema/conteo/preview → **`dropFile` en el `finally`**. Los errores se guardan en `mainError`
para no enmascarar el fallo real con el de la limpieza. Las **tablas** sobreviven a la carga (las
obsoletas se purgan en la carga siguiente, ver § 7); los archivos no.

### 6. Excel multi‑hoja = "Variante A": una tabla por pestaña
- **Hoja 0** → tabla base recibida (`excel_data`), por retrocompatibilidad con el plugin y las suites.
- **Hojas 1..N** → `excel_data_<hoja saneada>` (`buildSheetTableName`, worker líneas 399-412).
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

### 7. Ciclo de vida de las tablas gestionadas (purga entre cargas · Fase 9)
El Worker mantiene un **`Map tablasGestionadas`** (clave: `nombre.toLowerCase()`; valor: el nombre exacto
con el que se creó la tabla) que solo contiene las tablas creadas por una carga:

- `recordarTablaGestionada(tableName)` (worker líneas 196-199) se invoca tras cada `CREATE OR REPLACE`
  que ha tenido éxito: en `loadCSV`, en `loadParquet` y, hoja a hoja, en `loadExcel` (solo si
  `carga.estado === "ok"`).
- `limpiarTablasObsoletas(tablasVigentes)` (líneas 209-220) recorre `tablasGestionadas.entries()` y, si
  la clave no figura entre las tablas creadas por el libro vigente, ejecuta
  `DROP TABLE IF EXISTS <nombreOriginal>` y borra la clave. Es **best-effort**: si el `DROP` falla, la
  entrada permanece en el Map y se reintenta en la carga siguiente (la carga nunca se rompe por esto).
- Se invoca **solo tras una carga exitosa** (justo después de `if (mainError) throw mainError`), de modo
  que un archivo ilegible o un libro sin datos conserva intacto el dataset anterior.
- La clave en minúsculas deduplica (DuckDB compara identificadores sin distinguir mayúsculas) y el
  **valor con la grafía original** es lo que se cita en el `DROP`, así el SQL emitido es fiel a la tabla
  creada.
- **Tablas del usuario:** la consola SQL ejecuta cualquier sentencia vía `executeQuery`, que no registra
  nada; por tanto las tablas creadas a mano nunca entran en el Map y **jamás** se eliminan. Si el usuario
  elige un nombre que colisiona con una tabla gestionada, la siguiente carga lo reemplaza
  (`CREATE OR REPLACE`, comportamiento preexistente).
- Caso cubierto de raíz: **hoja 0 vacía** → el libro nuevo no crea `excel_data`, así que la tabla
  `excel_data` del archivo anterior (con datos obsoletos) se purga en esa misma carga en lugar de quedar
  como tabla fantasma consultable por el modelo.

### 8. Guardrail de solo lectura en el plugin (lista blanca + saneado)
`implementation.js` (bloque de guardrail, líneas 35-103) valida en cadena antes de enviar nada al motor:

1. **Saneado en una pasada (`saneaSQL`)** — un solo recorrido devuelve el SQL con los comentarios (de
   bloque y de línea) sustituidos por un espacio y los **literales vaciados** (`'x'` → `''`, `"x"` →
   `""`), conservando los delimitadores para no alterar los límites de palabra; acumula con `slice`
   (fragmentos, no carácter a carácter). Devuelve también `literalAbierto` para detectar comillas sin
   cerrar.
2. **Lista blanca (`LIMITADO`)** — `^(select|with|explain|show|describe)\b` sobre el texto saneado.
3. **Lista negra (`PROHIBIDO`)** — sobre el texto saneado: escritura, DDL, DCL y administración
   (`drop`, `alter`, `insert`, `update`, `delete`, `truncate`, `create`, `merge`, `attach`, `detach`,
   `copy`, `export`, `import`, `install`, `load`, `call`, `set`, `reset`, `vacuum`, `checkpoint`,
   `pragma`, `grant`, `revoke`, `use`, `begin`, `commit`, `rollback`, `transaction`, `prepare`,
   `execute`, `deallocate`).
4. **Apiladas y literales sin cerrar** — un `;` fuera de literales (más allá del separador final
   opcional) o `literalAbierto` ⇒ rechazo. Una llamada = una sentencia.

Efecto medido (banco de 67 consultas contra el archivo real): **0 falsos positivos** —los literales y los
identificadores entrecomillados ya no disparan la lista negra— y **0 falsos negativos** —se cierran
`TRUNCATE`, `INSTALL`, `LOAD`, `SET`, `CALL`, `VACUUM`, `CHECKPOINT`, `PRAGMA force_checkpoint`, `MERGE`,
`USE`, `BEGIN`, las sentencias apiladas y los dos modos de ocultar el verbo tras un comentario o un
literal—. Todo el guardrail vive **solo en el plugin**; el widget no restringe al usuario.

### 9. Puente postMessage sin estado
Cada llamada del plugin registra su propio listener y lo **retira** en los tres desenlaces posibles
(Fase 9): al emparejar la respuesta, al vencer el **timeout de 30 s** y si `window.parent.postMessage`
lanza (en ese caso también se cancela el temporizador con `clearTimeout`). Así no se acumulan handlers
entre invocaciones ni sobrevive ningún closure tras un timeout. El motor responde a `event.source` o, si
no hay, hace *broadcast* a todos los iframes.
Canal: `"tm-excel-engine"`.

### 9. Interacción de arrastre con discriminación clic vs. arrastre
Constantes (extensión líneas 20‑22): `UMBRAL_ARRASTRE_PX = 5`, `VENTANA_SUPRESION_MOUSE_MS = 400`,
`OPCIONES_TOUCH_MOVE = { passive: false }` (permite `preventDefault` y evitar scroll durante el arrastre).
Los `mousedown`/`mouseup` sintéticos que el navegador emite tras un toque se descartan por ventana
temporal; `touchend` se cancela para no duplicar la alternancia. Posición con `limitarDentroDeViewport`
(clamping) y reajuste en `resize` (más `setTimeout` de 320 ms para esperar la transición de 0.3 s).

### 10. Selector de hojas en la UI: estado mínimo + aislamiento condicional (Fase 6)
Tres variables de módulo sostienen todo el comportamiento multi-hoja de la extensión:
- `libroActual` — el `result` de la última carga con `hojas[]` (`null` si la carga no es un Excel
  multi-hoja, p. ej. CSV o Parquet).
- `hojaActiva` — nombre de tabla visible en este momento (`excel_data` o `excel_data_<hoja>`).
- `plantillaSqlActual` — última plantilla autogenerada, que permite detectar si el usuario editó la
  consola SQL.

**Aislamiento condicional (cambio no invasivo).** `renderizarSelectorHojas()` solo puebla la barra si
el libro tiene **más de una hoja con `estado: "ok"`**; en cualquier otro caso llama a
`ocultarSelectorHojas()` y el widget queda exactamente como en Fase 5. La barra vive **dentro de
`.tmee-body`**, de modo que el modo icono —que oculta el cuerpo con `display: none !important`— la
esconde sin tocar el FAB de 50×50 px.

**Sustitución conservadora del SQL** (`aplicarPlantillaSql(tablaNueva, tablaAnterior)`):
1. Si es una carga nueva (`tablaAnterior === null`) o la consola sigue siendo idéntica a
   `plantillaSqlActual`, se reescribe la plantilla completa (`SELECT * FROM <tabla> LIMIT 10;`).
2. Si el usuario la editó, se sustituye **solo** el token del `FROM` mediante
   `(FROM\s+)<tablaAnterior>(?=$|[^A-Za-z0-9_])`. El *lookahead* es imprescindible: sin él, venir de
   `excel_data` corrompería `excel_data_Ventas`. Si la tabla va entrecomillada no hay coincidencia y el
   texto del usuario se respeta literalmente.

**Cambio de hoja sin recargar.** `seleccionarHoja(tabla)` valida la tabla contra `libroActual.hojas`
(solo `estado: "ok"`), actualiza los metadatos, reapunta la consola y reutiliza `ejecutarSQL()`: no hay
mensaje nuevo al Worker ni re-procesado del VFS, porque cada pestaña ya es una tabla viva.
`resetearSelectorHojas()` se invoca al arrancar `procesarArchivo` para que un libro nuevo —o una carga
fallida— no deje hojas ni plantillas del anterior.

## Rutas críticas de implementación

| Necesidad | Punto de entrada |
|---|---|
| Arranque (DOM listo) | extensión líneas 844‑848 → `injectWidget()` + `initWorker()` |
| Inyectar el widget | `injectWidget()` líneas 138‑255 (CSS id `tmee-widget-styles` en 144‑196; nodo `#tmee-widget-root` en 199‑251) |
| Eventos de UI | `setupEvents()` líneas 258‑284 + `setupDrag()` líneas 325‑331 |
| Cargar archivo | `procesarArchivo(file)` líneas 447‑478 (CSV texto / Parquet‑Excel binario + reset del selector) |
| Renderizar una carga | `renderizarResultado(result)` líneas 485‑498 + `actualizarMetaHoja()` 501‑516 |
| Selector de hojas (Fase 6) | `renderizarSelectorHojas()` 525‑565 · `seleccionarHoja()` 617‑629 · `aplicarPlantillaSql()` 639‑657 · `resetearSelectorHojas()` 605‑610 |
| Ejecutar SQL manual | `ejecutarSQL()` líneas 665‑681 |
| Pintar tabla | `renderizarTablaSQL(rows)` líneas 683‑719 (máx. 100 filas) |
| Enviar al chat | `inyectarEnChat()` líneas 770‑819 (`buscarInputChat` → `construirTablaMarkdown`) |
| Exportar CSV | `exportarCSV()` líneas 821‑842 (`tmee_resultados.csv`) |
| Escapado Markdown | `escaparCeldaMarkdown` líneas 727‑733 (`\` → `\\`, `|` → `\|`, saltos → `<br>`) |
| Puente IA | listener global líneas 109‑135 |
| Tipos de mensaje del Worker | `self.onmessage` worker líneas 622-661 |

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
- `#tmee-dropzone` anuncia "`.csv, .parquet, .xlsx`" (línea 211) mientras el `accept` incluye `.xls`
  (línea 213): divergencia menor de copy, sin impacto funcional.
- `sanitizeTableName` (worker líneas 171-177) **no** recorta `_` de los extremos, mientras
  `sanitizeSheetSuffix` (líneas 382-392) sí: comportamiento intencional o descuido, por confirmar.

