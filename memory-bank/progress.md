# Progress — TypingMind Excel Engine

> Qué funciona, qué falta, estado actual e issues conocidos.
> Estado evaluado el **2026‑09‑30**; commits del producto **`d876db4`** (Fase 6 · UI desplegada) y
> **`be03349`** (Fase 8 · guardrail robusto del plugin: saneado + lista blanca + lista negra ampliada) más el
> commit de documentación que sincroniza este Memory Bank.

## Estado actual

**El soporte multi-hoja está 100 % cerrado en las tres piezas (Worker, Plugin IA y UI Widget) y las
dependencias externas quedan fijadas.** SheetJS pasa de etiqueta flotante a **`0.18.5`** (`686a46e`),
así que DuckDB‑WASM `1.29.0` y SheetJS `0.18.5` son las dos únicas versiones que carga el Worker. Las
tres suites headless están en verde (**85/85**, 44/44 y 25/25) y GitHub Pages sirve la build de Fase 6
(cabecera “Fase 6” en la extensión, que no cambió en Fase 7). El README (**527 líneas**) ya no describe
archivos inexistentes, declara la licencia MIT oficial y sus listas de sentencias permitidas y de
descubrimiento de esquema coinciden con el plugin real. El repositorio está **limpio y sincronizado**
con `origin/main`. **Fase 8 (`be03349`):** el guardrail del plugin pasa de lista negra textual a cuatro
comprobaciones en cadena (saneado, lista blanca, lista negra ampliada y bloqueo de apiladas), con
**0 falsos positivos y 0 falsos negativos** sobre un banco de 67 consultas.

## Qué funciona (verificado)

### Motor (worker)
- Inicialización de DuckDB‑WASM 1.29.0 con Blob Worker anidado (`initDuckDB` → versión
  `v1.0-phase-1b+multisheet`).
- **CSV:** `registerFileText` → `CREATE OR REPLACE TABLE` con `read_csv_auto(sample_size=-1,
  ignore_errors=true)` → conteo, esquema (`information_schema.columns`) y preview de 10 filas.
- **Parquet:** `registerFileBuffer` → `read_parquet` (llega por transferable).
- **Excel:** `XLSX.read` + `sheet_to_csv` por pestaña → **una tabla por hoja**, con SheetJS **fijado a
  `0.18.5`** (`SHEETJS_PACKAGE`, worker líneas 5 y 7) desde `686a46e`.
- **Multi‑hoja: 100 % (Worker).** Una tabla por pestaña, aislamiento de fallos por hoja y detalle
  completo en `hojas[]` (`hojas_totales`, `hojas_cargadas`, `hojas_omitidas`, `hojas_con_error`).
- **Normalización de tipos** en la frontera `postMessage` (BIGINT fuera de rango → string,
  HugeInt/Decimal, fechas ISO, binarios a arrays).
- **Limpieza del VFS** (`dropFile`) en el `finally` de las tres rutas de carga.

### Extensión (UI + puente)
- Widget inyectado con estado, dropzone, consola SQL, tabla (100 filas) y metadatos.
- Minimizar/maximizar con clic o tap en la cabecera; modo icono flotante de 50×50 px.
- Arrastre con ratón y táctil, umbral de 5 px, supresión del ratón sintético y clamping al viewport.
- **⬇️ CSV** → `tmee_resultados.csv` (UTF‑8, comillas escapadas).
- **💬 Enviar a TM** → tabla Markdown (máx. 50 filas) en el input del chat + evento `input` para
  React + auto‑minimizado; fallback a portapapeles.
- Puente `postMessage` del canal `tm-excel-engine` operativo (incluye *broadcast* a iframes si no hay
  `event.source`).
- **Multi‑hoja: 100 % (UI Widget).** Barra compacta `#tmee-sheets-bar` entre los metadatos y la consola
  SQL, con **desplegable nativo** `#tmee-sheets-select` (una opción por pestaña del libro).
- **Aislamiento condicional:** la barra solo se muestra si el libro tiene **más de una hoja con datos**
  (`cargadas > 1`); con CSV, Parquet o Excel de una sola hoja el widget es idéntico al de Fase 5.
- **Hojas no seleccionables:** las omitidas o con error se listan **deshabilitadas** y etiquetadas
  (`omitida (vacía)`, `omitida (sin datos)`) y las de solo cabeceras se marcan como tales; el resumen
  (`Hojas 3/5 · 2 vacías`) avisa en ámbar y el tooltip detalla el desenlace de cada pestaña.
- **Cambio de hoja sin recargar:** reapunta la consola a `excel_data_<hoja>` y reconsulta mediante
  `executeQuery` (una única petición nueva; el archivo **no** se vuelve a subir ni se reprocesa el VFS).
- **Sustitución conservadora del SQL:** la plantilla se reescribe solo si el usuario no la editó; si la
  editó, se sustituye **únicamente** el token `FROM <tabla>` y su consulta se conserva intacta.
- **Metadato de columnas** (`#tmee-meta-cols`) junto a tabla y registros de la hoja activa.

### Plugin IA
- `query_excel_data` declarada en el nivel superior (requisito de TypingMind), `plugin.json` válido.
- Guardrail de solo lectura **Fase 8** (bloque de guardrail, líneas 35-103): saneado de comentarios y
  literales en una pasada (`saneaSQL`) + lista blanca (`SELECT`, `WITH`, `EXPLAIN`, `SHOW`, `DESCRIBE`) +
  lista negra ampliada + rechazo de sentencias apiladas y de literales sin cerrar; **0 FP / 0 FN** sobre el
  banco de 67 consultas. Timeout de 30 s, listener por llamada y respuesta Markdown ≤ 50 filas (contrato
  verificado 22/22).
- Spec multi‑hoja con descubrimiento de tablas (`information_schema.tables` / `SHOW TABLES`,
  `DESCRIBE` / `information_schema.columns`); ya **no** advierte sobre palabras reservadas dentro de
  literales ni pide concatenar con `||` (el saneado lo hace innecesario).
- **Multi‑hoja: 100 % (Plugin IA).** El spec descubre todas las tablas del libro y trabaja
  indistintamente con `excel_data` y `excel_data_<hoja>` sin cambios adicionales.

### Verificaciones ejecutadas hoy (evidencia)

| Suite | Checks | Resultado |
|---|---|---|
| `tests/fase6-multisheet-verificacion.html` | 85 | **OK 85 / FAIL 0** |
| `tests/fase5-verificacion.html` | 44 | **OK 44 / FAIL 0** |
| `tests/fase3b-verificacion.html` | 25 | **OK 25 / FAIL 0** |
| `node --check` (worker, plugin, extensión) | 3 | **OK (exit 0 ×3)** |
| jsDelivr: `xlsx@0.18.5` vs. URL sin pin (`HEAD`) | 1 | **881 727 bytes idénticos** |
| BOM/CRLF/líneas (README, worker, implementation) | 3 | **sin BOM añadido, 0 LF sueltos, sin doble codificación** |
| Banco del guardrail (67 consultas sobre el plugin real) | 67 | **0 FP / 0 FN** |
| Contrato del plugin (canal, 30 s, Markdown, 50 filas) | 22 | **22/22 OK** |

## Deriva de documentación (corregida)

`README.md` **ya está alineado con el repositorio**. En el commit de sincronización del Memory Bank se
eliminaron del árbol las líneas de `index.html`, `.nojekyll`,
`backup/duckdb-worker.v0.4.23-contaminado.js` y `duckdb/typingmind-excel-engine-v0.3-test.js`; se
quitaron las menciones a `.nojekyll` en la guía de instalación y a las carpetas legado en los hitos; y
la sección § Licencia declara ya la **MIT oficial** (`LICENSE`). Se conserva el registro de lo
detectado como evidencia del proceso:

`README.md` describía un repositorio que ya no existía en varios puntos. Verificado con
`Get-ChildItem -Force` y búsquedas en el archivo:

| Línea del README | Afirma | Realidad verificada |
|---|---|---|
| 414 | existe `index.html` (historial de versiones) | **No existe** (eliminado en la limpieza de `test-data`/`index.html`) |
| 415, 249 | existe `.nojekyll` "ya incluido" | **No existe** en el repo |
| 426‑427 | existen `backup/duckdb-worker.v0.4.23-contaminado.js` y `duckdb/typingmind-excel-engine-v0.3-test.js` | **Ninguno existe** |
| 477 | hitos: "ver carpetas `duckdb/` y `backup/`, no usar" | Las carpetas **no existen** |
| 498‑500 | "el repositorio **aún no incluye un archivo `LICENSE`**" | **`LICENSE` MIT © 2026 DoomsdayAILabs sí existe** |
| 94‑95 | el plugin permite solo `SELECT`, `WITH` y `EXPLAIN` | **corregido en `686a46e`**: el README (líneas 379 y 488) y `plugin/README.md` ya documentan `SHOW` y `DESCRIBE` junto a `SELECT`/`WITH`/`EXPLAIN` (la lista negra no los bloquea) |
| 101‑102 y 336‑337 | el modelo empieza con `SELECT * FROM excel_data LIMIT 1;` para descubrir columnas | **corregido en `686a46e`**: el README ya ordena `information_schema.tables` / `SHOW TABLES` y después `information_schema.columns` / `DESCRIBE`, igual que el spec |

Deriva adicional fuera del README: `tests/sql-test-runner.html` mantiene el título
**"Batería de pruebas SQL — v0.4.20"** (línea 30) aunque el producto es v1.0.

**Estado de la corrección:** las **siete filas de la tabla están corregidas**. Las dos últimas se
cerraron en `686a46e` (lista de sentencias permitidas y orden de descubrimiento del esquema). La única
deriva documental que queda es el título desactualizado de `sql-test-runner.html`.

## Qué falta / pendientes reales

1. **`tests/sql-test-runner.html`:** necesita que el usuario **seleccione a mano** `TEUs.xlsx` o
   `RequerimientoPrueba`; esos datos de ejemplo ya no están en el repo, así que la batería SQL más
   amplia (**38 casos**: 20 + 18) **no es automatizable hoy**. Ver `testing-strategy.md`.
2. **Versión del producto:** el Worker declara `v1.0-phase-1b+multisheet` mientras el producto es
   `1.0`; no hay un único identificador de versión.
3. **Título de `tests/sql-test-runner.html`:** anuncia "Batería de pruebas SQL — v0.4.20" con el
   producto en v1.0.
4. ~~**SheetJS sin pin de versión** en `duckdb-worker.js` (línea 6).~~ **Completado en `686a46e`:**
   fijado a `0.18.5` con `SHEETJS_PACKAGE` (worker, línea 5) e `importScripts(SHEETJS_PACKAGE)` (línea
   7). Queda como riesgo residual la ausencia de SRI (ver `techContext.md`).

## Issues y riesgos conocidos

- **Tablas huérfanas entre cargas:** cada carga recrea la tabla base pero **no elimina** las
  `excel_data_<hoja>` de un libro anterior con más pestañas (documentado en README líneas 379‑385).
  El **selector de hojas no las muestra** (se construye desde `result.hojas` del libro vigente), pero
  siguen siendo consultables con `SHOW TABLES` y desde la consola SQL.
- ~~**Guardrail = lista negra textual.**~~ **Resuelto en `be03349` (Fase 8).** El guardrail es ahora lista
  blanca (`SELECT`, `WITH`, `EXPLAIN`, `SHOW`, `DESCRIBE`) + saneado de comentarios y literales + lista
  negra ampliada, que sí bloquea `TRUNCATE`, `INSTALL`, `LOAD`, `SET`, `CALL`, `VACUUM`, `CHECKPOINT`,
  `PRAGMA`, `MERGE`, `USE`, `BEGIN`/`COMMIT`/`ROLLBACK`… además de las que ya bloqueaba.
  *Nota:* el widget (uso manual) sigue ejecutando cualquier sentencia a propósito: la restricción solo
  aplica a lo que la IA puede lanzar por sí misma.
- ~~**Falsos positivos del guardrail.**~~ **Resuelto en `be03349` (Fase 8):** los literales
  (`WHERE estado = 'DELETE'`) y los identificadores entrecomillados (`SELECT "Update" FROM …`) ya **no** se
  bloquean, porque `saneaSQL` vacía los literales antes de aplicar la lista negra. El spec del plugin ya no
  pide al modelo evitar esas palabras ni concatenar con `||`.
- **Riesgo residual del guardrail textual (Fase 8):** es heurístico, no un parser SQL. Los dos modos de
  ocultar un verbo destructivo (tras un comentario o tras un literal con comilla dentro de `$$...$$`) se
  cierran con el saneado en una pasada y el rechazo de literales sin cerrar; el endurecimiento estructural
  (conexión DuckDB restringida para el plugin) queda como línea futura.
- **Sin persistencia:** recargar la página descarta datos y motor.
- **`file://` no funciona** para pruebas: hace falta servidor local.
- **Dependencia de jsDelivr** en el primer arranque (CDN bloqueado ⇒ motor inoperante).
- **`DROP TABLE IF EXISTS` best‑effort** al descartar una tabla parcial: si falla, queda una tabla
  parcial y la hoja se reporta igualmente como error (comportamiento intencional).

## Evolución de decisiones (desde `git log`)

| Hito | Commit |
|---|---|
| Motor DuckDB‑WASM en Worker + proxy Blob (Fase 1B) | `c7a79d2` |
| Parquet (1C) y Excel con SheetJS (1D) | `b7eaa49`, `c4278a6` |
| `read_csv_auto` con escaneo completo | `28489a3` |
| Puente IA + botón CSV | `17eb201` |
| Botones toggle e inyección al chat (Fase 3B) | `e6ce0fc`, `435d940` |
| UI móvil: FAB + arrastre (Fase 5) | `43960c3` |
| Documentación y empaquetado v1.0 (Fase 6) + limpieza de legado | `e3a162c`, `54b9ed3`, `41f97bc` |
| `LICENSE` MIT añadido | `57d0667` |
| **Multi‑hoja en el Worker** (aislamiento por hoja, Variante A) | `142478e` |
| Spec multi‑hoja con descubrimiento de tablas | `e174d79` |
| README sincronizado con `SHOW`/`DESCRIBE` y sin referencias a `test-data` | `38521f4`, `5067b16` |
| **Memory Bank inicializado** (6 archivos núcleo + 2 de contexto) y `.clinerules/` versionados | `425e5fd` |
| **Fase 6 · UI**: selector de hojas multi-hoja en el widget + suite de 85 checks | `d876db4` |
| Memory Bank sincronizado con Fase 6 y README purgado (deriva + licencia MIT) | commit de docs posterior a `d876db4` |
| **Fase 7 · deps/docs**: pin de SheetJS `0.18.5` + guardrail y README con `SHOW`/`DESCRIBE` | `686a46e` |
| Memory Bank sincronizado con Fase 7 (dependencias fijadas y referencias de README) | commit de docs posterior a `686a46e` |
| **Fase 8 · plugin**: guardrail robusto (saneado + lista blanca + lista negra ampliada + apiladas) | `be03349` |
| Memory Bank sincronizado con Fase 8 (guardrail, spec y docs) | commit de docs posterior a `be03349` |

## Por confirmar

- Destino de los datos de ejemplo eliminados (`test-data`): ¿existen fuera del repo para uso manual?
- Si se quiere `index.html` de vuelta como historial de versiones descargables (ya no lo anuncia el
  README, así que ahora es una decisión de producto).
- Estado de GitHub Pages como Pages "clásico" vs. Actions: no hay workflow en el repo (la propagación
  de `d876db4` tardó unos minutos, coherente con una caché de build).
- **Caso borde del selector:** libro con 1 hoja cargada y 1 vacía → la barra no se muestra y la omisión
  pasa desapercibida (un cambio de una línea, pendiente de decisión).

