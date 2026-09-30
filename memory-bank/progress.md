# Progress — TypingMind Excel Engine

> Qué funciona, qué falta, estado actual e issues conocidos.
> Estado evaluado el **2026‑09‑30**; commit de referencia **`425e5fd`** (el código de producto no ha
> cambiado desde `5067b16`).

## Estado actual

**v1.0 cerrado y desplegado.** Las dos suites headless automatizadas están en verde (44/44 y 25/25) y
los archivos servidos por GitHub Pages son idénticos a `main`. El trabajo pendiente es de
**documentación y dependencias**, no de funcionalidad. El repositorio está **limpio y sincronizado**
(HEAD = `origin/main` = `425e5fd`).

## Qué funciona (verificado)

### Motor (worker)
- Inicialización de DuckDB‑WASM 1.29.0 con Blob Worker anidado (`initDuckDB` → versión
  `v1.0-phase-1b+multisheet`).
- **CSV:** `registerFileText` → `CREATE OR REPLACE TABLE` con `read_csv_auto(sample_size=-1,
  ignore_errors=true)` → conteo, esquema (`information_schema.columns`) y preview de 10 filas.
- **Parquet:** `registerFileBuffer` → `read_parquet` (llega por transferable).
- **Excel:** `XLSX.read` + `sheet_to_csv` por pestaña → **una tabla por hoja**.
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

### Plugin IA
- `query_excel_data` declarada en el nivel superior (requisito de TypingMind), `plugin.json` válido.
- Guardrail de solo lectura, timeout de 30 s, listener por llamada y respuesta Markdown ≤ 50 filas.
- Spec multi‑hoja con descubrimiento de tablas (`information_schema.tables` / `SHOW TABLES`,
  `DESCRIBE` / `information_schema.columns`) y aviso del saneado textual del guardrail.

### Verificaciones ejecutadas hoy (evidencia)

| Suite | Checks | Resultado |
|---|---|---|
| `tests/fase5-verificacion.html` | 44 | **OK 44 / FAIL 0** |
| `tests/fase3b-verificacion.html` | 25 | **OK 25 / FAIL 0** |

## Deriva de documentación (detectada, NO corregida)

`README.md` describe un repositorio que ya no existe en varios puntos. Verificado con
`Get-ChildItem -Force` y búsquedas en el archivo:

| Línea del README | Afirma | Realidad verificada |
|---|---|---|
| 414 | existe `index.html` (historial de versiones) | **No existe** (eliminado en la limpieza de `test-data`/`index.html`) |
| 415, 249 | existe `.nojekyll` "ya incluido" | **No existe** en el repo |
| 426‑427 | existen `backup/duckdb-worker.v0.4.23-contaminado.js` y `duckdb/typingmind-excel-engine-v0.3-test.js` | **Ninguno existe** |
| 477 | hitos: "ver carpetas `duckdb/` y `backup/`, no usar" | Las carpetas **no existen** |
| 498‑500 | "el repositorio **aún no incluye un archivo `LICENSE`**" | **`LICENSE` MIT © 2026 DoomsdayAILabs sí existe** |
| 83 | el plugin permite solo `SELECT`, `WITH` y `EXPLAIN` | el propio README (líneas 362, 468) y `plugin/README.md` documentan también `SHOW` y `DESCRIBE` (la lista negra no los bloquea) |
| 89‑90, 320 | el modelo empieza con `SELECT * FROM excel_data LIMIT 1;` para descubrir columnas | el spec actual ordena empezar por `information_schema.tables` / `SHOW TABLES` |

Deriva adicional fuera del README: `tests/sql-test-runner.html` mantiene el título
**"Batería de pruebas SQL — v0.4.20"** (línea 30) aunque el producto es v1.0.

## Qué falta / pendientes reales

1. **README:** corregir la deriva de la tabla anterior y decidir el texto de licencia.
2. **`tests/sql-test-runner.html`:** necesita que el usuario **seleccione a mano** `TEUs.xlsx` o
   `RequerimientoPrueba`; esos datos de ejemplo ya no están en el repo, así que la batería SQL más
   amplia (**38 casos**: 20 + 18) **no es automatizable hoy**. Ver `testing-strategy.md`.
3. **SheetJS sin pin de versión** en `duckdb-worker.js` (línea 6).
4. **Versión del producto:** el Worker declara `v1.0-phase-1b+multisheet` mientras el producto es
   `1.0`; no hay un único identificador de versión.

## Issues y riesgos conocidos

- **Tablas huérfanas entre cargas:** cada carga recrea la tabla base pero **no elimina** las
  `excel_data_<hoja>` de un libro anterior con más pestañas (documentado en README líneas 375‑381).
- **Guardrail = lista negra textual.** Solo se bloquean `drop`, `alter`, `insert`, `update`, `delete`,
  `create`, `copy`, `attach`, `export`, `pragma force`. Otras sentencias de DuckDB que no estén en esa
  lista y no sean de lectura (p. ej. `TRUNCATE`, `INSTALL`, `LOAD`, `SET`, `VACUUM`, `CHECKPOINT`)
  **no son rechazadas por el plugin**. El README documenta la lista como el mecanismo de defensa, así
  que **está por confirmar si es un límite asumido o un endurecimiento pendiente**.
  *Nota:* el widget (uso manual) ejecuta cualquier sentencia a propósito: la restricción solo aplica a
  lo que la IA puede lanzar por sí misma.
- **Falsos positivos del guardrail:** bloquea también si una de esas palabras aparece dentro de un
  identificador o literal (`SELECT "Update" FROM …`) — documentado y mitigado en el prompt del spec.
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

## Por confirmar

- Destino de los datos de ejemplo eliminados (`test-data`): ¿existen fuera del repo para uso manual?
- Si se quiere `index.html` de vuelta como historial de versiones descargables (el README lo anuncia).
- Estado de GitHub Pages como Pages "clásico" vs. Actions: no hay workflow en el repo.

