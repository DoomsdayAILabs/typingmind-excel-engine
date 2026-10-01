# Active Context — TypingMind Excel Engine

> Foco de trabajo actual, cambios recientes, siguientes pasos y decisiones activas.
> Última actualización: **2026‑09‑30** (Fase 8 · guardrail blindado del plugin, commit `be03349`; 0 FP / 0 FN sobre el banco de 67 consultas).

## Foco actual

**Fase 8 completada: el guardrail del plugin deja de ser una lista negra textual.** `plugin/implementation.js`
reescribe la validación en cuatro comprobaciones en cadena: saneado de comentarios y literales en una pasada
(`saneaSQL`), lista blanca `LIMITADO` (`SELECT`/`WITH`/`EXPLAIN`/`SHOW`/`DESCRIBE`), lista negra ampliada
`PROHIBIDO` y rechazo de sentencias apiladas y de literales sin cerrar. Medido sobre un banco de **67
consultas**: **0 falsos positivos** (antes 14: literales como `WHERE estado = 'DELETE'` e identificadores
como `SELECT "Update"`) y **0 falsos negativos** (antes 16: `TRUNCATE`, `INSTALL`, `LOAD`, `SET`, `CALL`,
`VACUUM`, `CHECKPOINT`, `PRAGMA force_checkpoint`, `MERGE`, `USE`, `BEGIN`, sentencias apiladas y dos modos
de ocultar el verbo tras un comentario o un literal). El contrato del plugin (canal, timeout de 30 s,
Markdown ≤ 50 filas y mensajes) queda intacto (22/22 comprobaciones) y las tres suites headless siguen en
verde. Commit del producto: **`be03349`**.

**Fase 7 completada: estabilidad de dependencias y consistencia documental.** SheetJS queda **fijado a
`0.18.5`** en el Worker (constante `SHEETJS_PACKAGE`; era la única dependencia sin pin), el guardrail
del plugin anuncia las cinco sentencias permitidas (`SELECT`, `WITH`, `EXPLAIN`, `SHOW`, `DESCRIBE`) y
el README alinea su lista de sentencias (líneas 94‑95) y el descubrimiento de esquema
(`information_schema.tables` / `SHOW TABLES` → `information_schema.columns` / `DESCRIBE`), pasando a
**527 líneas**. Con Fase 6 ya cerrada, el soporte multi-hoja sigue **100 % en las tres piezas** (Worker,
Plugin IA y UI).

## Estado del repositorio (verificado en esta sesión)

- Commit del producto (Fase 8 · guardrail): **`be03349`** — `feat(plugin): guardrail robusto de solo
  lectura con lista blanca y saneado de literales (Fase 8)` (1 archivo, +81 / -2).
- Commits previos del producto: **`686a46e`** (Fase 7 · pin de SheetJS `0.18.5`) y **`d876db4`** (Fase 6 · UI).
- Este documento se actualiza en el commit `docs: sincronizar spec de plugin, docs y memory-bank con
  guardrail Fase 8`, inmediatamente posterior a `be03349`.
- **Working tree limpio** y rama `main` sincronizada con `origin/main` tras el push.
- Sin `package.json`, sin CI, sin `.github/`. Repo estático servido por GitHub Pages.
- La extensión se mantiene en **849 líneas**; el Worker pasa a **600 líneas** al añadir la constante
  `SHEETJS_PACKAGE` y `importScripts(SHEETJS_PACKAGE)`.

## Verificación ejecutada en esta sesión

| Comprobación | Comando | Resultado |
|---|---|---|
| Sintaxis (3 archivos) | `node --check duckdb-worker.js` · `plugin/implementation.js` · `typingmind-excel-engine-v1.0.js` | OK (exit 0 ×3) |
| Suite Fase 6 | `node tests/run-headless.js tests/fase6-multisheet-verificacion.html` | **85/85 OK** (exit 0) |
| Suite Fase 5 | `node tests/run-headless.js tests/fase5-verificacion.html` | **44/44 OK** (exit 0) |
| Suite Fase 3B | `node tests/run-headless.js tests/fase3b-verificacion.html` | **25/25 OK** (exit 0) |
| Banco del guardrail (67 consultas sobre el plugin real) | `node -` (arnés `vm` en memoria) | **0 FP / 0 FN** (exit 0) |
| Contrato del plugin (canal, 30 s, Markdown, 50 filas) | mismo arnés sobre `plugin/implementation.js` | **22/22 OK** (exit 0) |
| Push al remoto | `git ls-remote origin main` → hash de `686a46e` y del commit de docs | remoto sincronizado |
| GitHub Pages | descarga de `typingmind-excel-engine-v1.0.js` desde Pages | cabecera **“Fase 6”** (la extensión no cambió en Fase 7) |
| Artefacto de jsDelivr | `HEAD .../xlsx@0.18.5/dist/xlsx.full.min.js` vs. URL sin pin | 200 · **881 727 bytes en ambos** (equivalencia exacta) |
| `latest` de npm `xlsx` | `https://data.jsdelivr.com/v1/packages/npm/xlsx` | `{"latest":"0.18.5"}` |
| README | BOM + líneas + conteo CRLF/LF | sin BOM · **527 líneas** · 527 CRLF · 0 LF sueltos · 32 445 bytes |
| Worker | BOM + líneas + conteo CRLF/LF | con BOM · **600 líneas** · 600 CRLF · 0 LF sueltos |
| Integridad de caracteres | conteo de `U+FFFD` y `U+00C3` en los 3 archivos | **0 y 0** (sin doble codificación) |

Ver detalle en `testing-strategy.md` y `techContext.md`.

## Decisiones activas

1. **El Memory Bank vive en `memory-bank/` en la raíz del repo** (6 archivos núcleo + 2 de contexto:
   `plugin-bridge-protocol.md` y `testing-strategy.md`).
2. **Regla de oro heredada del README (§ Contribuir):** no tocar `typingmind-excel-engine-v1.0.js` ni
   `duckdb-worker.js` sin dejar **las suites de `tests/` en verde** (hoy son **tres**: Fase 6, Fase 5
   y Fase 3B); y mantener sincronizados `README.md` (producto) y `plugin/README.md` (integración IA).
3. **La retrocompatibilidad del nombre `excel_data` es intocable**: el plugin, las suites SQL y la
   documentación dependen de que la hoja 0 conserve el nombre base.
4. **Ningún dato sale del navegador**: no introducir backend, telemetría ni subida de archivos.
5. **Se conservan los JSDoc de la extensión** (~40 KB): decisión expresa de mantener los comentarios
   internos; el tamaño se considera óptimo frente a su valor de documentación.
6. **Visibilidad de la barra de hojas:** el selector aparece **solo** si el libro tiene **más de una
   hoja con datos** (`cargadas > 1`). Con CSV, Parquet o un Excel de una sola hoja la barra queda
   oculta y el widget es idéntico al de Fase 5 (regla aprobada y cubierta por tests).
7. **Sustitución conservadora del SQL:** al cambiar de hoja se reescribe la plantilla solo si el
   usuario no la editó; si la editó, se sustituye **únicamente** el token `FROM <tabla>`. Con la tabla
   entrecomillada no hay coincidencia y el texto del usuario se respeta literalmente.
8. **El README ya está purgado:** se eliminaron las referencias a `index.html`, `.nojekyll`, `backup/`
   y `duckdb/`, y la sección de licencia declara la **MIT oficial** (`LICENSE`, © 2026
   DoomsdayAILabs).
9. **Guardrail del plugin (Fase 8) con cuatro decisiones asumidas:** `PRAGMA` bloqueado (el esquema se
   consulta con `DESCRIBE` o `information_schema`), prohibida la forma `(SELECT ...)` (la sentencia debe
   empezar por palabra clave formal), sin soporte de cadenas `$$...$$` (escáner simple y determinista) y
   rechazo estricto de sentencias apiladas (`;` fuera de literales): una llamada = una consulta.

## Próximos pasos candidatos (no iniciados, por prioridad sugerida)

1. ~~**Endurecer el guardrail del plugin**~~ **Completado en `be03349` (Fase 8):** lista blanca + saneado
   de comentarios y literales + lista negra ampliada (escritura, DDL, DCL y administración) + bloqueo de
   sentencias apiladas; 0 FP / 0 FN sobre el banco de 67 consultas.
2. **Cerrar el hueco de la suite SQL manual:** incorporar fixtures mínimos (o generar un CSV/XLSX en
   memoria) para que `tests/sql-test-runner.html` (38 casos) sea ejecutable en headless.
3. **Decidir si se restaura `index.html`** como historial de versiones descargables: ya no se anuncia
   en el README, así que pasó de ser una deriva documental a una decisión de producto.
4. **Título de `tests/sql-test-runner.html`:** sigue anunciando "v0.4.20" con el producto en v1.0 (la
   única deriva documental que queda, ver `progress.md`).
5. **Suite headless del guardrail:** portar el banco de 67 consultas a
   `tests/fase8-guardrail-verificacion.html` (convención `RES:`) para que la cobertura del plugin entre en
   la batería automática (ver `testing-strategy.md` § Pendiente recomendado).

## Patrones y aprendizajes de esta sesión

- **Edición byte-segura con BOM/CRLF:** para tocar `duckdb-worker.js` (UTF‑8 **con** BOM) y los `.md`
  (sin BOM) se usó `ReadAllText` + `WriteAllText` con `UTF8Encoding($true/$false)` y sustituciones con
  **aserción de unicidad** previa: si el ancla no aparece exactamente una vez, se aborta sin escribir.
- **Trampa de codepoints al inyectar acentos:** `después` se escribió con `[char]0xF3` (ó) en lugar de
  `[char]0xE9` (é) y el error **solo** se vio al revisar el diff. La verificación de integridad
  (`U+FFFD == 0` y `U+00C3 == 0`) y el recuento por codepoint lo detectan de forma mecánica.
- **Anclas con guiones invisibles:** `projectbrief.md` usa guiones ASCII (`2026-09-30`) mientras
  `activeContext.md` y `progress.md` usan `U+2011`; un ancla copiada del render falla. El dump previo
  con codepoints escapados (`<U+XXXX>`) elimina la ambigüedad.
- **`FileReader` real ≠ determinista en headless:** con `--virtual-time-budget=30000` la lectura de
  archivos no siempre termina dentro de la ventana medida y dejaba cargas a medias
  (`estado = "Cargando..."`, 26 checks en rojo). La suite de Fase 6 **stubea `FileReader`** —igual que
  ya hacía con `fetch` y `Worker`— porque su objeto es la UI, no el I/O del navegador.
- **Por qué el FAB no cambia:** la barra de hojas vive **dentro** de `.tmee-body`, que el modo icono
  oculta con `display: none !important`; el test lo mide (50×50 px) y también con el rect a 0.
- **`core.autocrlf=true`:** el working tree conserva CRLF (README: 527 líneas / 527 CRLF / 0 LF) pero
  el blob del repo se normaliza a LF, así que GitHub Pages sirve ~1 byte menos por línea que el disco.
  **No es un problema de sincronía.**
- **`git push` puede devolver `exit 1` en falso:** PowerShell interpreta el progreso de git en `stderr`
  como `NativeCommandError` y oculta el mensaje real. Repetido con `cmd /c "git push origin main 2>&1"`
  respondió `Everything up-to-date`: el push ya había funcionado. Confirmar siempre con `git ls-remote`
  antes de declarar un fallo.
- **Las lecturas por rango (`start_line`/`end_line`) pueden fallar sobre archivos ya cacheados** en la
  sesión; alternativa: leer el archivo completo o `Get-Content | Select-Object -Skip N -First M`.
- `git --no-pager` evita quedarse colgado en el paginador dentro de la shell no interactiva.

## Por confirmar

- **Caso borde del selector:** con **1 hoja cargada y 1 vacía** la barra permanece oculta (regla
  `cargadas <= 1`), así que la hoja omitida no se anuncia. Cambiarlo es una línea
  (`if (cargadas.length <= 1 && omitidas + conError === 0)`), pendiente de decisión del usuario.
- Ubicación/estado de los datos de ejemplo (`TEUs.xlsx`, `RequerimientoPrueba`) necesarios para la
  suite SQL manual: fueron retirados del repo.
- Si se desea restaurar `index.html` (historial de versiones descargables).
- Identificador único de versión: el Worker declara `v1.0-phase-1b+multisheet` mientras el producto es
  `1.0`; no hay un único número de versión.

## Riesgos vivos

- **SheetJS fijado pero sin SRI:** el pin a `0.18.5` elimina la deriva de versión, pero `importScripts`
  no admite `integrity`, así que un compromiso del CDN no se detectaría (ver `techContext.md`).
- **Guardrail textual (no es un parser SQL):** la Fase 8 lo endurece (lista blanca + saneado + lista negra
  ampliada + bloqueo de apiladas) con 0 FP / 0 FN sobre el banco de 67 consultas, pero sigue siendo
  heurístico. Queda como línea futura el endurecimiento estructural: una conexión DuckDB restringida para
  el plugin, que hoy comparte conexión con la consola manual del usuario.
- **Datos de ejemplo ausentes**: la suite SQL más completa (**38 casos** en dos baterías) requiere
  selección manual de archivos que ya no están en el repo → no es ejecutable sin intervención humana.
- **Selector silencioso con una sola hoja cargada:** si el libro tiene una única hoja con datos, la
  barra no se muestra y las hojas omitidas no se anuncian (ver § Por confirmar).
