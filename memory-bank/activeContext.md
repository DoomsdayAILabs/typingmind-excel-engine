# Active Context — TypingMind Excel Engine

> Foco de trabajo actual, cambios recientes, siguientes pasos y decisiones activas.
> Última actualización: **2026‑09‑30** (cierre de Fase 6 · UI y sincronización del Memory Bank).

## Foco actual

**Fase 6 completada: selector de hojas multi-hoja en la UI.** El widget ya permite alternar entre las
tablas `excel_data_<hoja>` de un libro con un desplegable nativo, sin recargar el archivo, con
sustitución conservadora del SQL editado y metadato de columnas. Está **desplegada** en `origin/main`
(`d876db4`) y servida por GitHub Pages. Con esta fase, el soporte multi-hoja queda **100 % cerrado en
las tres piezas** (Worker, Plugin IA y UI).

## Estado del repositorio (verificado en esta sesión)

- Commit del producto (Fase 6 · UI): **`d876db4`** — `feat(ui): selector desplegable multi-hoja en
  widget y suite Fase 6 (85 checks)` (3 archivos, 612 inserciones, 13 borrados).
- Este documento se actualiza en el commit `docs: sincronizar memory-bank con Fase 6 y purgar
  referencias obsoletas en README`, inmediatamente posterior a `d876db4`.
- **Working tree limpio** y rama `main` sincronizada con `origin/main` antes y después del push.
- Sin `package.json`, sin CI, sin `.github/`. Repo estático servido por GitHub Pages.
- La extensión pasó de **645 a 849 líneas** (30 969 → 40 399 bytes): reescritura íntegra del archivo
  con 216 inserciones y 12 borrados exactos (cabecera, regla `.tmee-meta-info` y las 4 líneas internas
  que `renderizarResultado` delegó en las funciones nuevas).

## Verificación ejecutada en esta sesión

| Comprobación | Comando | Resultado |
|---|---|---|
| Sintaxis extensión | `node --check typingmind-excel-engine-v1.0.js` | OK (exit 0) |
| Suite Fase 6 (nueva) | `node tests/run-headless.js tests/fase6-multisheet-verificacion.html` | **85/85 OK** (exit 0) |
| Suite Fase 5 | `node tests/run-headless.js tests/fase5-verificacion.html` | **44/44 OK** (exit 0) |
| Suite Fase 3B | `node tests/run-headless.js tests/fase3b-verificacion.html` | **25/25 OK** (exit 0) |
| Push al remoto | `git ls-remote origin main` → `d876db4f6d53c323cb82298b58d744073699d915` | remoto sincronizado |
| GitHub Pages | descarga de `typingmind-excel-engine-v1.0.js` desde Pages | cabecera **“Fase 6”**, 39 550 bytes |
| README (bytes) | BOM + conteo CRLF/LF | sin BOM · 528 CRLF · 0 LF sueltos |

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

## Próximos pasos candidatos (no iniciados, por prioridad sugerida)

1. **Fijar la versión de SheetJS** en `duckdb-worker.js` (hoy se descarga `xlsx` sin pin) — riesgo
   vivo de cadena de suministro.
2. **Endurecer el guardrail del plugin**: valorar lista blanca o incluir `TRUNCATE`, `INSTALL`, `LOAD`,
   `SET`, `VACUUM` y `CHECKPOINT`, que hoy no se rechazan.
3. **Cerrar el hueco de la suite SQL manual:** incorporar fixtures mínimos (o generar un CSV/XLSX en
   memoria) para que `tests/sql-test-runner.html` (38 casos) sea ejecutable en headless.
4. **Decidir si se restaura `index.html`** como historial de versiones descargables: ya no se anuncia
   en el README, así que pasó de ser una deriva documental a una decisión de producto.

## Patrones y aprendizajes de esta sesión

- **Reescritura íntegra sin diffs por regex:** el archivo se regeneró por bloques de ~5‑6 KB usando un
  marcador textual (`// __FIN__`) sustituido en cada paso; permitió auditar el resultado con
  `git diff` y comprobar que solo cambiaba lo previsto.
- **El diff como test de la reescritura:** los 12 borrados fueron exactamente la cabecera, la regla
  `.tmee-meta-info` y las 4 líneas internas de `renderizarResultado`; todo lo demás fue inserción pura,
  por lo que las suites de Fase 3B y Fase 5 siguieron en verde sin tocarlas.
- **`FileReader` real ≠ determinista en headless:** con `--virtual-time-budget=30000` la lectura de
  archivos no siempre termina dentro de la ventana medida y dejaba cargas a medias
  (`estado = "Cargando..."`, 26 checks en rojo). La suite de Fase 6 **stubea `FileReader`** —igual que
  ya hacía con `fetch` y `Worker`— porque su objeto es la UI, no el I/O del navegador.
- **Por qué el FAB no cambia:** la barra de hojas vive **dentro** de `.tmee-body`, que el modo icono
  oculta con `display: none !important`; el test lo mide (50×50 px) y también con el rect a 0.
- **`core.autocrlf=true`:** el working tree conserva CRLF (README: 528 CRLF / 0 LF) pero el blob del
  repo se normaliza a LF, así que GitHub Pages sirve 39 550 bytes frente a los 40 399 del disco; la
  diferencia es exactamente 1 byte por línea. **No es un problema de sincronía.**
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

- **SheetJS sin pin** (dependencia externa que puede cambiar sin aviso).
- **Guardrail por lista negra** en el plugin: solo bloquea las palabras listadas (ver
  `progress.md` § Issues y riesgos conocidos).
- **Datos de ejemplo ausentes**: la suite SQL más completa (**38 casos** en dos baterías) requiere
  selección manual de archivos que ya no están en el repo → no es ejecutable sin intervención humana.
- **Selector silencioso con una sola hoja cargada:** si el libro tiene una única hoja con datos, la
  barra no se muestra y las hojas omitidas no se anuncian (ver § Por confirmar).
