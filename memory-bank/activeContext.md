# Active Context — TypingMind Excel Engine

> Foco de trabajo actual, cambios recientes, siguientes pasos y decisiones activas.
> Última actualización: **2026‑09‑30** (inicialización del Memory Bank).

## Foco actual

**El producto v1.0 está cerrado y desplegado.** La tarea de esta sesión fue crear el Memory Bank a
partir de una inspección real del repositorio (sin inventar nada), versionarlo (`425e5fd`) y dejarlo
al día con ese commit de referencia. No se ha modificado ni una línea de código de producto: los
únicos archivos nuevos son los de `memory-bank/` y `.clinerules/`.

## Estado del repositorio (verificado hoy)

- Commit de referencia: **`425e5fd`** — `docs: inicializar memory bank` (incluye `.clinerules/` y los
  8 archivos de `memory-bank/`). Rama `main`, sincronizada con `origin/main`.
- **Working tree limpio:** no hay cambios pendientes ni archivos sin trackear. El código de producto
  no se ha tocado desde `5067b16`.
- Sin `package.json`, sin CI, sin `.github/`. Repo estático.

## Verificación ejecutada en esta sesión

| Comprobación | Comando | Resultado |
|---|---|---|
| Sintaxis extensión | `node --check typingmind-excel-engine-v1.0.js` | OK (exit 0) |
| Sintaxis worker | `node --check duckdb-worker.js` | OK (exit 0) |
| Sintaxis plugin | `node --check plugin/implementation.js` | OK (exit 0) |
| JSON del plugin | `JSON.parse(plugin.json)` | Válido |
| Suite Fase 5 | `node tests/run-headless.js tests/fase5-verificacion.html` | **44/44 OK** (exit 0) |
| Suite Fase 3B | `node tests/run-headless.js tests/fase3b-verificacion.html` | **25/25 OK** (exit 0) |
| Sincronía con GitHub Pages | comparación de texto descargado vs. local | **Idéntico** en ambos archivos de producto |

Ver detalle en `testing-strategy.md` y `techContext.md`.

## Decisiones activas

1. **El Memory Bank vive en `memory-bank/` en la raíz del repo** (6 archivos núcleo + 2 de contexto:
   `plugin-bridge-protocol.md` y `testing-strategy.md`).
2. **Regla de oro heredada del README (§ Contribuir):** no tocar `typingmind-excel-engine-v1.0.js` ni
   `duckdb-worker.js` sin dejar **las suites de `tests/` en verde**; y mantener sincronizados
   `README.md` (producto) y `plugin/README.md` (integración IA).
3. **La retrocompatibilidad del nombre `excel_data` es intocable**: el plugin, las suites SQL y la
   documentación dependen de que la hoja 0 conserve el nombre base.
4. **Ningún dato sale del navegador**: no introducir backend, telemetría ni subida de archivos.

## Próximos pasos candidatos (no iniciados, por prioridad sugerida)

1. **Sincronizar la documentación con la realidad del repo** (deriva detectada, lista exacta en
   `progress.md` § Deriva de documentación): README habla de `index.html`, `.nojekyll`, `backup/`,
   `duckdb/` y de "no hay LICENSE" — ninguno de esos puntos es cierto hoy.
2. **Fijar la versión de SheetJS** en `duckdb-worker.js` (hoy se descarga `xlsx` sin pin).
3. **Documentar cómo obtener los archivos de muestra** que exige `tests/sql-test-runner.html`
   (`TEUs.xlsx`, `RequerimientoPrueba`), porque ya no viven en el repo.

## Patrones y aprendizajes de esta sesión

- **Las lecturas por rango (`start_line`/`end_line`) no se aplican sobre archivos ya cacheados** en
  esta sesión: devolvieron "outdated". Alternativa que funcionó: `Get-Content | Select-Object -Skip N
  -First M` para revisar secciones concretas. Útil para no releer archivos completos.
- **El README no es fuente de verdad por sí solo**: describe un árbol de archivos que ya no existe.
  Para el estado real, manda el working tree + `git log` + la ejecución de las suites.
- **La comparación de hashes de archivos de texto entre local y un servidor da falsos negativos** por
  BOM/CRLF; hay que normalizar (quitar BOM, `\r\n` → `\n`) antes de comparar.
- `git --no-pager` evita quedarse colgado en el paginador dentro de la shell no interactiva.

## Por confirmar

- Ubicación/estado de los datos de ejemplo (`TEUs.xlsx`, `RequerimientoPrueba`) necesarios para la
  suite SQL manual: fueron retirados del repo.
- Si se desea restaurar `index.html` (historial de versiones descargables) que el README anuncia y
  `.nojekyll` que el README declara incluido.
- Autoridad final sobre la licencia: `LICENSE` (MIT © 2026 DoomsdayAILabs) vs. README § Licencia.

## Riesgos vivos

- **SheetJS sin pin** (dependencia externa que puede cambiar sin aviso).
- **Guardrail por lista negra** en el plugin: solo bloquea las palabras listadas (ver
  `progress.md` § Issues y riesgos conocidos).
- **Datos de ejemplo ausentes**: la suite SQL más completa (**38 casos** en dos baterías) requiere
  selección manual de archivos que ya no están en el repo → no es ejecutable sin intervención humana.
