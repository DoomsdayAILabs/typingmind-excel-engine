# Project Brief — TypingMind Excel Engine

> Documento fundacional. Fuente de verdad del alcance.
> Derivado de la inspección directa del repo; commit de referencia **`425e5fd`** (2026-09-30). El código
> de producto no ha cambiado desde `5067b16`. Nada aquí es especulativo.

## Qué es

**Extensión + plugin para TypingMind** que inyecta un motor de base de datos analítico (DuckDB‑WASM)
dentro de la propia interfaz de chat. El usuario carga CSV / Parquet / Excel en un widget flotante,
consulta con SQL real y la IA del chat puede consultar esos datos por sí misma mediante
*function calling* (`query_excel_data`). Todo el cómputo ocurre en el navegador.

Repo: `DoomsdayAILabs/typingmind-excel-engine` · rama `main` · producto en **v1.0**.

## Artefactos de producto (3 piezas)

| Pieza | Archivo | Dónde vive |
|---|---|---|
| Extensión UI + puente | `typingmind-excel-engine-v1.0.js` (645 líneas) | TypingMind → Settings → Extensions |
| Motor (Worker) | `duckdb-worker.js` (599 líneas) | GitHub Pages (lo descarga la extensión) |
| Plugin IA | `plugin/plugin.json` + `plugin/implementation.js` | TypingMind → Plugins |

## Objetivos núcleo

1. **Privacidad:** los datos nunca salen del equipo (sin backend, sin uploads, sin API keys de datos).
2. **Tamaño:** trabajar con archivos que no caben en el contexto de un LLM.
3. **Precisión:** el modelo no estima; consulta el dato con SQL y devuelve el resultado del motor.
4. **Agilidad:** explorar, verificar y exportar sin salir de TypingMind.

## Requisitos no negociables (verificados en código)

- **Cero backend:** todo el procesamiento ocurre en la ventana principal + Web Worker del navegador.
- **Ciclo de vida del VFS:** los archivos temporales se registran en el VFS de DuckDB y se **eliminan**
  (`dropFile`) en el `finally` de cada carga; no hay persistencia entre recargas.
- **Solo lectura para la IA:** el plugin bloquea sentencias de escritura; la consola SQL del *widget*
  (uso manual del usuario) no está restringida — es su propio motor local.
- **La IA nunca recibe el archivo:** solo tablas Markdown acotadas (50 filas) por consulta concreta.

## Formatos soportados

`.csv`, `.parquet`, `.xlsx`, `.xls` — (`accept` del input en `typingmind-excel-engine-v1.0.js` línea 194).
Excel multi‑hoja: **una tabla por pestaña** (`excel_data` + `excel_data_<hoja saneada>`).

## Fuera de alcance (documentado en README, líneas 373‑399)

- Persistencia e historial entre recargas.
- Trabajar con **dos archivos a la vez**: cada carga recrea la tabla base pero **no elimina** tablas
  sobrantes de cargas anteriores (hay que usar `UNION` o reexportar).
- Escritura de datos desde la IA.

## Criterios de aceptación verificables

- Widget inyectado, badge de estado y estado final **"Motor Listo"**.
- Carga correcta de los 4 formatos con metadatos, esquema y preview de 10 filas.
- `query_excel_data` devuelve Markdown con datos reales (máx. 50 filas).
- Suites de regresión en verde: `44/44` (Fase 5) y `25/25` (Fase 3B).

## Documentación y trazabilidad

- `README.md` — documentación de producto (506 líneas).
- `plugin/README.md` — integración del plugin / contrato del puente (161 líneas).
- `memory-bank/` — continuidad entre sesiones de trabajo (este directorio).

## Por confirmar

- **Titularidad de licencia:** `LICENSE` es **MIT © 2026 DoomsdayAILabs**, pero `README.md`
  (líneas 498‑500) afirma que el repo "aún no incluye un archivo `LICENSE`" y que los derechos quedan
  reservados. Falta decidir cuál de los dos textos es el vigente.
