# Testing Strategy — TypingMind Excel Engine

> Cómo se prueba el proyecto, qué cubre cada suite y qué no es automatizable hoy.
> Resultados ejecutados y verificados el **2026‑10‑01** sobre el producto con Fase 8 · guardrail (`be03349`) y las **7 suites automatizadas** (**363 checks** en verde): 22 de Fase 10 contra DuckDB-WASM real (`--serve`), 39 del runner SQL sintético, 50 de Fase 9 y 98 de Fase 8.

## Principios

- Los tests son **páginas HTML autocontenidas**: no hay framework, ni dependencias, ni `package.json`.
- Cada página publica su resultado en `document.title` con el prefijo **`RES:`** y el JSON
  URL‑encoded (`{ checks: [{ nombre, ok, extra }], logs: [] }`).
- `tests/run-headless.js` (Node) lanza Chrome/Edge con `--headless=new --dump-dom
  --virtual-time-budget=30000`, lee el `<title>` y reporta `OK/FAIL`; **sale con código 1 si hay
  algún FAIL**.
- Los tests del widget **stubean** `fetch` y `Worker` antes de cargar el motor, por lo que validan la
  lógica de UI/puente sin red ni DuckDB real. La suite de Fase 6 **stubea además `FileReader`**: con
  headless y `--virtual-time-budget` la lectura real de archivos no siempre termina dentro de la ventana
  medida y dejaba cargas a medias (`estado = "Cargando..."`).

## Cómo ejecutarlas

```powershell
# Fase 8 — guardrail de solo lectura y contrato postMessage (98 checks)
node tests/run-headless.js tests/fase8-guardrail-verificacion.html

# Fase 6 — selector de hojas multi-hoja en la UI (85 checks)
node tests/run-headless.js tests/fase6-multisheet-verificacion.html

# Fase 5 — UI móvil, FAB, arrastre, inyección (44 checks)
node tests/run-headless.js tests/fase5-verificacion.html

# Fase 3B — regresión de Markdown, CSV y puente postMessage (25 checks)
node tests/run-headless.js tests/fase3b-verificacion.html

# Fase 10 - verificacion real con DuckDB-WASM y SheetJS (servidor efimero) -> 22 checks
node tests/run-headless.js tests/fase10-duckdb-real-verificacion.html --serve --budget=180000

# Fase 10 - bateria SQL con datasets sinteticos en memoria (servidor efimero) -> 39 checks
node tests/run-headless.js 'tests/sql-test-runner.html?headless=1' --serve --budget=180000

# Fase 9 - ciclo de vida de las tablas gestionadas y de los listeners (50 checks)
node tests/run-headless.js tests/fase9-ciclo-vida-verificacion.html

# Ver también los checks que pasan y fijar el tamaño de ventana
node tests/run-headless.js tests/fase5-verificacion.html 900,760 --all
```

- **Dos modos de ejecución:** las **5 suites rápidas** (Fase 9, 8, 6, 5 y 3B) corren en `file://` con tiempo
  virtual y sin red (instantáneas); las **2 suites reales** (Fase 10 y el runner SQL) declaran
  `<meta name="tmee-live">` y se lanzan con `--serve --budget=180000`: el runner levanta un servidor HTTP
  efímero (puerto libre), la página descarga DuckDB-WASM y SheetJS de jsDelivr y envía su informe por
  `POST /__report`, que el runner espera en **tiempo real** (con tiempo virtual los temporizadores de la
  página se adelantarían al Worker que compila DuckDB-WASM).
- El runner autodetecta Chrome o Edge (o usa la variable de entorno `CHROME_PATH`).
  En esta máquina: Chrome en `C:/Program Files/Google/Chrome/Application/chrome.exe` y
  Edge en `C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe`.
- Ventana por defecto: `900,760`.

## Inventario de pruebas

| Archivo | Tipo | Cobertura | Estado |
|---|---|---|---|
| `tests/fase6-multisheet-verificacion.html` | automática headless | **85 checks**: barra oculta con CSV/Parquet/Excel de 1 hoja; libro de 5 pestañas → 5 opciones (3 habilitadas, 2 deshabilitadas) con etiquetas (`42 filas`, `omitida (vacía)`, `omitida (sin datos)`, `solo cabeceras`); resumen `Hojas 3/5 · 2 vacías` en ámbar + tooltip por hoja; orden DOM metadatos → barra → consola; selección inicial = tabla primaria (≠ hoja 0); cambio de hoja con metadatos y **una sola** consulta nueva (sin recargar el archivo); preservación de la consulta editada y de la tabla entrecomillada; selecciones vacías/ajenas al libro ignoradas; reset por archivo nuevo y por carga fallida; FAB 50×50 y arrastre de cabecera intactos; reglas CSS de Fase 6 y de Fase 5 | **85/85 OK** |
| `tests/fase5-verificacion.html` | automática headless | 44 checks: widget inyectado, modo icono 50×50, geometry/clamping, arrastre ratón y táctil, umbral clic vs. arrastre, supresión del ratón sintético, auto‑minimizado al enviar, reglas CSS exigidas | **44/44 OK** |
| `tests/fase3b-verificacion.html` | automática headless | 25 checks: botón "💬 Enviar a TM", escapado Markdown (pipe, salto de línea → `<br>`, `null`), límite de 50 filas + aviso de omitidas, evento `input` para React, suma al texto existente, alertas, fallback a portapapeles, exportación CSV, ausencia del botón ➖ | **25/25 OK** |
| `tests/fase10-duckdb-real-verificacion.html` | automatica headless (**22 checks**, `--serve`) | Worker real + DuckDB-WASM 1.29.0 + SheetJS 0.18.5 desde jsDelivr: initDuckDB real; Excel A de 3 hojas -> `SHOW TABLES` e `information_schema.tables` con 3 tablas; Excel B de 2 hojas -> **purga real de `excel_data_Resumen`**; hoja 0 vacia -> **eliminacion real de la tabla fantasma `excel_data`** (y `SELECT` falla con Catalog Error); `CREATE TABLE mia` sobrevive a la carga siguiente; **case-insensitivity real** del `DROP TABLE IF EXISTS "mitabla"`; idempotencia y catalogo final | **22/22 OK** |
| `tests/fase9-ciclo-vida-verificacion.html` | automática headless | **50 checks**: ciclo de vida de las tablas gestionadas del Worker (Worker real cargado como script global con dobles de `XLSX`/`db`/`conn` y catálogo en memoria): primera carga sin DROPs; libro con menos hojas → purga de la huérfana; **hoja 0 vacía → se elimina la tabla fantasma `excel_data`**; CSV/Parquet tras Excel; tabla del usuario intacta y fuera del Map; carga fallida que no destruye el dataset; reintento tras fallo de `DROP`; colisión `_2`; forma del SQL (`IF EXISTS` + grafía original); y los listeners del plugin retirados al responder, al vencer el timeout y al fallar `postMessage` | **50/50 OK** |
| `tests/fase8-guardrail-verificacion.html` | automática headless | **98 checks**: banco de **67 consultas** del guardrail: los 14 falsos positivos históricos (literales e identificadores entrecomillados con palabras reservadas) y los 16 falsos negativos históricos (`TRUNCATE`, `INSTALL`, `LOAD`, `SET`, `CALL`, `VACUUM`, `CHECKPOINT`, `PRAGMA force_checkpoint`, `MERGE`, `USE`, `BEGIN`, sentencias apiladas, verbo oculto tras comentario o tras literal con `$$...$$`), más 22 comprobaciones de contrato del plugin (nombre/spec, mensajes, canal `tm-excel-engine`, timeout de 30 s, Markdown, `\|` escapado, `null` → celda vacía, 0 filas, 50 de 60 filas) | **98/98 OK** |
| `tests/fase3b-visual.html` | manual/visual | capturas de los estados `#min`, `#drag`, `#send` (usa `File` en memoria con `ventas.csv`, sin red) | por confirmar visualmente |
| `tests/sql-test-runner.html` | automatica headless (**39 checks**) | **37 consultas SQL** en dos baterias sobre **datasets sinteticos en memoria** (sin archivos externos): `TESTS_TEUS` (**20 casos**, 18 filas con self-test aritmetico de las 8 constantes) y `buildRequerimientoTests(totalRows)` (**17 casos**, 120 filas deterministas) + **1 check de carga** + **1 self-test** = **39 checks** | **39/39 OK** |
| `test-widget.html` | manual | simulador de TypingMind (102 líneas) que carga el motor real; sirve para probar carga de archivos y widget con un servidor local | uso manual |
| `test-duckdb-worker.html` | manual | banco de pruebas del **Worker real** (`new Worker("./duckdb-worker.js")`, línea 121; `initDuckDB` con timeout de 120 s), con `<input type="file" accept=".xlsx, .xls">` (línea 98). Requiere servidor local (usa ruta relativa) | uso manual |

## Estado verificado hoy (evidencia literal)

```text
== tests/fase10-duckdb-real-verificacion.html (900,760) => checks: 22 | OK: 22 | FAIL: 0  (exit 0)
== tests/sql-test-runner.html?headless=1 (900,760) => checks: 39 | OK: 39 | FAIL: 0  (exit 0)
== tests/fase9-ciclo-vida-verificacion.html (900,760) => checks: 50 | OK: 50 | FAIL: 0  (exit 0)
== tests/fase8-guardrail-verificacion.html (900,760) => checks: 98 | OK: 98 | FAIL: 0  (exit 0)
== tests/fase6-multisheet-verificacion.html (900,760) => checks: 85 | OK: 85 | FAIL: 0   (exit 0)
== tests/fase5-verificacion.html (900,760) => checks: 44 | OK: 44 | FAIL: 0             (exit 0)
== tests/fase3b-verificacion.html (900,760) => checks: 25 | OK: 25 | FAIL: 0            (exit 0)
== banco del guardrail (67 consultas, plugin/implementation.js) => FP: 0 | FN: 0       (exit 0)
== contrato del plugin (canal, timeout 30 s, Markdown, 50 filas) => 22/22 OK           (exit 0)
```

Comprobaciones adicionales sin ejecución de navegador:

```text
node --check typingmind-excel-engine-v1.0.js   → exit 0
node --check duckdb-worker.js                  → exit 0
node --check plugin/implementation.js          → exit 0
JSON.parse(plugin/plugin.json)                 → JSON válido
```

## Limitaciones del sistema de pruebas

1. ~~**`tests/sql-test-runner.html` requería seleccionar a mano** `TEUs.xlsx` o `RequerimientoPrueba`.~~
   **Resuelto en Fase 10 (`f7e9510`):** la batería usa **datasets sintéticos en memoria** (18 filas TEUs
   con self-test aritmético + 120 filas Requerimiento deterministas), corre en headless con
   `?headless=1 --serve` y publica `RES:` (**39 checks**). El `<input type="file">` se conserva solo como
   modo manual opcional.
2. **Los tests del widget no ejercitan DuckDB real**: mockean `fetch` y `Worker`. La carga real de
   CSV/Parquet/Excel y la ejecución SQL solo se validan con las páginas manuales (o el runner SQL) y
   con navegador + servidor local.
3. **`file://` no sirve para pruebas manuales**: los Workers y los módulos ES quedan bloqueados; hay
   que usar `python -m http.server 8080` o `npx serve .`.
4. **Las transiciones CSS no avanzan en headless**: la suite de Fase 5 las desactiva para medir la
   geometría final del icono flotante y verifica la declaración `transition` sobre el texto del CSS.
5. **Cobertura del plugin (Fase 8):** el guardrail se valida ahora con la suite headless
   `tests/fase8-guardrail-verificacion.html` (**98/98 OK**), que carga `plugin/implementation.js` real en un
   contexto `vm` con `window`/`setTimeout` instrumentados: **banco de 67 consultas** del guardrail (0 FP /
   0 FN) más **22 comprobaciones de contrato** del plugin. Ya **no queda hueco de cobertura en `tests/`**
   para el plugin; la suite de Fase 3B sigue ejercitando el puente desde el lado del motor. La suite de Fase 9
(`tests/fase9-ciclo-vida-verificacion.html`, **50 checks**) cubre además el ciclo de vida de las tablas del
Worker (purga de huérfanas y tabla fantasma) y los tres desenlaces del listener del plugin.
6. **El I/O real de archivos no se prueba en headless:** la suite de Fase 6 stubea `FileReader` para ser
   determinista bajo `--virtual-time-budget`. La lectura real de CSV/XLSX se valida con las páginas
   manuales (`test-widget.html`, `test-duckdb-worker.html`) contra un servidor local.

## Protocolo al cambiar código (regla del repo)

1. Ejecutar **las 5 suites rápidas** (Fase 9, Fase 8, Fase 6, Fase 5 y Fase 3B) y dejarlas en verde antes de dar
   por bueno un cambio en `typingmind-excel-engine-v1.0.js` o `duckdb-worker.js`. Si el cambio toca el ciclo
   de vida de las tablas, el runner o la infraestructura de pruebas, añadir las **2 suites reales**
   (Fase 10 y el runner SQL) con `--serve --budget=180000`.
2. Ampliar los checks de `tests/` cuando se cambie el comportamiento del widget o del plugin.
3. Mantener sincronizados `README.md` (producto) y `plugin/README.md` (integración IA).

## Pendiente recomendado

- ~~Incorporar fixtures mínimos para la batería SQL y cerrar el hueco de cobertura.~~ **Completado en
  Fase 10 (`f7e9510`):** datasets sintéticos en memoria + `?headless=1 --serve` (**39 checks**).
- ~~**Convertir el banco del guardrail en suite headless**~~ **Completado:** el banco de 67 consultas y las 22
  comprobaciones de contrato viven desde ahora en `tests/fase8-guardrail-verificacion.html` (**98/98 OK**,
  convención `RES:`), así que la cobertura del plugin forma parte de la batería automática.
- ~~**Verificación manual contra DuckDB real (Fase 9).**~~ **Completado en Fase 10 (`f7e9510`):**
  `tests/fase10-duckdb-real-verificacion.html` (**22 checks**) lo comprueba de forma automática con el
  Worker y DuckDB-WASM reales (`--serve`).
