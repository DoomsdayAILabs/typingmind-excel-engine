# Testing Strategy — TypingMind Excel Engine

> Cómo se prueba el proyecto, qué cubre cada suite y qué no es automatizable hoy.
> Resultados ejecutados y verificados el **2026‑09‑30** (commit `5067b16`).

## Principios

- Los tests son **páginas HTML autocontenidas**: no hay framework, ni dependencias, ni `package.json`.
- Cada página publica su resultado en `document.title` con el prefijo **`RES:`** y el JSON
  URL‑encoded (`{ checks: [{ nombre, ok, extra }], logs: [] }`).
- `tests/run-headless.js` (Node) lanza Chrome/Edge con `--headless=new --dump-dom
  --virtual-time-budget=30000`, lee el `<title>` y reporta `OK/FAIL`; **sale con código 1 si hay
  algún FAIL**.
- Los tests del widget **stubean** `fetch` y `Worker` antes de cargar el motor, por lo que validan la
  lógica de UI/puente sin red ni DuckDB real.

## Cómo ejecutarlas

```powershell
# Fase 5 — UI móvil, FAB, arrastre, inyección (44 checks)
node tests/run-headless.js tests/fase5-verificacion.html

# Fase 3B — regresión de Markdown, CSV y puente postMessage (25 checks)
node tests/run-headless.js tests/fase3b-verificacion.html

# Ver también los checks que pasan y fijar el tamaño de ventana
node tests/run-headless.js tests/fase5-verificacion.html 900,760 --all
```

- El runner autodetecta Chrome o Edge (o usa la variable de entorno `CHROME_PATH`).
  En esta máquina: Chrome en `C:/Program Files/Google/Chrome/Application/chrome.exe` y
  Edge en `C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe`.
- Ventana por defecto: `900,760`.

## Inventario de pruebas

| Archivo | Tipo | Cobertura | Estado |
|---|---|---|---|
| `tests/fase5-verificacion.html` | automática headless | 44 checks: widget inyectado, modo icono 50×50, geometry/clamping, arrastre ratón y táctil, umbral clic vs. arrastre, supresión del ratón sintético, auto‑minimizado al enviar, reglas CSS exigidas | **44/44 OK** |
| `tests/fase3b-verificacion.html` | automática headless | 25 checks: botón "💬 Enviar a TM", escapado Markdown (pipe, salto de línea → `<br>`, `null`), límite de 50 filas + aviso de omitidas, evento `input` para React, suma al texto existente, alertas, fallback a portapapeles, exportación CSV, ausencia del botón ➖ | **25/25 OK** |
| `tests/fase3b-visual.html` | manual/visual | capturas de los estados `#min`, `#drag`, `#send` (usa `File` en memoria con `ventas.csv`, sin red) | por confirmar visualmente |
| `tests/sql-test-runner.html` | **manual** | **38 casos** en dos baterías: `TESTS_TEUS` (línea 291, **20 casos** sobre `TEUs.xlsx`, 18 filas de datos: COUNT, GROUP BY + SUM, WHERE, SUM FILTER, COUNT DISTINCT, STRING_AGG, concatenación, EXTRACT YEAR/MONTH, AVG/SUM/COUNT, porcentaje con NULLIF, ROW_NUMBER, suma acumulada, LAG/LEAD, CTEs + JOIN, subconsulta AVG, CASE, HAVING + JOIN, BETWEEN/IN/COALESCE) y `buildRequerimientoTests(totalRows)` (línea 534, **18 casos** sobre `RequerimientoPrueba`: Fecha, Area, Categoria, Empleado, Turno, Equipo; hasta 60 000 filas, 1‑2 min de carga) | **no ejecutable sin intervención humana** (ver abajo) |
| `test-widget.html` | manual | simulador de TypingMind (102 líneas) que carga el motor real; sirve para probar carga de archivos y widget con un servidor local | uso manual |
| `test-duckdb-worker.html` | manual | banco de pruebas del **Worker real** (`new Worker("./duckdb-worker.js")`, línea 121; `initDuckDB` con timeout de 120 s), con `<input type="file" accept=".xlsx, .xls">` (línea 98). Requiere servidor local (usa ruta relativa) | uso manual |

## Estado verificado hoy (evidencia literal)

```text
== tests/fase5-verificacion.html (900,760) => checks: 44 | OK: 44 | FAIL: 0      (exit 0)
== tests/fase3b-verificacion.html (900,760) => checks: 25 | OK: 25 | FAIL: 0      (exit 0)
```

Comprobaciones adicionales sin ejecución de navegador:

```text
node --check typingmind-excel-engine-v1.0.js   → exit 0
node --check duckdb-worker.js                  → exit 0
node --check plugin/implementation.js          → exit 0
JSON.parse(plugin/plugin.json)                 → JSON válido
```

## Limitaciones del sistema de pruebas

1. **`tests/sql-test-runner.html` requiere seleccionar a mano** `TEUs.xlsx` o `RequerimientoPrueba`
   (`<input id="excel-file" type="file" accept=".xlsx,.xls">`). Esos datos de ejemplo fueron
   eliminados del repo (`chore: limpieza definitiva de test-data e index.html`), así que la batería
   SQL **no se puede ejecutar de forma desatendida** hasta decidir si se reincorporan fixtures.
2. **Los tests del widget no ejercitan DuckDB real**: mockean `fetch` y `Worker`. La carga real de
   CSV/Parquet/Excel y la ejecución SQL solo se validan con las páginas manuales (o el runner SQL) y
   con navegador + servidor local.
3. **`file://` no sirve para pruebas manuales**: los Workers y los módulos ES quedan bloqueados; hay
   que usar `python -m http.server 8080` o `npx serve .`.
4. **Las transiciones CSS no avanzan en headless**: la suite de Fase 5 las desactiva para medir la
   geometría final del icono flotante y verifica la declaración `transition` sobre el texto del CSS.
5. **No hay cobertura automatizada del plugin** (`plugin/implementation.js`) más allá del puente que
   la suite de Fase 3B ejercita indirectamente desde el lado del motor.

## Protocolo al cambiar código (regla del repo)

1. Ejecutar **ambas** suites headless y dejarlas en verde antes de dar por bueno un cambio en
   `typingmind-excel-engine-v1.0.js` o `duckdb-worker.js`.
2. Ampliar los checks de `tests/` cuando se cambie el comportamiento del widget o del plugin.
3. Mantener sincronizados `README.md` (producto) y `plugin/README.md` (integración IA).

## Pendiente recomendado

- Incorporar fixtures mínimos (o generar un CSV/XLSX en memoria) para poder ejecutar la batería SQL
  en headless de forma desatendida y cerrar el único hueco grande de cobertura.
