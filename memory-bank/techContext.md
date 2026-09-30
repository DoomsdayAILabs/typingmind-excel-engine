# Tech Context — TypingMind Excel Engine

> Tecnologías, entorno de desarrollo, restricciones y dependencias.
> Verificado en la máquina local (Windows, 2026‑09‑30) sobre el commit `5067b16`.

## Stack

| Capa | Tecnología | Versión / origen |
|---|---|---|
| Motor SQL | **DuckDB‑WASM** | `1.29.0` — `https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm` (worker, línea 4) |
| Excel | **SheetJS (XLSX)** | `https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js` (worker, línea 6) — **sin versión fijada** |
| Concurrencia | Web Worker + Blob URL + Blob Worker anidado | API nativa del navegador |
| UI | JavaScript vanilla + CSS inyectado | sin framework, sin bundler |
| IA | TypingMind Plugins (function calling) | `plugin.json` formato nativo; `outputType: respond_to_ai` |
| Test runner | Node.js + Chrome/Edge headless | `tests/run-headless.js` (`execFileSync`, `--dump-dom`) |
| Hosting | GitHub Pages (`main`) | `https://doomsdayailabs.github.io/typingmind-excel-engine/` |
| Licencia | MIT © 2026 DoomsdayAILabs (`LICENSE`) | contradice el README, ver "Por confirmar" |

## Entorno de desarrollo

- **No hay `package.json`, ni `node_modules`, ni bundler, ni linter, ni CI.** El repo es 100 % estático.
- **Node.js v24.20.0** instalado localmente (necesario solo para `tests/run-headless.js`).
- Navegadores detectados en la máquina: **Chrome** en
  `C:/Program Files/Google/Chrome/Application/chrome.exe` y **Edge** en
  `C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe` (el runner los autodetecta o usa
  `CHROME_PATH`).
- **Servidor local obligatorio para pruebas manuales** (con `file://` el navegador bloquea Workers y
  módulos ES): `python -m http.server 8080` o `npx serve .`.
- Codificación: los archivos de producto están en **UTF‑8 con BOM y CRLF**; al desplegarse por Pages
  se sirven en LF (el texto es idéntico, ver más abajo).

## Comandos útiles (todos verificados hoy)

```powershell
# Comprobación de sintaxis (sin ejecutar nada)
node --check typingmind-excel-engine-v1.0.js      # exit 0
node --check duckdb-worker.js                     # exit 0
node --check plugin/implementation.js             # exit 0
node -e "JSON.parse(require('fs').readFileSync('plugin/plugin.json','utf8'))"   # JSON válido

# Suites headless (requieren Chrome o Edge)
node tests/run-headless.js tests/fase5-verificacion.html     # 44/44 OK
node tests/run-headless.js tests/fase3b-verificacion.html    # 25/25 OK
node tests/run-headless.js tests/fase5-verificacion.html 900,760 --all   # imprime todos los checks

# Servidor local para pruebas manuales
python -m http.server 8080
```

## Restricciones técnicas conocidas

1. **CORS en Workers:** no se puede `new Worker()` cross‑origin → Blob URL (y Blob Worker anidado
   para el binario de DuckDB).
2. **Dependencia de CDN en el primer arranque:** sin acceso a `cdn.jsdelivr.net` no se cargan
   DuckDB‑WASM ni SheetJS. Mitigación documentada: autoalojar los bundles y cambiar `DUCKDB_PACKAGE`.
3. **Sin persistencia:** al recargar la página se pierde el motor y los datos.
4. **`file://` inutiliza la app:** Workers y módulos ES quedan bloqueados.
5. **Excel se parsea completo en memoria** con SheetJS antes de crear la tabla; para ficheros muy
   grandes el camino recomendado es CSV/Parquet (escaneo optimizado de DuckDB).
6. **`WORKER_PATH` es una URL absoluta** (`typingmind-excel-engine-v1.0.js` línea 16): un fork debe
   actualizarla además de la URL del `<script>`.
7. **`read_csv_auto(sample_size=-1, ignore_errors=true)`**: escaneo completo para inferir tipos
   (coste de tiempo en archivos grandes) y tolerancia a filas malformadas.
8. **En headless las transiciones CSS no avanzan**: la suite de Fase 5 las desactiva y verifica la
   declaración `transition` sobre el CSS.

## Despliegue (estado verificado hoy)

Los dos archivos de producto servidos por GitHub Pages **coinciden línea a línea** con el working
tree en `5067b16` (comparación de texto normalizado; solo difieren en BOM/CRLF):

| Archivo | HTTP | ¿Texto idéntico al local? |
|---|---|---|
| `duckdb-worker.js` | 200 | **Sí** (600 líneas en ambos) |
| `typingmind-excel-engine-v1.0.js` | 200 | **Sí** (646 líneas en ambos) |

⇒ El despliegue está **sincronizado** con la rama `main`. El plugin NO se despliega por Pages: se
pega a mano en TypingMind (ver `plugin-bridge-protocol.md`).

## Política de dependencias / versiones

- DuckDB‑WASM está **fijado** a `1.29.0` (la descripción del producto presume esa versión).
- SheetJS se carga **sin pin de versión** (`xlsx/dist/...`): una publicación nueva del paquete puede
  cambiar el comportamiento del Worker sin tocar el repo. **Riesgo abierto** (ver `progress.md`).

## Por confirmar

- Si existe (o debe existir) un `.nojekyll`: el README lo declara "ya incluido", pero **no está en el
  repo**; Pages sirve igualmente los dos archivos (verificado con HTTP 200), así que hoy no bloquea.
- Versión exacta del producto en el badge del README: `1.0` para el producto y
  `v1.0-phase-1b+multisheet` para el Worker (`WORKER_VERSION`, worker línea 3).
