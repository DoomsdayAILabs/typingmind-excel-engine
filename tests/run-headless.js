/**
 * Runner headless de las verificaciones del proyecto (Chrome/Edge + --dump-dom).
 *
 * Uso:  node tests/run-headless.js <ruta-test-relativa[?query]> [ancho,alto] [--all] [--serve] [--budget=ms]
 *
 *   --all           imprime también los checks que pasan
 *   --serve         sirve el repositorio por HTTP efímero (127.0.0.1:0) en lugar de file:// y habilita
 *                   el canal POST /__report: las suites que lo usan (las que necesitan red y DuckDB-WASM
 *                   real) envían su resultado por HTTP y el runner cierra Chrome en cuanto lo recibe, sin
 *                   depender del tiempo virtual de --dump-dom.
 *   --budget=<ms>   presupuesto de tiempo virtual de Chrome (por defecto 30000) y, en modo de tiempo
 *                   real, tiempo máximo de espera del informe POST.
 *   --live          fuerza el modo de tiempo real (sin --dump-dom ni tiempo virtual). Se activa solo
 *                   con --serve, y también automáticamente si la página declara
 *                   <meta name="tmee-live" content="1">: las suites con trabajo asíncrono real (red y
 *                   DuckDB-WASM en un Worker) no pueden usar tiempo virtual, porque Chrome adelanta los
 *                   temporizadores de la página mientras el Worker sigue trabajando.
 *
 * Requiere Chrome o Edge instalado; se puede forzar la ruta con la variable CHROME_PATH.
 * El test debe publicar su resultado en document.title con el prefijo "RES:" (JSON URL-encoded).
 */
const { spawn } = require("child_process");
const fs = require("fs");
const http = require("http");
const path = require("path");

const REPO = path.resolve(__dirname, "..");

// Tipos MIME que necesita el servidor efímero (DuckDB-WASM sirve su .wasm por HTTP).
const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".wasm": "application/wasm",
  ".css": "text/css; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".ico": "image/x-icon",
  ".map": "application/json; charset=utf-8",
  ".txt": "text/plain; charset=utf-8",
  ".md": "text/markdown; charset=utf-8",
  ".csv": "text/csv; charset=utf-8"
};

const argumentos = process.argv.slice(2);
const rel = argumentos.find((a) => !a.startsWith("--") && !/^\d+,\d+$/.test(a));
if (!rel) {
  console.error("Uso: node tests/run-headless.js <ruta-test-relativa[?query]> [ancho,alto] [--all] [--serve] [--budget=ms]");
  process.exit(1);
}
const ventana = argumentos.find((a) => /^\d+,\d+$/.test(a)) || "900,760";
const verbose = argumentos.includes("--all");
const servir = argumentos.includes("--serve");
const presupuesto = (argumentos.find((a) => a.startsWith("--budget=")) || "").split("=")[1] || "30000";
const [relRuta, relConsulta] = rel.split("?");
const archivo = path.join(REPO, relRuta);
if (!fs.existsSync(archivo)) {
  console.error("No existe el archivo: " + archivo);
  process.exit(1);
}

// Modo de tiempo real: obligatorio para las suites con trabajo asíncrono real (red + Worker).
// Se declara con <meta name="tmee-live" content="1"> o con el flag --live.
const marcadorTiempoReal = /<meta[^>]+name=["']tmee-live["']/i.test(fs.readFileSync(archivo, "utf8"));
const tiempoReal = servir && (marcadorTiempoReal || argumentos.includes("--live"));

const candidatos = [
  process.env.CHROME_PATH,
  "C:/Program Files/Google/Chrome/Application/chrome.exe",
  "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
  (process.env.LOCALAPPDATA || "") + "/Google/Chrome/Application/chrome.exe",
  "C:/Program Files/Microsoft/Edge/Application/msedge.exe",
  "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
].filter(Boolean);
const navegador = candidatos.find((c) => fs.existsSync(c));
if (!navegador) {
  console.error("No se encontró Chrome/Edge. Define la variable CHROME_PATH.");
  process.exit(1);
}

// --- Servidor estático efímero + canal de informe (solo con --serve) ---
function levantarServidor() {
  let entregarInforme;
  const informe = new Promise((resolve) => { entregarInforme = resolve; });

  const servidor = http.createServer((req, res) => {
    if (req.method === "POST" && req.url.split("?")[0] === "/__report") {
      let cuerpo = "";
      req.on("data", (trozo) => { cuerpo += trozo; });
      req.on("end", () => {
        res.writeHead(204, { "Access-Control-Allow-Origin": "*" });
        res.end();
        entregarInforme(cuerpo);
      });
      return;
    }

    try {
      const url = new URL(req.url, "http://127.0.0.1");
      const destino = path.join(REPO, decodeURIComponent(url.pathname));
      if (!destino.startsWith(REPO)) throw new Error("ruta fuera del repositorio");
      const datos = fs.readFileSync(destino);
      res.writeHead(200, {
        "Content-Type": MIME[path.extname(destino).toLowerCase()] || "application/octet-stream",
        "Cache-Control": "no-store"
      });
      res.end(datos);
    } catch (error) {
      res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("404 " + req.url);
    }
  });

  return new Promise((resolve) => {
    servidor.listen(0, "127.0.0.1", () => resolve({ servidor: servidor, informe: informe, puerto: servidor.address().port }));
  });
}

function lanzarNavegador(url, modoTiempoReal) {
  const banderas = ["--headless=new", "--disable-gpu", "--no-sandbox", "--window-size=" + ventana];
  if (!modoTiempoReal) banderas.push("--dump-dom", "--virtual-time-budget=" + presupuesto);
  const hijo = spawn(navegador, banderas.concat([url]));
  let stdout = "";
  let stderr = "";
  hijo.stdout.on("data", (trozo) => { stdout += trozo; });
  hijo.stderr.on("data", (trozo) => { stderr += trozo; });
  const completado = new Promise((resolve) => hijo.on("close", () => resolve({ stdout: stdout, stderr: stderr })));
  return { hijo: hijo, completado: completado };
}

const esperar = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function leerResultado(texto) {
  const limpio = String(texto || "").trim();
  if (!limpio) return null;
  try { return JSON.parse(limpio); } catch (error) { /* puede venir URL-encoded */ }
  try { return JSON.parse(decodeURIComponent(limpio)); } catch (error) { return null; }
}

function extraerDelDom(html) {
  const coincidencia = String(html || "").match(/<title>RES:([^<]*)<\/title>/);
  return coincidencia ? coincidencia[1] : null;
}

(async function main() {
  let servidor = null;
  let informe = null;
  let url;

  if (servir) {
    const servicio = await levantarServidor();
    servidor = servicio.servidor;
    informe = servicio.informe;
    url = "http://127.0.0.1:" + servicio.puerto + "/" + relRuta.replace(/\\/g, "/") + (relConsulta ? "?" + relConsulta : "");
  } else {
    url = "file:///" + path.join(REPO, relRuta).replace(/\\/g, "/") + (relConsulta ? "?" + relConsulta : "");
  }

  console.log(
    "== " + rel + " (" + ventana + ", budget " + presupuesto + " ms, " +
    (servir ? "http" : "file") + ", " + (tiempoReal ? "tiempo real" : "tiempo virtual") + ") => " + url
  );
  const proceso = lanzarNavegador(url, tiempoReal);
  let cuerpo = null;

  if (servir) {
    cuerpo = await Promise.race([
      informe,
      proceso.completado.then(() => null),
      esperar(Number(presupuesto) + (tiempoReal ? 0 : 20000)).then(() => null)
    ]);
    proceso.hijo.kill();
  }

  const { stdout, stderr } = await proceso.completado;
  if (servidor) servidor.close();

  const carga = cuerpo !== null ? cuerpo : extraerDelDom(stdout);
  const res = leerResultado(carga);

  if (!res || !Array.isArray(res.checks)) {
    const titulo = String(stdout).match(/<title>([^<]*)<\/title>/);
    console.log("SIN RESULTADOS. title=" + (titulo ? titulo[1] : "(ninguno)") + "  informe POST=" + (cuerpo === null ? "no" : "si"));
    if (tiempoReal) console.log("modo de tiempo real: se esperaba un POST a /__report desde la página.");
    console.log("longitud DOM=" + String(stdout).length);
    if (stderr) console.log("stderr (200 primeros): " + String(stderr).slice(0, 200).replace(/\s+/g, " "));
    process.exit(2);
  }
  let ok = 0, ko = 0;
  for (const c of res.checks) {
    if (c.ok) ok++;
    else { ko++; console.log("FAIL -> " + c.nombre + (c.extra ? "   [" + c.extra + "]" : "")); }
    if (verbose) console.log((c.ok ? "ok  " : "FAIL") + " | " + c.nombre + (c.extra ? "  [" + c.extra + "]" : ""));
  }
  console.log("== " + rel + " (" + ventana + ") => checks: " + res.checks.length + " | OK: " + ok + " | FAIL: " + ko);
  process.exit(ko === 0 ? 0 : 1);
})();
