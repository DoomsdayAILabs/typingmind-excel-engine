/**
 * Runner headless de las verificaciones del widget (Chrome/Edge + --dump-dom).
 *
 * Uso:  node tests/run-headless.js tests/fase5-verificacion.html [ancho,alto] [--all]
 *       --all  imprime también los checks que pasan
 *
 * Requiere Chrome o Edge instalado; se puede forzar la ruta con la variable CHROME_PATH.
 * El test debe publicar su resultado en document.title con el prefijo "RES:" (JSON URL-encoded).
 */
const { execFileSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const REPO = path.resolve(__dirname, "..");
const rel = process.argv[2];
if (!rel) {
  console.error("Uso: node tests/run-headless.js <ruta-test-relativa> [ancho,alto] [--all]");
  process.exit(1);
}
const ventana = process.argv[3] && /^\d+,\d+$/.test(process.argv[3]) ? process.argv[3] : "900,760";
const verbose = process.argv.includes("--all");

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

const url = "file:///" + path.join(REPO, rel).replace(/\\/g, "/");
const salida = execFileSync(navegador, [
  "--headless=new", "--disable-gpu", "--no-sandbox", "--dump-dom",
  "--window-size=" + ventana, "--virtual-time-budget=30000", url
], { encoding: "utf8", maxBuffer: 128 * 1024 * 1024 });

const coincidencia = salida.match(/<title>RES:([^<]*)<\/title>/);
if (!coincidencia) {
  const titulo = salida.match(/<title>([^<]*)<\/title>/);
  console.log("SIN RESULTADOS en document.title. title=" + (titulo ? titulo[1] : "(ninguno)"));
  console.log("longitud DOM=" + salida.length);
  process.exit(2);
}

const res = JSON.parse(decodeURIComponent(coincidencia[1]));
let ok = 0, ko = 0;
for (const c of res.checks) {
  if (c.ok) ok++;
  else { ko++; console.log("FAIL -> " + c.nombre + (c.extra ? "   [" + c.extra + "]" : "")); }
  if (verbose) console.log((c.ok ? "ok  " : "FAIL") + " | " + c.nombre + (c.extra ? "  [" + c.extra + "]" : ""));
}
console.log("== " + rel + " (" + ventana + ") => checks: " + res.checks.length + " | OK: " + ok + " | FAIL: " + ko);
process.exit(ko === 0 ? 0 : 1);
