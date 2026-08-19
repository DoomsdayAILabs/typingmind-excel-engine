/*
 * TypingMind Excel Data Engine - Extension v0.2
 * Client-side only. No server required.
 *
 * Purpose:
 * - Select an XLSX file directly from the browser.
 * - Process it locally with DuckDB-Wasm.
 * - Never attach the selected XLSX to the TypingMind chat.
 * - Return only small query results that the user can copy to the chat.
 *
 * This is a prototype. It does NOT yet expose the local database directly
 * as an AI tool. That is planned for v0.2.
 */

(() => {
  "use strict";

  const APP_ID = "tm-excel-engine";
  const DB_VERSION = "1.5.4";
  const CDN_MODULE = `https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@${DB_VERSION}/dist/duckdb-browser.mjs`;

  let duckdbModule = null;
  let db = null;
  let conn = null;
  let currentFile = null;
  let currentFileName = null;
  let currentTable = null;

  const $ = (sel, root = document) => root.querySelector(sel);

  function escapeIdent(name) {
    return '"' + String(name).replaceAll('"', '""') + '"';
  }

  function escapeString(value) {
    return "'" + String(value).replaceAll("'", "''") + "'";
  }

  function formatBytes(bytes) {
    if (!Number.isFinite(bytes)) return "";
    const units = ["B", "KB", "MB", "GB"];
    let i = 0, n = bytes;
    while (n >= 1024 && i < units.length - 1) { n /= 1024; i++; }
    return `${n.toFixed(i ? 2 : 0)} ${units[i]}`;
  }

  function setStatus(text, kind = "info") {
    const el = $("#tmxe-status");
    if (!el) return;
    el.textContent = text;
    el.dataset.kind = kind;
  }

  function ensureStyles() {
    if ($("#tmxe-styles")) return;
    const style = document.createElement("style");
    style.id = "tmxe-styles";
    style.textContent = `
      #tmxe-launcher {
        position: fixed;
        right: 18px;
        bottom: 86px;
        z-index: 2147483000;
        border: 0;
        border-radius: 999px;
        padding: 11px 15px;
        background: #2563eb;
        color: white;
        font: 600 14px/1.1 system-ui,sans-serif;
        box-shadow: 0 6px 20px rgba(0,0,0,.25);
        cursor: pointer;
      }
      #tmxe-overlay {
        position: fixed;
        inset: 0;
        z-index: 2147483001;
        background: rgba(0,0,0,.48);
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 16px;
      }
      #tmxe-panel {
        width: min(900px, 96vw);
        max-height: 90vh;
        overflow: auto;
        background: Canvas;
        color: CanvasText;
        border: 1px solid rgba(127,127,127,.35);
        border-radius: 16px;
        padding: 18px;
        box-shadow: 0 20px 70px rgba(0,0,0,.35);
        font: 14px/1.45 system-ui,sans-serif;
      }
      #tmxe-panel h2 { margin: 0 0 6px; font-size: 20px; }
      #tmxe-help { opacity: .72; margin-bottom: 14px; }
      #tmxe-file { width: 100%; margin: 8px 0 12px; }
      #tmxe-query {
        width: 100%;
        min-height: 110px;
        resize: vertical;
        box-sizing: border-box;
        font: 13px/1.45 ui-monospace,SFMono-Regular,Menlo,monospace;
        padding: 10px;
        border-radius: 10px;
        border: 1px solid rgba(127,127,127,.45);
        background: transparent;
        color: inherit;
      }
      #tmxe-actions { display:flex; gap:8px; flex-wrap:wrap; margin:10px 0; }
      #tmxe-actions button, #tmxe-close {
        border: 1px solid rgba(127,127,127,.45);
        border-radius: 9px;
        padding: 8px 12px;
        background: transparent;
        color: inherit;
        cursor:pointer;
      }
      #tmxe-run { background:#2563eb !important; color:white !important; border-color:#2563eb !important; }
      #tmxe-result {
        white-space: pre-wrap;
        overflow:auto;
        max-height: 360px;
        padding: 12px;
        border-radius: 10px;
        background: rgba(127,127,127,.10);
        font: 12px/1.45 ui-monospace,SFMono-Regular,Menlo,monospace;
      }
      #tmxe-meta { margin: 8px 0; opacity:.85; }
      #tmxe-status { margin: 8px 0; min-height: 20px; }
      #tmxe-status[data-kind="error"] { color:#dc2626; }
      #tmxe-status[data-kind="ok"] { color:#16a34a; }
      @media (max-width: 600px) {
        #tmxe-launcher { right: 10px; bottom: 80px; }
        #tmxe-panel { padding: 14px; }
      }
    `;
    document.head.appendChild(style);
  }

  function makeButton() {
    if ($("#tmxe-launcher")) return;
    const b = document.createElement("button");
    b.id = "tmxe-launcher";
    b.type = "button";
    b.textContent = "📊 Excel Engine";
    b.title = "Procesar Excel localmente sin adjuntarlo al chat";
    b.addEventListener("click", openPanel);
    document.body.appendChild(b);
  }

  function openPanel() {
    if ($("#tmxe-overlay")) return;
    const overlay = document.createElement("div");
    overlay.id = "tmxe-overlay";
    overlay.innerHTML = `
      <div id="tmxe-panel" role="dialog" aria-modal="true">
        <div style="display:flex;justify-content:space-between;gap:12px;align-items:center">
          <div>
            <h2>📊 Excel Data Engine</h2>
            <div id="tmxe-help">
              El archivo se procesa en este dispositivo. No lo adjuntes al chat.
            </div>
          </div>
          <button id="tmxe-close" type="button">Cerrar</button>
        </div>

        <label><strong>1. Selecciona un archivo XLSX</strong></label>
        <input id="tmxe-file" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet">

        <div id="tmxe-meta">Sin archivo cargado.</div>
        <div id="tmxe-status">Esperando archivo…</div>

        <label><strong>2. Consulta SQL</strong></label>
        <textarea id="tmxe-query" spellcheck="false"
          placeholder="Ejemplo: SELECT COUNT(*) AS filas FROM excel_data"></textarea>

        <div id="tmxe-actions">
          <button id="tmxe-example" type="button">Consulta de ejemplo</button>
          <button id="tmxe-run" type="button">Ejecutar consulta</button>
          <button id="tmxe-copy" type="button">Copiar resultado</button>
        </div>

        <div><strong>Resultado</strong></div>
        <pre id="tmxe-result">—</pre>

        <div style="margin-top:10px;opacity:.7">
          v0.2 — DuckDB-Wasm local. No servidor propio.
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    $("#tmxe-close").addEventListener("click", () => overlay.remove());
    overlay.addEventListener("click", e => {
      if (e.target === overlay) overlay.remove();
    });

    $("#tmxe-file").addEventListener("change", handleFile);
    $("#tmxe-example").addEventListener("click", () => {
      $("#tmxe-query").value = currentTable
        ? `SELECT COUNT(*) AS filas FROM ${escapeIdent(currentTable)}`
        : "SELECT 'Selecciona primero un archivo' AS mensaje";
    });
    $("#tmxe-run").addEventListener("click", runQuery);
    $("#tmxe-copy").addEventListener("click", async () => {
      const text = $("#tmxe-result").textContent || "";
      try {
        await navigator.clipboard.writeText(text);
        setStatus("Resultado copiado al portapapeles.", "ok");
      } catch {
        setStatus("No se pudo copiar automáticamente. Selecciona y copia el resultado.", "error");
      }
    });
  }

  async function loadDuckDB() {
    if (db) return;
    setStatus("Cargando DuckDB-Wasm…");
    duckdbModule = await import(CDN_MODULE);
    const bundles = duckdbModule.getJsDelivrBundles();
    const bundle = await duckdbModule.selectBundle(bundles);

    const workerURL = URL.createObjectURL(
      new Blob([`importScripts("${bundle.mainWorker}");`], { type: "text/javascript" })
    );

    const worker = new Worker(workerURL);
    const logger = new duckdbModule.ConsoleLogger();
    db = new duckdbModule.AsyncDuckDB(logger, worker);

    await db.instantiate(bundle.mainModule, bundle.pthreadWorker);
    URL.revokeObjectURL(workerURL);

    conn = await db.connect();

    // DuckDB-Wasm provides an official Excel extension.
    try {
      await conn.query("LOAD excel;");
    } catch {
      // Some bundles may autoload the extension when read_xlsx is used.
    }
  }

  async function handleFile(event) {
    const file = event.target.files?.[0];
    if (!file) return;

    currentFile = file;
    currentFileName = file.name;
    currentTable = "excel_data";

    try {
      setStatus(`Leyendo ${file.name} localmente…`);
      $("#tmxe-meta").textContent = `${file.name} — ${formatBytes(file.size)}`;

      await loadDuckDB();

      const bytes = new Uint8Array(await file.arrayBuffer());
      await db.registerFileBuffer(currentFileName, bytes);

      setStatus("Detectando estructura del Excel…");

      // Create a logical table from the first worksheet.
      // Header inference is delegated to DuckDB.
      await conn.query(`
        CREATE OR REPLACE VIEW ${escapeIdent(currentTable)} AS
        SELECT * FROM read_xlsx(${escapeString(currentFileName)}, header = true);
      `);

      const count = await conn.query(`SELECT COUNT(*) AS filas FROM ${escapeIdent(currentTable)}`);
      const row = count.toArray()[0];
      const filas = row?.filas?.toString?.() ?? String(row?.filas ?? "?");

      const columns = await conn.query(`DESCRIBE ${escapeIdent(currentTable)}`);
      const names = columns.toArray().map(r => r.column_name);

      $("#tmxe-meta").textContent =
        `${file.name} — ${formatBytes(file.size)} — ${filas} filas — ${names.length} columnas`;

      $("#tmxe-query").value =
        `SELECT COUNT(*) AS filas FROM ${escapeIdent(currentTable)}`;

      setStatus("Archivo cargado y procesado localmente.", "ok");
      $("#tmxe-result").textContent =
        `Columnas:\n${names.join("\n")}\n\nFilas: ${filas}`;
    } catch (err) {
      console.error(err);
      setStatus(`Error: ${err?.message || err}`, "error");
      $("#tmxe-result").textContent = String(err?.stack || err);
    }
  }

  async function runQuery() {
    if (!conn || !currentTable) {
      setStatus("Primero selecciona un archivo XLSX.", "error");
      return;
    }

    const sql = $("#tmxe-query").value.trim();
    if (!sql) {
      setStatus("Escribe una consulta SQL.", "error");
      return;
    }

    try {
      setStatus("Ejecutando consulta local…");
      const start = performance.now();
      const result = await conn.query(sql);
      const elapsed = performance.now() - start;

      const rows = result.toArray().map(r => {
        const out = {};
        for (const [k, v] of Object.entries(r)) {
          out[k] = typeof v === "bigint" ? v.toString() : v;
        }
        return out;
      });

      $("#tmxe-result").textContent = JSON.stringify(rows, null, 2);
      setStatus(`Consulta completada en ${elapsed.toFixed(0)} ms.`, "ok");
    } catch (err) {
      console.error(err);
      setStatus(`SQL error: ${err?.message || err}`, "error");
      $("#tmxe-result").textContent = String(err?.stack || err);
    }
  }

  function init() {
    ensureStyles();
    makeButton();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init, { once: true });
  } else {
    init();
  }
})();
