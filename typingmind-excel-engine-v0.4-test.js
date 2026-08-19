/*
 * TypingMind Excel Data Engine - Extension v0.4 TEST
 *
 * Purpose:
 * - Load DuckDB-Wasm inside TypingMind.
 * - Select an XLSX file locally.
 * - Register the file in DuckDB-Wasm memory.
 * - Read the first worksheet using read_xlsx().
 * - Show row count, columns and a small preview.
 *
 * IMPORTANT:
 * - The XLSX is NOT uploaded to GitHub.
 * - The XLSX is NOT sent to the LLM.
 * - Processing occurs inside the browser.
 */

(() => {
  "use strict";

  const APP_ID = "tm-excel-engine-v04-test";

  /*
   * This is the npm package version used by the working v0.3 test.
   *
   * We keep the same package here because v0.3 already proved
   * that DuckDB-Wasm can load successfully inside TypingMind.
   */
  const DUCKDB_PACKAGE =
    "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

  let duckdb = null;
  let db = null;
  let conn = null;

  let currentFile = null;
  let currentFileName = null;
  let currentTable = "excel_data";

  function addStyles() {
    if (document.getElementById("tmxe-v04-style")) return;

    const style = document.createElement("style");

    style.id = "tmxe-v04-style";

    style.textContent = `
      #tmxe-v04-button {
        position: fixed;
        right: 18px;
        bottom: 86px;
        z-index: 2147483000;

        border: 0;
        border-radius: 999px;

        padding: 11px 16px;

        background: #2563eb;
        color: white;

        font: 600 14px/1.1 system-ui, sans-serif;

        box-shadow: 0 6px 20px rgba(0,0,0,.25);

        cursor: pointer;
      }

      #tmxe-v04-overlay {
        position: fixed;
        inset: 0;

        z-index: 2147483001;

        background: rgba(0,0,0,.48);

        display: flex;
        align-items: center;
        justify-content: center;

        padding: 16px;
      }

      #tmxe-v04-panel {
        width: min(850px, 96vw);
        max-height: 90vh;

        overflow: auto;

        background: Canvas;
        color: CanvasText;

        border: 1px solid rgba(127,127,127,.35);
        border-radius: 16px;

        padding: 20px;

        box-shadow: 0 20px 70px rgba(0,0,0,.35);

        font: 14px/1.45 system-ui, sans-serif;
      }

      #tmxe-v04-panel h2 {
        margin: 0 0 5px;
        font-size: 20px;
      }

      #tmxe-v04-help {
        opacity: .7;
        margin-bottom: 15px;
      }

      #tmxe-v04-file {
        width: 100%;
        margin: 8px 0 14px;
      }

      #tmxe-v04-meta {
        padding: 10px;
        border-radius: 9px;
        background: rgba(127,127,127,.10);
        margin-bottom: 10px;
      }

      #tmxe-v04-status {
        padding: 12px;
        border-radius: 9px;
        background: rgba(127,127,127,.10);
        white-space: pre-wrap;
        font: 13px/1.5 ui-monospace,
              SFMono-Regular,
              Menlo,
              monospace;
        margin-bottom: 12px;
      }

      #tmxe-v04-status[data-kind="ok"] {
        color: #16a34a;
      }

      #tmxe-v04-status[data-kind="error"] {
        color: #dc2626;
      }

      #tmxe-v04-result {
        white-space: pre-wrap;
        overflow: auto;
        max-height: 420px;

        padding: 12px;

        border-radius: 10px;

        background: rgba(127,127,127,.10);

        font: 12px/1.45 ui-monospace,
              SFMono-Regular,
              Menlo,
              monospace;
      }

      #tmxe-v04-actions {
        display: flex;
        gap: 8px;
        flex-wrap: wrap;
        margin: 10px 0;
      }

      #tmxe-v04-actions button,
      #tmxe-v04-close {
        border: 1px solid rgba(127,127,127,.45);

        border-radius: 9px;

        padding: 9px 13px;

        background: transparent;
        color: inherit;

        cursor: pointer;
      }

      #tmxe-v04-close {
        white-space: nowrap;
      }

      @media (max-width: 600px) {
        #tmxe-v04-button {
          right: 10px;
          bottom: 80px;
        }

        #tmxe-v04-panel {
          padding: 14px;
        }
      }
    `;

    document.head.appendChild(style);
  }

  function createButton() {
    if (document.getElementById("tmxe-v04-button")) return;

    const button = document.createElement("button");

    button.id = "tmxe-v04-button";
    button.type = "button";

    button.textContent = "📊 Excel v0.4";

    button.title = "Probar lectura local de Excel";

    button.addEventListener("click", openPanel);

    document.body.appendChild(button);
  }

  function openPanel() {
    if (document.getElementById("tmxe-v04-overlay")) return;

    const overlay = document.createElement("div");

    overlay.id = "tmxe-v04-overlay";

    overlay.innerHTML = `
      <div id="tmxe-v04-panel" role="dialog" aria-modal="true">

        <div style="
          display:flex;
          justify-content:space-between;
          align-items:flex-start;
          gap:12px;
        ">

          <div>
            <h2>📊 Excel Data Engine v0.4</h2>

            <div id="tmxe-v04-help">
              Prueba de lectura XLSX completamente local.
              El archivo no se adjunta al chat.
            </div>
          </div>

          <button id="tmxe-v04-close" type="button">
            Cerrar
          </button>

        </div>

        <strong>1. Selecciona un archivo XLSX</strong>

        <input
          id="tmxe-v04-file"
          type="file"
          accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        >

        <div id="tmxe-v04-meta">
          Sin archivo seleccionado.
        </div>

        <div id="tmxe-v04-status">
          Esperando archivo...
        </div>

        <div id="tmxe-v04-actions">

          <button id="tmxe-v04-count" type="button">
            🔢 Contar filas
          </button>

          <button id="tmxe-v04-preview" type="button">
            👁️ Ver primeras 5 filas
          </button>

          <button id="tmxe-v04-columns" type="button">
            📋 Ver columnas
          </button>

        </div>

        <strong>Resultado</strong>

        <pre id="tmxe-v04-result">—</pre>

        <div style="
          margin-top:12px;
          opacity:.65;
          font-size:12px;
        ">
          v0.4 — DuckDB-Wasm + XLSX local
        </div>

      </div>
    `;

    document.body.appendChild(overlay);

    document
      .getElementById("tmxe-v04-close")
      .addEventListener("click", () => {
        overlay.remove();
      });

    overlay.addEventListener("click", event => {
      if (event.target === overlay) {
        overlay.remove();
      }
    });

    document
      .getElementById("tmxe-v04-file")
      .addEventListener("change", handleFile);

    document
      .getElementById("tmxe-v04-count")
      .addEventListener("click", countRows);

    document
      .getElementById("tmxe-v04-preview")
      .addEventListener("click", previewRows);

    document
      .getElementById("tmxe-v04-columns")
      .addEventListener("click", showColumns);
  }

  function setStatus(text, kind = "") {
    const el =
      document.getElementById("tmxe-v04-status");

    if (!el) return;

    el.textContent = text;

    if (kind) {
      el.dataset.kind = kind;
    } else {
      delete el.dataset.kind;
    }
  }

  function setResult(value) {
    const el =
      document.getElementById("tmxe-v04-result");

    if (!el) return;

    if (typeof value === "string") {
      el.textContent = value;
    } else {
      el.textContent =
        JSON.stringify(value, null, 2);
    }
  }

  function formatBytes(bytes) {
    if (!Number.isFinite(bytes)) {
      return "";
    }

    const units = [
      "B",
      "KB",
      "MB",
      "GB"
    ];

    let i = 0;
    let value = bytes;

    while (
      value >= 1024 &&
      i < units.length - 1
    ) {
      value /= 1024;
      i++;
    }

    return `${value.toFixed(i ? 2 : 0)} ${units[i]}`;
  }

  function escapeIdent(name) {
    return '"' +
      String(name).replaceAll('"', '""') +
      '"';
  }

  function escapeString(value) {
    return "'" +
      String(value).replaceAll("'", "''") +
      "'";
  }

  async function loadDuckDB() {
    if (db && conn) {
      return;
    }

    setStatus(
      "1/5 — Cargando módulo DuckDB-Wasm..."
    );

    duckdb =
      await import(DUCKDB_PACKAGE);

    setStatus(
      "2/5 — Módulo cargado.\n" +
      "Seleccionando bundle..."
    );

    const bundles =
      duckdb.getJsDelivrBundles();

    const bundle =
      await duckdb.selectBundle(bundles);

    setStatus(
      "3/5 — Inicializando WebAssembly..."
    );

    const workerURL =
      URL.createObjectURL(
        new Blob(
          [
            `importScripts("${bundle.mainWorker}");`
          ],
          {
            type: "text/javascript"
          }
        )
      );

    const worker =
      new Worker(workerURL);

    const logger =
      new duckdb.ConsoleLogger();

    db =
      new duckdb.AsyncDuckDB(
        logger,
        worker
      );

    await db.instantiate(
      bundle.mainModule,
      bundle.pthreadWorker
    );

    URL.revokeObjectURL(workerURL);

    conn =
      await db.connect();

    setStatus(
      "4/5 — DuckDB listo.\n" +
      "Preparando extensión Excel..."
    );

    /*
     * Excel is an officially available DuckDB-Wasm extension.
     *
     * In Wasm, extension loading can happen automatically.
     * We try LOAD first so failures are explicit.
     */
    try {
      await conn.query("LOAD excel;");
    } catch (loadError) {

      console.warn(
        "[TM Excel Engine] LOAD excel failed; " +
        "continuing because DuckDB-Wasm supports autoloading.",
        loadError
      );
    }

    setStatus(
      "5/5 — DuckDB preparado.",
      "ok"
    );
  }

  async function handleFile(event) {
    const file =
      event.target.files?.[0];

    if (!file) return;

    currentFile = file;
    currentFileName =
      `tmxe_${Date.now()}_${file.name}`;

    try {

      setResult("Cargando archivo...");

      const meta =
        document.getElementById(
          "tmxe-v04-meta"
        );

      meta.textContent =
        `${file.name} — ${formatBytes(file.size)}`;

      await loadDuckDB();

      setStatus(
        "Registrando archivo XLSX en memoria local..."
      );

      const bytes =
        new Uint8Array(
          await file.arrayBuffer()
        );

      /*
       * IMPORTANT:
       * The file is registered directly in DuckDB-Wasm.
       * It is not uploaded to a remote server.
       */
      await db.registerFileBuffer(
        currentFileName,
        bytes
      );

      setStatus(
        "Leyendo la primera hoja del Excel..."
      );

      /*
       * Create a VIEW over the XLSX.
       *
       * header = true means the first row
       * is interpreted as column names.
       */
      await conn.query(`
        CREATE OR REPLACE VIEW
        ${escapeIdent(currentTable)}
        AS
        SELECT *
        FROM read_xlsx(
          ${escapeString(currentFileName)},
          header = true
        );
      `);

      /*
       * Basic metadata.
       */
      const countResult =
        await conn.query(`
          SELECT COUNT(*) AS filas
          FROM ${escapeIdent(currentTable)}
        `);

      const countRow =
        countResult.toArray()[0];

      const filas =
        countRow?.filas != null
          ? String(countRow.filas)
          : "?";

      const describeResult =
        await conn.query(`
          DESCRIBE ${escapeIdent(currentTable)}
        `);

      const columns =
        describeResult
          .toArray()
          .map(row => ({
            nombre: row.column_name,
            tipo: row.column_type
          }));

      meta.textContent =
        `${file.name} — ${formatBytes(file.size)} — ` +
        `${filas} filas — ` +
        `${columns.length} columnas`;

      setResult({
        archivo: file.name,
        tamaño: formatBytes(file.size),
        filas: filas,
        columnas: columns
      });

      setStatus(
        "✅ Excel cargado y procesado localmente.",
        "ok"
      );

    } catch (error) {

      console.error(
        `[${APP_ID}]`,
        error
      );

      setStatus(
        "❌ ERROR AL CARGAR EL EXCEL\n\n" +
        (
          error?.message ||
          String(error)
        ),
        "error"
      );

      setResult(
        error?.stack ||
        String(error)
      );
    }
  }

  async function countRows() {

    if (!conn || !currentFile) {

      setStatus(
        "Primero selecciona un archivo XLSX.",
        "error"
      );

      return;
    }

    try {

      setStatus(
        "Ejecutando COUNT(*) localmente..."
      );

      const start =
        performance.now();

      const result =
        await conn.query(`
          SELECT
            COUNT(*) AS filas
          FROM ${escapeIdent(currentTable)}
        `);

      const elapsed =
        performance.now() - start;

      const row =
        result.toArray()[0];

      const filas =
        row?.filas != null
          ? String(row.filas)
          : "?";

      setResult({
        consulta: "COUNT(*)",
        filas: filas,
        tiempo_ms:
          Number(elapsed.toFixed(2))
      });

      setStatus(
        `✅ COUNT(*) completado en ${elapsed.toFixed(0)} ms.`,
        "ok"
      );

    } catch (error) {

      console.error(error);

      setStatus(
        "❌ Error ejecutando COUNT(*)",
        "error"
      );

      setResult(
        error?.stack ||
        String(error)
      );
    }
  }

  async function previewRows() {

    if (!conn || !currentFile) {

      setStatus(
        "Primero selecciona un archivo XLSX.",
        "error"
      );

      return;
    }

    try {

      setStatus(
        "Leyendo las primeras 5 filas..."
      );

      const result =
        await conn.query(`
          SELECT *
          FROM ${escapeIdent(currentTable)}
          LIMIT 5
        `);

      const rows =
        result.toArray().map(row => {

          const output = {};

          for (
            const [key, value]
            of Object.entries(row)
          ) {

            output[key] =
              typeof value === "bigint"
                ? value.toString()
                : value;
          }

          return output;
        });

      setResult(rows);

      setStatus(
        "✅ Vista previa generada.",
        "ok"
      );

    } catch (error) {

      console.error(error);

      setStatus(
        "❌ Error obteniendo vista previa.",
        "error"
      );

      setResult(
        error?.stack ||
        String(error)
      );
    }
  }

  async function showColumns() {

    if (!conn || !currentFile) {

      setStatus(
        "Primero selecciona un archivo XLSX.",
        "error"
      );

      return;
    }

    try {

      const result =
        await conn.query(`
          DESCRIBE ${escapeIdent(currentTable)}
        `);

      const columns =
        result
          .toArray()
          .map(row => ({
            columna: row.column_name,
            tipo: row.column_type
          }));

      setResult(columns);

      setStatus(
        `✅ ${columns.length} columnas detectadas.`,
        "ok"
      );

    } catch (error) {

      console.error(error);

      setStatus(
        "❌ Error obteniendo columnas.",
        "error"
      );

      setResult(
        error?.stack ||
        String(error)
      );
    }
  }

  function init() {

    addStyles();

    createButton();

    console.log(
      `[${APP_ID}] cargado`
    );
  }

  if (
    document.readyState ===
    "loading"
  ) {

    document.addEventListener(
      "DOMContentLoaded",
      init,
      { once: true }
    );

  } else {

    init();
  }

})();
