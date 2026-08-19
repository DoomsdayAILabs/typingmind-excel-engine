/*
 * TypingMind Excel Engine - v0.3 TEST
 * DuckDB-Wasm loading test only.
 *
 * This version DOES NOT load Excel.
 * It only verifies that DuckDB-Wasm can be loaded
 * and execute a SQL query inside TypingMind.
 */

(() => {
  "use strict";

  const APP_ID = "tm-excel-engine-v03-test";

  // IMPORTANT:
  // This is the package version, not the DuckDB engine version.
  const DUCKDB_PACKAGE =
    "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

  let duckdb = null;
  let db = null;
  let conn = null;

  function addStyles() {
    if (document.getElementById("tmxe-v03-style")) return;

    const style = document.createElement("style");
    style.id = "tmxe-v03-style";

    style.textContent = `
      #tmxe-v03-button {
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

      #tmxe-v03-overlay {
        position: fixed;
        inset: 0;

        z-index: 2147483001;

        background: rgba(0,0,0,.48);

        display: flex;
        align-items: center;
        justify-content: center;

        padding: 16px;
      }

      #tmxe-v03-panel {
        width: min(700px, 95vw);
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

      #tmxe-v03-panel h2 {
        margin: 0 0 6px;
      }

      #tmxe-v03-status {
        margin: 15px 0;

        padding: 12px;

        border-radius: 10px;

        background: rgba(127,127,127,.10);

        white-space: pre-wrap;

        font: 13px/1.5 ui-monospace,
              SFMono-Regular,
              Menlo,
              monospace;
      }

      #tmxe-v03-status[data-kind="ok"] {
        color: #16a34a;
      }

      #tmxe-v03-status[data-kind="error"] {
        color: #dc2626;
      }

      #tmxe-v03-close,
      #tmxe-v03-run {
        border: 1px solid rgba(127,127,127,.45);

        border-radius: 9px;

        padding: 9px 13px;

        background: transparent;
        color: inherit;

        cursor: pointer;
      }

      #tmxe-v03-run {
        background: #2563eb;
        color: white;
        border-color: #2563eb;
      }

      #tmxe-v03-actions {
        display: flex;
        gap: 8px;
        flex-wrap: wrap;
      }
    `;

    document.head.appendChild(style);
  }

  function createButton() {
    if (document.getElementById("tmxe-v03-button")) return;

    const button = document.createElement("button");

    button.id = "tmxe-v03-button";
    button.type = "button";

    button.textContent = "🧪 DuckDB Test";

    button.title = "Probar DuckDB-Wasm";

    button.addEventListener("click", openPanel);

    document.body.appendChild(button);
  }

  function openPanel() {
    if (document.getElementById("tmxe-v03-overlay")) return;

    const overlay = document.createElement("div");

    overlay.id = "tmxe-v03-overlay";

    overlay.innerHTML = `
      <div id="tmxe-v03-panel">

        <div style="
          display:flex;
          justify-content:space-between;
          align-items:center;
          gap:12px;
        ">

          <div>
            <h2>🧪 DuckDB-Wasm Test</h2>

            <div style="opacity:.7">
              v0.3 — prueba de carga de DuckDB
            </div>
          </div>

          <button id="tmxe-v03-close">
            Cerrar
          </button>

        </div>

        <div id="tmxe-v03-status">
          Listo para probar DuckDB-Wasm.
        </div>

        <div id="tmxe-v03-actions">

          <button id="tmxe-v03-run">
            ▶ Cargar DuckDB y ejecutar SELECT 42
          </button>

        </div>

      </div>
    `;

    document.body.appendChild(overlay);

    document
      .getElementById("tmxe-v03-close")
      .addEventListener("click", () => {
        overlay.remove();
      });

    overlay.addEventListener("click", event => {
      if (event.target === overlay) {
        overlay.remove();
      }
    });

    document
      .getElementById("tmxe-v03-run")
      .addEventListener("click", runTest);
  }

  function setStatus(text, kind = "") {
    const el = document.getElementById("tmxe-v03-status");

    if (!el) return;

    el.textContent = text;

    if (kind) {
      el.dataset.kind = kind;
    } else {
      delete el.dataset.kind;
    }
  }

  async function runTest() {
    const button = document.getElementById("tmxe-v03-run");

    try {
      button.disabled = true;

      setStatus(
        "1/4 — Cargando módulo DuckDB-Wasm..."
      );

      duckdb = await import(DUCKDB_PACKAGE);

      setStatus(
        "2/4 — Módulo cargado correctamente.\n" +
        "Preparando bundle..."
      );

      const bundles = duckdb.getJsDelivrBundles();

      const bundle =
        await duckdb.selectBundle(bundles);

      setStatus(
        "3/4 — Bundle seleccionado.\n" +
        "Inicializando DuckDB..."
      );

      const workerURL = URL.createObjectURL(
        new Blob(
          [
            `importScripts("${bundle.mainWorker}");`
          ],
          {
            type: "text/javascript"
          }
        )
      );

      const worker = new Worker(workerURL);

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
        "4/4 — Ejecutando SQL..."
      );

      const result =
        await conn.query(`
          SELECT
            42 AS prueba,
            'DuckDB funcionando' AS mensaje
        `);

      const rows =
        result.toArray().map(row => {

          const output = {};

          for (const [key, value]
            of Object.entries(row)) {

            output[key] =
              typeof value === "bigint"
                ? value.toString()
                : value;
          }

          return output;
        });

      setStatus(
        "✅ PRUEBA EXITOSA\n\n" +
        JSON.stringify(rows, null, 2),
        "ok"
      );

    } catch (error) {

      console.error(
        "[TM Excel Engine v0.3]",
        error
      );

      setStatus(
        "❌ ERROR\n\n" +
        (error?.stack ||
         error?.message ||
         String(error)),
        "error"
      );

    } finally {

      button.disabled = false;
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
