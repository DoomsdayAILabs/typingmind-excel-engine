/*
 * TypingMind Excel Data Engine - Extension v0.4.2 TEST
 *
 * DIAGNOSTIC VERSION
 *
 * Purpose:
 * - Confirm DuckDB-Wasm modern version loads correctly.
 * - Test whether the "excel" extension can be loaded.
 * - Show the exact LOAD excel error.
 * - Inspect duckdb_extensions().
 *
 * IMPORTANT:
 * - This version DOES NOT load an XLSX file.
 * - It is only a diagnostic test.
 */

(() => {
  "use strict";

  const APP_ID = "tm-excel-engine-v042-test";

  /*
   * v0.4.2:
   * Upgrade from DuckDB-Wasm 1.29.0 to current stable 1.5.4.
   */
  const DUCKDB_PACKAGE =
    "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.5.4/+esm";

  let duckdb = null;
  let db = null;
  let conn = null;

  function addStyles() {
    if (document.getElementById("tmxe-v042-style")) return;

    const style = document.createElement("style");

    style.id = "tmxe-v042-style";

    style.textContent = `
      #tmxe-v042-button {
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

      #tmxe-v042-overlay {
        position: fixed;
        inset: 0;
        z-index: 2147483001;
        background: rgba(0,0,0,.48);
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 16px;
      }

      #tmxe-v042-panel {
        width: min(900px, 96vw);
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

      #tmxe-v042-panel h2 {
        margin: 0 0 5px;
        font-size: 20px;
      }

      #tmxe-v042-help {
        opacity: .7;
        margin-bottom: 15px;
      }

      #tmxe-v042-status {
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

      #tmxe-v042-status[data-kind="ok"] {
        color: #16a34a;
      }

      #tmxe-v042-status[data-kind="error"] {
        color: #dc2626;
      }

      #tmxe-v042-result {
        white-space: pre-wrap;
        overflow: auto;
        max-height: 500px;
        padding: 12px;
        border-radius: 10px;
        background: rgba(127,127,127,.10);
        font: 12px/1.45 ui-monospace,
          SFMono-Regular,
          Menlo,
          monospace;
      }

      #tmxe-v042-run,
      #tmxe-v042-close {
        border: 1px solid rgba(127,127,127,.45);
        border-radius: 9px;
        padding: 9px 13px;
        background: transparent;
        color: inherit;
        cursor: pointer;
      }

      #tmxe-v042-run {
        background: #2563eb !important;
        color: white !important;
        border-color: #2563eb !important;
      }

      #tmxe-v042-actions {
        display: flex;
        gap: 8px;
        flex-wrap: wrap;
        margin: 10px 0;
      }
    `;

    document.head.appendChild(style);
  }

  function createButton() {
    if (document.getElementById("tmxe-v042-button")) return;

    const button = document.createElement("button");

    button.id = "tmxe-v042-button";
    button.type = "button";
    button.textContent = "🧪 Excel v0.4.2";
    button.title = "Diagnóstico de extensión Excel";

    button.addEventListener("click", openPanel);

    document.body.appendChild(button);
  }

  function openPanel() {
    if (document.getElementById("tmxe-v042-overlay")) {
      return;
    }

    const overlay = document.createElement("div");

    overlay.id = "tmxe-v042-overlay";

    overlay.innerHTML = `
      <div
        id="tmxe-v042-panel"
        role="dialog"
        aria-modal="true"
      >

        <div style="
          display:flex;
          justify-content:space-between;
          align-items:flex-start;
          gap:12px;
        ">

          <div>
            <h2>🧪 Excel Engine v0.4.2 TEST</h2>

            <div id="tmxe-v042-help">
              Diagnóstico de DuckDB-Wasm y extensión Excel.
              DuckDB-Wasm 1.5.4.
            </div>
          </div>

          <button
            id="tmxe-v042-close"
            type="button"
          >
            Cerrar
          </button>

        </div>

        <div id="tmxe-v042-status">
          Listo para iniciar diagnóstico.
        </div>

        <div id="tmxe-v042-actions">

          <button
            id="tmxe-v042-run"
            type="button"
          >
            🔍 Ejecutar diagnóstico
          </button>

        </div>

        <strong>Resultado del diagnóstico</strong>

        <pre id="tmxe-v042-result">—</pre>

        <div style="
          margin-top:12px;
          opacity:.65;
          font-size:12px;
        ">
          v0.4.2 TEST — Solo diagnóstico.
          No se carga ningún archivo.
        </div>

      </div>
    `;

    document.body.appendChild(overlay);

    document
      .getElementById("tmxe-v042-close")
      .addEventListener(
        "click",
        () => overlay.remove()
      );

    overlay.addEventListener(
      "click",
      event => {
        if (event.target === overlay) {
          overlay.remove();
        }
      }
    );

    document
      .getElementById("tmxe-v042-run")
      .addEventListener(
        "click",
        runDiagnostic
      );
  }

  function setStatus(text, kind = "") {
    const el =
      document.getElementById("tmxe-v042-status");

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
      document.getElementById("tmxe-v042-result");

    if (!el) return;

    if (typeof value === "string") {
      el.textContent = value;
    } else {
      el.textContent =
        JSON.stringify(value, null, 2);
    }
  }

  async function runDiagnostic() {
    try {
      setResult("");

      setStatus(
        "1/6 — Cargando DuckDB-Wasm 1.5.4..."
      );

      duckdb = await import(DUCKDB_PACKAGE);

      setStatus(
        "2/6 — Seleccionando bundle..."
      );

      const bundles =
        duckdb.getJsDelivrBundles();

      const bundle =
        await duckdb.selectBundle(bundles);

      setStatus(
        "3/6 — Inicializando WebAssembly..."
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

      conn = await db.connect();

      setStatus(
        "4/6 — DuckDB conectado.\n" +
        "Probando LOAD excel..."
      );

      let loadResult = null;
      let loadError = null;

      try {
        loadResult =
          await conn.query("LOAD excel;");
      } catch (error) {
        loadError =
          error?.message ||
          String(error);
      }

      setStatus(
        "5/6 — Consultando estado de extensiones..."
      );

      let extensions = null;
      let extensionQueryError = null;

      try {
        const result =
          await conn.query(`
            SELECT
              extension_name,
              loaded,
              installed
            FROM duckdb_extensions()
            ORDER BY extension_name
          `);

        extensions =
          result.toArray();

      } catch (error) {
        extensionQueryError =
          error?.message ||
          String(error);
      }

      setStatus(
        "6/6 — Diagnóstico terminado.",
        loadError ? "error" : "ok"
      );

      setResult({
        duckdb_wasm:
          "CARGADO",

        duckdb_wasm_package:
          "1.5.4",

        conexion:
          "CONECTADA",

        load_excel:
          loadError
            ? "ERROR"
            : "OK",

        load_excel_error:
          loadError,

        load_excel_result:
          loadResult
            ? "Consulta ejecutada"
            : null,

        extension_excel:
          extensions
            ? extensions.filter(
                item =>
                  item.extension_name === "excel"
              )
            : null,

        todas_las_extensiones:
          extensions,

        error_consulta_extensiones:
          extensionQueryError
      });

    } catch (error) {

      console.error(
        `[${APP_ID}]`,
        error
      );

      setStatus(
        "❌ ERROR DURANTE EL DIAGNÓSTICO",
        "error"
      );

      setResult({
        mensaje:
          error?.message ||
          String(error),

        stack:
          error?.stack ||
          null
      });
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
    document.readyState === "loading"
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
