/*
 * TypingMind Excel Data Engine - Extension v0.4.3 TEST
 *
 * DEEP DIAGNOSTIC VERSION
 *
 * Purpose:
 * - Confirm DuckDB-Wasm 1.29.0 loads.
 * - Identify the embedded DuckDB version.
 * - Inspect DuckDB extension configuration.
 * - Inspect the Excel extension.
 * - Test LOAD excel.
 *
 * IMPORTANT:
 * - No XLSX file is loaded.
 * - Diagnostic only.
 */

(() => {
  "use strict";

  const APP_ID = "tm-excel-engine-v043-test";

  const DUCKDB_PACKAGE =
    "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

  let duckdb = null;
  let db = null;
  let conn = null;
  let worker = null;
  let workerURL = null;

  function addStyles() {
    if (document.getElementById("tmxe-v043-style")) return;

    const style = document.createElement("style");

    style.id = "tmxe-v043-style";

    style.textContent = `
      #tmxe-v043-button {
        position: fixed;
        right: 18px;
        bottom: 86px;
        z-index: 2147483000;
        border: 0;
        border-radius: 999px;
        padding: 11px 16px;
        background: #0891b2;
        color: white;
        font: 600 14px/1.1 system-ui, sans-serif;
        box-shadow: 0 6px 20px rgba(0,0,0,.25);
        cursor: pointer;
      }

      #tmxe-v043-overlay {
        position: fixed;
        inset: 0;
        z-index: 2147483001;
        background: rgba(0,0,0,.48);
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 16px;
      }

      #tmxe-v043-panel {
        width: min(1000px, 96vw);
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

      #tmxe-v043-panel h2 {
        margin: 0 0 5px;
        font-size: 20px;
      }

      #tmxe-v043-help {
        opacity: .7;
        margin-bottom: 15px;
      }

      #tmxe-v043-status {
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

      #tmxe-v043-status[data-kind="ok"] {
        color: #16a34a;
      }

      #tmxe-v043-status[data-kind="error"] {
        color: #dc2626;
      }

      #tmxe-v043-result {
        white-space: pre-wrap;
        overflow: auto;
        max-height: 600px;
        padding: 12px;
        border-radius: 10px;
        background: rgba(127,127,127,.10);
        font: 12px/1.45 ui-monospace,
          SFMono-Regular,
          Menlo,
          monospace;
      }

      #tmxe-v043-run,
      #tmxe-v043-close {
        border: 1px solid rgba(127,127,127,.45);
        border-radius: 9px;
        padding: 9px 13px;
        background: transparent;
        color: inherit;
        cursor: pointer;
      }

      #tmxe-v043-run {
        background: #0891b2 !important;
        color: white !important;
        border-color: #0891b2 !important;
      }

      #tmxe-v043-actions {
        display: flex;
        gap: 8px;
        flex-wrap: wrap;
        margin: 10px 0;
      }
    `;

    document.head.appendChild(style);
  }

  function createButton() {
    if (document.getElementById("tmxe-v043-button")) return;

    const button = document.createElement("button");

    button.id = "tmxe-v043-button";
    button.type = "button";
    button.textContent = "🧪 Excel v0.4.3";
    button.title = "Diagnóstico profundo de DuckDB";

    button.addEventListener("click", openPanel);

    document.body.appendChild(button);
  }

  function openPanel() {
    if (document.getElementById("tmxe-v043-overlay")) {
      return;
    }

    const overlay = document.createElement("div");

    overlay.id = "tmxe-v043-overlay";

    overlay.innerHTML = `
      <div
        id="tmxe-v043-panel"
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

            <h2>
              🧪 Excel Engine v0.4.3 TEST
            </h2>

            <div id="tmxe-v043-help">
              Diagnóstico profundo de DuckDB-Wasm 1.29.0.
              No se carga ningún Excel.
            </div>

          </div>

          <button
            id="tmxe-v043-close"
            type="button"
          >
            Cerrar
          </button>

        </div>

        <div id="tmxe-v043-status">
          Listo para iniciar diagnóstico.
        </div>

        <div id="tmxe-v043-actions">

          <button
            id="tmxe-v043-run"
            type="button"
          >
            🔍 Ejecutar diagnóstico
          </button>

        </div>

        <strong>Resultado del diagnóstico</strong>

        <pre id="tmxe-v043-result">—</pre>

        <div style="
          margin-top:12px;
          opacity:.65;
          font-size:12px;
        ">
          v0.4.3 TEST — Deep Diagnostic.
        </div>

      </div>
    `;

    document.body.appendChild(overlay);

    document
      .getElementById("tmxe-v043-close")
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
      .getElementById("tmxe-v043-run")
      .addEventListener(
        "click",
        runDiagnostic
      );
  }

  function setStatus(text, kind = "") {
    const el =
      document.getElementById("tmxe-v043-status");

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
      document.getElementById("tmxe-v043-result");

    if (!el) return;

    if (typeof value === "string") {
      el.textContent = value;
    } else {
      el.textContent =
        JSON.stringify(value, null, 2);
    }
  }

  async function safeQuery(sql) {
    try {
      const result = await conn.query(sql);

      return {
        ok: true,
        rows: result.toArray()
      };

    } catch (error) {

      return {
        ok: false,
        error:
          error?.message ||
          String(error)
      };
    }
  }

  async function runDiagnostic() {

    try {

      setResult("");

      /*
       * 1
       */
      setStatus(
        "1/9 — Cargando DuckDB-Wasm 1.29.0..."
      );

      duckdb =
        await import(DUCKDB_PACKAGE);

      /*
       * 2
       */
      setStatus(
        "2/9 — Inspeccionando bundle..."
      );

      const bundles =
        duckdb.getJsDelivrBundles();

      const bundle =
        await duckdb.selectBundle(bundles);

      /*
       * 3
       */
      setStatus(
        "3/9 — Inicializando WebAssembly..."
      );

      workerURL =
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

      worker =
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

      /*
       * 4
       */
      setStatus(
        "4/9 — Conectando DuckDB..."
      );

      conn =
        await db.connect();

      /*
       * 5
       */
      setStatus(
        "5/9 — Consultando versión de DuckDB..."
      );

      const version =
        await safeQuery(
          "SELECT version() AS version;"
        );

      /*
       * 6
       */
      setStatus(
        "6/9 — Consultando configuración..."
      );

      const settings =
        await safeQuery(`
          SELECT
            name,
            value
          FROM duckdb_settings()
          WHERE
            name ILIKE '%extension%'
            OR name ILIKE '%repository%'
          ORDER BY name;
        `);

      /*
       * 7
       */
      setStatus(
        "7/9 — Consultando extensiones..."
      );

      const extensions =
        await safeQuery(`
          SELECT
            extension_name,
            loaded,
            installed,
            install_mode,
            installed_from
          FROM duckdb_extensions()
          ORDER BY extension_name;
        `);

      /*
       * 8
       */
      setStatus(
        "8/9 — Probando LOAD excel..."
      );

      const loadExcel =
        await safeQuery(
          "LOAD excel;"
        );

      /*
       * 9
       */
      setStatus(
        "9/9 — Diagnóstico terminado.",
        loadExcel.ok
          ? "ok"
          : "error"
      );

      const excelAfterLoad =
        await safeQuery(`
          SELECT
            extension_name,
            loaded,
            installed,
            install_mode,
            installed_from
          FROM duckdb_extensions()
          WHERE extension_name = 'excel';
        `);

      setResult({

        app_id:
          APP_ID,

        duckdb_wasm:
          "CARGADO",

        duckdb_wasm_package:
          "1.29.0",

        bundle:
          {
            mainModule:
              bundle.mainModule,

            mainWorker:
              bundle.mainWorker,

            pthreadWorker:
              bundle.pthreadWorker
          },

        conexion:
          "CONECTADA",

        version_duckdb:
          version,

        configuracion_extensiones:
          settings,

        extensiones_antes_load:
          extensions,

        load_excel:
          loadExcel,

        extension_excel_despues_load:
          excelAfterLoad

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
