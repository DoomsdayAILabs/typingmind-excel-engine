/*
 * TypingMind Excel Data Engine - Extension v0.4.4 TEST
 *
 * XLSX LOCAL -> SHEETJS -> DUCKDB-WASM
 *
 * IMPORTANT:
 * - No DuckDB Excel extension.
 * - No server upload.
 * - No LLM.
 * - File remains in browser.
 */

(() => {
  "use strict";

  const APP_ID = "tm-excel-engine-v044-test";

  const DUCKDB_PACKAGE =
    "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

  const SHEETJS_URL =
    "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js";

  let duckdb = null;
  let db = null;
  let conn = null;

  function addStyles() {
    if (document.getElementById("tmxe-v044-style")) return;

    const style = document.createElement("style");

    style.id = "tmxe-v044-style";

    style.textContent = `
      #tmxe-v044-button {
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

      #tmxe-v044-overlay {
        position: fixed;
        inset: 0;
        z-index: 2147483001;
        background: rgba(0,0,0,.48);
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 16px;
      }

      #tmxe-v044-panel {
        width: min(1050px, 96vw);
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

      #tmxe-v044-status {
        padding: 12px;
        border-radius: 9px;
        background: rgba(127,127,127,.10);
        white-space: pre-wrap;
        font: 13px/1.5 ui-monospace, monospace;
        margin: 12px 0;
      }

      #tmxe-v044-status[data-kind="ok"] {
        color: #16a34a;
      }

      #tmxe-v044-status[data-kind="error"] {
        color: #dc2626;
      }

      #tmxe-v044-result {
        white-space: pre-wrap;
        overflow: auto;
        max-height: 550px;
        padding: 12px;
        border-radius: 10px;
        background: rgba(127,127,127,.10);
        font: 12px/1.45 ui-monospace, monospace;
      }

      #tmxe-v044-file {
        margin: 10px 0;
      }

      #tmxe-v044-run,
      #tmxe-v044-close {
        border: 1px solid rgba(127,127,127,.45);
        border-radius: 9px;
        padding: 9px 13px;
        background: transparent;
        color: inherit;
        cursor: pointer;
      }

      #tmxe-v044-run {
        background: #2563eb !important;
        color: white !important;
        border-color: #2563eb !important;
      }

      #tmxe-v044-actions {
        display: flex;
        gap: 8px;
        flex-wrap: wrap;
        margin: 10px 0;
      }
    `;

    document.head.appendChild(style);
  }

  function createButton() {
    if (document.getElementById("tmxe-v044-button")) return;

    const button = document.createElement("button");

    button.id = "tmxe-v044-button";
    button.type = "button";
    button.textContent = "📊 Excel v0.4.4";
    button.title = "Excel local -> DuckDB";

    button.addEventListener("click", openPanel);

    document.body.appendChild(button);
  }

  function openPanel() {
    if (document.getElementById("tmxe-v044-overlay")) {
      return;
    }

    const overlay = document.createElement("div");

    overlay.id = "tmxe-v044-overlay";

    overlay.innerHTML = `
      <div id="tmxe-v044-panel">

        <div style="
          display:flex;
          justify-content:space-between;
          align-items:flex-start;
          gap:12px;
        ">

          <div>
            <h2>📊 Excel Engine v0.4.4 TEST</h2>

            <div style="opacity:.7">
              XLSX local → SheetJS → DuckDB-Wasm
            </div>
          </div>

          <button id="tmxe-v044-close">
            Cerrar
          </button>

        </div>

        <input
          id="tmxe-v044-file"
          type="file"
          accept=".xlsx,.xls,.xlsb,.ods"
        />

        <div id="tmxe-v044-actions">

          <button id="tmxe-v044-run">
            🚀 Cargar Excel
          </button>

        </div>

        <div id="tmxe-v044-status">
          Selecciona un archivo Excel.
        </div>

        <strong>Resultado</strong>

        <pre id="tmxe-v044-result">—</pre>

      </div>
    `;

    document.body.appendChild(overlay);

    document
      .getElementById("tmxe-v044-close")
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
      .getElementById("tmxe-v044-run")
      .addEventListener(
        "click",
        loadExcel
      );
  }

  function setStatus(text, kind = "") {

    const el =
      document.getElementById(
        "tmxe-v044-status"
      );

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
      document.getElementById(
        "tmxe-v044-result"
      );

    if (!el) return;

    el.textContent =
      typeof value === "string"
        ? value
        : JSON.stringify(
            value,
            null,
            2
          );
  }

  async function loadSheetJS() {

    if (window.XLSX) {
      return window.XLSX;
    }

    return new Promise(
      (resolve, reject) => {

        const script =
          document.createElement("script");

        script.src = SHEETJS_URL;

        script.onload = () => {

          if (!window.XLSX) {
            reject(
              new Error(
                "SheetJS cargó pero XLSX no está disponible."
              )
            );

            return;
          }

          resolve(window.XLSX);
        };

        script.onerror = () => {

          reject(
            new Error(
              "No se pudo cargar SheetJS."
            )
          );

        };

        document.head.appendChild(script);

      }
    );
  }

  async function initDuckDB() {

    setStatus(
      "Inicializando DuckDB-Wasm..."
    );

    duckdb =
      await import(
        DUCKDB_PACKAGE
      );

    const bundles =
      duckdb.getJsDelivrBundles();

    const bundle =
      await duckdb.selectBundle(
        bundles
      );

    const workerURL =
      URL.createObjectURL(
        new Blob(
          [
            `importScripts("${bundle.mainWorker}");`
          ],
          {
            type:
              "text/javascript"
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

    URL.revokeObjectURL(
      workerURL
    );

    conn =
      await db.connect();

    return bundle;
  }

  async function loadExcel() {

    try {

      setResult("");

      const fileInput =
        document.getElementById(
          "tmxe-v044-file"
        );

      const file =
        fileInput?.files?.[0];

      if (!file) {

        setStatus(
          "❌ Primero selecciona un archivo Excel.",
          "error"
        );

        return;
      }

      /*
       * 1
       */

      setStatus(
        "1/7 — Cargando parser XLSX..."
      );

      const XLSX =
        await loadSheetJS();

      /*
       * 2
       */

      setStatus(
        "2/7 — Leyendo archivo local..."
      );

      const arrayBuffer =
        await file.arrayBuffer();

      /*
       * 3
       */

      setStatus(
        "3/7 — Analizando XLSX..."
      );

      const workbook =
        XLSX.read(
          arrayBuffer,
          {
            type: "array",
            dense: true
          }
        );

      /*
       * 4
       */

      const sheetNames =
        workbook.SheetNames;

      if (!sheetNames.length) {

        throw new Error(
          "El archivo no contiene hojas."
        );

      }

      const firstSheetName =
        sheetNames[0];

      const worksheet =
        workbook.Sheets[
          firstSheetName
        ];

      const rows =
        XLSX.utils.sheet_to_json(
          worksheet,
          {
            header: 1,
            defval: null,
            raw: true
          }
        );

      if (!rows.length) {

        throw new Error(
          "La primera hoja está vacía."
        );

      }

      /*
       * 5
       */

      setStatus(
        "4/7 — Preparando datos para DuckDB..."
      );

      const headers =
        rows[0].map(
          (value, index) => {

            const text =
              String(
                value ??
                `column_${index + 1}`
              ).trim();

            return text ||
              `column_${index + 1}`;

          }
        );

      const dataRows =
        rows.slice(1);

      /*
       * 6
       */

      setStatus(
        "5/7 — Inicializando DuckDB-Wasm..."
      );

      await initDuckDB();

      /*
       * Crear tabla usando VALUES.
       *
       * Esta primera versión es deliberadamente
       * simple para validar todo el flujo.
       */

      setStatus(
        "6/7 — Insertando datos en DuckDB..."
      );

      const safeIdentifier =
        name =>
          `"${String(name)
            .replace(/"/g, '""')}"`;

      const uniqueHeaders = [];

      const usedHeaders =
        new Set();

      headers.forEach(
        (header, index) => {

          let name = header;

          if (usedHeaders.has(name)) {

            let counter = 2;

            while (
              usedHeaders.has(
                `${name}_${counter}`
              )
            ) {
              counter++;
            }

            name =
              `${name}_${counter}`;
          }

          usedHeaders.add(name);

          uniqueHeaders.push(name);

        }
      );

      const columnsSQL =
        uniqueHeaders
          .map(
            name =>
              `${safeIdentifier(name)} VARCHAR`
          )
          .join(", ");

      await conn.query(
        "DROP TABLE IF EXISTS excel_data;"
      );

      await conn.query(
        `CREATE TABLE excel_data (${columnsSQL});`
      );

      /*
       * Insertar filas.
       *
       * Solo para esta prueba.
       */

      for (
        let start = 0;
        start < dataRows.length;
        start += 500
      ) {

        const batch =
          dataRows.slice(
            start,
            start + 500
          );

        const values =
          batch.map(
            row => {

              const cells =
                uniqueHeaders.map(
                  (_, index) => {

                    const value =
                      row[index];

                    if (
                      value === null ||
                      value === undefined ||
                      value === ""
                    ) {
                      return "NULL";
                    }

                    const text =
                      String(value)
                        .replace(
                          /'/g,
                          "''"
                        );

                    return `'${text}'`;

                  }
                );

              return `(${cells.join(",")})`;

            }
          );

        if (values.length) {

          await conn.query(
            `INSERT INTO excel_data VALUES
             ${values.join(",")};`
          );

        }

      }

      /*
       * 7
       */

      setStatus(
        "7/7 — Ejecutando consultas SQL..."
      );

      const countResult =
        await conn.query(
          `
          SELECT COUNT(*) AS total_rows
          FROM excel_data;
          `
        );

      const previewResult =
        await conn.query(
          `
          SELECT *
          FROM excel_data
          LIMIT 10;
          `
        );

      const columnsResult =
        await conn.query(
          `
          DESCRIBE excel_data;
          `
        );

      setStatus(
        "✅ Excel cargado y consultado localmente.",
        "ok"
      );

      setResult({

        archivo:
          file.name,

        tamano_bytes:
          file.size,

        hoja:
          firstSheetName,

        hojas:
          sheetNames,

        filas_detectadas:
          dataRows.length,

        columnas_detectadas:
          uniqueHeaders.length,

        encabezados:
          uniqueHeaders,

        duckdb_table:
          "excel_data",

        count_sql:
          countResult.toArray(),

        schema:
          columnsResult.toArray(),

        preview:
          previewResult.toArray()

      });

    } catch (error) {

      console.error(
        `[${APP_ID}]`,
        error
      );

      setStatus(
        "❌ ERROR",
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
