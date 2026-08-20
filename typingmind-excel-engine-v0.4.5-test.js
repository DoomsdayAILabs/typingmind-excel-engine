/*
 * TypingMind Excel Data Engine - Extension v0.4.5 TEST
 *
 * XLSX LOCAL -> SHEETJS -> DUCKDB-WASM
 *
 * v0.4.5
 *
 * Changes from v0.4.4:
 * - Fix BigInt JSON serialization.
 * - Preserve local-only processing.
 * - No DuckDB Excel extension.
 * - No LLM communication.
 * - No file upload.
 */

(() => {
  "use strict";

  const APP_ID =
    "tm-excel-engine-v045-test";

  /*
   * DuckDB-Wasm version already proven
   * to load correctly in TypingMind.
   */
  const DUCKDB_PACKAGE =
    "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

  /*
   * SheetJS browser build.
   */
  const SHEETJS_URL =
    "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js";

  let duckdb = null;
  let db = null;
  let conn = null;

  /*
   * -------------------------------------------------------
   * STYLES
   * -------------------------------------------------------
   */

  function addStyles() {

    if (
      document.getElementById(
        "tmxe-v045-style"
      )
    ) {
      return;
    }

    const style =
      document.createElement("style");

    style.id =
      "tmxe-v045-style";

    style.textContent = `

      #tmxe-v045-button {

        position: fixed;

        right: 18px;

        bottom: 86px;

        z-index: 2147483000;

        border: 0;

        border-radius: 999px;

        padding: 11px 16px;

        background: #2563eb;

        color: white;

        font:
          600 14px/1.1
          system-ui,
          sans-serif;

        box-shadow:
          0 6px 20px
          rgba(0,0,0,.25);

        cursor: pointer;
      }

      #tmxe-v045-overlay {

        position: fixed;

        inset: 0;

        z-index: 2147483001;

        background:
          rgba(0,0,0,.48);

        display: flex;

        align-items: center;

        justify-content: center;

        padding: 16px;
      }

      #tmxe-v045-panel {

        width:
          min(1050px, 96vw);

        max-height: 90vh;

        overflow: auto;

        background: Canvas;

        color: CanvasText;

        border:
          1px solid
          rgba(127,127,127,.35);

        border-radius: 16px;

        padding: 20px;

        box-shadow:
          0 20px 70px
          rgba(0,0,0,.35);

        font:
          14px/1.45
          system-ui,
          sans-serif;
      }

      #tmxe-v045-panel h2 {

        margin:
          0 0 5px;

        font-size: 20px;
      }

      #tmxe-v045-help {

        opacity: .7;

        margin-bottom: 12px;
      }

      #tmxe-v045-file {

        display: block;

        width: 100%;

        margin:
          12px 0;

        padding: 8px;

        border:
          1px solid
          rgba(127,127,127,.35);

        border-radius: 8px;

        box-sizing: border-box;
      }

      #tmxe-v045-status {

        padding: 12px;

        border-radius: 9px;

        background:
          rgba(127,127,127,.10);

        white-space: pre-wrap;

        font:
          13px/1.5
          ui-monospace,
          SFMono-Regular,
          Menlo,
          monospace;

        margin:
          12px 0;
      }

      #tmxe-v045-status[data-kind="ok"] {

        color: #16a34a;
      }

      #tmxe-v045-status[data-kind="error"] {

        color: #dc2626;
      }

      #tmxe-v045-result {

        white-space: pre-wrap;

        overflow: auto;

        max-height: 550px;

        padding: 12px;

        border-radius: 10px;

        background:
          rgba(127,127,127,.10);

        font:
          12px/1.45
          ui-monospace,
          SFMono-Regular,
          Menlo,
          monospace;
      }

      #tmxe-v045-run,
      #tmxe-v045-close {

        border:
          1px solid
          rgba(127,127,127,.45);

        border-radius: 9px;

        padding:
          9px 13px;

        background: transparent;

        color: inherit;

        cursor: pointer;
      }

      #tmxe-v045-run {

        background:
          #2563eb !important;

        color:
          white !important;

        border-color:
          #2563eb !important;
      }

      #tmxe-v045-actions {

        display: flex;

        gap: 8px;

        flex-wrap: wrap;

        margin:
          10px 0;
      }

    `;

    document.head.appendChild(style);
  }

  /*
   * -------------------------------------------------------
   * BUTTON
   * -------------------------------------------------------
   */

  function createButton() {

    if (
      document.getElementById(
        "tmxe-v045-button"
      )
    ) {
      return;
    }

    const button =
      document.createElement("button");

    button.id =
      "tmxe-v045-button";

    button.type =
      "button";

    button.textContent =
      "📊 Excel v0.4.5";

    button.title =
      "Excel local → DuckDB";

    button.addEventListener(
      "click",
      openPanel
    );

    document.body.appendChild(
      button
    );
  }

  /*
   * -------------------------------------------------------
   * PANEL
   * -------------------------------------------------------
   */

  function openPanel() {

    if (
      document.getElementById(
        "tmxe-v045-overlay"
      )
    ) {
      return;
    }

    const overlay =
      document.createElement("div");

    overlay.id =
      "tmxe-v045-overlay";

    overlay.innerHTML = `

      <div
        id="tmxe-v045-panel"
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
              📊 Excel Engine v0.4.5 TEST
            </h2>

            <div
              id="tmxe-v045-help"
            >
              XLSX local → SheetJS →
              DuckDB-Wasm
            </div>

          </div>

          <button
            id="tmxe-v045-close"
            type="button"
          >
            Cerrar
          </button>

        </div>

        <input
          id="tmxe-v045-file"
          type="file"
          accept=".xlsx,.xls,.xlsb,.ods"
        />

        <div
          id="tmxe-v045-actions"
        >

          <button
            id="tmxe-v045-run"
            type="button"
          >
            🚀 Cargar Excel
          </button>

        </div>

        <div
          id="tmxe-v045-status"
        >
          Selecciona un archivo Excel.
        </div>

        <strong>
          Resultado
        </strong>

        <pre
          id="tmxe-v045-result"
        >—</pre>

        <div style="
          margin-top:12px;
          opacity:.65;
          font-size:12px;
        ">

          v0.4.5 TEST —
          procesamiento completamente local.

        </div>

      </div>

    `;

    document.body.appendChild(
      overlay
    );

    /*
     * Close button
     */

    document
      .getElementById(
        "tmxe-v045-close"
      )
      .addEventListener(
        "click",
        () => overlay.remove()
      );

    /*
     * Close clicking outside
     */

    overlay.addEventListener(
      "click",
      event => {

        if (
          event.target === overlay
        ) {

          overlay.remove();

        }

      }
    );

    /*
     * Run
     */

    document
      .getElementById(
        "tmxe-v045-run"
      )
      .addEventListener(
        "click",
        loadExcel
      );
  }

  /*
   * -------------------------------------------------------
   * STATUS
   * -------------------------------------------------------
   */

  function setStatus(
    text,
    kind = ""
  ) {

    const el =
      document.getElementById(
        "tmxe-v045-status"
      );

    if (!el) {
      return;
    }

    el.textContent =
      text;

    if (kind) {

      el.dataset.kind =
        kind;

    } else {

      delete el.dataset.kind;

    }
  }

  /*
   * -------------------------------------------------------
   * RESULT
   *
   * Important:
   * DuckDB COUNT() and other numeric
   * aggregations can return BigInt.
   *
   * JSON.stringify() cannot serialize
   * BigInt by default.
   *
   * We convert BigInt to string.
   * -------------------------------------------------------
   */

  function setResult(
    value
  ) {

    const el =
      document.getElementById(
        "tmxe-v045-result"
      );

    if (!el) {
      return;
    }

    if (
      typeof value === "string"
    ) {

      el.textContent =
        value;

      return;
    }

    el.textContent =
      JSON.stringify(
        value,

        (
          key,
          currentValue
        ) => {

          if (
            typeof currentValue ===
            "bigint"
          ) {

            return currentValue.toString();

          }

          return currentValue;

        },

        2
      );
  }

  /*
   * -------------------------------------------------------
   * LOAD SHEETJS
   * -------------------------------------------------------
   */

  async function loadSheetJS() {

    /*
     * Already loaded
     */

    if (
      window.XLSX
    ) {

      return window.XLSX;

    }

    /*
     * Search existing script
     */

    const existing =
      document.querySelector(
        `script[src="${SHEETJS_URL}"]`
      );

    if (existing) {

      await new Promise(
        (
          resolve,
          reject
        ) => {

          existing.addEventListener(
            "load",
            resolve,
            { once: true }
          );

          existing.addEventListener(
            "error",
            reject,
            { once: true }
          );

        }
      );

      if (
        window.XLSX
      ) {

        return window.XLSX;

      }

    }

    /*
     * Create script
     */

    await new Promise(
      (
        resolve,
        reject
      ) => {

        const script =
          document.createElement(
            "script"
          );

        script.src =
          SHEETJS_URL;

        script.onload =
          resolve;

        script.onerror =
          () => {

            reject(
              new Error(
                "No se pudo cargar SheetJS desde CDN."
              )
            );

          };

        document.head.appendChild(
          script
        );

      }
    );

    if (
      !window.XLSX
    ) {

      throw new Error(
        "SheetJS cargó pero window.XLSX no está disponible."
      );

    }

    return window.XLSX;
  }

  /*
   * -------------------------------------------------------
   * INIT DUCKDB
   * -------------------------------------------------------
   */

  async function initDuckDB() {

    setStatus(
      "5/8 — Inicializando DuckDB-Wasm..."
    );

    /*
     * Import DuckDB-Wasm
     */

    duckdb =
      await import(
        DUCKDB_PACKAGE
      );

    /*
     * Select bundle
     */

    const bundles =
      duckdb.getJsDelivrBundles();

    const bundle =
      await duckdb.selectBundle(
        bundles
      );

    /*
     * Worker
     */

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
      new Worker(
        workerURL
      );

    /*
     * Logger
     */

    const logger =
      new duckdb.ConsoleLogger();

    /*
     * Database
     */

    db =
      new duckdb.AsyncDuckDB(
        logger,
        worker
      );

    /*
     * Instantiate
     */

    await db.instantiate(
      bundle.mainModule,
      bundle.pthreadWorker
    );

    URL.revokeObjectURL(
      workerURL
    );

    /*
     * Connection
     */

    conn =
      await db.connect();

    return bundle;
  }

  /*
   * -------------------------------------------------------
   * SQL IDENTIFIER
   * -------------------------------------------------------
   */

  function safeIdentifier(
    name
  ) {

    return (
      `"${
        String(name)
          .replace(
            /"/g,
            '""'
          )
      }"`
    );

  }

  /*
   * -------------------------------------------------------
   * NORMALIZE HEADERS
   * -------------------------------------------------------
   */

  function normalizeHeaders(
    headers
  ) {

    const result = [];

    const used =
      new Set();

    headers.forEach(
      (
        header,
        index
      ) => {

        let name =
          String(
            header ??
            `column_${index + 1}`
          ).trim();

        if (!name) {

          name =
            `column_${index + 1}`;

        }

        const original =
          name;

        let counter =
          2;

        while (
          used.has(name)
        ) {

          name =
            `${original}_${counter}`;

          counter++;

        }

        used.add(name);

        result.push(
          name
        );

      }
    );

    return result;
  }

  /*
   * -------------------------------------------------------
   * CONVERT VALUE TO SQL
   * -------------------------------------------------------
   */

  function sqlValue(
    value
  ) {

    /*
     * NULL
     */

    if (
      value === null ||
      value === undefined ||
      value === ""
    ) {

      return "NULL";

    }

    /*
     * Number
     */

    if (
      typeof value ===
      "number"
    ) {

      if (
        Number.isFinite(value)
      ) {

        return String(value);

      }

      return "NULL";
    }

    /*
     * Boolean
     */

    if (
      typeof value ===
      "boolean"
    ) {

      return value
        ? "TRUE"
        : "FALSE";

    }

    /*
     * Date
     */

    if (
      value instanceof Date
    ) {

      const iso =
        value
          .toISOString()
          .replace(
            /'/g,
            "''"
          );

      return `'${iso}'`;

    }

    /*
     * Everything else
     * becomes VARCHAR.
     */

    const text =
      String(value)
        .replace(
          /'/g,
          "''"
        );

    return `'${text}'`;
  }

  /*
   * -------------------------------------------------------
   * LOAD EXCEL
   * -------------------------------------------------------
   */

  async function loadExcel() {

    try {

      setResult("");

      /*
       * File
       */

      const fileInput =
        document.getElementById(
          "tmxe-v045-file"
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
        "1/8 — Cargando parser XLSX..."
      );

      const XLSX =
        await loadSheetJS();

      /*
       * 2
       */

      setStatus(
        "2/8 — Leyendo archivo local..."
      );

      const arrayBuffer =
        await file.arrayBuffer();

      /*
       * 3
       */

      setStatus(
        "3/8 — Analizando estructura XLSX..."
      );

      const workbook =
        XLSX.read(
          arrayBuffer,
          {
            type:
              "array",

            dense:
              true
          }
        );

      /*
       * Sheets
       */

      const sheetNames =
        workbook.SheetNames;

      if (
        !sheetNames ||
        !sheetNames.length
      ) {

        throw new Error(
          "El archivo no contiene hojas."
        );

      }

      /*
       * First sheet
       */

      const firstSheetName =
        sheetNames[0];

      const worksheet =
        workbook.Sheets[
          firstSheetName
        ];

      if (!worksheet) {

        throw new Error(
          "No se pudo obtener la primera hoja."
        );

      }

      /*
       * 4
       */

      setStatus(
        "4/8 — Convirtiendo hoja a filas..."
      );

      const rows =
        XLSX.utils.sheet_to_json(
          worksheet,
          {
            header:
              1,

            defval:
              null,

            raw:
              true
          }
        );

      if (
        !rows.length
      ) {

        throw new Error(
          "La primera hoja está vacía."
        );

      }

      /*
       * Header
       */

      const headers =
        normalizeHeaders(
          rows[0]
        );

      /*
       * Data
       */

      const dataRows =
        rows.slice(1);

      /*
       * 5
       */

      setStatus(
        "5/8 — Inicializando DuckDB-Wasm..."
      );

      const bundle =
        await initDuckDB();

      /*
       * 6
       */

      setStatus(
        "6/8 — Creando tabla excel_data..."
      );

      /*
       * v0.4.5 still uses VARCHAR
       * for maximum compatibility.
       *
       * Type inference comes later.
       */

      const columnsSQL =
        headers
          .map(
            name =>
              `${safeIdentifier(name)} VARCHAR`
          )
          .join(", ");

      await conn.query(
        "DROP TABLE IF EXISTS excel_data;"
      );

      await conn.query(
        `
        CREATE TABLE excel_data
        (${columnsSQL});
        `
      );

      /*
       * 7
       */

      setStatus(
        "7/8 — Insertando datos localmente..."
      );

      /*
       * Insert in batches.
       *
       * Current test batch:
       * 500 rows.
       */

      const BATCH_SIZE =
        500;

      let insertedRows =
        0;

      for (
        let start = 0;
        start < dataRows.length;
        start += BATCH_SIZE
      ) {

        const batch =
          dataRows.slice(
            start,
            start + BATCH_SIZE
          );

        const values =
          batch.map(
            row => {

              const cells =
                headers.map(
                  (
                    _,
                    index
                  ) => {

                    return sqlValue(
                      row[index]
                    );

                  }
                );

              return (
                `(${cells.join(",")})`
              );

            }
          );

        if (
          values.length
        ) {

          await conn.query(
            `
            INSERT INTO excel_data
            VALUES
            ${values.join(",")};
            `
          );

          insertedRows +=
            batch.length;

        }

      }

      /*
       * 8
       */

      setStatus(
        "8/8 — Ejecutando consultas SQL..."
      );

      /*
       * COUNT
       */

      const countResult =
        await conn.query(
          `
          SELECT
            COUNT(*) AS total_rows
          FROM excel_data;
          `
        );

      /*
       * DESCRIBE
       */

      const schemaResult =
        await conn.query(
          `
          DESCRIBE excel_data;
          `
        );

      /*
       * Preview
       */

      const previewResult =
        await conn.query(
          `
          SELECT *
          FROM excel_data
          LIMIT 10;
          `
        );

      /*
       * Database version
       */

      const versionResult =
        await conn.query(
          `
          SELECT
            version() AS duckdb_version;
          `
        );

      /*
       * Success
       */

      setStatus(
        "✅ Excel cargado y consultado localmente.",
        "ok"
      );

      /*
       * Final result
       */

      setResult({

        engine:
          APP_ID,

        procesamiento:
          "LOCAL",

        archivo:
          file.name,

        tamano_bytes:
          file.size,

        hojas:
          sheetNames,

        hoja_principal:
          firstSheetName,

        filas_detectadas:
          dataRows.length,

        filas_insertadas:
          insertedRows,

        columnas_detectadas:
          headers.length,

        encabezados:
          headers,

        duckdb_table:
          "excel_data",

        duckdb_version:
          versionResult.toArray(),

        count_sql:
          countResult.toArray(),

        schema:
          schemaResult.toArray(),

        preview:
          previewResult.toArray(),

        bundle:
          {
            mainModule:
              bundle.mainModule,

            mainWorker:
              bundle.mainWorker,

            pthreadWorker:
              bundle.pthreadWorker
          }

      });

    } catch (
      error
    ) {

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

        nombre:
          error?.name ||
          null,

        stack:
          error?.stack ||
          null

      });

    }

  }

  /*
   * -------------------------------------------------------
   * INIT
   * -------------------------------------------------------
   */

  function init() {

    addStyles();

    createButton();

    console.log(
      `[${APP_ID}] cargado`
    );

  }

  /*
   * -------------------------------------------------------
   * START
   * -------------------------------------------------------
   */

  if (
    document.readyState ===
    "loading"
  ) {

    document.addEventListener(
      "DOMContentLoaded",
      init,
      {
        once:
          true
      }
    );

  } else {

    init();

  }

})();
