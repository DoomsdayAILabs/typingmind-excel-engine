/*
 * TypingMind Excel Data Engine - Extension v0.5.0 TEST
 *
 * Objetivo:
 * - Cargar XLSX localmente.
 * - Detectar filas reales.
 * - Inferir tipos.
 * - Crear tabla DuckDB: excel_data.
 * - Ejecutar SQL manual sobre excel_data.
 * - Devolver resultados JSON seguros.
 *
 * TODO LOCAL:
 * Excel -> navegador -> DuckDB-Wasm -> SQL
 *
 * NO se envía el Excel a ningún servidor.
 */

(() => {
  "use strict";

  const APP_ID = "tm-excel-engine-v050-test";

  const DUCKDB_PACKAGE =
    "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

  const XLSX_PACKAGE =
    "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";

  let duckdb = null;
  let db = null;
  let conn = null;
  let currentFile = null;

  const TABLE_NAME = "excel_data";

  /* =========================================================
     UTILIDADES
     ========================================================= */

  function safeJSON(value) {
    if (typeof value === "bigint") {
      return value.toString();
    }

    if (value instanceof Date) {
      return value.toISOString();
    }

    if (Array.isArray(value)) {
      return value.map(safeJSON);
    }

    if (value && typeof value === "object") {
      const out = {};

      for (const [key, val] of Object.entries(value)) {
        out[key] = safeJSON(val);
      }

      return out;
    }

    return value;
  }

  function jsonString(value) {
    return JSON.stringify(
      safeJSON(value),
      null,
      2
    );
  }

  function escapeIdentifier(name) {
    return '"' +
      String(name)
        .replace(/"/g, '""') +
      '"';
  }

  function normalizeHeader(value, index) {
    let text =
      value === null ||
      value === undefined
        ? ""
        : String(value).trim();

    if (!text) {
      text = `columna_${index + 1}`;
    }

    return text;
  }

  function isEmpty(value) {
    return (
      value === null ||
      value === undefined ||
      String(value).trim() === ""
    );
  }

  function isEmptyRow(row) {
    if (!Array.isArray(row)) return true;

    return row.every(isEmpty);
  }

  /* =========================================================
     ESTILOS
     ========================================================= */

  function addStyles() {

    if (
      document.getElementById(
        "tmxe-v050-style"
      )
    ) {
      return;
    }

    const style =
      document.createElement("style");

    style.id =
      "tmxe-v050-style";

    style.textContent = `

      #tmxe-v050-button {
        position: fixed;
        right: 18px;
        bottom: 86px;
        z-index: 2147483000;

        border: 0;
        border-radius: 999px;

        padding: 11px 16px;

        background: #7c3aed;
        color: white;

        font: 600 14px/1.1 system-ui, sans-serif;

        box-shadow:
          0 6px 20px rgba(0,0,0,.25);

        cursor: pointer;
      }

      #tmxe-v050-overlay {
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

      #tmxe-v050-panel {
        width: min(1050px, 96vw);
        max-height: 92vh;

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
          system-ui, sans-serif;
      }

      #tmxe-v050-panel h2 {
        margin:
          0 0 5px;

        font-size: 20px;
      }

      #tmxe-v050-help {
        opacity: .7;
        margin-bottom: 15px;
      }

      #tmxe-v050-status {
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

        margin-bottom: 12px;
      }

      #tmxe-v050-status[data-kind="ok"] {
        color: #16a34a;
      }

      #tmxe-v050-status[data-kind="error"] {
        color: #dc2626;
      }

      #tmxe-v050-result {
        white-space: pre-wrap;

        overflow: auto;

        max-height: 400px;

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

      #tmxe-v050-sql {
        width: 100%;
        min-height: 150px;

        box-sizing: border-box;

        resize: vertical;

        padding: 12px;

        border-radius: 10px;

        border:
          1px solid
          rgba(127,127,127,.45);

        background: Canvas;
        color: CanvasText;

        font:
          13px/1.5
          ui-monospace,
          SFMono-Regular,
          Menlo,
          monospace;
      }

      #tmxe-v050-actions {
        display: flex;

        gap: 8px;

        flex-wrap: wrap;

        margin: 10px 0;
      }

      #tmxe-v050-run,
      #tmxe-v050-load,
      #tmxe-v050-count,
      #tmxe-v050-preview,
      #tmxe-v050-summary,
      #tmxe-v050-close {
        border:
          1px solid
          rgba(127,127,127,.45);

        border-radius: 9px;

        padding: 9px 13px;

        background: transparent;

        color: inherit;

        cursor: pointer;
      }

      #tmxe-v050-run {
        background:
          #7c3aed !important;

        color:
          white !important;

        border-color:
          #7c3aed !important;
      }

      #tmxe-v050-load {
        background:
          #2563eb !important;

        color:
          white !important;

        border-color:
          #2563eb !important;
      }

      #tmxe-v050-file {
        display: none;
      }

      .tmxe-v050-section {
        margin-top: 15px;
      }

      .tmxe-v050-label {
        display: block;

        margin-bottom: 6px;

        font-weight: 700;
      }

      .tmxe-v050-small {
        opacity: .65;

        font-size: 12px;
      }
    `;

    document.head.appendChild(style);
  }

  /* =========================================================
     INTERFAZ
     ========================================================= */

  function createButton() {

    if (
      document.getElementById(
        "tmxe-v050-button"
      )
    ) {
      return;
    }

    const button =
      document.createElement("button");

    button.id =
      "tmxe-v050-button";

    button.type =
      "button";

    button.textContent =
      "📊 Excel v0.5.0";

    button.title =
      "Excel Data Engine v0.5.0";

    button.addEventListener(
      "click",
      openPanel
    );

    document.body.appendChild(button);
  }

  function openPanel() {

    if (
      document.getElementById(
        "tmxe-v050-overlay"
      )
    ) {
      return;
    }

    const overlay =
      document.createElement("div");

    overlay.id =
      "tmxe-v050-overlay";

    overlay.innerHTML = `

      <div
        id="tmxe-v050-panel"
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
              📊 Excel Data Engine v0.5.0
            </h2>

            <div id="tmxe-v050-help">
              Excel local + DuckDB-Wasm + SQL
            </div>

          </div>

          <button
            id="tmxe-v050-close"
            type="button"
          >
            Cerrar
          </button>

        </div>

        <div id="tmxe-v050-status">
          Motor listo.
          Carga un archivo Excel para comenzar.
        </div>

        <input
          id="tmxe-v050-file"
          type="file"
          accept=".xlsx,.xls"
        />

        <div class="tmxe-v050-section">

          <span class="tmxe-v050-label">
            1. Archivo Excel
          </span>

          <div id="tmxe-v050-file-name"
               class="tmxe-v050-small">
            Ningún archivo cargado.
          </div>

          <div id="tmxe-v050-actions"
               class="tmxe-v050-actions">

            <button
              id="tmxe-v050-load"
              type="button"
            >
              📂 Cargar Excel
            </button>

            <button
              id="tmxe-v050-count"
              type="button"
            >
              🔢 COUNT
            </button>

            <button
              id="tmxe-v050-preview"
              type="button"
            >
              👁 Vista previa
            </button>

            <button
              id="tmxe-v050-summary"
              type="button"
            >
              📊 Resumen
            </button>

          </div>

        </div>

        <div class="tmxe-v050-section">

          <span class="tmxe-v050-label">
            2. SQL manual
          </span>

          <textarea
            id="tmxe-v050-sql"
          >SELECT
    EXTRACT(YEAR FROM "Fecha") AS año,
    SUM("Total TEU's") AS total
FROM excel_data
GROUP BY año
ORDER BY año;</textarea>

          <div
            class="tmxe-v050-small"
            style="margin-top:6px;"
          >
            La consulta se ejecuta localmente
            sobre la tabla <b>excel_data</b>.
          </div>

          <div
            id="tmxe-v050-actions"
            class="tmxe-v050-actions"
          >

            <button
              id="tmxe-v050-run"
              type="button"
            >
              ▶ Ejecutar SQL local
            </button>

          </div>

        </div>

        <div class="tmxe-v050-section">

          <strong>
            Resultado
          </strong>

          <pre
            id="tmxe-v050-result"
          >—</pre>

        </div>

        <div
          class="tmxe-v050-small"
          style="margin-top:12px;"
        >
          v0.5.0 TEST —
          procesamiento local.
          El archivo no se envía al servidor.
        </div>

      </div>
    `;

    document.body.appendChild(
      overlay
    );

    document
      .getElementById(
        "tmxe-v050-close"
      )
      .addEventListener(
        "click",
        () => overlay.remove()
      );

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

    document
      .getElementById(
        "tmxe-v050-load"
      )
      .addEventListener(
        "click",
        () => {

          document
            .getElementById(
              "tmxe-v050-file"
            )
            .click();

        }
      );

    document
      .getElementById(
        "tmxe-v050-file"
      )
      .addEventListener(
        "change",
        loadExcel
      );

    document
      .getElementById(
        "tmxe-v050-count"
      )
      .addEventListener(
        "click",
        () =>
          executeSQL(
            `
              SELECT
                COUNT(*) AS registros
              FROM excel_data;
            `
          )
      );

    document
      .getElementById(
        "tmxe-v050-preview"
      )
      .addEventListener(
        "click",
        () =>
          executeSQL(
            `
              SELECT *
              FROM excel_data
              LIMIT 10;
            `
          )
      );

    document
      .getElementById(
        "tmxe-v050-summary"
      )
      .addEventListener(
        "click",
        () =>
          executeSQL(
            `
              SELECT *
              FROM excel_data
              LIMIT 10;
            `
          )
      );

    document
      .getElementById(
        "tmxe-v050-run"
      )
      .addEventListener(
        "click",
        runManualSQL
      );
  }

  function setStatus(
    text,
    kind = ""
  ) {

    const el =
      document.getElementById(
        "tmxe-v050-status"
      );

    if (!el) return;

    el.textContent =
      text;

    if (kind) {
      el.dataset.kind =
        kind;
    } else {
      delete el.dataset.kind;
    }
  }

  function setResult(value) {

    const el =
      document.getElementById(
        "tmxe-v050-result"
      );

    if (!el) return;

    if (
      typeof value ===
      "string"
    ) {

      el.textContent =
        value;

    } else {

      el.textContent =
        jsonString(value);

    }
  }

  /* =========================================================
     DUCKDB
     ========================================================= */

  async function initializeDuckDB() {

    if (
      duckdb &&
      db &&
      conn
    ) {
      return;
    }

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
      new Worker(
        workerURL
      );

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
  }

  /* =========================================================
     XLSX
     ========================================================= */

  async function loadExcel(event) {

    const file =
      event.target.files?.[0];

    if (!file) return;

    try {

      currentFile =
        file;

      setResult("");

      setStatus(
        "1/7 — Cargando librería XLSX..."
      );

      const XLSX =
        await import(
          XLSX_PACKAGE
        );

      setStatus(
        "2/7 — Leyendo archivo local..."
      );

      const buffer =
        await file.arrayBuffer();

      const workbook =
        XLSX.read(
          buffer,
          {
            type:
              "array",
            cellDates:
              true
          }
        );

      const sheets =
        workbook.SheetNames;

      if (
        !sheets.length
      ) {
        throw new Error(
          "El archivo no contiene hojas."
        );
      }

      const mainSheet =
        sheets[0];

      const worksheet =
        workbook.Sheets[
          mainSheet
        ];

      setStatus(
        "3/7 — Analizando filas..."
      );

      const rows =
        XLSX.utils.sheet_to_json(
          worksheet,
          {
            header: 1,
            defval: null,
            raw: true
          }
        );

      if (
        !rows.length
      ) {
        throw new Error(
          "La hoja está vacía."
        );
      }

      const headerRow =
        rows[0];

      const headers =
        headerRow.map(
          normalizeHeader
        );

      const dataRows =
        rows
          .slice(1)
          .filter(
            row =>
              !isEmptyRow(row)
          );

      setStatus(
        "4/7 — Inicializando DuckDB..."
      );

      await initializeDuckDB();

      setStatus(
        "5/7 — Creando tabla local..."
      );

      await createDuckDBTable(
        headers,
        dataRows
      );

      setStatus(
        "6/7 — Verificando datos..."
      );

      const countResult =
        await executeSQLInternal(
          `
            SELECT
              COUNT(*) AS registros
            FROM ${TABLE_NAME};
          `,
          false
        );

      const count =
        countResult.resultado?.[0]
          ?.registros ?? null;

      document
        .getElementById(
          "tmxe-v050-file-name"
        )
        .textContent =
          `${file.name} — ${count} registros`;

      setStatus(
        "7/7 — Excel cargado correctamente.",
        "ok"
      );

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
          sheets,

        hoja_principal:
          mainSheet,

        filas_fisicas_detectadas:
          rows.length,

        filas_reales:
          dataRows.length,

        filas_vacias_ignoradas:
          rows.length -
          1 -
          dataRows.length,

        filas_insertadas:
          Number(count),

        columnas_detectadas:
          headers.length,

        encabezados:
          headers,

        duckdb_table:
          TABLE_NAME,

        count_sql:
          countResult
      });

    } catch (error) {

      console.error(
        `[${APP_ID}]`,
        error
      );

      setStatus(
        "❌ ERROR AL CARGAR EXCEL",
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

  /* =========================================================
     INFERENCIA DE TIPOS
     ========================================================= */

  function inferColumnType(
    values,
    header
  ) {

    const nonEmpty =
      values.filter(
        value =>
          !isEmpty(value)
      );

    if (
      !nonEmpty.length
    ) {
      return "VARCHAR";
    }

    const headerLower =
      String(header)
        .toLowerCase();

    const dateHeader =
      /fecha|date|día|dia|mes|month|year|año/
        .test(
          headerLower
        );

    const allDates =
      nonEmpty.every(
        value =>
          value instanceof Date ||
          (
            typeof value ===
              "number" &&
            Number.isFinite(value) &&
            value > 20000 &&
            value < 100000
          )
      );

    if (
      dateHeader &&
      allDates
    ) {
      return "DATE";
    }

    const allIntegers =
      nonEmpty.every(
        value => {

          if (
            typeof value ===
            "number"
          ) {
            return (
              Number.isFinite(value) &&
              Number.isInteger(value)
            );
          }

          return /^-?\d+$/
            .test(
              String(value).trim()
            );
        }
      );

    if (allIntegers) {
      return "BIGINT";
    }

    const allNumbers =
      nonEmpty.every(
        value => {

          if (
            typeof value ===
            "number"
          ) {
            return Number.isFinite(
              value
            );
          }

          const n =
            Number(
              String(value)
                .trim()
                .replace(/,/g, "")
            );

          return Number.isFinite(n);
        }
      );

    if (allNumbers) {
      return "DOUBLE";
    }

    return "VARCHAR";
  }

  /* =========================================================
     CREACIÓN TABLA
     ========================================================= */

  async function createDuckDBTable(
    headers,
    rows
  ) {

    await conn.query(
      `DROP TABLE IF EXISTS ${TABLE_NAME};`
    );

    const types =
      headers.map(
        (_, columnIndex) =>
          inferColumnType(
            rows.map(
              row =>
                row[columnIndex]
            ),
            headers[columnIndex]
          )
      );

    const columnsSQL =
      headers
        .map(
          (header, index) =>
            `${escapeIdentifier(
              header
            )} ${types[index]}`
        )
        .join(", ");

    await conn.query(
      `
        CREATE TABLE ${TABLE_NAME} (
          ${columnsSQL}
        );
      `
    );

    if (
      rows.length === 0
    ) {
      return;
    }

    const quotedHeaders =
      headers
        .map(
          escapeIdentifier
        )
        .join(", ");

    const valuesSQL =
      rows.map(
        row => {

          const values =
            headers.map(
              (header, index) => {

                const value =
                  row[index];

                const type =
                  types[index];

                return sqlValue(
                  value,
                  type
                );
              }
            );

          return `(${values.join(", ")})`;
        }
      )
      .join(",\n");

    await conn.query(
      `
        INSERT INTO
          ${TABLE_NAME}
          (${quotedHeaders})
        VALUES
          ${valuesSQL};
      `
    );
  }

  function sqlValue(
    value,
    type
  ) {

    if (isEmpty(value)) {
      return "NULL";
    }

    if (
      type ===
      "DATE"
    ) {

      if (
        value instanceof Date
      ) {

        const iso =
          value
            .toISOString()
            .slice(0, 10);

        return `DATE '${iso}'`;
      }

      if (
        typeof value ===
        "number"
      ) {

        const date =
          excelSerialToDate(
            value
          );

        const iso =
          date
            .toISOString()
            .slice(0, 10);

        return `DATE '${iso}'`;
      }

      return `DATE '${String(
        value
      ).slice(0, 10)}'`;
    }

    if (
      type ===
      "BIGINT"
    ) {

      const n =
        BigInt(
          String(value)
            .trim()
            .replace(/,/g, "")
        );

      return n.toString();
    }

    if (
      type ===
      "DOUBLE"
    ) {

      const n =
        Number(
          String(value)
            .trim()
            .replace(/,/g, "")
        );

      if (
        !Number.isFinite(n)
      ) {
        return "NULL";
      }

      return String(n);
    }

    return `'${String(value)
      .replace(/'/g, "''")}'`;
  }

  function excelSerialToDate(
    serial
  ) {

    const epoch =
      Date.UTC(
        1899,
        11,
        30
      );

    return new Date(
      epoch +
      serial * 86400000
    );
  }

  /* =========================================================
     SQL
     ========================================================= */

  function validateSQL(
    sql
  ) {

    const cleaned =
      String(sql)
        .trim()
        .replace(/;+\s*$/, "");

    if (!cleaned) {
      throw new Error(
        "La consulta SQL está vacía."
      );
    }

    const firstWord =
      cleaned
        .split(/\s+/)[0]
        .toUpperCase();

    const allowed = [
      "SELECT",
      "WITH",
      "DESCRIBE",
      "SHOW",
      "SUMMARIZE"
    ];

    if (
      !allowed.includes(
        firstWord
      )
    ) {

      throw new Error(
        "Por seguridad, v0.5.0 TEST solo permite consultas SELECT, WITH, DESCRIBE, SHOW o SUMMARIZE."
      );
    }

    const forbidden =
      /\b(INSERT|UPDATE|DELETE|DROP|ALTER|CREATE|TRUNCATE|COPY|ATTACH|DETACH|INSTALL|LOAD|EXPORT|IMPORT)\b/i;

    if (
      forbidden.test(
        cleaned
      )
    ) {

      throw new Error(
        "La consulta contiene una operación no permitida."
      );
    }

    return cleaned;
  }

  async function executeSQLInternal(
    sql,
    addMetadata = true
  ) {

    if (!conn) {

      throw new Error(
        "DuckDB no está inicializado. Carga primero un Excel."
      );
    }

    const safeSQL =
      validateSQL(
        sql
      );

    const result =
      await conn.query(
        safeSQL
      );

    const rows =
      result.toArray()
        .map(
          row =>
            safeJSON(row)
        );

    return {
      ...(addMetadata
        ? {
            procesamiento:
              "LOCAL",

            sql:
              safeSQL
          }
        : {}),

      filas_resultado:
        rows.length,

      resultado:
        rows
    };
  }

  async function executeSQL(
    sql
  ) {

    try {

      if (!conn) {

        setStatus(
          "Carga primero un archivo Excel.",
          "error"
        );

        return;
      }

      setStatus(
        "Ejecutando SQL local..."
      );

      setResult("");

      const result =
        await executeSQLInternal(
          sql,
          true
        );

      setStatus(
        "Consulta ejecutada correctamente — procesamiento LOCAL.",
        "ok"
      );

      setResult(
        result
      );

      console.log(
        `[${APP_ID}]`,
        result
      );

    } catch (error) {

      console.error(
        `[${APP_ID}] SQL`,
        error
      );

      setStatus(
        "❌ ERROR EN SQL",
        "error"
      );

      setResult({
        procesamiento:
          "LOCAL",

        error:
          error?.message ||
          String(error),

        sql:
          sql
      });
    }
  }

  async function runManualSQL() {

    const textarea =
      document.getElementById(
        "tmxe-v050-sql"
      );

    if (!textarea) return;

    await executeSQL(
      textarea.value
    );
  }

  /* =========================================================
     INIT
     ========================================================= */

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
      {
        once: true
      }
    );

  } else {

    init();

  }

})();
