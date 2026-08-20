/*
 * TypingMind Excel Data Engine - Extension v0.4.9.1 TEST
 *
 * OBJETIVO:
 * - Cargar XLSX localmente.
 * - Procesar Excel con JavaScript.
 * - Crear tabla DuckDB: excel_data.
 * - Inferir tipos básicos.
 * - Ejecutar SQL escrito manualmente por el usuario.
 * - Serializar correctamente BIGINT, DATE, TIMESTAMP,
 *   arrays y estructuras DuckDB.
 *
 * TODO EL PROCESAMIENTO DEL EXCEL ES LOCAL.
 */

(() => {
  "use strict";

  const APP_ID = "tm-excel-engine-v0491-test";

  const VERSION = "0.4.9.1-test";

  const DUCKDB_PACKAGE =
    "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

  let duckdb = null;
  let db = null;
  let conn = null;

  let currentFile = null;
  let currentTable = "excel_data";

  /* =========================================================
     UTILIDADES
     ========================================================= */

  function safeString(value) {
    if (value === null || value === undefined) {
      return "";
    }

    return String(value);
  }

  function isEmptyRow(row) {
    if (!Array.isArray(row)) return true;

    return row.every(
      value =>
        value === null ||
        value === undefined ||
        String(value).trim() === ""
    );
  }

  function escapeIdentifier(name) {
    return `"${String(name).replace(/"/g, '""')}"`;
  }

  function escapeSQLString(value) {
    return String(value).replace(/'/g, "''");
  }

  /* =========================================================
     SERIALIZACIÓN ROBUSTA
     ========================================================= */

  function serializeValue(value) {
    if (value === null || value === undefined) {
      return value;
    }

    if (typeof value === "bigint") {
      const numberValue = Number(value);

      if (
        Number.isSafeInteger(numberValue)
      ) {
        return numberValue;
      }

      return value.toString();
    }

    if (value instanceof Date) {
      return value.toISOString();
    }

    if (
      typeof value === "number" ||
      typeof value === "string" ||
      typeof value === "boolean"
    ) {
      return value;
    }

    if (ArrayBuffer.isView(value)) {
      return Array.from(value).map(serializeValue);
    }

    if (Array.isArray(value)) {
      return value.map(serializeValue);
    }

    if (typeof value === "object") {
      const result = {};

      for (const [key, item] of Object.entries(value)) {
        result[key] = serializeValue(item);
      }

      return result;
    }

    return String(value);
  }

  function serializeRows(rows) {
    return rows.map(row => serializeValue(row));
  }

  function resultToRows(result) {
    if (!result) return [];

    try {
      return serializeRows(result.toArray());
    } catch (error) {
      console.warn(
        `[${APP_ID}] Error convirtiendo resultado:`,
        error
      );

      return [];
    }
  }

  /* =========================================================
     ESTILOS
     ========================================================= */

  function addStyles() {
    if (
      document.getElementById(
        "tmxe-v0491-style"
      )
    ) {
      return;
    }

    const style =
      document.createElement("style");

    style.id =
      "tmxe-v0491-style";

    style.textContent = `
      #tmxe-v0491-button {
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

      #tmxe-v0491-overlay {
        position: fixed;
        inset: 0;

        z-index: 2147483001;

        background: rgba(0,0,0,.48);

        display: flex;

        align-items: center;
        justify-content: center;

        padding: 16px;
      }

      #tmxe-v0491-panel {
        width: min(1100px, 96vw);
        max-height: 92vh;

        overflow: auto;

        background: Canvas;
        color: CanvasText;

        border:
          1px solid rgba(127,127,127,.35);

        border-radius: 16px;

        padding: 20px;

        box-shadow:
          0 20px 70px rgba(0,0,0,.35);

        font:
          14px/1.45 system-ui, sans-serif;
      }

      #tmxe-v0491-panel h2 {
        margin: 0 0 5px;
        font-size: 20px;
      }

      #tmxe-v0491-help {
        opacity: .7;
        margin-bottom: 15px;
      }

      #tmxe-v0491-status {
        padding: 12px;

        border-radius: 9px;

        background:
          rgba(127,127,127,.10);

        white-space: pre-wrap;

        font:
          13px/1.5 ui-monospace,
          SFMono-Regular,
          Menlo,
          monospace;

        margin-bottom: 12px;
      }

      #tmxe-v0491-status[data-kind="ok"] {
        color: #16a34a;
      }

      #tmxe-v0491-status[data-kind="error"] {
        color: #dc2626;
      }

      #tmxe-v0491-file {
        display: none;
      }

      #tmxe-v0491-sql {
        width: 100%;
        min-height: 190px;

        box-sizing: border-box;

        resize: vertical;

        padding: 12px;

        border:
          1px solid rgba(127,127,127,.40);

        border-radius: 10px;

        background:
          rgba(127,127,127,.08);

        color: inherit;

        font:
          13px/1.5 ui-monospace,
          SFMono-Regular,
          Menlo,
          Consolas,
          monospace;
      }

      #tmxe-v0491-result {
        white-space: pre-wrap;

        overflow: auto;

        max-height: 500px;

        padding: 12px;

        border-radius: 10px;

        background:
          rgba(127,127,127,.10);

        font:
          12px/1.45 ui-monospace,
          SFMono-Regular,
          Menlo,
          Consolas,
          monospace;
      }

      #tmxe-v0491-actions {
        display: flex;

        gap: 8px;

        flex-wrap: wrap;

        margin: 10px 0;
      }

      #tmxe-v0491-actions button,
      #tmxe-v0491-close,
      #tmxe-v0491-load,
      #tmxe-v0491-run {
        border:
          1px solid rgba(127,127,127,.45);

        border-radius: 9px;

        padding: 9px 13px;

        background: transparent;

        color: inherit;

        cursor: pointer;
      }

      #tmxe-v0491-load,
      #tmxe-v0491-run {
        background: #7c3aed !important;
        color: white !important;
        border-color: #7c3aed !important;
      }

      #tmxe-v0491-small {
        opacity: .65;
        font-size: 12px;
      }

      #tmxe-v0491-section {
        margin-top: 16px;
      }

      #tmxe-v0491-label {
        display: block;
        font-weight: 700;
        margin-bottom: 7px;
      }

      #tmxe-v0491-file-name {
        margin-left: 8px;
        opacity: .75;
      }

      #tmxe-v0491-presets {
        display: flex;
        gap: 7px;
        flex-wrap: wrap;
        margin: 8px 0;
      }

      #tmxe-v0491-presets button {
        border:
          1px solid rgba(127,127,127,.40);

        border-radius: 8px;

        padding: 7px 10px;

        background:
          rgba(127,127,127,.08);

        color: inherit;

        cursor: pointer;
      }
    `;

    document.head.appendChild(style);
  }

  /* =========================================================
     UI
     ========================================================= */

  function createButton() {
    if (
      document.getElementById(
        "tmxe-v0491-button"
      )
    ) {
      return;
    }

    const button =
      document.createElement("button");

    button.id =
      "tmxe-v0491-button";

    button.type = "button";

    button.textContent =
      "📊 Excel Engine";

    button.title =
      "Excel Data Engine v0.4.9.1";

    button.addEventListener(
      "click",
      openPanel
    );

    document.body.appendChild(button);
  }

  function openPanel() {
    if (
      document.getElementById(
        "tmxe-v0491-overlay"
      )
    ) {
      return;
    }

    const overlay =
      document.createElement("div");

    overlay.id =
      "tmxe-v0491-overlay";

    overlay.innerHTML = `
      <div
        id="tmxe-v0491-panel"
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
              📊 Excel Data Engine
              v0.4.9.1 TEST
            </h2>

            <div id="tmxe-v0491-help">
              Excel local + DuckDB-Wasm +
              SQL manual.
            </div>

          </div>

          <button
            id="tmxe-v0491-close"
            type="button"
          >
            Cerrar
          </button>

        </div>

        <div id="tmxe-v0491-status">
          Motor listo.
          Selecciona un archivo Excel.
        </div>

        <div id="tmxe-v0491-actions">

          <input
            id="tmxe-v0491-file"
            type="file"
            accept=".xlsx,.xls"
          />

          <button
            id="tmxe-v0491-load"
            type="button"
          >
            📂 Cargar Excel
          </button>

          <span
            id="tmxe-v0491-file-name"
          >
            Ningún archivo seleccionado
          </span>

        </div>

        <div id="tmxe-v0491-section">

          <label
            id="tmxe-v0491-label"
            for="tmxe-v0491-sql"
          >
            SQL
          </label>

          <textarea
            id="tmxe-v0491-sql"
            spellcheck="false"
          >SELECT
    "Año",
    SUM("Total TEU's") AS total
FROM excel_data
GROUP BY "Año"
ORDER BY "Año";</textarea>

        </div>

        <div id="tmxe-v0491-presets">

          <button
            type="button"
            data-sql="SELECT COUNT(*) AS registros FROM excel_data;"
          >
            COUNT
          </button>

          <button
            type="button"
            data-sql="SELECT * FROM excel_data LIMIT 10;"
          >
            Vista previa
          </button>

          <button
            type="button"
            data-sql='SELECT
    SUM("Local") AS "Local__sum",
    AVG("Local") AS "Local__avg",
    MIN("Local") AS "Local__min",
    MAX("Local") AS "Local__max",
    SUM("Transshipment") AS "Transshipment__sum",
    AVG("Transshipment") AS "Transshipment__avg",
    MIN("Transshipment") AS "Transshipment__min",
    MAX("Transshipment") AS "Transshipment__max",
    SUM("Total TEU''s") AS "Total TEU''s__sum",
    AVG("Total TEU''s") AS "Total TEU''s__avg",
    MIN("Total TEU''s") AS "Total TEU''s__min",
    MAX("Total TEU''s") AS "Total TEU''s__max"
FROM excel_data;'
          >
            Resumen
          </button>

          <button
            type="button"
            data-sql='SELECT
    "Año",
    SUM("Total TEU''s") AS total
FROM excel_data
GROUP BY "Año"
ORDER BY "Año";'
          >
            Total por año
          </button>

        </div>

        <div id="tmxe-v0491-actions">

          <button
            id="tmxe-v0491-run"
            type="button"
          >
            ▶ Ejecutar SQL
          </button>

        </div>

        <div id="tmxe-v0491-section">

          <strong>
            Resultado
          </strong>

          <pre
            id="tmxe-v0491-result"
          >—</pre>

        </div>

        <div id="tmxe-v0491-small">
          v0.4.9.1 TEST — Todo el
          procesamiento se realiza localmente.
        </div>

      </div>
    `;

    document.body.appendChild(overlay);

    document
      .getElementById(
        "tmxe-v0491-close"
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
        "tmxe-v0491-load"
      )
      .addEventListener(
        "click",
        () => {
          document
            .getElementById(
              "tmxe-v0491-file"
            )
            .click();
        }
      );

    document
      .getElementById(
        "tmxe-v0491-file"
      )
      .addEventListener(
        "change",
        handleFileSelection
      );

    document
      .getElementById(
        "tmxe-v0491-run"
      )
      .addEventListener(
        "click",
        executeSQL
      );

    document
      .querySelectorAll(
        "#tmxe-v0491-presets button"
      )
      .forEach(button => {

        button.addEventListener(
          "click",
          () => {

            const sql =
              button.dataset.sql;

            document
              .getElementById(
                "tmxe-v0491-sql"
              )
              .value = sql;

          }
        );

      });
  }

  function setStatus(
    text,
    kind = ""
  ) {
    const element =
      document.getElementById(
        "tmxe-v0491-status"
      );

    if (!element) return;

    element.textContent = text;

    if (kind) {
      element.dataset.kind =
        kind;
    } else {
      delete element.dataset.kind;
    }
  }

  function setResult(value) {
    const element =
      document.getElementById(
        "tmxe-v0491-result"
      );

    if (!element) return;

    try {

      if (
        typeof value === "string"
      ) {

        element.textContent =
          value;

      } else {

        element.textContent =
          JSON.stringify(
            serializeValue(value),
            null,
            2
          );

      }

    } catch (error) {

      element.textContent =
        String(value);

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
      "1/5 — Cargando DuckDB-Wasm..."
    );

    duckdb =
      await import(
        DUCKDB_PACKAGE
      );

    setStatus(
      "2/5 — Seleccionando bundle..."
    );

    const bundles =
      duckdb.getJsDelivrBundles();

    const bundle =
      await duckdb.selectBundle(
        bundles
      );

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

    await conn.query(
      "SELECT 1;"
    );

    setStatus(
      "4/5 — DuckDB conectado."
    );
  }

  /* =========================================================
     ARCHIVO
     ========================================================= */

  async function handleFileSelection(
    event
  ) {

    const file =
      event.target.files?.[0];

    if (!file) return;

    currentFile = file;

    const nameElement =
      document.getElementById(
        "tmxe-v0491-file-name"
      );

    if (nameElement) {
      nameElement.textContent =
        file.name;
    }

    try {

      await loadExcel(
        file
      );

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

  async function loadExcel(
    file
  ) {

    await initializeDuckDB();

    setStatus(
      "5/5 — Leyendo archivo Excel..."
    );

    const buffer =
      await file.arrayBuffer();

    /*
     * XLSX parser.
     */

    const XLSX =
      await import(
        "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm"
      );

    const workbook =
      XLSX.read(
        buffer,
        {
          type: "array",
          cellDates: false
        }
      );

    const sheets =
      workbook.SheetNames;

    if (!sheets.length) {
      throw new Error(
        "El archivo no contiene hojas."
      );
    }

    const sheetName =
      sheets[0];

    const sheet =
      workbook.Sheets[
        sheetName
      ];

    const matrix =
      XLSX.utils.sheet_to_json(
        sheet,
        {
          header: 1,
          defval: null,
          raw: true
        }
      );

    const physicalRows =
      matrix.length;

    /*
     * Buscar encabezados.
     */

    let headerIndex = -1;

    for (
      let i = 0;
      i < matrix.length;
      i++
    ) {

      if (
        !isEmptyRow(matrix[i])
      ) {

        headerIndex = i;

        break;
      }
    }

    if (
      headerIndex < 0
    ) {

      throw new Error(
        "No se encontró una fila de encabezados."
      );

    }

    const headers =
      matrix[headerIndex].map(
        (value, index) => {

          const text =
            safeString(value)
              .trim();

          return (
            text ||
            `Column_${index + 1}`
          );

        }
      );

    /*
     * Filas reales.
     */

    const dataRows =
      matrix
        .slice(
          headerIndex + 1
        )
        .filter(
          row =>
            !isEmptyRow(row)
        );

    /*
     * Inferencia.
     */

    const inferred =
      inferColumnTypes(
        headers,
        dataRows
      );

    /*
     * Crear tabla.
     */

    await conn.query(
      `DROP TABLE IF EXISTS ${escapeIdentifier(currentTable)};`
    );

    const definitions =
      headers.map(
        (header, index) => {

          return `${escapeIdentifier(
            header
          )} ${inferred[index].sqlType}`;

        }
      );

    await conn.query(
      `CREATE TABLE ${escapeIdentifier(
        currentTable
      )} (${definitions.join(", ")});`
    );

    /*
     * Insertar datos.
     */

    for (
      const row of dataRows
    ) {

      const values =
        headers.map(
          (header, index) => {

            const value =
              row[index];

            return sqlLiteral(
              value,
              inferred[index].sqlType
            );

          }
        );

      await conn.query(
        `INSERT INTO ${escapeIdentifier(
          currentTable
        )}
        VALUES (${values.join(", ")});`
      );

    }

    const countResult =
      await conn.query(
        `SELECT COUNT(*) AS registros
         FROM ${escapeIdentifier(
           currentTable
         )};`
      );

    const count =
      resultToRows(
        countResult
      );

    setStatus(
      "✅ Excel cargado correctamente.\n" +
      `Archivo: ${file.name}\n` +
      `Hoja: ${sheetName}\n` +
      `Registros: ${
        count[0]?.registros ?? "?"
      }`,
      "ok"
    );

    setResult({
      procesamiento: "LOCAL",
      archivo: file.name,
      tamano_bytes: file.size,
      hojas: sheets,
      hoja_principal: sheetName,
      filas_fisicas_detectadas:
        physicalRows,
      filas_reales:
        dataRows.length,
      filas_vacias_ignoradas:
        physicalRows -
        headerIndex -
        1 -
        dataRows.length,
      filas_insertadas:
        dataRows.length,
      columnas_detectadas:
        headers.length,
      encabezados:
        headers,
      tipos_inferidos:
        inferred,
      duckdb_table:
        currentTable,
      count:
        count
    });
  }

  /* =========================================================
     INFERENCIA DE TIPOS
     ========================================================= */

  function inferColumnTypes(
    headers,
    rows
  ) {

    return headers.map(
      (header, index) => {

        const values =
          rows
            .map(
              row =>
                row[index]
            )
            .filter(
              value =>
                value !== null &&
                value !== undefined &&
                String(value).trim() !== ""
            );

        const headerLower =
          header
            .toLowerCase()
            .trim();

        /*
         * FECHA
         */

        const dateHeader =
          /fecha|date|día|dia/.test(
            headerLower
          );

        const allDates =
          values.length > 0 &&
          values.every(
            value =>
              isExcelDateValue(
                value
              )
          );

        if (
          dateHeader &&
          allDates
        ) {

          return {
            column_index: index,
            column_name: header,
            duckdb_type: "DATE",
            sqlType: "DATE",
            confidence: "high",
            reason:
              "encabezado y valores compatibles con fecha",
            non_empty_values:
              values.length
          };

        }

        /*
         * INTEGER
         */

        const allIntegers =
          values.length > 0 &&
          values.every(
            value =>
              isIntegerValue(
                value
              )
          );

        if (
          allIntegers
        ) {

          return {
            column_index: index,
            column_name: header,
            duckdb_type: "BIGINT",
            sqlType: "BIGINT",
            confidence: "high",
            reason:
              "todos los valores son enteros",
            non_empty_values:
              values.length
          };

        }

        /*
         * DOUBLE
         */

        const allNumbers =
          values.length > 0 &&
          values.every(
            value =>
              isNumericValue(
                value
              )
          );

        if (
          allNumbers
        ) {

          return {
            column_index: index,
            column_name: header,
            duckdb_type: "DOUBLE",
            sqlType: "DOUBLE",
            confidence: "high",
            reason:
              "todos los valores son numéricos",
            non_empty_values:
              values.length
          };

        }

        return {
          column_index: index,
          column_name: header,
          duckdb_type: "VARCHAR",
          sqlType: "VARCHAR",
          confidence: "medium",
          reason:
            "valores tratados como texto",
          non_empty_values:
            values.length
        };

      }
    );
  }

  function isIntegerValue(
    value
  ) {

    if (
      typeof value === "number"
    ) {
      return (
        Number.isFinite(value) &&
        Number.isInteger(value)
      );
    }

    const text =
      String(value).trim();

    return /^-?\d+$/.test(
      text
    );
  }

  function isNumericValue(
    value
  ) {

    if (
      typeof value === "number"
    ) {
      return Number.isFinite(
        value
      );
    }

    const text =
      String(value).trim();

    return (
      text !== "" &&
      !Number.isNaN(
        Number(text)
      )
    );
  }

  function isExcelDateValue(
    value
  ) {

    /*
     * Excel serial date.
     */

    if (
      typeof value === "number"
    ) {

      return (
        value > 20000 &&
        value < 100000
      );

    }

    /*
     * JavaScript Date.
     */

    if (
      value instanceof Date
    ) {
      return !isNaN(
        value.getTime()
      );
    }

    /*
     * Texto ISO.
     */

    if (
      typeof value === "string"
    ) {

      return /^\d{4}-\d{1,2}-\d{1,2}$/.test(
        value.trim()
      );

    }

    return false;
  }

  /* =========================================================
     CONVERSIÓN A SQL
     ========================================================= */

  function excelSerialToDate(
    serial
  ) {

    const utcDays =
      Math.floor(
        serial - 25569
      );

    const date =
      new Date(
        utcDays * 86400 * 1000
      );

    const year =
      date.getUTCFullYear();

    const month =
      String(
        date.getUTCMonth() + 1
      ).padStart(2, "0");

    const day =
      String(
        date.getUTCDate()
      ).padStart(2, "0");

    return `${year}-${month}-${day}`;
  }

  function normalizeDate(
    value
  ) {

    if (
      typeof value === "number"
    ) {

      return excelSerialToDate(
        value
      );

    }

    if (
      value instanceof Date
    ) {

      const year =
        value.getUTCFullYear();

      const month =
        String(
          value.getUTCMonth() + 1
        ).padStart(2, "0");

      const day =
        String(
          value.getUTCDate()
        ).padStart(2, "0");

      return `${year}-${month}-${day}`;
    }

    return String(value).trim();
  }

  function sqlLiteral(
    value,
    sqlType
  ) {

    if (
      value === null ||
      value === undefined ||
      String(value).trim() === ""
    ) {
      return "NULL";
    }

    if (
      sqlType === "DATE"
    ) {

      const date =
        normalizeDate(
          value
        );

      return `DATE '${escapeSQLString(
        date
      )}'`;
    }

    if (
      sqlType === "BIGINT"
    ) {

      const number =
        Number(value);

      if (
        Number.isSafeInteger(
          number
        )
      ) {

        return String(
          number
        );

      }

      return "NULL";
    }

    if (
      sqlType === "DOUBLE"
    ) {

      const number =
        Number(value);

      if (
        Number.isFinite(
          number
        )
      ) {

        return String(
          number
        );

      }

      return "NULL";
    }

    return `'${escapeSQLString(
      value
    )}'`;
  }

  /* =========================================================
     EJECUCIÓN SQL
     ========================================================= */

  async function executeSQL() {

    const sqlElement =
      document.getElementById(
        "tmxe-v0491-sql"
      );

    if (!sqlElement) {
      return;
    }

    const sql =
      sqlElement.value.trim();

    if (!sql) {

      setStatus(
        "Escribe una consulta SQL.",
        "error"
      );

      return;
    }

    if (!conn) {

      setStatus(
        "Primero debes cargar un archivo Excel.",
        "error"
      );

      setResult({
        error:
          "DuckDB no está conectado."
      });

      return;
    }

    try {

      setStatus(
        "⏳ Ejecutando SQL localmente..."
      );

      const start =
        performance.now();

      const result =
        await conn.query(
          sql
        );

      const elapsed =
        performance.now() -
        start;

      const rows =
        resultToRows(
          result
        );

      setStatus(
        "✅ Consulta ejecutada correctamente.\n" +
        `Filas resultado: ${rows.length}\n` +
        `Tiempo: ${elapsed.toFixed(2)} ms`,
        "ok"
      );

      setResult({
        procesamiento: "LOCAL",
        sql: sql,
        filas_resultado:
          rows.length,
        tiempo_ms:
          Number(
            elapsed.toFixed(2)
          ),
        resultado:
          rows
      });

    } catch (error) {

      console.error(
        `[${APP_ID}] SQL ERROR`,
        error
      );

      setStatus(
        "❌ ERROR EN LA CONSULTA SQL",
        "error"
      );

      setResult({
        procesamiento:
          "LOCAL",

        sql:
          sql,

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
     INICIO
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
