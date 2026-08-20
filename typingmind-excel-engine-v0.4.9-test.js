/*
 * TypingMind Excel Data Engine - Extension v0.4.9 TEST
 *
 * OBJETIVO:
 * - Mantener la carga local de XLSX.
 * - Mantener DuckDB-Wasm 1.29.0 / DuckDB v1.1.1.
 * - Corregir la serialización de resultados DuckDB.
 * - Convertir BigInt, TypedArray, Date y otros tipos a JSON limpio.
 * - Probar COUNT, SUM, AVG y GROUP BY.
 *
 * IMPORTANTE:
 * - El archivo Excel permanece LOCAL.
 * - No se envía el Excel a ningún servidor.
 * - Esta versión todavía es una versión de prueba.
 */

(() => {
  "use strict";

  const APP_ID = "tm-excel-engine-v049-test";

  const DUCKDB_PACKAGE =
    "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

  let duckdb = null;
  let db = null;
  let conn = null;

  let currentFile = null;
  let currentFileName = null;

  /* ============================================================
     ESTILOS
     ============================================================ */

  function addStyles() {
    if (document.getElementById("tmxe-v049-style")) return;

    const style = document.createElement("style");

    style.id = "tmxe-v049-style";

    style.textContent = `
      #tmxe-v049-button {
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

        box-shadow: 0 6px 20px rgba(0,0,0,.25);

        cursor: pointer;
      }

      #tmxe-v049-overlay {
        position: fixed;
        inset: 0;

        z-index: 2147483001;

        background: rgba(0,0,0,.48);

        display: flex;
        align-items: center;
        justify-content: center;

        padding: 16px;
      }

      #tmxe-v049-panel {
        width: min(980px, 96vw);
        max-height: 92vh;

        overflow: auto;

        background: Canvas;
        color: CanvasText;

        border: 1px solid rgba(127,127,127,.35);
        border-radius: 16px;

        padding: 20px;

        box-shadow: 0 20px 70px rgba(0,0,0,.35);

        font: 14px/1.45 system-ui, sans-serif;
      }

      #tmxe-v049-panel h2 {
        margin: 0 0 5px;
        font-size: 20px;
      }

      #tmxe-v049-help {
        opacity: .7;
        margin-bottom: 15px;
      }

      #tmxe-v049-status {
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

      #tmxe-v049-status[data-kind="ok"] {
        color: #16a34a;
      }

      #tmxe-v049-status[data-kind="error"] {
        color: #dc2626;
      }

      #tmxe-v049-result {
        white-space: pre-wrap;

        overflow: auto;

        max-height: 520px;

        padding: 12px;

        border-radius: 10px;

        background: rgba(127,127,127,.10);

        font: 12px/1.45 ui-monospace,
              SFMono-Regular,
              Menlo,
              monospace;
      }

      #tmxe-v049-run,
      #tmxe-v049-close,
      #tmxe-v049-test {
        border: 1px solid rgba(127,127,127,.45);

        border-radius: 9px;

        padding: 9px 13px;

        background: transparent;

        color: inherit;

        cursor: pointer;
      }

      #tmxe-v049-run {
        background: #7c3aed !important;
        color: white !important;
        border-color: #7c3aed !important;
      }

      #tmxe-v049-test {
        background: #2563eb !important;
        color: white !important;
        border-color: #2563eb !important;
      }

      #tmxe-v049-actions {
        display: flex;

        gap: 8px;

        flex-wrap: wrap;

        margin: 10px 0;
      }

      #tmxe-v049-file {
        margin: 8px 0;
      }
    `;

    document.head.appendChild(style);
  }

  /* ============================================================
     BOTÓN
     ============================================================ */

  function createButton() {
    if (document.getElementById("tmxe-v049-button")) return;

    const button = document.createElement("button");

    button.id = "tmxe-v049-button";

    button.type = "button";

    button.textContent = "📊 Excel v0.4.9";

    button.title = "Excel Data Engine v0.4.9 TEST";

    button.addEventListener("click", openPanel);

    document.body.appendChild(button);
  }

  /* ============================================================
     PANEL
     ============================================================ */

  function openPanel() {
    if (document.getElementById("tmxe-v049-overlay")) {
      return;
    }

    const overlay = document.createElement("div");

    overlay.id = "tmxe-v049-overlay";

    overlay.innerHTML = `
      <div
        id="tmxe-v049-panel"
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
              📊 Excel Engine v0.4.9
            </h2>

            <div id="tmxe-v049-help">
              Prueba de serialización y consultas SQL locales.
            </div>

          </div>

          <button
            id="tmxe-v049-close"
            type="button"
          >
            Cerrar
          </button>

        </div>

        <input
          id="tmxe-v049-file"
          type="file"
          accept=".xlsx,.xls"
        />

        <div id="tmxe-v049-actions">

          <button
            id="tmxe-v049-run"
            type="button"
          >
            📥 Cargar Excel
          </button>

          <button
            id="tmxe-v049-test"
            type="button"
            disabled
          >
            🧪 Probar SQL
          </button>

        </div>

        <div id="tmxe-v049-status">
          Selecciona un archivo Excel.
        </div>

        <strong>Resultado</strong>

        <pre id="tmxe-v049-result">—</pre>

        <div style="
          margin-top:12px;
          opacity:.65;
          font-size:12px;
        ">
          v0.4.9 TEST — procesamiento local.
        </div>

      </div>
    `;

    document.body.appendChild(overlay);

    document
      .getElementById("tmxe-v049-close")
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
      .getElementById("tmxe-v049-run")
      .addEventListener(
        "click",
        loadExcel
      );

    document
      .getElementById("tmxe-v049-test")
      .addEventListener(
        "click",
        runTests
      );
  }

  /* ============================================================
     UI
     ============================================================ */

  function setStatus(text, kind = "") {
    const el =
      document.getElementById(
        "tmxe-v049-status"
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
        "tmxe-v049-result"
      );

    if (!el) return;

    try {
      el.textContent =
        typeof value === "string"
          ? value
          : JSON.stringify(
              normalizeForJSON(value),
              null,
              2
            );
    } catch (error) {
      el.textContent =
        "ERROR SERIALIZANDO RESULTADO:\n\n" +
        String(error);
    }
  }

  /* ============================================================
     NORMALIZADOR
     ============================================================ */

  function normalizeForJSON(value, seen = new WeakSet()) {

    if (value === null || value === undefined) {
      return value;
    }

    /*
     * BigInt
     */
    if (typeof value === "bigint") {

      const numberValue = Number(value);

      if (
        Number.isSafeInteger(numberValue)
      ) {
        return numberValue;
      }

      return value.toString();
    }

    /*
     * Number
     */
    if (typeof value === "number") {

      if (!Number.isFinite(value)) {
        return null;
      }

      return value;
    }

    /*
     * String / Boolean
     */
    if (
      typeof value === "string" ||
      typeof value === "boolean"
    ) {
      return value;
    }

    /*
     * Date
     */
    if (value instanceof Date) {

      if (Number.isNaN(value.getTime())) {
        return null;
      }

      return value.toISOString();
    }

    /*
     * TypedArray / ArrayBuffer
     *
     * DuckDB-Wasm puede devolver algunos valores
     * mediante estructuras binarias.
     */
    if (
      ArrayBuffer.isView(value) &&
      !(value instanceof DataView)
    ) {

      return Array.from(value).map(
        item => normalizeForJSON(item, seen)
      );
    }

    if (value instanceof ArrayBuffer) {
      return Array.from(
        new Uint8Array(value)
      );
    }

    /*
     * Array
     */
    if (Array.isArray(value)) {

      return value.map(
        item => normalizeForJSON(item, seen)
      );
    }

    /*
     * Objetos
     */
    if (typeof value === "object") {

      if (seen.has(value)) {
        return "[Circular]";
      }

      seen.add(value);

      const output = {};

      for (const key of Object.keys(value)) {

        output[key] =
          normalizeForJSON(
            value[key],
            seen
          );
      }

      return output;
    }

    /*
     * Fallback
     */
    return String(value);
  }

  /* ============================================================
     DUCKDB
     ============================================================ */

  async function initializeDuckDB() {

    if (db && conn) {
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
  }

  /* ============================================================
     EXCEL
     ============================================================ */

  async function loadExcel() {

    const fileInput =
      document.getElementById(
        "tmxe-v049-file"
      );

    if (
      !fileInput ||
      !fileInput.files ||
      !fileInput.files.length
    ) {

      setStatus(
        "Selecciona primero un archivo XLSX.",
        "error"
      );

      return;
    }

    try {

      currentFile =
        fileInput.files[0];

      currentFileName =
        currentFile.name;

      await initializeDuckDB();

      setStatus(
        "4/5 — Leyendo Excel localmente..."
      );

      const arrayBuffer =
        await currentFile.arrayBuffer();

      /*
       * Importamos XLSX mediante SheetJS.
       *
       * El archivo permanece en memoria local.
       */
      const XLSX =
        await import(
          "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm"
        );

      const workbook =
        XLSX.read(
          arrayBuffer,
          {
            type: "array",
            cellDates: true
          }
        );

      const sheetNames =
        workbook.SheetNames;

      if (!sheetNames.length) {
        throw new Error(
          "El archivo no contiene hojas."
        );
      }

      const sheetName =
        sheetNames[0];

      const sheet =
        workbook.Sheets[sheetName];

      const rows =
        XLSX.utils.sheet_to_json(
          sheet,
          {
            header: 1,
            defval: null,
            raw: true
          }
        );

      if (!rows.length) {
        throw new Error(
          "La hoja está vacía."
        );
      }

      /*
       * Primera fila = encabezados.
       */
      const rawHeaders =
        rows[0];

      const headers =
        rawHeaders.map(
          (header, index) => {

            let name =
              header === null ||
              header === undefined ||
              String(header).trim() === ""
                ? `Column_${index + 1}`
                : String(header).trim();

            return name;
          }
        );

      /*
       * Elimina filas completamente vacías.
       */
      const dataRows =
        rows
          .slice(1)
          .filter(row =>
            row.some(
              value =>
                value !== null &&
                value !== undefined &&
                String(value).trim() !== ""
            )
          );

      /*
       * Inferencia sencilla de tipos.
       */
      const inferred =
        inferColumnTypes(
          headers,
          dataRows
        );

      /*
       * Crear tabla SQL.
       */
      await conn.query(
        `DROP TABLE IF EXISTS excel_data;`
      );

      const columnDefinitions =
        headers.map(
          (header, index) =>
            `"${escapeIdentifier(header)}" ${inferred[index].sqlType}`
        );

      await conn.query(
        `CREATE TABLE excel_data (${columnDefinitions.join(", ")});`
      );

      /*
       * Insertar filas.
       */
      for (const row of dataRows) {

        const values =
          headers.map(
            (_, index) =>
              row[index] ?? null
          );

        const sqlValues =
          values.map(
            value =>
              sqlLiteral(value)
          );

        await conn.query(
          `INSERT INTO excel_data VALUES (${sqlValues.join(", ")});`
        );
      }

      setStatus(
        "5/5 — Excel cargado correctamente.",
        "ok"
      );

      const countResult =
        await conn.query(
          `SELECT COUNT(*) AS registros
           FROM excel_data;`
        );

      const previewResult =
        await conn.query(
          `SELECT *
           FROM excel_data
           LIMIT 10;`
        );

      const schemaResult =
        await conn.query(
          `DESCRIBE excel_data;`
        );

      const result = {

        procesamiento: "LOCAL",

        archivo:
          currentFileName,

        tamano_bytes:
          currentFile.size,

        hojas:
          sheetNames,

        hoja_principal:
          sheetName,

        filas_fisicas_detectadas:
          rows.length,

        filas_reales:
          dataRows.length,

        filas_vacias_ignoradas:
          rows.length -
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
          "excel_data",

        count:
          normalizeQueryResult(
            countResult
          ),

        schema:
          normalizeQueryResult(
            schemaResult
          ),

        preview:
          normalizeQueryResult(
            previewResult
          )
      };

      setResult(result);

      const testButton =
        document.getElementById(
          "tmxe-v049-test"
        );

      if (testButton) {
        testButton.disabled = false;
      }

    } catch (error) {

      console.error(
        `[${APP_ID}]`,
        error
      );

      setStatus(
        "❌ ERROR AL CARGAR EL EXCEL",
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

  /* ============================================================
     INFERENCIA DE TIPOS
     ============================================================ */

  function inferColumnTypes(
    headers,
    rows
  ) {

    return headers.map(
      (header, columnIndex) => {

        const values =
          rows
            .map(row => row[columnIndex])
            .filter(
              value =>
                value !== null &&
                value !== undefined &&
                String(value).trim() !== ""
            );

        if (!values.length) {

          return {
            column_index:
              columnIndex,

            column_name:
              header,

            duckdb_type:
              "VARCHAR",

            sqlType:
              "VARCHAR",

            confidence:
              "low",

            reason:
              "columna sin valores",

            non_empty_values:
              0
          };
        }

        /*
         * Fecha
         */
        const dateLike =
          values.every(
            value =>
              value instanceof Date ||
              isDateLike(value)
          );

        if (
          dateLike &&
          /fecha|date|día|dia/i.test(header)
        ) {

          return {
            column_index:
              columnIndex,

            column_name:
              header,

            duckdb_type:
              "DATE",

            sqlType:
              "DATE",

            confidence:
              "high",

            reason:
              "encabezado y valores compatibles con fecha",

            non_empty_values:
              values.length
          };
        }

        /*
         * Enteros
         */
        const integers =
          values.every(
            value =>
              typeof value === "number" &&
              Number.isInteger(value)
          );

        if (integers) {

          return {
            column_index:
              columnIndex,

            column_name:
              header,

            duckdb_type:
              "BIGINT",

            sqlType:
              "BIGINT",

            confidence:
              "high",

            reason:
              "todos los valores son enteros",

            non_empty_values:
              values.length
          };
        }

        /*
         * Decimal
         */
        const numbers =
          values.every(
            value =>
              typeof value === "number" &&
              Number.isFinite(value)
          );

        if (numbers) {

          return {
            column_index:
              columnIndex,

            column_name:
              header,

            duckdb_type:
              "DOUBLE",

            sqlType:
              "DOUBLE",

            confidence:
              "high",

            reason:
              "todos los valores son numéricos",

            non_empty_values:
              values.length
          };
        }

        /*
         * Texto
         */
        return {
          column_index:
            columnIndex,

          column_name:
            header,

          duckdb_type:
            "VARCHAR",

          sqlType:
            "VARCHAR",

          confidence:
            "medium",

          reason:
            "valores tratados como texto",

          non_empty_values:
            values.length
        };
      }
    );
  }

  function isDateLike(value) {

    if (value instanceof Date) {
      return !Number.isNaN(
        value.getTime()
      );
    }

    if (typeof value !== "string") {
      return false;
    }

    const trimmed =
      value.trim();

    if (!trimmed) {
      return false;
    }

    return /^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/.test(
      trimmed
    );
  }

  /* ============================================================
     SQL HELPERS
     ============================================================ */

  function escapeIdentifier(value) {

    return String(value)
      .replace(/"/g, '""');
  }

  function sqlLiteral(value) {

    if (
      value === null ||
      value === undefined ||
      value === ""
    ) {
      return "NULL";
    }

    if (value instanceof Date) {

      const iso =
        value
          .toISOString()
          .slice(0, 10);

      return `'${iso}'::DATE`;
    }

    if (
      typeof value === "number" &&
      Number.isFinite(value)
    ) {
      return String(value);
    }

    if (typeof value === "bigint") {
      return value.toString();
    }

    const text =
      String(value)
        .replace(/'/g, "''");

    return `'${text}'`;
  }

  /* ============================================================
     RESULTADOS DUCKDB
     ============================================================ */

  function normalizeQueryResult(result) {

    if (!result) {
      return [];
    }

    let rows;

    try {

      if (
        typeof result.toArray === "function"
      ) {

        rows =
          result.toArray();

      } else if (
        Array.isArray(result)
      ) {

        rows = result;

      } else {

        rows = [result];
      }

    } catch (error) {

      return {
        error:
          "No se pudo convertir el resultado DuckDB",

        detalle:
          String(error)
      };
    }

    return normalizeForJSON(rows);
  }

  /* ============================================================
     PRUEBAS SQL
     ============================================================ */

  async function executeSQL(sql) {

    if (!conn) {
      throw new Error(
        "DuckDB no está conectado."
      );
    }

    const result =
      await conn.query(sql);

    return {
      procesamiento:
        "LOCAL",

      sql,

      filas_resultado:
        result.numRows !== undefined
          ? Number(result.numRows)
          : undefined,

      resultado:
        normalizeQueryResult(result)
    };
  }

  async function runTests() {

    try {

      setStatus(
        "Ejecutando pruebas SQL..."
      );

      const tests = {};

      /*
       * TEST 1
       */
      tests.count =
        await executeSQL(`
          SELECT
            COUNT(*) AS registros
          FROM excel_data;
        `);

      /*
       * TEST 2
       */
      tests.sum =
        await executeSQL(`
          SELECT
            SUM("Total TEU's") AS total
          FROM excel_data;
        `);

      /*
       * TEST 3
       */
      tests.average =
        await executeSQL(`
          SELECT
            AVG("Total TEU's") AS promedio
          FROM excel_data;
        `);

      /*
       * TEST 4
       */
      tests.year_group =
        await executeSQL(`
          SELECT
            EXTRACT(YEAR FROM "Fecha") AS año,
            SUM("Total TEU's") AS total
          FROM excel_data
          GROUP BY año
          ORDER BY año;
        `);

      /*
       * TEST 5
       */
      tests.preview =
        await executeSQL(`
          SELECT *
          FROM excel_data
          LIMIT 10;
        `);

      setStatus(
        "✅ Pruebas terminadas correctamente.",
        "ok"
      );

      setResult(tests);

    } catch (error) {

      console.error(
        `[${APP_ID}]`,
        error
      );

      setStatus(
        "❌ ERROR EN LAS PRUEBAS SQL",
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

  /* ============================================================
     INIT
     ============================================================ */

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