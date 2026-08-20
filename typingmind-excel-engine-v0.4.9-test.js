/*
 * TypingMind Excel Data Engine - Extension v0.4.9 TEST
 *
 * OBJETIVO:
 * - Cargar XLSX localmente en el navegador.
 * - NO utilizar la extensión Excel de DuckDB.
 * - Procesar el XLSX completamente de forma LOCAL.
 * - Crear tabla DuckDB: excel_data
 * - Inferir tipos.
 * - Convertir fechas correctamente a DATE.
 * - Corregir serialización de BIGINT.
 * - Ejecutar consultas SQL de prueba.
 *
 * v0.4.9 TEST
 */

(() => {
  "use strict";

  const APP_ID = "tm-excel-engine-v049-test";

  const DUCKDB_PACKAGE =
    "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

  const XLSX_PACKAGE =
    "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";

  let duckdb = null;
  let XLSX = null;

  let db = null;
  let conn = null;

  let currentFile = null;
  let currentTable = "excel_data";

  // ============================================================
  // UTILIDADES
  // ============================================================

  function safeString(value) {
    if (value === null || value === undefined) {
      return null;
    }

    if (typeof value === "string") {
      return value;
    }

    if (typeof value === "bigint") {
      return value.toString();
    }

    if (value instanceof Date) {
      return value.toISOString();
    }

    return String(value);
  }

  function serializeValue(value) {
    if (value === null || value === undefined) {
      return null;
    }

    if (typeof value === "bigint") {
      return value.toString();
    }

    if (value instanceof Date) {
      return value.toISOString();
    }

    if (ArrayBuffer.isView(value)) {
      return Array.from(value).map(v =>
        typeof v === "bigint" ? v.toString() : v
      );
    }

    if (value instanceof ArrayBuffer) {
      return Array.from(new Uint8Array(value));
    }

    if (Array.isArray(value)) {
      return value.map(serializeValue);
    }

    if (typeof value === "object") {
      const out = {};

      for (const [key, val] of Object.entries(value)) {
        out[key] = serializeValue(val);
      }

      return out;
    }

    return value;
  }

  function safeJSONStringify(value) {
    return JSON.stringify(
      serializeValue(value),
      null,
      2
    );
  }

  function escapeIdentifier(name) {
    return `"${String(name).replace(/"/g, '""')}"`;
  }

  function escapeSQLString(value) {
    return String(value)
      .replace(/\\/g, "\\\\")
      .replace(/'/g, "''");
  }

  function isEmpty(value) {
    return (
      value === null ||
      value === undefined ||
      String(value).trim() === ""
    );
  }

  function normalizeHeader(value, index) {
    let header = safeString(value);

    if (!header || !header.trim()) {
      header = `Column_${index + 1}`;
    }

    return header.trim();
  }

  // ============================================================
  // FECHAS
  // ============================================================

  function excelSerialToDate(serial) {
    if (
      typeof serial !== "number" ||
      !Number.isFinite(serial)
    ) {
      return null;
    }

    // Sistema de fechas de Excel 1900.
    const utcDays = Math.floor(serial - 25569);

    const utcValue =
      utcDays * 86400 * 1000;

    const date = new Date(utcValue);

    if (Number.isNaN(date.getTime())) {
      return null;
    }

    return date;
  }

  function dateToSQL(date) {
    if (!(date instanceof Date)) {
      return null;
    }

    if (Number.isNaN(date.getTime())) {
      return null;
    }

    const y = date.getUTCFullYear();
    const m = String(
      date.getUTCMonth() + 1
    ).padStart(2, "0");

    const d = String(
      date.getUTCDate()
    ).padStart(2, "0");

    return `${y}-${m}-${d}`;
  }

  function normalizeDateValue(value) {
    if (value instanceof Date) {
      return dateToSQL(value);
    }

    if (typeof value === "number") {
      const date =
        excelSerialToDate(value);

      return date
        ? dateToSQL(date)
        : null;
    }

    const text =
      String(value ?? "").trim();

    if (!text) {
      return null;
    }

    // YYYY-MM-DD
    if (
      /^\d{4}-\d{1,2}-\d{1,2}$/.test(text)
    ) {
      const parts = text.split("-");

      return (
        `${parts[0]}-` +
        `${String(parts[1]).padStart(2, "0")}-` +
        `${String(parts[2]).padStart(2, "0")}`
      );
    }

    // DD/MM/YYYY
    let match =
      text.match(
        /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
      );

    if (match) {
      const day = match[1];
      const month = match[2];
      const year = match[3];

      return (
        `${year}-` +
        `${String(month).padStart(2, "0")}-` +
        `${String(day).padStart(2, "0")}`
      );
    }

    const parsed =
      new Date(text);

    if (!Number.isNaN(parsed.getTime())) {
      return dateToSQL(parsed);
    }

    return null;
  }

  // ============================================================
  // INFERENCIA DE TIPOS
  // ============================================================

  function looksLikeInteger(value) {
    if (typeof value === "number") {
      return (
        Number.isFinite(value) &&
        Number.isInteger(value)
      );
    }

    const text =
      String(value ?? "").trim();

    return /^[-+]?\d+$/.test(text);
  }

  function looksLikeNumber(value) {
    if (typeof value === "number") {
      return Number.isFinite(value);
    }

    const text =
      String(value ?? "")
        .trim()
        .replace(/,/g, "");

    if (!text) return false;

    return (
      /^[-+]?(?:\d+\.?\d*|\.\d+)$/.test(text)
    );
  }

  function looksLikeDate(value, header) {
    const headerText =
      String(header ?? "").toLowerCase();

    const dateHeader =
      /fecha|date|day|mes|month|año|year/.test(
        headerText
      );

    if (value instanceof Date) {
      return true;
    }

    if (
      dateHeader &&
      typeof value === "number" &&
      value > 20000 &&
      value < 100000
    ) {
      return true;
    }

    if (typeof value === "string") {
      const text = value.trim();

      if (
        /^\d{4}-\d{1,2}-\d{1,2}$/.test(text) ||
        /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(text)
      ) {
        return true;
      }
    }

    return false;
  }

  function inferColumnType(
    header,
    values,
    columnIndex
  ) {
    const nonEmpty =
      values.filter(v => !isEmpty(v));

    if (!nonEmpty.length) {
      return {
        column_index: columnIndex,
        column_name: header,
        duckdb_type: "VARCHAR",
        confidence: "low",
        reason: "columna vacía",
        non_empty_values: 0
      };
    }

    const dateMatches =
      nonEmpty.filter(v =>
        looksLikeDate(v, header)
      ).length;

    const integerMatches =
      nonEmpty.filter(
        looksLikeInteger
      ).length;

    const numberMatches =
      nonEmpty.filter(
        looksLikeNumber
      ).length;

    const total = nonEmpty.length;

    // FECHA
    if (dateMatches === total) {
      return {
        column_index: columnIndex,
        column_name: header,
        duckdb_type: "DATE",
        confidence: "high",
        reason:
          "encabezado y valores compatibles con fecha",
        non_empty_values: total
      };
    }

    // ENTERO
    if (integerMatches === total) {
      return {
        column_index: columnIndex,
        column_name: header,
        duckdb_type: "BIGINT",
        confidence: "high",
        reason:
          "todos los valores son enteros",
        non_empty_values: total
      };
    }

    // DECIMAL
    if (numberMatches === total) {
      return {
        column_index: columnIndex,
        column_name: header,
        duckdb_type: "DOUBLE",
        confidence: "high",
        reason:
          "todos los valores son numéricos",
        non_empty_values: total
      };
    }

    return {
      column_index: columnIndex,
      column_name: header,
      duckdb_type: "VARCHAR",
      confidence: "medium",
      reason:
        "valores mixtos o no numéricos",
      non_empty_values: total
    };
  }

  // ============================================================
  // CONVERSIÓN PARA DUCKDB
  // ============================================================

  function valueToSQL(
    value,
    type
  ) {
    if (isEmpty(value)) {
      return "NULL";
    }

    if (type === "DATE") {
      const date =
        normalizeDateValue(value);

      if (!date) {
        return "NULL";
      }

      return `DATE '${date}'`;
    }

    if (type === "BIGINT") {
      let text =
        String(value)
          .trim()
          .replace(/,/g, "");

      if (
        !/^[+-]?\d+$/.test(text)
      ) {
        return "NULL";
      }

      return text;
    }

    if (type === "DOUBLE") {
      let text =
        String(value)
          .trim()
          .replace(/,/g, "");

      const number =
        Number(text);

      if (!Number.isFinite(number)) {
        return "NULL";
      }

      return String(number);
    }

    return `'${escapeSQLString(
      String(value)
    )}'`;
  }

  // ============================================================
  // UI
  // ============================================================

  function addStyles() {
    if (
      document.getElementById(
        "tmxe-v049-style"
      )
    ) {
      return;
    }

    const style =
      document.createElement("style");

    style.id =
      "tmxe-v049-style";

    style.textContent = `
      #tmxe-v049-button {
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
        width: min(1000px, 96vw);
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

        max-height: 600px;

        padding: 12px;

        border-radius: 10px;

        background: rgba(127,127,127,.10);

        font: 12px/1.45 ui-monospace,
          SFMono-Regular,
          Menlo,
          monospace;
      }

      #tmxe-v049-actions {
        display: flex;
        gap: 8px;
        flex-wrap: wrap;
        margin: 10px 0;
      }

      #tmxe-v049-run,
      #tmxe-v049-close,
      #tmxe-v049-query {
        border: 1px solid rgba(127,127,127,.45);

        border-radius: 9px;

        padding: 9px 13px;

        background: transparent;

        color: inherit;

        cursor: pointer;
      }

      #tmxe-v049-run {
        background: #2563eb !important;
        color: white !important;
        border-color: #2563eb !important;
      }

      #tmxe-v049-file {
        display: none;
      }

      #tmxe-v049-file-label {
        display: inline-block;

        border: 1px solid rgba(127,127,127,.45);

        border-radius: 9px;

        padding: 9px 13px;

        cursor: pointer;
      }
    `;

    document.head.appendChild(style);
  }

  function createButton() {
    if (
      document.getElementById(
        "tmxe-v049-button"
      )
    ) {
      return;
    }

    const button =
      document.createElement("button");

    button.id =
      "tmxe-v049-button";

    button.type = "button";

    button.textContent =
      "📊 Excel v0.4.9";

    button.title =
      "Excel Data Engine v0.4.9";

    button.addEventListener(
      "click",
      openPanel
    );

    document.body.appendChild(button);
  }

  function openPanel() {
    if (
      document.getElementById(
        "tmxe-v049-overlay"
      )
    ) {
      return;
    }

    const overlay =
      document.createElement("div");

    overlay.id =
      "tmxe-v049-overlay";

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
              📊 Excel Data Engine v0.4.9
            </h2>

            <div id="tmxe-v049-help">
              Diagnóstico de tipos, fechas y
              serialización DuckDB-Wasm.
            </div>
          </div>

          <button
            id="tmxe-v049-close"
            type="button"
          >
            Cerrar
          </button>

        </div>

        <div id="tmxe-v049-status">
          Selecciona un archivo Excel.
        </div>

        <div id="tmxe-v049-actions">

          <label
            id="tmxe-v049-file-label"
            for="tmxe-v049-file"
          >
            📁 Seleccionar Excel
          </label>

          <input
            id="tmxe-v049-file"
            type="file"
            accept=".xlsx,.xls"
          />

          <button
            id="tmxe-v049-run"
            type="button"
          >
            🔍 Cargar y diagnosticar
          </button>

        </div>

        <strong>
          Resultado del diagnóstico
        </strong>

        <pre id="tmxe-v049-result">—</pre>

        <div style="
          margin-top:12px;
          opacity:.65;
          font-size:12px;
        ">
          v0.4.9 TEST — procesamiento 100% local.
        </div>

      </div>
    `;

    document.body.appendChild(overlay);

    document
      .getElementById(
        "tmxe-v049-close"
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
        "tmxe-v049-run"
      )
      .addEventListener(
        "click",
        loadExcel
      );
  }

  function setStatus(
    text,
    kind = ""
  ) {
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

    el.textContent =
      typeof value === "string"
        ? value
        : safeJSONStringify(value);
  }

  // ============================================================
  // DUCKDB
  // ============================================================

  async function initDuckDB() {
    if (
      db &&
      conn
    ) {
      return;
    }

    setStatus(
      "1/8 — Cargando DuckDB-Wasm..."
    );

    duckdb =
      await import(
        DUCKDB_PACKAGE
      );

    setStatus(
      "2/8 — Seleccionando bundle..."
    );

    const bundles =
      duckdb.getJsDelivrBundles();

    const bundle =
      await duckdb.selectBundle(
        bundles
      );

    setStatus(
      "3/8 — Inicializando WebAssembly..."
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

    return bundle;
  }

  // ============================================================
  // XLSX
  // ============================================================

  async function loadXLSX() {
    setStatus(
      "4/8 — Cargando lector XLSX..."
    );

    XLSX =
      await import(
        XLSX_PACKAGE
      );
  }

  async function readWorkbook(file) {
    const buffer =
      await file.arrayBuffer();

    const workbook =
      XLSX.read(
        buffer,
        {
          type: "array",
          cellDates: true,
          cellNF: true,
          cellText: false
        }
      );

    return workbook;
  }

  function extractSheetData(
    workbook
  ) {
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
          raw: true,
          blankrows: true
        }
      );

    const physicalRows =
      matrix.length;

    // Buscar la primera fila que tenga
    // contenido real.
    let headerIndex = -1;

    for (
      let i = 0;
      i < matrix.length;
      i++
    ) {
      const row =
        matrix[i] || [];

      if (
        row.some(
          cell => !isEmpty(cell)
        )
      ) {
        headerIndex = i;
        break;
      }
    }

    if (headerIndex < 0) {
      throw new Error(
        "No se encontró una fila de encabezados."
      );
    }

    const rawHeaders =
      matrix[headerIndex];

    const headers =
      rawHeaders.map(
        normalizeHeader
      );

    const rows = [];

    let emptyRows = 0;

    for (
      let i = headerIndex + 1;
      i < matrix.length;
      i++
    ) {
      const sourceRow =
        matrix[i] || [];

      const row =
        headers.map(
          (_, columnIndex) =>
            sourceRow[columnIndex] ??
            null
        );

      const hasData =
        row.some(
          value => !isEmpty(value)
        );

      if (!hasData) {
        emptyRows++;
        continue;
      }

      rows.push(row);
    }

    return {
      sheets,
      sheetName,
      physicalRows,
      headerIndex,
      headers,
      rows,
      emptyRows
    };
  }

  // ============================================================
  // CREACIÓN TABLA
  // ============================================================

  async function createTable(
    headers,
    rows,
    inferredTypes
  ) {
    await conn.query(
      `DROP TABLE IF EXISTS ${currentTable};`
    );

    const definitions =
      headers.map(
        (header, index) => {
          const type =
            inferredTypes[index]
              .duckdb_type;

          return (
            `${escapeIdentifier(header)} ${type}`
          );
        }
      );

    await conn.query(
      `CREATE TABLE ${currentTable} (${definitions.join(", ")});`
    );

    // Insertar en lotes.
    const BATCH_SIZE = 250;

    for (
      let start = 0;
      start < rows.length;
      start += BATCH_SIZE
    ) {
      const batch =
        rows.slice(
          start,
          start + BATCH_SIZE
        );

      const valuesSQL =
        batch.map(row => {
          const values =
            headers.map(
              (_, columnIndex) =>
                valueToSQL(
                  row[columnIndex],
                  inferredTypes[
                    columnIndex
                  ].duckdb_type
                )
            );

          return `(${values.join(", ")})`;
        });

      await conn.query(
        `INSERT INTO ${currentTable} VALUES ${valuesSQL.join(", ")};`
      );
    }
  }

  // ============================================================
  // DIAGNÓSTICO
  // ============================================================

  async function runDiagnostics(
    bundle,
    workbookInfo,
    inferredTypes
  ) {
    setStatus(
      "7/8 — Ejecutando pruebas SQL..."
    );

    const versionResult =
      await conn.query(
        "SELECT version() AS duckdb_version;"
      );

    const countResult =
      await conn.query(
        `SELECT COUNT(*) AS registros FROM ${currentTable};`
      );

    const schemaResult =
      await conn.query(
        `DESCRIBE ${currentTable};`
      );

    const typeResult =
      await conn.query(`
        SELECT
          typeof("Fecha") AS tipo_fecha,
          typeof("Local") AS tipo_local,
          typeof("Transshipment") AS tipo_transshipment,
          typeof("Total TEU's") AS tipo_total
        FROM ${currentTable}
        LIMIT 1;
      `);

    const dateTestResult =
      await conn.query(`
        SELECT
          MIN("Fecha") AS fecha_min,
          MAX("Fecha") AS fecha_max,
          EXTRACT(YEAR FROM "Fecha")
            AS año
        FROM ${currentTable}
        GROUP BY año
        ORDER BY año;
      `);

    const yearlyResult =
      await conn.query(`
        SELECT
          EXTRACT(YEAR FROM "Fecha") AS año,
          SUM("Total TEU's") AS total
        FROM ${currentTable}
        GROUP BY año
        ORDER BY año;
      `);

    const previewResult =
      await conn.query(`
        SELECT *
        FROM ${currentTable}
        LIMIT 10;
      `);

    const summaryResult =
      await conn.query(`
        SELECT
          SUM("Local") AS "Local__sum",
          AVG("Local") AS "Local__avg",
          MIN("Local") AS "Local__min",
          MAX("Local") AS "Local__max",

          SUM("Transshipment")
            AS "Transshipment__sum",
          AVG("Transshipment")
            AS "Transshipment__avg",
          MIN("Transshipment")
            AS "Transshipment__min",
          MAX("Transshipment")
            AS "Transshipment__max",

          SUM("Total TEU's")
            AS "Total TEU's__sum",
          AVG("Total TEU's")
            AS "Total TEU's__avg",
          MIN("Total TEU's")
            AS "Total TEU's__min",
          MAX("Total TEU's")
            AS "Total TEU's__max"

        FROM ${currentTable};
      `);

    setStatus(
      "8/8 — Diagnóstico terminado.",
      "ok"
    );

    return {
      engine:
        APP_ID,

      procesamiento:
        "LOCAL",

      archivo:
        currentFile?.name || null,

      tamano_bytes:
        currentFile?.size || null,

      hojas:
        workbookInfo.sheets,

      hoja_principal:
        workbookInfo.sheetName,

      filas_fisicas_detectadas:
        workbookInfo.physicalRows,

      fila_encabezados:
        workbookInfo.headerIndex + 1,

      filas_reales:
        workbookInfo.rows.length,

      filas_vacias_ignoradas:
        workbookInfo.emptyRows,

      filas_insertadas:
        countResult.toArray(),

      columnas_detectadas:
        workbookInfo.headers.length,

      encabezados:
        workbookInfo.headers,

      duckdb_table:
        currentTable,

      duckdb_version:
        versionResult.toArray(),

      tipos_inferidos:
        inferredTypes,

      schema_duckdb:
        schemaResult.toArray(),

      prueba_tipos_runtime:
        typeResult.toArray(),

      prueba_fechas:
        dateTestResult.toArray(),

      agrupacion_anual:
        yearlyResult.toArray(),

      resumen:
        summaryResult.toArray(),

      preview:
        previewResult.toArray(),

      bundle: {
        mainModule:
          bundle.mainModule,

        mainWorker:
          bundle.mainWorker,

        pthreadWorker:
          bundle.pthreadWorker
      }
    };
  }

  // ============================================================
  // CARGAR EXCEL
  // ============================================================

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
        "Selecciona primero un archivo Excel.",
        "error"
      );
      return;
    }

    currentFile =
      fileInput.files[0];

    try {
      setResult("");

      const bundle =
        await initDuckDB();

      await loadXLSX();

      setStatus(
        "5/8 — Leyendo archivo Excel localmente..."
      );

      const workbook =
        await readWorkbook(
          currentFile
        );

      const workbookInfo =
        extractSheetData(
          workbook
        );

      setStatus(
        "6/8 — Detectando tipos de columnas..."
      );

      const inferredTypes =
        workbookInfo.headers.map(
          (header, index) =>
            inferColumnType(
              header,
              workbookInfo.rows.map(
                row => row[index]
              ),
              index
            )
        );

      await createTable(
        workbookInfo.headers,
        workbookInfo.rows,
        inferredTypes
      );

      const result =
        await runDiagnostics(
          bundle,
          workbookInfo,
          inferredTypes
        );

      setResult(result);

      console.log(
        `[${APP_ID}]`,
        serializeValue(result)
      );

    } catch (error) {

      console.error(
        `[${APP_ID}]`,
        error
      );

      setStatus(
        "❌ ERROR DURANTE EL PROCESAMIENTO",
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

  // ============================================================
  // INIT
  // ============================================================

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
