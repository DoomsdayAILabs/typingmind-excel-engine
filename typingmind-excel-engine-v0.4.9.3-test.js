/*
 * TypingMind Excel Data Engine
 * Extension v0.4.9.3 TEST
 *
 * OBJETIVO
 * --------
 * Procesamiento local de archivos Excel mediante:
 *
 * XLSX -> JavaScript -> DuckDB-Wasm -> SQL
 *
 * El archivo Excel NO se envía al modelo.
 *
 * v0.4.9.3
 *
 * CAMBIOS:
 * - Corrige consulta pragma_version()
 * - DuckDB-Wasm 1.29.0
 * - DuckDB interno 1.1.1
 * - Serialización segura de BIGINT
 * - Lectura local XLSX
 * - Detección de filas reales
 * - Inferencia de tipos
 * - Creación de excel_data
 * - COUNT
 * - Preview
 * - Resumen
 * - Total por año
 * - SQL manual
 */

(() => {
  "use strict";

  const APP_ID = "tm-excel-engine-v0493-test";

  const DUCKDB_PACKAGE =
    "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

  const XLSX_PACKAGE =
    "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";

  let duckdb = null;
  let db = null;
  let conn = null;

  let workbook = null;
  let currentFile = null;

  let tableName = "excel_data";
  let headers = [];
  let rows = [];
  let inferredTypes = [];

  let initialized = false;

  /*
   * ------------------------------------------------------------
   * UTILIDADES
   * ------------------------------------------------------------
   */

  function $(id) {
    return document.getElementById(id);
  }

  function safeString(value) {
    if (value === null || value === undefined) {
      return null;
    }

    if (typeof value === "string") {
      return value;
    }

    if (typeof value === "number") {
      if (Number.isNaN(value)) return null;
      if (!Number.isFinite(value)) return String(value);
      return String(value);
    }

    if (typeof value === "bigint") {
      return value.toString();
    }

    if (value instanceof Date) {
      return value.toISOString();
    }

    return String(value);
  }

  /*
   * Convierte resultados de DuckDB a valores JSON seguros.
   *
   * Esto corrige casos como:
   *
   * [1535012,0,0,0]
   *
   * cuando DuckDB-Wasm devuelve ciertos BIGINT.
   */

  function normalizeValue(value) {
    if (value === null || value === undefined) {
      return value;
    }

    if (typeof value === "bigint") {
      const n = Number(value);

      if (
        Number.isSafeInteger(n) &&
        String(n) === value.toString()
      ) {
        return n;
      }

      return value.toString();
    }

    if (value instanceof Date) {
      return value.toISOString();
    }

    if (ArrayBuffer.isView(value)) {
      /*
       * Algunos resultados de DuckDB-Wasm pueden aparecer
       * como TypedArray.
       */

      if (
        value.length === 4 &&
        Number(value[1]) === 0 &&
        Number(value[2]) === 0 &&
        Number(value[3]) === 0
      ) {
        return normalizeValue(value[0]);
      }

      return Array.from(value).map(normalizeValue);
    }

    if (Array.isArray(value)) {
      /*
       * Caso especial observado:
       *
       * [1535012, 0, 0, 0]
       *
       * Lo interpretamos como un valor escalar
       * cuando solamente el primer elemento contiene datos.
       */

      if (
        value.length === 4 &&
        value.slice(1).every(v => {
          return Number(v) === 0;
        })
      ) {
        return normalizeValue(value[0]);
      }

      return value.map(normalizeValue);
    }

    if (typeof value === "object") {
      const result = {};

      for (const key of Object.keys(value)) {
        result[key] = normalizeValue(value[key]);
      }

      return result;
    }

    return value;
  }

  function rowsToObjects(result) {
    if (!result) return [];

    const rawRows = result.toArray();

    return rawRows.map(row => {
      const normalized = {};

      for (const key of Object.keys(row)) {
        normalized[key] = normalizeValue(row[key]);
      }

      return normalized;
    });
  }

  function stringifySafe(value) {
    return JSON.stringify(
      normalizeValue(value),
      null,
      2
    );
  }

  /*
   * Escapa identificadores SQL.
   *
   * Ejemplo:
   * Total TEU's
   *
   * se convierte en:
   * "Total TEU's"
   */

  function quoteIdentifier(name) {
    return `"${String(name).replace(/"/g, '""')}"`;
  }

  /*
   * Escapa valores para SQL.
   */

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
        value.toISOString().slice(0, 10);

      return `DATE '${iso}'`;
    }

    if (typeof value === "number") {
      if (!Number.isFinite(value)) {
        return "NULL";
      }

      return String(value);
    }

    if (typeof value === "bigint") {
      return value.toString();
    }

    const text = String(value)
      .replace(/\\/g, "\\\\")
      .replace(/'/g, "''");

    return `'${text}'`;
  }

  /*
   * Normaliza nombres de columnas duplicados.
   */

  function normalizeHeaders(rawHeaders) {
    const used = new Map();

    return rawHeaders.map((header, index) => {
      let name =
        header === null ||
        header === undefined ||
        String(header).trim() === ""
          ? `Column_${index + 1}`
          : String(header).trim();

      const base = name;

      const count =
        used.get(base) || 0;

      if (count > 0) {
        name = `${base}_${count + 1}`;
      }

      used.set(base, count + 1);

      return name;
    });
  }

  /*
   * ------------------------------------------------------------
   * FECHAS
   * ------------------------------------------------------------
   */

  function excelSerialToDate(serial) {
    if (
      typeof serial !== "number" ||
      !Number.isFinite(serial)
    ) {
      return null;
    }

    /*
     * Excel usa 1899-12-30 como referencia práctica
     * para el sistema de fechas 1900.
     */

    const epoch =
      Date.UTC(1899, 11, 30);

    const ms =
      epoch +
      Math.round(serial * 86400000);

    return new Date(ms);
  }

  function isDateHeader(header) {
    const h =
      String(header)
        .toLowerCase()
        .trim();

    return (
      h.includes("fecha") ||
      h.includes("date") ||
      h.includes("día") ||
      h.includes("dia")
    );
  }

  function normalizeDateValue(value) {
    if (value === null || value === undefined) {
      return null;
    }

    if (value instanceof Date) {
      if (Number.isNaN(value.getTime())) {
        return null;
      }

      return value;
    }

    if (typeof value === "number") {
      return excelSerialToDate(value);
    }

    const text =
      String(value).trim();

    if (!text) return null;

    /*
     * ISO
     */

    if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(text)) {
      const d = new Date(`${text}T00:00:00Z`);

      if (!Number.isNaN(d.getTime())) {
        return d;
      }
    }

    /*
     * DD/MM/YYYY
     */

    let match =
      text.match(
        /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
      );

    if (match) {
      const day = Number(match[1]);
      const month = Number(match[2]) - 1;
      const year = Number(match[3]);

      const d =
        new Date(
          Date.UTC(
            year,
            month,
            day
          )
        );

      if (!Number.isNaN(d.getTime())) {
        return d;
      }
    }

    /*
     * MM/DD/YYYY
     *
     * Se intenta únicamente como fallback.
     */

    match =
      text.match(
        /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
      );

    if (match) {
      const month = Number(match[1]) - 1;
      const day = Number(match[2]);
      const year = Number(match[3]);

      const d =
        new Date(
          Date.UTC(
            year,
            month,
            day
          )
        );

      if (!Number.isNaN(d.getTime())) {
        return d;
      }
    }

    return null;
  }

  /*
   * ------------------------------------------------------------
   * INFERENCIA DE TIPOS
   * ------------------------------------------------------------
   */

  function isIntegerLike(value) {
    if (
      value === null ||
      value === undefined ||
      String(value).trim() === ""
    ) {
      return false;
    }

    if (typeof value === "number") {
      return Number.isInteger(value);
    }

    return /^[-+]?\d+$/.test(
      String(value).trim()
    );
  }

  function isNumberLike(value) {
    if (
      value === null ||
      value === undefined ||
      String(value).trim() === ""
    ) {
      return false;
    }

    if (typeof value === "number") {
      return Number.isFinite(value);
    }

    const text =
      String(value)
        .trim()
        .replace(/,/g, "");

    return (
      text !== "" &&
      Number.isFinite(Number(text))
    );
  }

  function inferColumnType(
    columnIndex,
    columnName,
    dataRows
  ) {
    const values =
      dataRows
        .map(row => row[columnIndex])
        .filter(
          value =>
            value !== null &&
            value !== undefined &&
            String(value).trim() !== ""
        );

    const nonEmpty =
      values.length;

    if (nonEmpty === 0) {
      return {
        column_index: columnIndex,
        column_name: columnName,
        duckdb_type: "VARCHAR",
        sqlType: "VARCHAR",
        confidence: "low",
        reason: "columna sin valores",
        non_empty_values: 0
      };
    }

    /*
     * FECHA
     */

    if (
      isDateHeader(columnName)
    ) {
      let validDates = 0;

      for (const value of values) {
        if (
          normalizeDateValue(value)
        ) {
          validDates++;
        }
      }

      if (
        validDates === nonEmpty
      ) {
        return {
          column_index: columnIndex,
          column_name: columnName,
          duckdb_type: "DATE",
          sqlType: "DATE",
          confidence: "high",
          reason:
            "encabezado y valores compatibles con fecha",
          non_empty_values: nonEmpty
        };
      }
    }

    /*
     * INTEGER
     */

    if (
      values.every(
        isIntegerLike
      )
    ) {
      return {
        column_index: columnIndex,
        column_name: columnName,
        duckdb_type: "BIGINT",
        sqlType: "BIGINT",
        confidence: "high",
        reason:
          "todos los valores son enteros",
        non_empty_values: nonEmpty
      };
    }

    /*
     * DOUBLE
     */

    if (
      values.every(
        isNumberLike
      )
    ) {
      return {
        column_index: columnIndex,
        column_name: columnName,
        duckdb_type: "DOUBLE",
        sqlType: "DOUBLE",
        confidence: "medium",
        reason:
          "valores numéricos compatibles con decimal",
        non_empty_values: nonEmpty
      };
    }

    /*
     * TEXTO
     */

    return {
      column_index: columnIndex,
      column_name: columnName,
      duckdb_type: "VARCHAR",
      sqlType: "VARCHAR",
      confidence: "medium",
      reason:
        "valores tratados como texto",
      non_empty_values: nonEmpty
    };
  }

  function inferTypes() {
    inferredTypes =
      headers.map(
        (header, index) =>
          inferColumnType(
            index,
            header,
            rows
          )
      );

    return inferredTypes;
  }

  /*
   * ------------------------------------------------------------
   * PREPARACIÓN DE VALORES
   * ------------------------------------------------------------
   */

  function prepareValue(
    value,
    type
  ) {
    if (
      value === null ||
      value === undefined ||
      String(value).trim() === ""
    ) {
      return null;
    }

    if (type === "DATE") {
      const date =
        normalizeDateValue(value);

      if (!date) return null;

      return date;
    }

    if (type === "BIGINT") {
      const text =
        String(value)
          .trim()
          .replace(/,/g, "");

      if (!/^[-+]?\d+$/.test(text)) {
        return null;
      }

      return BigInt(text);
    }

    if (type === "DOUBLE") {
      const n =
        Number(
          String(value)
            .trim()
            .replace(/,/g, "")
        );

      return Number.isFinite(n)
        ? n
        : null;
    }

    return String(value);
  }

  /*
   * ------------------------------------------------------------
   * ESTILOS
   * ------------------------------------------------------------
   */

  function addStyles() {
    if (
      document.getElementById(
        "tmxe-v0493-style"
      )
    ) {
      return;
    }

    const style =
      document.createElement(
        "style"
      );

    style.id =
      "tmxe-v0493-style";

    style.textContent = `
      #tmxe-v0493-button {
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

      #tmxe-v0493-overlay {
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

      #tmxe-v0493-panel {
        width: min(1050px, 97vw);
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
          system-ui,
          sans-serif;
      }

      #tmxe-v0493-panel h2 {
        margin:
          0 0 4px;

        font-size:
          20px;
      }

      #tmxe-v0493-help {
        opacity: .7;
        margin-bottom: 14px;
      }

      #tmxe-v0493-status {
        padding: 12px;

        border-radius: 9px;

        background:
          rgba(127,127,127,.10);

        white-space:
          pre-wrap;

        font:
          13px/1.5
          ui-monospace,
          SFMono-Regular,
          Menlo,
          monospace;

        margin-bottom:
          12px;
      }

      #tmxe-v0493-status[data-kind="ok"] {
        color: #16a34a;
      }

      #tmxe-v0493-status[data-kind="error"] {
        color: #dc2626;
      }

      #tmxe-v0493-actions,
      #tmxe-v0493-sql-actions {
        display:
          flex;

        gap:
          8px;

        flex-wrap:
          wrap;

        margin:
          10px 0;
      }

      #tmxe-v0493-panel button {
        border:
          1px solid
          rgba(127,127,127,.45);

        border-radius:
          9px;

        padding:
          9px 13px;

        background:
          transparent;

        color:
          inherit;

        cursor:
          pointer;
      }

      #tmxe-v0493-load,
      #tmxe-v0493-run-sql {
        background:
          #7c3aed !important;

        color:
          white !important;

        border-color:
          #7c3aed !important;
      }

      #tmxe-v0493-file {
        display:
          none;
      }

      #tmxe-v0493-file-label {
        display:
          inline-block;

        border:
          1px solid
          rgba(127,127,127,.45);

        border-radius:
          9px;

        padding:
          9px 13px;

        cursor:
          pointer;
      }

      #tmxe-v0493-sql {
        width:
          100%;

        min-height:
          150px;

        resize:
          vertical;

        box-sizing:
          border-box;

        padding:
          12px;

        border-radius:
          10px;

        border:
          1px solid
          rgba(127,127,127,.4);

        background:
          rgba(127,127,127,.06);

        color:
          inherit;

        font:
          13px/1.5
          ui-monospace,
          SFMono-Regular,
          Menlo,
          monospace;
      }

      #tmxe-v0493-result {
        white-space:
          pre-wrap;

        overflow:
          auto;

        max-height:
          500px;

        padding:
          12px;

        border-radius:
          10px;

        background:
          rgba(127,127,127,.10);

        font:
          12px/1.45
          ui-monospace,
          SFMono-Regular,
          Menlo,
          monospace;
      }

      #tmxe-v0493-file-name {
        margin:
          8px 0;

        opacity:
          .75;
      }

      .tmxe-v0493-section {
        margin-top:
          18px;
      }

      .tmxe-v0493-small {
        opacity:
          .65;

        font-size:
          12px;
      }
    `;

    document.head.appendChild(
      style
    );
  }

  /*
   * ------------------------------------------------------------
   * INTERFAZ
   * ------------------------------------------------------------
   */

  function createButton() {
    if (
      document.getElementById(
        "tmxe-v0493-button"
      )
    ) {
      return;
    }

    const button =
      document.createElement(
        "button"
      );

    button.id =
      "tmxe-v0493-button";

    button.type =
      "button";

    button.textContent =
      "📊 Excel v0.4.9.3";

    button.title =
      "TypingMind Excel Data Engine";

    button.addEventListener(
      "click",
      openPanel
    );

    document.body.appendChild(
      button
    );
  }

  function openPanel() {
    if (
      document.getElementById(
        "tmxe-v0493-overlay"
      )
    ) {
      return;
    }

    const overlay =
      document.createElement(
        "div"
      );

    overlay.id =
      "tmxe-v0493-overlay";

    overlay.innerHTML = `
      <div
        id="tmxe-v0493-panel"
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
              📊 TypingMind Excel Data Engine
            </h2>

            <div id="tmxe-v0493-help">
              v0.4.9.3 TEST —
              Excel local + DuckDB-Wasm + SQL
            </div>
          </div>

          <button
            id="tmxe-v0493-close"
            type="button"
          >
            Cerrar
          </button>

        </div>

        <div id="tmxe-v0493-status">
          Motor listo.
        </div>

        <div id="tmxe-v0493-actions">

          <label
            id="tmxe-v0493-file-label"
            for="tmxe-v0493-file"
          >
            📁 Seleccionar Excel
          </label>

          <input
            id="tmxe-v0493-file"
            type="file"
            accept=".xlsx,.xls,.xlsm"
          >

          <button
            id="tmxe-v0493-load"
            type="button"
          >
            🚀 Cargar y analizar
          </button>

        </div>

        <div
          id="tmxe-v0493-file-name"
        >
          Ningún archivo seleccionado.
        </div>

        <div class="tmxe-v0493-section">

          <strong>
            Pruebas rápidas
          </strong>

          <div id="tmxe-v0493-actions">

            <button
              id="tmxe-v0493-count"
              type="button"
            >
              🔢 COUNT
            </button>

            <button
              id="tmxe-v0493-preview"
              type="button"
            >
              👁 Vista previa
            </button>

            <button
              id="tmxe-v0493-summary"
              type="button"
            >
              📈 Resumen
            </button>

            <button
              id="tmxe-v0493-year"
              type="button"
            >
              📅 Total por año
            </button>

            <button
              id="tmxe-v0493-version"
              type="button"
            >
              ℹ️ Versión DuckDB
            </button>

          </div>

        </div>

        <div class="tmxe-v0493-section">

          <strong>
            SQL manual
          </strong>

          <textarea
            id="tmxe-v0493-sql"
            spellcheck="false"
          >SELECT
    "Año",
    SUM("Total TEU's") AS total
FROM excel_data
GROUP BY "Año"
ORDER BY "Año";</textarea>

          <div id="tmxe-v0493-sql-actions">

            <button
              id="tmxe-v0493-run-sql"
              type="button"
            >
              ▶ Ejecutar SQL
            </button>

            <button
              id="tmxe-v0493-clear-sql"
              type="button"
            >
              Limpiar
            </button>

          </div>

        </div>

        <div class="tmxe-v0493-section">

          <strong>
            Resultado
          </strong>

          <pre id="tmxe-v0493-result">—</pre>

        </div>

        <div class="tmxe-v0493-small">
          v0.4.9.3 TEST —
          El archivo y las consultas se procesan
          localmente en el navegador.
        </div>

      </div>
    `;

    document.body.appendChild(
      overlay
    );

    $("tmxe-v0493-close")
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

    $("tmxe-v0493-file")
      .addEventListener(
        "change",
        handleFileSelection
      );

    $("tmxe-v0493-load")
      .addEventListener(
        "click",
        loadExcel
      );

    $("tmxe-v0493-count")
      .addEventListener(
        "click",
        () =>
          runPresetSQL(
            `
SELECT COUNT(*) AS registros
FROM excel_data;
            `
          )
      );

    $("tmxe-v0493-preview")
      .addEventListener(
        "click",
        () =>
          runPresetSQL(
            `
SELECT *
FROM excel_data
LIMIT 10;
            `
          )
      );

    $("tmxe-v0493-summary")
      .addEventListener(
        "click",
        runSummary
      );

    $("tmxe-v0493-year")
      .addEventListener(
        "click",
        runYearSummary
      );

    $("tmxe-v0493-version")
      .addEventListener(
        "click",
        runVersion
      );

    $("tmxe-v0493-run-sql")
      .addEventListener(
        "click",
        runManualSQL
      );

    $("tmxe-v0493-clear-sql")
      .addEventListener(
        "click",
        () => {
          $("tmxe-v0493-sql").value =
            "";
        }
      );
  }

  function setStatus(
    text,
    kind = ""
  ) {
    const el =
      $("tmxe-v0493-status");

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
      $("tmxe-v0493-result");

    if (!el) return;

    if (
      typeof value === "string"
    ) {
      el.textContent =
        value;
      return;
    }

    try {
      el.textContent =
        stringifySafe(value);
    } catch (error) {
      el.textContent =
        String(value);
    }
  }

  /*
   * ------------------------------------------------------------
   * ARCHIVO
   * ------------------------------------------------------------
   */

  function handleFileSelection(
    event
  ) {
    const file =
      event.target.files?.[0];

    if (!file) {
      currentFile =
        null;

      $("tmxe-v0493-file-name")
        .textContent =
          "Ningún archivo seleccionado.";

      return;
    }

    currentFile =
      file;

    $("tmxe-v0493-file-name")
      .textContent =
        `Archivo seleccionado: ${file.name}
Tamaño: ${file.size.toLocaleString()} bytes`;
  }

  /*
   * ------------------------------------------------------------
   * DUCKDB
   * ------------------------------------------------------------
   */

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
      "2/5 — Seleccionando bundle DuckDB..."
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

    initialized =
      true;

    return bundle;
  }

  /*
   * ------------------------------------------------------------
   * LEER EXCEL
   * ------------------------------------------------------------
   */

  async function readExcelFile(
    file
  ) {
    setStatus(
      "4/5 — Leyendo archivo Excel localmente..."
    );

    const XLSX =
      await import(
        XLSX_PACKAGE
      );

    const buffer =
      await file.arrayBuffer();

    workbook =
      XLSX.read(
        buffer,
        {
          type:
            "array",
          cellDates:
            true,
          cellNF:
            false,
          cellText:
            false
        }
      );

    const sheetNames =
      workbook.SheetNames || [];

    if (
      sheetNames.length === 0
    ) {
      throw new Error(
        "El archivo no contiene hojas."
      );
    }

    const sheetName =
      sheetNames[0];

    const sheet =
      workbook.Sheets[
        sheetName
      ];

    const matrix =
      XLSX.utils.sheet_to_json(
        sheet,
        {
          header:
            1,
          defval:
            null,
          raw:
            true,
          blankrows:
            true
        }
      );

    if (
      !matrix ||
      matrix.length === 0
    ) {
      throw new Error(
        "La hoja principal está vacía."
      );
    }

    /*
     * Buscar primera fila con contenido.
     */

    let headerIndex =
      -1;

    for (
      let i = 0;
      i < matrix.length;
      i++
    ) {
      const row =
        matrix[i];

      const hasContent =
        Array.isArray(row) &&
        row.some(
          value =>
            value !== null &&
            value !== undefined &&
            String(value).trim() !== ""
        );

      if (hasContent) {
        headerIndex =
          i;

        break;
      }
    }

    if (
      headerIndex === -1
    ) {
      throw new Error(
        "No se encontró una fila de encabezados."
      );
    }

    headers =
      normalizeHeaders(
        matrix[headerIndex]
      );

    const columnCount =
      headers.length;

    const physicalRows =
      matrix.length -
      headerIndex -
      1;

    const dataRows =
      [];

    let emptyRows =
      0;

    for (
      let i =
        headerIndex + 1;
      i < matrix.length;
      i++
    ) {
      const original =
        matrix[i] || [];

      const row =
        [];

      for (
        let c = 0;
        c < columnCount;
        c++
      ) {
        row.push(
          original[c] ??
          null
        );
      }

      const hasContent =
        row.some(
          value =>
            value !== null &&
            value !== undefined &&
            String(value).trim() !== ""
        );

      if (!hasContent) {
        emptyRows++;
        continue;
      }

      dataRows.push(
        row
      );
    }

    rows =
      dataRows;

    inferTypes();

    return {
      sheetNames,
      sheetName,
      physicalRows,
      realRows:
        rows.length,
      emptyRows,
      columnCount
    };
  }

  /*
   * ------------------------------------------------------------
   * CREAR TABLA DUCKDB
   * ------------------------------------------------------------
   */

  async function createDuckDBTable() {
    if (!conn) {
      throw new Error(
        "DuckDB no está conectado."
      );
    }

    /*
     * Eliminar tabla anterior.
     */

    await conn.query(
      `DROP TABLE IF EXISTS ${quoteIdentifier(tableName)};`
    );

    const definitions =
      inferredTypes.map(
        type =>
          `${quoteIdentifier(type.column_name)} ${type.sqlType}`
      );

    const createSQL =
      `
CREATE TABLE ${quoteIdentifier(tableName)} (
  ${definitions.join(",\n  ")}
);
      `;

    await conn.query(
      createSQL
    );
  }

  /*
   * ------------------------------------------------------------
   * INSERTAR DATOS POR LOTES
   * ------------------------------------------------------------
   */

  async function insertRows() {
    const batchSize =
      1000;

    const total =
      rows.length;

    for (
      let start = 0;
      start < total;
      start += batchSize
    ) {
      const end =
        Math.min(
          start + batchSize,
          total
        );

      const batch =
        rows.slice(
          start,
          end
        );

      const values =
        batch.map(
          row => {
            const prepared =
              row.map(
                (value, index) =>
                  prepareValue(
                    value,
                    inferredTypes[index]
                      .sqlType
                  )
              );

            return (
              "(" +
              prepared
                .map(sqlLiteral)
                .join(", ") +
              ")"
            );
          }
        );

      const sql =
        `
INSERT INTO ${quoteIdentifier(tableName)}
VALUES
${values.join(",\n")};
        `;

      await conn.query(
        sql
      );

      setStatus(
        `5/5 — Insertando datos localmente...\n` +
        `${end.toLocaleString()} / ${total.toLocaleString()} registros`
      );

      /*
       * Permitir que el navegador
       * respire entre lotes.
       */

      await new Promise(
        resolve =>
          setTimeout(
            resolve,
            0
          )
      );
    }
  }

  /*
   * ------------------------------------------------------------
   * CARGAR EXCEL
   * ------------------------------------------------------------
   */

  async function loadExcel() {
    try {
      setResult("");

      if (!currentFile) {
        throw new Error(
          "Primero selecciona un archivo Excel."
        );
      }

      if (
        !/\.(xlsx|xls|xlsm)$/i.test(
          currentFile.name
        )
      ) {
        throw new Error(
          "El archivo debe ser XLSX, XLS o XLSM."
        );
      }

      const bundle =
        await initializeDuckDB();

      const excelInfo =
        await readExcelFile(
          currentFile
        );

      await createDuckDBTable();

      await insertRows();

      /*
       * COUNT REAL
       */

      const countResult =
        await conn.query(
          `
SELECT COUNT(*) AS registros
FROM ${quoteIdentifier(tableName)};
          `
        );

      const countRows =
        rowsToObjects(
          countResult
        );

      /*
       * SCHEMA
       */

      const schemaResult =
        await conn.query(
          `
DESCRIBE ${quoteIdentifier(tableName)};
          `
        );

      const schema =
        rowsToObjects(
          schemaResult
        );

      /*
       * PREVIEW
       */

      const previewResult =
        await conn.query(
          `
SELECT *
FROM ${quoteIdentifier(tableName)}
LIMIT 10;
          `
        );

      const preview =
        rowsToObjects(
          previewResult
        );

      /*
       * VERSIÓN CORRECTA
       *
       * IMPORTANTE:
       *
       * pragma_version() NO tiene una columna
       * llamada "version".
       *
       * Utilizamos library_version.
       */

      let duckdbVersion =
        null;

      try {
        const versionResult =
          await conn.query(
            `
SELECT
  library_version AS duckdb_version
FROM pragma_version();
            `
          );

        duckdbVersion =
          rowsToObjects(
            versionResult
          );
      } catch (error) {
        duckdbVersion = {
          error:
            error?.message ||
            String(error)
        };
      }

      setStatus(
        "✅ Excel cargado y procesado localmente.",
        "ok"
      );

      setResult({
        engine:
          APP_ID,

        procesamiento:
          "LOCAL",

        archivo:
          currentFile.name,

        tamano_bytes:
          currentFile.size,

        hojas:
          excelInfo.sheetNames,

        hoja_principal:
          excelInfo.sheetName,

        filas_fisicas_detectadas:
          excelInfo.physicalRows,

        filas_reales:
          excelInfo.realRows,

        filas_vacias_ignoradas:
          excelInfo.emptyRows,

        filas_insertadas:
          rows.length,

        columnas_detectadas:
          excelInfo.columnCount,

        encabezados:
          headers,

        tipos_inferidos:
          inferredTypes,

        duckdb_table:
          tableName,

        duckdb_version:
          duckdbVersion,

        count:
          countRows,

        schema:
          schema,

        preview:
          preview,

        bundle: {
          mainModule:
            bundle.mainModule,

          mainWorker:
            bundle.mainWorker,

          pthreadWorker:
            bundle.pthreadWorker
        }
      });

    } catch (error) {
      console.error(
        `[${APP_ID}]`,
        error
      );

      setStatus(
        "❌ ERROR AL CARGAR EL ARCHIVO",
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

  /*
   * ------------------------------------------------------------
   * SQL
   * ------------------------------------------------------------
   */

  async function executeSQL(
    sql
  ) {
    if (!conn) {
      throw new Error(
        "Primero debes cargar un archivo Excel."
      );
    }

    const cleanSQL =
      String(sql || "")
        .trim();

    if (!cleanSQL) {
      throw new Error(
        "La consulta SQL está vacía."
      );
    }

    const start =
      performance.now();

    const result =
      await conn.query(
        cleanSQL
      );

    const elapsed =
      performance.now() -
      start;

    const data =
      rowsToObjects(
        result
      );

    return {
      procesamiento:
        "LOCAL",

      sql:
        cleanSQL,

      filas_resultado:
        data.length,

      tiempo_ms:
        Number(
          elapsed.toFixed(1)
        ),

      resultado:
        data
    };
  }

  async function runPresetSQL(
    sql
  ) {
    try {
      setStatus(
        "⏳ Ejecutando SQL localmente..."
      );

      const result =
        await executeSQL(
          sql
        );

      setStatus(
        "✅ Consulta ejecutada localmente.",
        "ok"
      );

      setResult(
        result
      );

    } catch (error) {
      console.error(
        `[${APP_ID}] SQL`,
        error
      );

      setStatus(
        "❌ ERROR EN LA CONSULTA SQL",
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

  async function runManualSQL() {
    const sql =
      $("tmxe-v0493-sql")
        ?.value || "";

    await runPresetSQL(
      sql
    );
  }

  /*
   * ------------------------------------------------------------
   * COUNT
   * ------------------------------------------------------------
   */

  /*
   * ------------------------------------------------------------
   * RESUMEN
   * ------------------------------------------------------------
   */

  async function runSummary() {
    try {
      if (!conn) {
        throw new Error(
          "Primero debes cargar un Excel."
        );
      }

      setStatus(
        "⏳ Generando resumen estadístico local..."
      );

      const numericColumns =
        inferredTypes.filter(
          item =>
            item.sqlType ===
              "BIGINT" ||
            item.sqlType ===
              "DOUBLE"
        );

      if (
        numericColumns.length === 0
      ) {
        throw new Error(
          "No se detectaron columnas numéricas."
        );
      }

      const expressions =
        [];

      for (
        const column of numericColumns
      ) {
        const q =
          quoteIdentifier(
            column.column_name
          );

        expressions.push(
          `COUNT(${q}) AS ${quoteIdentifier(
            column.column_name +
              "__count"
          )}`
        );

        expressions.push(
          `SUM(${q}) AS ${quoteIdentifier(
            column.column_name +
              "__sum"
          )}`
        );

        expressions.push(
          `AVG(${q}) AS ${quoteIdentifier(
            column.column_name +
              "__avg"
          )}`
        );

        expressions.push(
          `MIN(${q}) AS ${quoteIdentifier(
            column.column_name +
              "__min"
          )}`
        );

        expressions.push(
          `MAX(${q}) AS ${quoteIdentifier(
            column.column_name +
              "__max"
          )}`
        );
      }

      const sql =
        `
SELECT
  ${expressions.join(",\n  ")}
FROM ${quoteIdentifier(tableName)};
        `;

      const result =
        await executeSQL(
          sql
        );

      setStatus(
        "✅ Resumen generado localmente.",
        "ok"
      );

      setResult(
        result
      );

    } catch (error) {
      setStatus(
        "❌ ERROR EN EL RESUMEN",
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

  /*
   * ------------------------------------------------------------
   * TOTAL POR AÑO
   * ------------------------------------------------------------
   */

  async function runYearSummary() {
    try {
      if (!conn) {
        throw new Error(
          "Primero debes cargar un Excel."
        );
      }

      const yearColumn =
        inferredTypes.find(
          item =>
            String(
              item.column_name
            )
              .trim()
              .toLowerCase() ===
            "año"
        );

      const dateColumn =
        inferredTypes.find(
          item =>
            item.sqlType ===
            "DATE"
        );

      const totalColumn =
        inferredTypes.find(
          item =>
            String(
              item.column_name
            )
              .toLowerCase()
              .includes("total")
        );

      /*
       * Preferimos una columna "Año"
       * si existe.
       */

      let yearExpression;

      if (yearColumn) {
        yearExpression =
          quoteIdentifier(
            yearColumn.column_name
          );
      } else if (dateColumn) {
        yearExpression =
          `EXTRACT(YEAR FROM ${quoteIdentifier(
            dateColumn.column_name
          )})`;
      } else {
        throw new Error(
          "No se encontró una columna Año ni una columna Fecha."
        );
      }

      if (!totalColumn) {
        throw new Error(
          "No se encontró una columna cuyo nombre contenga 'Total'."
        );
      }

      const total =
        quoteIdentifier(
          totalColumn.column_name
        );

      const sql =
        `
SELECT
  ${yearExpression} AS año,
  SUM(${total}) AS total
FROM ${quoteIdentifier(tableName)}
GROUP BY ${yearExpression}
ORDER BY ${yearExpression};
        `;

      setStatus(
        "⏳ Calculando total por año localmente..."
      );

      const result =
        await executeSQL(
          sql
        );

      setStatus(
        "✅ Total por año calculado localmente.",
        "ok"
      );

      setResult(
        result
      );

    } catch (error) {
      setStatus(
        "❌ ERROR EN TOTAL POR AÑO",
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

  /*
   * ------------------------------------------------------------
   * VERSIÓN DUCKDB
   * ------------------------------------------------------------
   */

  async function runVersion() {
    try {
      if (!conn) {
        throw new Error(
          "DuckDB todavía no está conectado."
        );
      }

      /*
       * CORRECTO PARA DUCKDB 1.1.1:
       *
       * SELECT library_version
       * FROM pragma_version();
       */

      const sql =
        `
SELECT
  library_version AS duckdb_version
FROM pragma_version();
        `;

      const result =
        await executeSQL(
          sql
        );

      setStatus(
        "✅ Versión DuckDB consultada.",
        "ok"
      );

      setResult(
        result
      );

    } catch (error) {
      setStatus(
        "❌ ERROR CONSULTANDO VERSIÓN",
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

  /*
   * ------------------------------------------------------------
   * INIT
   * ------------------------------------------------------------
   */

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
