/*
 * TypingMind Excel Data Engine - Extension v0.4.9.3 TEST
 *
 * OBJETIVO:
 * - Excel local en navegador
 * - SheetJS para lectura XLSX
 * - DuckDB-Wasm 1.29.0
 * - Tabla local: excel_data
 * - SQL manual
 * - Normalización robusta de resultados
 *
 * CORRECCIÓN PRINCIPAL v0.4.9.3:
 * - Convierte correctamente BigInt
 * - Convierte TypedArrays
 * - Convierte objetos Arrow/vectoriales
 * - Evita resultados como:
 *      "total": [1535012, 0, 0, 0]
 * - Resultado esperado:
 *      "total": 1535012
 *
 * TODO EL PROCESAMIENTO ES LOCAL.
 */

(() => {
  "use strict";

  const APP_ID = "tm-excel-engine-v0493-test";

  const VERSION = "v0.4.9.3 TEST";

  const DUCKDB_PACKAGE =
    "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

  const XLSX_PACKAGE =
    "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";

  let duckdb = null;
  let db = null;
  let conn = null;

  let currentFile = null;
  let currentRows = [];
  let currentHeaders = [];
  let currentSheetName = null;
  let currentSchema = [];
  let currentTypeInfo = [];

  /*
   * ============================================================
   * UTILIDADES
   * ============================================================
   */

  function isTypedArray(value) {
    return (
      value != null &&
      typeof value === "object" &&
      ArrayBuffer.isView(value) &&
      !(value instanceof DataView)
    );
  }

  function isNumericKey(key) {
    return /^\d+$/.test(String(key));
  }

  function normalizeValue(value, depth = 0) {

    if (depth > 12) {
      return String(value);
    }

    if (value === null || value === undefined) {
      return value;
    }

    /*
     * BigInt
     */
    if (typeof value === "bigint") {
      const n = Number(value);

      if (Number.isSafeInteger(n)) {
        return n;
      }

      return value.toString();
    }

    /*
     * Fechas
     */
    if (value instanceof Date) {
      return value.toISOString();
    }

    /*
     * TypedArray
     *
     * Ejemplo problemático:
     * Uint32Array [1535012,0,0,0]
     *
     * Si representa un valor escalar de 4 bytes,
     * intentamos recuperar el valor correcto.
     */
    if (isTypedArray(value)) {

      /*
       * Si tiene un solo elemento, devolverlo directamente.
       */
      if (value.length === 1) {
        return normalizeValue(value[0], depth + 1);
      }

      /*
       * Algunos resultados Arrow/Decimal pueden llegar
       * como bloques de 4 posiciones.
       *
       * Si solamente el primer elemento tiene información
       * y los demás son cero, tratamos el primero como escalar.
       */
      if (
        value.length === 4 &&
        Number(value[1]) === 0 &&
        Number(value[2]) === 0 &&
        Number(value[3]) === 0
      ) {
        return normalizeValue(value[0], depth + 1);
      }

      return Array.from(value).map(v =>
        normalizeValue(v, depth + 1)
      );
    }

    /*
     * Array normal
     */
    if (Array.isArray(value)) {

      /*
       * Si llega un array de cuatro posiciones
       * con solamente el primer valor utilizado,
       * normalizarlo como escalar.
       */
      if (
        value.length === 4 &&
        Number(value[1]) === 0 &&
        Number(value[2]) === 0 &&
        Number(value[3]) === 0
      ) {
        return normalizeValue(value[0], depth + 1);
      }

      return value.map(v =>
        normalizeValue(v, depth + 1)
      );
    }

    /*
     * Objetos especiales / Arrow.
     */
    if (typeof value === "object") {

      const keys = Object.keys(value);

      /*
       * Detectar objeto tipo:
       *
       * {
       *   "0": 1535012,
       *   "1": 0,
       *   "2": 0,
       *   "3": 0
       * }
       */
      if (
        keys.length === 4 &&
        keys.every(isNumericKey)
      ) {

        const ordered = keys
          .sort((a, b) => Number(a) - Number(b))
          .map(k => value[k]);

        if (
          Number(ordered[1]) === 0 &&
          Number(ordered[2]) === 0 &&
          Number(ordered[3]) === 0
        ) {
          return normalizeValue(
            ordered[0],
            depth + 1
          );
        }

        return ordered.map(v =>
          normalizeValue(v, depth + 1)
        );
      }

      /*
       * Detectar objetos con toJSON.
       */
      if (
        typeof value.toJSON === "function"
      ) {
        try {
          return normalizeValue(
            value.toJSON(),
            depth + 1
          );
        } catch (_) {}
      }

      /*
       * Objeto normal.
       */
      const result = {};

      for (const key of keys) {
        result[key] =
          normalizeValue(
            value[key],
            depth + 1
          );
      }

      return result;
    }

    /*
     * Números, strings y booleanos.
     */
    return value;
  }

  function normalizeRows(rows) {

    if (!Array.isArray(rows)) {
      rows = Array.from(rows || []);
    }

    return rows.map(row =>
      normalizeValue(row)
    );
  }

  function safeJSONStringify(value) {

    return JSON.stringify(
      normalizeValue(value),
      (key, val) => {

        if (typeof val === "bigint") {
          const n = Number(val);

          return Number.isSafeInteger(n)
            ? n
            : val.toString();
        }

        return val;
      },
      2
    );
  }

  /*
   * ============================================================
   * ESTILOS
   * ============================================================
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
      document.createElement("style");

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
          rgba(0,0,0,.55);

        display: flex;

        align-items: center;
        justify-content: center;

        padding: 16px;
      }

      #tmxe-v0493-panel {
        width: min(1100px, 96vw);
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
        margin: 0;
        font-size: 20px;
      }

      #tmxe-v0493-help {
        opacity: .7;
        margin-top: 4px;
        margin-bottom: 15px;
      }

      #tmxe-v0493-status {
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

      #tmxe-v0493-status[data-kind="ok"] {
        color: #16a34a;
      }

      #tmxe-v0493-status[data-kind="error"] {
        color: #dc2626;
      }

      #tmxe-v0493-file {
        display: none;
      }

      #tmxe-v0493-load {
        border: 0;
        border-radius: 9px;

        padding: 10px 15px;

        background: #7c3aed;
        color: white;

        cursor: pointer;

        font-weight: 600;
      }

      #tmxe-v0493-sql {
        width: 100%;
        min-height: 180px;

        box-sizing: border-box;

        resize: vertical;

        padding: 12px;

        border-radius: 10px;

        border:
          1px solid
          rgba(127,127,127,.45);

        background:
          rgba(127,127,127,.08);

        color: inherit;

        font:
          13px/1.5
          ui-monospace,
          SFMono-Regular,
          Menlo,
          monospace;
      }

      #tmxe-v0493-actions {
        display: flex;

        flex-wrap: wrap;

        gap: 8px;

        margin: 10px 0;
      }

      #tmxe-v0493-actions button {
        border:
          1px solid
          rgba(127,127,127,.45);

        border-radius: 9px;

        padding: 9px 13px;

        background: transparent;

        color: inherit;

        cursor: pointer;
      }

      #tmxe-v0493-actions
      #tmxe-v0493-execute {
        background: #7c3aed;
        color: white;
        border-color: #7c3aed;
      }

      #tmxe-v0493-result {
        white-space: pre-wrap;

        overflow: auto;

        max-height: 480px;

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

      #tmxe-v0493-info {
        margin: 10px 0;

        padding: 10px;

        border-radius: 9px;

        background:
          rgba(127,127,127,.08);

        font-size: 13px;
      }

      .tmxe-v0493-label {
        font-weight: 700;
        margin-top: 15px;
        margin-bottom: 6px;
      }
    `;

    document.head.appendChild(style);
  }

  /*
   * ============================================================
   * BOTÓN
   * ============================================================
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
      document.createElement("button");

    button.id =
      "tmxe-v0493-button";

    button.type = "button";

    button.textContent =
      "📊 Excel v0.4.9.3";

    button.title =
      "Excel Data Engine v0.4.9.3";

    button.addEventListener(
      "click",
      openPanel
    );

    document.body.appendChild(button);
  }

  /*
   * ============================================================
   * PANEL
   * ============================================================
   */

  function openPanel() {

    if (
      document.getElementById(
        "tmxe-v0493-overlay"
      )
    ) {
      return;
    }

    const overlay =
      document.createElement("div");

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
              📊 Excel Data Engine
              ${VERSION}
            </h2>

            <div id="tmxe-v0493-help">
              Excel local + DuckDB-Wasm + SQL manual.
            </div>

          </div>

          <button
            id="tmxe-v0493-close"
            type="button"
            style="
              border:1px solid rgba(127,127,127,.45);
              border-radius:9px;
              padding:9px 13px;
              background:transparent;
              color:inherit;
              cursor:pointer;
            "
          >
            Cerrar
          </button>

        </div>

        <div
          id="tmxe-v0493-status"
        >
          Motor listo.
          Carga un archivo Excel para comenzar.
        </div>

        <div style="
          display:flex;
          align-items:center;
          gap:10px;
          flex-wrap:wrap;
        ">

          <input
            id="tmxe-v0493-file"
            type="file"
            accept=".xlsx,.xls,.xlsm"
          >

          <button
            id="tmxe-v0493-load"
            type="button"
          >
            📁 Cargar Excel
          </button>

          <span
            id="tmxe-v0493-filename"
            style="opacity:.75"
          >
            Ningún archivo cargado.
          </span>

        </div>

        <div
          id="tmxe-v0493-info"
        >
          Tabla DuckDB:
          <code>excel_data</code>
        </div>

        <div class="tmxe-v0493-label">
          SQL
        </div>

        <textarea
          id="tmxe-v0493-sql"
          spellcheck="false"
        >SELECT COUNT(*) AS registros
FROM excel_data;</textarea>

        <div id="tmxe-v0493-actions">

          <button
            id="tmxe-v0493-count"
            type="button"
          >
            COUNT
          </button>

          <button
            id="tmxe-v0493-preview"
            type="button"
          >
            Vista previa
          </button>

          <button
            id="tmxe-v0493-summary"
            type="button"
          >
            Resumen
          </button>

          <button
            id="tmxe-v0493-year"
            type="button"
          >
            Total por año
          </button>

          <button
            id="tmxe-v0493-execute"
            type="button"
          >
            ▶ Ejecutar SQL
          </button>

        </div>

        <div class="tmxe-v0493-label">
          Resultado
        </div>

        <pre
          id="tmxe-v0493-result"
        >—</pre>

        <div style="
          margin-top:12px;
          opacity:.6;
          font-size:12px;
        ">
          ${VERSION} —
          Todo el procesamiento se realiza localmente.
        </div>

      </div>
    `;

    document.body.appendChild(overlay);

    document
      .getElementById(
        "tmxe-v0493-close"
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

    const fileInput =
      document.getElementById(
        "tmxe-v0493-file"
      );

    document
      .getElementById(
        "tmxe-v0493-load"
      )
      .addEventListener(
        "click",
        () => fileInput.click()
      );

    fileInput.addEventListener(
      "change",
      loadExcel
    );

    document
      .getElementById(
        "tmxe-v0493-execute"
      )
      .addEventListener(
        "click",
        executeSQL
      );

    document
      .getElementById(
        "tmxe-v0493-count"
      )
      .addEventListener(
        "click",
        () => {

          setSQL(`
SELECT COUNT(*) AS registros
FROM excel_data;
          `.trim());

          executeSQL();
        }
      );

    document
      .getElementById(
        "tmxe-v0493-preview"
      )
      .addEventListener(
        "click",
        () => {

          setSQL(`
SELECT *
FROM excel_data
LIMIT 10;
          `.trim());

          executeSQL();
        }
      );

    document
      .getElementById(
        "tmxe-v0493-summary"
      )
      .addEventListener(
        "click",
        summarySQL
      );

    document
      .getElementById(
        "tmxe-v0493-year"
      )
      .addEventListener(
        "click",
        totalPorAno
      );
  }

  /*
   * ============================================================
   * UI
   * ============================================================
   */

  function setStatus(
    text,
    kind = ""
  ) {

    const el =
      document.getElementById(
        "tmxe-v0493-status"
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
        "tmxe-v0493-result"
      );

    if (!el) return;

    if (
      typeof value === "string"
    ) {
      el.textContent = value;
      return;
    }

    try {

      el.textContent =
        safeJSONStringify(value);

    } catch (error) {

      el.textContent =
        String(value);

    }
  }

  function setSQL(sql) {

    const el =
      document.getElementById(
        "tmxe-v0493-sql"
      );

    if (el) {
      el.value = sql;
    }
  }

  /*
   * ============================================================
   * CARGA DE LIBRERÍAS
   * ============================================================
   */

  async function ensureDuckDB() {

    if (
      duckdb &&
      db &&
      conn
    ) {
      return;
    }

    setStatus(
      "1/4 — Cargando DuckDB-Wasm..."
    );

    duckdb =
      await import(
        DUCKDB_PACKAGE
      );

    setStatus(
      "2/4 — Seleccionando bundle DuckDB..."
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

    setStatus(
      "3/4 — DuckDB conectado."
    );
  }

  async function ensureXLSX() {

    return await import(
      XLSX_PACKAGE
    );
  }

  /*
   * ============================================================
   * UTILIDADES EXCEL
   * ============================================================
   */

  function cleanHeader(
    value,
    index
  ) {

    let name =
      value == null
        ? ""
        : String(value).trim();

    if (!name) {
      name = `Columna_${index + 1}`;
    }

    return name;
  }

  function makeUniqueHeaders(
    headers
  ) {

    const used =
      new Map();

    return headers.map(
      (header, index) => {

        let name =
          cleanHeader(
            header,
            index
          );

        const count =
          used.get(name) || 0;

        if (count > 0) {
          name =
            `${name}_${count + 1}`;
        }

        used.set(
          cleanHeader(
            header,
            index
          ),
          count + 1
        );

        return name;
      }
    );
  }

  function isEmptyValue(value) {

    return (
      value === null ||
      value === undefined ||
      String(value).trim() === ""
    );
  }

  function isEmptyRow(row) {

    return row.every(
      value =>
        isEmptyValue(value)
    );
  }

  function excelSerialToDate(
    serial
  ) {

    if (
      typeof serial !== "number" ||
      !Number.isFinite(serial)
    ) {
      return null;
    }

    /*
     * Excel epoch.
     */
    const epoch =
      Date.UTC(
        1899,
        11,
        30
      );

    const ms =
      epoch +
      serial * 86400000;

    const date =
      new Date(ms);

    if (
      Number.isNaN(
        date.getTime()
      )
    ) {
      return null;
    }

    return date;
  }

  function formatDateForDuckDB(
    value
  ) {

    if (
      value instanceof Date
    ) {
      return value
        .toISOString()
        .slice(0, 10);
    }

    if (
      typeof value === "number"
    ) {

      const date =
        excelSerialToDate(
          value
        );

      if (date) {
        return date
          .toISOString()
          .slice(0, 10);
      }
    }

    const text =
      String(value)
        .trim();

    /*
     * ISO
     */
    if (
      /^\d{4}-\d{1,2}-\d{1,2}$/.test(
        text
      )
    ) {

      const parts =
        text.split("-");

      return (
        `${parts[0]}-` +
        `${parts[1].padStart(2, "0")}-` +
        `${parts[2].padStart(2, "0")}`
      );
    }

    return null;
  }

  function looksLikeDateHeader(
    header
  ) {

    const h =
      String(header)
        .toLowerCase();

    return (
      h.includes("fecha") ||
      h.includes("date") ||
      h.includes("día") ||
      h.includes("dia")
    );
  }

  function inferColumnType(
    header,
    values,
    index
  ) {

    const nonEmpty =
      values.filter(
        v => !isEmptyValue(v)
      );

    if (
      nonEmpty.length === 0
    ) {

      return {
        column_index: index,
        column_name: header,
        duckdb_type: "VARCHAR",
        sqlType: "VARCHAR",
        confidence: "low",
        reason: "sin valores",
        non_empty_values: 0
      };
    }

    /*
     * DATE
     */
    if (
      looksLikeDateHeader(header)
    ) {

      const dates =
        nonEmpty.filter(
          value =>
            formatDateForDuckDB(
              value
            ) !== null
        );

      if (
        dates.length ===
        nonEmpty.length
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
            nonEmpty.length
        };
      }
    }

    /*
     * INTEGER
     */
    const allIntegers =
      nonEmpty.every(
        value => {

          if (
            typeof value === "number"
          ) {
            return (
              Number.isFinite(value) &&
              Number.isInteger(value)
            );
          }

          const text =
            String(value)
              .trim();

          return (
            /^[-+]?\d+$/.test(text)
          );
        }
      );

    if (allIntegers) {

      return {
        column_index: index,
        column_name: header,
        duckdb_type: "BIGINT",
        sqlType: "BIGINT",
        confidence: "high",
        reason:
          "todos los valores son enteros",
        non_empty_values:
          nonEmpty.length
      };
    }

    /*
     * DOUBLE
     */
    const allNumbers =
      nonEmpty.every(
        value => {

          if (
            typeof value === "number"
          ) {
            return Number.isFinite(value);
          }

          const normalized =
            String(value)
              .trim()
              .replace(/,/g, "");

          return (
            normalized !== "" &&
            Number.isFinite(
              Number(normalized)
            )
          );
        }
      );

    if (allNumbers) {

      return {
        column_index: index,
        column_name: header,
        duckdb_type: "DOUBLE",
        sqlType: "DOUBLE",
        confidence: "high",
        reason:
          "todos los valores son numéricos",
        non_empty_values:
          nonEmpty.length
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
        nonEmpty.length
    };
  }

  /*
   * ============================================================
   * SQL HELPERS
   * ============================================================
   */

  function quoteIdentifier(
    name
  ) {

    return (
      '"' +
      String(name)
        .replace(/"/g, '""') +
      '"'
    );
  }

  function escapeSQLString(
    value
  ) {

    return String(value)
      .replace(/'/g, "''");
  }

  function valueForSQL(
    value,
    type
  ) {

    if (
      value === null ||
      value === undefined ||
      String(value).trim() === ""
    ) {
      return "NULL";
    }

    if (
      type === "DATE"
    ) {

      const date =
        formatDateForDuckDB(
          value
        );

      if (date) {
        return `DATE '${date}'`;
      }

      return "NULL";
    }

    if (
      type === "BIGINT"
    ) {

      const text =
        String(value)
          .trim()
          .replace(/,/g, "");

      if (
        /^[-+]?\d+$/.test(text)
      ) {
        return text;
      }

      return "NULL";
    }

    if (
      type === "DOUBLE"
    ) {

      const text =
        String(value)
          .trim()
          .replace(/,/g, "");

      if (
        Number.isFinite(
          Number(text)
        )
      ) {
        return text;
      }

      return "NULL";
    }

    return `'${escapeSQLString(
      value
    )}'`;
  }

  /*
   * ============================================================
   * CREAR TABLA
   * ============================================================
   */

  async function createDuckDBTable(
    rows,
    headers,
    typeInfo
  ) {

    await conn.query(
      `DROP TABLE IF EXISTS excel_data;`
    );

    const columnsSQL =
      headers.map(
        (header, index) =>
          `${quoteIdentifier(header)} ${
            typeInfo[index].sqlType
          }`
      ).join(",\n");

    await conn.query(
      `
      CREATE TABLE excel_data (
        ${columnsSQL}
      );
      `
    );

    /*
     * Insertar por lotes para no construir
     * una consulta gigantesca.
     */
    const BATCH_SIZE = 500;

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
        batch.map(
          row => {

            const values =
              headers.map(
                (_, index) =>
                  valueForSQL(
                    row[index],
                    typeInfo[index].sqlType
                  )
              );

            return `(${values.join(",")})`;
          }
        ).join(",\n");

      if (valuesSQL) {

        await conn.query(
          `
          INSERT INTO excel_data
          VALUES
          ${valuesSQL};
          `
        );
      }
    }
  }

  /*
   * ============================================================
   * CARGAR EXCEL
   * ============================================================
   */

  async function loadExcel(
    event
  ) {

    const file =
      event.target.files?.[0];

    if (!file) {
      return;
    }

    try {

      currentFile = file;

      setStatus(
        "Preparando carga local..."
      );

      setResult("");

      /*
       * DuckDB
       */
      await ensureDuckDB();

      /*
       * SheetJS
       */
      setStatus(
        "4/6 — Leyendo Excel localmente..."
      );

      const XLSX =
        await ensureXLSX();

      const arrayBuffer =
        await file.arrayBuffer();

      const workbook =
        XLSX.read(
          arrayBuffer,
          {
            type: "array",
            cellDates: false
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

      currentSheetName =
        sheets[0];

      const worksheet =
        workbook.Sheets[
          currentSheetName
        ];

      /*
       * Matriz completa.
       */
      const matrix =
        XLSX.utils.sheet_to_json(
          worksheet,
          {
            header: 1,
            defval: null,
            raw: true,
            blankrows: true
          }
        );

      const physicalRows =
        matrix.length;

      /*
       * Encontrar primera fila no vacía.
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
        headerIndex === -1
      ) {
        throw new Error(
          "No se encontró una fila de encabezados."
        );
      }

      currentHeaders =
        makeUniqueHeaders(
          matrix[headerIndex]
        );

      const dataMatrix =
        matrix.slice(
          headerIndex + 1
        );

      /*
       * Filas reales.
       */
      currentRows =
        dataMatrix.filter(
          row =>
            !isEmptyRow(row)
        );

      /*
       * Asegurar cantidad de columnas.
       */
      currentRows =
        currentRows.map(
          row => {

            const normalized =
              new Array(
                currentHeaders.length
              ).fill(null);

            for (
              let i = 0;
              i <
              currentHeaders.length;
              i++
            ) {
              normalized[i] =
                row[i] ?? null;
            }

            return normalized;
          }
        );

      /*
       * Inferencia.
       */
      currentTypeInfo =
        currentHeaders.map(
          (header, index) =>
            inferColumnType(
              header,
              currentRows.map(
                row => row[index]
              ),
              index
            )
        );

      setStatus(
        "5/6 — Creando tabla DuckDB local..."
      );

      await createDuckDBTable(
        currentRows,
        currentHeaders,
        currentTypeInfo
      );

      /*
       * Verificación.
       */
      const countResult =
        await conn.query(
          `
          SELECT COUNT(*) AS registros
          FROM excel_data;
          `
        );

      const countRows =
        normalizeRows(
          countResult.toArray()
        );

      const schemaResult =
        await conn.query(
          `
          DESCRIBE excel_data;
          `
        );

      currentSchema =
        normalizeRows(
          schemaResult.toArray()
        );

      const count =
        countRows[0]?.registros ?? 0;

      const previewResult =
        await conn.query(
          `
          SELECT *
          FROM excel_data
          LIMIT 10;
          `
        );

      const preview =
        normalizeRows(
          previewResult.toArray()
        );

      setStatus(
        "6/6 — Excel cargado y consultado localmente.",
        "ok"
      );

      const filenameEl =
        document.getElementById(
          "tmxe-v0493-filename"
        );

      if (filenameEl) {
        filenameEl.textContent =
          file.name;
      }

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
          currentSheetName,

        filas_fisicas_detectadas:
          physicalRows,

        filas_reales:
          currentRows.length,

        filas_vacias_ignoradas:
          physicalRows -
          headerIndex -
          1 -
          currentRows.length,

        filas_insertadas:
          Number(count),

        columnas_detectadas:
          currentHeaders.length,

        encabezados:
          currentHeaders,

        tipos_inferidos:
          currentTypeInfo,

        duckdb_table:
          "excel_data",

        duckdb_version:
          normalizeRows(
            (
              await conn.query(
                "SELECT version AS duckdb_version FROM pragma_version();"
              )
            ).toArray()
          ),

        schema:
          currentSchema,

        preview:
          preview

      });

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

  /*
   * ============================================================
   * EJECUTAR SQL
   * ============================================================
   */

  async function executeSQL() {

    if (!conn) {

      setStatus(
        "Primero carga un archivo Excel.",
        "error"
      );

      return;
    }

    const sqlEl =
      document.getElementById(
        "tmxe-v0493-sql"
      );

    const sql =
      sqlEl?.value?.trim();

    if (!sql) {

      setStatus(
        "Escribe una consulta SQL.",
        "error"
      );

      return;
    }

    try {

      setStatus(
        "Ejecutando SQL localmente..."
      );

      const start =
        performance.now();

      const result =
        await conn.query(
          sql
        );

      const end =
        performance.now();

      const rawRows =
        result.toArray();

      const rows =
        normalizeRows(
          rawRows
        );

      setStatus(
        "Consulta SQL ejecutada correctamente.",
        "ok"
      );

      setResult({

        procesamiento:
          "LOCAL",

        sql:
          sql,

        filas_resultado:
          rows.length,

        tiempo_ms:
          Number(
            (end - start)
              .toFixed(1)
          ),

        resultado:
          rows

      });

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

  /*
   * ============================================================
   * RESUMEN
   * ============================================================
   */

  async function summarySQL() {

    if (!currentHeaders.length) {

      setStatus(
        "Primero carga un Excel.",
        "error"
      );

      return;
    }

    const numericColumns =
      currentTypeInfo.filter(
        item =>
          item.sqlType === "BIGINT" ||
          item.sqlType === "DOUBLE"
      );

    if (
      numericColumns.length === 0
    ) {

      setStatus(
        "No se detectaron columnas numéricas.",
        "error"
      );

      return;
    }

    const expressions =
      numericColumns.map(
        item => {

          const c =
            quoteIdentifier(
              item.column_name
            );

          return `
SUM(${c}) AS ${quoteIdentifier(
            item.column_name +
            "__sum"
          )},
AVG(${c}) AS ${quoteIdentifier(
            item.column_name +
            "__avg"
          )},
MIN(${c}) AS ${quoteIdentifier(
            item.column_name +
            "__min"
          )},
MAX(${c}) AS ${quoteIdentifier(
            item.column_name +
            "__max"
          )}
          `.trim();

        }
      ).join(",\n");

    const sql = `
SELECT
    ${expressions}
FROM excel_data;
    `.trim();

    setSQL(sql);

    await executeSQL();
  }

  /*
   * ============================================================
   * TOTAL POR AÑO
   * ============================================================
   */

  async function totalPorAno() {

    if (!currentHeaders.length) {

      setStatus(
        "Primero carga un Excel.",
        "error"
      );

      return;
    }

    const yearColumn =
      currentHeaders.find(
        h =>
          String(h)
            .trim()
            .toLowerCase() ===
          "año"
      );

    const dateColumn =
      currentHeaders.find(
        h =>
          String(h)
            .trim()
            .toLowerCase() ===
          "fecha"
      );

    const totalColumn =
      currentHeaders.find(
        h =>
          String(h)
            .trim()
            .toLowerCase() ===
          "total teu's"
      );

    if (
      !totalColumn
    ) {

      setStatus(
        'No se encontró la columna "Total TEU\'s".',
        "error"
      );

      return;
    }

    let yearExpression;

    /*
     * Si existe Año y es numérico,
     * usamos esa columna directamente.
     */
    if (
      yearColumn
    ) {

      yearExpression =
        quoteIdentifier(
          yearColumn
        );

    } else if (
      dateColumn
    ) {

      yearExpression =
        `EXTRACT(YEAR FROM ${quoteIdentifier(
          dateColumn
        )})`;

    } else {

      setStatus(
        "No se encontró una columna de año o fecha.",
        "error"
      );

      return;
    }

    const sql = `
SELECT
    ${yearExpression} AS "Año",
    SUM(${quoteIdentifier(
      totalColumn
    )}) AS total
FROM excel_data
GROUP BY ${yearExpression}
ORDER BY ${yearExpression};
    `.trim();

    setSQL(sql);

    await executeSQL();
  }

  /*
   * ============================================================
   * INICIO
   * ============================================================
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
      { once: true }
    );

  } else {

    init();

  }

})();
