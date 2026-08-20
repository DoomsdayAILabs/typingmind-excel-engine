/*
 * TypingMind Excel Data Engine
 * Extension v0.4.10 TEST
 *
 * OBJETIVO
 * --------
 * Motor Excel local para TypingMind.
 *
 * - Excel se procesa completamente en el navegador.
 * - No se sube el archivo Excel a ningún servidor.
 * - DuckDB-Wasm ejecuta las consultas SQL localmente.
 * - XLSX se interpreta mediante SheetJS.
 * - Los resultados se normalizan antes de enviarlos/mostrarlos.
 *
 * CAMBIO PRINCIPAL v0.4.10
 * -------------------------
 * Corrección robusta de resultados DuckDB-Wasm:
 *
 * - BigInt
 * - BigInt64Array
 * - Uint32Array
 * - Int32Array
 * - Float64Array
 * - ArrayBuffer
 * - Date
 * - objetos Arrow
 * - valores numéricos
 * - NULL
 *
 * Ejemplo:
 *
 * ANTES:
 *
 * "total": [1535012, 0, 0, 0]
 *
 * AHORA:
 *
 * "total": 1535012
 *
 * DuckDB-Wasm:
 * @duckdb/duckdb-wasm@1.29.0
 *
 * XLSX:
 * xlsx@0.18.5
 */

(() => {
  "use strict";

  const APP_ID = "tm-excel-engine-v0410-test";

  const VERSION = "v0.4.10-test";

  const DUCKDB_PACKAGE =
    "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

  const XLSX_PACKAGE =
    "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";

  let duckdb = null;
  let XLSX = null;

  let db = null;
  let conn = null;

  let currentFile = null;
  let currentWorkbook = null;

  let currentSheetName = null;
  let currentHeaders = [];
  let currentRows = [];

  let worker = null;
  let workerURL = null;

  /*
   * ============================================================
   * UTILIDADES
   * ============================================================
   */

  function safeString(value) {
    if (value === null || value === undefined) {
      return "";
    }

    return String(value);
  }

  function isTypedArray(value) {
    return (
      value instanceof Int8Array ||
      value instanceof Uint8Array ||
      value instanceof Uint8ClampedArray ||
      value instanceof Int16Array ||
      value instanceof Uint16Array ||
      value instanceof Int32Array ||
      value instanceof Uint32Array ||
      value instanceof Float32Array ||
      value instanceof Float64Array ||
      value instanceof BigInt64Array ||
      value instanceof BigUint64Array
    );
  }

  /*
   * ------------------------------------------------------------
   * NORMALIZADOR ROBUSTO DE VALORES DUCKDB
   * ------------------------------------------------------------
   *
   * Esta es una de las partes principales de v0.4.10.
   */

  function normalizeValue(value) {

    if (value === null || value === undefined) {
      return null;
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
     * Date
     */
    if (value instanceof Date) {
      if (!Number.isNaN(value.getTime())) {
        return value.toISOString();
      }

      return null;
    }

    /*
     * ArrayBuffer
     */
    if (value instanceof ArrayBuffer) {
      return Array.from(
        new Uint8Array(value)
      );
    }

    /*
     * TypedArray
     */
    if (isTypedArray(value)) {

      const arr = Array.from(value);

      /*
       * Caso especial:
       *
       * DuckDB/Arrow puede devolver valores
       * numéricos como un TypedArray de longitud 4.
       *
       * Si parece ser un valor escalar codificado
       * de esa forma, intentamos extraerlo.
       */

      if (arr.length === 1) {
        return normalizeValue(arr[0]);
      }

      /*
       * Para Uint32Array/Int32Array/etc.
       * devolvemos un array normal.
       */

      return arr.map(
        item => normalizeValue(item)
      );
    }

    /*
     * Array normal
     */
    if (Array.isArray(value)) {

      return value.map(
        item => normalizeValue(item)
      );
    }

    /*
     * Objetos
     */
    if (typeof value === "object") {

      /*
       * Algunos objetos numéricos pueden
       * tener propiedades 0,1,2,3.
       *
       * Detectamos estructuras que realmente
       * representan un único valor numérico.
       */

      const keys = Object.keys(value);

      const numericKeys =
        keys.filter(
          key => /^\d+$/.test(key)
        );

      if (
        numericKeys.length > 0 &&
        numericKeys.length === keys.length
      ) {

        const ordered =
          numericKeys
            .sort(
              (a, b) =>
                Number(a) - Number(b)
            )
            .map(
              key => value[key]
            );

        /*
         * Si tenemos:
         *
         * {
         *   0: 1535012,
         *   1: 0,
         *   2: 0,
         *   3: 0
         * }
         *
         * y los elementos posteriores son 0,
         * asumimos que el primero representa
         * el valor escalar.
         */

        if (
          ordered.length > 1 &&
          ordered.slice(1).every(
            item =>
              item === 0 ||
              item === 0n
          )
        ) {
          return normalizeValue(
            ordered[0]
          );
        }

        /*
         * Si todos los elementos son valores
         * reales, devolvemos array.
         */

        return ordered.map(
          item => normalizeValue(item)
        );
      }

      /*
       * Objetos normales:
       * copiar propiedad por propiedad.
       */

      const result = {};

      for (const key of keys) {

        result[key] =
          normalizeValue(value[key]);

      }

      return result;
    }

    /*
     * number
     */
    if (typeof value === "number") {

      if (
        Number.isNaN(value) ||
        !Number.isFinite(value)
      ) {
        return null;
      }

      return value;
    }

    /*
     * boolean / string
     */
    if (
      typeof value === "boolean" ||
      typeof value === "string"
    ) {
      return value;
    }

    return String(value);
  }

  /*
   * ------------------------------------------------------------
   * NORMALIZAR FILA COMPLETA
   * ------------------------------------------------------------
   */

  function normalizeRow(row) {

    const result = {};

    if (!row) {
      return result;
    }

    for (const key of Object.keys(row)) {

      result[key] =
        normalizeValue(row[key]);

    }

    return result;
  }

  /*
   * ------------------------------------------------------------
   * NORMALIZAR RESULTADO DUCKDB
   * ------------------------------------------------------------
   */

  function normalizeDuckDBResult(result) {

    if (!result) {
      return [];
    }

    let rows = [];

    try {

      rows =
        result.toArray();

    } catch (error) {

      console.warn(
        `[${APP_ID}] No se pudo ejecutar toArray()`,
        error
      );

      return [];
    }

    return rows.map(
      row => normalizeRow(row)
    );
  }

  /*
   * ------------------------------------------------------------
   * JSON SEGURO
   * ------------------------------------------------------------
   */

  function safeJSONStringify(value) {

    try {

      return JSON.stringify(
        value,
        (key, val) =>
          normalizeValue(val),
        2
      );

    } catch (error) {

      return JSON.stringify(
        {
          error:
            "No fue posible serializar el resultado",
          mensaje:
            error?.message ||
            String(error)
        },
        null,
        2
      );

    }
  }

  /*
   * ============================================================
   * ESTILOS
   * ============================================================
   */

  function addStyles() {

    if (
      document.getElementById(
        "tmxe-v0410-style"
      )
    ) {
      return;
    }

    const style =
      document.createElement("style");

    style.id =
      "tmxe-v0410-style";

    style.textContent = `

      #tmxe-v0410-button {
        position: fixed;
        right: 18px;
        bottom: 86px;
        z-index: 2147483000;

        border: 0;
        border-radius: 999px;

        padding: 11px 16px;

        background: #7c3aed;
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

      #tmxe-v0410-overlay {
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

      #tmxe-v0410-panel {

        width:
          min(1000px, 96vw);

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

      #tmxe-v0410-panel h2 {
        margin:
          0 0 5px;

        font-size: 20px;
      }

      #tmxe-v0410-help {

        opacity: .7;

        margin-bottom: 15px;
      }

      #tmxe-v0410-status {

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

        margin-bottom: 12px;
      }

      #tmxe-v0410-status[data-kind="ok"] {
        color: #16a34a;
      }

      #tmxe-v0410-status[data-kind="error"] {
        color: #dc2626;
      }

      #tmxe-v0410-status[data-kind="busy"] {
        color: #7c3aed;
      }

      #tmxe-v0410-result {

        white-space: pre-wrap;

        overflow: auto;

        max-height: 420px;

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

      #tmxe-v0410-sql {

        width: 100%;

        min-height: 150px;

        box-sizing: border-box;

        resize: vertical;

        padding: 12px;

        border-radius: 10px;

        border:
          1px solid
          rgba(127,127,127,.35);

        background: Canvas;

        color: CanvasText;

        font:
          13px/1.5
          ui-monospace,
          SFMono-Regular,
          Menlo,
          monospace;
      }

      #tmxe-v0410-file {

        display: none;
      }

      .tmxe-v0410-actions {

        display: flex;

        gap: 8px;

        flex-wrap: wrap;

        margin: 10px 0;
      }

      .tmxe-v0410-btn {

        border:
          1px solid
          rgba(127,127,127,.45);

        border-radius: 9px;

        padding: 9px 13px;

        background: transparent;

        color: inherit;

        cursor: pointer;
      }

      .tmxe-v0410-primary {

        background: #7c3aed !important;

        color: white !important;

        border-color:
          #7c3aed !important;
      }

      .tmxe-v0410-success {

        background: #16a34a !important;

        color: white !important;

        border-color:
          #16a34a !important;
      }

      .tmxe-v0410-section {

        margin-top: 16px;

        padding-top: 12px;

        border-top:
          1px solid
          rgba(127,127,127,.25);
      }

      .tmxe-v0410-label {

        display: block;

        margin-bottom: 7px;

        font-weight: 600;
      }

    `;

    document.head.appendChild(style);
  }

  /*
   * ============================================================
   * INTERFAZ
   * ============================================================
   */

  function createButton() {

    if (
      document.getElementById(
        "tmxe-v0410-button"
      )
    ) {
      return;
    }

    const button =
      document.createElement("button");

    button.id =
      "tmxe-v0410-button";

    button.type = "button";

    button.textContent =
      "📊 Excel Engine " +
      VERSION.replace("-test", "");

    button.title =
      "TypingMind Excel Data Engine " +
      VERSION;

    button.addEventListener(
      "click",
      openPanel
    );

    document.body.appendChild(button);
  }

  function openPanel() {

    if (
      document.getElementById(
        "tmxe-v0410-overlay"
      )
    ) {
      return;
    }

    const overlay =
      document.createElement("div");

    overlay.id =
      "tmxe-v0410-overlay";

    overlay.innerHTML = `

      <div
        id="tmxe-v0410-panel"
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

            <div id="tmxe-v0410-help">

              ${VERSION} —
              Procesamiento 100 % local.
              Excel → DuckDB-Wasm → SQL.

            </div>

          </div>

          <button
            id="tmxe-v0410-close"
            class="tmxe-v0410-btn"
            type="button"
          >
            Cerrar
          </button>

        </div>

        <div
          id="tmxe-v0410-status"
        >
          Motor listo.
          Selecciona un archivo Excel.
        </div>

        <div class="tmxe-v0410-actions">

          <input
            id="tmxe-v0410-file"
            type="file"
            accept=".xlsx,.xls,.xlsm"
          >

          <button
            id="tmxe-v0410-load"
            class="tmxe-v0410-btn tmxe-v0410-primary"
            type="button"
          >
            📂 Cargar Excel
          </button>

          <button
            id="tmxe-v0410-count"
            class="tmxe-v0410-btn"
            type="button"
          >
            🔢 COUNT
          </button>

          <button
            id="tmxe-v0410-preview"
            class="tmxe-v0410-btn"
            type="button"
          >
            👁 Vista previa
          </button>

          <button
            id="tmxe-v0410-summary"
            class="tmxe-v0410-btn"
            type="button"
          >
            📈 Resumen
          </button>

          <button
            id="tmxe-v0410-year"
            class="tmxe-v0410-btn"
            type="button"
          >
            📅 Total por año
          </button>

        </div>

        <div class="tmxe-v0410-section">

          <label
            class="tmxe-v0410-label"
            for="tmxe-v0410-sql"
          >
            Ejecutar SQL
          </label>

          <textarea
            id="tmxe-v0410-sql"
          >SELECT COUNT(*) AS registros
FROM excel_data;</textarea>

          <div class="tmxe-v0410-actions">

            <button
              id="tmxe-v0410-run-sql"
              class="tmxe-v0410-btn tmxe-v0410-success"
              type="button"
            >
              ▶ Ejecutar SQL
            </button>

            <button
              id="tmxe-v0410-schema"
              class="tmxe-v0410-btn"
              type="button"
            >
              🧱 Ver esquema
            </button>

          </div>

        </div>

        <div class="tmxe-v0410-section">

          <strong>
            Resultado
          </strong>

          <pre
            id="tmxe-v0410-result"
          >—</pre>

        </div>

        <div style="
          margin-top:12px;
          opacity:.65;
          font-size:12px;
        ">

          ${VERSION}<br>
          DuckDB-Wasm 1.29.0<br>
          DuckDB interno 1.1.1<br>
          Procesamiento local.

        </div>

      </div>
    `;

    document.body.appendChild(
      overlay
    );

    document
      .getElementById(
        "tmxe-v0410-close"
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
        "tmxe-v0410-load"
      )
      .addEventListener(
        "click",
        () => {

          document
            .getElementById(
              "tmxe-v0410-file"
            )
            .click();

        }
      );

    document
      .getElementById(
        "tmxe-v0410-file"
      )
      .addEventListener(
        "change",
        event => {

          const file =
            event.target.files?.[0];

          if (file) {
            loadExcel(file);
          }

        }
      );

    document
      .getElementById(
        "tmxe-v0410-run-sql"
      )
      .addEventListener(
        "click",
        executeSQL
      );

    document
      .getElementById(
        "tmxe-v0410-count"
      )
      .addEventListener(
        "click",
        countRows
      );

    document
      .getElementById(
        "tmxe-v0410-preview"
      )
      .addEventListener(
        "click",
        previewRows
      );

    document
      .getElementById(
        "tmxe-v0410-summary"
      )
      .addEventListener(
        "click",
        summaryRows
      );

    document
      .getElementById(
        "tmxe-v0410-year"
      )
      .addEventListener(
        "click",
        totalByYear
      );

    document
      .getElementById(
        "tmxe-v0410-schema"
      )
      .addEventListener(
        "click",
        showSchema
      );
  }

  /*
   * ============================================================
   * STATUS / RESULTADO
   * ============================================================
   */

  function setStatus(
    text,
    kind = ""
  ) {

    const element =
      document.getElementById(
        "tmxe-v0410-status"
      );

    if (!element) {
      return;
    }

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
        "tmxe-v0410-result"
      );

    if (!element) {
      return;
    }

    element.textContent =
      typeof value === "string"
        ? value
        : safeJSONStringify(value);
  }

  /*
   * ============================================================
   * DUCKDB
   * ============================================================
   */

  async function initializeDuckDB() {

    if (
      db &&
      conn
    ) {
      return;
    }

    setStatus(
      "1/4 — Cargando DuckDB-Wasm...",
      "busy"
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

    setStatus(
      "2/4 — Inicializando WebAssembly...",
      "busy"
    );

    workerURL =
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

    worker =
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

    conn =
      await db.connect();

    setStatus(
      "3/4 — DuckDB conectado.",
      "ok"
    );

    return bundle;
  }

  /*
   * ============================================================
   * XLSX
   * ============================================================
   */

  async function loadXLSXLibrary() {

    if (XLSX) {
      return XLSX;
    }

    setStatus(
      "Cargando lector XLSX...",
      "busy"
    );

    XLSX =
      await import(
        XLSX_PACKAGE
      );

    /*
     * Algunas versiones ESM colocan XLSX
     * dentro de default.
     */

    if (
      XLSX.default &&
      XLSX.default.read
    ) {
      XLSX =
        XLSX.default;
    }

    return XLSX;
  }

  /*
   * ============================================================
   * DETECCIÓN DE TIPOS
   * ============================================================
   */

  function isEmptyValue(value) {

    return (
      value === null ||
      value === undefined ||
      String(value).trim() === ""
    );
  }

  function isIntegerValue(value) {

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

    if (!text) {
      return false;
    }

    return /^[-+]?\d+$/.test(text);
  }

  function isNumberValue(value) {

    if (
      typeof value === "number"
    ) {

      return Number.isFinite(value);

    }

    const text =
      String(value)
        .trim()
        .replace(/,/g, "");

    if (!text) {
      return false;
    }

    return (
      /^[-+]?(?:\d+\.?\d*|\.\d+)$/.test(
        text
      )
    );
  }

  function isDateValue(value) {

    if (
      value instanceof Date
    ) {
      return !Number.isNaN(
        value.getTime()
      );
    }

    if (
      typeof value === "number"
    ) {

      /*
       * Excel serial date aproximado.
       */

      return (
        value > 20000 &&
        value < 100000
      );
    }

    const text =
      String(value).trim();

    if (!text) {
      return false;
    }

    /*
     * Formatos ISO.
     */

    if (
      /^\d{4}-\d{1,2}-\d{1,2}$/.test(
        text
      )
    ) {
      return true;
    }

    return false;
  }

  function inferColumnType(
    header,
    values
  ) {

    const nonEmpty =
      values.filter(
        value =>
          !isEmptyValue(value)
      );

    if (
      nonEmpty.length === 0
    ) {
      return {
        duckdb_type: "VARCHAR",
        confidence: "low",
        reason:
          "sin valores no vacíos",
        non_empty_values: 0
      };
    }

    const headerLower =
      String(header)
        .toLowerCase();

    const dateHeader =
      /fecha|date|día|dia/.test(
        headerLower
      );

    const allDates =
      nonEmpty.every(
        value =>
          isDateValue(value)
      );

    if (
      dateHeader &&
      allDates
    ) {

      return {
        duckdb_type: "DATE",
        sqlType: "DATE",
        confidence: "high",
        reason:
          "encabezado y valores compatibles con fecha",
        non_empty_values:
          nonEmpty.length
      };
    }

    const allIntegers =
      nonEmpty.every(
        value =>
          isIntegerValue(value)
      );

    if (
      allIntegers
    ) {

      return {
        duckdb_type: "BIGINT",
        sqlType: "BIGINT",
        confidence: "high",
        reason:
          "todos los valores son enteros",
        non_empty_values:
          nonEmpty.length
      };
    }

    const allNumbers =
      nonEmpty.every(
        value =>
          isNumberValue(value)
      );

    if (
      allNumbers
    ) {

      return {
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
   * IDENTIFICADORES SQL
   * ============================================================
   */

  function quoteIdentifier(name) {

    return (
      '"' +
      String(name)
        .replace(/"/g, '""') +
      '"'
    );
  }

  /*
   * ============================================================
   * CONVERSIÓN DE FECHAS
   * ============================================================
   */

  function excelSerialToDate(
    serial
  ) {

    /*
     * Sistema de fechas Excel 1900.
     */

    const excelEpoch =
      Date.UTC(
        1899,
        11,
        30
      );

    const milliseconds =
      Number(serial) *
      24 *
      60 *
      60 *
      1000;

    return new Date(
      excelEpoch +
      milliseconds
    );
  }

  function dateToSQL(value) {

    let date = null;

    if (
      value instanceof Date
    ) {

      date = value;

    } else if (
      typeof value === "number"
    ) {

      date =
        excelSerialToDate(
          value
        );

    } else {

      const text =
        String(value)
          .trim();

      if (
        /^\d{4}-\d{1,2}-\d{1,2}$/.test(
          text
        )
      ) {

        date =
          new Date(
            text + "T00:00:00Z"
          );

      }

    }

    if (
      !date ||
      Number.isNaN(
        date.getTime()
      )
    ) {

      return null;

    }

    return (
      date
        .toISOString()
        .slice(0, 10)
    );
  }

  /*
   * ============================================================
   * VALOR SQL
   * ============================================================
   */

  function sqlValue(
    value,
    type
  ) {

    if (
      isEmptyValue(value)
    ) {
      return "NULL";
    }

    if (
      type === "DATE"
    ) {

      const date =
        dateToSQL(value);

      if (!date) {
        return "NULL";
      }

      return (
        "DATE '" +
        date +
        "'"
      );
    }

    if (
      type === "BIGINT"
    ) {

      const text =
        String(value)
          .trim()
          .replace(/,/g, "");

      if (
        /^[-+]?\d+$/.test(
          text
        )
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

      const number =
        Number(text);

      if (
        Number.isFinite(number)
      ) {

        return String(number);

      }

      return "NULL";
    }

    /*
     * VARCHAR
     */

    return (
      "'" +
      String(value)
        .replace(/'/g, "''") +
      "'"
    );
  }

  /*
   * ============================================================
   * CARGAR EXCEL
   * ============================================================
   */

  async function loadExcel(
    file
  ) {

    try {

      setResult("");

      currentFile = file;

      setStatus(
        "1/8 — Preparando archivo...",
        "busy"
      );

      const bundle =
        await initializeDuckDB();

      await loadXLSXLibrary();

      setStatus(
        "4/8 — Leyendo XLSX localmente...",
        "busy"
      );

      const buffer =
        await file.arrayBuffer();

      currentWorkbook =
        XLSX.read(
          buffer,
          {
            type: "array",
            cellDates: false
          }
        );

      const sheets =
        currentWorkbook.SheetNames;

      if (
        !sheets ||
        sheets.length === 0
      ) {

        throw new Error(
          "El archivo no contiene hojas."
        );

      }

      currentSheetName =
        sheets[0];

      const sheet =
        currentWorkbook.Sheets[
          currentSheetName
        ];

      /*
       * header: 1 devuelve matriz.
       */

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

      if (
        physicalRows === 0
      ) {

        throw new Error(
          "La hoja está vacía."
        );

      }

      /*
       * Buscar la primera fila que
       * tenga contenido suficiente para
       * considerarla encabezado.
       */

      let headerIndex = 0;

      for (
        let i = 0;
        i < matrix.length;
        i++
      ) {

        const row =
          matrix[i] || [];

        const nonEmpty =
          row.filter(
            value =>
              !isEmptyValue(value)
          );

        if (
          nonEmpty.length >= 1
        ) {

          headerIndex = i;
          break;

        }

      }

      const rawHeaders =
        matrix[
          headerIndex
        ] || [];

      /*
       * Determinar número de columnas.
       */

      let columnCount =
        rawHeaders.length;

      for (
        let i =
          headerIndex + 1;
        i < matrix.length;
        i++
      ) {

        columnCount =
          Math.max(
            columnCount,
            (
              matrix[i] || []
            ).length
          );

      }

      currentHeaders =
        [];

      for (
        let i = 0;
        i < columnCount;
        i++
      ) {

        let header =
          rawHeaders[i];

        if (
          isEmptyValue(header)
        ) {

          header =
            `Column_${i + 1}`;

        }

        currentHeaders.push(
          String(header)
            .trim()
        );

      }

      /*
       * Garantizar encabezados únicos.
       */

      const used =
        new Map();

      currentHeaders =
        currentHeaders.map(
          header => {

            const count =
              used.get(header) || 0;

            used.set(
              header,
              count + 1
            );

            if (
              count === 0
            ) {

              return header;

            }

            return (
              header +
              "_" +
              (count + 1)
            );

          }
        );

      /*
       * Filas reales.
       */

      currentRows = [];

      for (
        let i =
          headerIndex + 1;
        i < matrix.length;
        i++
      ) {

        const sourceRow =
          matrix[i] || [];

        const normalized =
          [];

        let hasData = false;

        for (
          let c = 0;
          c < columnCount;
          c++
        ) {

          const value =
            sourceRow[c];

          normalized.push(
            value ?? null
          );

          if (
            !isEmptyValue(value)
          ) {

            hasData = true;

          }

        }

        if (
          hasData
        ) {

          currentRows.push(
            normalized
          );

        }

      }

      setStatus(
        "5/8 — Filas reales detectadas: " +
        currentRows.length,
        "busy"
      );

      /*
       * Inferencia de tipos.
       */

      const inferredTypes =
        currentHeaders.map(
          (header, index) => {

            const values =
              currentRows.map(
                row => row[index]
              );

            return {
              column_index:
                index,
              column_name:
                header,
              ...inferColumnType(
                header,
                values
              )
            };

          }
        );

      /*
       * Crear tabla.
       */

      await conn.query(
        "DROP TABLE IF EXISTS excel_data;"
      );

      const definitions =
        inferredTypes
          .map(
            info =>
              `${quoteIdentifier(
                info.column_name
              )} ${info.sqlType || info.duckdb_type}`
          )
          .join(", ");

      await conn.query(
        `CREATE TABLE excel_data (${definitions});`
      );

      setStatus(
        "6/8 — Insertando datos en DuckDB...",
        "busy"
      );

      /*
       * Insertar por lotes.
       *
       * Esto evita construir una sola
       * consulta gigantesca.
       */

      const BATCH_SIZE = 500;

      for (
        let start = 0;
        start < currentRows.length;
        start += BATCH_SIZE
      ) {

        const batch =
          currentRows.slice(
            start,
            start + BATCH_SIZE
          );

        if (
          batch.length === 0
        ) {
          continue;
        }

        const valuesSQL =
          batch
            .map(
              row =>
                "(" +
                row
                  .map(
                    (value, index) =>
                      sqlValue(
                        value,
                        inferredTypes[
                          index
                        ].sqlType ||
                        inferredTypes[
                          index
                        ].duckdb_type
                      )
                  )
                  .join(",") +
                ")"
            )
            .join(",");

        await conn.query(
          `INSERT INTO excel_data VALUES ${valuesSQL};`
        );

      }

      setStatus(
        "7/8 — Validando tabla local...",
        "busy"
      );

      const countResult =
        await conn.query(
          `
          SELECT
            COUNT(*) AS registros
          FROM excel_data;
          `
        );

      const countRows =
        normalizeDuckDBResult(
          countResult
        );

      /*
       * Schema.
       */

      const schemaResult =
        await conn.query(
          `
          DESCRIBE excel_data;
          `
        );

      const schema =
        normalizeDuckDBResult(
          schemaResult
        );

      /*
       * Preview.
       */

      const previewResult =
        await conn.query(
          `
          SELECT *
          FROM excel_data
          LIMIT 10;
          `
        );

      const preview =
        normalizeDuckDBResult(
          previewResult
        );

      setStatus(
        "8/8 — Excel cargado y consultado localmente.",
        "ok"
      );

      setResult({

        engine:
          VERSION,

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
          currentRows.length,

        columnas_detectadas:
          currentHeaders.length,

        encabezados:
          currentHeaders,

        tipos_inferidos:
          inferredTypes,

        duckdb_table:
          "excel_data",

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

    try {

      if (
        !conn
      ) {

        await initializeDuckDB();

      }

      if (
        !currentRows.length
      ) {

        throw new Error(
          "Primero debes cargar un archivo Excel."
        );

      }

      const sql =
        document
          .getElementById(
            "tmxe-v0410-sql"
          )
          .value
          .trim();

      if (!sql) {

        throw new Error(
          "El SQL está vacío."
        );

      }

      setStatus(
        "Ejecutando SQL localmente...",
        "busy"
      );

      const start =
        performance.now();

      const result =
        await conn.query(
          sql
        );

      const end =
        performance.now();

      const rows =
        normalizeDuckDBResult(
          result
        );

      setStatus(
        "SQL ejecutado correctamente.",
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
            (
              end - start
            ).toFixed(1)
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

  /*
   * ============================================================
   * COUNT
   * ============================================================
   */

  async function countRows() {

    const textarea =
      document.getElementById(
        "tmxe-v0410-sql"
      );

    textarea.value =
      `
SELECT
    COUNT(*) AS registros
FROM excel_data;
      `.trim();

    await executeSQL();
  }

  /*
   * ============================================================
   * PREVIEW
   * ============================================================
   */

  async function previewRows() {

    const textarea =
      document.getElementById(
        "tmxe-v0410-sql"
      );

    textarea.value =
      `
SELECT *
FROM excel_data
LIMIT 10;
      `.trim();

    await executeSQL();
  }

  /*
   * ============================================================
   * RESUMEN
   * ============================================================
   */

  async function summaryRows() {

    try {

      if (
        !conn ||
        !currentRows.length
      ) {

        throw new Error(
          "Primero debes cargar un Excel."
        );

      }

      /*
       * Obtener columnas numéricas.
       */

      const describeResult =
        await conn.query(
          `DESCRIBE excel_data;`
        );

      const schema =
        normalizeDuckDBResult(
          describeResult
        );

      const numericColumns =
        schema.filter(
          item =>
            [
              "BIGINT",
              "INTEGER",
              "DOUBLE",
              "FLOAT",
              "DECIMAL",
              "HUGEINT"
            ].some(
              type =>
                String(
                  item.column_type
                )
                  .toUpperCase()
                  .includes(type)
            )
        );

      if (
        numericColumns.length === 0
      ) {

        throw new Error(
          "No se encontraron columnas numéricas."
        );

      }

      const expressions =
        [];

      for (
        const column of numericColumns
      ) {

        const name =
          quoteIdentifier(
            column.column_name
          );

        const aliasBase =
          column.column_name
            .replace(
              /[^a-zA-Z0-9_áéíóúÁÉÍÓÚñÑ]/g,
              "_"
            );

        expressions.push(
          `COUNT(${name}) AS "${aliasBase}__count"`
        );

        expressions.push(
          `SUM(${name}) AS "${aliasBase}__sum"`
        );

        expressions.push(
          `AVG(${name}) AS "${aliasBase}__avg"`
        );

        expressions.push(
          `MIN(${name}) AS "${aliasBase}__min"`
        );

        expressions.push(
          `MAX(${name}) AS "${aliasBase}__max"`
        );

      }

      const sql =
        `
SELECT
    ${expressions.join(",\n    ")}
FROM excel_data;
        `.trim();

      const textarea =
        document.getElementById(
          "tmxe-v0410-sql"
        );

      textarea.value =
        sql;

      await executeSQL();

    } catch (error) {

      setStatus(
        "❌ ERROR EN RESUMEN",
        "error"
      );

      setResult({
        mensaje:
          error?.message ||
          String(error)
      });

    }
  }

  /*
   * ============================================================
   * TOTAL POR AÑO
   * ============================================================
   *
   * Importante:
   *
   * Primero intentamos utilizar la columna
   * "Año" si existe.
   *
   * Esto evita el problema anterior de
   * GROUP BY / Fecha.
   */

  async function totalByYear() {

    try {

      if (
        !conn ||
        !currentRows.length
      ) {

        throw new Error(
          "Primero debes cargar un Excel."
        );

      }

      const schemaResult =
        await conn.query(
          `DESCRIBE excel_data;`
        );

      const schema =
        normalizeDuckDBResult(
          schemaResult
        );

      const hasYear =
        schema.some(
          item =>
            String(
              item.column_name
            ).toLowerCase() ===
            "año"
        );

      const hasTotal =
        schema.some(
          item =>
            String(
              item.column_name
            ).toLowerCase() ===
            "total teu's"
        );

      if (
        !hasTotal
      ) {

        throw new Error(
          'No se encontró la columna "Total TEU\'s".'
        );

      }

      let sql;

      if (
        hasYear
      ) {

        sql =
          `
SELECT
    "Año",
    SUM("Total TEU's") AS total
FROM excel_data
GROUP BY "Año"
ORDER BY "Año";
          `.trim();

      } else {

        const hasFecha =
          schema.some(
            item =>
              String(
                item.column_name
              ).toLowerCase() ===
              "fecha"
          );

        if (
          !hasFecha
        ) {

          throw new Error(
            'No existe "Año" ni "Fecha".'
          );

        }

        sql =
          `
SELECT
    EXTRACT(
      YEAR FROM "Fecha"
    ) AS año,
    SUM("Total TEU's") AS total
FROM excel_data
GROUP BY
    EXTRACT(
      YEAR FROM "Fecha"
    )
ORDER BY año;
          `.trim();

      }

      const textarea =
        document.getElementById(
          "tmxe-v0410-sql"
        );

      textarea.value =
        sql;

      await executeSQL();

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
   * ============================================================
   * ESQUEMA
   * ============================================================
   */

  async function showSchema() {

    try {

      if (
        !conn ||
        !currentRows.length
      ) {

        throw new Error(
          "Primero debes cargar un Excel."
        );

      }

      const sql =
        `
DESCRIBE excel_data;
        `.trim();

      document
        .getElementById(
          "tmxe-v0410-sql"
        )
        .value =
        sql;

      await executeSQL();

    } catch (error) {

      setStatus(
        "❌ ERROR AL OBTENER ESQUEMA",
        "error"
      );

      setResult({
        mensaje:
          error?.message ||
          String(error)
      });

    }
  }

  /*
   * ============================================================
   * INICIALIZACIÓN
   * ============================================================
   */

  function init() {

    addStyles();

    createButton();

    console.log(
      `[${APP_ID}] cargado correctamente`
    );

    console.log(
      `[${APP_ID}] versión ${VERSION}`
    );

  }

  /*
   * ============================================================
   * ARRANQUE
   * ============================================================
   */

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
