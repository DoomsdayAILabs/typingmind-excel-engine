/*
 * TypingMind Excel Data Engine
 * Extension v0.4.11 TEST
 *
 * OBJETIVO
 * --------
 * Motor Excel local para TypingMind.
 *
 * - El Excel permanece en el navegador.
 * - No se sube el archivo a ningún servidor.
 * - SheetJS lee el XLSX localmente.
 * - DuckDB-Wasm ejecuta SQL localmente.
 * - La tabla local se llama: excel_data
 *
 * CAMBIO PRINCIPAL v0.4.11
 * -------------------------
 * Corrección definitiva de resultados DuckDB-Wasm.
 *
 * v0.4.10 intentaba interpretar directamente:
 *
 *     row["total"]
 *
 * En Arrow/DuckDB-Wasm eso puede devolver una estructura
 * interna que termina apareciendo como:
 *
 *     [1535012, 0, 0, 0]
 *
 * v0.4.11 utiliza:
 *
 *     row.toJSON()
 *
 * para obtener el valor real de cada fila antes de
 * normalizarlo.
 *
 * Esto sigue el patrón recomendado por DuckDB-Wasm:
 *
 *     result.toArray().map(row => row.toJSON())
 *
 * DuckDB-Wasm:
 * @duckdb/duckdb-wasm@1.29.0
 *
 * SheetJS:
 * xlsx@0.18.5
 */

(() => {
  "use strict";

  const APP_ID = "tm-excel-engine-v0411-test";
  const VERSION = "v0.4.11-test";

  const DUCKDB_PACKAGE =
    "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

  const XLSX_PACKAGE =
    "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";

  let duckdb = null;
  let XLSX = null;

  let db = null;
  let conn = null;

  let worker = null;
  let workerURL = null;

  let currentFile = null;
  let currentWorkbook = null;

  let currentSheetName = null;
  let currentHeaders = [];
  let currentRows = [];

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

  function isEmptyValue(value) {
    if (value === null || value === undefined) {
      return true;
    }

    if (typeof value === "string") {
      return value.trim() === "";
    }

    return false;
  }

  /*
   * ============================================================
   * NORMALIZACIÓN DE VALORES
   * ============================================================
   *
   * Esta función NO intenta interpretar objetos numéricos
   * como escalares.
   *
   * Ese comportamiento fue eliminado en v0.4.11.
   *
   * Primero obtenemos el JSON real de Arrow mediante
   * row.toJSON().
   */

  function normalizeValue(value) {
    if (value === null || value === undefined) {
      return null;
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
     * Date
     */
    if (value instanceof Date) {
      if (Number.isNaN(value.getTime())) {
        return null;
      }

      return value.toISOString();
    }

    /*
     * ArrayBuffer
     */
    if (value instanceof ArrayBuffer) {
      return Array.from(new Uint8Array(value));
    }

    /*
     * TypedArrays
     */
    if (
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
    ) {
      return Array.from(value).map(normalizeValue);
    }

    /*
     * Array
     */
    if (Array.isArray(value)) {
      return value.map(normalizeValue);
    }

    /*
     * Object
     */
    if (typeof value === "object") {
      const result = {};

      for (const key of Object.keys(value)) {
        result[key] = normalizeValue(value[key]);
      }

      return result;
    }

    /*
     * Number
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
     * Boolean / String
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
   * ============================================================
   * NORMALIZAR FILA ARROW
   * ============================================================
   *
   * IMPORTANTE:
   *
   * v0.4.11 utiliza row.toJSON().
   *
   * Esto evita acceder directamente a las propiedades
   * internas de Arrow.
   */

  function normalizeArrowRow(row) {
    if (!row) {
      return {};
    }

    let jsonRow = null;

    /*
     * Método recomendado por DuckDB-Wasm / Arrow.
     */
    if (typeof row.toJSON === "function") {
      try {
        jsonRow = row.toJSON();
      } catch (error) {
        console.warn(
          `[${APP_ID}] Error usando row.toJSON()`,
          error
        );
      }
    }

    /*
     * Fallback solamente si toJSON no existe.
     */
    if (
      jsonRow === null ||
      jsonRow === undefined
    ) {
      jsonRow = row;
    }

    return normalizeValue(jsonRow);
  }

  /*
   * ============================================================
   * NORMALIZAR RESULTADO DUCKDB
   * ============================================================
   */

  function normalizeDuckDBResult(result) {
    if (!result) {
      return [];
    }

    let rows = [];

    try {
      rows = result.toArray();
    } catch (error) {
      console.error(
        `[${APP_ID}] Error ejecutando result.toArray()`,
        error
      );

      return [];
    }

    return rows.map(normalizeArrowRow);
  }

  /*
   * ============================================================
   * JSON SEGURO
   * ============================================================
   */

  function safeJSONStringify(value) {
    try {
      return JSON.stringify(
        normalizeValue(value),
        null,
        2
      );
    } catch (error) {
      return JSON.stringify(
        {
          error:
            "No fue posible serializar el resultado.",
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
        "tmxe-v0411-style"
      )
    ) {
      return;
    }

    const style = document.createElement("style");

    style.id = "tmxe-v0411-style";

    style.textContent = `
      #tmxe-v0411-button {
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

      #tmxe-v0411-overlay {
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

      #tmxe-v0411-panel {
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

      #tmxe-v0411-panel h2 {
        margin: 0 0 5px;
        font-size: 20px;
      }

      #tmxe-v0411-help {
        opacity: .7;
        margin-bottom: 15px;
      }

      #tmxe-v0411-status {
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

      #tmxe-v0411-status[data-kind="ok"] {
        color: #16a34a;
      }

      #tmxe-v0411-status[data-kind="error"] {
        color: #dc2626;
      }

      #tmxe-v0411-status[data-kind="busy"] {
        color: #7c3aed;
      }

      #tmxe-v0411-result {
        white-space: pre-wrap;

        overflow: auto;

        max-height: 500px;

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

      #tmxe-v0411-sql {
        width: 100%;

        min-height: 150px;

        resize: vertical;

        box-sizing: border-box;

        padding: 12px;

        border:
          1px solid
          rgba(127,127,127,.35);

        border-radius: 10px;

        background: Canvas;
        color: CanvasText;

        font:
          13px/1.5
          ui-monospace,
          SFMono-Regular,
          Menlo,
          monospace;
      }

      .tmxe-v0411-actions {
        display: flex;

        gap: 8px;

        flex-wrap: wrap;

        margin: 10px 0;
      }

      .tmxe-v0411-btn {
        border:
          1px solid
          rgba(127,127,127,.45);

        border-radius: 9px;

        padding: 9px 13px;

        background: transparent;

        color: inherit;

        cursor: pointer;
      }

      .tmxe-v0411-btn:hover {
        opacity: .85;
      }

      .tmxe-v0411-primary {
        background: #7c3aed !important;
        color: white !important;
        border-color: #7c3aed !important;
      }

      .tmxe-v0411-success {
        background: #16a34a !important;
        color: white !important;
        border-color: #16a34a !important;
      }

      .tmxe-v0411-section {
        margin-top: 16px;
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
        "tmxe-v0411-button"
      )
    ) {
      return;
    }

    const button =
      document.createElement("button");

    button.id =
      "tmxe-v0411-button";

    button.type = "button";

    button.textContent =
      "📊 Excel v0.4.11";

    button.title =
      "Excel Data Engine v0.4.11";

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
        "tmxe-v0411-overlay"
      )
    ) {
      return;
    }

    const overlay =
      document.createElement("div");

    overlay.id =
      "tmxe-v0411-overlay";

    overlay.innerHTML = `
      <div
        id="tmxe-v0411-panel"
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

            <div id="tmxe-v0411-help">
              Procesamiento Excel + DuckDB
              completamente local.
            </div>

          </div>

          <button
            id="tmxe-v0411-close"
            class="tmxe-v0411-btn"
            type="button"
          >
            Cerrar
          </button>

        </div>

        <div id="tmxe-v0411-status">
          Listo para cargar un archivo Excel.
        </div>

        <div class="tmxe-v0411-actions">

          <input
            id="tmxe-v0411-file"
            type="file"
            accept=".xlsx,.xls,.xlsm"
            style="display:none"
          >

          <button
            id="tmxe-v0411-load"
            class="tmxe-v0411-btn tmxe-v0411-primary"
            type="button"
          >
            📂 Cargar Excel
          </button>

          <button
            id="tmxe-v0411-count"
            class="tmxe-v0411-btn"
            type="button"
          >
            🔢 COUNT
          </button>

          <button
            id="tmxe-v0411-preview"
            class="tmxe-v0411-btn"
            type="button"
          >
            👁 Vista previa
          </button>

          <button
            id="tmxe-v0411-summary"
            class="tmxe-v0411-btn"
            type="button"
          >
            📈 Resumen
          </button>

          <button
            id="tmxe-v0411-year"
            class="tmxe-v0411-btn"
            type="button"
          >
            📅 Total por año
          </button>

          <button
            id="tmxe-v0411-schema"
            class="tmxe-v0411-btn"
            type="button"
          >
            🧱 Esquema
          </button>

        </div>

        <div class="tmxe-v0411-section">

          <strong>
            SQL
          </strong>

          <textarea
            id="tmxe-v0411-sql"
            spellcheck="false"
          >SELECT COUNT(*) AS registros
FROM excel_data;</textarea>

          <div class="tmxe-v0411-actions">

            <button
              id="tmxe-v0411-run-sql"
              class="tmxe-v0411-btn tmxe-v0411-success"
              type="button"
            >
              ▶ Ejecutar SQL
            </button>

          </div>

        </div>

        <div class="tmxe-v0411-section">

          <strong>
            Resultado
          </strong>

          <pre
            id="tmxe-v0411-result"
          >—</pre>

        </div>

        <div style="
          margin-top:12px;
          opacity:.65;
          font-size:12px;
        ">
          ${VERSION}<br>
          DuckDB-Wasm 1.29.0<br>
          DuckDB interno v1.1.1<br>
          SheetJS 0.18.5<br>
          Procesamiento 100% local.
        </div>

      </div>
    `;

    document.body.appendChild(overlay);

    document
      .getElementById(
        "tmxe-v0411-close"
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
        "tmxe-v0411-load"
      )
      .addEventListener(
        "click",
        () => {
          document
            .getElementById(
              "tmxe-v0411-file"
            )
            .click();
        }
      );

    document
      .getElementById(
        "tmxe-v0411-file"
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
        "tmxe-v0411-run-sql"
      )
      .addEventListener(
        "click",
        executeSQL
      );

    document
      .getElementById(
        "tmxe-v0411-count"
      )
      .addEventListener(
        "click",
        countRows
      );

    document
      .getElementById(
        "tmxe-v0411-preview"
      )
      .addEventListener(
        "click",
        previewRows
      );

    document
      .getElementById(
        "tmxe-v0411-summary"
      )
      .addEventListener(
        "click",
        summaryRows
      );

    document
      .getElementById(
        "tmxe-v0411-year"
      )
      .addEventListener(
        "click",
        totalByYear
      );

    document
      .getElementById(
        "tmxe-v0411-schema"
      )
      .addEventListener(
        "click",
        showSchema
      );
  }

  /*
   * ============================================================
   * STATUS
   * ============================================================
   */

  function setStatus(
    text,
    kind = ""
  ) {
    const element =
      document.getElementById(
        "tmxe-v0411-status"
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

  /*
   * ============================================================
   * RESULTADO
   * ============================================================
   */

  function setResult(value) {
    const element =
      document.getElementById(
        "tmxe-v0411-result"
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
      return null;
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
      "busy"
    );

    return bundle;
  }

  /*
   * ============================================================
   * SHEETJS
   * ============================================================
   */

  async function loadXLSXLibrary() {
    if (XLSX) {
      return;
    }

    XLSX =
      await import(
        XLSX_PACKAGE
      );
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
   * FECHAS EXCEL
   * ============================================================
   */

  function excelSerialToDate(
    serial
  ) {
    const excelEpoch =
      Date.UTC(
        1899,
        11,
        30
      );

    return new Date(
      excelEpoch +
      Number(serial) *
      24 *
      60 *
      60 *
      1000
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
        /^\d{4}-\d{1,2}-\d{1,2}$/
          .test(text)
      ) {
        date =
          new Date(
            text +
            "T00:00:00Z"
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

    return date
      .toISOString()
      .slice(0, 10);
  }

  /*
   * ============================================================
   * DETECCIÓN DE TIPOS
   * ============================================================
   */

  function isIntegerValue(
    value
  ) {
    if (
      value === null ||
      value === undefined ||
      value === ""
    ) {
      return false;
    }

    if (
      typeof value === "number"
    ) {
      return Number.isInteger(
        value
      );
    }

    return /^[-+]?\d+$/.test(
      String(value)
        .trim()
        .replace(/,/g, "")
    );
  }

  function isNumberValue(
    value
  ) {
    if (
      value === null ||
      value === undefined ||
      value === ""
    ) {
      return false;
    }

    if (
      typeof value === "number"
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

  function looksLikeDate(
    value
  ) {
    if (
      value instanceof Date
    ) {
      return true;
    }

    if (
      typeof value === "number"
    ) {
      return (
        value >= 20000 &&
        value <= 80000
      );
    }

    const text =
      String(value)
        .trim();

    return /^\d{4}-\d{1,2}-\d{1,2}$/
      .test(text);
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
        sqlType: "VARCHAR",
        confidence: "low",
        reason:
          "columna vacía",
        non_empty_values: 0
      };
    }

    const headerText =
      String(header)
        .toLowerCase();

    const dateHeader =
      /fecha|date|día|dia/.test(
        headerText
      );

    const allDates =
      nonEmpty.every(
        looksLikeDate
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
        isIntegerValue
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
        isNumberValue
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

      return `DATE '${date}'`;
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
      const number =
        Number(
          String(value)
            .trim()
            .replace(/,/g, "")
        );

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

      currentFile =
        file;

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
       * Buscar encabezado.
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
       * Determinar columnas.
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

      /*
       * Encabezados.
       */

      currentHeaders = [];

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
       * Encabezados únicos.
       */

      const used =
        new Map();

      currentHeaders =
        currentHeaders.map(
          header => {
            const count =
              used.get(header) ||
              0;

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

        const normalized = [];

        let hasData =
          false;

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
       * Inferencia.
       */

      const inferredTypes =
        currentHeaders.map(
          (
            header,
            index
          ) => {
            const values =
              currentRows.map(
                row =>
                  row[index]
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
       * Eliminar tabla anterior.
       */

      try {
        await conn.query(
          `DROP TABLE IF EXISTS excel_data;`
        );
      } catch (_) {}

      /*
       * Crear tabla.
       */

      const definitions =
        inferredTypes
          .map(
            info =>
              `${quoteIdentifier(
                info.column_name
              )} ${
                info.sqlType ||
                info.duckdb_type
              }`
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
       */

      const BATCH_SIZE =
        500;

      for (
        let start = 0;
        start <
        currentRows.length;
        start +=
          BATCH_SIZE
      ) {
        const batch =
          currentRows.slice(
            start,
            start +
              BATCH_SIZE
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
                    (
                      value,
                      index
                    ) =>
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

      /*
       * COUNT
       */

      const countResult =
        await conn.query(
          `
          SELECT
            COUNT(*) AS registros
          FROM excel_data;
          `
        );

      const count =
        normalizeDuckDBResult(
          countResult
        );

      /*
       * SCHEMA
       */

      const schemaResult =
        await conn.query(
          `DESCRIBE excel_data;`
        );

      const schema =
        normalizeDuckDBResult(
          schemaResult
        );

      /*
       * PREVIEW
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
          Math.max(
            0,
            physicalRows -
            headerIndex -
            1 -
            currentRows.length
          ),

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

          count,

        schema:
          schema,

        preview:
          preview,

        bundle: {
          mainModule:
            bundle?.mainModule ||
            null,

          mainWorker:
            bundle?.mainWorker ||
            null,

          pthreadWorker:
            bundle?.pthreadWorker ||
            null
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
      if (!conn) {
        await initializeDuckDB();
      }

      if (
        !currentRows.length
      ) {
        throw new Error(
          "Primero debes cargar un archivo Excel."
        );
      }

      const textarea =
        document.getElementById(
          "tmxe-v0411-sql"
        );

      const sql =
        textarea.value.trim();

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

      /*
       * AQUÍ ESTÁ EL CAMBIO CRÍTICO.
       *
       * normalizeDuckDBResult()
       *
       * utiliza row.toJSON()
       * antes de normalizar.
       */

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
        "tmxe-v0411-sql"
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
        "tmxe-v0411-sql"
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

      const schemaResult =
        await conn.query(
          `DESCRIBE excel_data;`
        );

      const schema =
        normalizeDuckDBResult(
          schemaResult
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
                  .includes(
                    type
                  )
            )
        );

      if (
        numericColumns.length === 0
      ) {
        throw new Error(
          "No se encontraron columnas numéricas."
        );
      }

      const expressions = [];

      for (
        const column
        of numericColumns
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
    ${expressions.join(
      ",\n    "
    )}
FROM excel_data;
        `.trim();

      document
        .getElementById(
          "tmxe-v0411-sql"
        )
        .value = sql;

      await executeSQL();

    } catch (error) {
      setStatus(
        "❌ ERROR EN RESUMEN",
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
   * TOTAL POR AÑO
   * ============================================================
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
            )
              .toLowerCase() ===
            "año"
        );

      const hasTotal =
        schema.some(
          item =>
            String(
              item.column_name
            )
              .toLowerCase() ===
            "total teus"
        );

      if (!hasTotal) {
        throw new Error(
          'No se encontró la columna "Total TEU\'s".'
        );
      }

      let sql;

      if (hasYear) {
        sql =
          `
SELECT
    "Año",
    SUM("Total TEUs") AS total
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
              )
                .toLowerCase() ===
              "fecha"
          );

        if (!hasFecha) {
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
    SUM("Total TEUs") AS total
FROM excel_data
GROUP BY
    EXTRACT(
      YEAR FROM "Fecha"
    )
ORDER BY año;
          `.trim();
      }

      document
        .getElementById(
          "tmxe-v0411-sql"
        )
        .value = sql;

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
          "tmxe-v0411-sql"
        )
        .value = sql;

      await executeSQL();

    } catch (error) {
      setStatus(
        "❌ ERROR AL OBTENER ESQUEMA",
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
