/*
 * TypingMind Excel Data Engine
 * v0.4.9.2 TEST
 *
 * OBJETIVO
 * ------------------------------------------------------------
 * Excel local -> DuckDB-Wasm local -> SQL local
 *
 * Esta versión corrige:
 *
 * 1. SQL del botón "Total por año".
 * 2. Conversión de HUGEINT/BIGINT de DuckDB-Wasm.
 * 3. Resultados tipo [1535012,0,0,0].
 *
 * NO se envía el Excel a ningún servidor.
 *
 * Dependencias:
 * - DuckDB-Wasm 1.29.0
 * - SheetJS XLSX 0.18.5
 */

(() => {
  "use strict";

  const APP_ID = "tm-excel-engine-v0492-test";

  const DUCKDB_PACKAGE =
    "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

  const XLSX_PACKAGE =
    "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";

  const BUTTON_ID = "tmxe-v0492-button";
  const OVERLAY_ID = "tmxe-v0492-overlay";
  const STYLE_ID = "tmxe-v0492-style";

  let duckdb = null;
  let XLSX = null;

  let db = null;
  let conn = null;

  let currentFile = null;
  let currentTable = "excel_data";

  let currentSchema = [];
  let currentRows = 0;

  /* =========================================================
   * ESTILOS
   * ======================================================= */

  function addStyles() {

    if (document.getElementById(STYLE_ID)) {
      return;
    }

    const style = document.createElement("style");

    style.id = STYLE_ID;

    style.textContent = `
      #${BUTTON_ID} {
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

      #${BUTTON_ID}:hover {
        filter: brightness(1.08);
      }

      #${OVERLAY_ID} {
        position: fixed;
        inset: 0;

        z-index: 2147483001;

        background: rgba(0,0,0,.48);

        display: flex;

        align-items: center;
        justify-content: center;

        padding: 16px;
      }

      #tmxe-v0492-panel {
        width: min(1100px, 97vw);

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

      #tmxe-v0492-panel h2 {
        margin: 0 0 4px;
        font-size: 20px;
      }

      #tmxe-v0492-help {
        opacity: .7;
        margin-bottom: 14px;
      }

      #tmxe-v0492-status {
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

      #tmxe-v0492-status[data-kind="ok"] {
        color: #16a34a;
      }

      #tmxe-v0492-status[data-kind="error"] {
        color: #dc2626;
      }

      #tmxe-v0492-file {
        margin: 10px 0;
        font-weight: 600;
      }

      #tmxe-v0492-sql {
        width: 100%;

        min-height: 170px;

        box-sizing: border-box;

        resize: vertical;

        padding: 12px;

        border-radius: 10px;

        border: 1px solid rgba(127,127,127,.45);

        background: rgba(127,127,127,.08);

        color: inherit;

        font: 13px/1.5 ui-monospace,
              SFMono-Regular,
              Menlo,
              monospace;
      }

      #tmxe-v0492-actions {
        display: flex;

        flex-wrap: wrap;

        gap: 8px;

        margin: 10px 0;
      }

      #tmxe-v0492-actions button,
      #tmxe-v0492-close {
        border: 1px solid rgba(127,127,127,.45);

        border-radius: 9px;

        padding: 9px 13px;

        background: transparent;

        color: inherit;

        cursor: pointer;
      }

      #tmxe-v0492-load,
      #tmxe-v0492-run {
        background: #7c3aed !important;
        color: white !important;
        border-color: #7c3aed !important;
      }

      #tmxe-v0492-result {
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

      #tmxe-v0492-info {
        display: grid;

        grid-template-columns:
          repeat(auto-fit,minmax(160px,1fr));

        gap: 8px;

        margin: 10px 0;
      }

      .tmxe-v0492-card {
        padding: 10px;

        border-radius: 9px;

        background: rgba(127,127,127,.08);
      }

      .tmxe-v0492-card small {
        display: block;
        opacity: .65;
        margin-bottom: 3px;
      }

      .tmxe-v0492-card strong {
        font-size: 15px;
      }

      #tmxe-v0492-input {
        display: none;
      }
    `;

    document.head.appendChild(style);
  }

  /* =========================================================
   * UI
   * ======================================================= */

  function createButton() {

    if (document.getElementById(BUTTON_ID)) {
      return;
    }

    const button = document.createElement("button");

    button.id = BUTTON_ID;

    button.type = "button";

    button.textContent = "📊 Excel v0.4.9.2";

    button.title =
      "Excel Data Engine v0.4.9.2";

    button.addEventListener(
      "click",
      openPanel
    );

    document.body.appendChild(button);
  }

  function openPanel() {

    if (document.getElementById(OVERLAY_ID)) {
      return;
    }

    const overlay = document.createElement("div");

    overlay.id = OVERLAY_ID;

    overlay.innerHTML = `
      <div
        id="tmxe-v0492-panel"
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
              📊 Excel Data Engine v0.4.9.2 TEST
            </h2>

            <div id="tmxe-v0492-help">
              Excel local + DuckDB-Wasm + SQL manual.
              Todo el procesamiento se realiza localmente.
            </div>

          </div>

          <button
            id="tmxe-v0492-close"
            type="button"
          >
            Cerrar
          </button>

        </div>

        <div id="tmxe-v0492-status">
          Listo.
        </div>

        <div id="tmxe-v0492-info">

          <div class="tmxe-v0492-card">
            <small>Archivo</small>
            <strong id="tmxe-v0492-file-name">—</strong>
          </div>

          <div class="tmxe-v0492-card">
            <small>Tabla</small>
            <strong>excel_data</strong>
          </div>

          <div class="tmxe-v0492-card">
            <small>Registros</small>
            <strong id="tmxe-v0492-row-count">—</strong>
          </div>

          <div class="tmxe-v0492-card">
            <small>Procesamiento</small>
            <strong>LOCAL</strong>
          </div>

        </div>

        <input
          id="tmxe-v0492-input"
          type="file"
          accept=".xlsx,.xls"
        />

        <div style="
          display:flex;
          gap:8px;
          align-items:center;
          flex-wrap:wrap;
        ">

          <button
            id="tmxe-v0492-load"
            type="button"
          >
            📂 Cargar Excel
          </button>

          <span id="tmxe-v0492-file">
            Ningún archivo cargado.
          </span>

        </div>

        <div style="margin-top:14px;">
          <strong>SQL</strong>
        </div>

        <textarea
          id="tmxe-v0492-sql"
          spellcheck="false"
        >SELECT
    COUNT(*) AS registros
FROM excel_data;</textarea>

        <div id="tmxe-v0492-actions">

          <button
            id="tmxe-v0492-count"
            type="button"
          >
            COUNT
          </button>

          <button
            id="tmxe-v0492-preview"
            type="button"
          >
            Vista previa
          </button>

          <button
            id="tmxe-v0492-summary"
            type="button"
          >
            Resumen
          </button>

          <button
            id="tmxe-v0492-year"
            type="button"
          >
            Total por año
          </button>

        </div>

        <button
          id="tmxe-v0492-run"
          type="button"
        >
          ▶ Ejecutar SQL
        </button>

        <div style="margin-top:16px;">
          <strong>Resultado</strong>
        </div>

        <pre id="tmxe-v0492-result">—</pre>

        <div style="
          margin-top:12px;
          opacity:.65;
          font-size:12px;
        ">
          v0.4.9.2 TEST —
          DuckDB-Wasm 1.29.0 —
          procesamiento 100 % local.
        </div>

      </div>
    `;

    document.body.appendChild(overlay);

    document
      .getElementById("tmxe-v0492-close")
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
      .getElementById("tmxe-v0492-load")
      .addEventListener(
        "click",
        () => {
          document
            .getElementById("tmxe-v0492-input")
            .click();
        }
      );

    document
      .getElementById("tmxe-v0492-input")
      .addEventListener(
        "change",
        handleFile
      );

    document
      .getElementById("tmxe-v0492-run")
      .addEventListener(
        "click",
        executeSQL
      );

    document
      .getElementById("tmxe-v0492-count")
      .addEventListener(
        "click",
        () => setSQL(`
SELECT
    COUNT(*) AS registros
FROM excel_data;
`)
      );

    document
      .getElementById("tmxe-v0492-preview")
      .addEventListener(
        "click",
        () => setSQL(`
SELECT *
FROM excel_data
LIMIT 10;
`)
      );

    document
      .getElementById("tmxe-v0492-summary")
      .addEventListener(
        "click",
        () => setSQL(buildSummarySQL())
      );

    document
      .getElementById("tmxe-v0492-year")
      .addEventListener(
        "click",
        () => setSQL(buildYearSQL())
      );
  }

  function setStatus(text, kind = "") {

    const element =
      document.getElementById(
        "tmxe-v0492-status"
      );

    if (!element) {
      return;
    }

    element.textContent = text;

    if (kind) {
      element.dataset.kind = kind;
    } else {
      delete element.dataset.kind;
    }
  }

  function setResult(value) {

    const element =
      document.getElementById(
        "tmxe-v0492-result"
      );

    if (!element) {
      return;
    }

    try {

      const normalized =
        normalizeForJSON(value);

      element.textContent =
        typeof normalized === "string"
          ? normalized
          : JSON.stringify(
              normalized,
              null,
              2
            );

    } catch (error) {

      element.textContent =
        String(value);
    }
  }

  function setSQL(sql) {

    const textarea =
      document.getElementById(
        "tmxe-v0492-sql"
      );

    if (!textarea) {
      return;
    }

    textarea.value =
      String(sql).trim();

    setStatus(
      "SQL preparado. Presiona «Ejecutar SQL»."
    );
  }

  /* =========================================================
   * NORMALIZACIÓN DE VALORES DUCKDB-WASM
   * ======================================================= */

  function normalizeForJSON(value) {

    if (value === null ||
        value === undefined) {

      return value;
    }

    if (typeof value === "bigint") {

      const numeric =
        Number(value);

      if (
        Number.isSafeInteger(numeric)
      ) {
        return numeric;
      }

      return value.toString();
    }

    if (value instanceof Date) {

      return value.toISOString();
    }

    if (
      typeof ArrayBuffer !== "undefined" &&
      ArrayBuffer.isView(value)
    ) {

      return normalizeTypedArray(value);
    }

    if (Array.isArray(value)) {

      return value.map(
        item => normalizeForJSON(item)
      );
    }

    if (typeof value === "object") {

      const keys =
        Object.keys(value);

      /*
       * DuckDB-Wasm puede exponer ciertos
       * valores numéricos grandes como una
       * estructura de 4 palabras.
       *
       * Ejemplo observado:
       *
       * {
       *   0: 1535012,
       *   1: 0,
       *   2: 0,
       *   3: 0
       * }
       *
       * Se intenta reconstruir únicamente
       * estructuras numéricas de este tipo.
       */

      if (
        keys.length === 4 &&
        keys.every(
          (key, index) =>
            key === String(index)
        ) &&
        keys.every(
          key =>
            typeof value[key] ===
              "number" ||
            typeof value[key] ===
              "bigint"
        )
      ) {

        return normalizeFourWordInteger(
          keys.map(
            key => value[key]
          )
        );
      }

      const result = {};

      for (const key of keys) {

        result[key] =
          normalizeForJSON(
            value[key]
          );
      }

      return result;
    }

    return value;
  }

  function normalizeTypedArray(value) {

    const array =
      Array.from(value);

    /*
     * Si es un vector de un solo valor,
     * devolver directamente el valor.
     */

    if (array.length === 1) {

      return normalizeForJSON(
        array[0]
      );
    }

    /*
     * Para resultados de DuckDB,
     * conservar arrays normales.
     */

    return array.map(
      item => normalizeForJSON(item)
    );
  }

  function normalizeFourWordInteger(parts) {

    try {

      const nums =
        parts.map(
          value =>
            typeof value === "bigint"
              ? value
              : BigInt(value)
        );

      /*
       * Caso habitual observado:
       *
       * [valor, 0, 0, 0]
       *
       * Para valores positivos pequeños
       * basta devolver la primera palabra.
       */

      if (
        nums[1] === 0n &&
        nums[2] === 0n &&
        nums[3] === 0n
      ) {

        const n =
          Number(nums[0]);

        if (
          Number.isSafeInteger(n)
        ) {
          return n;
        }

        return nums[0].toString();
      }

      /*
       * Reconstrucción de un entero
       * de hasta 128 bits usando palabras
       * de 32 bits.
       */

      let result = 0n;

      for (
        let i = 3;
        i >= 0;
        i--
      ) {

        result =
          (result << 32n) +
          nums[i];
      }

      const number =
        Number(result);

      if (
        Number.isSafeInteger(number)
      ) {

        return number;
      }

      return result.toString();

    } catch (error) {

      return parts.map(
        value =>
          typeof value === "bigint"
            ? value.toString()
            : value
      );
    }
  }

  /* =========================================================
   * DUCKDB
   * ======================================================= */

  async function initializeDuckDB() {

    if (conn) {
      return;
    }

    setStatus(
      "1/4 — Cargando DuckDB-Wasm..."
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
      "2/4 — Inicializando WebAssembly..."
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

    /*
     * Intentamos obtener la extensión
     * Excel, pero NO dependemos de ella.
     *
     * El XLSX se procesa mediante SheetJS
     * y los datos se insertan localmente
     * en DuckDB.
     */

    try {

      await conn.query(
        "INSTALL excel;"
      );

    } catch (_) {
      /*
       * No es un error fatal.
       *
       * DuckDB-Wasm 1.1.1 no dispone
       * de la extensión Excel remota
       * en este entorno.
       */
    }

    setStatus(
      "3/4 — DuckDB conectado."
    );
  }

  /* =========================================================
   * XLSX
   * ======================================================= */

  async function initializeXLSX() {

    if (XLSX) {
      return;
    }

    XLSX =
      await import(
        XLSX_PACKAGE
      );
  }

  /* =========================================================
   * CARGAR ARCHIVO
   * ======================================================= */

  async function handleFile(event) {

    const file =
      event.target.files &&
      event.target.files[0];

    if (!file) {
      return;
    }

    try {

      await loadExcel(file);

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

  async function loadExcel(file) {

    setResult("");

    setStatus(
      "Preparando carga local..."
    );

    await initializeDuckDB();

    await initializeXLSX();

    currentFile = file;

    const arrayBuffer =
      await file.arrayBuffer();

    setStatus(
      "4/4 — Leyendo XLSX localmente..."
    );

    const workbook =
      XLSX.read(
        arrayBuffer,
        {
          type: "array",
          cellDates: true,
          cellNF: true,
          cellText: false
        }
      );

    const sheetNames =
      workbook.SheetNames || [];

    if (!sheetNames.length) {

      throw new Error(
        "El archivo no contiene hojas."
      );
    }

    const sheetName =
      sheetNames[0];

    const worksheet =
      workbook.Sheets[
        sheetName
      ];

    if (!worksheet) {

      throw new Error(
        "No se pudo obtener la hoja principal."
      );
    }

    const range =
      worksheet["!ref"];

    const physicalRows =
      range
        ? XLSX.utils.decode_range(
            range
          ).e.r + 1
        : 0;

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

    /*
     * Buscar automáticamente la primera
     * fila que parezca encabezado.
     */

    const headerIndex =
      findHeaderRow(matrix);

    if (headerIndex < 0) {

      throw new Error(
        "No se encontró una fila de encabezados válida."
      );
    }

    const rawHeaders =
      matrix[headerIndex] || [];

    const headers =
      makeUniqueHeaders(
        rawHeaders
      );

    const dataRows =
      matrix.slice(
        headerIndex + 1
      );

    const realRows =
      dataRows.filter(
        row =>
          !isEmptyRow(
            row,
            headers.length
          )
      );

    currentRows =
      realRows.length;

    currentSchema =
      inferSchema(
        headers,
        realRows
      );

    await createDuckDBTable(
      headers,
      currentSchema,
      realRows
    );

    updateFileInfo(
      file.name,
      currentRows
    );

    setStatus(
      "✅ Excel cargado y consultado localmente.",
      "ok"
    );

    /*
     * Diagnóstico inicial.
     */

    const countResult =
      await conn.query(
        `
        SELECT COUNT(*) AS registros
        FROM ${quoteIdentifier(currentTable)};
        `
      );

    const countRows =
      normalizeRows(
        countResult.toArray()
      );

    setResult({
      procesamiento: "LOCAL",

      archivo:
        file.name,

      tamano_bytes:
        file.size,

      hojas:
        sheetNames,

      hoja_principal:
        sheetName,

      filas_fisicas_detectadas:
        physicalRows,

      filas_reales:
        realRows.length,

      filas_vacias_ignoradas:
        Math.max(
          0,
          physicalRows -
          headerIndex -
          1 -
          realRows.length
        ),

      filas_insertadas:
        realRows.length,

      columnas_detectadas:
        headers.length,

      encabezados:
        headers,

      tipos_inferidos:
        currentSchema,

      duckdb_table:
        currentTable,

      duckdb_version:
        await getDuckDBVersion(),

      count:
        countRows
    });
  }

  /* =========================================================
   * DETECTAR ENCABEZADOS
   * ======================================================= */

  function findHeaderRow(matrix) {

    for (
      let i = 0;
      i < matrix.length;
      i++
    ) {

      const row =
        matrix[i];

      if (!Array.isArray(row)) {
        continue;
      }

      const nonEmpty =
        row.filter(
          value =>
            value !== null &&
            value !== undefined &&
            String(value).trim() !== ""
        );

      /*
       * Para una tabla Excel normal,
       * 2 o más encabezados es suficiente.
       */

      if (nonEmpty.length >= 2) {

        return i;
      }
    }

    return -1;
  }

  /* =========================================================
   * ENCABEZADOS
   * ======================================================= */

  function makeUniqueHeaders(rawHeaders) {

    const used =
      new Map();

    return rawHeaders.map(
      (header, index) => {

        let name =
          header === null ||
          header === undefined ||
          String(header).trim() === ""
            ? `Column_${index + 1}`
            : String(header).trim();

        if (
          used.has(name)
        ) {

          const count =
            used.get(name) + 1;

          used.set(
            name,
            count
          );

          name =
            `${name}_${count}`;

        } else {

          used.set(
            name,
            1
          );
        }

        return name;
      }
    );
  }

  /* =========================================================
   * FILAS VACÍAS
   * ======================================================= */

  function isEmptyRow(
    row,
    columnCount
  ) {

    for (
      let i = 0;
      i < columnCount;
      i++
    ) {

      const value =
        row?.[i];

      if (
        value !== null &&
        value !== undefined &&
        String(value).trim() !== ""
      ) {

        return false;
      }
    }

    return true;
  }

  /* =========================================================
   * INFERENCIA DE TIPOS
   * ======================================================= */

  function inferSchema(
    headers,
    rows
  ) {

    return headers.map(
      (header, columnIndex) => {

        const values =
          rows
            .map(
              row =>
                row?.[columnIndex]
            )
            .filter(
              value =>
                value !== null &&
                value !== undefined &&
                String(value).trim() !== ""
            );

        const nonEmptyValues =
          values.length;

        if (
          looksLikeDateColumn(
            header,
            values
          )
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
              nonEmptyValues
          };
        }

        if (
          values.length > 0 &&
          values.every(
            value =>
              isIntegerValue(
                value
              )
          )
        ) {

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
              nonEmptyValues
          };
        }

        if (
          values.length > 0 &&
          values.every(
            value =>
              isNumberValue(
                value
              )
          )
        ) {

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
              nonEmptyValues
          };
        }

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
            nonEmptyValues
        };
      }
    );
  }

  function looksLikeDateColumn(
    header,
    values
  ) {

    const normalized =
      String(header)
        .toLowerCase()
        .normalize("NFD")
        .replace(
          /[\u0300-\u036f]/g,
          ""
        );

    const dateHeader =
      normalized.includes("fecha") ||
      normalized.includes("date") ||
      normalized.includes("dia");

    if (!dateHeader) {
      return false;
    }

    if (!values.length) {
      return false;
    }

    return values.every(
      value =>
        isDateValue(value)
    );
  }

  function isDateValue(value) {

    if (
      value instanceof Date &&
      !Number.isNaN(
        value.getTime()
      )
    ) {

      return true;
    }

    if (
      typeof value === "number"
    ) {

      /*
       * Excel serial dates típicas.
       */

      return (
        value > 20000 &&
        value < 100000
      );
    }

    const parsed =
      Date.parse(
        String(value)
      );

    return !Number.isNaN(parsed);
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

    if (
      typeof value === "bigint"
    ) {

      return true;
    }

    const text =
      String(value).trim();

    return /^[-+]?\d+$/.test(
      text
    );
  }

  function isNumberValue(value) {

    if (
      typeof value === "number"
    ) {

      return Number.isFinite(value);
    }

    const text =
      String(value).trim();

    return (
      text !== "" &&
      Number.isFinite(
        Number(text)
      )
    );
  }

  /* =========================================================
   * DUCKDB TABLE
   * ======================================================= */

  async function createDuckDBTable(
    headers,
    schema,
    rows
  ) {

    if (!conn) {

      throw new Error(
        "DuckDB no está conectado."
      );
    }

    try {

      await conn.query(
        `DROP TABLE IF EXISTS ${quoteIdentifier(currentTable)};`
      );

    } catch (_) {}

    const definitions =
      schema.map(
        column =>
          `${quoteIdentifier(column.column_name)} ${column.sqlType}`
      );

    await conn.query(
      `
      CREATE TABLE ${quoteIdentifier(currentTable)}
      (
        ${definitions.join(",\n")}
      );
      `
    );

    if (!rows.length) {
      return;
    }

    /*
     * Insertar en lotes.
     *
     * Esto evita construir una consulta
     * gigantesca para archivos grandes.
     */

    const BATCH_SIZE = 1000;

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

      const valueGroups =
        batch.map(
          row => {

            const values =
              headers.map(
                (header, index) => {

                  const column =
                    schema[index];

                  const value =
                    row?.[index];

                  return sqlLiteral(
                    value,
                    column.sqlType
                  );
                }
              );

            return `(${values.join(",")})`;
          }
        );

      await conn.query(
        `
        INSERT INTO ${quoteIdentifier(currentTable)}
        VALUES
        ${valueGroups.join(",\n")};
        `
      );
    }
  }

  /* =========================================================
   * SQL LITERALS
   * ======================================================= */

  function sqlLiteral(
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

    if (type === "DATE") {

      const date =
        excelValueToDate(
          value
        );

      if (!date) {
        return "NULL";
      }

      const yyyy =
        date.getUTCFullYear();

      const mm =
        String(
          date.getUTCMonth() + 1
        ).padStart(2, "0");

      const dd =
        String(
          date.getUTCDate()
        ).padStart(2, "0");

      return `DATE '${yyyy}-${mm}-${dd}'`;
    }

    if (type === "BIGINT") {

      const n =
        BigInt(
          String(value)
            .trim()
        );

      return n.toString();
    }

    if (type === "DOUBLE") {

      const n =
        Number(value);

      if (!Number.isFinite(n)) {
        return "NULL";
      }

      return String(n);
    }

    return `'${escapeSQLString(
      String(value)
    )}'`;
  }

  function escapeSQLString(value) {

    return String(value)
      .replace(
        /'/g,
        "''"
      );
  }

  function excelValueToDate(
    value
  ) {

    if (
      value instanceof Date
    ) {

      return value;
    }

    if (
      typeof value === "number" &&
      value > 20000 &&
      value < 100000
    ) {

      /*
       * Excel serial date.
       *
       * Excel epoch:
       * 1899-12-30
       */

      const milliseconds =
        Math.round(
          (
            value -
            25569
          ) *
          86400000
        );

      return new Date(
        milliseconds
      );
    }

    const parsed =
      new Date(
        String(value)
      );

    if (
      Number.isNaN(
        parsed.getTime()
      )
    ) {

      return null;
    }

    return parsed;
  }

  function quoteIdentifier(
    name
  ) {

    return `"${String(name)
      .replace(
        /"/g,
        '""'
      )}"`;
  }

  /* =========================================================
   * SQL PREDEFINIDO
   * ======================================================= */

  function buildYearSQL() {

    /*
     * IMPORTANTE:
     *
     * No usamos EXTRACT(YEAR FROM Fecha)
     * aquí porque el Excel ya contiene
     * una columna "Año".
     *
     * Además evitamos cualquier problema
     * de alias o GROUP BY.
     */

    return `
SELECT
    "Año",
    SUM("Total TEU's") AS total
FROM excel_data
GROUP BY "Año"
ORDER BY "Año";
`;
  }

  function buildSummarySQL() {

    if (!currentSchema.length) {

      return `
SELECT
    COUNT(*) AS registros
FROM excel_data;
`;
    }

    const numericColumns =
      currentSchema.filter(
        column =>
          column.sqlType === "BIGINT" ||
          column.sqlType === "DOUBLE"
      );

    if (!numericColumns.length) {

      return `
SELECT
    COUNT(*) AS registros
FROM excel_data;
`;
    }

    const expressions =
      numericColumns.map(
        column => {

          const q =
            quoteIdentifier(
              column.column_name
            );

          const alias =
            quoteIdentifier(
              `${column.column_name}__sum`
            );

          return `
    SUM(${q}) AS ${alias},
    AVG(${q}) AS ${quoteIdentifier(
      `${column.column_name}__avg`
    )},
    MIN(${q}) AS ${quoteIdentifier(
      `${column.column_name}__min`
    )},
    MAX(${q}) AS ${quoteIdentifier(
      `${column.column_name}__max`
    )}`;
        }
      );

    return `
SELECT
${expressions.join(",\n")}
FROM excel_data;
`;
  }

  /* =========================================================
   * EJECUTAR SQL
   * ======================================================= */

  async function executeSQL() {

    const textarea =
      document.getElementById(
        "tmxe-v0492-sql"
      );

    if (!textarea) {
      return;
    }

    const sql =
      textarea.value.trim();

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
        mensaje:
          "DuckDB no está conectado o no existe una tabla excel_data."
      });

      return;
    }

    try {

      setStatus(
        "⏳ Ejecutando consulta SQL local..."
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

      const rawRows =
        result.toArray();

      const rows =
        normalizeRows(
          rawRows
        );

      setStatus(
        `✅ Consulta ejecutada localmente en ${elapsed.toFixed(1)} ms.`,
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
            elapsed.toFixed(1)
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
   * NORMALIZAR FILAS
   * ======================================================= */

  function normalizeRows(
    rows
  ) {

    return rows.map(
      row =>
        normalizeForJSON(
          row
        )
    );
  }

  /* =========================================================
   * VERSIÓN DUCKDB
   * ======================================================= */

  async function getDuckDBVersion() {

    try {

      const result =
        await conn.query(
          "SELECT version() AS duckdb_version;"
        );

      return normalizeRows(
        result.toArray()
      );

    } catch (error) {

      return {
        error:
          error?.message ||
          String(error)
      };
    }
  }

  /* =========================================================
   * INFO UI
   * ======================================================= */

  function updateFileInfo(
    fileName,
    rowCount
  ) {

    const fileElement =
      document.getElementById(
        "tmxe-v0492-file"
      );

    const nameElement =
      document.getElementById(
        "tmxe-v0492-file-name"
      );

    const countElement =
      document.getElementById(
        "tmxe-v0492-row-count"
      );

    if (fileElement) {

      fileElement.textContent =
        fileName;
    }

    if (nameElement) {

      nameElement.textContent =
        fileName;
    }

    if (countElement) {

      countElement.textContent =
        String(rowCount);
    }
  }

  /* =========================================================
   * INICIALIZACIÓN
   * ======================================================= */

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
