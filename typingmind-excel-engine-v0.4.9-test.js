/*
 * TypingMind Excel Data Engine - Extension v0.4.9 TEST
 *
 * RESULT NORMALIZER
 *
 * Objetivo:
 * - Mantener carga local de Excel.
 * - Mantener DuckDB-Wasm.
 * - Mantener consultas SQL locales.
 * - Convertir resultados DuckDB a JSON limpio.
 * - Resolver BigInt / TypedArray / Date / valores DuckDB.
 *
 * Procesamiento:
 * EXCEL -> NAVEGADOR -> DUCKDB-WASM -> SQL -> RESULTADO NORMALIZADO
 *
 * NO se envía el Excel al servidor.
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
  let currentTable = "excel_data";

  /* =========================================================
     NORMALIZADOR DE RESULTADOS DUCKDB
     ========================================================= */

  function normalizeDuckDBValue(value) {

    if (value === null || value === undefined) {
      return null;
    }

    /*
     * BigInt
     */
    if (typeof value === "bigint") {
      const number = Number(value);

      if (
        Number.isSafeInteger(number) &&
        BigInt(number) === value
      ) {
        return number;
      }

      return value.toString();
    }

    /*
     * Date
     */
    if (value instanceof Date) {
      return value.toISOString();
    }

    /*
     * TypedArray / ArrayBuffer
     *
     * DuckDB puede devolver algunos resultados
     * numéricos como estructuras binarias.
     */
    if (
      ArrayBuffer.isView(value) &&
      !(value instanceof DataView)
    ) {

      if (value.length === 1) {
        return normalizeDuckDBValue(value[0]);
      }

      return Array.from(value).map(
        normalizeDuckDBValue
      );
    }

    /*
     * Array
     */
    if (Array.isArray(value)) {
      return value.map(
        normalizeDuckDBValue
      );
    }

    /*
     * Objeto
     */
    if (
      typeof value === "object"
    ) {

      const result = {};

      for (const key of Object.keys(value)) {

        result[key] =
          normalizeDuckDBValue(
            value[key]
          );
      }

      return result;
    }

    /*
     * Número normal
     */
    if (typeof value === "number") {

      if (Number.isNaN(value)) {
        return null;
      }

      if (!Number.isFinite(value)) {
        return null;
      }

      return value;
    }

    /*
     * String / boolean
     */
    return value;
  }


  function normalizeRows(rows) {

    if (!rows) {
      return [];
    }

    return rows.map(row => {

      const normalized = {};

      for (const key of Object.keys(row)) {

        normalized[key] =
          normalizeDuckDBValue(
            row[key]
          );
      }

      return normalized;
    });
  }


  /*
   * JSON.stringify seguro.
   *
   * Aunque quede algún BigInt inesperado,
   * nunca debe romper la interfaz.
   */
  function safeStringify(value) {

    return JSON.stringify(
      value,
      (key, currentValue) => {

        if (
          typeof currentValue === "bigint"
        ) {

          const number =
            Number(currentValue);

          if (
            Number.isSafeInteger(number) &&
            BigInt(number) === currentValue
          ) {
            return number;
          }

          return currentValue.toString();
        }

        if (
          ArrayBuffer.isView(currentValue) &&
          !(currentValue instanceof DataView)
        ) {
          return Array.from(currentValue);
        }

        return currentValue;
      },
      2
    );
  }


  /* =========================================================
     ESTILOS
     ========================================================= */

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

      #tmxe-v049-overlay {

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

      #tmxe-v049-panel {

        width:
          min(1000px, 96vw);

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

      #tmxe-v049-panel h2 {

        margin:
          0 0 5px;

        font-size: 20px;
      }

      #tmxe-v049-help {

        opacity: .7;

        margin-bottom: 15px;
      }

      #tmxe-v049-status {

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

      #tmxe-v049-status[data-kind="ok"] {
        color: #16a34a;
      }

      #tmxe-v049-status[data-kind="error"] {
        color: #dc2626;
      }

      #tmxe-v049-result {

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

      #tmxe-v049-actions {

        display: flex;

        gap: 8px;

        flex-wrap: wrap;

        margin: 10px 0;
      }

      #tmxe-v049-run,
      #tmxe-v049-close {

        border:
          1px solid
          rgba(127,127,127,.45);

        border-radius: 9px;

        padding: 9px 13px;

        background: transparent;

        color: inherit;

        cursor: pointer;
      }

      #tmxe-v049-run {

        background:
          #7c3aed !important;

        color:
          white !important;

        border-color:
          #7c3aed !important;
      }

      #tmxe-v049-file {

        margin:
          8px 0 12px;

        width: 100%;
      }

      #tmxe-v049-sql {

        width: 100%;

        min-height: 110px;

        box-sizing: border-box;

        padding: 10px;

        border-radius: 9px;

        border:
          1px solid
          rgba(127,127,127,.35);

        background: transparent;

        color: inherit;

        font:
          13px/1.5
          ui-monospace,
          SFMono-Regular,
          Menlo,
          monospace;
      }

    `;

    document.head.appendChild(style);
  }


  /* =========================================================
     BOTÓN
     ========================================================= */

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

    button.type =
      "button";

    button.textContent =
      "📊 Excel v0.4.9";

    button.title =
      "Excel Data Engine v0.4.9";

    button.addEventListener(
      "click",
      openPanel
    );

    document.body.appendChild(
      button
    );
  }


  /* =========================================================
     PANEL
     ========================================================= */

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

              Procesamiento local con
              DuckDB-Wasm.
              Esta versión agrega
              normalización de resultados.

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
          accept=".xlsx,.xls,.csv"
        />

        <div id="tmxe-v049-actions">

          <button
            id="tmxe-v049-run"
            type="button"
          >
            📂 Cargar Excel
          </button>

        </div>

        <div id="tmxe-v049-status">
          Listo.
        </div>

        <strong>
          Consulta SQL
        </strong>

        <textarea
          id="tmxe-v049-sql"
        >SELECT COUNT(*) AS registros
FROM excel_data;</textarea>

        <div id="tmxe-v049-actions">

          <button
            id="tmxe-v049-query"
            type="button"
          >
            ▶ Ejecutar SQL
          </button>

        </div>

        <strong>
          Resultado
        </strong>

        <pre id="tmxe-v049-result">—</pre>

        <div style="
          margin-top:12px;
          opacity:.65;
          font-size:12px;
        ">

          v0.4.9 TEST —
          Result Normalizer.
          Excel y consultas procesados
          localmente.

        </div>

      </div>
    `;

    document.body.appendChild(
      overlay
    );

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

    document
      .getElementById(
        "tmxe-v049-query"
      )
      .addEventListener(
        "click",
        runSQL
      );
  }


  /* =========================================================
     UI
     ========================================================= */

  function setStatus(
    text,
    kind = ""
  ) {

    const el =
      document.getElementById(
        "tmxe-v049-status"
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


  function setResult(
    value
  ) {

    const el =
      document.getElementById(
        "tmxe-v049-result"
      );

    if (!el) {
      return;
    }

    try {

      const normalized =
        normalizeDuckDBValue(
          value
        );

      el.textContent =
        safeStringify(
          normalized
        );

    } catch (error) {

      el.textContent =
        String(error);
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
      "INSTALL icu;"
    ).catch(() => {});

    await conn.query(
      "LOAD icu;"
    ).catch(() => {});

    setStatus(
      "4/5 — DuckDB conectado."
    );
  }


  /* =========================================================
     LECTURA XLSX
     ========================================================= */

  async function loadExcel() {

    const input =
      document.getElementById(
        "tmxe-v049-file"
      );

    if (
      !input ||
      !input.files ||
      !input.files.length
    ) {

      setStatus(
        "Selecciona primero un archivo Excel.",
        "error"
      );

      return;
    }

    try {

      const file =
        input.files[0];

      currentFile =
        file;

      await initializeDuckDB();

      setStatus(
        "5/5 — Leyendo archivo Excel localmente..."
      );

      const arrayBuffer =
        await file.arrayBuffer();

      /*
       * XLSX se procesa mediante SheetJS.
       *
       * La biblioteca se carga únicamente en el navegador.
       */
      if (
        !window.XLSX
      ) {

        await loadScript(
          "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"
        );
      }

      const workbook =
        window.XLSX.read(
          arrayBuffer,
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

      const rows =
        window.XLSX.utils.sheet_to_json(
          sheet,
          {
            header: 1,
            defval: null,
            raw: true
          }
        );

      /*
       * Buscar primera fila con contenido.
       */
      let headerIndex = -1;

      for (
        let i = 0;
        i < rows.length;
        i++
      ) {

        const row =
          rows[i];

        if (
          row &&
          row.some(
            value =>
              value !== null &&
              value !== undefined &&
              String(value).trim() !== ""
          )
        ) {

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

      const headers =
        makeUniqueHeaders(
          rows[headerIndex]
        );

      const dataRows =
        rows
          .slice(headerIndex + 1)
          .filter(
            row =>
              row &&
              row.some(
                value =>
                  value !== null &&
                  value !== undefined &&
                  String(value).trim() !== ""
              )
          );

      /*
       * Convertir datos a objetos.
       */
      const objects =
        dataRows.map(
          row => {

            const obj = {};

            for (
              let i = 0;
              i < headers.length;
              i++
            ) {

              obj[headers[i]] =
                row[i] ?? null;
            }

            return obj;
          }
        );

      /*
       * Eliminar tabla anterior.
       */
      await conn.query(
        `DROP TABLE IF EXISTS "${currentTable}";`
      );

      /*
       * Crear tabla inicialmente como VARCHAR.
       *
       * Posteriormente inferimos tipos.
       */
      const columnsSQL =
        headers
          .map(
            name =>
              `"${escapeIdentifier(name)}" VARCHAR`
          )
          .join(", ");

      await conn.query(
        `CREATE TABLE "${currentTable}" (${columnsSQL});`
      );

      /*
       * Insertar filas.
       */
      for (
        const row of objects
      ) {

        const columnNames =
          headers
            .map(
              h =>
                `"${escapeIdentifier(h)}"`
            )
            .join(", ");

        const values =
          headers
            .map(
              h =>
                sqlLiteral(
                  row[h]
                )
            )
            .join(", ");

        await conn.query(
          `INSERT INTO "${currentTable}"
           (${columnNames})
           VALUES (${values});`
        );
      }

      /*
       * Intentar inferencia de tipos.
       */
      await inferColumnTypes(
        headers,
        objects
      );

      /*
       * Resultado inicial.
       */
      const countResult =
        await conn.query(
          `SELECT COUNT(*) AS registros
           FROM "${currentTable}";`
        );

      const rowsResult =
        normalizeRows(
          countResult.toArray()
        );

      setStatus(
        "✅ Excel cargado y procesado localmente.",
        "ok"
      );

      setResult({

        procesamiento:
          "LOCAL",

        archivo:
          file.name,

        tamano_bytes:
          file.size,

        hojas:
          sheets,

        hoja_principal:
          sheetName,

        filas_fisicas_detectadas:
          rows.length,

        filas_reales:
          objects.length,

        filas_insertadas:
          objects.length,

        columnas_detectadas:
          headers.length,

        encabezados:
          headers,

        duckdb_table:
          currentTable,

        count:
          rowsResult

      });

    } catch (error) {

      console.error(
        `[${APP_ID}]`,
        error
      );

      setStatus(
        "❌ Error al cargar Excel.",
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

  async function inferColumnTypes(
    headers,
    objects
  ) {

    /*
     * Para esta versión:
     *
     * - fechas -> DATE
     * - enteros -> BIGINT
     * - decimales -> DOUBLE
     * - resto -> VARCHAR
     *
     * Solo convertimos cuando todos los
     * valores no vacíos son compatibles.
     */

    for (
      let i = 0;
      i < headers.length;
      i++
    ) {

      const header =
        headers[i];

      const values =
        objects
          .map(
            row =>
              row[header]
          )
          .filter(
            value =>
              value !== null &&
              value !== undefined &&
              String(value).trim() !== ""
          );

      if (!values.length) {
        continue;
      }

      /*
       * Detectar fecha Excel.
       */
      if (
        isDateColumn(
          header,
          values
        )
      ) {

        await conn.query(
          `ALTER TABLE "${currentTable}"
           ALTER COLUMN "${escapeIdentifier(header)}"
           SET DATA TYPE DATE
           USING try_cast("${escapeIdentifier(header)}" AS DATE);`
        ).catch(
          () => {}
        );

        continue;
      }

      /*
       * Enteros.
       */
      if (
        values.every(
          isIntegerValue
        )
      ) {

        await conn.query(
          `ALTER TABLE "${currentTable}"
           ALTER COLUMN "${escapeIdentifier(header)}"
           SET DATA TYPE BIGINT
           USING try_cast("${escapeIdentifier(header)}" AS BIGINT);`
        ).catch(
          () => {}
        );

        continue;
      }

      /*
       * Decimales.
       */
      if (
        values.every(
          isNumericValue
        )
      ) {

        await conn.query(
          `ALTER TABLE "${currentTable}"
           ALTER COLUMN "${escapeIdentifier(header)}"
           SET DATA TYPE DOUBLE
           USING try_cast("${escapeIdentifier(header)}" AS DOUBLE);`
        ).catch(
          () => {}
        );
      }
    }
  }


  function isIntegerValue(
    value
  ) {

    if (
      typeof value === "number"
    ) {

      return Number.isInteger(
        value
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
      Number.isFinite(
        Number(text)
      )
    );
  }


  function isDateColumn(
    header,
    values
  ) {

    const name =
      String(header)
        .toLowerCase();

    const headerSuggestsDate =
      /fecha|date|día|dia|time/.test(
        name
      );

    if (
      !headerSuggestsDate
    ) {
      return false;
    }

    return values.every(
      value => {

        if (
          typeof value === "number"
        ) {

          return (
            value > 20000 &&
            value < 100000
          );
        }

        return false;
      }
    );
  }


  /* =========================================================
     SQL
     ========================================================= */

  async function runSQL() {

    if (!conn) {

      setStatus(
        "Primero debes cargar un archivo Excel.",
        "error"
      );

      return;
    }

    const sqlElement =
      document.getElementById(
        "tmxe-v049-sql"
      );

    const sql =
      sqlElement.value.trim();

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

      const result =
        await conn.query(
          sql
        );

      const rawRows =
        result.toArray();

      const normalizedRows =
        normalizeRows(
          rawRows
        );

      setStatus(
        "✅ Consulta ejecutada localmente.",
        "ok"
      );

      setResult({

        procesamiento:
          "LOCAL",

        sql:
          sql,

        filas_resultado:
          normalizedRows.length,

        resultado:
          normalizedRows

      });

    } catch (error) {

      console.error(
        `[${APP_ID}] SQL`,
        error
      );

      setStatus(
        "❌ Error en consulta SQL.",
        "error"
      );

      setResult({

        procesamiento:
          "LOCAL",

        sql:
          sql,

        error:
          error?.message ||
          String(error)

      });
    }
  }


  /* =========================================================
     UTILIDADES
     ========================================================= */

  function escapeIdentifier(
    value
  ) {

    return String(value)
      .replaceAll(
        '"',
        '""'
      );
  }


  function sqlLiteral(
    value
  ) {

    if (
      value === null ||
      value === undefined ||
      value === ""
    ) {

      return "NULL";
    }

    /*
     * Excel serial date.
     *
     * No convertimos aquí:
     * la inferencia posterior se encarga
     * del tipo DATE.
     */

    if (
      typeof value === "number"
    ) {

      if (
        Number.isFinite(value)
      ) {

        return String(value);
      }

      return "NULL";
    }

    const text =
      String(value)
        .replaceAll(
          "'",
          "''"
        );

    return `'${text}'`;
  }


  function makeUniqueHeaders(
    row
  ) {

    const result = [];

    const used =
      new Map();

    for (
      let i = 0;
      i < row.length;
      i++
    ) {

      let name =
        row[i] === null ||
        row[i] === undefined ||
        String(row[i]).trim() === ""
          ? `column_${i + 1}`
          : String(row[i]).trim();

      const count =
        used.get(name) || 0;

      if (count > 0) {

        name =
          `${name}_${count + 1}`;
      }

      used.set(
        String(row[i]),
        count + 1
      );

      result.push(
        name
      );
    }

    return result;
  }


  async function loadScript(
    src
  ) {

    await new Promise(
      (resolve, reject) => {

        const script =
          document.createElement(
            "script"
          );

        script.src =
          src;

        script.onload =
          resolve;

        script.onerror =
          () =>
            reject(
              new Error(
                `No se pudo cargar: ${src}`
              )
            );

        document.head.appendChild(
          script
        );
      }
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
