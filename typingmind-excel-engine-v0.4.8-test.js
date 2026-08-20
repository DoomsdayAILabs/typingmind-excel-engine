```javascript
/*
 * TypingMind Excel Data Engine - Extension v0.4.8 TEST
 *
 * OBJETIVO:
 * - Cargar un XLSX localmente.
 * - Procesarlo completamente dentro del navegador.
 * - Crear tabla DuckDB: excel_data.
 * - Inferir tipos básicos.
 * - Ejecutar SQL manual contra excel_data.
 * - Mostrar resultados tabulares.
 *
 * IMPORTANTE:
 * - El archivo NO se sube a ningún servidor.
 * - No utiliza la extensión oficial "excel" de DuckDB.
 * - XLSX se interpreta localmente mediante SheetJS.
 * - DuckDB-Wasm se utiliza para consultas SQL locales.
 *
 * v0.4.8 TEST
 */

(() => {
  "use strict";

  const APP_ID = "tm-excel-engine-v048-test";

  /*
   * DuckDB-Wasm estable que ya funcionó en v0.4.7.
   */
  const DUCKDB_PACKAGE =
    "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

  /*
   * SheetJS para lectura local de XLSX.
   *
   * Solo se descarga la biblioteca.
   * El archivo Excel permanece en el navegador.
   */
  const XLSX_PACKAGE =
    "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";

  let duckdb = null;
  let db = null;
  let conn = null;
  let XLSX = null;

  let currentFile = null;
  let currentSheet = null;
  let currentRows = [];
  let currentHeaders = [];
  let currentSchema = [];

  function addStyles() {
    if (document.getElementById("tmxe-v048-style")) return;

    const style = document.createElement("style");

    style.id = "tmxe-v048-style";

    style.textContent = `
      #tmxe-v048-button {
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

      #tmxe-v048-overlay {
        position: fixed;
        inset: 0;

        z-index: 2147483001;

        background: rgba(0,0,0,.48);

        display: flex;
        align-items: center;
        justify-content: center;

        padding: 16px;
      }

      #tmxe-v048-panel {
        width: min(1100px, 96vw);
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

      #tmxe-v048-panel h2 {
        margin: 0 0 5px;
        font-size: 20px;
      }

      #tmxe-v048-help {
        opacity: .7;
        margin-bottom: 15px;
      }

      #tmxe-v048-status {
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

      #tmxe-v048-status[data-kind="ok"] {
        color: #16a34a;
      }

      #tmxe-v048-status[data-kind="error"] {
        color: #dc2626;
      }

      #tmxe-v048-actions {
        display: flex;
        gap: 8px;
        flex-wrap: wrap;
        margin: 10px 0;
      }

      #tmxe-v048-actions button {
        border: 1px solid rgba(127,127,127,.45);

        border-radius: 9px;

        padding: 9px 13px;

        background: transparent;

        color: inherit;

        cursor: pointer;
      }

      #tmxe-v048-load {
        background: #2563eb !important;
        color: white !important;
        border-color: #2563eb !important;
      }

      #tmxe-v048-sql {
        width: 100%;
        min-height: 130px;

        box-sizing: border-box;

        resize: vertical;

        padding: 12px;

        border-radius: 10px;

        border: 1px solid rgba(127,127,127,.4);

        background: Canvas;
        color: CanvasText;

        font: 13px/1.5 ui-monospace,
              SFMono-Regular,
              Menlo,
              monospace;
      }

      #tmxe-v048-result {
        margin-top: 8px;

        overflow: auto;

        max-height: 450px;

        border: 1px solid rgba(127,127,127,.25);

        border-radius: 10px;
      }

      #tmxe-v048-result table {
        border-collapse: collapse;
        width: 100%;
        min-width: 500px;
      }

      #tmxe-v048-result th,
      #tmxe-v048-result td {
        padding: 8px 10px;

        border-bottom: 1px solid rgba(127,127,127,.2);

        text-align: left;

        white-space: nowrap;
      }

      #tmxe-v048-result th {
        position: sticky;
        top: 0;

        background: Canvas;

        font-weight: 700;
      }

      #tmxe-v048-result td {
        font-family:
          ui-monospace,
          SFMono-Regular,
          Menlo,
          monospace;

        font-size: 12px;
      }

      #tmxe-v048-info {
        margin: 12px 0;

        padding: 12px;

        border-radius: 10px;

        background: rgba(127,127,127,.08);

        font: 12px/1.5 ui-monospace,
              SFMono-Regular,
              Menlo,
              monospace;

        white-space: pre-wrap;
      }

      #tmxe-v048-examples {
        display: flex;
        gap: 6px;
        flex-wrap: wrap;
        margin: 8px 0 12px;
      }

      #tmxe-v048-examples button {
        border: 1px solid rgba(127,127,127,.35);
        border-radius: 8px;
        padding: 6px 9px;
        background: transparent;
        color: inherit;
        cursor: pointer;
        font-size: 12px;
      }
    `;

    document.head.appendChild(style);
  }

  function createButton() {
    if (document.getElementById("tmxe-v048-button")) return;

    const button = document.createElement("button");

    button.id = "tmxe-v048-button";
    button.type = "button";
    button.textContent = "📊 Excel v0.4.8";
    button.title = "Excel Data Engine v0.4.8";

    button.addEventListener("click", openPanel);

    document.body.appendChild(button);
  }

  function openPanel() {
    if (document.getElementById("tmxe-v048-overlay")) return;

    const overlay = document.createElement("div");

    overlay.id = "tmxe-v048-overlay";

    overlay.innerHTML = `
      <div
        id="tmxe-v048-panel"
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
            <h2>📊 Excel Engine v0.4.8 TEST</h2>

            <div id="tmxe-v048-help">
              Excel local + DuckDB-Wasm + SQL local.
            </div>
          </div>

          <button
            id="tmxe-v048-close"
            type="button"
          >
            Cerrar
          </button>

        </div>

        <div id="tmxe-v048-status">
          Listo. Selecciona un archivo Excel.
        </div>

        <div id="tmxe-v048-actions">

          <button
            id="tmxe-v048-load"
            type="button"
          >
            📂 Cargar Excel
          </button>

          <button
            id="tmxe-v048-test"
            type="button"
          >
            ⚡ Ejecutar consulta
          </button>

        </div>

        <input
          id="tmxe-v048-file"
          type="file"
          accept=".xlsx,.xls,.xlsb"
          style="display:none"
        />

        <div id="tmxe-v048-info">
          No hay archivo cargado.
        </div>

        <strong>Consulta SQL</strong>

        <textarea
          id="tmxe-v048-sql"
          spellcheck="false"
        >SELECT *
FROM excel_data
LIMIT 10;</textarea>

        <div id="tmxe-v048-examples">

          <button data-sql="
SELECT COUNT(*) AS registros
FROM excel_data;
          ">
            COUNT
          </button>

          <button data-sql="
SELECT *
FROM excel_data
LIMIT 10;
          ">
            Primeros 10
          </button>

          <button data-sql="
SELECT *
FROM excel_data
ORDER BY 1
LIMIT 10;
          ">
            Ordenar
          </button>

          <button data-sql="
SELECT *
FROM excel_data
WHERE 1 = 1
LIMIT 10;
          ">
            WHERE
          </button>

          <button data-sql="
SELECT
  EXTRACT(YEAR FROM &quot;Fecha&quot;) AS año,
  SUM(&quot;Total TEU's&quot;) AS total
FROM excel_data
GROUP BY año
ORDER BY año;
          ">
            Agrupar año
          </button>

        </div>

        <strong>Resultado</strong>

        <div id="tmxe-v048-result">
          <div style="padding:15px;opacity:.65;">
            Ejecuta una consulta.
          </div>
        </div>

        <div style="
          margin-top:12px;
          opacity:.65;
          font-size:12px;
        ">
          v0.4.8 TEST — El Excel permanece local.
          DuckDB ejecuta las consultas dentro del navegador.
        </div>

      </div>
    `;

    document.body.appendChild(overlay);

    document
      .getElementById("tmxe-v048-close")
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
      .getElementById("tmxe-v048-load")
      .addEventListener(
        "click",
        () => {
          document
            .getElementById("tmxe-v048-file")
            .click();
        }
      );

    document
      .getElementById("tmxe-v048-file")
      .addEventListener(
        "change",
        loadExcel
      );

    document
      .getElementById("tmxe-v048-test")
      .addEventListener(
        "click",
        executeSQL
      );

    document
      .querySelectorAll(
        "#tmxe-v048-examples button"
      )
      .forEach(button => {

        button.addEventListener(
          "click",
          () => {

            const sql =
              button.dataset.sql.trim();

            document
              .getElementById(
                "tmxe-v048-sql"
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

    const el =
      document.getElementById(
        "tmxe-v048-status"
      );

    if (!el) return;

    el.textContent = text;

    if (kind) {
      el.dataset.kind = kind;
    } else {
      delete el.dataset.kind;
    }
  }

  function setInfo(text) {

    const el =
      document.getElementById(
        "tmxe-v048-info"
      );

    if (!el) return;

    el.textContent = text;
  }

  function setResultMessage(
    text,
    kind = ""
  ) {

    const el =
      document.getElementById(
        "tmxe-v048-result"
      );

    if (!el) return;

    el.innerHTML = "";

    const div =
      document.createElement("div");

    div.style.padding = "15px";
    div.style.whiteSpace = "pre-wrap";

    if (kind === "error") {
      div.style.color = "#dc2626";
    }

    div.textContent = text;

    el.appendChild(div);
  }

  /*
   * Convierte valores DuckDB a algo
   * seguro para mostrar.
   *
   * Evita el problema:
   * "Do not know how to serialize a BigInt"
   */
  function safeValue(value) {

    if (typeof value === "bigint") {
      return value.toString();
    }

    if (value instanceof Date) {
      return value.toISOString();
    }

    if (
      value &&
      typeof value === "object"
    ) {

      try {
        return JSON.stringify(
          value,
          (_, v) =>
            typeof v === "bigint"
              ? v.toString()
              : v
        );

      } catch {
        return String(value);
      }
    }

    return value;
  }

  function rowsToSafeObjects(
    rows
  ) {

    return rows.map(row => {

      const output = {};

      for (const key of Object.keys(row)) {
        output[key] =
          safeValue(row[key]);
      }

      return output;
    });
  }

  function escapeHTML(value) {

    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function renderTable(
    rows
  ) {

    const container =
      document.getElementById(
        "tmxe-v048-result"
      );

    if (!container) return;

    container.innerHTML = "";

    if (!rows.length) {

      const div =
        document.createElement("div");

      div.style.padding = "15px";
      div.textContent =
        "La consulta no devolvió registros.";

      container.appendChild(div);

      return;
    }

    const columns =
      Object.keys(rows[0]);

    const table =
      document.createElement("table");

    const thead =
      document.createElement("thead");

    const headerRow =
      document.createElement("tr");

    columns.forEach(column => {

      const th =
        document.createElement("th");

      th.textContent = column;

      headerRow.appendChild(th);

    });

    thead.appendChild(headerRow);

    const tbody =
      document.createElement("tbody");

    rows.forEach(row => {

      const tr =
        document.createElement("tr");

      columns.forEach(column => {

        const td =
          document.createElement("td");

        td.textContent =
          safeValue(row[column]);

        tr.appendChild(td);

      });

      tbody.appendChild(tr);

    });

    table.appendChild(thead);
    table.appendChild(tbody);

    container.appendChild(table);
  }

  async function initializeDuckDB() {

    if (conn) return;

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

    return bundle;
  }

  async function initializeXLSX() {

    if (XLSX) return;

    setStatus(
      "Cargando lector XLSX..."
    );

    XLSX =
      await import(
        XLSX_PACKAGE
      );
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
      Number(serial) *
      86400000
    );
  }

  function looksLikeDateHeader(
    name
  ) {

    const normalized =
      String(name)
        .toLowerCase()
        .normalize("NFD")
        .replace(
          /[\u0300-\u036f]/g,
          ""
        );

    return (
      normalized.includes("fecha") ||
      normalized.includes("date") ||
      normalized.includes("dia") ||
      normalized.includes("day")
    );
  }

  function looksLikeDateValue(
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
      typeof value === "number" &&
      value > 20000 &&
      value < 80000
    ) {
      return true;
    }

    const text =
      String(value).trim();

    if (
      /^\d{4}-\d{1,2}-\d{1,2}$/.test(
        text
      )
    ) {
      return true;
    }

    if (
      /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/.test(
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
          value !== null &&
          value !== undefined &&
          String(value).trim() !== ""
      );

    if (!nonEmpty.length) {
      return {
        duckdb_type: "VARCHAR",
        confidence: "low",
        reason: "columna vacía"
      };
    }

    if (
      looksLikeDateHeader(header) &&
      nonEmpty.every(
        looksLikeDateValue
      )
    ) {
      return {
        duckdb_type: "DATE",
        confidence: "high",
        reason:
          "encabezado y valores compatibles con fecha"
      };
    }

    const allIntegers =
      nonEmpty.every(value => {

        if (
          typeof value === "number"
        ) {
          return (
            Number.isFinite(value) &&
            Number.isInteger(value)
          );
        }

        return /^[-+]?\d+$/.test(
          String(value).trim()
        );
      });

    if (allIntegers) {
      return {
        duckdb_type: "BIGINT",
        confidence: "high",
        reason:
          "todos los valores son enteros"
      };
    }

    const allNumbers =
      nonEmpty.every(value => {

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
      });

    if (allNumbers) {
      return {
        duckdb_type: "DOUBLE",
        confidence: "high",
        reason:
          "todos los valores son numéricos"
      };
    }

    return {
      duckdb_type: "VARCHAR",
      confidence: "medium",
      reason:
        "valores mixtos o texto"
    };
  }

  function convertValue(
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

    if (type === "BIGINT") {

      try {

        return BigInt(
          String(value)
            .trim()
            .replace(/,/g, "")
        );

      } catch {

        return null;

      }
    }

    if (type === "DOUBLE") {

      const number =
        Number(
          String(value)
            .trim()
            .replace(/,/g, "")
        );

      return Number.isFinite(number)
        ? number
        : null;
    }

    if (type === "DATE") {

      if (
        typeof value === "number"
      ) {

        const date =
          excelSerialToDate(value);

        return date
          .toISOString()
          .slice(0, 10);

      }

      const text =
        String(value).trim();

      const parsed =
        new Date(text);

      if (!Number.isNaN(
        parsed.getTime()
      )) {

        return parsed
          .toISOString()
          .slice(0, 10);

      }

      return null;
    }

    return String(value);
  }

  function quoteIdentifier(
    name
  ) {

    return `"${String(name)
      .replaceAll('"', '""')}"`;
  }

  function makeUniqueHeaders(
    headers
  ) {

    const used =
      new Map();

    return headers.map(
      (header, index) => {

        let base =
          String(
            header ??
            ""
          ).trim();

        if (!base) {
          base =
            `Column_${index + 1}`;
        }

        const count =
          used.get(base) || 0;

        used.set(
          base,
          count + 1
        );

        if (count === 0) {
          return base;
        }

        return `${base}_${count + 1}`;
      }
    );
  }

  async function createDuckDBTable(
    rows,
    headers,
    schema
  ) {

    setStatus(
      "Creando tabla DuckDB local..."
    );

    try {
      await conn.query(
        "DROP TABLE IF EXISTS excel_data;"
      );
    } catch {}

    const columnDefinitions =
      headers.map(
        (header, index) => {

          return (
            quoteIdentifier(header) +
            " " +
            schema[index].duckdb_type
          );

        }
      );

    await conn.query(
      `
      CREATE TABLE excel_data (
        ${columnDefinitions.join(",\n")}
      );
      `
    );

    /*
     * Insertamos por lotes.
     *
     * Esto evita construir una consulta gigante
     * para archivos grandes.
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

      const valuesSQL =
        batch.map(row => {

          const values =
            headers.map(
              (header, index) => {

                const value =
                  convertValue(
                    row[index],
                    schema[index]
                      .duckdb_type
                  );

                if (
                  value === null
                ) {
                  return "NULL";
                }

                if (
                  schema[index]
                    .duckdb_type ===
                  "BIGINT"
                ) {
                  return String(value);
                }

                if (
                  schema[index]
                    .duckdb_type ===
                  "DOUBLE"
                ) {
                  return String(value);
                }

                return `'${String(value)
                  .replaceAll(
                    "'",
                    "''"
                  )}'`;
              }
            );

          return `(${values.join(",")})`;

        }).join(",");

      if (valuesSQL) {

        await conn.query(
          `
          INSERT INTO excel_data
          VALUES ${valuesSQL};
          `
        );

      }

      setStatus(
        `Cargando datos localmente...\n` +
        `${Math.min(
          start + BATCH_SIZE,
          rows.length
        )} / ${rows.length} registros`
      );
    }
  }

  async function loadExcel(
    event
  ) {

    const file =
      event.target.files?.[0];

    if (!file) return;

    try {

      setResultMessage(
        "Procesando archivo local..."
      );

      currentFile = file;

      setStatus(
        "1/6 — Cargando bibliotecas..."
      );

      await initializeDuckDB();

      await initializeXLSX();

      setStatus(
        "2/6 — Leyendo Excel localmente..."
      );

      const buffer =
        await file.arrayBuffer();

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

      currentSheet =
        sheets[0];

      const worksheet =
        workbook.Sheets[
          currentSheet
        ];

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

      if (!matrix.length) {
        throw new Error(
          "La hoja está vacía."
        );
      }

      setStatus(
        "3/6 — Detectando encabezados y filas..."
      );

      /*
       * Primera fila = encabezados.
       */
      const rawHeaders =
        matrix[0];

      currentHeaders =
        makeUniqueHeaders(
          rawHeaders
        );

      const dataRows =
        matrix
          .slice(1)
          .filter(row =>
            row.some(
              value =>
                value !== null &&
                value !== undefined &&
                String(value).trim() !== ""
            )
          );

      currentRows =
        dataRows;

      /*
       * Inferencia usando hasta 100 filas.
       */
      const sampleRows =
        dataRows.slice(
          0,
          Math.min(
            100,
            dataRows.length
          )
        );

      currentSchema =
        currentHeaders.map(
          (header, index) => {

            const values =
              sampleRows.map(
                row => row[index]
              );

            const inferred =
              inferColumnType(
                header,
                values
              );

            return {
              column_index: index,
              column_name: header,
              ...inferred,
              non_empty_values:
                values.filter(
                  value =>
                    value !== null &&
                    value !== undefined &&
                    String(value).trim() !== ""
                ).length
            };

          }
        );

      setStatus(
        "4/6 — Insertando datos en DuckDB local..."
      );

      await createDuckDBTable(
        currentRows,
        currentHeaders,
        currentSchema
      );

      setStatus(
        "5/6 — Verificando tabla..."
      );

      const countResult =
        await conn.query(
          `
          SELECT COUNT(*) AS registros
          FROM excel_data;
          `
        );

      const countRows =
        rowsToSafeObjects(
          countResult.toArray()
        );

      const schemaResult =
        await conn.query(
          `
          DESCRIBE excel_data;
          `
        );

      const duckSchema =
        rowsToSafeObjects(
          schemaResult.toArray()
        );

      setInfo(
        [
          `Archivo: ${file.name}`,
          `Tamaño: ${file.size.toLocaleString()} bytes`,
          `Hoja: ${currentSheet}`,
          `Filas físicas: ${matrix.length - 1}`,
          `Filas reales: ${dataRows.length}`,
          `Columnas: ${currentHeaders.length}`,
          ``,
          `Tabla DuckDB: excel_data`,
          `Registros insertados: ${
            countRows[0]?.registros ?? "?"
          }`
        ].join("\n")
      );

      setStatus(
        "6/6 — Excel cargado correctamente.\n" +
        "Procesamiento 100% LOCAL.",
        "ok"
      );

      /*
       * Consulta inicial automática.
       */
      document
        .getElementById(
          "tmxe-v048-sql"
        )
        .value =
        `
SELECT *
FROM excel_data
LIMIT 10;
        `.trim();

      await executeSQL();

    } catch (error) {

      console.error(
        `[${APP_ID}]`,
        error
      );

      setStatus(
        "❌ ERROR AL CARGAR EXCEL",
        "error"
      );

      setResultMessage(
        error?.message ||
        String(error),
        "error"
      );

    } finally {

      event.target.value = "";

    }
  }

  async function executeSQL() {

    if (!conn) {

      setResultMessage(
        "Primero debes cargar un archivo Excel.",
        "error"
      );

      return;
    }

    const textarea =
      document.getElementById(
        "tmxe-v048-sql"
      );

    const sql =
      textarea?.value?.trim();

    if (!sql) {

      setResultMessage(
        "La consulta SQL está vacía.",
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

      const elapsed =
        performance.now() -
        start;

      const rows =
        result.toArray();

      const safeRows =
        rowsToSafeObjects(
          rows
        );

      renderTable(
        safeRows
      );

      setStatus(
        [
          "Consulta ejecutada correctamente.",
          `Filas devueltas: ${safeRows.length}`,
          `Tiempo: ${elapsed.toFixed(2)} ms`,
          "Procesamiento: LOCAL"
        ].join("\n"),
        "ok"
      );

    } catch (error) {

      console.error(
        `[${APP_ID}] SQL ERROR`,
        error
      );

      setStatus(
        "❌ ERROR DE SQL",
        "error"
      );

      setResultMessage(
        error?.message ||
        String(error),
        "error"
      );
    }
  }

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
```
