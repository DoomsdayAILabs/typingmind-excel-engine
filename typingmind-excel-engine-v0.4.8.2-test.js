/*
 * TypingMind Excel Data Engine
 * v0.4.8.2 TEST
 *
 * OBJETIVO
 * --------
 * 1. Cargar un XLSX localmente.
 * 2. Leerlo mediante JavaScript.
 * 3. Detectar filas reales.
 * 4. Inferir tipos.
 * 5. Crear tabla excel_data en DuckDB-Wasm.
 * 6. Ejecutar SQL personalizado LOCALMENTE.
 *
 * NO utiliza:
 *   LOAD excel;
 *
 * El archivo nunca se envía al modelo.
 */

(() => {
  "use strict";

  const APP_ID = "tm-excel-engine-v0482-test";

  const DUCKDB_PACKAGE =
    "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

  const XLSX_PACKAGE =
    "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";

  let duckdb = null;
  let db = null;
  let conn = null;
  let workbook = null;
  let currentData = null;
  let currentHeaders = null;

  function addStyles() {
    if (document.getElementById("tmxe-v0482-style")) return;

    const style = document.createElement("style");

    style.id = "tmxe-v0482-style";

    style.textContent = `
      #tmxe-v0482-button {
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

      #tmxe-v0482-overlay {
        position: fixed;
        inset: 0;
        z-index: 2147483001;

        background: rgba(0,0,0,.48);

        display: flex;
        align-items: center;
        justify-content: center;

        padding: 16px;
      }

      #tmxe-v0482-panel {
        width: min(1050px, 96vw);
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

      #tmxe-v0482-panel h2 {
        margin: 0 0 5px;
        font-size: 20px;
      }

      #tmxe-v0482-help {
        opacity: .7;
        margin-bottom: 15px;
      }

      #tmxe-v0482-status {
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

      #tmxe-v0482-status[data-kind="ok"] {
        color: #16a34a;
      }

      #tmxe-v0482-status[data-kind="error"] {
        color: #dc2626;
      }

      #tmxe-v0482-result {
        white-space: pre-wrap;

        overflow: auto;

        max-height: 380px;

        padding: 12px;

        border-radius: 10px;

        background: rgba(127,127,127,.10);

        font: 12px/1.45 ui-monospace,
          SFMono-Regular,
          Menlo,
          monospace;
      }

      #tmxe-v0482-sql {
        width: 100%;
        min-height: 130px;

        box-sizing: border-box;

        resize: vertical;

        padding: 12px;

        border-radius: 10px;

        border: 1px solid rgba(127,127,127,.35);

        background: Canvas;
        color: CanvasText;

        font: 13px/1.45 ui-monospace,
          SFMono-Regular,
          Menlo,
          monospace;
      }

      #tmxe-v0482-actions {
        display: flex;
        gap: 8px;
        flex-wrap: wrap;
        margin: 10px 0;
      }

      #tmxe-v0482-actions button {
        border: 1px solid rgba(127,127,127,.45);

        border-radius: 9px;

        padding: 9px 13px;

        background: transparent;

        color: inherit;

        cursor: pointer;
      }

      #tmxe-v0482-load {
        background: #2563eb !important;
        color: white !important;
        border-color: #2563eb !important;
      }

      #tmxe-v0482-run {
        background: #16a34a !important;
        color: white !important;
        border-color: #16a34a !important;
      }

      #tmxe-v0482-close {
        border: 1px solid rgba(127,127,127,.45);

        border-radius: 9px;

        padding: 9px 13px;

        background: transparent;

        color: inherit;

        cursor: pointer;
      }

      #tmxe-v0482-file {
        display: none;
      }

      #tmxe-v0482-info {
        padding: 12px;
        margin: 10px 0;

        border-radius: 10px;

        background: rgba(37,99,235,.08);

        border: 1px solid rgba(37,99,235,.20);
      }
    `;

    document.head.appendChild(style);
  }

  function createButton() {
    if (document.getElementById("tmxe-v0482-button")) return;

    const button = document.createElement("button");

    button.id = "tmxe-v0482-button";

    button.type = "button";

    button.textContent = "📊 Excel Engine v0.4.8.2";

    button.title =
      "Excel Engine — procesamiento local";

    button.addEventListener(
      "click",
      openPanel
    );

    document.body.appendChild(button);
  }

  function openPanel() {
    if (
      document.getElementById(
        "tmxe-v0482-overlay"
      )
    ) {
      return;
    }

    const overlay =
      document.createElement("div");

    overlay.id =
      "tmxe-v0482-overlay";

    overlay.innerHTML = `
      <div
        id="tmxe-v0482-panel"
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
              📊 Excel Data Engine v0.4.8.2
            </h2>

            <div id="tmxe-v0482-help">
              Excel → JavaScript → DuckDB-Wasm → SQL LOCAL
            </div>

          </div>

          <button
            id="tmxe-v0482-close"
            type="button"
          >
            Cerrar
          </button>

        </div>

        <div id="tmxe-v0482-status">
          Listo.
        </div>

        <input
          id="tmxe-v0482-file"
          type="file"
          accept=".xlsx,.xls"
        />

        <div id="tmxe-v0482-actions">

          <button id="tmxe-v0482-load">
            📂 Cargar Excel
          </button>

        </div>

        <div id="tmxe-v0482-info">
          <strong>Estado del motor</strong>
          <br>
          Archivo: —
          <br>
          Tabla DuckDB: —
          <br>
          Registros: —
        </div>

        <strong>SQL local</strong>

        <textarea
          id="tmxe-v0482-sql"
          spellcheck="false"
        >SELECT COUNT(*) AS registros
FROM excel_data;</textarea>

        <div id="tmxe-v0482-actions">

          <button id="tmxe-v0482-run">
            ▶ Ejecutar SQL LOCAL
          </button>

          <button
            id="tmxe-v0482-count"
          >
            COUNT(*)
          </button>

          <button
            id="tmxe-v0482-preview"
          >
            Vista previa
          </button>

          <button
            id="tmxe-v0482-summary"
          >
            Resumen
          </button>

        </div>

        <strong>Resultado</strong>

        <pre id="tmxe-v0482-result">—</pre>

        <div style="
          margin-top:12px;
          opacity:.65;
          font-size:12px;
        ">
          v0.4.8.2 TEST — Procesamiento local.
          El XLSX no se envía al modelo.
        </div>

      </div>
    `;

    document.body.appendChild(overlay);

    document
      .getElementById("tmxe-v0482-close")
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
      .getElementById("tmxe-v0482-load")
      .addEventListener(
        "click",
        () => {
          document
            .getElementById("tmxe-v0482-file")
            .click();
        }
      );

    document
      .getElementById("tmxe-v0482-file")
      .addEventListener(
        "change",
        loadExcel
      );

    document
      .getElementById("tmxe-v0482-run")
      .addEventListener(
        "click",
        runSQL
      );

    document
      .getElementById("tmxe-v0482-count")
      .addEventListener(
        "click",
        () => {
          document
            .getElementById("tmxe-v0482-sql")
            .value =
              `SELECT COUNT(*) AS registros
FROM excel_data;`;

          runSQL();
        }
      );

    document
      .getElementById("tmxe-v0482-preview")
      .addEventListener(
        "click",
        () => {
          document
            .getElementById("tmxe-v0482-sql")
            .value =
              `SELECT *
FROM excel_data
LIMIT 10;`;

          runSQL();
        }
      );

    document
      .getElementById("tmxe-v0482-summary")
      .addEventListener(
        "click",
        () => {
          if (!currentHeaders) return;

          const numericColumns =
            currentHeaders
              .filter(
                header =>
                  header.type === "BIGINT" ||
                  header.type === "DOUBLE"
              )
              .map(
                header =>
                  `SUM("${header.name}") AS "${header.name}__sum",
AVG("${header.name}") AS "${header.name}__avg",
MIN("${header.name}") AS "${header.name}__min",
MAX("${header.name}") AS "${header.name}__max"`
              );

          if (!numericColumns.length) {
            setResult(
              "No se detectaron columnas numéricas."
            );
            return;
          }

          document
            .getElementById("tmxe-v0482-sql")
            .value =
              `SELECT
${numericColumns.join(",\n")}
FROM excel_data;`;

          runSQL();
        }
      );
  }

  function setStatus(
    text,
    kind = ""
  ) {
    const el =
      document.getElementById(
        "tmxe-v0482-status"
      );

    if (!el) return;

    el.textContent = text;

    if (kind) {
      el.dataset.kind = kind;
    } else {
      delete el.dataset.kind;
    }
  }

  function normalizeValue(value) {
    if (typeof value === "bigint") {
      return value.toString();
    }

    if (
      value &&
      typeof value === "object"
    ) {
      if (
        typeof value.toISOString ===
        "function"
      ) {
        return value.toISOString();
      }
    }

    return value;
  }

  function serialize(value) {
    return JSON.stringify(
      value,
      (key, val) => {
        if (typeof val === "bigint") {
          return val.toString();
        }

        return normalizeValue(val);
      },
      2
    );
  }

  function setResult(value) {
    const el =
      document.getElementById(
        "tmxe-v0482-result"
      );

    if (!el) return;

    if (
      typeof value === "string"
    ) {
      el.textContent = value;
      return;
    }

    el.textContent = serialize(value);
  }

  function updateInfo(
    fileName,
    rowCount
  ) {
    const el =
      document.getElementById(
        "tmxe-v0482-info"
      );

    if (!el) return;

    el.innerHTML = `
      <strong>Estado del motor</strong>
      <br>
      Archivo: ${escapeHTML(fileName)}
      <br>
      Tabla DuckDB: <code>excel_data</code>
      <br>
      Registros: ${rowCount}
    `;
  }

  function escapeHTML(value) {
    return String(value)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
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

    const milliseconds =
      Number(serial) *
      86400000;

    return new Date(
      epoch + milliseconds
    );
  }

  function isDateLikeHeader(
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
      normalized.includes("dia")
    );
  }

  function inferColumnType(
    name,
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
        name,
        type: "VARCHAR",
        confidence: "low",
        reason: "sin valores"
      };
    }

    if (
      isDateLikeHeader(name)
    ) {
      const compatible =
        nonEmpty.every(
          value => {
            if (
              value instanceof Date
            ) {
              return true;
            }

            if (
              typeof value ===
              "number"
            ) {
              return true;
            }

            const d =
              new Date(value);

            return !Number.isNaN(
              d.getTime()
            );
          }
        );

      if (compatible) {
        return {
          name,
          type: "DATE",
          confidence: "high",
          reason:
            "encabezado y valores compatibles con fecha"
        };
      }
    }

    const allIntegers =
      nonEmpty.every(
        value => {
          if (
            typeof value ===
            "number"
          ) {
            return (
              Number.isFinite(value) &&
              Number.isInteger(value)
            );
          }

          return /^[-+]?\d+$/.test(
            String(value).trim()
          );
        }
      );

    if (allIntegers) {
      return {
        name,
        type: "BIGINT",
        confidence: "high",
        reason:
          "todos los valores son enteros"
      };
    }

    const allNumbers =
      nonEmpty.every(
        value => {
          const n =
            Number(value);

          return Number.isFinite(n);
        }
      );

    if (allNumbers) {
      return {
        name,
        type: "DOUBLE",
        confidence: "high",
        reason:
          "todos los valores son numéricos"
      };
    }

    return {
      name,
      type: "VARCHAR",
      confidence: "medium",
      reason:
        "valores no compatibles con tipo numérico"
    };
  }

  function escapeIdentifier(
    value
  ) {
    return String(value)
      .replaceAll('"', '""');
  }

  function escapeString(
    value
  ) {
    return String(value)
      .replaceAll("'", "''");
  }

  function convertForDuckDB(
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

    if (type === "BIGINT") {
      const n =
        Number(value);

      if (
        Number.isInteger(n) &&
        Number.isSafeInteger(n)
      ) {
        return String(n);
      }

      return "NULL";
    }

    if (type === "DOUBLE") {
      const n =
        Number(value);

      if (
        Number.isFinite(n)
      ) {
        return String(n);
      }

      return "NULL";
    }

    if (type === "DATE") {
      let date = null;

      if (
        value instanceof Date
      ) {
        date = value;
      } else if (
        typeof value ===
        "number"
      ) {
        date =
          excelSerialToDate(
            value
          );
      } else {
        date =
          new Date(value);
      }

      if (
        !date ||
        Number.isNaN(
          date.getTime()
        )
      ) {
        return "NULL";
      }

      const iso =
        date.toISOString()
          .slice(0, 10);

      return `DATE '${iso}'`;
    }

    return `'${escapeString(
      value
    )}'`;
  }

  async function initDuckDB() {
    if (
      duckdb &&
      db &&
      conn
    ) {
      return;
    }

    setStatus(
      "1/6 — Cargando DuckDB-Wasm..."
    );

    duckdb =
      await import(
        DUCKDB_PACKAGE
      );

    setStatus(
      "2/6 — Seleccionando bundle..."
    );

    const bundles =
      duckdb.getJsDelivrBundles();

    const bundle =
      await duckdb.selectBundle(
        bundles
      );

    setStatus(
      "3/6 — Inicializando WebAssembly..."
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
  }

  async function loadExcel(
    event
  ) {
    const file =
      event.target.files?.[0];

    if (!file) return;

    try {
      setResult("");

      setStatus(
        "Preparando DuckDB-Wasm..."
      );

      await initDuckDB();

      setStatus(
        "4/6 — Cargando lector XLSX..."
      );

      const XLSX =
        await import(
          XLSX_PACKAGE
        );

      setStatus(
        "5/6 — Leyendo Excel localmente..."
      );

      const buffer =
        await file.arrayBuffer();

      workbook =
        XLSX.read(
          buffer,
          {
            type: "array",
            cellDates: true
          }
        );

      const sheetName =
        workbook.SheetNames[0];

      const sheet =
        workbook.Sheets[
          sheetName
        ];

      const raw =
        XLSX.utils.sheet_to_json(
          sheet,
          {
            header: 1,
            defval: null,
            raw: true,
            blankrows: true
          }
        );

      if (
        !raw ||
        raw.length === 0
      ) {
        throw new Error(
          "La hoja está vacía."
        );
      }

      let headerIndex = 0;

      while (
        headerIndex <
          raw.length &&
        raw[headerIndex].every(
          value =>
            value === null ||
            value === undefined ||
            String(value).trim() === ""
        )
      ) {
        headerIndex++;
      }

      if (
        headerIndex >= raw.length
      ) {
        throw new Error(
          "No se encontró una fila de encabezados."
        );
      }

      const headers =
        raw[headerIndex].map(
          (value, index) => {
            const name =
              value === null ||
              value === undefined ||
              String(value).trim() === ""
                ? `columna_${index + 1}`
                : String(value).trim();

            return name;
          }
        );

      const dataRows =
        raw
          .slice(
            headerIndex + 1
          )
          .filter(
            row =>
              row.some(
                value =>
                  value !== null &&
                  value !== undefined &&
                  String(value).trim() !== ""
              )
          );

      if (!dataRows.length) {
        throw new Error(
          "No se encontraron filas con datos."
        );
      }

      const types =
        headers.map(
          (name, index) =>
            inferColumnType(
              name,
              dataRows.map(
                row =>
                  row[index]
              )
            )
        );

      setStatus(
        "6/6 — Creando tabla DuckDB local..."
      );

      try {
        await conn.query(
          "DROP TABLE IF EXISTS excel_data;"
        );
      } catch (_) {}

      const definitions =
        types
          .map(
            item =>
              `"${escapeIdentifier(
                item.name
              )}" ${item.type}`
          )
          .join(", ");

      await conn.query(
        `CREATE TABLE excel_data (${definitions});`
      );

      const columnList =
        headers
          .map(
            name =>
              `"${escapeIdentifier(
                name
              )}"`
          )
          .join(", ");

      const batchSize = 250;

      for (
        let start = 0;
        start < dataRows.length;
        start += batchSize
      ) {
        const batch =
          dataRows.slice(
            start,
            start + batchSize
          );

        const valuesSQL =
          batch
            .map(
              row => {
                const values =
                  headers.map(
                    (name, index) =>
                      convertForDuckDB(
                        row[index],
                        types[index].type
                      )
                  );

                return `(${values.join(
                  ", "
                )})`;
              }
            )
            .join(",\n");

        await conn.query(
          `INSERT INTO excel_data (${columnList})
           VALUES ${valuesSQL};`
        );
      }

      currentData =
        dataRows;

      currentHeaders =
        types;

      const countResult =
        await conn.query(
          `SELECT COUNT(*) AS registros
           FROM excel_data;`
        );

      const countRows =
        countResult.toArray();

      const schemaResult =
        await conn.query(
          `DESCRIBE excel_data;`
        );

      const schema =
        schemaResult.toArray();

      const previewResult =
        await conn.query(
          `SELECT *
           FROM excel_data
           LIMIT 10;`
        );

      const preview =
        previewResult
          .toArray();

      updateInfo(
        file.name,
        dataRows.length
      );

      setStatus(
        "Análisis terminado correctamente.",
        "ok"
      );

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
          workbook.SheetNames,

        hoja_principal:
          sheetName,

        filas_fisicas_detectadas:
          raw.length,

        filas_reales:
          dataRows.length,

        filas_vacias_ignoradas:
          raw.length -
          headerIndex -
          1 -
          dataRows.length,

        filas_insertadas:
          dataRows.length,

        columnas_detectadas:
          headers.length,

        encabezados:
          headers,

        duckdb_table:
          "excel_data",

        duckdb_version:
          (
            await conn.query(
              "SELECT version() AS duckdb_version;"
            )
          ).toArray(),

        tipos_inferidos:
          types.map(
            (item, index) => ({
              column_index:
                index,

              column_name:
                item.name,

              duckdb_type:
                item.type,

              confidence:
                item.confidence,

              reason:
                item.reason,

              non_empty_values:
                dataRows.filter(
                  row =>
                    row[index] !==
                      null &&
                    row[index] !==
                      undefined &&
                    String(
                      row[index]
                    ).trim() !== ""
                ).length
            })
          ),

        schema_duckdb:
          schema,

        count_sql:
          countRows,

        preview:
          preview
      });

      document
        .getElementById(
          "tmxe-v0482-sql"
        ).value =
        `SELECT COUNT(*) AS registros
FROM excel_data;`;

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

  async function runSQL() {
    if (!conn) {
      setResult(
        "Primero debes cargar un archivo Excel."
      );

      return;
    }

    const textarea =
      document.getElementById(
        "tmxe-v0482-sql"
      );

    const sql =
      textarea.value.trim();

    if (!sql) {
      setResult(
        "La consulta SQL está vacía."
      );

      return;
    }

    try {
      setStatus(
        "Ejecutando SQL LOCAL..."
      );

      const result =
        await conn.query(
          sql
        );

      const rows =
        result.toArray();

      setStatus(
        `SQL ejecutado correctamente. Filas devueltas: ${rows.length}`,
        "ok"
      );

      setResult({
        procesamiento:
          "LOCAL",

        sql:
          sql,

        filas_resultado:
          rows.length,

        resultado:
          rows
      });

    } catch (error) {
      console.error(
        `[${APP_ID}] SQL ERROR`,
        error
      );

      setStatus(
        "❌ ERROR SQL",
        "error"
      );

      setResult({
        mensaje:
          error?.message ||
          String(error),

        sql:
          sql
      });
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
      { once: true }
    );
  } else {
    init();
  }

})();
