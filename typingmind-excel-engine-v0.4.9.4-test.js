/*
 * TypingMind Excel Data Engine - Extension v0.4.9.4 TEST
 *
 * OBJETIVO:
 * - Cargar XLSX localmente.
 * - Procesar Excel completamente en el navegador.
 * - Crear tabla DuckDB local.
 * - Inferir tipos.
 * - Ejecutar SQL.
 * - Mostrar resultados seguros para JSON.
 *
 * CAMBIOS v0.4.9.4:
 * - Corrección de consulta de versión DuckDB.
 * - Usa SELECT version().
 * - Normalización robusta de resultados Arrow / BigInt.
 * - Evita errores de JSON.stringify(BigInt).
 * - Corrige resultados numéricos de SUM/AVG/etc.
 *
 * PROCESAMIENTO:
 * 100 % LOCAL
 */

(() => {
  "use strict";

  const APP_ID = "tm-excel-engine-v0494-test";

  const DUCKDB_PACKAGE =
    "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

  const XLSX_PACKAGE =
    "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";

  let duckdb = null;
  let db = null;
  let conn = null;

  let currentFile = null;
  let currentSheet = null;

  /* =========================================================
     UTILIDADES
     ========================================================= */

  function normalizeValue(value) {

    if (value === null || value === undefined) {
      return null;
    }

    if (typeof value === "bigint") {
      const n = Number(value);

      if (
        Number.isSafeInteger(n) &&
        BigInt(n) === value
      ) {
        return n;
      }

      return value.toString();
    }

    if (ArrayBuffer.isView(value)) {

      if (
        typeof value.toArray === "function"
      ) {
        return normalizeValue(
          value.toArray()
        );
      }

      return Array.from(value).map(
        normalizeValue
      );
    }

    if (Array.isArray(value)) {
      return value.map(normalizeValue);
    }

    if (
      typeof value === "object"
    ) {

      if (
        typeof value.toJSON === "function"
      ) {
        try {
          return normalizeValue(
            value.toJSON()
          );
        } catch (_) {}
      }

      const output = {};

      for (
        const [key, val]
        of Object.entries(value)
      ) {
        output[key] =
          normalizeValue(val);
      }

      return output;
    }

    return value;
  }


  function safeJSON(value) {

    return JSON.stringify(
      normalizeValue(value),
      null,
      2
    );
  }


  function getRows(table) {

    let rows = [];

    try {

      if (
        table &&
        typeof table.toArray ===
          "function"
      ) {

        rows = table.toArray();

      }

    } catch (_) {

      rows = [];

    }

    return rows.map(
      row => normalizeValue(row)
    );
  }


  function escapeIdentifier(name) {

    return `"${String(name)
      .replace(/"/g, '""')}"`;

  }


  function escapeString(value) {

    if (value === null ||
        value === undefined) {

      return "NULL";

    }

    return "'" +
      String(value)
        .replace(/'/g, "''") +
      "'";

  }


  function isEmptyRow(row) {

    if (!Array.isArray(row)) {
      return true;
    }

    return row.every(
      value =>
        value === null ||
        value === undefined ||
        String(value).trim() === ""
    );

  }


  function detectDate(value) {

    if (
      value === null ||
      value === undefined ||
      value === ""
    ) {
      return false;
    }

    if (
      value instanceof Date &&
      !isNaN(value.getTime())
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
      /^\d{1,2}[\/-]\d{1,2}[\/-]\d{4}$/.test(
        text
      )
    ) {
      return true;
    }

    return false;
  }


  function detectNumber(value) {

    if (
      value === null ||
      value === undefined ||
      value === ""
    ) {
      return false;
    }

    if (
      typeof value === "number" &&
      Number.isFinite(value)
    ) {
      return true;
    }

    const text =
      String(value)
        .trim()
        .replace(/,/g, "");

    if (text === "") {
      return false;
    }

    return Number.isFinite(
      Number(text)
    );
  }


  function detectInteger(value) {

    if (!detectNumber(value)) {
      return false;
    }

    return Number.isInteger(
      Number(
        String(value)
          .trim()
          .replace(/,/g, "")
      )
    );
  }


  function inferColumnType(
    name,
    values,
    index
  ) {

    const nonEmpty =
      values.filter(
        value =>
          value !== null &&
          value !== undefined &&
          String(value).trim() !== ""
      );

    if (nonEmpty.length === 0) {

      return {
        column_index: index,
        column_name: name,
        duckdb_type: "VARCHAR",
        sqlType: "VARCHAR",
        confidence: "low",
        reason:
          "columna vacía",
        non_empty_values: 0
      };

    }

    const lower =
      String(name)
        .toLowerCase();

    const dateHeader =
      lower.includes("fecha") ||
      lower.includes("date");

    const allDates =
      nonEmpty.every(
        detectDate
      );

    const allIntegers =
      nonEmpty.every(
        detectInteger
      );

    const allNumbers =
      nonEmpty.every(
        detectNumber
      );

    if (
      dateHeader &&
      allDates
    ) {

      return {
        column_index: index,
        column_name: name,
        duckdb_type: "DATE",
        sqlType: "DATE",
        confidence: "high",
        reason:
          "encabezado y valores compatibles con fecha",
        non_empty_values:
          nonEmpty.length
      };

    }

    if (allDates) {

      return {
        column_index: index,
        column_name: name,
        duckdb_type: "DATE",
        sqlType: "DATE",
        confidence: "medium",
        reason:
          "valores compatibles con fecha",
        non_empty_values:
          nonEmpty.length
      };

    }

    if (allIntegers) {

      return {
        column_index: index,
        column_name: name,
        duckdb_type: "BIGINT",
        sqlType: "BIGINT",
        confidence: "high",
        reason:
          "todos los valores son enteros",
        non_empty_values:
          nonEmpty.length
      };

    }

    if (allNumbers) {

      return {
        column_index: index,
        column_name: name,
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
      column_name: name,
      duckdb_type: "VARCHAR",
      sqlType: "VARCHAR",
      confidence: "medium",
      reason:
        "valores tratados como texto",
      non_empty_values:
        nonEmpty.length
    };

  }


  function excelDateToISO(value) {

    if (
      value instanceof Date &&
      !isNaN(value.getTime())
    ) {

      const y =
        value.getFullYear();

      const m =
        String(
          value.getMonth() + 1
        ).padStart(2, "0");

      const d =
        String(
          value.getDate()
        ).padStart(2, "0");

      return `${y}-${m}-${d}`;

    }

    if (
      typeof value === "number"
    ) {

      const date =
        XLSX.SSF.parse_date_code(
          value
        );

      if (date) {

        return [
          date.y,
          String(date.m)
            .padStart(2, "0"),
          String(date.d)
            .padStart(2, "0")
        ].join("-");

      }

    }

    const text =
      String(value)
        .trim();

    if (
      /^\d{4}-\d{1,2}-\d{1,2}$/.test(
        text
      )
    ) {

      const parts =
        text.split("-");

      return [
        parts[0],
        parts[1].padStart(2, "0"),
        parts[2].padStart(2, "0")
      ].join("-");

    }

    const slash =
      text.match(
        /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/
      );

    if (slash) {

      return [
        slash[3],
        slash[2].padStart(2, "0"),
        slash[1].padStart(2, "0")
      ].join("-");

    }

    return null;
  }


  /* =========================================================
     ESTILOS
     ========================================================= */

  function addStyles() {

    if (
      document.getElementById(
        "tmxe-v0494-style"
      )
    ) {
      return;
    }

    const style =
      document.createElement("style");

    style.id =
      "tmxe-v0494-style";

    style.textContent = `
      #tmxe-v0494-button {
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

      #tmxe-v0494-overlay {
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

      #tmxe-v0494-panel {
        width: min(1050px, 96vw);
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
          system-ui, sans-serif;
      }

      #tmxe-v0494-panel h2 {
        margin:
          0 0 5px;

        font-size: 20px;
      }

      #tmxe-v0494-help {
        opacity: .7;
        margin-bottom: 15px;
      }

      #tmxe-v0494-status {
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

      #tmxe-v0494-status[data-kind="ok"] {
        color: #16a34a;
      }

      #tmxe-v0494-status[data-kind="error"] {
        color: #dc2626;
      }

      #tmxe-v0494-result {
        white-space: pre-wrap;

        overflow: auto;

        max-height: 450px;

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

      #tmxe-v0494-sql {
        width: 100%;
        min-height: 150px;

        box-sizing: border-box;

        resize: vertical;

        padding: 12px;

        border-radius: 10px;

        border:
          1px solid
          rgba(127,127,127,.4);

        background: Canvas;
        color: CanvasText;

        font:
          13px/1.45
          ui-monospace,
          SFMono-Regular,
          Menlo,
          monospace;
      }

      #tmxe-v0494-file {
        margin:
          10px 0;
      }

      #tmxe-v0494-actions {
        display: flex;

        gap: 8px;

        flex-wrap: wrap;

        margin: 10px 0;
      }

      #tmxe-v0494-actions button {
        border:
          1px solid
          rgba(127,127,127,.45);

        border-radius: 9px;

        padding: 9px 13px;

        background: transparent;

        color: inherit;

        cursor: pointer;
      }

      #tmxe-v0494-actions button.primary {
        background: #7c3aed;
        color: white;
        border-color: #7c3aed;
      }

      #tmxe-v0494-close {
        border:
          1px solid
          rgba(127,127,127,.45);

        border-radius: 9px;

        padding: 9px 13px;

        background: transparent;

        color: inherit;

        cursor: pointer;
      }

      #tmxe-v0494-info {
        margin:
          12px 0;

        padding: 12px;

        border-radius: 10px;

        background:
          rgba(127,127,127,.08);
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
        "tmxe-v0494-button"
      )
    ) {
      return;
    }

    const button =
      document.createElement("button");

    button.id =
      "tmxe-v0494-button";

    button.type =
      "button";

    button.textContent =
      "📊 Excel v0.4.9.4";

    button.title =
      "Excel Data Engine";

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
        "tmxe-v0494-overlay"
      )
    ) {
      return;
    }

    const overlay =
      document.createElement("div");

    overlay.id =
      "tmxe-v0494-overlay";

    overlay.innerHTML = `

      <div
        id="tmxe-v0494-panel"
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
              v0.4.9.4 TEST
            </h2>

            <div id="tmxe-v0494-help">
              Procesamiento local con
              DuckDB-Wasm.
            </div>

          </div>

          <button
            id="tmxe-v0494-close"
            type="button"
          >
            Cerrar
          </button>

        </div>

        <div id="tmxe-v0494-status">
          Motor listo.
        </div>

        <input
          id="tmxe-v0494-file"
          type="file"
          accept=".xlsx,.xls,.xlsm"
        />

        <div id="tmxe-v0494-actions">

          <button
            id="tmxe-v0494-load"
            class="primary"
            type="button"
          >
            📂 Cargar Excel
          </button>

          <button
            id="tmxe-v0494-count"
            type="button"
          >
            🔢 COUNT
          </button>

          <button
            id="tmxe-v0494-preview"
            type="button"
          >
            👁 Vista previa
          </button>

          <button
            id="tmxe-v0494-summary"
            type="button"
          >
            📈 Resumen
          </button>

        </div>

        <div id="tmxe-v0494-info">
          <strong>SQL manual</strong>

          <textarea
            id="tmxe-v0494-sql"
          >SELECT
    "Año",
    SUM("Total TEU's") AS total
FROM excel_data
GROUP BY "Año"
ORDER BY "Año";</textarea>

          <div id="tmxe-v0494-actions">

            <button
              id="tmxe-v0494-run-sql"
              class="primary"
              type="button"
            >
              ▶ Ejecutar SQL
            </button>

          </div>
        </div>

        <strong>
          Resultado
        </strong>

        <pre
          id="tmxe-v0494-result"
        >—</pre>

        <div style="
          margin-top:12px;
          opacity:.65;
          font-size:12px;
        ">
          v0.4.9.4 TEST —
          Todo el procesamiento del Excel
          se realiza localmente.
        </div>

      </div>
    `;

    document.body.appendChild(
      overlay
    );

    document
      .getElementById(
        "tmxe-v0494-close"
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
        "tmxe-v0494-load"
      )
      .addEventListener(
        "click",
        loadExcel
      );

    document
      .getElementById(
        "tmxe-v0494-count"
      )
      .addEventListener(
        "click",
        runCount
      );

    document
      .getElementById(
        "tmxe-v0494-preview"
      )
      .addEventListener(
        "click",
        runPreview
      );

    document
      .getElementById(
        "tmxe-v0494-summary"
      )
      .addEventListener(
        "click",
        runSummary
      );

    document
      .getElementById(
        "tmxe-v0494-run-sql"
      )
      .addEventListener(
        "click",
        runManualSQL
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
        "tmxe-v0494-status"
      );

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
      document.getElementById(
        "tmxe-v0494-result"
      );

    if (!el) return;

    if (
      typeof value === "string"
    ) {

      el.textContent =
        value;

      return;

    }

    el.textContent =
      safeJSON(value);

  }


  /* =========================================================
     DUCKDB
     ========================================================= */

  async function initDuckDB() {

    if (
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

    /*
     * IMPORTANTE:
     *
     * No usamos:
     *
     * SELECT version
     * FROM pragma_version();
     *
     * porque la columna no se llama "version".
     *
     * Usamos version(), que devuelve
     * directamente la versión.
     */

    const versionResult =
      await conn.query(
        "SELECT version() AS duckdb_version;"
      );

    return {
      bundle,
      version:
        getRows(versionResult)
    };

  }


  /* =========================================================
     CARGAR EXCEL
     ========================================================= */

  async function loadExcel() {

    const input =
      document.getElementById(
        "tmxe-v0494-file"
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

      setResult("");

      const file =
        input.files[0];

      currentFile =
        file;

      setStatus(
        "Inicializando motor local..."
      );

      const engine =
        await initDuckDB();

      setStatus(
        "4/5 — Leyendo Excel localmente..."
      );

      const xlsx =
        await import(
          XLSX_PACKAGE
        );

      const buffer =
        await file.arrayBuffer();

      const workbook =
        xlsx.read(
          buffer,
          {
            type: "array",
            cellDates: true
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

      currentSheet =
        sheets[0];

      const worksheet =
        workbook.Sheets[
          currentSheet
        ];

      const matrix =
        xlsx.utils.sheet_to_json(
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

      if (
        physicalRows === 0
      ) {

        throw new Error(
          "La hoja está vacía."
        );

      }

      /*
       * Primera fila no completamente vacía
       * = encabezados.
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

          headerIndex =
            i;

          break;

        }

      }

      if (
        headerIndex === -1
      ) {

        throw new Error(
          "No se encontraron encabezados."
        );

      }

      const headers =
        matrix[
          headerIndex
        ].map(
          (value, index) => {

            const text =
              value === null ||
              value === undefined ||
              String(value).trim() === ""
                ? `Columna_${index + 1}`
                : String(value).trim();

            return text;

          }
        );

      const rawRows =
        matrix.slice(
          headerIndex + 1
        );

      const rows =
        rawRows.filter(
          row =>
            !isEmptyRow(row)
        );

      /*
       * Normalizamos longitud.
       */

      const normalizedRows =
        rows.map(
          row => {

            const output =
              [];

            for (
              let i = 0;
              i < headers.length;
              i++
            ) {

              output.push(
                row[i] ??
                null
              );

            }

            return output;

          }
        );

      const types =
        headers.map(
          (name, index) =>
            inferColumnType(
              name,
              normalizedRows.map(
                row => row[index]
              ),
              index
            )
        );

      /*
       * Crear tabla.
       */

      try {

        await conn.query(
          "DROP TABLE IF EXISTS excel_data;"
        );

      } catch (_) {}

      const definitions =
        headers.map(
          (name, index) =>
            `${escapeIdentifier(name)} ${types[index].sqlType}`
        ).join(", ");

      await conn.query(
        `CREATE TABLE excel_data (${definitions});`
      );

      /*
       * Insertar filas por lotes.
       */

      const BATCH_SIZE = 500;

      for (
        let start = 0;
        start < normalizedRows.length;
        start += BATCH_SIZE
      ) {

        const batch =
          normalizedRows.slice(
            start,
            start + BATCH_SIZE
          );

        const values =
          batch.map(
            row => {

              const cells =
                row.map(
                  (value, index) => {

                    if (
                      value === null ||
                      value === undefined ||
                      String(value).trim() === ""
                    ) {

                      return "NULL";

                    }

                    const type =
                      types[index].sqlType;

                    if (
                      type === "DATE"
                    ) {

                      const iso =
                        excelDateToISO(
                          value
                        );

                      if (!iso) {
                        return "NULL";
                      }

                      return `DATE ${escapeString(iso)}`;

                    }

                    if (
                      type === "BIGINT"
                    ) {

                      const number =
                        Number(
                          String(value)
                            .replace(/,/g, "")
                            .trim()
                        );

                      if (
                        !Number.isFinite(number)
                      ) {
                        return "NULL";
                      }

                      return String(
                        Math.trunc(number)
                      );

                    }

                    if (
                      type === "DOUBLE"
                    ) {

                      const number =
                        Number(
                          String(value)
                            .replace(/,/g, "")
                            .trim()
                        );

                      if (
                        !Number.isFinite(number)
                      ) {
                        return "NULL";
                      }

                      return String(
                        number
                      );

                    }

                    return escapeString(
                      value
                    );

                  }
                );

              return `(${cells.join(", ")})`;

            }
          ).join(",\n");

        if (values) {

          await conn.query(
            `INSERT INTO excel_data VALUES ${values};`
          );

        }

      }

      setStatus(
        "5/5 — Excel cargado correctamente.",
        "ok"
      );

      const countResult =
        await conn.query(
          "SELECT COUNT(*) AS registros FROM excel_data;"
        );

      const schemaResult =
        await conn.query(
          "DESCRIBE excel_data;"
        );

      const previewResult =
        await conn.query(
          "SELECT * FROM excel_data LIMIT 10;"
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
          currentSheet,

        filas_fisicas_detectadas:
          physicalRows,

        filas_reales:
          normalizedRows.length,

        filas_vacias_ignoradas:
          physicalRows -
          headerIndex -
          1 -
          normalizedRows.length,

        filas_insertadas:
          normalizedRows.length,

        columnas_detectadas:
          headers.length,

        encabezados:
          headers,

        tipos_inferidos:
          types,

        duckdb_table:
          "excel_data",

        duckdb_version:
          engine.version,

        count:
          getRows(countResult),

        schema:
          getRows(schemaResult),

        preview:
          getRows(previewResult)

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


  /* =========================================================
     SQL
     ========================================================= */

  async function ensureLoaded() {

    if (!conn) {

      throw new Error(
        "Primero debes cargar un archivo Excel."
      );

    }

    const check =
      await conn.query(
        `
        SELECT COUNT(*) AS registros
        FROM excel_data;
        `
      );

    return getRows(check);

  }


  async function executeSQL(
    sql
  ) {

    const start =
      performance.now();

    const table =
      await conn.query(
        sql
      );

    const elapsed =
      performance.now() -
      start;

    return {

      procesamiento:
        "LOCAL",

      sql,

      filas_resultado:
        table.numRows,

      tiempo_ms:
        Number(
          elapsed.toFixed(1)
        ),

      resultado:
        getRows(table)

    };

  }


  async function runCount() {

    try {

      await ensureLoaded();

      const result =
        await executeSQL(
          `
          SELECT
            COUNT(*) AS registros
          FROM excel_data;
          `
        );

      setStatus(
        "COUNT ejecutado localmente.",
        "ok"
      );

      setResult(
        result
      );

    } catch (error) {

      showSQLError(
        error
      );

    }

  }


  async function runPreview() {

    try {

      await ensureLoaded();

      const result =
        await executeSQL(
          `
          SELECT *
          FROM excel_data
          LIMIT 10;
          `
        );

      setStatus(
        "Vista previa ejecutada localmente.",
        "ok"
      );

      setResult(
        result
      );

    } catch (error) {

      showSQLError(
        error
      );

    }

  }


  async function runSummary() {

    try {

      await ensureLoaded();

      const schema =
        await conn.query(
          "DESCRIBE excel_data;"
        );

      const columns =
        getRows(schema);

      const numeric =
        columns.filter(
          column =>
            [
              "BIGINT",
              "INTEGER",
              "DOUBLE",
              "DECIMAL",
              "HUGEINT"
            ].some(
              type =>
                String(
                  column.column_type
                ).includes(type)
            )
        );

      if (
        numeric.length === 0
      ) {

        setResult({
          procesamiento:
            "LOCAL",

          mensaje:
            "No se encontraron columnas numéricas."
        });

        return;

      }

      const expressions =
        numeric.map(
          column => {

            const name =
              escapeIdentifier(
                column.column_name
              );

            const alias =
              escapeIdentifier(
                `${column.column_name}__sum`
              );

            return `
              SUM(${name})
              AS ${alias},

              AVG(${name})
              AS ${escapeIdentifier(
                `${column.column_name}__avg`
              )},

              MIN(${name})
              AS ${escapeIdentifier(
                `${column.column_name}__min`
              )},

              MAX(${name})
              AS ${escapeIdentifier(
                `${column.column_name}__max`
              )}
            `;

          }
        ).join(",\n");

      const result =
        await executeSQL(
          `
          SELECT
            ${expressions}
          FROM excel_data;
          `
        );

      setStatus(
        "Resumen ejecutado localmente.",
        "ok"
      );

      setResult(
        result
      );

    } catch (error) {

      showSQLError(
        error
      );

    }

  }


  async function runManualSQL() {

    const textarea =
      document.getElementById(
        "tmxe-v0494-sql"
      );

    const sql =
      textarea
        ? textarea.value.trim()
        : "";

    if (!sql) {

      setStatus(
        "Escribe una consulta SQL.",
        "error"
      );

      return;

    }

    try {

      await ensureLoaded();

      const result =
        await executeSQL(
          sql
        );

      setStatus(
        "SQL ejecutado correctamente en LOCAL.",
        "ok"
      );

      setResult(
        result
      );

    } catch (error) {

      showSQLError(
        error
      );

    }

  }


  function showSQLError(
    error
  ) {

    console.error(
      `[${APP_ID}] SQL`,
      error
    );

    setStatus(
      "❌ ERROR EN SQL",
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
