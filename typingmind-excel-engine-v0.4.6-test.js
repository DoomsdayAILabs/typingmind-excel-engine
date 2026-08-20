/*
 * TypingMind Excel Data Engine - Extension v0.4.6 TEST
 *
 * XLSX LOCAL -> SHEETJS -> TYPE INFERENCE -> DUCKDB-WASM
 *
 * v0.4.6
 *
 * Objetivos:
 * - Leer XLSX localmente.
 * - Detectar tipos de columnas.
 * - Convertir fechas seriales de Excel.
 * - Crear tabla DuckDB con tipos reales.
 * - Ejecutar estadísticas SQL.
 * - Mostrar esquema y preview.
 *
 * NO:
 * - No utiliza LOAD excel.
 * - No sube el Excel.
 * - No envía datos al LLM.
 * - No utiliza servidor propio.
 */

(() => {
  "use strict";

  const APP_ID =
    "tm-excel-engine-v046-test";

  const DUCKDB_PACKAGE =
    "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

  const SHEETJS_URL =
    "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js";

  let duckdb = null;
  let db = null;
  let conn = null;

  /*
   * ======================================================
   * ESTILOS
   * ======================================================
   */

  function addStyles() {

    if (
      document.getElementById(
        "tmxe-v046-style"
      )
    ) {
      return;
    }

    const style =
      document.createElement("style");

    style.id =
      "tmxe-v046-style";

    style.textContent = `

      #tmxe-v046-button {
        position: fixed;
        right: 18px;
        bottom: 86px;
        z-index: 2147483000;

        border: 0;
        border-radius: 999px;

        padding: 11px 16px;

        background: #2563eb;
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

      #tmxe-v046-overlay {
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

      #tmxe-v046-panel {
        width:
          min(1100px, 96vw);

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

      #tmxe-v046-panel h2 {
        margin:
          0 0 5px;

        font-size: 20px;
      }

      #tmxe-v046-help {
        opacity: .7;
        margin-bottom: 12px;
      }

      #tmxe-v046-file {
        display: block;

        width: 100%;

        margin:
          12px 0;

        padding: 8px;

        border:
          1px solid
          rgba(127,127,127,.35);

        border-radius: 8px;

        box-sizing: border-box;
      }

      #tmxe-v046-status {
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

        margin:
          12px 0;
      }

      #tmxe-v046-status[data-kind="ok"] {
        color: #16a34a;
      }

      #tmxe-v046-status[data-kind="error"] {
        color: #dc2626;
      }

      #tmxe-v046-result {
        white-space: pre-wrap;

        overflow: auto;

        max-height: 600px;

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

      #tmxe-v046-run,
      #tmxe-v046-close {
        border:
          1px solid
          rgba(127,127,127,.45);

        border-radius: 9px;

        padding:
          9px 13px;

        background: transparent;

        color: inherit;

        cursor: pointer;
      }

      #tmxe-v046-run {
        background:
          #2563eb !important;

        color:
          white !important;

        border-color:
          #2563eb !important;
      }

      #tmxe-v046-actions {
        display: flex;

        gap: 8px;

        flex-wrap: wrap;

        margin:
          10px 0;
      }

    `;

    document.head.appendChild(style);
  }

  /*
   * ======================================================
   * BOTÓN
   * ======================================================
   */

  function createButton() {

    if (
      document.getElementById(
        "tmxe-v046-button"
      )
    ) {
      return;
    }

    const button =
      document.createElement("button");

    button.id =
      "tmxe-v046-button";

    button.type =
      "button";

    button.textContent =
      "📊 Excel v0.4.6";

    button.title =
      "Excel local → DuckDB";

    button.addEventListener(
      "click",
      openPanel
    );

    document.body.appendChild(
      button
    );
  }

  /*
   * ======================================================
   * PANEL
   * ======================================================
   */

  function openPanel() {

    if (
      document.getElementById(
        "tmxe-v046-overlay"
      )
    ) {
      return;
    }

    const overlay =
      document.createElement("div");

    overlay.id =
      "tmxe-v046-overlay";

    overlay.innerHTML = `

      <div
        id="tmxe-v046-panel"
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
              📊 Excel Engine v0.4.6 TEST
            </h2>

            <div id="tmxe-v046-help">
              XLSX local → detección de tipos →
              DuckDB-Wasm → SQL
            </div>

          </div>

          <button
            id="tmxe-v046-close"
            type="button"
          >
            Cerrar
          </button>

        </div>

        <input
          id="tmxe-v046-file"
          type="file"
          accept=".xlsx,.xls,.xlsb,.ods"
        />

        <div
          id="tmxe-v046-actions"
        >

          <button
            id="tmxe-v046-run"
            type="button"
          >
            🚀 Analizar Excel
          </button>

        </div>

        <div
          id="tmxe-v046-status"
        >
          Selecciona un archivo Excel.
        </div>

        <strong>
          Resultado
        </strong>

        <pre
          id="tmxe-v046-result"
        >—</pre>

        <div style="
          margin-top:12px;
          opacity:.65;
          font-size:12px;
        ">

          v0.4.6 TEST —
          procesamiento completamente local.

        </div>

      </div>

    `;

    document.body.appendChild(
      overlay
    );

    document
      .getElementById(
        "tmxe-v046-close"
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
        "tmxe-v046-run"
      )
      .addEventListener(
        "click",
        loadExcel
      );
  }

  /*
   * ======================================================
   * STATUS
   * ======================================================
   */

  function setStatus(
    text,
    kind = ""
  ) {

    const el =
      document.getElementById(
        "tmxe-v046-status"
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

  /*
   * ======================================================
   * JSON SEGURO
   *
   * DuckDB puede devolver BigInt.
   * ======================================================
   */

  function safeJSONStringify(
    value
  ) {

    return JSON.stringify(
      value,
      (
        key,
        currentValue
      ) => {

        if (
          typeof currentValue ===
          "bigint"
        ) {

          return currentValue.toString();

        }

        return currentValue;
      },
      2
    );
  }

  function setResult(
    value
  ) {

    const el =
      document.getElementById(
        "tmxe-v046-result"
      );

    if (!el) {
      return;
    }

    if (
      typeof value === "string"
    ) {

      el.textContent =
        value;

      return;
    }

    el.textContent =
      safeJSONStringify(
        value
      );
  }

  /*
   * ======================================================
   * SHEETJS
   * ======================================================
   */

  async function loadSheetJS() {

    if (
      window.XLSX
    ) {

      return window.XLSX;

    }

    const existing =
      document.querySelector(
        `script[src="${SHEETJS_URL}"]`
      );

    if (existing) {

      await new Promise(
        (
          resolve,
          reject
        ) => {

          existing.addEventListener(
            "load",
            resolve,
            { once: true }
          );

          existing.addEventListener(
            "error",
            reject,
            { once: true }
          );

        }
      );

      if (
        window.XLSX
      ) {

        return window.XLSX;

      }
    }

    await new Promise(
      (
        resolve,
        reject
      ) => {

        const script =
          document.createElement(
            "script"
          );

        script.src =
          SHEETJS_URL;

        script.onload =
          resolve;

        script.onerror =
          () => {

            reject(
              new Error(
                "No se pudo cargar SheetJS."
              )
            );

          };

        document.head.appendChild(
          script
        );

      }
    );

    if (
      !window.XLSX
    ) {

      throw new Error(
        "SheetJS cargó pero XLSX no está disponible."
      );

    }

    return window.XLSX;
  }

  /*
   * ======================================================
   * DUCKDB
   * ======================================================
   */

  async function initDuckDB() {

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

  /*
   * ======================================================
   * IDENTIFICADOR SQL
   * ======================================================
   */

  function safeIdentifier(
    name
  ) {

    return (
      `"${
        String(name)
          .replace(
            /"/g,
            '""'
          )
      }"`
    );
  }

  /*
   * ======================================================
   * ENCABEZADOS
   * ======================================================
   */

  function normalizeHeaders(
    headers
  ) {

    const result = [];

    const used =
      new Set();

    headers.forEach(
      (
        header,
        index
      ) => {

        let name =
          String(
            header ??
            `column_${index + 1}`
          ).trim();

        if (!name) {
          name =
            `column_${index + 1}`;
        }

        const original =
          name;

        let counter =
          2;

        while (
          used.has(name)
        ) {

          name =
            `${original}_${counter}`;

          counter++;
        }

        used.add(name);

        result.push(
          name
        );

      }
    );

    return result;
  }

  /*
   * ======================================================
   * FECHA SERIAL DE EXCEL
   *
   * Excel normalmente utiliza:
   *
   * 1 = 1900-01-01
   *
   * con el conocido bug del año 1900.
   *
   * SheetJS trabaja con seriales cuando
   * raw=true.
   * ======================================================
   */

  function excelSerialToDate(
    serial
  ) {

    if (
      typeof serial !==
      "number"
    ) {
      return null;
    }

    if (
      !Number.isFinite(serial)
    ) {
      return null;
    }

    /*
     * Rango razonable para fechas Excel.
     *
     * Evitamos interpretar números
     * normales como fechas.
     */

    if (
      serial < 20000 ||
      serial > 80000
    ) {
      return null;
    }

    /*
     * Excel 1900 date system.
     */

    const utcDays =
      Math.floor(serial);

    const fractional =
      serial -
      utcDays;

    const milliseconds =
      Math.round(
        fractional *
        86400000
      );

    const epoch =
      Date.UTC(
        1899,
        11,
        30
      );

    const date =
      new Date(
        epoch +
        utcDays *
        86400000 +
        milliseconds
      );

    if (
      Number.isNaN(
        date.getTime()
      )
    ) {
      return null;
    }

    return date;
  }

  /*
   * ======================================================
   * DETECTAR SI EL ENCABEZADO PARECE FECHA
   * ======================================================
   */

  function headerLooksLikeDate(
    header
  ) {

    const text =
      String(
        header || ""
      )
        .toLowerCase()
        .trim();

    const patterns = [

      "fecha",
      "date",
      "datetime",
      "timestamp",
      "día",
      "dia"

    ];

    return patterns.some(
      pattern =>
        text.includes(
          pattern
        )
    );
  }

  /*
   * ======================================================
   * DETECCIÓN DE TIPO
   * ======================================================
   */

  function inferColumnType(
    header,
    values
  ) {

    const nonEmpty =
      values.filter(
        value =>
          value !== null &&
          value !== undefined &&
          value !== ""
      );

    if (
      nonEmpty.length === 0
    ) {

      return {
        duckdb_type:
          "VARCHAR",

        confidence:
          "none",

        reason:
          "columna vacía"
      };

    }

    /*
     * FECHA
     */

    if (
      headerLooksLikeDate(
        header
      )
    ) {

      const dateValues =
        nonEmpty.map(
          value =>
            excelSerialToDate(
              value
            )
        );

      const validDates =
        dateValues.filter(
          date =>
            date !== null
        );

      if (
        validDates.length /
          nonEmpty.length >=
        0.8
      ) {

        return {
          duckdb_type:
            "DATE",

          confidence:
            "high",

          reason:
            "encabezado y valores compatibles con fecha"
        };

      }

    }

    /*
     * BOOLEAN
     */

    const booleanCompatible =
      nonEmpty.every(
        value => {

          if (
            typeof value ===
            "boolean"
          ) {
            return true;
          }

          const text =
            String(value)
              .toLowerCase()
              .trim();

          return (
            text === "true" ||
            text === "false" ||
            text === "sí" ||
            text === "si" ||
            text === "no"
          );

        }
      );

    if (
      booleanCompatible
    ) {

      return {
        duckdb_type:
          "BOOLEAN",

        confidence:
          "high",

        reason:
          "valores booleanos"
      };

    }

    /*
     * NUMÉRICO
     */

    const numericCompatible =
      nonEmpty.every(
        value => {

          if (
            typeof value ===
            "number"
          ) {

            return Number.isFinite(
              value
            );

          }

          const text =
            String(value)
              .trim()
              .replace(
                /,/g,
                ""
              );

          if (
            text === ""
          ) {
            return true;
          }

          return (
            Number.isFinite(
              Number(text)
            )
          );

        }
      );

    if (
      numericCompatible
    ) {

      const numbers =
        nonEmpty.map(
          value =>
            Number(
              String(value)
                .replace(
                  /,/g,
                  ""
                )
            )
        );

      const allIntegers =
        numbers.every(
          number =>
            Number.isInteger(
              number
            )
        );

      if (
        allIntegers
      ) {

        return {
          duckdb_type:
            "BIGINT",

          confidence:
            "high",

          reason:
            "todos los valores son enteros"
        };

      }

      return {
        duckdb_type:
          "DOUBLE",

        confidence:
          "high",

        reason:
          "valores numéricos con decimales"
      };

    }

    /*
     * TEXTO
     */

    return {
      duckdb_type:
        "VARCHAR",

      confidence:
        "high",

      reason:
        "valores no compatibles con tipos anteriores"
    };
  }

  /*
   * ======================================================
   * INFERENCIA DE TODAS LAS COLUMNAS
   * ======================================================
   */

  function inferSchema(
    headers,
    dataRows
  ) {

    return headers.map(
      (
        header,
        columnIndex
      ) => {

        const values =
          dataRows.map(
            row =>
              row[columnIndex]
          );

        const inferred =
          inferColumnType(
            header,
            values
          );

        return {

          column_index:
            columnIndex,

          column_name:
            header,

          duckdb_type:
            inferred.duckdb_type,

          confidence:
            inferred.confidence,

          reason:
            inferred.reason,

          non_empty_values:
            values.filter(
              value =>
                value !== null &&
                value !== undefined &&
                value !== ""
            ).length

        };

      }
    );
  }

  /*
   * ======================================================
   * CONVERSIÓN DE VALORES SEGÚN TIPO
   * ======================================================
   */

  function convertValue(
    value,
    type
  ) {

    /*
     * NULL
     */

    if (
      value === null ||
      value === undefined ||
      value === ""
    ) {

      return null;
    }

    /*
     * DATE
     */

    if (
      type === "DATE"
    ) {

      const date =
        excelSerialToDate(
          value
        );

      if (
        date
      ) {

        return (
          date
            .toISOString()
            .slice(
              0,
              10
            )
        );

      }

      /*
       * Si ya viene como fecha
       * o texto compatible.
       */

      if (
        value instanceof Date
      ) {

        return (
          value
            .toISOString()
            .slice(
              0,
              10
            )
        );

      }

      return null;
    }

    /*
     * BIGINT
     */

    if (
      type === "BIGINT"
    ) {

      const number =
        Number(
          String(value)
            .replace(
              /,/g,
              ""
            )
        );

      if (
        Number.isFinite(
          number
        )
      ) {

        return Math.trunc(
          number
        );

      }

      return null;
    }

    /*
     * DOUBLE
     */

    if (
      type === "DOUBLE"
    ) {

      const number =
        Number(
          String(value)
            .replace(
              /,/g,
              ""
            )
        );

      if (
        Number.isFinite(
          number
        )
      ) {

        return number;

      }

      return null;
    }

    /*
     * BOOLEAN
     */

    if (
      type === "BOOLEAN"
    ) {

      if (
        typeof value ===
        "boolean"
      ) {

        return value;
      }

      const text =
        String(value)
          .toLowerCase()
          .trim();

      return (
        text === "true" ||
        text === "sí" ||
        text === "si"
      );
    }

    /*
     * VARCHAR
     */

    return String(
      value
    );
  }

  /*
   * ======================================================
   * SQL VALUE
   * ======================================================
   */

  function sqlValue(
    value,
    type
  ) {

    const converted =
      convertValue(
        value,
        type
      );

    if (
      converted === null ||
      converted === undefined
    ) {

      return "NULL";
    }

    if (
      type === "BIGINT" ||
      type === "DOUBLE"
    ) {

      return String(
        converted
      );
    }

    if (
      type === "BOOLEAN"
    ) {

      return converted
        ? "TRUE"
        : "FALSE";
    }

    const escaped =
      String(
        converted
      )
        .replace(
          /'/g,
          "''"
        );

    return `'${escaped}'`;
  }

  /*
   * ======================================================
   * CARGA PRINCIPAL
   * ======================================================
   */

  async function loadExcel() {

    try {

      setResult("");

      /*
       * --------------------------------------------------
       * ARCHIVO
       * --------------------------------------------------
       */

      const fileInput =
        document.getElementById(
          "tmxe-v046-file"
        );

      const file =
        fileInput?.files?.[0];

      if (!file) {

        setStatus(
          "❌ Primero selecciona un archivo Excel.",
          "error"
        );

        return;
      }

      /*
       * --------------------------------------------------
       * 1/10
       * --------------------------------------------------
       */

      setStatus(
        "1/10 — Cargando SheetJS..."
      );

      const XLSX =
        await loadSheetJS();

      /*
       * --------------------------------------------------
       * 2/10
       * --------------------------------------------------
       */

      setStatus(
        "2/10 — Leyendo archivo local..."
      );

      const arrayBuffer =
        await file.arrayBuffer();

      /*
       * --------------------------------------------------
       * 3/10
       * --------------------------------------------------
       */

      setStatus(
        "3/10 — Analizando XLSX..."
      );

      const workbook =
        XLSX.read(
          arrayBuffer,
          {
            type:
              "array",

            dense:
              true
          }
        );

      const sheetNames =
        workbook.SheetNames;

      if (
        !sheetNames.length
      ) {

        throw new Error(
          "El archivo no contiene hojas."
        );

      }

      /*
       * Primera hoja
       */

      const sheetName =
        sheetNames[0];

      const worksheet =
        workbook.Sheets[
          sheetName
        ];

      /*
       * --------------------------------------------------
       * 4/10
       * --------------------------------------------------
       */

      setStatus(
        "4/10 — Extrayendo filas y encabezados..."
      );

      const rows =
        XLSX.utils.sheet_to_json(
          worksheet,
          {
            header:
              1,

            defval:
              null,

            raw:
              true
          }
        );

      if (
        !rows.length
      ) {

        throw new Error(
          "La hoja está vacía."
        );

      }

      const headers =
        normalizeHeaders(
          rows[0]
        );

      const dataRows =
        rows.slice(1);

      /*
       * --------------------------------------------------
       * 5/10
       * --------------------------------------------------
       */

      setStatus(
        "5/10 — Detectando tipos de columnas..."
      );

      const inferredSchema =
        inferSchema(
          headers,
          dataRows
        );

      /*
       * --------------------------------------------------
       * 6/10
       * --------------------------------------------------
       */

      setStatus(
        "6/10 — Inicializando DuckDB-Wasm..."
      );

      const bundle =
        await initDuckDB();

      /*
       * --------------------------------------------------
       * 7/10
       * --------------------------------------------------
       */

      setStatus(
        "7/10 — Creando tabla con tipos reales..."
      );

      await conn.query(
        "DROP TABLE IF EXISTS excel_data;"
      );

      const columnsSQL =
        inferredSchema
          .map(
            column =>
              `${safeIdentifier(
                column.column_name
              )} ${column.duckdb_type}`
          )
          .join(", ");

      await conn.query(
        `
        CREATE TABLE excel_data
        (${columnsSQL});
        `
      );

      /*
       * --------------------------------------------------
       * 8/10
       * --------------------------------------------------
       */

      setStatus(
        "8/10 — Insertando datos localmente..."
      );

      const BATCH_SIZE =
        500;

      let insertedRows =
        0;

      for (
        let start = 0;
        start < dataRows.length;
        start += BATCH_SIZE
      ) {

        const batch =
          dataRows.slice(
            start,
            start + BATCH_SIZE
          );

        const values =
          batch.map(
            row => {

              const cells =
                inferredSchema.map(
                  column =>
                    sqlValue(
                      row[
                        column.column_index
                      ],
                      column.duckdb_type
                    )
                );

              return (
                `(${cells.join(",")})`
              );
            }
          );

        if (
          values.length
        ) {

          await conn.query(
            `
            INSERT INTO excel_data
            VALUES
            ${values.join(",")};
            `
          );

          insertedRows +=
            batch.length;
        }
      }

      /*
       * --------------------------------------------------
       * 9/10
       * --------------------------------------------------
       */

      setStatus(
        "9/10 — Ejecutando SQL de prueba..."
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

      /*
       * SCHEMA
       */

      const schemaResult =
        await conn.query(
          `
          DESCRIBE excel_data;
          `
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

      /*
       * ESTADÍSTICAS NUMÉRICAS
       *
       * Generamos automáticamente
       * estadísticas para columnas
       * BIGINT y DOUBLE.
       */

      const numericColumns =
        inferredSchema.filter(
          column =>
            column.duckdb_type ===
              "BIGINT" ||
            column.duckdb_type ===
              "DOUBLE"
        );

      let numericStats =
        [];

      if (
        numericColumns.length
      ) {

        const expressions =
          numericColumns
            .map(
              column => {

                const identifier =
                  safeIdentifier(
                    column.column_name
                  );

                return `
                  COUNT(${identifier})
                    AS ${safeIdentifier(
                      `${column.column_name}__count`
                    )},

                  SUM(${identifier})
                    AS ${safeIdentifier(
                      `${column.column_name}__sum`
                    )},

                  AVG(${identifier})
                    AS ${safeIdentifier(
                      `${column.column_name}__avg`
                    )},

                  MIN(${identifier})
                    AS ${safeIdentifier(
                      `${column.column_name}__min`
                    )},

                  MAX(${identifier})
                    AS ${safeIdentifier(
                      `${column.column_name}__max`
                    )}
                `;
              }
            )
            .join(",\n");

        const statsResult =
          await conn.query(
            `
            SELECT
              ${expressions}
            FROM excel_data;
            `
          );

        numericStats =
          statsResult.toArray();
      }

      /*
       * Versión DuckDB
       */

      const versionResult =
        await conn.query(
          `
          SELECT
            version() AS duckdb_version;
          `
        );

      /*
       * --------------------------------------------------
       * 10/10
       * --------------------------------------------------
       */

      setStatus(
        "10/10 — Análisis terminado correctamente.",
        "ok"
      );

      /*
       * --------------------------------------------------
       * RESULTADO
       * --------------------------------------------------
       */

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
          sheetNames,

        hoja_principal:
          sheetName,

        filas_detectadas:
          dataRows.length,

        filas_insertadas:
          insertedRows,

        columnas_detectadas:
          headers.length,

        encabezados:
          headers,

        duckdb_table:
          "excel_data",

        duckdb_version:
          versionResult.toArray(),

        /*
         * Tipo inferido antes
         * de crear la tabla.
         */

        tipos_inferidos:
          inferredSchema,

        /*
         * Esquema real creado
         * por DuckDB.
         */

        schema_duckdb:
          schemaResult.toArray(),

        /*
         * COUNT SQL
         */

        count_sql:
          countResult.toArray(),

        /*
         * Estadísticas
         */

        estadisticas_numericas:
          numericStats,

        /*
         * Primeros 10 registros
         */

        preview:
          previewResult.toArray(),

        bundle:
          {
            mainModule:
              bundle.mainModule,

            mainWorker:
              bundle.mainWorker,

            pthreadWorker:
              bundle.pthreadWorker
          }

      });

    } catch (
      error
    ) {

      console.error(
        `[${APP_ID}]`,
        error
      );

      setStatus(
        "❌ ERROR DURANTE EL ANÁLISIS",
        "error"
      );

      setResult({

        mensaje:
          error?.message ||
          String(error),

        nombre:
          error?.name ||
          null,

        stack:
          error?.stack ||
          null

      });
    }
  }

  /*
   * ======================================================
   * INIT
   * ======================================================
   */

  function init() {

    addStyles();

    createButton();

    console.log(
      `[${APP_ID}] cargado`
    );
  }

  /*
   * ======================================================
   * START
   * ======================================================
   */

  if (
    document.readyState ===
    "loading"
  ) {

    document.addEventListener(
      "DOMContentLoaded",
      init,
      {
        once:
          true
      }
    );

  } else {

    init();

  }

})();
