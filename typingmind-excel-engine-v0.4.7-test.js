/*
 * TypingMind Excel Data Engine - Extension v0.4.7 TEST
 *
 * OBJETIVO
 * --------
 * XLSX local
 *   ↓
 * SheetJS
 *   ↓
 * eliminación de filas vacías
 *   ↓
 * inferencia de tipos
 *   ↓
 * DuckDB-Wasm
 *   ↓
 * SQL analítico
 *
 * CAMBIOS v0.4.7
 * --------------
 * - Distingue filas físicas de filas reales.
 * - Ignora filas completamente vacías.
 * - Conserva tipos DATE / BIGINT / DOUBLE / BOOLEAN / VARCHAR.
 * - Convierte fechas para mostrar YYYY-MM-DD.
 * - Agrega consultas analíticas SQL.
 * - Maneja BigInt correctamente.
 *
 * PROCESAMIENTO:
 * 100 % LOCAL
 *
 * NO utiliza:
 * - LOAD excel
 * - servidor
 * - API
 * - subida del archivo
 * - LLM
 */

(() => {

  "use strict";

  const APP_ID =
    "tm-excel-engine-v047-test";

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
        "tmxe-v047-style"
      )
    ) {
      return;
    }

    const style =
      document.createElement("style");

    style.id =
      "tmxe-v047-style";

    style.textContent = `

      #tmxe-v047-button {
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

      #tmxe-v047-overlay {
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

      #tmxe-v047-panel {
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

      #tmxe-v047-panel h2 {
        margin:
          0 0 5px;

        font-size: 20px;
      }

      #tmxe-v047-help {
        opacity: .7;
        margin-bottom: 12px;
      }

      #tmxe-v047-file {
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

      #tmxe-v047-status {
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

      #tmxe-v047-status[data-kind="ok"] {
        color: #16a34a;
      }

      #tmxe-v047-status[data-kind="error"] {
        color: #dc2626;
      }

      #tmxe-v047-result {
        white-space: pre-wrap;

        overflow: auto;

        max-height: 650px;

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

      #tmxe-v047-run,
      #tmxe-v047-close {
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

      #tmxe-v047-run {
        background:
          #2563eb !important;

        color:
          white !important;

        border-color:
          #2563eb !important;
      }

      #tmxe-v047-actions {
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
        "tmxe-v047-button"
      )
    ) {
      return;
    }

    const button =
      document.createElement("button");

    button.id =
      "tmxe-v047-button";

    button.type =
      "button";

    button.textContent =
      "📊 Excel v0.4.7";

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
        "tmxe-v047-overlay"
      )
    ) {
      return;
    }

    const overlay =
      document.createElement("div");

    overlay.id =
      "tmxe-v047-overlay";

    overlay.innerHTML = `

      <div
        id="tmxe-v047-panel"
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
              📊 Excel Engine v0.4.7 TEST
            </h2>

            <div id="tmxe-v047-help">

              Filas reales + tipos +
              fechas + SQL analítico

            </div>

          </div>

          <button
            id="tmxe-v047-close"
            type="button"
          >
            Cerrar
          </button>

        </div>

        <input
          id="tmxe-v047-file"
          type="file"
          accept=".xlsx,.xls,.xlsb,.ods"
        />

        <div
          id="tmxe-v047-actions"
        >

          <button
            id="tmxe-v047-run"
            type="button"
          >
            🚀 Analizar Excel
          </button>

        </div>

        <div
          id="tmxe-v047-status"
        >
          Selecciona un archivo Excel.
        </div>

        <strong>
          Resultado
        </strong>

        <pre
          id="tmxe-v047-result"
        >—</pre>

        <div style="
          margin-top:12px;
          opacity:.65;
          font-size:12px;
        ">

          v0.4.7 TEST —
          procesamiento completamente local.

        </div>

      </div>

    `;

    document.body.appendChild(
      overlay
    );

    document
      .getElementById(
        "tmxe-v047-close"
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
        "tmxe-v047-run"
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
        "tmxe-v047-status"
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
   * SERIALIZACIÓN SEGURA
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
        "tmxe-v047-result"
      );

    if (!el) {
      return;
    }

    if (
      typeof value ===
      "string"
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
   * SQL IDENTIFIER
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
   * NORMALIZAR ENCABEZADOS
   * ======================================================
   */

  function normalizeHeaders(
    headers
  ) {

    const result =
      [];

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
   * FILA VACÍA
   * ======================================================
   */

  function isEmptyRow(
    row
  ) {

    if (
      !Array.isArray(row)
    ) {

      return true;

    }

    return row.every(
      value =>
        value === null ||
        value === undefined ||
        value === ""
    );

  }


  /*
   * ======================================================
   * FECHA EXCEL
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
      !Number.isFinite(
        serial
      )
    ) {

      return null;

    }

    if (
      serial < 20000 ||
      serial > 80000
    ) {

      return null;

    }

    const epoch =
      Date.UTC(
        1899,
        11,
        30
      );

    const date =
      new Date(
        epoch +
        Math.round(
          serial *
          86400000
        )
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
   * ENCABEZADO FECHA
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

    return [

      "fecha",
      "date",
      "datetime",
      "timestamp",
      "día",
      "dia"

    ].some(
      pattern =>
        text.includes(
          pattern
        )
    );

  }


  /*
   * ======================================================
   * INFERENCIA DE TIPOS
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
      !nonEmpty.length
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

      const validDates =
        nonEmpty.filter(
          value =>
            excelSerialToDate(
              value
            ) !== null
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

          return (
            text !== "" &&
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

      const integers =
        numbers.every(
          number =>
            Number.isInteger(
              number
            )
        );

      if (
        integers
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
        "valores tratados como texto"

    };

  }


  /*
   * ======================================================
   * INFERIR ESQUEMA
   * ======================================================
   */

  function inferSchema(
    headers,
    rows
  ) {

    return headers.map(
      (
        header,
        columnIndex
      ) => {

        const values =
          rows.map(
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
   * CONVERSIÓN
   * ======================================================
   */

  function convertValue(
    value,
    type
  ) {

    if (
      value === null ||
      value === undefined ||
      value === ""
    ) {

      return null;

    }


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

      return Number.isFinite(
        number
      )
        ? Math.trunc(number)
        : null;

    }


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

      return Number.isFinite(
        number
      )
        ? number
        : null;

    }


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
   * LIMPIAR RESULTADOS PARA PRESENTACIÓN
   * ======================================================
   */

  function cleanResultValue(
    value,
    columnType = null
  ) {

    if (
      typeof value ===
      "bigint"
    ) {

      return value.toString();

    }

    /*
     * DuckDB DATE puede llegar
     * como Date o timestamp.
     */

    if (
      columnType === "DATE" &&
      typeof value ===
      "number"
    ) {

      const date =
        new Date(
          value
        );

      if (
        !Number.isNaN(
          date.getTime()
        )
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

    }

    return value;

  }


  function cleanRows(
    rows,
    schema
  ) {

    return rows.map(
      row => {

        const output =
          {};

        schema.forEach(
          column => {

            const name =
              column.column_name;

            let value =
              row[name];

            value =
              cleanResultValue(
                value,
                column.duckdb_type
              );

            output[name] =
              value;

          }
        );

        return output;

      }
    );

  }


  /*
   * ======================================================
   * CARGAR EXCEL
   * ======================================================
   */

  async function loadExcel() {

    try {

      setResult("");

      const fileInput =
        document.getElementById(
          "tmxe-v047-file"
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
       * 1
       */

      setStatus(
        "1/12 — Cargando SheetJS..."
      );

      const XLSX =
        await loadSheetJS();


      /*
       * 2
       */

      setStatus(
        "2/12 — Leyendo archivo local..."
      );

      const arrayBuffer =
        await file.arrayBuffer();


      /*
       * 3
       */

      setStatus(
        "3/12 — Analizando estructura XLSX..."
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


      const sheetName =
        sheetNames[0];

      const worksheet =
        workbook.Sheets[
          sheetName
        ];


      /*
       * 4
       */

      setStatus(
        "4/12 — Extrayendo filas..."
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


      /*
       * Primera fila
       */

      const headers =
        normalizeHeaders(
          rows[0]
        );


      /*
       * Todas las filas físicas
       */

      const physicalRows =
        rows.slice(1);


      /*
       * Filas reales
       */

      const realRows =
        physicalRows.filter(
          row =>
            !isEmptyRow(
              row
            )
        );


      const emptyRows =
        physicalRows.length -
        realRows.length;


      /*
       * 5
       */

      setStatus(
        "5/12 — Filas físicas: " +
        physicalRows.length +
        "\n" +
        "Filas reales: " +
        realRows.length +
        "\n" +
        "Filas vacías ignoradas: " +
        emptyRows
      );


      /*
       * 6
       */

      setStatus(
        "6/12 — Detectando tipos..."
      );

      const inferredSchema =
        inferSchema(
          headers,
          realRows
        );


      /*
       * 7
       */

      setStatus(
        "7/12 — Inicializando DuckDB-Wasm..."
      );

      const bundle =
        await initDuckDB();


      /*
       * 8
       */

      setStatus(
        "8/12 — Creando tabla DuckDB..."
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
          .join(",");


      await conn.query(
        `
        CREATE TABLE excel_data
        (${columnsSQL});
        `
      );


      /*
       * 9
       */

      setStatus(
        "9/12 — Insertando datos localmente..."
      );


      const BATCH_SIZE =
        500;


      let insertedRows =
        0;


      for (
        let start = 0;
        start < realRows.length;
        start += BATCH_SIZE
      ) {

        const batch =
          realRows.slice(
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
                "(" +
                cells.join(",") +
                ")"
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
       * 10
       */

      setStatus(
        "10/12 — Ejecutando consultas SQL..."
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
       * DESCRIBE
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
       * RESUMEN
       */

      const numericColumns =
        inferredSchema.filter(
          column =>
            column.duckdb_type ===
              "BIGINT" ||
            column.duckdb_type ===
              "DOUBLE"
        );


      let resumenSQL =
        null;

      let resumen =
        [];


      if (
        numericColumns.length
      ) {

        const expressions =
          numericColumns
            .map(
              column => {

                const id =
                  safeIdentifier(
                    column.column_name
                  );

                return `

                  COUNT(${id})
                    AS ${safeIdentifier(
                      column.column_name +
                      "__count"
                    )},

                  SUM(${id})
                    AS ${safeIdentifier(
                      column.column_name +
                      "__sum"
                    )},

                  AVG(${id})
                    AS ${safeIdentifier(
                      column.column_name +
                      "__avg"
                    )},

                  MIN(${id})
                    AS ${safeIdentifier(
                      column.column_name +
                      "__min"
                    )},

                  MAX(${id})
                    AS ${safeIdentifier(
                      column.column_name +
                      "__max"
                    )}

                `;

              }
            )
            .join(",");


        resumenSQL = `
          SELECT
            ${expressions}
          FROM excel_data;
        `;


        const resumenResult =
          await conn.query(
            resumenSQL
          );


        resumen =
          resumenResult.toArray();

      }


      /*
       * 11
       *
       * AGRUPACIÓN TEMPORAL
       */

      let agrupacionAnual =
        [];

      const dateColumn =
        inferredSchema.find(
          column =>
            column.duckdb_type ===
            "DATE"
        );


      if (
        dateColumn
      ) {

        const dateId =
          safeIdentifier(
            dateColumn.column_name
          );


        const totalColumn =
          numericColumns.length
            ? safeIdentifier(
                numericColumns[
                  numericColumns.length - 1
                ].column_name
              )
            : null;


        if (
          totalColumn
        ) {

          const yearResult =
            await conn.query(
              `
              SELECT

                EXTRACT(
                  YEAR FROM
                  ${dateId}
                ) AS año,

                COUNT(*) AS registros,

                SUM(
                  ${totalColumn}
                ) AS total

              FROM excel_data

              GROUP BY año

              ORDER BY año;
              `
            );


          agrupacionAnual =
            yearResult.toArray();

        }

      }


      /*
       * 12
       */

      setStatus(
        "12/12 — Análisis terminado correctamente.",
        "ok"
      );


      /*
       * Limpiar preview
       */

      const preview =
        cleanRows(
          previewResult.toArray(),
          inferredSchema
        );


      /*
       * RESULTADO FINAL
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


        /*
         * FILAS
         */

        filas_fisicas_detectadas:
          physicalRows.length,

        filas_reales:
          realRows.length,

        filas_vacias_ignoradas:
          emptyRows,

        filas_insertadas:
          insertedRows,


        /*
         * COLUMNAS
         */

        columnas_detectadas:
          headers.length,

        encabezados:
          headers,


        /*
         * DUCKDB
         */

        duckdb_table:
          "excel_data",

        duckdb_version:
          (
            await conn.query(
              `
              SELECT
                version()
                AS duckdb_version;
              `
            )
          ).toArray(),


        /*
         * TIPOS
         */

        tipos_inferidos:
          inferredSchema,

        schema_duckdb:
          schemaResult.toArray(),


        /*
         * COUNT
         */

        count_sql:
          countResult.toArray(),


        /*
         * ESTADÍSTICAS
         */

        resumen_sql:
          resumenSQL,

        resumen:
          resumen,


        /*
         * AGRUPACIÓN
         */

        agrupacion_anual:
          agrupacionAnual,


        /*
         * PREVIEW
         */

        preview:
          preview,


        /*
         * BUNDLE
         */

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
