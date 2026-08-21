/*
 * TypingMind Excel Data Engine
 * Extension v0.4.16 TEST
 *
 * CAMBIOS v0.4.16
 *
 * - Mantiene todas las correcciones de v0.4.15.
 *
 * - CORRECCIÓN PRINCIPAL:
 *   Normalización robusta de resultados DuckDB-Wasm / Arrow.
 *
 *   Corrige resultados que pueden aparecer como:
 *
 *      [948618, 0, 0, 0]
 *
 *   y los convierte en:
 *
 *      948618
 *
 * - La normalización ya NO depende del nombre
 *   del campo para determinar si un array debe
 *   convertirse en escalar.
 *
 * - Maneja:
 *   Array
 *   TypedArray
 *   Uint8Array
 *   Int32Array
 *   BigInt64Array
 *   Float64Array
 *   DataView
 *   objetos Arrow con toArray()
 *   objetos con length
 *
 * - Protege especialmente resultados escalares
 *   provenientes de SUM, AVG, MIN, MAX, COUNT
 *   y expresiones matemáticas.
 *
 * - Mantiene DATE como DATE.
 *
 * - Mantiene:
 *   Año             -> BIGINT
 *   Local           -> BIGINT
 *   Transshipment   -> BIGINT
 *   Total TEUs      -> BIGINT
 *
 * - Detecta variantes:
 *   Total TEUs
 *   Total TEU's
 *   Total TEU´s
 *   Total TEU
 *   Total
 *
 * - information_schema.columns utiliza data_type.
 *
 * - Procesamiento 100% LOCAL.
 *
 * - DuckDB-Wasm 1.29.0
 * - DuckDB interno 1.1.1
 * - SheetJS 0.18.5
 */

(() => {
  "use strict";

  const APP_ID =
    "tm-excel-engine-v0416-test";

  const VERSION =
    "v0.4.16-test";

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
  let currentInferredTypes = [];


  /* =========================================================
   * UTILIDADES GENERALES
   * ========================================================= */

  function safeString(value) {
    return value === null ||
      value === undefined
      ? ""
      : String(value);
  }


  function isEmptyValue(value) {
    return (
      value === null ||
      value === undefined ||
      (
        typeof value === "string" &&
        value.trim() === ""
      )
    );
  }


  function normalizeHeaderName(value) {
    return safeString(value)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/[’'´`]/g, "")
      .replace(/[^a-z0-9]+/g, "")
      .trim();
  }


  /* =========================================================
   * v0.4.16
   * NORMALIZACIÓN DUCKDB / ARROW
   * ========================================================= */

  /*
   * Determina si un valor es un TypedArray.
   *
   * Se excluyen:
   *   DataView
   *   ArrayBuffer
   */
  function isTypedArray(value) {
    return (
      value !== null &&
      value !== undefined &&
      ArrayBuffer.isView(value) &&
      !(value instanceof DataView)
    );
  }


  /*
   * Convierte diferentes estructuras binarias
   * o Arrow a un array JavaScript normal.
   */
  function tryConvertToArray(value) {

    if (Array.isArray(value)) {
      return value.slice();
    }

    if (isTypedArray(value)) {
      try {
        return Array.from(value);
      } catch (_) {
        return null;
      }
    }

    return null;
  }


  /*
   * Detecta el patrón que DuckDB-Wasm puede devolver
   * para una única celda:
   *
   * [valor, 0, 0, 0]
   *
   * El cuarto elemento puede representar metadatos,
   * padding o información interna del vector.
   *
   * IMPORTANTE:
   *
   * No asumimos que todos los arrays de 4 elementos
   * son escalares.
   *
   * Solo se convierte cuando:
   *
   *   - length === 4
   *   - primer valor es numérico
   *   - los restantes son cero
   *
   * También se acepta BigInt.
   */
  function normalizeFourElementScalar(value) {

    const array =
      tryConvertToArray(value);

    if (!array) {
      return {
        matched: false,
        value: null
      };
    }

    if (array.length !== 4) {
      return {
        matched: false,
        value: null
      };
    }

    const first =
      array[0];

    const rest =
      array.slice(1);

    const firstIsNumeric =
      typeof first === "number" ||
      typeof first === "bigint" ||
      (
        typeof first === "string" &&
        first.trim() !== "" &&
        Number.isFinite(
          Number(first)
        )
      );

    if (!firstIsNumeric) {
      return {
        matched: false,
        value: null
      };
    }

    const restAreZero =
      rest.every(
        (v) => {

          if (
            typeof v === "bigint"
          ) {
            return v === 0n;
          }

          if (
            typeof v === "number"
          ) {
            return v === 0;
          }

          if (
            typeof v === "string"
          ) {
            return (
              v.trim() === "0"
            );
          }

          return false;
        }
      );

    if (!restAreZero) {
      return {
        matched: false,
        value: null
      };
    }

    return {
      matched: true,
      value: first
    };
  }


  /*
   * Convierte objetos que puedan comportarse
   * como vectores Arrow.
   */
  function tryNormalizeArrowLike(value) {

    if (
      value === null ||
      value === undefined
    ) {
      return {
        matched: false,
        value: null
      };
    }

    /*
     * Algunos objetos Arrow exponen toArray().
     */
    if (
      typeof value.toArray === "function"
    ) {
      try {

        const converted =
          value.toArray();

        const scalar =
          normalizeFourElementScalar(
            converted
          );

        if (scalar.matched) {
          return scalar;
        }

        /*
         * Si toArray devuelve un único
         * elemento, podemos utilizarlo.
         */
        if (
          Array.isArray(converted) &&
          converted.length === 1
        ) {
          return {
            matched: true,
            value: converted[0]
          };
        }

      } catch (_) {
        // Continuar.
      }
    }

    /*
     * Algunos wrappers pueden exponer
     * length + acceso numérico.
     */
    if (
      typeof value.length === "number" &&
      value.length === 4
    ) {

      try {

        const array = [];

        for (
          let i = 0;
          i < value.length;
          i++
        ) {
          array.push(
            value[i]
          );
        }

        const scalar =
          normalizeFourElementScalar(
            array
          );

        if (scalar.matched) {
          return scalar;
        }

      } catch (_) {
        // Continuar.
      }
    }

    return {
      matched: false,
      value: null
    };
  }


  /*
   * Normaliza cualquier valor proveniente
   * de DuckDB-Wasm.
   */
  function normalizeValue(
    value,
    key = ""
  ) {

    if (
      value === null ||
      value === undefined
    ) {
      return null;
    }


    /*
     * BigInt
     */
    if (
      typeof value === "bigint"
    ) {

      const n =
        Number(value);

      return Number.isSafeInteger(n)
        ? n
        : value.toString();
    }


    /*
     * Date
     */
    if (
      value instanceof Date
    ) {

      return Number.isNaN(
        value.getTime()
      )
        ? null
        : value.toISOString();
    }


    /*
     * ArrayBuffer
     */
    if (
      value instanceof ArrayBuffer
    ) {

      return Array.from(
        new Uint8Array(value)
      );
    }


    /*
     * TypedArray / DataView
     */
    if (
      ArrayBuffer.isView(value)
    ) {

      /*
       * DataView
       */
      if (
        value instanceof DataView
      ) {
        return value;
      }

      const array =
        Array.from(value);

      /*
       * PRIMERA PRIORIDAD:
       * detectar [valor,0,0,0].
       */
      const scalar =
        normalizeFourElementScalar(
          array
        );

      if (scalar.matched) {

        return normalizeValue(
          scalar.value,
          key
        );
      }

      return array.map(
        (v) =>
          normalizeValue(
            v,
            key
          )
      );
    }


    /*
     * Array normal
     */
    if (
      Array.isArray(value)
    ) {

      /*
       * PRIMERA PRIORIDAD:
       * detectar [valor,0,0,0].
       */
      const scalar =
        normalizeFourElementScalar(
          value
        );

      if (scalar.matched) {

        return normalizeValue(
          scalar.value,
          key
        );
      }

      return value.map(
        (v) =>
          normalizeValue(
            v,
            key
          )
      );
    }


    /*
     * Objetos Arrow-like.
     */
    if (
      typeof value === "object"
    ) {

      const arrow =
        tryNormalizeArrowLike(
          value
        );

      if (arrow.matched) {

        return normalizeValue(
          arrow.value,
          key
        );
      }

      const output = {};

      for (
        const [k, v]
        of Object.entries(value)
      ) {

        output[k] =
          normalizeValue(
            v,
            k
          );
      }

      return output;
    }


    return value;
  }


  function normalizeRows(rows) {

    if (
      !Array.isArray(rows)
    ) {
      return [];
    }

    return rows.map(
      (row) =>
        normalizeValue(row)
    );
  }


  function jsonSafe(value) {

    return JSON.parse(
      JSON.stringify(
        value,
        (_, v) => {

          if (
            typeof v === "bigint"
          ) {

            const n =
              Number(v);

            return Number.isSafeInteger(n)
              ? n
              : v.toString();
          }

          return v;
        }
      )
    );
  }


  function setResult(value) {

    const safe =
      jsonSafe(
        normalizeValue(value)
      );

    const output =
      document.getElementById(
        "tm-excel-result"
      );

    if (!output) {
      return;
    }

    output.textContent =
      JSON.stringify(
        safe,
        null,
        2
      );
  }


  function showError(error) {

    const message =
      error &&
      error.message
        ? error.message
        : String(error);

    const stack =
      error &&
      error.stack
        ? error.stack
        : "";

    setResult({
      mensaje: message,
      stack: stack
    });
  }


  /* =========================================================
   * SQL
   * ========================================================= */

  function escapeSqlIdentifier(name) {

    return (
      '"' +
      safeString(name)
        .replace(
          /"/g,
          '""'
        ) +
      '"'
    );
  }


  function escapeSqlString(value) {

    return safeString(value)
      .replace(
        /'/g,
        "''"
      );
  }


  /* =========================================================
   * ENCABEZADOS
   * ========================================================= */

  function makeUniqueHeaders(headers) {

    const result = [];
    const used = new Map();

    headers.forEach(
      (
        header,
        index
      ) => {

        let name =
          safeString(header)
            .trim();

        if (!name) {
          name =
            `Columna_${index + 1}`;
        }

        const normalized =
          name.toLowerCase();

        if (
          !used.has(normalized)
        ) {

          used.set(
            normalized,
            1
          );

          result.push(name);

          return;
        }

        const count =
          used.get(
            normalized
          ) + 1;

        used.set(
          normalized,
          count
        );

        result.push(
          `${name}_${count}`
        );
      }
    );

    return result;
  }


  function detectColumn(
    headers,
    candidates
  ) {

    if (
      !Array.isArray(headers)
    ) {
      return null;
    }

    const normalizedHeaders =
      headers.map(
        (h) =>
          normalizeHeaderName(h)
      );

    for (
      const candidate
      of candidates
    ) {

      const target =
        normalizeHeaderName(
          candidate
        );

      const index =
        normalizedHeaders.indexOf(
          target
        );

      if (index >= 0) {
        return headers[index];
      }
    }

    return null;
  }


  function detectSpecialColumns() {

    const yearColumn =
      detectColumn(
        currentHeaders,
        [
          "Año",
          "Ano",
          "Year"
        ]
      );

    const dateColumn =
      detectColumn(
        currentHeaders,
        [
          "Fecha",
          "Date"
        ]
      );

    const totalColumn =
      detectColumn(
        currentHeaders,
        [
          "Total TEUs",
          "Total TEU's",
          "Total TEU´s",
          "Total TEU",
          "Total"
        ]
      );

    const localColumn =
      detectColumn(
        currentHeaders,
        [
          "Local"
        ]
      );

    const transshipmentColumn =
      detectColumn(
        currentHeaders,
        [
          "Transshipment",
          "Transhipment",
          "Transshipment TEUs"
        ]
      );

    const monthColumn =
      detectColumn(
        currentHeaders,
        [
          "Mes",
          "Month"
        ]
      );

    return {
      año: yearColumn,
      fecha: dateColumn,
      mes: monthColumn,
      total_teus: totalColumn,
      local: localColumn,
      transshipment:
        transshipmentColumn
    };
  }


  /* =========================================================
   * INFERENCIA DE TIPOS
   * ========================================================= */

  function isDateHeader(name) {

    const normalized =
      normalizeHeaderName(name);

    return [
      "fecha",
      "date",
      "fechacombinada",
      "fechaevento"
    ].includes(
      normalized
    );
  }


  function looksLikeDateString(value) {

    if (
      value instanceof Date
    ) {
      return true;
    }

    if (
      typeof value === "number" ||
      typeof value === "bigint"
    ) {
      return false;
    }

    const text =
      safeString(value)
        .trim();

    if (!text) {
      return false;
    }

    if (
      /^\d{4}-\d{1,2}-\d{1,2}$/
        .test(text)
    ) {
      return true;
    }

    if (
      /^\d{1,2}[\/-]\d{1,2}[\/-]\d{4}$/
        .test(text)
    ) {
      return true;
    }

    if (
      /^[A-Za-zÁÉÍÓÚáéíóú]+[\/-]\d{4}$/
        .test(text)
    ) {
      return true;
    }

    return false;
  }


  function looksLikeInteger(value) {

    if (
      typeof value === "bigint"
    ) {
      return true;
    }

    if (
      typeof value === "number"
    ) {
      return Number.isInteger(value);
    }

    const text =
      safeString(value)
        .trim();

    if (!text) {
      return false;
    }

    return /^-?\d+$/
      .test(
        text.replace(
          /,/g,
          ""
        )
      );
  }


  function looksLikeDouble(value) {

    if (
      typeof value === "number"
    ) {
      return Number.isFinite(value);
    }

    const text =
      safeString(value)
        .replace(/,/g, "")
        .trim();

    if (!text) {
      return false;
    }

    return Number.isFinite(
      Number(text)
    );
  }


  function inferColumnType(
    columnIndex,
    columnName,
    rows
  ) {

    const values =
      rows
        .map(
          (row) =>
            row[columnIndex]
        )
        .filter(
          (v) =>
            !isEmptyValue(v)
        );

    const nonEmptyValues =
      values.length;


    if (
      isDateHeader(columnName)
    ) {

      return {
        column_index:
          columnIndex,
        column_name:
          columnName,
        duckdb_type:
          "DATE",
        sqlType:
          "DATE",
        confidence:
          "high",
        reason:
          "encabezado identificado como fecha",
        non_empty_values:
          nonEmptyValues
      };
    }


    const normalizedName =
      normalizeHeaderName(
        columnName
      );


    if (
      normalizedName === "ano" ||
      normalizedName === "year"
    ) {

      return {
        column_index:
          columnIndex,
        column_name:
          columnName,
        duckdb_type:
          "BIGINT",
        sqlType:
          "BIGINT",
        confidence:
          "high",
        reason:
          "columna identificada como año",
        non_empty_values:
          nonEmptyValues
      };
    }


    if (
      [
        "local",
        "transshipment",
        "transhipment",
        "totalteus",
        "totalteu",
        "total"
      ].includes(
        normalizedName
      )
    ) {

      const allIntegers =
        values.length > 0 &&
        values.every(
          looksLikeInteger
        );

      if (allIntegers) {

        return {
          column_index:
            columnIndex,
          column_name:
            columnName,
          duckdb_type:
            "BIGINT",
          sqlType:
            "BIGINT",
          confidence:
            "high",
          reason:
            "columna numérica identificada por encabezado",
          non_empty_values:
            nonEmptyValues
        };
      }

      const allNumeric =
        values.length > 0 &&
        values.every(
          looksLikeDouble
        );

      if (allNumeric) {

        return {
          column_index:
            columnIndex,
          column_name:
            columnName,
          duckdb_type:
            "DOUBLE",
          sqlType:
            "DOUBLE",
          confidence:
            "high",
          reason:
            "columna numérica identificada por encabezado",
          non_empty_values:
            nonEmptyValues
        };
      }
    }


    if (
      values.length > 0 &&
      values.every(
        looksLikeDateString
      )
    ) {

      return {
        column_index:
          columnIndex,
        column_name:
          columnName,
        duckdb_type:
          "DATE",
        sqlType:
          "DATE",
        confidence:
          "medium",
        reason:
          "todos los valores son compatibles con fecha",
        non_empty_values:
          nonEmptyValues
      };
    }


    if (
      values.length > 0 &&
      values.every(
        looksLikeInteger
      )
    ) {

      return {
        column_index:
          columnIndex,
        column_name:
          columnName,
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
        looksLikeDouble
      )
    ) {

      return {
        column_index:
          columnIndex,
        column_name:
          columnName,
        duckdb_type:
          "DOUBLE",
        sqlType:
          "DOUBLE",
        confidence:
          "medium",
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
        columnName,
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


  /* =========================================================
   * FECHAS
   * ========================================================= */

  function excelSerialToDate(
    serial
  ) {

    const n =
      Number(serial);

    if (
      !Number.isFinite(n)
    ) {
      return null;
    }

    const excelEpoch =
      Date.UTC(
        1899,
        11,
        30
      );

    const millis =
      excelEpoch +
      Math.round(
        n * 86400000
      );

    return new Date(
      millis
    );
  }


  function normalizeExcelDate(
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
      value < 80000
    ) {

      return excelSerialToDate(
        value
      );
    }

    const text =
      safeString(value)
        .trim();

    if (!text) {
      return null;
    }

    if (
      /^\d{4}-\d{1,2}-\d{1,2}$/
        .test(text)
    ) {

      const d =
        new Date(
          `${text}T00:00:00Z`
        );

      return Number.isNaN(
        d.getTime()
      )
        ? null
        : d;
    }

    const d =
      new Date(text);

    return Number.isNaN(
      d.getTime()
    )
      ? null
      : d;
  }


  function dateToSql(date) {

    if (
      !(date instanceof Date)
    ) {
      return null;
    }

    const y =
      date.getUTCFullYear();

    const m =
      String(
        date.getUTCMonth() + 1
      ).padStart(
        2,
        "0"
      );

    const d =
      String(
        date.getUTCDate()
      ).padStart(
        2,
        "0"
      );

    return `${y}-${m}-${d}`;
  }


  /* =========================================================
   * CONVERSIÓN SQL
   * ========================================================= */

  function valueToSql(
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
        normalizeExcelDate(
          value
        );

      if (!date) {
        return "NULL";
      }

      return `DATE '${dateToSql(
        date
      )}'`;
    }


    if (
      type === "BIGINT"
    ) {

      const text =
        safeString(value)
          .replace(
            /,/g,
            ""
          )
          .trim();

      const n =
        Number(text);

      if (
        !Number.isFinite(n)
      ) {
        return "NULL";
      }

      return String(
        Math.trunc(n)
      );
    }


    if (
      type === "DOUBLE"
    ) {

      const text =
        safeString(value)
          .replace(
            /,/g,
            ""
          )
          .trim();

      const n =
        Number(text);

      if (
        !Number.isFinite(n)
      ) {
        return "NULL";
      }

      return String(n);
    }


    return `'${escapeSqlString(
      value
    )}'`;
  }


  function buildCreateTableSql(
    headers,
    inferredTypes
  ) {

    const columns =
      headers.map(
        (
          header,
          index
        ) => {

          const type =
            inferredTypes[index]
              ?.sqlType ||
            "VARCHAR";

          return (
            escapeSqlIdentifier(
              header
            ) +
            " " +
            type
          );
        }
      );

    return `
      CREATE TABLE excel_data (
        ${columns.join(
          ",\n"
        )}
      );
    `;
  }


  function buildInsertSql(
    headers,
    rows,
    inferredTypes
  ) {

    const columnSql =
      headers
        .map(
          escapeSqlIdentifier
        )
        .join(", ");

    const statements = [];

    for (
      const row
      of rows
    ) {

      const values =
        headers.map(
          (
            _,
            index
          ) =>
            valueToSql(
              row[index],
              inferredTypes[index]
                ?.sqlType ||
              "VARCHAR"
            )
        );

      statements.push(
        `(${values.join(
          ", "
        )})`
      );
    }

    return `
      INSERT INTO excel_data
      (${columnSql})
      VALUES
      ${statements.join(
        ",\n"
      )};
    `;
  }


  /* =========================================================
   * DUCKDB
   * ========================================================= */

  async function loadLibraries() {

    if (!duckdb) {

      duckdb =
        await import(
          DUCKDB_PACKAGE
        );
    }

    if (!XLSX) {

      XLSX =
        await import(
          XLSX_PACKAGE
        );
    }
  }


  async function createDuckDB() {

    if (
      db &&
      conn
    ) {
      return;
    }

    const bundles =
      duckdb.getJsDelivrBundles();

    const bundle =
      bundles.eh;

    workerURL =
      URL.createObjectURL(
        new Blob(
          [
            `importScripts("${bundle.mainWorker}");`
          ],
          {
            type:
              "application/javascript"
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
  }


  async function resetDatabase() {

    if (!conn) {
      return;
    }

    try {

      await conn.query(
        "DROP TABLE IF EXISTS excel_data;"
      );

    } catch (_) {
      // Ignorar.
    }
  }


  async function queryRows(sql) {

    const result =
      await conn.query(
        sql
      );

    const rows =
      result.toArray();

    /*
     * v0.4.16:
     *
     * La normalización se realiza
     * inmediatamente después de convertir
     * el resultado Arrow a array.
     */
    return normalizeRows(
      rows
    );
  }


  async function getDuckDBVersion() {

    try {

      return await queryRows(
        "SELECT * FROM pragma_version();"
      );

    } catch (error) {

      return [
        {
          duckdb_version_error:
            error.message
        }
      ];
    }
  }


  async function getSchema() {

    return queryRows(`
      SELECT
        column_name,
        data_type,
        column_default,
        scope_name,
        collation_name
      FROM information_schema.columns
      WHERE table_name = 'excel_data'
      ORDER BY ordinal_position;
    `);
  }


  /* =========================================================
   * EXCEL
   * ========================================================= */

  function readWorksheet(
    worksheet
  ) {

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

    if (
      !matrix.length
    ) {

      return {
        headers: [],
        rows: [],
        physicalRows: 0,
        realRows: 0,
        emptyRows: 0
      };
    }

    const rawHeaders =
      matrix[0] || [];

    const headers =
      makeUniqueHeaders(
        rawHeaders
      );

    const data =
      matrix.slice(1);

    const normalizedRows =
      data.map(
        (row) => {

          const result = [];

          for (
            let i = 0;
            i < headers.length;
            i++
          ) {

            result.push(
              row?.[i] ?? null
            );
          }

          return result;
        }
      );

    const realRows =
      normalizedRows.filter(
        (row) =>
          row.some(
            (value) =>
              !isEmptyValue(value)
          )
      );

    const emptyRows =
      normalizedRows.length -
      realRows.length;

    return {
      headers,
      rows:
        realRows,
      physicalRows:
        matrix.length,
      realRows:
        realRows.length,
      emptyRows
    };
  }


  async function loadExcel(
    file
  ) {

    try {

      await loadLibraries();

      await createDuckDB();

      await resetDatabase();

      currentFile =
        file;

      const buffer =
        await file.arrayBuffer();

      const workbook =
        XLSX.read(
          buffer,
          {
            type: "array",
            cellDates: true,
            raw: true
          }
        );

      currentWorkbook =
        workbook;

      const sheetNames =
        workbook.SheetNames ||
        [];

      if (
        !sheetNames.length
      ) {

        throw new Error(
          "El archivo Excel no contiene hojas."
        );
      }

      currentSheetName =
        sheetNames[0];

      const worksheet =
        workbook.Sheets[
          currentSheetName
        ];

      const parsed =
        readWorksheet(
          worksheet
        );

      currentHeaders =
        parsed.headers;

      currentRows =
        parsed.rows;

      if (
        !currentHeaders.length
      ) {

        throw new Error(
          "No se encontraron encabezados."
        );
      }

      if (
        !currentRows.length
      ) {

        throw new Error(
          "No se encontraron filas con datos."
        );
      }

      currentInferredTypes =
        currentHeaders.map(
          (
            header,
            index
          ) =>
            inferColumnType(
              index,
              header,
              currentRows
            )
        );

      const createSql =
        buildCreateTableSql(
          currentHeaders,
          currentInferredTypes
        );

      await conn.query(
        createSql
      );

      const insertSql =
        buildInsertSql(
          currentHeaders,
          currentRows,
          currentInferredTypes
        );

      await conn.query(
        insertSql
      );

      const version =
        await getDuckDBVersion();

      const schema =
        await getSchema();

      const count =
        await queryRows(`
          SELECT
            COUNT(*) AS registros
          FROM excel_data;
        `);

      const preview =
        await queryRows(`
          SELECT *
          FROM excel_data
          LIMIT 10;
        `);

      const specialColumns =
        detectSpecialColumns();

      const result = {

        procesamiento:
          "LOCAL",

        engine:
          APP_ID,

        version:
          VERSION,

        archivo:
          file.name,

        tamano_bytes:
          file.size,

        hojas:
          sheetNames,

        hoja_principal:
          currentSheetName,

        filas_fisicas_detectadas:
          parsed.physicalRows,

        filas_reales:
          parsed.realRows,

        filas_vacias_ignoradas:
          parsed.emptyRows,

        filas_insertadas:
          currentRows.length,

        columnas_detectadas:
          currentHeaders.length,

        encabezados:
          currentHeaders,

        tipos_inferidos:
          currentInferredTypes,

        columnas_especiales_detectadas:
          specialColumns,

        duckdb_table:
          "excel_data",

        duckdb_version:
          version,

        count:
          count,

        schema:
          schema,

        preview:
          preview,

        bundle: {

          mainModule:
            "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/dist/duckdb-eh.wasm",

          mainWorker:
            "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/dist/duckdb-browser-eh.worker.js",

          pthreadWorker:
            null
        }
      };

      setResult(
        result
      );

      updateStatus(
        `Excel cargado localmente: ${file.name} — ${currentRows.length} registros`
      );

    } catch (error) {

      showError(
        error
      );

      updateStatus(
        "Error al cargar Excel"
      );
    }
  }


  /* =========================================================
   * EJECUCIÓN SQL
   * ========================================================= */

  async function executeSql(
    sql,
    label = "SQL"
  ) {

    try {

      if (!conn) {

        throw new Error(
          "Primero debes cargar un archivo Excel."
        );
      }

      if (
        !sql ||
        !sql.trim()
      ) {

        throw new Error(
          "Introduce una consulta SQL."
        );
      }

      const start =
        performance.now();

      const rows =
        await queryRows(
          sql
        );

      const elapsed =
        performance.now() -
        start;

      setResult({

        procesamiento:
          "LOCAL",

        sql:
          sql,

        filas_resultado:
          rows.length,

        tiempo_ms:
          Number(
            elapsed.toFixed(
              1
            )
          ),

        resultado:
          rows
      });

      updateStatus(
        `${label} ejecutado localmente`
      );

      return rows;

    } catch (error) {

      showError(
        error
      );

      updateStatus(
        `Error en ${label}`
      );

      throw error;
    }
  }


  /* =========================================================
   * CONSULTAS PREDEFINIDAS
   * ========================================================= */

  async function countRows() {

    return executeSql(
      `
      SELECT
        COUNT(*) AS registros
      FROM excel_data;
      `,
      "COUNT"
    );
  }


  async function previewRows() {

    return executeSql(
      `
      SELECT *
      FROM excel_data
      LIMIT 10;
      `,
      "Vista previa"
    );
  }


  async function summary() {

    try {

      if (!conn) {

        throw new Error(
          "Primero debes cargar un archivo Excel."
        );
      }

      const schema =
        await getSchema();

      const numericColumns =
        schema.filter(
          (column) =>
            /BIGINT|DOUBLE|DECIMAL|INTEGER|HUGEINT|FLOAT/i
              .test(
                safeString(
                  column.data_type
                )
              )
        );

      if (
        !numericColumns.length
      ) {

        throw new Error(
          "No se encontraron columnas numéricas."
        );
      }

      const expressions =
        [];

      for (
        const column
        of numericColumns
      ) {

        const name =
          column.column_name;

        const q =
          escapeSqlIdentifier(
            name
          );

        expressions.push(
          `SUM(${q}) AS ${escapeSqlIdentifier(
            `${name}__sum`
          )}`
        );

        expressions.push(
          `AVG(${q}) AS ${escapeSqlIdentifier(
            `${name}__avg`
          )}`
        );

        expressions.push(
          `MIN(${q}) AS ${escapeSqlIdentifier(
            `${name}__min`
          )}`
        );

        expressions.push(
          `MAX(${q}) AS ${escapeSqlIdentifier(
            `${name}__max`
          )}`
        );
      }

      const sql = `
        SELECT
          ${expressions.join(
            ",\n"
          )}
        FROM excel_data;
      `;

      return executeSql(
        sql,
        "Resumen"
      );

    } catch (error) {

      showError(
        error
      );

      updateStatus(
        "Error al generar resumen"
      );
    }
  }


  async function totalByYear() {

    try {

      if (!conn) {

        throw new Error(
          "Primero debes cargar un archivo Excel."
        );
      }

      const special =
        detectSpecialColumns();

      if (!special.año) {

        throw new Error(
          `No se encontró la columna de año. Columnas disponibles: ${currentHeaders.join(
            ", "
          )}`
        );
      }

      if (!special.total_teus) {

        throw new Error(
          `No se encontró la columna de Total TEUs. Columnas disponibles: ${currentHeaders.join(
            ", "
          )}`
        );
      }

      const yearSql =
        escapeSqlIdentifier(
          special.año
        );

      const totalSql =
        escapeSqlIdentifier(
          special.total_teus
        );

      const sql = `
        SELECT
          ${yearSql} AS año,
          SUM(${totalSql}) AS total
        FROM excel_data
        GROUP BY ${yearSql}
        ORDER BY ${yearSql};
      `;

      return executeSql(
        sql,
        "Total por año"
      );

    } catch (error) {

      showError(
        error
      );

      updateStatus(
        "Error en Total por año"
      );
    }
  }


  async function totalByMonth() {

    try {

      if (!conn) {

        throw new Error(
          "Primero debes cargar un archivo Excel."
        );
      }

      const special =
        detectSpecialColumns();

      if (!special.año) {

        throw new Error(
          "No se encontró la columna Año."
        );
      }

      if (!special.mes) {

        throw new Error(
          "No se encontró la columna Mes."
        );
      }

      if (!special.total_teus) {

        throw new Error(
          "No se encontró la columna Total TEUs."
        );
      }

      const yearSql =
        escapeSqlIdentifier(
          special.año
        );

      const monthSql =
        escapeSqlIdentifier(
          special.mes
        );

      const totalSql =
        escapeSqlIdentifier(
          special.total_teus
        );

      const sql = `
        SELECT
          ${yearSql} AS año,
          ${monthSql} AS mes,
          SUM(${totalSql}) AS total
        FROM excel_data
        GROUP BY
          ${yearSql},
          ${monthSql}
        ORDER BY
          ${yearSql},
          ${monthSql};
      `;

      return executeSql(
        sql,
        "Total por mes"
      );

    } catch (error) {

      showError(
        error
      );

      updateStatus(
        "Error en Total por mes"
      );
    }
  }


  async function totalByLocal() {

    try {

      if (!conn) {

        throw new Error(
          "Primero debes cargar un archivo Excel."
        );
      }

      const special =
        detectSpecialColumns();

      if (!special.local) {

        throw new Error(
          "No se encontró la columna Local."
        );
      }

      if (!special.total_teus) {

        throw new Error(
          "No se encontró la columna Total TEUs."
        );
      }

      const localSql =
        escapeSqlIdentifier(
          special.local
        );

      const totalSql =
        escapeSqlIdentifier(
          special.total_teus
        );

      const sql = `
        SELECT
          ${localSql} AS local,
          SUM(${totalSql}) AS total
        FROM excel_data
        GROUP BY ${localSql}
        ORDER BY total DESC;
      `;

      return executeSql(
        sql,
        "Total por Local"
      );

    } catch (error) {

      showError(
        error
      );

      updateStatus(
        "Error en Total por Local"
      );
    }
  }


  async function totalTransshipmentByYear() {

    try {

      if (!conn) {

        throw new Error(
          "Primero debes cargar un archivo Excel."
        );
      }

      const special =
        detectSpecialColumns();

      if (!special.año) {

        throw new Error(
          "No se encontró la columna Año."
        );
      }

      if (!special.transshipment) {

        throw new Error(
          "No se encontró la columna Transshipment."
        );
      }

      if (!special.total_teus) {

        throw new Error(
          "No se encontró la columna Total TEUs."
        );
      }

      const yearSql =
        escapeSqlIdentifier(
          special.año
        );

      const transSql =
        escapeSqlIdentifier(
          special.transshipment
        );

      const totalSql =
        escapeSqlIdentifier(
          special.total_teus
        );

      const sql = `
        SELECT
          ${yearSql} AS año,
          SUM(${transSql}) AS total_transshipment,
          SUM(${totalSql}) AS total_teus
        FROM excel_data
        GROUP BY ${yearSql}
        ORDER BY ${yearSql};
      `;

      return executeSql(
        sql,
        "Transshipment por año"
      );

    } catch (error) {

      showError(
        error
      );

      updateStatus(
        "Error en Transshipment por año"
      );
    }
  }


  /* =========================================================
   * VERIFICACIÓN LOCAL + TRANSSHIPMENT
   * ========================================================= */

  async function localVsTransshipmentByYear() {

    try {

      if (!conn) {

        throw new Error(
          "Primero debes cargar un archivo Excel."
        );
      }

      const special =
        detectSpecialColumns();

      if (!special.año) {

        throw new Error(
          "No se encontró la columna Año."
        );
      }

      if (!special.local) {

        throw new Error(
          "No se encontró la columna Local."
        );
      }

      if (!special.transshipment) {

        throw new Error(
          "No se encontró la columna Transshipment."
        );
      }

      if (!special.total_teus) {

        throw new Error(
          "No se encontró la columna Total TEUs."
        );
      }

      const yearSql =
        escapeSqlIdentifier(
          special.año
        );

      const localSql =
        escapeSqlIdentifier(
          special.local
        );

      const transSql =
        escapeSqlIdentifier(
          special.transshipment
        );

      const totalSql =
        escapeSqlIdentifier(
          special.total_teus
        );

      const sql = `
        SELECT
          ${yearSql} AS año,
          SUM(${localSql}) AS total_local,
          SUM(${transSql}) AS total_transshipment,
          SUM(${totalSql}) AS total_teus
        FROM excel_data
        GROUP BY ${yearSql}
        ORDER BY ${yearSql};
      `;

      return executeSql(
        sql,
        "Local vs Transshipment por año"
      );

    } catch (error) {

      showError(
        error
      );

      updateStatus(
        "Error en Local vs Transshipment"
      );
    }
  }


  /* =========================================================
   * VERIFICAR TOTAL = LOCAL + TRANSSHIPMENT
   * ========================================================= */

  async function verifyTotalComposition() {

    try {

      if (!conn) {

        throw new Error(
          "Primero debes cargar un archivo Excel."
        );
      }

      const special =
        detectSpecialColumns();

      if (!special.año) {

        throw new Error(
          "No se encontró la columna Año."
        );
      }

      if (!special.local) {

        throw new Error(
          "No se encontró la columna Local."
        );
      }

      if (!special.transshipment) {

        throw new Error(
          "No se encontró la columna Transshipment."
        );
      }

      if (!special.total_teus) {

        throw new Error(
          "No se encontró la columna Total TEUs."
        );
      }

      const yearSql =
        escapeSqlIdentifier(
          special.año
        );

      const localSql =
        escapeSqlIdentifier(
          special.local
        );

      const transSql =
        escapeSqlIdentifier(
          special.transshipment
        );

      const totalSql =
        escapeSqlIdentifier(
          special.total_teus
        );

      const sql = `
        SELECT
          ${yearSql} AS año,
          SUM(${localSql}) AS local,
          SUM(${transSql}) AS transshipment,
          SUM(${totalSql}) AS total_teus,
          SUM(${localSql})
            + SUM(${transSql})
            AS calculado,
          SUM(${totalSql})
            - (
              SUM(${localSql})
              + SUM(${transSql})
            )
            AS diferencia
        FROM excel_data
        GROUP BY ${yearSql}
        ORDER BY ${yearSql};
      `;

      return executeSql(
        sql,
        "Verificación Total TEUs"
      );

    } catch (error) {

      showError(
        error
      );

      updateStatus(
        "Error en verificación"
      );
    }
  }


  /* =========================================================
   * ESQUEMA
   * ========================================================= */

  async function showSchema() {

    try {

      if (!conn) {

        throw new Error(
          "Primero debes cargar un archivo Excel."
        );
      }

      const schema =
        await getSchema();

      setResult({

        procesamiento:
          "LOCAL",

        tabla:
          "excel_data",

        columnas:
          schema,

        columnas_especiales_detectadas:
          detectSpecialColumns(),

        tipos_inferidos:
          currentInferredTypes
      });

      updateStatus(
        "Esquema mostrado"
      );

    } catch (error) {

      showError(
        error
      );
    }
  }


  /* =========================================================
   * SQL DESDE EDITOR
   * ========================================================= */

  async function runSqlFromEditor() {

    const editor =
      document.getElementById(
        "tm-excel-sql"
      );

    if (!editor) {

      showError(
        new Error(
          "No se encontró el editor SQL."
        )
      );

      return;
    }

    const sql =
      editor.value.trim();

    await executeSql(
      sql,
      "SQL"
    );
  }


  function updateStatus(
    message
  ) {

    const status =
      document.getElementById(
        "tm-excel-status"
      );

    if (status) {

      status.textContent =
        message;
    }
  }


  /* =========================================================
   * UI
   * ========================================================= */

  function createUI() {

    const old =
      document.getElementById(
        "tm-excel-engine-panel"
      );

    if (old) {
      old.remove();
    }

    const panel =
      document.createElement(
        "div"
      );

    panel.id =
      "tm-excel-engine-panel";

    panel.innerHTML = `

      <div
        style="
          font-family:Arial,sans-serif;
          width:100%;
          box-sizing:border-box;
        "
      >

        <div
          style="
            font-size:18px;
            font-weight:bold;
            margin-bottom:6px;
          "
        >
          📊 Excel Data Engine
        </div>

        <div
          style="
            font-size:12px;
            opacity:.75;
            margin-bottom:10px;
          "
        >
          ${VERSION} —
          Todo el procesamiento se realiza localmente.
        </div>

        <div
          id="tm-excel-status"
          style="
            padding:8px;
            margin-bottom:10px;
            border:1px solid #ddd;
            border-radius:6px;
          "
        >
          Motor listo. Carga un archivo Excel.
        </div>

        <input
          id="tm-excel-file"
          type="file"
          accept=".xlsx,.xls,.xlsm"
          style="
            width:100%;
            margin-bottom:10px;
          "
        />

        <div
          style="
            display:grid;
            grid-template-columns:
              repeat(2,minmax(0,1fr));
            gap:6px;
            margin-bottom:10px;
          "
        >

          <button
            id="tm-excel-count"
            type="button"
          >
            COUNT
          </button>

          <button
            id="tm-excel-preview"
            type="button"
          >
            Vista previa
          </button>

          <button
            id="tm-excel-summary"
            type="button"
          >
            Resumen
          </button>

          <button
            id="tm-excel-year"
            type="button"
          >
            Total por año
          </button>

          <button
            id="tm-excel-month"
            type="button"
          >
            Total por mes
          </button>

          <button
            id="tm-excel-local"
            type="button"
          >
            Total por Local
          </button>

          <button
            id="tm-excel-trans"
            type="button"
          >
            Transshipment por año
          </button>

          <button
            id="tm-excel-local-trans"
            type="button"
          >
            Local vs Transshipment
          </button>

          <button
            id="tm-excel-verify"
            type="button"
          >
            Verificar Total
          </button>

          <button
            id="tm-excel-schema"
            type="button"
          >
            Esquema
          </button>

        </div>

        <div
          style="
            font-weight:bold;
            margin-bottom:5px;
          "
        >
          Ejecutar SQL
        </div>

        <textarea
          id="tm-excel-sql"
          rows="8"
          spellcheck="false"
          style="
            width:100%;
            box-sizing:border-box;
            font-family:monospace;
            font-size:12px;
            padding:8px;
            border:1px solid #ccc;
            border-radius:6px;
            resize:vertical;
          "
        >SELECT
    "Año",
    SUM("Local") AS total_local,
    SUM("Transshipment") AS total_transshipment,
    SUM("Total TEUs") AS total_teus,
    SUM("Total TEUs")
      - (
        SUM("Local")
        + SUM("Transshipment")
      ) AS diferencia
FROM excel_data
GROUP BY "Año"
ORDER BY "Año";</textarea>

        <button
          id="tm-excel-run-sql"
          type="button"
          style="
            margin-top:6px;
            width:100%;
          "
        >
          Ejecutar SQL
        </button>

        <div
          style="
            font-weight:bold;
            margin-top:12px;
            margin-bottom:5px;
          "
        >
          Resultado
        </div>

        <pre
          id="tm-excel-result"
          style="
            white-space:pre-wrap;
            word-break:break-word;
            max-height:500px;
            overflow:auto;
            padding:10px;
            border:1px solid #ddd;
            border-radius:6px;
            font-family:monospace;
            font-size:12px;
            background:#fafafa;
          "
        ></pre>

      </div>
    `;


    const possibleHosts = [
      "#right-sidebar",
      "[data-testid='right-sidebar']",
      ".right-sidebar",
      "aside"
    ];


    let host = null;


    for (
      const selector
      of possibleHosts
    ) {

      host =
        document.querySelector(
          selector
        );

      if (host) {
        break;
      }
    }


    if (host) {

      host.appendChild(
        panel
      );

    } else {

      panel.style.position =
        "fixed";

      panel.style.right =
        "10px";

      panel.style.top =
        "80px";

      panel.style.width =
        "420px";

      panel.style.maxHeight =
        "calc(100vh - 100px)";

      panel.style.overflow =
        "auto";

      panel.style.zIndex =
        "999999";

      panel.style.background =
        "white";

      panel.style.padding =
        "12px";

      panel.style.border =
        "1px solid #ccc";

      panel.style.borderRadius =
        "8px";

      panel.style.boxShadow =
        "0 4px 20px rgba(0,0,0,.15)";

      document.body.appendChild(
        panel
      );
    }


    document
      .getElementById(
        "tm-excel-file"
      )
      .addEventListener(
        "change",
        async (event) => {

          const file =
            event.target.files?.[0];

          if (!file) {
            return;
          }

          await loadExcel(
            file
          );
        }
      );


    document
      .getElementById(
        "tm-excel-count"
      )
      .addEventListener(
        "click",
        countRows
      );


    document
      .getElementById(
        "tm-excel-preview"
      )
      .addEventListener(
        "click",
        previewRows
      );


    document
      .getElementById(
        "tm-excel-summary"
      )
      .addEventListener(
        "click",
        summary
      );


    document
      .getElementById(
        "tm-excel-year"
      )
      .addEventListener(
        "click",
        totalByYear
      );


    document
      .getElementById(
        "tm-excel-month"
      )
      .addEventListener(
        "click",
        totalByMonth
      );


    document
      .getElementById(
        "tm-excel-local"
      )
      .addEventListener(
        "click",
        totalByLocal
      );


    document
      .getElementById(
        "tm-excel-trans"
      )
      .addEventListener(
        "click",
        totalTransshipmentByYear
      );


    document
      .getElementById(
        "tm-excel-local-trans"
      )
      .addEventListener(
        "click",
        localVsTransshipmentByYear
      );


    document
      .getElementById(
        "tm-excel-verify"
      )
      .addEventListener(
        "click",
        verifyTotalComposition
      );


    document
      .getElementById(
        "tm-excel-schema"
      )
      .addEventListener(
        "click",
        showSchema
      );


    document
      .getElementById(
        "tm-excel-run-sql"
      )
      .addEventListener(
        "click",
        runSqlFromEditor
      );
  }


  /* =========================================================
   * INICIALIZACIÓN
   * ========================================================= */

  async function initialize() {

    try {

      createUI();

      updateStatus(
        `Inicializando ${VERSION}...`
      );

      await loadLibraries();

      await createDuckDB();

      const version =
        await getDuckDBVersion();

      updateStatus(
        `${VERSION} listo — DuckDB-Wasm cargado localmente`
      );


      window.TMExcelEngine = {

        version:
          VERSION,

        app_id:
          APP_ID,

        duckdb_version:
          version,

        executeSql:
          executeSql,

        loadExcel:
          loadExcel,

        countRows:
          countRows,

        previewRows:
          previewRows,

        summary:
          summary,

        totalByYear:
          totalByYear,

        totalByMonth:
          totalByMonth,

        totalByLocal:
          totalByLocal,

        totalTransshipmentByYear:
          totalTransshipmentByYear,

        localVsTransshipmentByYear:
          localVsTransshipmentByYear,

        verifyTotalComposition:
          verifyTotalComposition,

        showSchema:
          showSchema,

        detectSpecialColumns:
          detectSpecialColumns,

        getSchema:
          getSchema
      };

    } catch (error) {

      showError(
        error
      );

      updateStatus(
        "No se pudo inicializar el motor"
      );
    }
  }


  /*
   * Evita instalar varias veces
   * la extensión.
   */
  if (
    window.__TM_EXCEL_ENGINE_0416_INITIALIZED
  ) {
    return;
  }

  window.__TM_EXCEL_ENGINE_0416_INITIALIZED =
    true;

  initialize();

})();
