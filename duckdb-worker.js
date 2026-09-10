/*
 * TypingMind Excel Data Engine
 * Extension v0.4.23 TEST
 *
 * CAMBIOS v0.4.23
 *
 * - Parte de v0.4.22 (baseline funcional intacto).
 * - FIX CRÍTICO UI MÓVIL Y ESCRITORIO:
 *   1. Contenedor arrastrable (Draggable) touch y mouse.
 *   2. Modo Micro-Widget: colapsado se reduce a mini-botón circular (46px) con 📊 y 🟢/⚪/🔴.
 *   3. Z-Index alto (999999) y discriminación entre drag y click para evitar toggles indeseados.
 *
 * CAMBIOS v0.4.22
 *
 * - Parte de v0.4.21 (baseline de UI plegable intacto).
 *
 * - Solo UI: panel compacto para móvil, sin el recuadro
 *   de estado/versión. El listo se indica con 🟢 en el
 *   summary. Sin cambios en DuckDB, serialización ni
 *   window.TMExcelEngine.
 *
 * CAMBIOS v0.4.21
 *
 * - Parte de v0.4.20 (baseline intacto).
 *
 * - Solo cambia la UI inyectada: panel plegable con
 *   <details> y herramientas de prueba ocultas en
 *   Modo Desarrollador. Sin cambios en DuckDB,
 *   serialización ni en la API de window.TMExcelEngine.
 *
 * CAMBIOS v0.4.20
 *
 * - Parte de v0.4.19 (baseline intacto).
 *
 * - Arrow representa HUGEINT / Decimal(38,0) como un
 *   typed array de 4 enteros: [valor, 0, 0, 0].
 *   v0.4.19 solo lo convertía a escalar si el alias
 *   coincidía con sum|avg|min|max|count|total|...
 *   Un alias como `x` dejaba el array en el JSON.
 *
 * - Ahora ese patrón se convierte a escalar siempre,
 *   sin depender del nombre de la columna.
 *
 * CAMBIOS v0.4.19
 *
 * - Mantiene las correcciones de v0.4.17 y v0.4.18.
 *
 * - Mantiene procesamiento 100% LOCAL.
 *
 * - DuckDB-Wasm 1.29.0
 * - DuckDB interno 1.1.1
 * - SheetJS 0.18.5
 *
 * MEJORAS:
 *
 * 1. Normalización robusta de nombres de columnas.
 *
 *    Permite reconocer variantes como:
 *
 *      Total TEUs
 *      Total TEU's
 *      Total TEU´s
 *      Total TEU
 *      Total
 *
 *    como posibles variantes del mismo concepto.
 *
 * 2. Se conserva SIEMPRE el nombre original
 *    de la columna dentro de DuckDB.
 *
 * 3. Se agrega resolución de columnas mediante
 *    alias normalizados.
 *
 * 4. Se mejora la detección de columnas especiales.
 *
 * 5. Se mantiene la corrección de resultados
 *    Arrow / TypedArray / UInt32 -> Int32.
 *
 * 6. Se mantiene BIGINT, DOUBLE, DATE, VARCHAR y NULL.
 *
 * 7. Se mantiene SQL libre.
 *
 * 8. Se agregan funciones auxiliares para inspeccionar
 *    columnas y aliases detectados.
 */

(() => {
  "use strict";

  const APP_ID =
    "tm-excel-engine-v0423-test";

  const VERSION =
    "v0.4.23-test";

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

  /*
   * ============================================================
   * UTILIDADES GENERALES
   * ============================================================
   */

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

  /*
   * Normaliza nombres únicamente para comparaciones.
   *
   * NO modifica el nombre original.
   *
   * Ejemplos:
   *
   * "Total TEUs"
   * "Total TEU's"
   * "Total TEU´s"
   *
   * terminan en una representación comparable.
   */

  function normalizeHeaderName(value) {
    return safeString(value)
      .normalize("NFD")
      .replace(
        /[\u0300-\u036f]/g,
        ""
      )
      .toLowerCase()
      .replace(
        /[’'´`]/g,
        ""
      )
      .replace(
        /[^a-z0-9]+/g,
        ""
      )
      .trim();
  }

  /*
   * ============================================================
   * NORMALIZACIÓN NUMÉRICA
   * ============================================================
   */

  function normalizeSigned32Number(value) {
    if (
      typeof value !== "number" ||
      !Number.isFinite(value)
    ) {
      return value;
    }

    if (
      Number.isInteger(value) &&
      value >= 4294967295 - 2147483648
    ) {
      return value - 4294967296;
    }

    return value;
  }

  function normalizeNumericArray(value) {
    if (
      !Array.isArray(value) ||
      value.length !== 4
    ) {
      return null;
    }

    const nums =
      value.map((v) => {
        if (typeof v === "bigint") {
          return Number(v);
        }

        if (typeof v === "number") {
          return v;
        }

        return Number(v);
      });

    if (
      !nums.every(
        Number.isFinite
      )
    ) {
      return null;
    }

    const signed =
      nums.map(
        normalizeSigned32Number
      );

    /*
     * Patrón:
     *
     * [valor, 0, 0, 0]
     */

    if (
      signed[1] === 0 &&
      signed[2] === 0 &&
      signed[3] === 0
    ) {
      return signed[0];
    }

    /*
     * Patrón observado:
     *
     * [valor, -1, -1, -1]
     */

    if (
      signed[1] === -1 &&
      signed[2] === -1 &&
      signed[3] === -1
    ) {
      return signed[0];
    }

    return null;
  }

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

    if (
      typeof value === "bigint"
    ) {
      const n =
        Number(value);

      return Number.isSafeInteger(n)
        ? n
        : value.toString();
    }

    if (
      value instanceof Date
    ) {
      return Number.isNaN(
        value.getTime()
      )
        ? null
        : value.toISOString();
    }

    if (
      value instanceof ArrayBuffer
    ) {
      return Array.from(
        new Uint8Array(value)
      );
    }

    if (
      ArrayBuffer.isView(value)
    ) {
      const arr =
        Array.from(value);

      const scalar =
        normalizeNumericArray(
          arr
        );

      if (scalar !== null) {
        return scalar;
      }

      return arr.map(
        (v) =>
          normalizeSigned32Number(
            v
          )
      );
    }

    if (
      Array.isArray(value)
    ) {
      const scalar =
        normalizeNumericArray(
          value
        );

      if (scalar !== null) {
        return scalar;
      }

      return value.map(
        (v) =>
          normalizeValue(
            v,
            key
          )
      );
    }

    if (
      typeof value === "object"
    ) {
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

  /*
   * ============================================================
   * RESULTADOS
   * ============================================================
   */

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
      mensaje:
        message,
      stack:
        stack
    });
  }

  /*
   * ============================================================
   * SQL
   * ============================================================
   */

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

  /*
   * ============================================================
   * ENCABEZADOS
   * ============================================================
   */

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
          normalizeHeaderName(
            name
          );

        if (
          !used.has(
            normalized
          )
        ) {
          used.set(
            normalized,
            1
          );

          result.push(
            name
          );

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

  /*
   * ============================================================
   * ALIASES DE COLUMNAS
   * ============================================================
   *
   * Esta función permite que distintas formas de escribir
   * un encabezado sean reconocidas como el mismo concepto.
   */

  function getColumnAliases() {
    return {
      año: [
        "Año",
        "Ano",
        "Year",
        "Año Fiscal",
        "Fiscal Year"
      ],

      fecha: [
        "Fecha",
        "Date",
        "Fecha Evento",
        "Event Date"
      ],

      mes: [
        "Mes",
        "Month"
      ],

      total_teus: [
        "Total TEUs",
        "Total TEU's",
        "Total TEU´s",
        "Total TEU",
        "Total TEUS",
        "TEUs Totales",
        "TEU Totales",
        "Total"
      ],

      local: [
        "Local",
        "Locales",
        "Local TEUs",
        "Local TEU"
      ],

      transshipment: [
        "Transshipment",
        "Transhipment",
        "Transshipment TEUs",
        "Transshipment TEU",
        "Transbordo"
      ]
    };
  }

  /*
   * Busca primero coincidencia exacta normalizada.
   */

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

  /*
   * Devuelve información detallada de resolución.
   */

  function resolveColumn(
    headers,
    candidates
  ) {
    if (
      !Array.isArray(headers)
    ) {
      return {
        found: false,
        column: null,
        matched_alias: null
      };
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
        return {
          found: true,
          column: headers[index],
          matched_alias: candidate,
          normalized:
            normalizedHeaders[index]
        };
      }
    }

    return {
      found: false,
      column: null,
      matched_alias: null
    };
  }

  function detectSpecialColumns() {
    const aliases =
      getColumnAliases();

    const year =
      resolveColumn(
        currentHeaders,
        aliases.año
      );

    const date =
      resolveColumn(
        currentHeaders,
        aliases.fecha
      );

    const total =
      resolveColumn(
        currentHeaders,
        aliases.total_teus
      );

    const local =
      resolveColumn(
        currentHeaders,
        aliases.local
      );

    const transshipment =
      resolveColumn(
        currentHeaders,
        aliases.transshipment
      );

    const month =
      resolveColumn(
        currentHeaders,
        aliases.mes
      );

    return {
      año:
        year.column,

      fecha:
        date.column,

      mes:
        month.column,

      total_teus:
        total.column,

      local:
        local.column,

      transshipment:
        transshipment.column
    };
  }

  function getColumnResolutionDetails() {
    const aliases =
      getColumnAliases();

    return {
      año:
        resolveColumn(
          currentHeaders,
          aliases.año
        ),

      fecha:
        resolveColumn(
          currentHeaders,
          aliases.fecha
        ),

      mes:
        resolveColumn(
          currentHeaders,
          aliases.mes
        ),

      total_teus:
        resolveColumn(
          currentHeaders,
          aliases.total_teus
        ),

      local:
        resolveColumn(
          currentHeaders,
          aliases.local
        ),

      transshipment:
        resolveColumn(
          currentHeaders,
          aliases.transshipment
        )
    };
  }

  /*
   * ============================================================
   * INFERENCIA DE TIPOS
   * ============================================================
   */

  function isDateHeader(name) {
    const normalized =
      normalizeHeaderName(
        name
      );

    return [
      "fecha",
      "date",
      "fechacombinada",
      "fechaevento",
      "eventdate"
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

    if (
      /^[A-Za-zÁÉÍÓÚáéíóú]+[\/-]\d{4}$/.test(
        text
      )
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
      return Number.isInteger(
        value
      );
    }

    const text =
      safeString(value)
        .trim();

    if (!text) {
      return false;
    }

    return /^-?\d+$/.test(
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
      return Number.isFinite(
        value
      );
    }

    const text =
      safeString(value)
        .replace(
          /,/g,
          ""
        )
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
      isDateHeader(
        columnName
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

    const numericSpecialNames = [
      "local",
      "locales",
      "localteus",
      "localteu",
      "transshipment",
      "transhipment",
      "transshipmentteus",
      "transshipmentteu",
      "transbordo",
      "totalteus",
      "totalteu",
      "totalteus",
      "teustotales",
      "teutotales",
      "total"
    ];

    if (
      numericSpecialNames.includes(
        normalizedName
      )
    ) {
      const allIntegers =
        values.length > 0 &&
        values.every(
          looksLikeInteger
        );

      if (
        allIntegers
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

      if (
        allNumeric
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

  /*
   * ============================================================
   * FECHAS
   * ============================================================
   */

  function excelSerialToDate(serial) {
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

  function normalizeExcelDate(value) {
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
      /^\d{4}-\d{1,2}-\d{1,2}$/.test(
        text
      )
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

  /*
   * ============================================================
   * SQL DE VALORES
   * ============================================================
   */

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
      const row of rows
    ) {
      const values =
        headers.map(
          (
            _,
            index
          ) =>
            valueToSql(
              row[index],
              inferredTypes[
                index
              ]?.sqlType ||
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

  /*
   * ============================================================
   * LIBRERÍAS
   * ============================================================
   */

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

  /*
   * ============================================================
   * DUCKDB
   * ============================================================
   */

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

    return normalizeRows(
      rows
    );
  }

  async function getDuckDBVersion() {
    try {
      return await queryRows(
        "SELECT * FROM pragma_version();"
      );
    } catch (
      error
    ) {
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

  /*
   * ============================================================
   * LECTURA DEL EXCEL
   * ============================================================
   */

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
        (
          row
        ) => {
          const result = [];

          for (
            let i = 0;
            i < headers.length;
            i++
          ) {
            result.push(
              row?.[i] ??
              null
            );
          }

          return result;
        }
      );

    const realRows =
      normalizedRows.filter(
        (
          row
        ) =>
          row.some(
            (
              value
            ) =>
              !isEmptyValue(
                value
              )
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

  /*
   * ============================================================
   * CARGAR EXCEL
   * ============================================================
   */

  async function loadExcel(file) {
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
            type:
              "array",

            cellDates:
              true,

            raw:
              true
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

      const resolution =
        getColumnResolutionDetails();

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

        resolucion_columnas:
          resolution,

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

    } catch (
      error
    ) {
      showError(
        error
      );

      updateStatus(
        "Error al cargar Excel"
      );
    }
  }

  /*
   * ============================================================
   * SQL LIBRE
   * ============================================================
   */

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

    } catch (
      error
    ) {
      showError(
        error
      );

      updateStatus(
        `Error en ${label}`
      );

      throw error;
    }
  }

  /*
   * ============================================================
   * CONSULTAS PREDEFINIDAS
   * ============================================================
   */

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
          (
            column
          ) =>
            /BIGINT|DOUBLE|DECIMAL|INTEGER|HUGEINT|FLOAT/i.test(
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

    } catch (
      error
    ) {
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

      if (
        !special.año
      ) {
        throw new Error(
          `No se encontró la columna de año. Columnas disponibles: ${currentHeaders.join(
            ", "
          )}`
        );
      }

      if (
        !special.total_teus
      ) {
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

    } catch (
      error
    ) {
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

      if (
        !special.año
      ) {
        throw new Error(
          "No se encontró la columna Año."
        );
      }

      if (
        !special.mes
      ) {
        throw new Error(
          "No se encontró la columna Mes."
        );
      }

      if (
        !special.total_teus
      ) {
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

    } catch (
      error
    ) {
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

      if (
        !special.local
      ) {
        throw new Error(
          "No se encontró la columna Local."
        );
      }

      if (
        !special.total_teus
      ) {
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

    } catch (
      error
    ) {
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

      if (
        !special.año
      ) {
        throw new Error(
          "No se encontró la columna Año."
        );
      }

      if (
        !special.transshipment
      ) {
        throw new Error(
          "No se encontró la columna Transshipment."
        );
      }

      if (
        !special.total_teus
      ) {
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

    } catch (
      error
    ) {
      showError(
        error
      );

      updateStatus(
        "Error en Transshipment por año"
      );
    }
  }

  /*
   * ============================================================
   * ESQUEMA
   * ============================================================
   */

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

        resolucion_columnas:
          getColumnResolutionDetails(),

        tipos_inferidos:
          currentInferredTypes
      });

      updateStatus(
        "Esquema mostrado"
      );

    } catch (
      error
    ) {
      showError(
        error
      );
    }
  }

  /*
   * ============================================================
   * DIAGNÓSTICO DE COLUMNAS
   * ============================================================
   */

  async function showColumnResolution() {
    try {
      if (!conn) {
        throw new Error(
          "Primero debes cargar un archivo Excel."
        );
      }

      setResult({
        procesamiento:
          "LOCAL",

        encabezados_originales:
          currentHeaders,

        aliases_disponibles:
          getColumnAliases(),

        resolucion:
          getColumnResolutionDetails(),

        columnas_especiales:
          detectSpecialColumns()
      });

      updateStatus(
        "Resolución de columnas mostrada"
      );

    } catch (
      error
    ) {
      showError(
        error
      );
    }
  }

  /*
   * ============================================================
   * EDITOR SQL
   * ============================================================
   */

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

  /*
   * ============================================================
   * UI
   * ============================================================
   */

  function updateStatus(message) {
    const status =
      document.getElementById(
        "tm-excel-status"
      );

    if (status) {
      status.textContent =
        message;
    }

    const summary =
      document.getElementById(
        "tm-excel-engine-summary"
      );

    if (summary) {
      summary.title =
        safeString(message);
    }

    const dot =
      document.getElementById(
        "tm-excel-ready-dot"
      );

    if (dot) {
      const text =
        safeString(message)
          .toLowerCase();

      const isError =
        text.includes("error") ||
        text.includes("no se pudo");

      const isReady =
        !isError &&
        (
          text.includes("listo") ||
          text.includes("excel cargado") ||
          text.includes("ejecutado localmente")
        );

      dot.textContent =
        isError
          ? "🔴"
          : isReady
            ? "🟢"
            : "⚪";
    }
  }

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

    panel.style.cssText = `
      max-width:95vw;
      width:100%;
      min-width:0;
      box-sizing:border-box;
    `;

    panel.innerHTML = `
      <details
        id="tm-excel-engine-details"
        open
        style="
          font-family:Arial,sans-serif;
          max-width:95vw;
          width:100%;
          min-width:0;
          box-sizing:border-box;
          font-size:14px;
          padding:10px;
          overflow-wrap:break-word;
          overflow-x:hidden;
        "
      >
        <summary
          id="tm-excel-engine-summary"
          style="font-weight:bold; cursor:grab; user-select:none; display:flex; align-items:center; justify-content:space-between; list-style:none;"
        ><span id="tm-excel-title-content" style="display:flex; align-items:center; gap:6px;"><span id="tm-excel-icon" style="font-size:16px;">📊</span><span id="tm-excel-title-text">Excel Data Engine</span></span><span id="tm-excel-ready-dot" style="font-size:12px;">⚪</span></summary>

        <div
          id="tm-excel-status"
          hidden
        ></div>

        <input
          id="tm-excel-file"
          type="file"
          accept=".xlsx,.xls,.xlsm"
          style="
            display:block;
            max-width:100%;
            width:100%;
            box-sizing:border-box;
            margin-top:8px;
            margin-bottom:10px;
          "
        />

        <details>
          <summary style="cursor:pointer; color:gray; font-size:0.9em;">🛠️ Modo Desarrollador / Pruebas</summary>

          <div
            style="
              display:grid;
              grid-template-columns:
                repeat(2,minmax(0,1fr));
              gap:6px;
              margin-top:8px;
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
              id="tm-excel-schema"
              type="button"
            >
              Esquema
            </button>

            <button
              id="tm-excel-resolution"
              type="button"
            >
              Resolución columnas
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
            rows="10"
            spellcheck="false"
            style="
              width:100%;
              max-width:100%;
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
    "Mes",
    EXTRACT(MONTH FROM "Fecha") AS numero_mes,
    SUM("Total TEUs") AS total_mes
FROM excel_data
GROUP BY
    "Año",
    "Mes",
    EXTRACT(MONTH FROM "Fecha")
ORDER BY
    "Año",
    numero_mes;</textarea>

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
              max-width:100%;
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
        </details>
      </details>
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

      document.body.appendChild(panel);
    }

    // --- COMPORTAMIENTO MICRO-WIDGET Y DRAGGABLE v0.4.23 ---
    const details = panel.querySelector("#tm-excel-engine-details");
    const summary = panel.querySelector("#tm-excel-engine-summary");
    const titleText = panel.querySelector("#tm-excel-title-text");
    const readyDot = panel.querySelector("#tm-excel-ready-dot");

    function applyMicroWidget() {
      if (!details) return;
      if (details.open) {
        panel.style.width = "min(92vw, 420px)";
        panel.style.height = "auto";
        panel.style.minWidth = "0";
        panel.style.minHeight = "0";
        panel.style.borderRadius = "10px";
        panel.style.padding = "12px";
        panel.style.boxShadow = "0 8px 30px rgba(0,0,0,.2)";
        if (titleText) titleText.style.display = "inline";
        if (summary) {
          summary.style.justifyContent = "space-between";
          summary.style.padding = "0";
          summary.style.height = "auto";
        }
        if (readyDot) {
          readyDot.style.position = "static";
        }
      } else {
        panel.style.width = "46px";
        panel.style.height = "46px";
        panel.style.minWidth = "46px";
        panel.style.minHeight = "46px";
        panel.style.borderRadius = "50%";
        panel.style.padding = "0";
        panel.style.overflow = "hidden";
        panel.style.boxShadow = "0 4px 16px rgba(0,0,0,.25)";
        if (titleText) titleText.style.display = "none";
        if (summary) {
          summary.style.width = "100%";
          summary.style.height = "100%";
          summary.style.display = "flex";
          summary.style.alignItems = "center";
          summary.style.justifyContent = "center";
          summary.style.position = "relative";
          summary.style.padding = "0";
        }
        if (readyDot) {
          readyDot.style.position = "absolute";
          readyDot.style.top = "5px";
          readyDot.style.right = "5px";
          readyDot.style.fontSize = "10px";
        }
      }
    }

    if (details) {
      details.addEventListener("toggle", applyMicroWidget);
    }

    let isDragging = false;
    let startX = 0, startY = 0;
    let origLeft = 0, origTop = 0;
    let didMove = false;

    function getPoint(e) {
      if (e.touches && e.touches.length > 0) {
        return { x: e.touches[0].clientX, y: e.touches[0].clientY };
      }
      return { x: e.clientX, y: e.clientY };
    }

    function onDragStart(e) {
      if (e.target && (e.target.tagName === "INPUT" || e.target.tagName === "BUTTON" || e.target.tagName === "TEXTAREA")) {
        return;
      }
      isDragging = true;
      didMove = false;
      const pt = getPoint(e);
      startX = pt.x;
      startY = pt.y;

      const rect = panel.getBoundingClientRect();
      origLeft = rect.left;
      origTop = rect.top;

      panel.style.position = "fixed";
      panel.style.left = origLeft + "px";
      panel.style.top = origTop + "px";
      panel.style.right = "auto";
      panel.style.bottom = "auto";
      panel.style.cursor = "grabbing";
    }

    function onDragMove(e) {
      if (!isDragging) return;
      const pt = getPoint(e);
      const dx = pt.x - startX;
      const dy = pt.y - startY;

      if (Math.hypot(dx, dy) > 5) {
        didMove = true;
      }

      let nx = origLeft + dx;
      let ny = origTop + dy;

      const w = panel.offsetWidth || 46;
      const h = panel.offsetHeight || 46;

      nx = Math.max(4, Math.min(window.innerWidth - w - 4, nx));
      ny = Math.max(4, Math.min(window.innerHeight - h - 4, ny));

      panel.style.left = nx + "px";
      panel.style.top = ny + "px";
    }

    function onDragEnd() {
      if (!isDragging) return;
      isDragging = false;
      panel.style.cursor = "";
      if (summary) summary.style.cursor = "grab";
    }

    if (summary) {
      summary.addEventListener("click", (e) => {
        if (didMove) {
          e.preventDefault();
          e.stopPropagation();
          didMove = false;
        }
      }, true);

      summary.addEventListener("mousedown", onDragStart);
      summary.addEventListener("touchstart", onDragStart, { passive: true });
    }

    document.addEventListener("mousemove", onDragMove);
    document.addEventListener("touchmove", onDragMove, { passive: true });
    document.addEventListener("mouseup", onDragEnd);
    document.addEventListener("touchend", onDragEnd);
    // --- FIN COMPORTAMIENTO v0.4.23 ---

    document
      .getElementById(
        "tm-excel-file"
      )
      .addEventListener(
        "change",
        async (
          event
        ) => {
          const file =
            event.target
              .files?.[0];

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
        "tm-excel-schema"
      )
      .addEventListener(
        "click",
        showSchema
      );

    document
      .getElementById(
        "tm-excel-resolution"
      )
      .addEventListener(
        "click",
        showColumnResolution
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

  /*
   * ============================================================
   * INICIALIZACIÓN
   * ============================================================
   */

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

        showSchema:
          showSchema,

        showColumnResolution:
          showColumnResolution,

        detectSpecialColumns:
          detectSpecialColumns,

        getColumnResolutionDetails:
          getColumnResolutionDetails,

        getSchema:
          getSchema
      };

    } catch (
      error
    ) {
      showError(
        error
      );

      updateStatus(
        "No se pudo inicializar el motor"
      );
    }
  }

  /*
   * ============================================================
   * PROTECCIÓN CONTRA DOBLE INSTALACIÓN
   * ============================================================
   */

  if (
    window.__TM_EXCEL_ENGINE_0419_INITIALIZED
  ) {
    return;
  }

  window.__TM_EXCEL_ENGINE_0419_INITIALIZED =
    true;

  initialize();

})();
/* ============================================================
 * CARGA DE DATOS PARQUET
 * ============================================================ */

async function loadParquet(parquetData, tableName = "excel_data") {
  if (!conn) {
    await initDuckDB();
  }

  if (!parquetData) {
    throw new Error("No se proporcionaron datos Parquet.");
  }

  const cleanTableName = tableName.replace(/[^a-zA-Z0-9_]/g, "_");

  const start = performance.now();

  await conn.query(`DROP TABLE IF EXISTS ${cleanTableName};`);

  const createSql = `
    CREATE TABLE ${cleanTableName} AS
    SELECT * FROM read_parquet('${parquetData}');
  `;

  await conn.query(createSql);
  const elapsed = performance.now() - start;

  const countResult = await conn.query(`SELECT COUNT(*) AS registros FROM ${cleanTableName};`);
  const countRows = countResult.toArray();

  const schemaResult = await conn.query(`
    SELECT column_name, data_type
    FROM information_schema.columns
    WHERE table_name = '${cleanTableName}'
    ORDER BY ordinal_position;
  `);
  const schema = schemaResult.toArray();

  const previewResult = await conn.query(`SELECT * FROM ${cleanTableName} LIMIT 10;`);
  const preview = previewResult.toArray();

  return {
    procesamiento: "LOCAL_NATIVO_WORKER",
    formato: "PARQUET",
    tabla: cleanTableName,
    registros: countRows[0]?.registros ?? 0,
    tiempo_carga_ms: Number(elapsed.toFixed(1)),
    columnas: schema.map((c) => c.column_name),
    esquema: schema,
    preview: preview
  };
}

/* ============================================================
 * DISPATCHER DE MENSAJES
 * ============================================================ */

self.onmessage = async (event) => {
  const { type, ...payload } = event.data || {};

  try {
    let result;

    switch (type) {
      case "initDuckDB":
        result = await initDuckDB();
        break;

      case "executeQuery":
        result = await executeQuery(payload.sql);
        break;

      case "loadCSV":
        result = await loadCSV(payload.csvData, payload.tableName);
        break;

      case "loadParquet":
        result = await loadParquet(payload.parquetData, payload.tableName);
        break;

      default:
        throw new Error(`Tipo de acción desconocido: ${type}`);
    }

    self.postMessage({
      success: true,
      data: result
    });
  } catch (error) {
    self.postMessage({
      success: false,
      error: error.message || "Error desconocido"
    });
  }
};
