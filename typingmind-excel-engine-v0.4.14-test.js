/*
 * TypingMind Excel Data Engine
 * Extension v0.4.14 TEST
 *
 * CAMBIOS v0.4.14
 * - Compatible con DuckDB-Wasm 1.29.0 / DuckDB 1.1.1.
 * - Corrige information_schema.columns:
 *   usa DATA_TYPE en lugar de COLUMN_TYPE.
 * - Corrige inferencia de fechas.
 * - Los números como 87840, 31391, 119231, etc.
 *   NO se interpretan como fechas salvo que la columna
 *   tenga encabezado de fecha.
 * - Detecta "Total TEUs", "Total TEU's" y variantes.
 * - Total por año usa los encabezados reales.
 * - Total por mes usa los encabezados reales.
 * - Total por Local usa los encabezados reales.
 * - Normaliza resultados Arrow/DuckDB-Wasm.
 * - Mantiene procesamiento 100% LOCAL.
 */

(() => {
  "use strict";

  const APP_ID = "tm-excel-engine-v0414-test";
  const VERSION = "v0.4.14-test";

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
  let currentTypes = [];

  function safeString(value) {
    return value === null || value === undefined
      ? ""
      : String(value);
  }

  function isEmptyValue(value) {
    return (
      value === null ||
      value === undefined ||
      (typeof value === "string" &&
        value.trim() === "")
    );
  }

  /*
   * Normaliza encabezados únicamente para comparar.
   * No modifica el nombre real.
   */
  function normalizeHeaderName(value) {
    return safeString(value)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/[’'´`]/g, "")
      .replace(/[^a-z0-9]+/g, "")
      .trim();
  }

  function escapeSqlIdentifier(name) {
    return (
      '"' +
      safeString(name).replace(/"/g, '""') +
      '"'
    );
  }

  function escapeSqlString(value) {
    return safeString(value).replace(/'/g, "''");
  }

  /*
   * Convierte valores especiales de DuckDB-Wasm.
   */
  function normalizeNumericArray(value) {
    if (!Array.isArray(value)) {
      return null;
    }

    if (value.length !== 4) {
      return null;
    }

    const nums = value.map(v => {
      if (typeof v === "bigint") {
        return Number(v);
      }

      return Number(v);
    });

    if (!nums.every(Number.isFinite)) {
      return null;
    }

    if (
      nums[1] === 0 &&
      nums[2] === 0 &&
      nums[3] === 0
    ) {
      return nums[0];
    }

    return null;
  }

  function normalizeValue(value, key = "") {
    if (
      value === null ||
      value === undefined
    ) {
      return null;
    }

    if (typeof value === "bigint") {
      const n = Number(value);

      return Number.isSafeInteger(n)
        ? n
        : value.toString();
    }

    if (value instanceof Date) {
      return Number.isNaN(value.getTime())
        ? null
        : value.toISOString();
    }

    if (value instanceof ArrayBuffer) {
      return Array.from(
        new Uint8Array(value)
      );
    }

    if (ArrayBuffer.isView(value)) {
      const arr = Array.from(value);
      const scalar =
        normalizeNumericArray(arr);

      if (
        scalar !== null &&
        /sum|avg|min|max|count|total|registros/i.test(
          key
        )
      ) {
        return scalar;
      }

      return arr.map(v =>
        normalizeValue(v, key)
      );
    }

    if (Array.isArray(value)) {
      const scalar =
        normalizeNumericArray(value);

      if (
        scalar !== null &&
        /sum|avg|min|max|count|total|registros/i.test(
          key
        )
      ) {
        return scalar;
      }

      return value.map(v =>
        normalizeValue(v, key)
      );
    }

    if (typeof value === "object") {
      const output = {};

      for (const [k, v] of Object.entries(value)) {
        output[k] = normalizeValue(v, k);
      }

      return output;
    }

    return value;
  }

  function normalizeRows(rows) {
    if (!Array.isArray(rows)) {
      return [];
    }

    return rows.map(row =>
      normalizeValue(row)
    );
  }

  function jsonSafe(value) {
    return JSON.parse(
      JSON.stringify(value, (_, v) => {
        if (typeof v === "bigint") {
          const n = Number(v);

          return Number.isSafeInteger(n)
            ? n
            : v.toString();
        }

        return v;
      })
    );
  }

  function setResult(value) {
    const output =
      document.getElementById(
        "tm-excel-result"
      );

    if (!output) {
      return;
    }

    output.textContent =
      JSON.stringify(
        jsonSafe(
          normalizeValue(value)
        ),
        null,
        2
      );
  }

  function showError(error) {
    setResult({
      mensaje:
        error?.message ||
        String(error),

      stack:
        error?.stack ||
        ""
    });
  }

  function makeUniqueHeaders(headers) {
    const result = [];
    const used = new Map();

    headers.forEach((header, index) => {
      let name =
        safeString(header).trim();

      if (!name) {
        name =
          `Columna_${index + 1}`;
      }

      const key =
        name.toLowerCase();

      if (!used.has(key)) {
        used.set(key, 1);
        result.push(name);
        return;
      }

      const count =
        used.get(key) + 1;

      used.set(key, count);

      result.push(
        `${name}_${count}`
      );
    });

    return result;
  }

  function detectColumn(headers, candidates) {
    if (!Array.isArray(headers)) {
      return null;
    }

    const normalized =
      headers.map(h =>
        normalizeHeaderName(h)
      );

    for (const candidate of candidates) {
      const target =
        normalizeHeaderName(candidate);

      const index =
        normalized.indexOf(target);

      if (index >= 0) {
        return headers[index];
      }
    }

    return null;
  }

  function detectSpecialColumns() {
    return {
      año: detectColumn(
        currentHeaders,
        [
          "Año",
          "Ano",
          "Year"
        ]
      ),

      fecha: detectColumn(
        currentHeaders,
        [
          "Fecha",
          "Date"
        ]
      ),

      total_teus: detectColumn(
        currentHeaders,
        [
          "Total TEUs",
          "Total TEU's",
          "Total TEU´s",
          "Total TEU",
          "Total"
        ]
      ),

      local: detectColumn(
        currentHeaders,
        ["Local"]
      ),

      transshipment: detectColumn(
        currentHeaders,
        [
          "Transshipment",
          "Transhipment",
          "Transshipment TEUs"
        ]
      )
    };
  }

  function isDateHeader(name) {
    const normalized =
      normalizeHeaderName(name);

    return [
      "fecha",
      "date",
      "fechacombinada",
      "fechaevento"
    ].includes(normalized);
  }

  /*
   * IMPORTANTE:
   *
   * Los números NO se consideran fechas
   * únicamente por estar dentro del rango
   * de números seriales de Excel.
   *
   * Una fecha numérica solamente será DATE
   * si el encabezado indica que es una fecha.
   */
  function looksLikeDate(value) {
    if (value instanceof Date) {
      return true;
    }

    if (typeof value === "number") {
      return false;
    }

    const text =
      safeString(value).trim();

    if (!text) {
      return false;
    }

    return (
      /^\d{4}-\d{1,2}-\d{1,2}$/.test(text) ||
      /^\d{1,2}[\/-]\d{1,2}[\/-]\d{4}$/.test(text) ||
      /^[A-Za-zÁÉÍÓÚáéíóú]+[\/-]\d{4}$/.test(text)
    );
  }

  function looksLikeInteger(value) {
    if (typeof value === "bigint") {
      return true;
    }

    if (typeof value === "number") {
      return Number.isInteger(value);
    }

    const text =
      safeString(value)
        .trim()
        .replace(/,/g, "");

    if (!text) {
      return false;
    }

    return /^-?\d+$/.test(text);
  }

  function inferColumnType(
    columnIndex,
    columnName,
    rows
  ) {
    const values =
      rows
        .map(row =>
          row[columnIndex]
        )
        .filter(
          value =>
            !isEmptyValue(value)
        );

    const nonEmptyValues =
      values.length;

    if (!nonEmptyValues) {
      return {
        column_index: columnIndex,
        column_name: columnName,
        duckdb_type: "VARCHAR",
        sqlType: "VARCHAR",
        confidence: "low",
        reason: "columna sin valores",
        non_empty_values: 0
      };
    }

    /*
     * PRIMERO:
     * Si el encabezado es fecha,
     * podemos aceptar fechas numéricas
     * de Excel.
     */
    if (isDateHeader(columnName)) {
      const dateCompatible =
        values.every(
          value =>
            value instanceof Date ||
            typeof value === "number" ||
            looksLikeDate(value)
        );

      if (dateCompatible) {
        return {
          column_index: columnIndex,
          column_name: columnName,
          duckdb_type: "DATE",
          sqlType: "DATE",
          confidence: "high",
          reason:
            "encabezado de fecha y valores compatibles",
          non_empty_values:
            nonEmptyValues
        };
      }
    }

    /*
     * DESPUÉS:
     * Detectamos enteros.
     *
     * Esto evita que:
     * Local = 87840
     * Transshipment = 31391
     * Total TEUs = 119231
     *
     * sean interpretados como fechas.
     */
    if (
      values.every(
        looksLikeInteger
      )
    ) {
      return {
        column_index: columnIndex,
        column_name: columnName,
        duckdb_type: "BIGINT",
        sqlType: "BIGINT",
        confidence: "high",
        reason:
          "todos los valores son enteros",
        non_empty_values:
          nonEmptyValues
      };
    }

    const numeric =
      values.every(value => {
        if (
          typeof value === "number"
        ) {
          return Number.isFinite(
            value
          );
        }

        const text =
          safeString(value)
            .replace(/,/g, "")
            .trim();

        return (
          text !== "" &&
          Number.isFinite(
            Number(text)
          )
        );
      });

    if (numeric) {
      return {
        column_index: columnIndex,
        column_name: columnName,
        duckdb_type: "DOUBLE",
        sqlType: "DOUBLE",
        confidence: "high",
        reason:
          "todos los valores son numéricos",
        non_empty_values:
          nonEmptyValues
      };
    }

    return {
      column_index: columnIndex,
      column_name: columnName,
      duckdb_type: "VARCHAR",
      sqlType: "VARCHAR",
      confidence: "medium",
      reason:
        "valores tratados como texto",
      non_empty_values:
        nonEmptyValues
    };
  }

  function excelSerialToDate(serial) {
    const n = Number(serial);

    if (!Number.isFinite(n)) {
      return null;
    }

    const epoch =
      Date.UTC(1899, 11, 30);

    return new Date(
      epoch +
      Math.round(
        n * 86400000
      )
    );
  }

  function normalizeExcelDate(value) {
    if (value instanceof Date) {
      return value;
    }

    if (
      typeof value === "number" &&
      value > 20000 &&
      value < 80000
    ) {
      return excelSerialToDate(value);
    }

    const text =
      safeString(value).trim();

    if (!text) {
      return null;
    }

    if (
      /^\d{4}-\d{1,2}-\d{1,2}$/.test(text)
    ) {
      const date =
        new Date(
          `${text}T00:00:00Z`
        );

      return Number.isNaN(
        date.getTime()
      )
        ? null
        : date;
    }

    const date =
      new Date(text);

    return Number.isNaN(
      date.getTime()
    )
      ? null
      : date;
  }

  function dateToSql(date) {
    if (!(date instanceof Date)) {
      return null;
    }

    return (
      date.getUTCFullYear() +
      "-" +
      String(
        date.getUTCMonth() + 1
      ).padStart(2, "0") +
      "-" +
      String(
        date.getUTCDate()
      ).padStart(2, "0")
    );
  }

  function valueToSql(
    value,
    type
  ) {
    if (isEmptyValue(value)) {
      return "NULL";
    }

    if (type === "DATE") {
      const date =
        normalizeExcelDate(value);

      if (!date) {
        return "NULL";
      }

      return (
        "DATE '" +
        dateToSql(date) +
        "'"
      );
    }

    if (
      type === "BIGINT" ||
      type === "DOUBLE"
    ) {
      const n =
        Number(
          safeString(value)
            .replace(/,/g, "")
            .trim()
        );

      if (!Number.isFinite(n)) {
        return "NULL";
      }

      return type === "BIGINT"
        ? String(Math.trunc(n))
        : String(n);
    }

    return (
      "'" +
      escapeSqlString(value) +
      "'"
    );
  }

  function buildCreateTableSql(
    headers,
    types
  ) {
    const columns =
      headers.map(
        (header, index) =>
          `${escapeSqlIdentifier(
            header
          )} ${
            types[index]?.sqlType ||
            "VARCHAR"
          }`
      );

    return `
      CREATE TABLE excel_data (
        ${columns.join(",\n")}
      );
    `;
  }

  function buildInsertSql(
    headers,
    rows,
    types
  ) {
    const columnSql =
      headers
        .map(
          escapeSqlIdentifier
        )
        .join(", ");

    const values =
      rows.map(row =>
        "(" +
        headers
          .map(
            (_, index) =>
              valueToSql(
                row[index],
                types[index]
                  ?.sqlType ||
                  "VARCHAR"
              )
          )
          .join(", ") +
        ")"
      );

    return `
      INSERT INTO excel_data
      (${columnSql})
      VALUES
      ${values.join(",\n")};
    `;
  }

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
    if (db && conn) {
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
      new Worker(workerURL);

    db =
      new duckdb.AsyncDuckDB(
        new duckdb.ConsoleLogger(),
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

    await conn.query(
      "DROP TABLE IF EXISTS excel_data;"
    );
  }

  async function queryRows(sql) {
    const result =
      await conn.query(sql);

    return normalizeRows(
      result.toArray()
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

  /*
   * CORRECCIÓN IMPORTANTE:
   *
   * DuckDB 1.1.1 usa DATA_TYPE,
   * no COLUMN_TYPE en
   * information_schema.columns.
   */
  async function getSchema() {
    return queryRows(`
      SELECT
        column_name,
        data_type AS column_type,
        is_nullable AS null,
        column_default AS default_value,
        ordinal_position
      FROM information_schema.columns
      WHERE table_name = 'excel_data'
      ORDER BY ordinal_position;
    `);
  }

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

    if (!matrix.length) {
      return {
        headers: [],
        rows: [],
        physicalRows: 0,
        realRows: 0,
        emptyRows: 0
      };
    }

    const headers =
      makeUniqueHeaders(
        matrix[0] || []
      );

    const data =
      matrix.slice(1);

    const normalized =
      data.map(row =>
        headers.map(
          (_, index) =>
            row?.[index] ??
            null
        )
      );

    const realRows =
      normalized.filter(
        row =>
          row.some(
            value =>
              !isEmptyValue(value)
          )
      );

    return {
      headers,
      rows: realRows,
      physicalRows:
        matrix.length,
      realRows:
        realRows.length,
      emptyRows:
        normalized.length -
        realRows.length
    };
  }

  async function loadExcel(file) {
    try {
      await loadLibraries();
      await createDuckDB();
      await resetDatabase();

      currentFile = file;

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

      const sheets =
        workbook.SheetNames || [];

      if (!sheets.length) {
        throw new Error(
          "El archivo Excel no contiene hojas."
        );
      }

      currentSheetName =
        sheets[0];

      const parsed =
        readWorksheet(
          workbook.Sheets[
            currentSheetName
          ]
        );

      currentHeaders =
        parsed.headers;

      currentRows =
        parsed.rows;

      if (!currentHeaders.length) {
        throw new Error(
          "No se encontraron encabezados."
        );
      }

      if (!currentRows.length) {
        throw new Error(
          "No se encontraron filas con datos."
        );
      }

      currentTypes =
        currentHeaders.map(
          (header, index) =>
            inferColumnType(
              index,
              header,
              currentRows
            )
        );

      await conn.query(
        buildCreateTableSql(
          currentHeaders,
          currentTypes
        )
      );

      await conn.query(
        buildInsertSql(
          currentHeaders,
          currentRows,
          currentTypes
        )
      );

      const count =
        await queryRows(`
          SELECT COUNT(*) AS registros
          FROM excel_data;
        `);

      const schema =
        await getSchema();

      const preview =
        await queryRows(`
          SELECT *
          FROM excel_data
          LIMIT 10;
        `);

      setResult({
        procesamiento: "LOCAL",
        engine: APP_ID,
        version: VERSION,
        archivo: file.name,
        tamano_bytes: file.size,
        hojas: sheets,
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
          currentTypes,
        columnas_especiales_detectadas:
          detectSpecialColumns(),
        duckdb_table:
          "excel_data",
        duckdb_version:
          await getDuckDBVersion(),
        count,
        schema,
        preview
      });

      updateStatus(
        `Excel cargado localmente: ${file.name} — ${currentRows.length} registros`
      );

    } catch (error) {
      showError(error);
      updateStatus(
        "Error al cargar Excel"
      );
    }
  }

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

      if (!sql?.trim()) {
        throw new Error(
          "Introduce una consulta SQL."
        );
      }

      const start =
        performance.now();

      const rows =
        await queryRows(sql);

      const elapsed =
        performance.now() -
        start;

      setResult({
        procesamiento: "LOCAL",
        sql,
        filas_resultado:
          rows.length,
        tiempo_ms:
          Number(
            elapsed.toFixed(1)
          ),
        resultado: rows
      });

      updateStatus(
        `${label} ejecutado localmente`
      );

      return rows;

    } catch (error) {
      showError(error);

      updateStatus(
        `Error en ${label}`
      );
    }
  }

  async function countRows() {
    return executeSql(
      `
      SELECT COUNT(*) AS registros
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
      const schema =
        await getSchema();

      const numeric =
        schema.filter(
          c =>
            /BIGINT|DOUBLE|DECIMAL|INTEGER|HUGEINT|FLOAT|SMALLINT/i.test(
              safeString(
                c.column_type
              )
            )
        );

      if (!numeric.length) {
        throw new Error(
          "No se encontraron columnas numéricas."
        );
      }

      const expressions =
        numeric.flatMap(c => {
          const q =
            escapeSqlIdentifier(
              c.column_name
            );

          return [
            `SUM(${q}) AS ${escapeSqlIdentifier(
              c.column_name +
              "__sum"
            )}`,
            `AVG(${q}) AS ${escapeSqlIdentifier(
              c.column_name +
              "__avg"
            )}`,
            `MIN(${q}) AS ${escapeSqlIdentifier(
              c.column_name +
              "__min"
            )}`,
            `MAX(${q}) AS ${escapeSqlIdentifier(
              c.column_name +
              "__max"
            )}`
          ];
        });

      return executeSql(
        `
        SELECT
          ${expressions.join(",\n")}
        FROM excel_data;
        `,
        "Resumen"
      );

    } catch (error) {
      showError(error);
    }
  }

  async function totalByYear() {
    const special =
      detectSpecialColumns();

    if (!special.año) {
      throw new Error(
        "No se encontró la columna Año."
      );
    }

    if (!special.total_teus) {
      throw new Error(
        "No se encontró la columna Total TEUs."
      );
    }

    const year =
      escapeSqlIdentifier(
        special.año
      );

    const total =
      escapeSqlIdentifier(
        special.total_teus
      );

    return executeSql(
      `
      SELECT
        ${year} AS año,
        SUM(${total}) AS total
      FROM excel_data
      GROUP BY ${year}
      ORDER BY ${year};
      `,
      "Total por año"
    );
  }

  async function totalByMonth() {
    const special =
      detectSpecialColumns();

    const month =
      detectColumn(
        currentHeaders,
        ["Mes", "Month"]
      );

    if (!special.año) {
      throw new Error(
        "No se encontró la columna Año."
      );
    }

    if (!month) {
      throw new Error(
        "No se encontró la columna Mes."
      );
    }

    if (!special.total_teus) {
      throw new Error(
        "No se encontró la columna Total TEUs."
      );
    }

    const year =
      escapeSqlIdentifier(
        special.año
      );

    const monthSql =
      escapeSqlIdentifier(
        month
      );

    const total =
      escapeSqlIdentifier(
        special.total_teus
      );

    return executeSql(
      `
      SELECT
        ${year} AS año,
        ${monthSql} AS mes,
        SUM(${total}) AS total
      FROM excel_data
      GROUP BY
        ${year},
        ${monthSql}
      ORDER BY
        ${year},
        ${monthSql};
      `,
      "Total por mes"
    );
  }

  async function totalByLocal() {
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

    const local =
      escapeSqlIdentifier(
        special.local
      );

    const total =
      escapeSqlIdentifier(
        special.total_teus
      );

    return executeSql(
      `
      SELECT
        ${local} AS local,
        SUM(${total}) AS total
      FROM excel_data
      GROUP BY ${local}
      ORDER BY total DESC;
      `,
      "Total por Local"
    );
  }

  async function showSchema() {
    try {
      setResult({
        procesamiento: "LOCAL",
        tabla: "excel_data",
        columnas:
          await getSchema(),
        columnas_especiales_detectadas:
          detectSpecialColumns()
      });
    } catch (error) {
      showError(error);
    }
  }

  async function runSqlFromEditor() {
    const editor =
      document.getElementById(
        "tm-excel-sql"
      );

    if (!editor) {
      return;
    }

    await executeSql(
      editor.value.trim(),
      "SQL"
    );
  }

  function updateStatus(message) {
    const status =
      document.getElementById(
        "tm-excel-status"
      );

    if (status) {
      status.textContent =
        message;
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

    panel.innerHTML = `
      <div style="
        font-family:Arial,sans-serif;
        width:100%;
        box-sizing:border-box;
      ">

        <div style="
          font-size:18px;
          font-weight:bold;
          margin-bottom:6px;
        ">
          📊 Excel Data Engine
        </div>

        <div style="
          font-size:12px;
          opacity:.75;
          margin-bottom:10px;
        ">
          ${VERSION} —
          Todo el procesamiento se realiza localmente.
        </div>

        <div id="tm-excel-status" style="
          padding:8px;
          margin-bottom:10px;
          border:1px solid #ddd;
          border-radius:6px;
        ">
          Motor listo. Carga un archivo Excel.
        </div>

        <input
          id="tm-excel-file"
          type="file"
          accept=".xlsx,.xls,.xlsm"
          style="width:100%;margin-bottom:10px;"
        />

        <div style="
          display:grid;
          grid-template-columns:repeat(2,minmax(0,1fr));
          gap:6px;
          margin-bottom:10px;
        ">

          <button id="tm-excel-count">
            COUNT
          </button>

          <button id="tm-excel-preview">
            Vista previa
          </button>

          <button id="tm-excel-summary">
            Resumen
          </button>

          <button id="tm-excel-year">
            Total por año
          </button>

          <button id="tm-excel-month">
            Total por mes
          </button>

          <button id="tm-excel-local">
            Total por Local
          </button>

          <button id="tm-excel-schema">
            Esquema
          </button>

        </div>

        <div style="
          font-weight:bold;
          margin-bottom:5px;
        ">
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
    SUM("Total TEUs") AS total
FROM excel_data
GROUP BY "Año"
ORDER BY "Año";</textarea>

        <button
          id="tm-excel-run-sql"
          style="
            margin-top:6px;
            width:100%;
          "
        >
          Ejecutar SQL
        </button>

        <div style="
          font-weight:bold;
          margin-top:12px;
          margin-bottom:5px;
        ">
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

    const host =
      document.querySelector(
        "#right-sidebar"
      ) ||
      document.querySelector(
        "[data-testid='right-sidebar']"
      ) ||
      document.querySelector(
        ".right-sidebar"
      );

    if (host) {
      host.appendChild(panel);
    } else {
      Object.assign(
        panel.style,
        {
          position: "fixed",
          right: "10px",
          top: "80px",
          width: "420px",
          maxHeight:
            "calc(100vh - 100px)",
          overflow: "auto",
          zIndex: "999999",
          background: "white",
          padding: "12px",
          border: "1px solid #ccc",
          borderRadius: "8px",
          boxShadow:
            "0 4px 20px rgba(0,0,0,.15)"
        }
      );

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
        async event => {
          const file =
            event.target.files?.[0];

          if (file) {
            await loadExcel(file);
          }
        }
      );

    document
      .getElementById(
        "tm-excel-count"
      )
      .onclick = countRows;

    document
      .getElementById(
        "tm-excel-preview"
      )
      .onclick = previewRows;

    document
      .getElementById(
        "tm-excel-summary"
      )
      .onclick = summary;

    document
      .getElementById(
        "tm-excel-year"
      )
      .onclick = totalByYear;

    document
      .getElementById(
        "tm-excel-month"
      )
      .onclick = totalByMonth;

    document
      .getElementById(
        "tm-excel-local"
      )
      .onclick = totalByLocal;

    document
      .getElementById(
        "tm-excel-schema"
      )
      .onclick = showSchema;

    document
      .getElementById(
        "tm-excel-run-sql"
      )
      .onclick =
      runSqlFromEditor;
  }

  async function initialize() {
    try {
      createUI();

      updateStatus(
        `Inicializando ${VERSION}...`
      );

      await loadLibraries();
      await createDuckDB();

      updateStatus(
        `${VERSION} listo — DuckDB-Wasm cargado localmente`
      );

      window.TMExcelEngine = {
        version: VERSION,
        app_id: APP_ID,
        executeSql,
        loadExcel,
        countRows,
        previewRows,
        summary,
        totalByYear,
        totalByMonth,
        totalByLocal,
        showSchema,
        detectSpecialColumns,
        getSchema
      };

    } catch (error) {
      showError(error);

      updateStatus(
        "No se pudo inicializar el motor"
      );
    }
  }

  if (
    window.__TM_EXCEL_ENGINE_0414_INITIALIZED
  ) {
    return;
  }

  window.__TM_EXCEL_ENGINE_0414_INITIALIZED =
    true;

  initialize();

})();
