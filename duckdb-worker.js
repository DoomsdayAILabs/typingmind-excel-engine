/*
 * DuckDB Web Worker Engine — TypingMind Excel Engine
 * FASE 1: Carga nativa asíncrona CSV y Parquet + Ejecución SQL
 *
 * Mantiene la ejecución SQL y descompresión/carga de archivos pesados
 * en un hilo secundario (Web Worker) para evitar congelar la UI de TypingMind.
 */

const DUCKDB_PACKAGE = "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

let duckdb = null;
let db = null;
let conn = null;
let isInitialized = false;

/* ============================================================
 * NORMALIZACIÓN DE VALORES Y RESULTADOS (DE v0.4.23)
 * ============================================================ */

function safeString(value) {
  return value === null || value === undefined ? "" : String(value);
}

function normalizeSigned32Number(value) {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return value;
  }
  if (Number.isInteger(value) && value >= 4294967295 - 2147483648) {
    return value - 4294967296;
  }
  return value;
}

function normalizeNumericArray(value) {
  if (!Array.isArray(value) || value.length !== 4) {
    return null;
  }
  const nums = value.map((v) => {
    if (typeof v === "bigint") return Number(v);
    if (typeof v === "number") return v;
    return Number(v);
  });
  if (!nums.every(Number.isFinite)) {
    return null;
  }
  const signed = nums.map(normalizeSigned32Number);

  // Patrón HUGEINT/Decimal Arrow: [valor, 0, 0, 0]
  if (signed[1] === 0 && signed[2] === 0 && signed[3] === 0) {
    return signed[0];
  }
  // Patrón HUGEINT/Decimal Arrow negativo: [valor, -1, -1, -1]
  if (signed[1] === -1 && signed[2] === -1 && signed[3] === -1) {
    return signed[0];
  }
  return null;
}

function normalizeValue(value, key = "") {
  if (value === null || value === undefined) {
    return null;
  }
  if (typeof value === "bigint") {
    const n = Number(value);
    return Number.isSafeInteger(n) ? n : value.toString();
  }
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? null : value.toISOString();
  }
  if (value instanceof ArrayBuffer) {
    return Array.from(new Uint8Array(value));
  }
  if (ArrayBuffer.isView(value)) {
    const arr = Array.from(value);
    const scalar = normalizeNumericArray(arr);
    if (scalar !== null) return scalar;
    return arr.map((v) => normalizeSigned32Number(v));
  }
  if (Array.isArray(value)) {
    const scalar = normalizeNumericArray(value);
    if (scalar !== null) return scalar;
    return value.map((v) => normalizeValue(v, key));
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
  if (!Array.isArray(rows)) return [];
  return rows.map((row) => normalizeValue(row));
}

function escapeSqlIdentifier(name) {
  return '"' + safeString(name).replace(/"/g, '""') + '"';
}

function escapeSqlString(value) {
  return safeString(value).replace(/'/g, "''");
}

/* ============================================================
 * INICIALIZACIÓN DE DUCKDB EN WORKER
 * ============================================================ */

async function initDuckDB() {
  if (isInitialized && db && conn) {
    return { initialized: true, already: true };
  }

  duckdb = await import(DUCKDB_PACKAGE);
  const bundles = duckdb.getJsDelivrBundles();
  const bundle = bundles.eh;

  const workerURL = URL.createObjectURL(
    new Blob([`importScripts("${bundle.mainWorker}");`], {
      type: "application/javascript"
    })
  );

  const worker = new Worker(workerURL);
  const logger = new duckdb.ConsoleLogger();

  db = new duckdb.AsyncDuckDB(logger, worker);
  await db.instantiate(bundle.mainModule, bundle.pthreadWorker);
  conn = await db.connect();

  isInitialized = true;

  const versionResult = await conn.query("SELECT * FROM pragma_version();");
  const duckdbVersion = normalizeRows(versionResult.toArray());

  return {
    initialized: true,
    duckdb_version: duckdbVersion
  };
}

/* ============================================================
 * 1. executeQuery — EJECUCIÓN LIBRE DE SQL
 * ============================================================ */

async function executeQuery(sql) {
  if (!conn) {
    await initDuckDB();
  }
  if (!sql || !sql.trim()) {
    throw new Error("Se requiere una consulta SQL válida.");
  }

  const start = performance.now();
  const result = await conn.query(sql);
  const rows = normalizeRows(result.toArray());
  const elapsed = performance.now() - start;

  return {
    sql: sql,
    filas_resultado: rows.length,
    tiempo_ms: Number(elapsed.toFixed(1)),
    resultado: rows
  };
}

/* ============================================================
 * 2. loadCSV — CARGA NATIVA DE ARCHIVOS CSV (Soporta Millones)
 * ============================================================ */

async function loadCSV(fileName, buffer, tableName = "excel_data", options = {}) {
  if (!conn) {
    await initDuckDB();
  }
  if (!buffer) {
    throw new Error("No se proporcionó el buffer del archivo CSV.");
  }

  const cleanFileName = safeString(fileName || "datos.csv").replace(/[^a-zA-Z0-9._-]/g, "_");
  const cleanTableName = escapeSqlIdentifier(tableName || "excel_data");

  const uint8Data = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  await db.registerFileBuffer(cleanFileName, uint8Data);

  const start = performance.now();

  await conn.query(`DROP TABLE IF EXISTS ${cleanTableName};`);

  // Configuración de opciones nativas para read_csv_auto
  let extraOptions = "";
  if (options.header === false) {
    extraOptions += ", header=false";
  }
  if (options.delim) {
    extraOptions += `, delim='${escapeSqlString(options.delim)}'`;
  }

  const createSql = `
    CREATE TABLE ${cleanTableName} AS
    SELECT * FROM read_csv_auto('${cleanFileName}'${extraOptions});
  `;

  await conn.query(createSql);
  const elapsed = performance.now() - start;

  const countResult = await conn.query(`SELECT COUNT(*) AS registros FROM ${cleanTableName};`);
  const countRows = normalizeRows(countResult.toArray());

  const schemaResult = await conn.query(`
    SELECT column_name, data_type
    FROM information_schema.columns
    WHERE table_name = '${tableName || "excel_data"}'
    ORDER BY ordinal_position;
  `);
  const schema = normalizeRows(schemaResult.toArray());

  const previewResult = await conn.query(`SELECT * FROM ${cleanTableName} LIMIT 10;`);
  const preview = normalizeRows(previewResult.toArray());

  return {
    procesamiento: "LOCAL_NATIVO_WORKER",
    formato: "CSV",
    archivo: cleanFileName,
    tabla: tableName || "excel_data",
    tamano_bytes: uint8Data.length,
    registros: countRows[0]?.registros ?? 0,
    tiempo_carga_ms: Number(elapsed.toFixed(1)),
    columnas: schema.map((c) => c.column_name),
    esquema: schema,
    preview: preview
  };
}

/* ============================================================
 * 3. loadParquet — CARGA NATIVA DE ARCHIVOS PARQUET
 * ============================================================ */

async function loadParquet(fileName, buffer, tableName = "excel_data") {
  if (!conn) {
    await initDuckDB();
  }
  if (!buffer) {
    throw new Error("No se proporcionó el buffer del archivo Parquet.");
  }

  const cleanFileName = safeString(fileName || "datos.parquet").replace(/[^a-zA-Z0-9._-]/g, "_");
  const cleanTableName = escapeSqlIdentifier(tableName || "excel_data");

  const uint8Data = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  await db.registerFileBuffer(cleanFileName, uint8Data);

  const start = performance.now();

  await conn.query(`DROP TABLE IF EXISTS ${cleanTableName};`);

  const createSql = `
    CREATE TABLE ${cleanTableName} AS
    SELECT * FROM read_parquet('${cleanFileName}');
  `;

  await conn.query(createSql);
  const elapsed = performance.now() - start;

  const countResult = await conn.query(`SELECT COUNT(*) AS registros FROM ${cleanTableName};`);
  const countRows = normalizeRows(countResult.toArray());

  const schemaResult = await conn.query(`
    SELECT column_name, data_type
    FROM information_schema.columns
    WHERE table_name = '${tableName || "excel_data"}'
    ORDER BY ordinal_position;
  `);
  const schema = normalizeRows(schemaResult.toArray());

  const previewResult = await conn.query(`SELECT * FROM ${cleanTableName} LIMIT 10;`);
  const preview = normalizeRows(previewResult.toArray());

  return {
    procesamiento: "LOCAL_NATIVO_WORKER",
    formato: "PARQUET",
    archivo: cleanFileName,
    tabla: tableName || "excel_data",
    tamano_bytes: uint8Data.length,
    registros: countRows[0]?.registros ?? 0,
    tiempo_carga_ms: Number(elapsed.toFixed(1)),
    columnas: schema.map((c) => c.column_name),
    esquema: schema,
    preview: preview
  };
}

/* ============================================================
 * DISPATCHER DE MENSAJES VÍA postMessage
 * ============================================================ */

self.onmessage = async (event) => {
  const { id, type, ...payload } = event.data || {};

  if (!id) {
    return;
  }

  try {
    let result = null;

    switch (type) {
      case "init":
        result = await initDuckDB();
        break;

      case "executeQuery":
        result = await executeQuery(payload.sql);
        break;

      case "loadCSV":
        result = await loadCSV(
          payload.fileName,
          payload.buffer,
          payload.tableName,
          payload.options
        );
        break;

      case "loadParquet":
        result = await loadParquet(
          payload.fileName,
          payload.buffer,
          payload.tableName
        );
        break;

      default:
        throw new Error(`Tipo de acción desconocido en worker: ${type}`);
    }

    self.postMessage({
      id,
      type: "response",
      success: true,
      result
    });
  } catch (error) {
    self.postMessage({
      id,
      type: "response",
      success: false,
      error: error?.message || String(error),
      stack: error?.stack || ""
    });
  }
};
