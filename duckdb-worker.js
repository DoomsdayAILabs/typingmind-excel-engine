"use strict";

const WORKER_VERSION = "v1.0-phase-1b+multisheet";
const DUCKDB_PACKAGE = "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";
const SHEETJS_PACKAGE = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";

importScripts(SHEETJS_PACKAGE);

let db = null;
let conn = null;
let runtimeWorker = null;
let runtimeWorkerUrl = null;

async function initDuckDB() {
  if (db && conn) {
    return {
      version: WORKER_VERSION,
      status: "ready"
    };
  }

  const duckdb = await import(DUCKDB_PACKAGE);

  if (
    typeof duckdb.getJsDelivrBundles !== "function" ||
    typeof duckdb.selectBundle !== "function" ||
    typeof duckdb.AsyncDuckDB !== "function"
  ) {
    throw new Error("DuckDB-Wasm no se importó correctamente");
  }

  const bundles = duckdb.getJsDelivrBundles();
  const bundle = await duckdb.selectBundle(bundles);

  runtimeWorkerUrl = URL.createObjectURL(
    new Blob(
      [`importScripts("${bundle.mainWorker}");`],
      { type: "text/javascript" }
    )
  );

  runtimeWorker = new Worker(runtimeWorkerUrl);
  const logger = new duckdb.ConsoleLogger();
  db = new duckdb.AsyncDuckDB(logger, runtimeWorker);

  await db.instantiate(bundle.mainModule, bundle.pthreadWorker);
  conn = await db.connect();



  return {
    version: WORKER_VERSION,
    status: "ready",
    bundle: {
      mainModule: bundle.mainModule,
      mainWorker: bundle.mainWorker,
      pthreadWorker: bundle.pthreadWorker ?? null
    }
  };
}

function bigintToSafeOutput(value) {
  if (value >= BigInt(Number.MIN_SAFE_INTEGER) && value <= BigInt(Number.MAX_SAFE_INTEGER)) {
    return Number(value);
  }
  return value.toString();
}

function maybeNormalizeHugeIntArray(arr) {
  if (!Array.isArray(arr) || arr.length !== 4) {
    return null;
  }

  const normalized = arr.map((v) => {
    if (typeof v === "bigint") return v;
    if (typeof v === "number" && Number.isFinite(v)) return BigInt(Math.trunc(v));
    return null;
  });

  if (normalized.some((v) => v === null)) {
    return null;
  }

  const [a, b, c, d] = normalized;

  const isPositivePattern = b === 0n && c === 0n && d === 0n;
  const isNegativePattern = b === -1n && c === -1n && d === -1n;

  if (!isPositivePattern && !isNegativePattern) {
    return null;
  }

  return bigintToSafeOutput(a);
}

function normalizeValue(value) {
  if (value === null || value === undefined) {
    return null;
  }

  if (typeof value === "bigint") {
    return bigintToSafeOutput(value);
  }

  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? null : value.toISOString();
  }

  if (value instanceof ArrayBuffer) {
    return Array.from(new Uint8Array(value));
  }

  if (typeof DataView !== "undefined" && value instanceof DataView) {
    return Array.from(new Uint8Array(value.buffer, value.byteOffset, value.byteLength));
  }

  if (ArrayBuffer.isView(value)) {
    const arr = Array.from(value);
    const hugeint = maybeNormalizeHugeIntArray(arr);
    if (hugeint !== null) {
      return hugeint;
    }
    return arr.map((item) => normalizeValue(item));
  }

  if (Array.isArray(value)) {
    const hugeint = maybeNormalizeHugeIntArray(value);
    if (hugeint !== null) {
      return hugeint;
    }
    return value.map((item) => normalizeValue(item));
  }

  if (typeof value === "object") {
    const out = {};
    for (const [key, val] of Object.entries(value)) {
      out[key] = normalizeValue(val);
    }
    return out;
  }

  return value;
}

function normalizeRows(rows) {
  if (!Array.isArray(rows)) {
    return normalizeValue(rows);
  }
  return rows.map((row) => normalizeValue(row));
}

async function executeQuery(sql) {
  if (!conn) {
    throw new Error("DuckDB no está inicializado");
  }

  if (typeof sql !== "string" || !sql.trim()) {
    throw new Error("SQL inválido");
  }

  const result = await conn.query(sql);
  const rows = result.toArray();
  return normalizeRows(rows);
}
function sanitizeTableName(tableName) {
  const name = String(tableName || "excel_data");
  const sanitized = name.replace(/[^A-Za-z0-9_]/g, "_");
  if (!sanitized) return "excel_data";
  if (/^[0-9]/.test(sanitized)) return `t_${sanitized}`;
  return sanitized;
}

function quoteIdentifier(identifier) {
  const str = String(identifier);
  return `"${str.replace(/"/g, '""')}"`;
}

function quoteStringLiteral(value) {
  return `'${String(value).replace(/'/g, "''")}'`;
}

async function loadCSV(csvData, tableName, requestId) {
  if (!db || !conn) {
    throw new Error("DuckDB no está inicializado");
  }

  if (typeof csvData !== "string" || !csvData.trim()) {
    throw new Error("CSV inválido");
  }

  const virtualName = `upload_${requestId}.csv`;
  let fileRegistered = false;
  let mainError = null;
  let result;

  try {
    await db.registerFileText(virtualName, csvData);
    fileRegistered = true;

    const sanitizedTable = sanitizeTableName(tableName);
    const quotedTable = quoteIdentifier(sanitizedTable);

    await conn.query(
      `CREATE OR REPLACE TABLE ${quotedTable} AS
       SELECT * FROM read_csv_auto(${quoteStringLiteral(virtualName)}, sample_size=-1, ignore_errors=true)`
    );

    const countResult = await conn.query(`SELECT COUNT(*) AS registros FROM ${quotedTable}`);
    const countRows = normalizeRows(countResult.toArray());
    const registros = countRows[0]?.registros ?? 0;

    const schemaResult = await conn.query(
      `SELECT column_name, data_type
       FROM information_schema.columns
       WHERE table_name = ${quoteStringLiteral(sanitizedTable)}
       ORDER BY ordinal_position`
    );
    const schemaRows = schemaResult.toArray();

    const previewResult = await conn.query(`SELECT * FROM ${quotedTable} LIMIT 10`);
    const previewRows = previewResult.toArray();

    result = {
      procesamiento: "LOCAL_NATIVO_WORKER",
      formato: "CSV",
      tabla: sanitizedTable,
      registros,
      columnas: schemaRows.map(row => row.column_name),
      esquema: normalizeRows(schemaRows),
      preview: normalizeRows(previewRows)
    };
  } catch (error) {
    mainError = error;
  } finally {
    if (fileRegistered) {
      try {
        await db.dropFile(virtualName);
      } catch (unregisterError) {
        if (!mainError) {
          throw unregisterError;
        }
      }
    }
  }

  if (mainError) {
    throw mainError;
  }

  return result;
}
async function loadParquet(parquetData, tableName, requestId) {
  if (!db || !conn) {
    throw new Error("DuckDB no está inicializado");
  }

  if (!parquetData) {
    throw new Error("Parquet data inválida");
  }

  const virtualName = `upload_${requestId}.parquet`;
  let fileRegistered = false;
  let mainError = null;
  let result;

  try {
    const uint8Array = parquetData instanceof Uint8Array ? parquetData : new Uint8Array(parquetData);
    await db.registerFileBuffer(virtualName, uint8Array);
    fileRegistered = true;

    const sanitizedTable = sanitizeTableName(tableName);
    const quotedTable = quoteIdentifier(sanitizedTable);

    await conn.query(
      `CREATE OR REPLACE TABLE ${quotedTable} AS SELECT * FROM read_parquet(${quoteStringLiteral(virtualName)})`
    );

    const countResult = await conn.query(`SELECT COUNT(*) AS registros FROM ${quotedTable}`);
    const countRows = normalizeRows(countResult.toArray());
    const registros = countRows[0]?.registros ?? 0;

    const schemaResult = await conn.query(
      `SELECT column_name, data_type
       FROM information_schema.columns
       WHERE table_name = ${quoteStringLiteral(sanitizedTable)}
       ORDER BY ordinal_position`
    );
    const schemaRows = schemaResult.toArray();

    const previewResult = await conn.query(`SELECT * FROM ${quotedTable} LIMIT 10`);
    const previewRows = previewResult.toArray();

    result = {
      procesamiento: "LOCAL_NATIVO_WORKER",
      formato: "PARQUET",
      tabla: sanitizedTable,
      registros,
      columnas: schemaRows.map(row => row.column_name),
      esquema: normalizeRows(schemaRows),
      preview: normalizeRows(previewRows)
    };
  } catch (error) {
    mainError = error;
  } finally {
    if (fileRegistered) {
      try {
        await db.dropFile(virtualName);
      } catch (unregisterError) {
        if (!mainError) {
          throw unregisterError;
        }
      }
    }
  }

  if (mainError) {
    throw mainError;
  }

  return result;
}

// --- Soporte multi-hoja (Excel): una tabla independiente por pestaña ---

/**
 * Sufijo determinista para la tabla de una hoja. Aplica las mismas reglas de saneado
 * que `sanitizeTableName` (no alfanuméricos -> "_", prefijo `t_` si empieza por dígito)
 * y cae a la posición de la hoja cuando el nombre no aporta ningún carácter útil.
 */
function sanitizeSheetSuffix(sheetName, index) {
  const limpio = String(sheetName ?? "")
    .replace(/[^A-Za-z0-9_]/g, "_")
    .replace(/^_+|_+$/g, "");

  if (!limpio) {
    return `hoja_${index + 1}`;
  }

  return /^[0-9]/.test(limpio) ? `t_${limpio}` : limpio;
}

/**
 * Tabla destino de una hoja: `<base>_<sufijo>`. DuckDB compara identificadores sin
 * distinguir mayúsculas, así que la deduplicación es case-insensitive y añade un
 * contador incremental (_2, _3…) cuando dos pestañas colisionan tras sanearse.
 */
function buildSheetTableName(baseTable, sheetName, index, usados) {
  const sufijo = sanitizeSheetSuffix(sheetName, index);
  let candidato = `${baseTable}_${sufijo}`;
  let contador = 1;

  while (usados.has(candidato.toLowerCase())) {
    contador += 1;
    candidato = `${baseTable}_${sufijo}_${contador}`;
  }

  usados.add(candidato.toLowerCase());

  return candidato;
}

/** Entrada homogénea de `hojas[]` para cualquier desenlace (ok / omitida / error). */
function entradaHoja(nombreHoja, carga) {
  const entrada = {
    nombre: nombreHoja,
    tabla: carga.tabla ?? null,
    registros: carga.registros ?? 0,
    columnas: carga.columnas ?? [],
    estado: carga.estado
  };

  if (carga.motivo) entrada.motivo = carga.motivo;
  if (carga.error) entrada.error = carga.error;
  if (carga.estado === "ok" && !carga.registros) entrada.advertencia = "solo_cabeceras";

  return entrada;
}

/** Descarta una tabla creada a medias sin enmascarar el error de la hoja. */
async function descartarTablaParcial(tableName) {
  try {
    await conn.query(`DROP TABLE IF EXISTS ${quoteIdentifier(tableName)}`);
  } catch (error) {
    // Best-effort: la hoja ya se reporta como error y la limpieza no puede abortar la carga.
  }
}

/**
 * Materializa una pestaña como tabla propia (CSV temporal en el VFS -> read_csv_auto).
 * Nunca lanza: devuelve el desenlace de esa hoja para que el resto del libro siga cargando.
 * `sample_size=-1` + `ignore_errors=true` alinean la inferencia con la ruta CSV.
 */
async function materializarHojaExcel(worksheet, tableName, virtualName, ficherosVFS) {
  const csvData = XLSX.utils.sheet_to_csv(worksheet);

  if (!csvData || !csvData.trim()) {
    return { estado: "omitida", motivo: "hoja_sin_datos" };
  }

  await db.registerFileBuffer(virtualName, new TextEncoder().encode(csvData));
  ficherosVFS.push(virtualName);

  const quotedTable = quoteIdentifier(tableName);
  let tablaCreada = false;

  try {
    await conn.query(
      `CREATE OR REPLACE TABLE ${quotedTable} AS
       SELECT * FROM read_csv_auto(${quoteStringLiteral(virtualName)}, sample_size=-1, ignore_errors=true)`
    );
    tablaCreada = true;

    const countResult = await conn.query(`SELECT COUNT(*) AS registros FROM ${quotedTable}`);
    const countRows = normalizeRows(countResult.toArray());
    const registros = countRows[0]?.registros ?? 0;

    const schemaResult = await conn.query(
      `SELECT column_name, data_type
       FROM information_schema.columns
       WHERE table_name = ${quoteStringLiteral(tableName)}
       ORDER BY ordinal_position`
    );

    return {
      estado: "ok",
      tabla: tableName,
      registros,
      columnas: schemaResult.toArray().map(row => row.column_name)
    };
  } catch (error) {
    if (tablaCreada) {
      await descartarTablaParcial(tableName);
    }

    return {
      estado: "error",
      motivo: "hoja_error_tabla",
      error: error?.message ? error.message : String(error)
    };
  }
}

async function loadExcel(excelBuffer, tableName, requestId) {
  if (!db || !conn) {
    throw new Error("DuckDB no está inicializado");
  }

  if (!excelBuffer) {
    throw new Error("Excel data inválida");
  }

  const ficherosVFS = [];
  let mainError = null;
  let result;

  try {
    const workbook = XLSX.read(excelBuffer, { type: "array" });
    const nombresHojas = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];

    if (nombresHojas.length === 0) {
      throw new Error("El archivo Excel no contiene hojas");
    }

    const baseTable = sanitizeTableName(tableName);
    const nombresUsados = new Set();
    const hojas = [];

    // Se itera el libro en su orden natural, hoja a hoja y de forma secuencial.
    for (let index = 0; index < nombresHojas.length; index++) {
      const nombreHoja = nombresHojas[index];
      const worksheet = workbook.Sheets ? workbook.Sheets[nombreHoja] : null;

      if (!worksheet || !worksheet["!ref"]) {
        hojas.push(entradaHoja(nombreHoja, { estado: "omitida", motivo: "hoja_vacia" }));
        continue;
      }

      // Variante A: la hoja 0 ocupa el nombre base recibido (excel_data) para conservar
      // la compatibilidad con el plugin y las suites SQL; las hojas 1..N se registran
      // como excel_data_<hoja saneada>. Si la hoja 0 está vacía o falla no existirá
      // `excel_data`: el widget sigue operativo porque usa la primera tabla cargada.
      const tablaDestino = index === 0
        ? baseTable
        : buildSheetTableName(baseTable, nombreHoja, index, nombresUsados);

      let carga;

      try {
        carga = await materializarHojaExcel(
          worksheet,
          tablaDestino,
          `upload_${requestId}_${index}.csv`,
          ficherosVFS
        );
      } catch (error) {
        // Red de seguridad: el fallo de una hoja jamás aborta la carga del libro.
        carga = {
          estado: "error",
          motivo: "hoja_error_vfs",
          error: error?.message ? error.message : String(error)
        };
      }

      hojas.push(entradaHoja(nombreHoja, carga));
    }

    const hojasOk = hojas.filter((hoja) => hoja.estado === "ok");

    if (hojasOk.length === 0) {
      throw new Error("El archivo Excel no contiene hojas con datos");
    }

    const primaria = hojasOk[0];
    const quotedPrimaria = quoteIdentifier(primaria.tabla);

    const schemaResult = await conn.query(
      `SELECT column_name, data_type
       FROM information_schema.columns
       WHERE table_name = ${quoteStringLiteral(primaria.tabla)}
       ORDER BY ordinal_position`
    );
    const schemaRows = schemaResult.toArray();

    const previewResult = await conn.query(`SELECT * FROM ${quotedPrimaria} LIMIT 10`);
    const previewRows = previewResult.toArray();

    result = {
      procesamiento: "LOCAL_NATIVO_WORKER",
      formato: "EXCEL",
      tabla: primaria.tabla,
      registros: primaria.registros,
      columnas: schemaRows.map(row => row.column_name),
      esquema: normalizeRows(schemaRows),
      preview: normalizeRows(previewRows),
      hojas,
      hojas_totales: nombresHojas.length,
      hojas_cargadas: hojasOk.length,
      hojas_omitidas: hojas.filter(hoja => hoja.estado === "omitida").length,
      hojas_con_error: hojas.filter(hoja => hoja.estado === "error").length
    };
  } catch (error) {
    mainError = error;
  } finally {
    for (const virtualName of ficherosVFS) {
      try {
        await db.dropFile(virtualName);
      } catch (unregisterError) {
        if (!mainError) {
          mainError = unregisterError;
        }
      }
    }
  }

  if (mainError) {
    throw mainError;
  }

  return result;
}

self.onmessage = async (event) => {
  const { type, requestId, ...payload } = event.data ?? {};

  try {
    let data;

    switch (type) {
      case "initDuckDB":
        data = await initDuckDB();
        break;
      case "executeQuery":
        data = await executeQuery(payload.sql);
        break;
      case "loadCSV":
        data = await loadCSV(payload.csvData, payload.tableName, requestId);
        break;
      case "loadParquet":
        data = await loadParquet(payload.parquetData, payload.tableName, requestId);
        break;
      case "loadExcel":
        data = await loadExcel(payload.excelBuffer, payload.tableName, requestId);
        break;
      default:
        throw new Error(`Tipo de mensaje no reconocido: ${String(type)}`);
    }

    self.postMessage({
      success: true,
      requestId,
      data
    });
  } catch (error) {
    self.postMessage({
      success: false,
      requestId,
      error: error?.message ? error.message : String(error)
    });
  }
};
