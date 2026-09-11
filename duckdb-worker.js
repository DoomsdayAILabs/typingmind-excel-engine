"use strict";

const WORKER_VERSION = "v1.0-phase-1a";
const DUCKDB_PACKAGE = "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

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
      case "loadParquet":
        throw new Error(`Acción no disponible en Fase 1A: ${type}`);
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
