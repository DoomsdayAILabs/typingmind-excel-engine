/*
 * TypingMind Excel Bridge — Comunicación con DuckDB Worker
 *
 * Maneja la comunicación asíncrona entre el motor principal
 * y el worker de DuckDB (duckdb-worker.js).
 */

const WORKER_PATH = "./duckdb-worker.js";

let worker = null;
let messageQueue = {};
let nextMessageId = 1;

/* ============================================================
 * INICIALIZACIÓN DEL WORKER
 * ============================================================ */

function initializeWorker() {
  if (worker) {
    return;
  }

  worker = new Worker(WORKER_PATH);

  worker.onmessage = (event) => {
    const { id, type, success, result, error, stack } = event.data || {};

    if (!id || !messageQueue[id]) {
      return;
    }

    const { resolve, reject } = messageQueue[id];
    delete messageQueue[id];

    if (type === "response" && success) {
      resolve(result);
    } else if (type === "response" && !success) {
      const err = new Error(error || "Error en worker");
      err.stack = stack;
      reject(err);
    } else {
      reject(new Error(`Tipo de mensaje desconocido: ${type}`));
    }
  };

  worker.onerror = (error) => {
    console.error("Error en DuckDB Worker:", error);
    for (const id in messageQueue) {
      messageQueue[id].reject(
        new Error("Error en worker: " + (error.message || "Error desconocido"))
      );
    }
    messageQueue = {};
  };
}

/* ============================================================
 * ENVÍO DE MENSAJES AL WORKER
 * ============================================================ */

function sendToWorker(type, payload = {}) {
  return new Promise((resolve, reject) => {
    if (!worker) {
      initializeWorker();
    }

    const id = nextMessageId++;
    messageQueue[id] = { resolve, reject };

    try {
      worker.postMessage({
        id,
        type,
        ...payload
      });
    } catch (error) {
      delete messageQueue[id];
      reject(error);
    }
  });
}

/* ============================================================
 * API PÚBLICA
 * ============================================================ */

const DuckDBBridge = {
  init: () => sendToWorker("init"),
  executeQuery: (sql) => sendToWorker("executeQuery", { sql }),
  loadCSV: (fileName, buffer, tableName, options) =>
    sendToWorker("loadCSV", { fileName, buffer, tableName, options }),
  loadParquet: (fileName, buffer, tableName) =>
    sendToWorker("loadParquet", { fileName, buffer, tableName }),
  terminate: () => {
    if (worker) {
      worker.terminate();
      worker = null;
      messageQueue = {};
    }
  }
};

// Exponer la API al motor principal
window.TMDuckDBBridge = DuckDBBridge;
