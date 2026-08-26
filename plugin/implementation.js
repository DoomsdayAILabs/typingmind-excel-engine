const CHANNEL = "tm-excel-engine";
const DEFAULT_MAX_ROWS = 200;
const HARD_MAX_ROWS = 500;
const TIMEOUT_MS = 180000;

function assertReadOnlySql(sql) {
  const text = String(sql || "").trim();
  if (!text) {
    throw new Error("Falta el SQL. Usa action=query y envía una consulta SELECT o WITH.");
  }
  const stripped = text.replace(/\/\*[\s\S]*?\*\//g, " ").replace(/--.*$/gm, " ").trim();
  if (!/^(with|select|explain)\b/i.test(stripped)) {
    throw new Error("Solo se permiten consultas de lectura (SELECT, WITH o EXPLAIN) contra excel_data.");
  }
  if (/\b(drop|alter|insert|update|delete|attach|copy|export|pragma\s+force|create\s+table|create\s+view)\b/i.test(stripped)) {
    throw new Error("Esta herramienta es de solo lectura. No se permiten cambios a la tabla excel_data.");
  }
  return text;
}

function callEngine(payload) {
  return new Promise((resolve, reject) => {
    const requestId = `tm-excel-${Date.now()}-${Math.random().toString(16).slice(2)}`;
    const timer = setTimeout(() => {
      window.removeEventListener("message", onMessage);
      reject(new Error("El motor Excel no respondió a tiempo. ¿Están instaladas las Extensions v0.4.20 y el puente, y recargaste TypingMind?"));
    }, TIMEOUT_MS);

    function onMessage(event) {
      const data = event.data;
      if (!data || data.channel !== CHANNEL || data.requestId !== requestId) return;
      window.removeEventListener("message", onMessage);
      clearTimeout(timer);
      if (data.error) {
        reject(new Error(data.error));
        return;
      }
      resolve(data.result);
    }

    window.addEventListener("message", onMessage);
    window.parent.postMessage({
      channel: CHANNEL,
      requestId,
      ...payload
    }, "*");
  });
}

async function query_excel_data(params) {
  const action = params?.action || "query";
  const maxRows = Math.min(HARD_MAX_ROWS, Math.max(1, Number(params?.max_rows) || DEFAULT_MAX_ROWS));

  if (action === "schema") {
    return callEngine({ action: "schema" });
  }

  if (action === "preview") {
    return callEngine({ action: "preview", maxRows: 10 });
  }

  if (action === "query") {
    const sql = assertReadOnlySql(params?.sql);
    return callEngine({ action: "query", sql, maxRows });
  }

  throw new Error("action debe ser schema, preview o query.");
}
