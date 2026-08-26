/*
 * Puente TypingMind Plugin ↔ Excel Engine v0.4.20
 *
 * Instalar como Extension APARTE del motor.
 * No modifica typingmind-excel-engine-v0.4.20-test.js.
 *
 * El plugin corre en un iframe. Este script vive en la página
 * de TypingMind, escucha postMessage y llama a window.TMExcelEngine.
 */
(() => {
  "use strict";

  const CHANNEL = "tm-excel-engine";

  function reply(source, requestId, payload) {
    if (!source || typeof source.postMessage !== "function") return;
    source.postMessage({ channel: CHANNEL, requestId, ...payload }, "*");
  }

  function clipRows(rows, maxRows) {
    const list = Array.isArray(rows) ? rows : [];
    const limit = Math.min(500, Math.max(1, Number(maxRows) || 200));
    const clipped = list.length > limit;
    return {
      procesamiento: "LOCAL",
      filas_resultado: list.length,
      filas_enviadas_al_llm: clipped ? limit : list.length,
      recortado: clipped,
      resultado: clipped ? list.slice(0, limit) : list
    };
  }

  window.addEventListener("message", async (event) => {
    const data = event.data;
    if (!data || data.channel !== CHANNEL || !data.requestId) return;

    const engine = window.TMExcelEngine;
    if (!engine) {
      reply(event.source, data.requestId, {
        error: "El motor Excel no está inicializado. Instala typingmind-excel-engine-v0.4.20-test.js como Extension y recarga TypingMind."
      });
      return;
    }

    try {
      if (data.action === "schema") {
        const schema = await engine.getSchema();
        reply(event.source, data.requestId, {
          result: {
            procesamiento: "LOCAL",
            tabla: "excel_data",
            version_motor: engine.version,
            columnas: schema
          }
        });
        return;
      }

      if (data.action === "preview") {
        const rows = await engine.previewRows();
        reply(event.source, data.requestId, {
          result: clipRows(rows, data.maxRows || 10)
        });
        return;
      }

      if (data.action === "query") {
        const rows = await engine.executeSql(data.sql, "plugin");
        reply(event.source, data.requestId, {
          result: clipRows(rows, data.maxRows)
        });
        return;
      }

      reply(event.source, data.requestId, {
        error: "Acción no reconocida por el puente Excel."
      });
    } catch (error) {
      reply(event.source, data.requestId, {
        error: error?.message || String(error)
      });
    }
  });
})();
