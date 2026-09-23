/**
 * TypingMind Plugin — query_excel_data (Excel Data Engine local · DuckDB-WASM)
 *
 * REQUISITO DE TYPINGMIND (docs oficiales, sección "JavaScript Code"):
 * el código de implementación debe declarar en el NIVEL SUPERIOR una función con
 * EXACTAMENTE el mismo nombre que el campo `name` del openaiSpec ("query_excel_data").
 * Por eso el cuerpo blindado va envuelto en `async function query_excel_data(params) { ... }`.
 * Nunca pegues sentencias `return` sueltas en el nivel superior: son un SyntaxError.
 *
 * Arquitectura: el plugin se ejecuta dentro de un iframe sandbox y no tiene acceso a
 * DuckDB. Envía la consulta al motor (Extensión, ventana principal) con
 * `window.parent.postMessage` y recibe las filas por el canal "tm-excel-engine".
 */
async function query_excel_data(params) {
  return new Promise((resolve) => {
    try {
      const sql = params?.sql_query || params?.sql || params?.query;
      if (!sql || typeof sql !== "string") {
        return resolve("Error: Consulta SQL no válida o vacía. Asegúrate de generar una consulta SELECT.");
      }

      const stripped = sql.replace(/\/\*[\s\S]*?\*\//g, " ").replace(/--.*$/gm, " ").trim();
      if (/\b(drop|alter|insert|update|delete|attach|copy|export|create|pragma\s+force)\b/i.test(stripped)) {
        return resolve("Error: Solo se permiten consultas de lectura (SELECT, WITH, EXPLAIN).");
      }

      const reqId = "req_" + Date.now() + "_" + Math.floor(Math.random() * 1000);
      const timer = setTimeout(() => {
        resolve("Error: Timeout. El motor local no respondió. Verifica que la extensión esté cargada y un archivo Excel esté procesado.");
      }, 30000);

      function onMessage(event) {
        if (event.data && event.data.channel === "tm-excel-engine" && event.data.requestId === reqId) {
          window.removeEventListener("message", onMessage);
          clearTimeout(timer);

          if (!event.data.success) {
            return resolve("Error SQL del motor local: " + event.data.error);
          }

          const rows = event.data.data;
          if (!Array.isArray(rows)) return resolve("Error interno: Datos no procesables.");
          if (rows.length === 0) return resolve("Consulta ejecutada con éxito (0 resultados).");

          const limit = Math.min(rows.length, 50);
          const displayRows = rows.slice(0, limit);
          const columns = Object.keys(displayRows[0]);

          let md = `Resultado (${limit} de ${rows.length} filas):\n\n`;
          md += "| " + columns.join(" | ") + " |\n";
          md += "| " + columns.map(() => "---").join(" | ") + " |\n";

          displayRows.forEach(row => {
            md += "| " + columns.map(col => {
              let val = row[col];
              return val !== null && val !== undefined ? String(val).replace(/\|/g, "\\|").replace(/\n/g, " ") : "";
            }).join(" | ") + " |\n";
          });

          resolve(md);
        }
      }

      window.addEventListener("message", onMessage);
      window.parent.postMessage({ channel: "tm-excel-engine", action: "execute_sql", sql: sql, requestId: reqId }, "*");
    } catch (e) {
      resolve("Error crítico en la ejecución del plugin: " + (e.message || String(e)));
    }
  });
}
