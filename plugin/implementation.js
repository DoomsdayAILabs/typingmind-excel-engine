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
 *
 * GUARDRAIL DE SOLO LECTURA (Fase 8 · comprobaciones en cadena sobre el SQL recibido):
 *   1. Saneado en una sola pasada: fuera los comentarios (de bloque y de línea) y vaciados
 *      los literales ('...' y "...") conservando sus delimitadores. Así lo entrecomillado deja
 *      de producir falsos positivos y lo que queda fuera ya no puede ocultar un verbo destructivo.
 *   2. Lista blanca: el texto saneado debe empezar por SELECT, WITH, EXPLAIN, SHOW o DESCRIBE.
 *   3. Lista negra ampliada sobre el texto saneado: escritura, DDL, DCL y administración
 *      del motor (drop, alter, insert, update, delete, truncate, create, merge, attach, detach,
 *      copy, export, import, install, load, call, set, reset, vacuum, checkpoint, pragma,
 *      grant, revoke, use, begin, commit, rollback, transaction, prepare, execute, deallocate).
 *   4. Rechazo de sentencias apiladas (punto y coma fuera de literales) y de literales sin cerrar.
 * Decisiones asumidas: PRAGMA queda bloqueado (usar DESCRIBE / information_schema) y la
 * sentencia debe empezar por palabra clave formal (no se admite "(SELECT ...)").
 */
async function query_excel_data(params) {
  return new Promise((resolve) => {
    try {
      const sql = params?.sql_query || params?.sql || params?.query;
      if (!sql || typeof sql !== "string") {
        return resolve("Error: Consulta SQL no válida o vacía. Asegúrate de generar una consulta SELECT.");
      }

      // --- Guardrail de solo lectura: lista blanca + lista negra saneada ---
      // 1) En una sola pasada se eliminan los comentarios (de bloque y de línea) y se vacían
      //    los literales ('...' y "..."), conservando los delimitadores para no alterar los
      //    límites de palabra. Señala también los literales sin cerrar (escaneo no fiable).
      const saneaSQL = (texto) => {
        let out = "";
        let i = 0;
        let corte = 0;
        let literalAbierto = false;
        const n = texto.length;
        while (i < n) {
          const ch = texto[i];
          const next = texto[i + 1];
          if (ch === "'" || ch === '"') {
            const q = ch;
            out += texto.slice(corte, i) + q + q;
            i++;
            let cerrado = false;
            while (i < n) {
              if (texto[i] === q) {
                if (texto[i + 1] === q) { i += 2; continue; }
                i++;
                cerrado = true;
                break;
              }
              i++;
            }
            if (!cerrado) literalAbierto = true;
            corte = i;
            continue;
          }
          if (ch === "/" && next === "*") {
            out += texto.slice(corte, i) + " ";
            i += 2;
            while (i < n && !(texto[i] === "*" && texto[i + 1] === "/")) i++;
            i += 2;
            corte = i;
            continue;
          }
          if (ch === "-" && next === "-") {
            out += texto.slice(corte, i) + " ";
            i += 2;
            while (i < n && texto[i] !== "\n" && texto[i] !== "\r") i++;
            corte = i;
            continue;
          }
          i++;
        }
        return { sql: out + texto.slice(corte), literalAbierto: literalAbierto };
      };

      // 2) Lista blanca: la sentencia debe empezar por una de las cinco formas de lectura.
      const LIMITADO = /^(select|with|explain|show|describe)\b/i;
      // 3) Lista negra ampliada: escritura, DDL, DCL y administración del motor.
      const PROHIBIDO = /\b(drop|alter|insert|update|delete|truncate|create|merge|attach|detach|copy|export|import|install|load|call|set|reset|vacuum|checkpoint|pragma|grant|revoke|use|begin|commit|rollback|transaction|prepare|execute|deallocate)\b/i;

      const saneado = saneaSQL(sql);
      const cuerpo = saneado.sql.replace(/;\s*$/, "").trim();
      // 4) Rechazo si hay un literal sin cerrar, sentencias apiladas, un verbo prohibido o
      //    una forma inicial distinta de SELECT/WITH/EXPLAIN/SHOW/DESCRIBE.
      const rechazada =
        saneado.literalAbierto ||
        cuerpo.indexOf(";") !== -1 ||
        !LIMITADO.test(cuerpo) ||
        PROHIBIDO.test(cuerpo);

      if (rechazada) {
        return resolve("Error: Solo se permiten consultas de lectura (SELECT, WITH, EXPLAIN, SHOW, DESCRIBE).");
      }

      // --- Puente postMessage con el motor local (canal "tm-excel-engine") ---
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
