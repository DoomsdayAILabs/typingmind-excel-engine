/**
 * TypingMind Excel Engine — v1.0.js
 * Script principal de inyección del Widget visual e integración con DuckDB-WASM Worker.
 */
(function () {
  "use strict";

  const WORKER_PATH = "./duckdb-worker.js";
  let worker = null;
  const pending = new Map();
  let nextRequestId = 1;
  let ultimosResultados = []; // Almacena la última consulta para inyectarla

  // --- 1. Inicialización del Web Worker ---
  function initWorker() {
    try {
      updateStatus("Inicializando Web Worker...", "pending");
      worker = new Worker(WORKER_PATH);

      worker.onmessage = (event) => {
        const { requestId, success, data, error } = event.data || {};
        const pendingItem = pending.get(requestId);
        if (!pendingItem) return;

        clearTimeout(pendingItem.timer);
        pending.delete(requestId);

        if (success) pendingItem.resolve(data);
        else pendingItem.reject(new Error(error));
      };

      worker.onerror = (event) => {
        updateStatus("Error en el Web Worker", "error");
        pending.clear();
      };

      sendRequest("initDuckDB", {}, 120000)
        .then(() => updateStatus("Motor DuckDB Listo", "success"))
        .catch((err) => updateStatus("Fallo al iniciar DuckDB: " + err.message, "error"));

    } catch (e) {
      updateStatus("Error de inicialización", "error");
    }
  }

  function sendRequest(type, payload = {}, timeoutMs = 60000, transferables = []) {
    const requestId = nextRequestId++;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        pending.delete(requestId);
        reject(new Error(`Timeout de ${timeoutMs}ms para (${type})`));
      }, timeoutMs);

      pending.set(requestId, { resolve, reject, timer, type });
      worker.postMessage({ type, requestId, ...payload }, transferables);
    });
  }

  // --- 2. Inyección de la UI del Widget ---
  function injectWidget() {
    if (document.getElementById("tmee-widget-root")) return;

    const styleEl = document.createElement("style");
    styleEl.id = "tmee-widget-styles";
    styleEl.innerHTML = `
      #tmee-widget-root { position: fixed; bottom: 20px; left: 20px; width: 400px; max-height: 85vh; background-color: #1e2026; border: 1px solid #374151; border-radius: 12px; box-shadow: 0 10px 25px -5px rgba(0,0,0,0.5); color: #f3f4f6; font-family: -apple-system, sans-serif; display: flex; flex-direction: column; overflow: hidden; z-index: 99999; }
      .tmee-header { padding: 12px 16px; background-color: #111318; border-bottom: 1px solid #374151; display: flex; justify-content: space-between; align-items: center; }
      .tmee-title { font-size: 14px; font-weight: 600; display: flex; align-items: center; gap: 6px; }
      .tmee-header-controls { display: flex; align-items: center; gap: 8px; }
      .tmee-btn-icon { background: transparent; border: none; color: #9ca3af; cursor: pointer; font-size: 14px; padding: 0 4px; transition: color 0.2s; }
      .tmee-btn-icon:hover { color: #f3f4f6; }
      .tmee-status-badge { font-size: 11px; padding: 2px 8px; border-radius: 12px; font-weight: 500; background-color: #374151; color: #9ca3af; }
      .tmee-status-badge.pending { background-color: rgba(245,158,11,0.2); color: #fbbf24; }
      .tmee-status-badge.success { background-color: rgba(16,185,129,0.2); color: #34d399; }
      .tmee-status-badge.error { background-color: rgba(239,68,68,0.2); color: #f87171; }
      .tmee-body { padding: 16px; overflow-y: auto; flex: 1; display: flex; flex-direction: column; gap: 12px; transition: all 0.3s ease; }
      .tmee-dropzone { border: 2px dashed #4b5563; border-radius: 8px; padding: 20px 16px; text-align: center; cursor: pointer; background-color: rgba(255,255,255,0.02); transition: all 0.2s; }
      .tmee-dropzone:hover, .tmee-dropzone.dragover { border-color: #10b981; background-color: rgba(16,185,129,0.05); }
      .tmee-dropzone-btn { background-color: #374151; color: #f3f4f6; border: none; padding: 6px 12px; border-radius: 6px; font-size: 11px; cursor: pointer; margin-top: 8px; }
      .tmee-result-container { display: none; flex-direction: column; gap: 12px; border-top: 1px solid #374151; padding-top: 12px; }
      .tmee-meta-info { font-size: 12px; background-color: rgba(255,255,255,0.03); border: 1px solid #374151; border-radius: 6px; padding: 8px 12px; display: flex; justify-content: space-between; }
      .tmee-sql-box { display: flex; flex-direction: column; gap: 8px; margin-top: 8px; }
      .tmee-textarea { background: #111318; color: #10b981; border: 1px solid #374151; border-radius: 6px; padding: 10px; font-family: monospace; font-size: 12px; resize: vertical; min-height: 70px; width: 100%; box-sizing: border-box;}
      .tmee-textarea:focus { outline: none; border-color: #10b981; }
      .tmee-btn-primary { background-color: #10b981; color: #111318; border: none; padding: 8px 16px; border-radius: 6px; font-size: 12px; font-weight: bold; cursor: pointer; transition: background 0.2s; width: 100%; }
      .tmee-btn-primary:hover { background-color: #059669; }
      .tmee-btn-inject { background-color: #3b82f6; color: #ffffff; width: auto; padding: 4px 10px; font-size: 11px; display: flex; align-items: center; gap: 4px;}
      .tmee-btn-inject:hover { background-color: #2563eb; }
      .tmee-results-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px; }
      .tmee-table-title { font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em; color: #9ca3af; }
      .tmee-table-wrapper { max-height: 200px; overflow: auto; border: 1px solid #374151; border-radius: 6px; background-color: #111318; }
      .tmee-table { width: 100%; border-collapse: collapse; font-size: 11px; text-align: left; }
      .tmee-table th { background-color: #1f2937; color: #f3f4f6; padding: 6px 8px; border-bottom: 1px solid #374151; position: sticky; top: 0; }
      .tmee-table td { padding: 6px 8px; border-bottom: 1px solid #2d3748; color: #d1d5db; white-space: nowrap; }
    `;
    document.head.appendChild(styleEl);

    const rootEl = document.createElement("div");
    rootEl.id = "tmee-widget-root";
    rootEl.innerHTML = `
      <div class="tmee-header">
        <div class="tmee-title"><span style="color:#10b981">📊</span> TM Excel Engine v1.0</div>
        <div class="tmee-header-controls">
          <span id="tmee-status" class="tmee-status-badge">Iniciando...</span>
          <button id="tmee-btn-toggle" class="tmee-btn-icon" title="Minimizar/Maximizar">➖</button>
        </div>
      </div>
      <div class="tmee-body" id="tmee-body">
        <div id="tmee-dropzone" class="tmee-dropzone">
          <div style="font-size:24px; margin-bottom:4px">📥</div>
          <div style="font-size:12px; color:#9ca3af">Arrastra tu archivo (.csv, .parquet, .xlsx)</div>
          <button id="tmee-select-btn" class="tmee-dropzone-btn">Seleccionar Archivo</button>
          <input type="file" id="tmee-file-input" accept=".csv,.parquet,.xlsx,.xls" style="display:none;" />
        </div>
        
        <div id="tmee-result" class="tmee-result-container">
          <div class="tmee-meta-info">
            <span>Tabla: <strong style="color:#10b981" id="tmee-meta-table">-</strong></span>
            <span>Registros: <strong style="color:#10b981" id="tmee-meta-rows">-</strong></span>
          </div>
          
          <!-- Consola SQL -->
          <div class="tmee-sql-box">
            <div class="tmee-table-title">Consola SQL</div>
            <textarea id="tmee-sql-input" class="tmee-textarea" spellcheck="false"></textarea>
            <button id="tmee-btn-run-sql" class="tmee-btn-primary">▶ Ejecutar Consulta</button>
          </div>

          <!-- Resultados SQL -->
          <div style="margin-top: 8px;">
            <div class="tmee-results-header">
              <div class="tmee-table-title" id="tmee-sql-results-title">Resultados</div>
              <button id="tmee-btn-inject" class="tmee-btn-primary tmee-btn-inject">💬 Enviar a TM</button>
            </div>
            <div class="tmee-table-wrapper">
              <table class="tmee-table">
                <thead id="tmee-sql-headers"></thead>
                <tbody id="tmee-sql-body"></tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    `;
    document.body.appendChild(rootEl);
    setupEvents();
  }

  // --- 3. Lógica de Eventos ---
  function setupEvents() {
    const dropzone = document.getElementById("tmee-dropzone");
    const fileInput = document.getElementById("tmee-file-input");
    const bodyEl = document.getElementById("tmee-body");
    const toggleBtn = document.getElementById("tmee-btn-toggle");
    
    // Controles de ventana
    toggleBtn.addEventListener("click", () => {
      if (bodyEl.style.display === "none") {
        bodyEl.style.display = "flex";
        toggleBtn.textContent = "➖";
      } else {
        bodyEl.style.display = "none";
        toggleBtn.textContent = "➕";
      }
    });

    document.getElementById("tmee-select-btn").addEventListener("click", () => fileInput.click());
    fileInput.addEventListener("change", (e) => e.target.files[0] && procesarArchivo(e.target.files[0]));
    
    dropzone.addEventListener("dragover", (e) => { e.preventDefault(); dropzone.classList.add("dragover"); });
    dropzone.addEventListener("dragleave", () => dropzone.classList.remove("dragover"));
    dropzone.addEventListener("drop", (e) => {
      e.preventDefault(); dropzone.classList.remove("dragover");
      if (e.dataTransfer.files[0]) procesarArchivo(e.dataTransfer.files[0]);
    });

    document.getElementById("tmee-btn-run-sql").addEventListener("click", ejecutarSQL);
    document.getElementById("tmee-btn-inject").addEventListener("click", inyectarEnChat);
  }

  function updateStatus(text, statusClass = "") {
    const statusEl = document.getElementById("tmee-status");
    if (!statusEl) return;
    statusEl.textContent = text;
    statusEl.className = "tmee-status-badge " + statusClass;
  }

  // --- 4. Carga de Archivos ---
  function procesarArchivo(file) {
    const extension = file.name.split('.').pop().toLowerCase();
    const tableName = "excel_data";
    document.getElementById("tmee-result").style.display = "none";
    updateStatus("Cargando...", "pending");
    
    const reader = new FileReader();
    
    if (extension === "csv") {
      reader.onload = async (e) => {
        try {
          const result = await sendRequest("loadCSV", { csvData: e.target.result, tableName }, 120000);
          renderizarResultado(result);
          updateStatus("CSV Cargado", "success");
        } catch (err) { updateStatus("Error CSV", "error"); }
      };
      reader.readAsText(file);
    } else if (["parquet", "xlsx", "xls"].includes(extension)) {
      reader.onload = async (e) => {
        try {
          const typeMap = { "parquet": "loadParquet", "xlsx": "loadExcel", "xls": "loadExcel" };
          const payload = extension === "parquet" ? { parquetData: e.target.result, tableName } : { excelBuffer: e.target.result, tableName };
          const result = await sendRequest(typeMap[extension], payload, 120000, [e.target.result]);
          renderizarResultado(result);
          updateStatus("Archivo Cargado", "success");
        } catch (err) { updateStatus("Error de archivo", "error"); }
      };
      reader.readAsArrayBuffer(file);
    }
  }

  function renderizarResultado(result) {
    document.getElementById("tmee-result").style.display = "flex";
    document.getElementById("tmee-meta-table").textContent = result.tabla;
    document.getElementById("tmee-meta-rows").textContent = Number(result.registros).toLocaleString();
    
    const sqlInput = document.getElementById("tmee-sql-input");
    sqlInput.value = `SELECT * FROM ${result.tabla} LIMIT 10;`;
    ejecutarSQL();
  }

  // --- 5. Motor SQL Dinámico ---
  async function ejecutarSQL() {
    const query = document.getElementById("tmee-sql-input").value;
    if (!query.trim()) return;

    try {
      updateStatus("Ejecutando SQL...", "pending");
      const rows = await sendRequest("executeQuery", { sql: query }, 60000);
      ultimosResultados = rows; // Guardar para la inyección
      renderizarTablaSQL(rows);
      updateStatus("Consulta OK", "success");
    } catch (err) {
      updateStatus("Error SQL", "error");
      ultimosResultados = [];
      alert("Error en la consulta SQL:\n" + err.message);
    }
  }

  function renderizarTablaSQL(rows) {
    const headersEl = document.getElementById("tmee-sql-headers");
    const bodyEl = document.getElementById("tmee-sql-body");
    const titleEl = document.getElementById("tmee-sql-results-title");
    
    headersEl.innerHTML = "";
    bodyEl.innerHTML = "";

    if (!rows || rows.length === 0) {
      titleEl.textContent = "Resultados (0 filas)";
      headersEl.innerHTML = "<tr><th style='text-align:center; color:#9ca3af'>Sin resultados</th></tr>";
      return;
    }

    const displayRows = rows.slice(0, 100);
    titleEl.textContent = `Resultados (${displayRows.length}${rows.length > 100 ? '+' : ''} filas)`;
    
    const columns = Object.keys(rows[0]);
    const trHead = document.createElement("tr");
    columns.forEach(col => {
      const th = document.createElement("th");
      th.textContent = col;
      trHead.appendChild(th);
    });
    headersEl.appendChild(trHead);

    displayRows.forEach(row => {
      const tr = document.createElement("tr");
      columns.forEach(col => {
        const td = document.createElement("td");
        const val = row[col];
        td.textContent = val !== null && val !== undefined ? String(val) : "";
        tr.appendChild(td);
      });
      bodyEl.appendChild(tr);
    });
  }

  // --- 6. Inyección a TypingMind ---
  function inyectarEnChat() {
    if (!ultimosResultados || ultimosResultados.length === 0) {
      alert("No hay datos para inyectar. Ejecuta una consulta primero.");
      return;
    }

    // Formatear a Markdown (máximo 50 filas)
    const limit = Math.min(ultimosResultados.length, 50);
    const rows = ultimosResultados.slice(0, limit);
    const columns = Object.keys(rows[0]);
    
    let md = `*Resultados extraídos de la tabla local (${limit} filas):*\n\n`;
    md += "| " + columns.join(" | ") + " |\n";
    md += "| " + columns.map(() => "---").join(" | ") + " |\n";
    
    rows.forEach(row => {
      md += "| " + columns.map(col => {
        let val = row[col];
        if (val === null || val === undefined) return "";
        return String(val).replace(/\|/g, "-").replace(/\n/g, " ");
      }).join(" | ") + " |\n";
    });

    // Buscar el input nativo de TypingMind
    const chatInput = document.querySelector('textarea[data-testid="chat-input"]') || 
                      document.querySelector('textarea[placeholder*="Type"]') || 
                      document.querySelector('[contenteditable="true"]');

    if (chatInput) {
      if (chatInput.tagName === 'TEXTAREA') {
        chatInput.value += (chatInput.value ? "\n\n" : "") + md;
      } else {
        chatInput.textContent += (chatInput.textContent ? "\n\n" : "") + md;
      }
      
      // Despachar evento para que React actualice el estado del botón de envío
      chatInput.dispatchEvent(new Event('input', { bubbles: true }));
      
      // Minimizar el widget automáticamente para ver el chat
      document.getElementById("tmee-body").style.display = "none";
      document.getElementById("tmee-btn-toggle").textContent = "➕";
    } else {
      // Fallback si el DOM de TM cambia: Copiar al portapapeles
      navigator.clipboard.writeText(md).then(() => {
        alert("Tabla copiada al portapapeles. Pégala en el chat.");
      }).catch(() => {
        alert("No se encontró el chat. Imprimiendo en consola.");
        console.log(md);
      });
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => { injectWidget(); initWorker(); });
  } else {
    injectWidget(); initWorker();
  }
})();