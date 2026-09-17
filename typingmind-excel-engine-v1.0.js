/**
 * TypingMind Excel Engine — v1.0.js
 * Script principal de inyección del Widget visual e integración con DuckDB-WASM Worker.
 */
(function () {
  "use strict";

  // Configuración del Worker
  const WORKER_PATH = "./duckdb-worker.js";
  let worker = null;
  const pending = new Map();
  let nextRequestId = 1;

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

        const elapsed = Math.round(performance.now() - pendingItem.startedAt);

        if (success) {
          console.log(`[TMEE] Respuesta #${requestId} OK en ${elapsed}ms`);
          pendingItem.resolve(data);
        } else {
          console.error(`[TMEE] Error #${requestId} en ${elapsed}ms:`, error);
          pendingItem.reject(new Error(error));
        }
      };

      worker.onerror = (event) => {
        console.error("[TMEE] Error crítico del Web Worker:", event);
        updateStatus("Error en el Web Worker", "error");
        for (const [requestId, item] of pending.entries()) {
          clearTimeout(item.timer);
          item.reject(new Error(`Worker error para requestId ${requestId}`));
        }
        pending.clear();
      };

      // Inicializamos DuckDB en el worker de inmediato
      sendRequest("initDuckDB", {}, 120000)
        .then(() => {
          updateStatus("Motor DuckDB Listo", "success");
        })
        .catch((err) => {
          console.error("[TMEE] Fallo al inicializar DuckDB:", err);
          updateStatus("Fallo al iniciar DuckDB: " + err.message, "error");
        });

    } catch (e) {
      console.error("[TMEE] No se pudo instanciar el Web Worker:", e);
      updateStatus("Error de inicialización", "error");
    }
  }

  // Enviar peticiones asíncronas usando Promesas con soporte para Transferables
  function sendRequest(type, payload = {}, timeoutMs = 60000, transferables = []) {
    const requestId = nextRequestId++;
    const startedAt = performance.now();

    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        pending.delete(requestId);
        reject(new Error(`Timeout de ${timeoutMs}ms para requestId ${requestId} (${type})`));
      }, timeoutMs);

      pending.set(requestId, { resolve, reject, timer, startedAt, type });

      worker.postMessage(
        {
          type,
          requestId,
          ...payload
        },
        transferables
      );
    });
  }

  // --- 2. Inyección de la UI del Widget ---
  function injectWidget() {
    // Evitar doble inyección
    if (document.getElementById("tmee-widget-root")) return;

    // Inyectar Estilos CSS
    const styleEl = document.createElement("style");
    styleEl.id = "tmee-widget-styles";
    styleEl.innerHTML = `
      #tmee-widget-root {
        position: fixed;
        bottom: 20px;
        left: 20px;
        width: 380px;
        max-height: 80vh;
        background-color: #1e2026;
        border: 1px solid #374151;
        border-radius: 12px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.5), 0 8px 10px -6px rgba(0, 0, 0, 0.5);
        color: #f3f4f6;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
        display: flex;
        flex-direction: column;
        overflow: hidden;
        z-index: 99999;
      }
      .tmee-header {
        padding: 12px 16px;
        background-color: #111318;
        border-bottom: 1px solid #374151;
        display: flex;
        justify-content: space-between;
        align-items: center;
      }
      .tmee-title {
        font-size: 14px;
        font-weight: 600;
        letter-spacing: 0.025em;
        display: flex;
        align-items: center;
        gap: 6px;
      }
      .tmee-title-icon {
        color: #10b981;
        font-weight: bold;
      }
      .tmee-status-badge {
        font-size: 11px;
        padding: 2px 8px;
        border-radius: 12px;
        font-weight: 500;
        background-color: #374151;
        color: #9ca3af;
        transition: all 0.3s ease;
      }
      .tmee-status-badge.pending {
        background-color: rgba(245, 158, 11, 0.2);
        color: #fbbf24;
      }
      .tmee-status-badge.success {
        background-color: rgba(16, 185, 129, 0.2);
        color: #34d399;
      }
      .tmee-status-badge.error {
        background-color: rgba(239, 68, 68, 0.2);
        color: #f87171;
      }
      .tmee-body {
        padding: 16px;
        overflow-y: auto;
        flex: 1;
        display: flex;
        flex-direction: column;
        gap: 12px;
      }
      .tmee-dropzone {
        border: 2px dashed #4b5563;
        border-radius: 8px;
        padding: 24px 16px;
        text-align: center;
        cursor: pointer;
        background-color: rgba(255, 255, 255, 0.02);
        transition: all 0.2s ease;
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 8px;
      }
      .tmee-dropzone:hover, .tmee-dropzone.dragover {
        border-color: #10b981;
        background-color: rgba(16, 185, 129, 0.05);
      }
      .tmee-dropzone-icon {
        font-size: 24px;
        color: #9ca3af;
      }
      .tmee-dropzone:hover .tmee-dropzone-icon, .tmee-dropzone.dragover .tmee-dropzone-icon {
        color: #10b981;
      }
      .tmee-dropzone-text {
        font-size: 12px;
        color: #9ca3af;
      }
      .tmee-dropzone-btn {
        background-color: #374151;
        color: #f3f4f6;
        border: none;
        padding: 6px 12px;
        border-radius: 6px;
        font-size: 11px;
        cursor: pointer;
        font-weight: 500;
        transition: background 0.2s;
      }
      .tmee-dropzone-btn:hover {
        background-color: #4b5563;
      }
      .tmee-result-container {
        display: none; /* Se muestra al cargar un archivo */
        flex-direction: column;
        gap: 12px;
        border-top: 1px solid #374151;
        padding-top: 12px;
      }
      .tmee-meta-info {
        font-size: 12px;
        background-color: rgba(255, 255, 255, 0.03);
        border: 1px solid #374151;
        border-radius: 6px;
        padding: 8px 12px;
        display: flex;
        flex-direction: column;
        gap: 4px;
      }
      .tmee-meta-row {
        display: flex;
        justify-content: space-between;
      }
      .tmee-meta-label {
        color: #9ca3af;
      }
      .tmee-meta-value {
        font-weight: 500;
        color: #10b981;
      }
      .tmee-table-title {
        font-size: 11px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #9ca3af;
        margin-bottom: 4px;
      }
      .tmee-table-wrapper {
        max-height: 150px;
        overflow-x: auto;
        overflow-y: auto;
        border: 1px solid #374151;
        border-radius: 6px;
        background-color: #111318;
      }
      .tmee-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 11px;
        text-align: left;
      }
      .tmee-table th {
        background-color: #1f2937;
        color: #f3f4f6;
        padding: 6px 8px;
        border-bottom: 1px solid #374151;
        font-weight: 600;
        position: sticky;
        top: 0;
      }
      .tmee-table td {
        padding: 6px 8px;
        border-bottom: 1px solid #2d3748;
        color: #d1d5db;
        white-space: nowrap;
      }
      .tmee-table tr:last-child td {
        border-bottom: none;
      }
      .tmee-table tr:hover td {
        background-color: rgba(255, 255, 255, 0.02);
      }
    `;
    document.head.appendChild(styleEl);

    // Crear Estructura HTML del Widget
    const rootEl = document.createElement("div");
    rootEl.id = "tmee-widget-root";
    rootEl.innerHTML = `
      <div class="tmee-header">
        <div class="tmee-title">
          <span class="tmee-title-icon">📊</span> TM Excel Engine v1.0
        </div>
        <span id="tmee-status" class="tmee-status-badge">Inicializando...</span>
      </div>
      <div class="tmee-body">
        <div id="tmee-dropzone" class="tmee-dropzone">
          <div class="tmee-dropzone-icon">📥</div>
          <div class="tmee-dropzone-text">Arrastra y suelta tu archivo aquí (.csv, .parquet, .xlsx)</div>
          <span style="font-size: 10px; color: #6b7280;">o también</span>
          <button id="tmee-select-btn" class="tmee-dropzone-btn">Seleccionar Archivo</button>
          <input type="file" id="tmee-file-input" accept=".csv,.parquet,.xlsx,.xls" style="display: none;" />
        </div>
        
        <!-- Contenedor del resultado de carga -->
        <div id="tmee-result" class="tmee-result-container">
          <div class="tmee-meta-info">
            <div class="tmee-meta-row">
              <span class="tmee-meta-label">Tabla SQL:</span>
              <span id="tmee-meta-table" class="tmee-meta-value">-</span>
            </div>
            <div class="tmee-meta-row">
              <span class="tmee-meta-label">Registros:</span>
              <span id="tmee-meta-rows" class="tmee-meta-value">-</span>
            </div>
            <div class="tmee-meta-row">
              <span class="tmee-meta-label">Columnas:</span>
              <span id="tmee-meta-cols" class="tmee-meta-value">-</span>
            </div>
          </div>

          <!-- Tabla Esquema -->
          <div>
            <div class="tmee-table-title">Esquema de Columnas</div>
            <div class="tmee-table-wrapper">
              <table id="tmee-schema-table" class="tmee-table">
                <thead>
                  <tr>
                    <th>Columna</th>
                    <th>Tipo de Dato</th>
                  </tr>
                </thead>
                <tbody></tbody>
              </table>
            </div>
          </div>

          <!-- Tabla Vista Previa -->
          <div>
            <div class="tmee-table-title">Previsualización de Datos (Primeras 5 filas)</div>
            <div class="tmee-table-wrapper">
              <table id="tmee-preview-table" class="tmee-table">
                <thead>
                  <tr id="tmee-preview-headers"></tr>
                </thead>
                <tbody id="tmee-preview-body"></tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    `;
    document.body.appendChild(rootEl);

    // Registrar Eventos de la UI
    setupEvents();
  }

  // --- 3. Lógica de Eventos y Arrastre de Archivos ---
  function setupEvents() {
    const dropzone = document.getElementById("tmee-dropzone");
    const fileInput = document.getElementById("tmee-file-input");
    const selectBtn = document.getElementById("tmee-select-btn");

    // Botón para seleccionar archivo de forma tradicional
    selectBtn.addEventListener("click", () => fileInput.click());
    fileInput.addEventListener("change", (e) => {
      if (e.target.files && e.target.files[0]) {
        procesarArchivo(e.target.files[0]);
      }
    });

    // Eventos Drag & Drop de Archivos
    dropzone.addEventListener("dragover", (e) => {
      e.preventDefault();
      dropzone.classList.add("dragover");
    });

    dropzone.addEventListener("dragleave", () => {
      dropzone.classList.remove("dragover");
    });

    dropzone.addEventListener("drop", (e) => {
      e.preventDefault();
      dropzone.classList.remove("dragover");
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        procesarArchivo(e.dataTransfer.files[0]);
      }
    });
  }

  // Actualizar el texto y el color del badge de estado
  function updateStatus(text, statusClass = "") {
    const statusEl = document.getElementById("tmee-status");
    if (!statusEl) return;
    statusEl.textContent = text;
    statusEl.className = "tmee-status-badge " + statusClass;
  }

  // --- 4. Lógica de Lectura e Invocación al Worker ---
  function procesarArchivo(file) {
    const extension = file.name.split('.').pop().toLowerCase();
    const tableName = "excel_data"; // Tabla principal del motor por convención

    updateStatus("Leyendo archivo...", "pending");
    ocultarResultado();

    const reader = new FileReader();

    if (extension === "csv") {
      reader.onload = async (e) => {
        try {
          const csvText = e.target.result;
          updateStatus("Cargando CSV...", "pending");
          const result = await sendRequest("loadCSV", { csvData: csvText, tableName }, 120000);
          renderizarResultado(result);
          updateStatus("CSV Cargado", "success");
        } catch (err) {
          updateStatus("Error CSV: " + err.message, "error");
        }
      };
      reader.onerror = () => updateStatus("Error al leer CSV", "error");
      reader.readAsText(file);

    } else if (extension === "parquet") {
      reader.onload = async (e) => {
        try {
          const arrayBuffer = e.target.result;
          updateStatus("Cargando Parquet...", "pending");
          // Transferimos el ArrayBuffer para optimizar memoria
          const result = await sendRequest(
            "loadParquet",
            { parquetData: arrayBuffer, tableName },
            120000,
            [arrayBuffer]
          );
          renderizarResultado(result);
          updateStatus("Parquet Cargado", "success");
        } catch (err) {
          updateStatus("Error Parquet: " + err.message, "error");
        }
      };
      reader.onerror = () => updateStatus("Error al leer Parquet", "error");
      reader.readAsArrayBuffer(file);

    } else if (extension === "xlsx" || extension === "xls") {
      reader.onload = async (e) => {
        try {
          const arrayBuffer = e.target.result;
          updateStatus("Cargando Excel...", "pending");
          // Transferimos el ArrayBuffer para optimizar memoria
          const result = await sendRequest(
            "loadExcel",
            { excelBuffer: arrayBuffer, tableName },
            120000,
            [arrayBuffer]
          );
          renderizarResultado(result);
          updateStatus("Excel Cargado", "success");
        } catch (err) {
          updateStatus("Error Excel: " + err.message, "error");
        }
      };
      reader.onerror = () => updateStatus("Error al leer Excel", "error");
      reader.readAsArrayBuffer(file);

    } else {
      updateStatus("Formato no soportado", "error");
      alert("Formato de archivo no soportado. Debe ser .csv, .parquet, .xlsx o .xls");
    }
  }

  // Ocultar paneles de previsualización al cargar un archivo nuevo
  function ocultarResultado() {
    document.getElementById("tmee-result").style.display = "none";
  }

  // --- 5. Renderizado de Esquema y Preview en la UI ---
  function renderizarResultado(result) {
    const resultContainer = document.getElementById("tmee-result");
    resultContainer.style.display = "flex";

    // Meta Info
    document.getElementById("tmee-meta-table").textContent = result.tabla;
    document.getElementById("tmee-meta-rows").textContent = Number(result.registros).toLocaleString();
    document.getElementById("tmee-meta-cols").textContent = result.columnas.length;

    // Tabla de Esquema (Nombre columna y tipo de dato)
    const schemaBody = document.querySelector("#tmee-schema-table tbody");
    schemaBody.innerHTML = "";
    if (result.esquema && Array.isArray(result.esquema)) {
      result.esquema.forEach((col) => {
        const row = document.createElement("tr");
        row.innerHTML = `
          <td><strong>${escapeHtml(col.column_name)}</strong></td>
          <td><code style="background:#1f2937;padding:2px 4px;border-radius:4px;color:#f43f5e;">${escapeHtml(col.data_type)}</code></td>
        `;
        schemaBody.appendChild(row);
      });
    }

    // Tabla de Vista Previa (Máximo primeros 5 registros para no saturar el widget)
    const previewHeaders = document.getElementById("tmee-preview-headers");
    const previewBody = document.getElementById("tmee-preview-body");
    previewHeaders.innerHTML = "";
    previewBody.innerHTML = "";

    if (result.columnas && Array.isArray(result.columnas) && result.preview && Array.isArray(result.preview)) {
      // 1. Cabeceras
      result.columnas.forEach((col) => {
        const th = document.createElement("th");
        th.textContent = col;
        previewHeaders.appendChild(th);
      });

      // 2. Filas (Tomamos máximo 5 para la visualización del widget compacto)
      const rowsToShow = result.preview.slice(0, 5);
      rowsToShow.forEach((rowObj) => {
        const tr = document.createElement("tr");
        result.columnas.forEach((col) => {
          const td = document.createElement("td");
          const val = rowObj[col];
          td.textContent = val !== null && val !== undefined ? String(val) : "";
          tr.appendChild(td);
        });
        previewBody.appendChild(tr);
      });
    }
  }

  // Función de escape básica para seguridad HTML
  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  // --- 6. Inicialización Automática al cargar el Script ---
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => {
      injectWidget();
      initWorker();
    });
  } else {
    injectWidget();
    initWorker();
  }

})();
