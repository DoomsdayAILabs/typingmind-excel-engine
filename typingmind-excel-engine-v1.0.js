/**
 * TypingMind Excel Engine — v1.0.js (Fase 6: Selector de hojas para libros multi-hoja)
 * Script principal de inyección del Widget visual e integración con DuckDB-WASM Worker.
 *
 * Cambios de esta fase:
 *  1) Los libros Excel multi-hoja exponen una barra compacta (.tmee-sheets-bar) entre los
 *     metadatos y la consola SQL: un desplegable nativo (#tmee-sheets-select) con las
 *     pestañas detectadas y un resumen "Hojas N/M · K vacías" (#tmee-sheets-summary).
 *  2) Alternar de hoja NO recarga el archivo: solo cambia la tabla (excel_data_<hoja>) de
 *     la consola SQL y se vuelve a consultar por el mismo Worker.
 *  3) La sustitución de la consulta es conservadora: si el usuario editó la consola SQL
 *     solo se reemplaza el token FROM <tabla>, sin tocar el resto de su consulta.
 *
 * Se conserva íntegro el comportamiento de la Fase 5: icono flotante de 50x50 px,
 * arrastre por cabecera (ratón + táctil) y minimizado por clic/tap.
 */
(function () {
  "use strict";

  const WORKER_PATH = "https://doomsdayailabs.github.io/typingmind-excel-engine/duckdb-worker.js";
  const MAX_FILAS_INYECCION = 50;

  // --- Fase 5: constantes de interacción ---
  const UMBRAL_ARRASTRE_PX = 5;                   // desplazamiento mínimo para considerarlo arrastre y no clic
  const VENTANA_SUPRESION_MOUSE_MS = 400;         // descarta el mousedown sintético que el navegador emite tras un toque
  const OPCIONES_TOUCH_MOVE = { passive: false }; // permite preventDefault para evitar el scroll durante el arrastre táctil

  let worker = null;
  let workerObjectUrl = null;
  const pending = new Map();
  let nextRequestId = 1;
  let ultimosResultados = [];

  // --- Fase 6: estado del libro cargado y de la hoja visible ---
  let libroActual = null;      // result de la última carga EXCEL multi-hoja (fuente de `hojas`)
  let hojaActiva = null;       // nombre de tabla de la hoja mostrada en este momento
  let plantillaSqlActual = ""; // última plantilla autogenerada (permite detectar si el usuario editó la consola)

  // --- Fase 5: estado del widget y del arrastre ---
  let rootEl = null;      // nodo #tmee-widget-root
  let headerEl = null;    // nodo .tmee-header (asa de arrastre y área de clic)
  let dragActivo = false; // hay un gesto en curso (mousedown/touchstart sin soltar)
  let isDragging = false; // true solo si el puntero se movió más de UMBRAL_ARRASTRE_PX
  let dragOffsetX = 0;    // distancia entre el puntero y el borde izquierdo del widget
  let dragOffsetY = 0;
  let dragStartX = 0;     // punto donde inició el gesto
  let dragStartY = 0;
  let ultimoTouchTs = 0;  // marca temporal del último touchstart

  // --- 1. Inicialización del Web Worker con Proxy Blob (CORS Bypass) ---
  async function initWorker() {
    try {
      updateStatus("Inicializando Worker...", "pending");

      const response = await fetch(WORKER_PATH);
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const workerCode = await response.text();

      const blob = new Blob([workerCode], { type: 'application/javascript' });
      workerObjectUrl = URL.createObjectURL(blob);
      worker = new Worker(workerObjectUrl);

      worker.onmessage = (event) => {
        const { requestId, success, data, error } = event.data || {};
        const pendingItem = pending.get(requestId);
        if (!pendingItem) return;

        clearTimeout(pendingItem.timer);
        pending.delete(requestId);

        if (success) pendingItem.resolve(data);
        else pendingItem.reject(new Error(error));
      };

      worker.onerror = () => {
        updateStatus("Error en Worker", "error");
        pending.clear();
      };

      sendRequest("initDuckDB", {}, 120000)
        .then(() => {
          updateStatus("Motor Listo", "success");
          if (workerObjectUrl) {
            URL.revokeObjectURL(workerObjectUrl);
            workerObjectUrl = null;
          }
        })
        .catch((err) => updateStatus("Fallo al iniciar: " + err.message, "error"));

    } catch (e) {
      updateStatus("Error de inicialización: " + e.message, "error");
    }
  }

  function sendRequest(type, payload = {}, timeoutMs = 60000, transferables = []) {
    const requestId = nextRequestId++;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        pending.delete(requestId);
        reject(new Error(`Timeout de ${timeoutMs}ms`));
      }, timeoutMs);

      pending.set(requestId, { resolve, reject, timer });
      worker.postMessage({ type, requestId, ...payload }, transferables);
    });
  }

  // --- 1.1 Puente de Comunicación para Plugins (TypingMind AI) ---
  window.addEventListener("message", async (event) => {
    if (event.data && event.data.channel === "tm-excel-engine" && event.data.action === "execute_sql") {
      try {
        const sql = event.data.sql;
        if (!sql) throw new Error("No SQL query provided");

        const rows = await sendRequest("executeQuery", { sql }, 60000);
        ultimosResultados = rows;
        renderizarTablaSQL(rows);
        updateStatus("Consulta IA OK", "success");

        if (event.source && typeof event.source.postMessage === "function") {
          event.source.postMessage({ channel: "tm-excel-engine", requestId: event.data.requestId, success: true, data: rows }, "*");
        } else {
          document.querySelectorAll('iframe').forEach(iframe => {
            iframe.contentWindow.postMessage({ channel: "tm-excel-engine", requestId: event.data.requestId, success: true, data: rows }, "*");
          });
        }
      } catch (err) {
        updateStatus("Error SQL IA", "error");
        ultimosResultados = [];
        if (event.source && typeof event.source.postMessage === "function") {
          event.source.postMessage({ channel: "tm-excel-engine", requestId: event.data.requestId, success: false, error: err.message }, "*");
        }
      }
    }
  });

  // --- 2. Inyección de la UI del Widget ---
  function injectWidget() {
    const existente = document.getElementById("tmee-widget-root");
    if (existente) { rootEl = existente; headerEl = rootEl.querySelector(".tmee-header"); return; }

    const styleEl = document.createElement("style");
    styleEl.id = "tmee-widget-styles";
    styleEl.innerHTML = `
      #tmee-widget-root { position: fixed; bottom: 20px; left: 20px; width: 400px; max-width: calc(100vw - 40px); max-height: 85vh; background-color: #1e2026; border: 1px solid #374151; border-radius: 12px; box-shadow: 0 10px 25px -5px rgba(0,0,0,0.5); color: #f3f4f6; font-family: -apple-system, sans-serif; display: flex; flex-direction: column; overflow: hidden; z-index: 99999; transition: width 0.3s, height 0.3s, border-radius 0.3s; }
      .tmee-header { padding: 12px 16px; background-color: #111318; border-bottom: 1px solid #374151; display: flex; justify-content: space-between; align-items: center; cursor: grab; -webkit-user-select: none; user-select: none; touch-action: none; }
      #tmee-widget-root.tmee-dragging, #tmee-widget-root.tmee-dragging .tmee-header { cursor: grabbing; }
      /* Fase 5: modo icono flotante (widget minimizado = botón circular de 50px) */
      #tmee-widget-root.tmee-minimized { width: 50px !important; height: 50px !important; border-radius: 25px !important; cursor: pointer; padding: 0; justify-content: center; align-items: center; background-color: #1e2026; transition: width 0.3s, height 0.3s, border-radius 0.3s; }
      #tmee-widget-root.tmee-minimized .tmee-body,
      #tmee-widget-root.tmee-minimized .tmee-header-controls,
      #tmee-widget-root.tmee-minimized .tmee-status-badge,
      #tmee-widget-root.tmee-minimized .tmee-title-text { display: none !important; }
      #tmee-widget-root.tmee-minimized .tmee-header { padding: 0; width: 100%; height: 100%; justify-content: center; border-bottom: none; background-color: transparent; }
      #tmee-widget-root.tmee-minimized .tmee-title { gap: 0; justify-content: center; }
      #tmee-widget-root.tmee-minimized .tmee-logo { font-size: 24px !important; margin: 0; display: block; text-align: center; width: 100%; }
      .tmee-title { font-size: 14px; font-weight: 600; display: flex; align-items: center; gap: 6px; }
      .tmee-logo { color: #10b981; font-size: 16px; line-height: 1; }
      .tmee-header-controls { display: flex; align-items: center; gap: 8px; }
      .tmee-status-badge { font-size: 11px; padding: 2px 8px; border-radius: 12px; font-weight: 500; background-color: #374151; color: #9ca3af; }
      .tmee-status-badge.pending { background-color: rgba(245,158,11,0.2); color: #fbbf24; }
      .tmee-status-badge.success { background-color: rgba(16,185,129,0.2); color: #34d399; }
      .tmee-status-badge.error { background-color: rgba(239,68,68,0.2); color: #f87171; }
      .tmee-body { padding: 16px; overflow-y: auto; flex: 1; display: flex; flex-direction: column; gap: 12px; transition: all 0.3s ease; }
      .tmee-dropzone { border: 2px dashed #4b5563; border-radius: 8px; padding: 20px 16px; text-align: center; cursor: pointer; background-color: rgba(255,255,255,0.02); transition: all 0.2s; }
      .tmee-dropzone:hover, .tmee-dropzone.dragover { border-color: #10b981; background-color: rgba(16,185,129,0.05); }
      .tmee-dropzone-btn { background-color: #374151; color: #f3f4f6; border: none; padding: 6px 12px; border-radius: 6px; font-size: 11px; cursor: pointer; margin-top: 8px; }

      .tmee-result-container { display: none; flex-direction: column; gap: 12px; border-top: 1px solid #374151; padding-top: 12px; }
      .tmee-meta-info { font-size: 12px; background-color: rgba(255,255,255,0.03); border: 1px solid #374151; border-radius: 6px; padding: 8px 12px; display: flex; flex-wrap: wrap; justify-content: space-between; gap: 4px 8px; }
      /* Fase 6: barra de hojas de un libro multi-hoja (oculta salvo que haya más de una hoja con datos) */
      .tmee-sheets-bar { display: flex; align-items: center; gap: 8px; }
      .tmee-sheets-bar[hidden] { display: none; }
      .tmee-select { flex: 1; min-width: 0; background-color: #111318; color: #f3f4f6; border: 1px solid #374151; border-radius: 6px; padding: 6px 8px; font-family: inherit; font-size: 11px; cursor: pointer; }
      .tmee-select:focus { outline: none; border-color: #10b981; }
      .tmee-select option:disabled { color: #6b7280; }
      .tmee-sheets-summary { font-size: 10px; color: #9ca3af; white-space: nowrap; }
      .tmee-sheets-summary.tmee-warn { color: #fbbf24; }
      .tmee-sheets-summary.tmee-error { color: #f87171; }
      .tmee-sql-box { display: flex; flex-direction: column; gap: 8px; margin-top: 8px; }
      .tmee-textarea { background: #111318; color: #10b981; border: 1px solid #374151; border-radius: 6px; padding: 10px; font-family: monospace; font-size: 12px; resize: vertical; min-height: 70px; width: 100%; box-sizing: border-box;}
      .tmee-textarea:focus { outline: none; border-color: #10b981; }
      .tmee-btn-primary { background-color: #10b981; color: #111318; border: none; padding: 8px 16px; border-radius: 6px; font-size: 12px; font-weight: bold; cursor: pointer; transition: background 0.2s; width: 100%; }
      .tmee-btn-primary:hover { background-color: #059669; }
      .tmee-btn-primary.tmee-btn-inject { width: auto; padding: 4px 10px; font-size: 11px; display: flex; align-items: center; gap: 4px; white-space: nowrap; }
      .tmee-btn-action { background-color: #3b82f6; color: #ffffff; width: auto; padding: 4px 10px; font-size: 11px; display: flex; align-items: center; gap: 4px; border: none; border-radius: 4px; cursor: pointer; }
      .tmee-btn-action:hover { background-color: #2563eb; }
      .tmee-btn-action.csv { background-color: #8b5cf6; }
      .tmee-btn-action.csv:hover { background-color: #7c3aed; }
      .tmee-results-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px; gap: 4px; }
      .tmee-table-title { font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em; color: #9ca3af; }
      .tmee-table-wrapper { max-height: 200px; overflow: auto; border: 1px solid #374151; border-radius: 6px; background-color: #111318; }
      .tmee-table { width: 100%; border-collapse: collapse; font-size: 11px; text-align: left; }
      .tmee-table th { background-color: #1f2937; color: #f3f4f6; padding: 6px 8px; border-bottom: 1px solid #374151; position: sticky; top: 0; }
      .tmee-table td { padding: 6px 8px; border-bottom: 1px solid #2d3748; color: #d1d5db; white-space: nowrap; }
    `;
    document.head.appendChild(styleEl);

    rootEl = document.createElement("div");
    rootEl.id = "tmee-widget-root";
    rootEl.innerHTML = `
      <div class="tmee-header" id="tmee-header" title="Arrastra para mover · Clic para minimizar">
        <div class="tmee-title"><span class="tmee-logo">📊</span> <span class="tmee-title-text">TM Excel Engine v1.0</span></div>
        <div class="tmee-header-controls">
          <span id="tmee-status" class="tmee-status-badge">Iniciando...</span>
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
            <span>Columnas: <strong style="color:#10b981" id="tmee-meta-cols">-</strong></span>
          </div>

          <div class="tmee-sheets-bar" id="tmee-sheets-bar" hidden>
            <select id="tmee-sheets-select" class="tmee-select" aria-label="Hoja del libro"></select>
            <span class="tmee-sheets-summary" id="tmee-sheets-summary"></span>
          </div>
          
          <div class="tmee-sql-box">
            <div class="tmee-table-title">Consola SQL</div>
            <textarea id="tmee-sql-input" class="tmee-textarea" spellcheck="false"></textarea>
            <button id="tmee-btn-run-sql" class="tmee-btn-primary">▶ Ejecutar Consulta</button>
          </div>

          <div style="margin-top: 8px;">
            <div class="tmee-results-header">
              <div class="tmee-table-title" id="tmee-results-title">Resultados</div>
              <div style="display:flex; gap:6px;">
                <button id="tmee-btn-export" class="tmee-btn-action csv">⬇️ CSV</button>
                <button id="tmee-btn-inject" class="tmee-btn-primary tmee-btn-inject" style="padding: 4px 8px; font-size: 10px;">💬 Enviar a TM</button>
              </div>
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
    headerEl = rootEl.querySelector(".tmee-header");
    setupEvents();
  }

  // --- 3. Lógica de Eventos ---
  function setupEvents() {
    const dropzone = document.getElementById("tmee-dropzone");
    const fileInput = document.getElementById("tmee-file-input");

    // Fase 5: la cabecera es el asa de arrastre y también el interruptor minimizar/maximizar
    setupDrag();

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
    document.getElementById("tmee-btn-export").addEventListener("click", exportarCSV);

    // Fase 6: el desplegable de hojas cambia la tabla activa sin recargar el archivo
    document.getElementById("tmee-sheets-select").addEventListener("change", (e) => seleccionarHoja(e.target.value));

    // Fase 5: al girar el móvil o redimensionar, el widget se mantiene dentro de la pantalla
    window.addEventListener("resize", reajustarPosicion);
  }

  // --- 3.1 Fase 5: arrastre del widget por la pantalla (ratón + táctil) ---
  /** Normaliza las coordenadas del puntero para eventos de ratón y eventos táctiles. */
  function obtenerPuntoPuntero(e) {
    if (e.touches && e.touches.length > 0) return { x: e.touches[0].clientX, y: e.touches[0].clientY };
    if (e.changedTouches && e.changedTouches.length > 0) return { x: e.changedTouches[0].clientX, y: e.changedTouches[0].clientY };
    return { x: e.clientX, y: e.clientY };
  }

  /** Evita que el widget quede total o parcialmente fuera de la ventana visible. */
  function limitarDentroDeViewport(left, top) {
    if (!rootEl) return { left: left, top: top };
    const ancho = rootEl.offsetWidth || 50;
    const alto = rootEl.offsetHeight || 50;
    const maxLeft = Math.max(0, window.innerWidth - ancho);
    const maxTop = Math.max(0, window.innerHeight - alto);
    return {
      left: Math.min(Math.max(0, left), maxLeft),
      top: Math.min(Math.max(0, top), maxTop)
    };
  }

  /** Recoloca el widget dentro de la pantalla (se usa al cambiar de tamaño o de orientación). */
  function reajustarPosicion() {
    // Si nunca se arrastró, el widget conserva su anclaje original (bottom/left)
    if (!rootEl || !rootEl.style.left) return;

    const ajustar = () => {
      if (!rootEl) return;
      const rect = rootEl.getBoundingClientRect();
      const destino = limitarDentroDeViewport(rect.left, rect.top);
      rootEl.style.left = destino.left + "px";
      rootEl.style.top = destino.top + "px";
    };

    requestAnimationFrame(ajustar);
    setTimeout(ajustar, 320); // al terminar la transición de tamaño (0.3s)
  }

  /** Registra el asa de arrastre: cabecera expandida o icono circular minimizado. */
  function setupDrag() {
    if (!headerEl) return;
    headerEl.addEventListener("mousedown", iniciarArrastre);
    headerEl.addEventListener("touchstart", iniciarArrastre, { passive: true });
    headerEl.addEventListener("mouseup", soltarSobreCabecera);
    headerEl.addEventListener("touchend", soltarSobreCabecera);
  }

  /** Inicia el gesto: guarda el offset del puntero respecto al widget y conecta los listeners de movimiento. */
  function iniciarArrastre(e) {
    if (!rootEl) return;

    if (e.type === "mousedown") {
      // El navegador emite mousedown/mouseup sintéticos tras un toque: se ignoran para no alternar dos veces
      if (Date.now() - ultimoTouchTs < VENTANA_SUPRESION_MOUSE_MS) return;
    } else if (e.type === "touchstart") {
      ultimoTouchTs = Date.now();
      if (e.touches && e.touches.length > 1) return; // gestos multitáctiles: no arrastrar
    }

    const punto = obtenerPuntoPuntero(e);
    const rect = rootEl.getBoundingClientRect();

    dragOffsetX = punto.x - rect.left;
    dragOffsetY = punto.y - rect.top;
    dragStartX = punto.x;
    dragStartY = punto.y;
    dragActivo = true;
    isDragging = false;

    document.addEventListener("mousemove", moverArrastre);
    document.addEventListener("touchmove", moverArrastre, OPCIONES_TOUCH_MOVE);
    document.addEventListener("mouseup", finalizarArrastre);
    document.addEventListener("touchend", finalizarArrastre);
    document.addEventListener("touchcancel", finalizarArrastre);
  }

  /** Actualiza la posición del widget mientras el puntero se mueve (listener de document). */
  function moverArrastre(e) {
    if (!rootEl || !dragActivo) return;

    const punto = obtenerPuntoPuntero(e);

    if (!isDragging) {
      const distancia = Math.hypot(punto.x - dragStartX, punto.y - dragStartY);
      // Menos de 5 px de recorrido todavía puede ser un clic/tap para minimizar o maximizar
      if (distancia <= UMBRAL_ARRASTRE_PX) return;

      isDragging = true;
      // Primer movimiento real: se abandona el anclaje inferior/derecho para posicionar libremente
      rootEl.style.bottom = "auto";
      rootEl.style.right = "auto";
      rootEl.classList.add("tmee-dragging");
    }

    // En táctil, evita que la página haga scroll mientras se arrastra el widget
    if (e.type === "touchmove" && e.cancelable) e.preventDefault();

    const destino = limitarDentroDeViewport(punto.x - dragOffsetX, punto.y - dragOffsetY);
    rootEl.style.left = destino.left + "px";
    rootEl.style.top = destino.top + "px";
  }

  /** Finaliza el gesto y desconecta los listeners de movimiento (listener de document). */
  function finalizarArrastre() {
    document.removeEventListener("mousemove", moverArrastre);
    document.removeEventListener("touchmove", moverArrastre, OPCIONES_TOUCH_MOVE);
    document.removeEventListener("mouseup", finalizarArrastre);
    document.removeEventListener("touchend", finalizarArrastre);
    document.removeEventListener("touchcancel", finalizarArrastre);

    if (rootEl) rootEl.classList.remove("tmee-dragging");
    dragActivo = false;
    isDragging = false;
  }

  /**
   * mouseup/touchend sobre la cabecera: si el gesto no se convirtió en arrastre,
   * se interpreta como clic y se alterna minimizar/maximizar.
   */
  function soltarSobreCabecera(e) {
    if (dragActivo && !isDragging) {
      // En táctil, cancelar el touchend evita que el navegador emita los eventos de ratón sintéticos
      if (e && e.type === "touchend" && e.cancelable) e.preventDefault();
      toggleWidget();
    }
    // La limpieza del gesto la realiza finalizarArrastre() desde el listener de document
  }

  // --- 3.2 Control de visibilidad (Minimizar / Maximizar) ---
  function estaMinimizado() {
    return !!rootEl && rootEl.classList.contains("tmee-minimized");
  }

  function toggleWidget() {
    setWidgetMinimizado(!estaMinimizado());
  }

  function setWidgetMinimizado(minimizado) {
    if (!rootEl) return;
    rootEl.classList.toggle("tmee-minimized", !!minimizado);
    sincronizarEstadoMinimizado();
  }

  /** Ajusta los elementos dependientes del estado minimizado (tooltip y posición dentro de la pantalla). */
  function sincronizarEstadoMinimizado() {
    if (headerEl) {
      headerEl.title = estaMinimizado()
        ? "Clic para expandir · Arrastra para mover"
        : "Arrastra para mover · Clic para minimizar";
    }
    reajustarPosicion();
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
    // Fase 6: cada carga (o su fallo) descarta las hojas y la plantilla del libro anterior
    resetearSelectorHojas();
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

  // --- 4.1 Renderizado del resultado de una carga (Fase 6: libro + hoja activa) ---
  /**
   * Pinta el desenlace de una carga en el orden del widget: metadatos, selector de hojas
   * (solo libros Excel multi-hoja), plantilla de consulta y vista previa de datos.
   */
  function renderizarResultado(result) {
    document.getElementById("tmee-result").style.display = "flex";

    // Fase 6: solo los libros Excel que traen `hojas[]` alimentan el selector
    libroActual = (result.formato === "EXCEL" && Array.isArray(result.hojas)) ? result : null;
    hojaActiva = result.tabla || null;

    actualizarMetaHoja(result.tabla, result.registros, result.columnas);
    renderizarSelectorHojas();

    // Carga nueva: tablaAnterior = null fuerza la reescritura de la plantilla
    aplicarPlantillaSql(result.tabla, null);
    ejecutarSQL();
  }

  /** Vuelca los metadatos de la hoja visible (nombre de tabla, registros y nº de columnas). */
  function actualizarMetaHoja(tabla, registros, columnas) {
    const tablaEl = document.getElementById("tmee-meta-table");
    const filasEl = document.getElementById("tmee-meta-rows");
    const colsEl = document.getElementById("tmee-meta-cols");

    if (tablaEl) tablaEl.textContent = tabla ? String(tabla) : "-";

    if (filasEl) {
      const total = Number(registros);
      filasEl.textContent = (registros === null || registros === undefined || Number.isNaN(total))
        ? "-"
        : total.toLocaleString();
    }

    if (colsEl) colsEl.textContent = Array.isArray(columnas) ? String(columnas.length) : "-";
  }

  // --- 4.2 Fase 6: selector de hojas del libro cargado ---

  /**
   * Dibuja la barra de hojas a partir de `result.hojas` del libro vigente.
   * La barra permanece oculta salvo que el libro tenga más de una hoja con datos: así los
   * CSV, los Parquet y los Excel de una sola hoja quedan exactamente igual que antes.
   */
  function renderizarSelectorHojas() {
    const barra = document.getElementById("tmee-sheets-bar");
    const select = document.getElementById("tmee-sheets-select");
    const resumen = document.getElementById("tmee-sheets-summary");
    if (!barra || !select || !resumen) return;

    const hojas = (libroActual && Array.isArray(libroActual.hojas)) ? libroActual.hojas : [];
    const cargadas = hojas.filter((hoja) => hoja.estado === "ok" && hoja.tabla);

    if (cargadas.length <= 1) {
      ocultarSelectorHojas();
      return;
    }

    select.innerHTML = "";
    hojas.forEach((hoja, indice) => {
      const opcion = document.createElement("option");
      opcion.textContent = etiquetaHoja(hoja, indice);

      if (hoja.estado === "ok" && hoja.tabla) {
        opcion.value = hoja.tabla;
        if (hoja.tabla === hojaActiva) opcion.selected = true;
      } else {
        // Las hojas omitidas o con error se listan como aviso, pero no son seleccionables
        opcion.disabled = true;
      }

      select.appendChild(opcion);
    });

    const omitidas = hojas.filter((hoja) => hoja.estado === "omitida").length;
    const conError = hojas.filter((hoja) => hoja.estado === "error").length;
    const partes = [`Hojas ${cargadas.length}/${hojas.length}`];
    if (omitidas) partes.push(`${omitidas} vacía${omitidas === 1 ? "" : "s"}`);
    if (conError) partes.push(`${conError} con error`);

    resumen.textContent = partes.join(" · ");
    resumen.className = "tmee-sheets-summary" + (conError ? " tmee-error" : omitidas ? " tmee-warn" : "");
    resumen.title = hojas.map((hoja) => `${hoja.nombre}: ${descripcionEstadoHoja(hoja)}`).join("\n");
    barra.hidden = false;
  }

  /** Etiqueta legible de una pestaña dentro del desplegable de hojas. */
  function etiquetaHoja(hoja, indice) {
    const nombre = (hoja && hoja.nombre) ? String(hoja.nombre) : `Hoja ${indice + 1}`;

    if (!hoja || hoja.estado === "omitida") {
      return `${nombre} · ${(hoja && hoja.motivo === "hoja_sin_datos") ? "omitida (sin datos)" : "omitida (vacía)"}`;
    }
    if (hoja.estado === "error") return `${nombre} · error`;

    const filas = Number(hoja.registros) || 0;
    return filas === 0 ? `${nombre} · solo cabeceras` : `${nombre} · ${filas.toLocaleString()} filas`;
  }

  /** Desenlace de una hoja descrito en texto (tooltip del resumen). */
  function descripcionEstadoHoja(hoja) {
    if (!hoja || !hoja.estado) return "sin información";
    if (hoja.estado === "omitida") return hoja.motivo === "hoja_sin_datos" ? "omitida (sin datos)" : "omitida (vacía)";
    if (hoja.estado === "error") return "error" + (hoja.error ? `: ${hoja.error}` : "");
    if (hoja.advertencia === "solo_cabeceras") return "cargada (solo cabeceras)";
    return `cargada (${(hoja.columnas || []).length} columnas)`;
  }

  /** Oculta la barra de hojas y descarta las opciones dibujadas. */
  function ocultarSelectorHojas() {
    const barra = document.getElementById("tmee-sheets-bar");
    const select = document.getElementById("tmee-sheets-select");
    const resumen = document.getElementById("tmee-sheets-summary");

    if (select) select.innerHTML = "";
    if (resumen) {
      resumen.textContent = "";
      resumen.className = "tmee-sheets-summary";
      resumen.title = "";
    }
    if (barra) barra.hidden = true;
  }

  /** Reinicia el estado del libro; se invoca al empezar cualquier carga nueva. */
  function resetearSelectorHojas() {
    libroActual = null;
    hojaActiva = null;
    plantillaSqlActual = "";
    ocultarSelectorHojas();
  }

  /**
   * Cambia la hoja visible sin recargar el archivo: los metadatos y la vista previa se
   * alimentan de la tabla que ese libro ya materializó en DuckDB (excel_data_<hoja>).
   * Solo se aceptan hojas con estado "ok"; las omitidas y las que fallaron se ignoran.
   */
  function seleccionarHoja(tabla) {
    if (!libroActual || !tabla || tabla === hojaActiva) return;

    const hoja = libroActual.hojas.find((candidata) => candidata.estado === "ok" && candidata.tabla === tabla);
    if (!hoja) return;

    const anterior = hojaActiva;
    hojaActiva = tabla;

    actualizarMetaHoja(hoja.tabla, hoja.registros, hoja.columnas);
    aplicarPlantillaSql(tabla, anterior);
    ejecutarSQL();
  }

  /**
   * Fija la consola SQL sobre una tabla y guarda la plantilla vigente.
   * @param {string} tablaNueva tabla que pasa a estar activa.
   * @param {?string} tablaAnterior tabla de la que se viene; null en una carga nueva.
   *
   * Regla conservadora aprobada: si el usuario editó la consulta, solo se sustituye el
   * token `FROM <tabla anterior>`, de modo que su SELECT completo se conserva intacto.
   */
  function aplicarPlantillaSql(tablaNueva, tablaAnterior = null) {
    const sqlInput = document.getElementById("tmee-sql-input");
    const plantilla = `SELECT * FROM ${tablaNueva} LIMIT 10;`;

    if (sqlInput) {
      const actual = sqlInput.value.trim();
      const sinEditar = !actual || actual === plantillaSqlActual;

      if (!tablaAnterior || sinEditar) {
        sqlInput.value = plantilla;
      } else {
        const patron = new RegExp(`(FROM\\s+)${escaparRegex(tablaAnterior)}(?=$|[^A-Za-z0-9_])`, "gi");
        const conSustitucion = sqlInput.value.replace(patron, `$1${tablaNueva}`);
        if (conSustitucion !== sqlInput.value) sqlInput.value = conSustitucion;
      }
    }

    plantillaSqlActual = plantilla;
  }

  /** Escapa un texto para poder usarlo como literal dentro de una RegExp. */
  function escaparRegex(texto) {
    return String(texto).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  // --- 5. Motor SQL Dinámico Manual ---
  async function ejecutarSQL() {
    const query = document.getElementById("tmee-sql-input").value;
    if (!query.trim()) return;

    try {
      updateStatus("Ejecutando SQL...", "pending");
      const rows = await sendRequest("executeQuery", { sql: query }, 60000);
      // Fase 3B: se conservan los resultados para poder enviarlos al chat
      ultimosResultados = rows;
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
    const titleEl = document.getElementById("tmee-results-title");

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

  // --- 6. Acciones Extras (Inyección de Contexto a la IA y Exportación CSV) ---

  /**
   * Escapa el contenido de una celda para que no rompa la sintaxis Markdown:
   * barras invertidas, pipes (|) y saltos de línea.
   */
  function escaparCeldaMarkdown(valor) {
    if (valor === null || valor === undefined) return "";
    return String(valor)
      .replace(/\\/g, "\\\\")
      .replace(/\|/g, "\\|")
      .replace(/\r\n|\r|\n/g, "<br>");
  }

  /**
   * Convierte las filas en una tabla Markdown, limitando el número de filas
   * para no saturar el contexto del LLM (máximo MAX_FILAS_INYECCION).
   */
  function construirTablaMarkdown(rows) {
    const filas = rows.slice(0, MAX_FILAS_INYECCION);
    const columnas = Object.keys(filas[0] || {});
    if (columnas.length === 0) return "";

    const total = rows.length;
    const aviso = total > filas.length
      ? ` — mostrando las primeras ${filas.length} de ${total} filas`
      : "";

    let md = `*Resultados del Excel Data Engine (${filas.length} fila(s)${aviso}):*\n\n`;
    md += "| " + columnas.map(escaparCeldaMarkdown).join(" | ") + " |\n";
    md += "| " + columnas.map(() => "---").join(" | ") + " |\n";
    filas.forEach((fila) => {
      md += "| " + columnas.map((col) => escaparCeldaMarkdown(fila[col])).join(" | ") + " |\n";
    });

    if (total > filas.length) {
      md += `\n_Se omitieron ${total - filas.length} filas adicionales para no saturar el contexto del modelo._\n`;
    }
    return md;
  }

  /** Localiza el input del chat de TypingMind (textarea o editor enriquecido). */
  function buscarInputChat() {
    return document.querySelector('textarea[data-testid="chat-input"]') ||
           document.querySelector('textarea[placeholder*="Type"]') ||
           document.querySelector('[contenteditable="true"]');
  }

  /** Fase 3B: envía los últimos resultados al chat de TypingMind como tabla Markdown. */
  function inyectarEnChat() {
    if (!ultimosResultados || ultimosResultados.length === 0) {
      alert("No hay resultados que enviar.\n\nEjecuta primero una consulta SQL en la consola del widget y vuelve a intentarlo.");
      return;
    }

    const md = construirTablaMarkdown(ultimosResultados);
    if (!md) {
      alert("Los resultados no contienen columnas que se puedan formatear.");
      return;
    }

    const chatInput = buscarInputChat();

    if (chatInput) {
      const esCampoSimple = chatInput.tagName === "TEXTAREA" || chatInput.tagName === "INPUT";
      const textoActual = esCampoSimple ? chatInput.value : chatInput.textContent;
      const separador = textoActual ? "\n\n" : "";

      if (esCampoSimple) chatInput.value = textoActual + separador + md;
      else chatInput.textContent = textoActual + separador + md;

      // Evento nativo para que React (TypingMind) sincronice su estado interno
      chatInput.dispatchEvent(new Event("input", { bubbles: true }));

      try {
        chatInput.focus();
        if (typeof chatInput.setSelectionRange === "function") {
          const fin = (chatInput.value || "").length;
          chatInput.setSelectionRange(fin, fin);
        }
      } catch (_ignorado) {
        // Algunos editores enriquecidos no permiten manipular la selección
      }

      // Fase 5: el minimizado automático solo activa el modo icono flotante
      if (rootEl) rootEl.classList.add("tmee-minimized");
      sincronizarEstadoMinimizado();
      updateStatus("Enviado al chat", "success");
      return;
    }

    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(md)
        .then(() => alert("No se encontró el input del chat de TypingMind.\n\nTabla copiada al portapapeles. Pégala en el chat con Ctrl+V."))
        .catch(() => alert("No se encontró el input del chat ni se pudo acceder al portapapeles.\n\nCopia manualmente:\n\n" + md));
    } else {
      alert("No se encontró el input del chat ni el portapapeles.\n\nCopia manualmente:\n\n" + md);
    }
  }

  function exportarCSV() {
    if (!ultimosResultados || ultimosResultados.length === 0) return alert("No hay datos para exportar.");
    const columns = Object.keys(ultimosResultados[0]);
    const header = columns.map(c => `"${String(c).replace(/"/g, '""')}"`).join(",");
    const rows = ultimosResultados.map(row => {
        return columns.map(col => {
            const val = row[col];
            if (val === null || val === undefined) return '""';
            return `"${String(val).replace(/"/g, '""')}"`;
        }).join(",");
    });
    const csvContent = [header, ...rows].join("\n");
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", "tmee_resultados.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => { injectWidget(); initWorker(); });
  } else {
    injectWidget(); initWorker();
  }
})();
