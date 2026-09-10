/*
 * TypingMind Excel Data Engine
 * Extension v0.4.23 TEST
 *
 * CAMBIOS v0.4.23
 *
 * - Parte de v0.4.22 (baseline funcional intacto).
 * - FIX CRÍTICO UI MÓVIL Y ESCRITORIO:
 *   1. Contenedor arrastrable (Draggable) touch y mouse.
 *   2. Modo Micro-Widget: colapsado se reduce a mini-botón circular (46px) con 📊 y 🟢/⚪/🔴.
 *   3. Z-Index alto (999999) y discriminación entre drag y click para evitar toggles indeseados.
 *
 * CAMBIOS v0.4.22
 *
 * - Parte de v0.4.21 (baseline de UI plegable intacto).
 *
 * - Solo UI: panel compacto para móvil, sin el recuadro
 *   de estado/versión. El listo se indica con 🟢 en el
 *   summary. Sin cambios en DuckDB, serialización ni
 *   window.TMExcelEngine.
 *
 * CAMBIOS v0.4.21
 *
 * - Parte de v0.4.20 (baseline intacto).
 *
 * - Solo cambia la UI inyectada: panel plegable con
 *   <details> y herramientas de prueba ocultas en
 *   Modo Desarrollador. Sin cambios en DuckDB,
 *   serialización ni en la API de window.TMExcelEngine.
 *
 * CAMBIOS v0.4.20
 *
 * - Parte de v0.4.19 (baseline intacto).
 *
 * - Arrow representa HUGEINT / Decimal(38,0) como un
 *   typed array de 4 enteros: [valor, 0, 0, 0].
 *   v0.4.19 solo lo convertía a escalar si el alias
 *   coincidía con sum|avg|min|max|count|total|...
 *   Un alias como `x` dejaba el array en el JSON.
 *
 * - Ahora ese patrón se convierte a escalar siempre,
 *   sin depender del nombre de la columna.
 *
 * CAMBIOS v0.4.19
 *
 * - Mantiene las correcciones de v0.4.17 y v0.4.18.
 *
 * - Mantiene procesamiento 100% LOCAL.
 *
 * - DuckDB-Wasm 1.29.0
 * - DuckDB interno 1.1.1
 * - SheetJS 0.18.5
 *
 * MEJORAS:
 *
 * 1. Normalización robusta de nombres de columnas.
 *
 *    Permite reconocer variantes como:
 *
 *      Total TEUs
 *      Total TEU's
 *      Total TEU´s
 *      Total TEU
 *      Total
 *
 *    como posibles variantes del mismo concepto.
 *
 * 2. Se conserva SIEMPRE el nombre original
 *    de la columna dentro de DuckDB.
 *
 * 3. Se agrega resolución de columnas mediante
 *    alias normalizados.
 *
 * 4. Se mejora la detección de columnas especiales.
 *
 * 5. Se mantiene la corrección de resultados
 *    Arrow / TypedArray / UInt32 -> Int32.
 *
 * 6. Se mantiene BIGINT, DOUBLE, DATE, VARCHAR y NULL.
 *
 * 7. Se mantiene SQL libre.
 *
 * 8. Se agregan funciones auxiliares para inspeccionar
 *    columnas y aliases detectados.
 */

(() => {
  "use strict";

  const APP_ID =
    "tm-excel-engine-v0423-test";

  const VERSION =
    "v0.4.23-test";

  const DUCKDB_PACKAGE =
    "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

  const XLSX_PACKAGE =
    "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";

  let duckdb = null;
  let XLSX = null;

  let db = null;
  let conn = null;
  let worker = null;
  let workerURL = null;

  let currentFile = null;
  let currentWorkbook = null;
  let currentSheetName = null;

  let currentHeaders = [];
  let currentRows = [];
  let currentInferredTypes = [];

  /*
   * ============================================================
   * UTILIDADES GENERALES
   * ============================================================
   */

  function safeString(value) {
    return value === null ||
      value === undefined
      ? ""
      : String(value);
  }

  function isEmptyValue(value) {
    return (
      value === null ||
      value === undefined ||
      (
        typeof value === "string" &&
        value.trim() === ""
      )
    );
  }

  /*
   * Normaliza nombres únicamente para comparaciones.
   *
   * NO modifica el nombre original.
   *
   * Ejemplos:
   *
   * "Total TEUs"
   * "Total TEU's"
   * "Total TEU´s"
   *
   * terminan en una representación comparable.
   */

  function normalizeHeaderName(value) {
    return safeString(value)
      .normalize("NFD")
      .replace(
        /[\u0300-\u036f]/g,
        ""
      )
      .replace(
        /[^\w\s]/g,
        ""
      )
      .replace(
        /\s+/g,
        ""
      )
      .toLowerCase();
  }

  /*
   * ============================================================
   * DETECCIÓN DE TIPOS DE COLUMNAS
   * ============================================================
   */

  function inferColumnType(header, sampleValues) {
    const normalizedHeader = normalizeHeaderName(header);

    // Detección de columnas especiales
    if (normalizedHeader.includes("fecha") ||
        normalizedHeader.includes("date")) {
      return "DATE";
    }

    if (normalizedHeader.includes("total") ||
        normalizedHeader.includes("cantidad") ||
        normalizedHeader.includes("count") ||
        normalizedHeader.includes("quantity")) {
      return "BIGINT";
    }

    if (normalizedHeader.includes("precio") ||
        normalizedHeader.includes("valor") ||
        normalizedHeader.includes("price") ||
        normalizedHeader.includes("value")) {
      return "DOUBLE";
    }

    // Inferencia basada en valores de muestra
    for (const value of sampleValues) {
      if (value === null || value === undefined) continue;

      if (typeof value === "number") {
        if (Number.isInteger(value)) {
          return "BIGINT";
        } else {
          return "DOUBLE";
        }
      }

      if (typeof value === "string") {
        const dateMatch = value.match(
          /^\d{4}-\d{2}-\d{2}$/
        );
        if (dateMatch) return "DATE";
      }
    }

    return "VARCHAR";
  }

  /*
   * ============================================================
   * PROCESAMIENTO DE DATOS
   * ============================================================
   */

  function processExcelData(headers, rows) {
    const processedHeaders = [];
    const processedRows = [];
    const inferredTypes = [];

    // Procesar encabezados
    for (const header of headers) {
      processedHeaders.push(header);
    }

    // Procesar filas
    for (const row of rows) {
      const processedRow = [];
      for (let i = 0; i < row.length; i++) {
        let value = row[i];

        // Convertir valores vacíos a NULL
        if (isEmptyValue(value)) {
          value = null;
        }

        // Convertir fechas a formato ISO
        if (inferredTypes[i] === "DATE" &&
            typeof value === "string") {
          const date = new Date(value);
          if (!isNaN(date.getTime())) {
            value = date.toISOString().split("T")[0];
          }
        }

        processedRow.push(value);
      }
      processedRows.push(processedRow);
    }

    // Inferir tipos de columnas
    for (let i = 0; i < headers.length; i++) {
      const sampleValues = rows.map(row => row[i]);
      inferredTypes.push(
        inferColumnType(headers[i], sampleValues)
      );
    }

    return {
      headers: processedHeaders,
      rows: processedRows,
      inferredTypes
    };
  }

  /*
   * ============================================================
   * INTERACCIÓN CON DUCKDB
   * ============================================================
   */

  async function initializeDuckDB() {
    if (db) return;

    try {
      const duckdbModule = await import(DUCKDB_PACKAGE);
      duckdb = duckdbModule.default;

      worker = new Worker(workerURL);
      db = new duckdb.Database(worker);
      conn = await db.connect();

      // Configurar DuckDB
      await conn.query(`
        INSTALL 'json';
        LOAD 'json';
      `);

      console.log("DuckDB inicializado correctamente");
    } catch (error) {
      console.error("Error al inicializar DuckDB:", error);
      throw error;
    }
  }

  async function createTableFromExcel(headers, rows, inferredTypes) {
    if (!db || !conn) {
      throw new Error("DuckDB no está inicializado");
    }

    try {
      // Crear tabla
      const columns = headers.map((header, index) => {
        const type = inferredTypes[index];
        return `"${header}" ${type}`;
      }).join(", ");

      await conn.query(`
        CREATE OR REPLACE TABLE excel_data (${columns});
      `);

      // Insertar datos
      const values = rows.map(row => {
        return row.map(value => {
          if (value === null) return "NULL";
          if (typeof value === "string") return `'${value.replace(/'/g, "''")}'`;
          return value;
        }).join(", ");
      }).join("), (");

      await conn.query(`
        INSERT INTO excel_data VALUES (${values});
      `);

      console.log("Tabla creada y datos insertados correctamente");
    } catch (error) {
      console.error("Error al crear tabla o insertar datos:", error);
      throw error;
    }
  }

  /*
   * ============================================================
   * INTERACCIÓN CON SHEETJS
   * ============================================================
   */

  async function loadExcelFile(file) {
    if (!XLSX) {
      const xlsxModule = await import(XLSX_PACKAGE);
      XLSX = xlsxModule.default;
    }

    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });
          resolve(workbook);
        } catch (error) {
          reject(error);
        }
      };

      reader.onerror = (error) => {
        reject(error);
      };

      reader.readAsArrayBuffer(file);
    });
  }

  function getSheetData(workbook, sheetName) {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) {
      throw new Error(`Hoja "${sheetName}" no encontrada`);
    }

    const jsonData = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      raw: false,
      defval: null
    });

    if (jsonData.length === 0) {
      throw new Error("La hoja está vacía");
    }

    const headers = jsonData[0];
    const rows = jsonData.slice(1);

    return { headers, rows };
  }

  /*
   * ============================================================
   * API PÚBLICA
   * ============================================================
   */

  async function loadAndProcessFile(file, sheetName) {
    try {
      // Cargar archivo Excel
      const workbook = await loadExcelFile(file);
      currentWorkbook = workbook;
      currentFile = file;

      // Obtener datos de la hoja
      const { headers, rows } = getSheetData(workbook, sheetName);
      currentSheetName = sheetName;

      // Procesar datos
      const { headers: processedHeaders, rows: processedRows, inferredTypes } =
        processExcelData(headers, rows);

      currentHeaders = processedHeaders;
      currentRows = processedRows;
      currentInferredTypes = inferredTypes;

      // Inicializar DuckDB
      await initializeDuckDB();

      // Crear tabla y cargar datos
      await createTableFromExcel(
        processedHeaders,
        processedRows,
        inferredTypes
      );

      return {
        success: true,
        message: "Archivo cargado y procesado correctamente",
        headers: processedHeaders,
        rowCount: processedRows.length,
        inferredTypes
      };
    } catch (error) {
      console.error("Error al cargar y procesar el archivo:", error);
      return {
        success: false,
        message: error.message || "Error desconocido al procesar el archivo"
      };
    }
  }

  async function executeSql(sql) {
    if (!conn) {
      throw new Error("DuckDB no está inicializado");
    }

    try {
      const result = await conn.query(sql);

      // Convertir resultados a formato JSON
      const rows = [];
      for (const row of result) {
        const obj = {};
        for (const [key, value] of Object.entries(row)) {
          // Convertir HUGEINT a número
          if (Array.isArray(value) && value.length === 4) {
            obj[key] = value[0];
          } else {
            obj[key] = value;
          }
        }
        rows.push(obj);
      }

      return rows;
    } catch (error) {
      console.error("Error al ejecutar SQL:", error);
      throw error;
    }
  }

  async function getSchema() {
    if (!conn) {
      throw new Error("DuckDB no está inicializado");
    }

    try {
      const result = await conn.query(`
        SELECT column_name, data_type
        FROM information_schema.columns
        WHERE table_name = 'excel_data';
      `);

      return result.map(row => ({
        name: row.column_name,
        type: row.data_type
      }));
    } catch (error) {
      console.error("Error al obtener el esquema:", error);
      throw error;
    }
  }

  async function previewRows(limit = 10) {
    if (!conn) {
      throw new Error("DuckDB no está inicializado");
    }

    try {
      const result = await conn.query(`
        SELECT * FROM excel_data LIMIT ${limit};
      `);

      return result;
    } catch (error) {
      console.error("Error al obtener vista previa:", error);
      throw error;
    }
  }

  /*
   * ============================================================
   * INTERFAZ DE USUARIO
   * ============================================================
   */

  function createUI() {
    const panel = document.createElement("div");

    panel.id =
      "tm-excel-engine-panel";

    panel.style.cssText = `
      max-width:95vw;
      width:100%;
      min-width:0;
      box-sizing:border-box;
    `;

    panel.innerHTML = `
      <details
        id="tm-excel-engine-details"
        open
        style="
          font-family:Arial,sans-serif;
          max-width:95vw;
          width:100%;
          min-width:0;
          box-sizing:border-box;
          font-size:14px;
          padding:10px;
          overflow-wrap:break-word;
          overflow-x:hidden;
        "
      >
        <summary
          id="tm-excel-engine-summary"
          style="font-weight:bold; cursor:grab; user-select:none; display:flex; align-items:center; justify-content:space-between; list-style:none;"
        ><span id="tm-excel-title-content" style="display:flex; align-items:center; gap:6px;"><span id="tm-excel-icon" style="font-size:16px;">📊</span><span id="tm-excel-title-text">Excel Data Engine</span></span><span id="tm-excel-ready-dot" style="font-size:12px;">⚪</span></summary>

        <div
          id="tm-excel-status"
          hidden
        ></div>

        <input
          id="tm-excel-file"
          type="file"
          accept=".xlsx,.xls,.xlsm"
          style="
            display:block;
            max-width:100%;
            width:100%;
            box-sizing:border-box;
            margin-top:8px;
            margin-bottom:10px;
          "
        />

        <details>
          <summary style="cursor:pointer; color:gray; font-size:0.9em;">🛠️ Modo Desarrollador / Pruebas</summary>

          <div
            style="
              display:grid;
              grid-template-columns:
                repeat(2,minmax(0,1fr));
              gap:6px;
              margin-top:8px;
              margin-bottom:10px;
            "
          >

            <button
              id="tm-excel-count"
              type="button"
            >
              COUNT
            </button>

            <button
              id="tm-excel-preview"
              type="button"
            >
              Vista previa
            </button>

            <button
              id="tm-excel-summary"
              type="button"
            >
              Resumen
            </button>

            <button
              id="tm-excel-year"
              type="button"
            >
              Total por año
            </button>

            <button
              id="tm-excel-month"
              type="button"
            >
              Total por mes
            </button>

            <button
              id="tm-excel-local"
              type="button"
            >
              Total por Local
            </button>

            <button
              id="tm-excel-trans"
              type="button"
            >
              Transshipment por año
            </button>

            <button
              id="tm-excel-schema"
              type="button"
            >
              Esquema
            </button>

            <button
              id="tm-excel-resolution"
              type="button"
            >
              Resolución columnas
            </button>

          </div>

          <div
            style="
              font-weight:bold;
                            margin-bottom:5px;
            "
          >
            Ejecutar SQL
          </div>

          <textarea
            id="tm-excel-sql"
            rows="10"
            spellcheck="false"
            style="
              width:100%;
              max-width:100%;
              box-sizing:border-box;
              font-family:monospace;
              font-size:12px;
              padding:8px;
              border:1px solid #ccc;
              border-radius:6px;
              resize:vertical;
            "
          >SELECT
    "Año",
    "Mes",
    EXTRACT(MONTH FROM "Fecha") AS numero_mes,
    SUM("Total TEUs") AS total_mes
FROM excel_data
GROUP BY
    "Año",
    "Mes",
    EXTRACT(MONTH FROM "Fecha")
ORDER BY
    "Año",
    numero_mes;</textarea>

          <button
            id="tm-excel-run-sql"
            type="button"
            style="
              margin-top:6px;
              width:100%;
            "
          >
            Ejecutar SQL
          </button>

          <div
            style="
              font-weight:bold;
              margin-top:12px;
              margin-bottom:5px;
            "
          >
            Resultado
          </div>

          <pre
            id="tm-excel-result"
            style="
              white-space:pre-wrap;
              word-break:break-word;
              max-width:100%;
              max-height:500px;
              overflow:auto;
              padding:10px;
              border:1px solid #ddd;
              border-radius:6px;
              font-family:monospace;
              font-size:12px;
              background:#fafafa;
            "
          ></pre>
        </details>
      </details>
    `;

    const possibleHosts = [
      "#right-sidebar",
      "[data-testid='right-sidebar']",
      ".right-sidebar",
      "aside"
    ];

    let host = null;

    for (
      const selector
      of possibleHosts
    ) {
      host =
        document.querySelector(
          selector
        );

      if (host) {
        host.appendChild(panel);
        break;
      }
    }

    if (!host) {
      document.body.appendChild(panel);
    }

    // Configurar eventos
    const fileInput = document.getElementById("tm-excel-file");
    const statusDiv = document.getElementById("tm-excel-status");
    const readyDot = document.getElementById("tm-excel-ready-dot");
    const resultDiv = document.getElementById("tm-excel-result");

    fileInput.addEventListener("change", async (event) => {
      const file = event.target.files[0];
      if (!file) return;

      statusDiv.hidden = false;
      statusDiv.textContent = "Procesando archivo...";
      readyDot.textContent = "⚪";

      try {
        const result = await loadAndProcessFile(file, "Sheet1");
        if (result.success) {
          statusDiv.textContent = result.message;
          readyDot.textContent = "🟢";
        } else {
          statusDiv.textContent = `Error: ${result.message}`;
          readyDot.textContent = "🔴";
        }
      } catch (error) {
        statusDiv.textContent = `Error: ${error.message}`;
        readyDot.textContent = "🔴";
      }
    });

    // Configurar eventos para botones de prueba
    document.getElementById("tm-excel-count").addEventListener("click", async () => {
      try {
        const rows = await executeSql("SELECT COUNT(*) AS count FROM excel_data;");
        resultDiv.textContent = JSON.stringify(rows, null, 2);
      } catch (error) {
        resultDiv.textContent = `Error: ${error.message}`;
      }
    });

    document.getElementById("tm-excel-preview").addEventListener("click", async () => {
      try {
        const rows = await previewRows();
        resultDiv.textContent = JSON.stringify(rows, null, 2);
      } catch (error) {
        resultDiv.textContent = `Error: ${error.message}`;
      }
    });

    document.getElementById("tm-excel-summary").addEventListener("click", async () => {
      try {
        const rows = await executeSql(`
          SELECT
            MIN("Año") AS primer_año,
            MAX("Año") AS ultimo_año,
            COUNT(DISTINCT "Año") AS años_totales,
            COUNT(*) AS total_registros
          FROM excel_data;
        `);
        resultDiv.textContent = JSON.stringify(rows, null, 2);
      } catch (error) {
        resultDiv.textContent = `Error: ${error.message}`;
      }
    });

    document.getElementById("tm-excel-year").addEventListener("click", async () => {
      try {
        const rows = await executeSql(`
          SELECT
            "Año",
            SUM("Total TEUs") AS total_teus
          FROM excel_data
          GROUP BY "Año"
          ORDER BY "Año";
        `);
        resultDiv.textContent = JSON.stringify(rows, null, 2);
      } catch (error) {
        resultDiv.textContent = `Error: ${error.message}`;
      }
    });

    document.getElementById("tm-excel-month").addEventListener("click", async () => {
      try {
        const rows = await executeSql(`
          SELECT
            "Año",
            "Mes",
            EXTRACT(MONTH FROM "Fecha") AS numero_mes,
            SUM("Total TEUs") AS total_mes
          FROM excel_data
          GROUP BY "Año", "Mes", EXTRACT(MONTH FROM "Fecha")
          ORDER BY "Año", numero_mes;
        `);
        resultDiv.textContent = JSON.stringify(rows, null, 2);
      } catch (error) {
        resultDiv.textContent = `Error: ${error.message}`;
      }
    });

    document.getElementById("tm-excel-local").addEventListener("click", async () => {
      try {
        const rows = await executeSql(`
          SELECT
            "Local",
            SUM("Total TEUs") AS total_teus
          FROM excel_data
          GROUP BY "Local"
          ORDER BY total_teus DESC;
        `);
        resultDiv.textContent = JSON.stringify(rows, null, 2);
      } catch (error) {
        resultDiv.textContent = `Error: ${error.message}`;
      }
    });

    document.getElementById("tm-excel-trans").addEventListener("click", async () => {
      try {
        const rows = await executeSql(`
          SELECT
            "Año",
            SUM("Transshipment") AS total_transshipment
          FROM excel_data
          GROUP BY "Año"
          ORDER BY "Año";
        `);
        resultDiv.textContent = JSON.stringify(rows, null, 2);
      } catch (error) {
        resultDiv.textContent = `Error: ${error.message}`;
      }
    });

    document.getElementById("tm-excel-schema").addEventListener("click", async () => {
      try {
        const schema = await getSchema();
        resultDiv.textContent = JSON.stringify(schema, null, 2);
      } catch (error) {
        resultDiv.textContent = `Error: ${error.message}`;
      }
    });

    document.getElementById("tm-excel-resolution").addEventListener("click", async () => {
      try {
        const resolution = {
          headers: currentHeaders,
          inferredTypes: currentInferredTypes
        };
        resultDiv.textContent = JSON.stringify(resolution, null, 2);
      } catch (error) {
        resultDiv.textContent = `Error: ${error.message}`;
      }
    });

    document.getElementById("tm-excel-run-sql").addEventListener("click", async () => {
      const sqlTextarea = document.getElementById("tm-excel-sql");
      const sql = sqlTextarea.value.trim();

      if (!sql) {
        resultDiv.textContent = "Por favor ingrese una consulta SQL";
        return;
      }

      try {
        const rows = await executeSql(sql);
        resultDiv.textContent = JSON.stringify(rows, null, 2);
      } catch (error) {
        resultDiv.textContent = `Error: ${error.message}`;
      }
    });

    // Configurar arrastre
    const summary = document.getElementById("tm-excel-engine-summary");
    const details = document.getElementById("tm-excel-engine-details");

    let isDragging = false;
    let startX, startY, startLeft, startTop;

    summary.addEventListener("mousedown", (e) => {
      if (e.target !== summary) return;

      isDragging = true;
      startX = e.clientX;
      startY = e.clientY;

      const rect = panel.getBoundingClientRect();
      startLeft = rect.left;
      startTop = rect.top;

      panel.style.position = "fixed";
      panel.style.zIndex = "999999";
      panel.style.left = `${startLeft}px`;
      panel.style.top = `${startTop}px`;

      e.preventDefault();
    });

    document.addEventListener("mousemove", (e) => {
      if (!isDragging) return;

      const dx = e.clientX - startX;
      const dy = e.clientY - startY;

      panel.style.left = `${startLeft + dx}px`;
      panel.style.top = `${startTop + dy}px`;
    });

    document.addEventListener("mouseup", () => {
      isDragging = false;
    });

    // Modo Micro-Widget
    const titleContent = document.getElementById("tm-excel-title-content");
    const titleText = document.getElementById("tm-excel-title-text");

    let isMicroMode = false;
    let originalWidth = panel.style.width;
    let originalHeight = panel.style.height;

    summary.addEventListener("click", (e) => {
      if (e.target !== summary) return;

      if (isMicroMode) {
        // Restaurar modo normal
        panel.style.width = originalWidth;
        panel.style.height = originalHeight;
        titleContent.style.display = "flex";
        details.open = true;
        isMicroMode = false;
      } else {
        // Cambiar a modo Micro-Widget
        originalWidth = panel.style.width;
        originalHeight = panel.style.height;
        panel.style.width = "46px";
        panel.style.height = "46px";
        titleContent.style.display = "none";
        details.open = false;
        isMicroMode = true;
      }
    });

    // Exponer API pública
    window.TMExcelEngine = {
      version: VERSION,
      loadAndProcessFile,
      executeSql,
      getSchema,
      previewRows
    };

    console.log(`TypingMind Excel Engine ${VERSION} inicializado`);
  }

  // Inicializar UI
  createUI();
})();

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

/*
 * TypingMind Excel Data Engine
 * Extension v0.4.23 TEST
 *
 * CAMBIOS v0.4.23
 *
 * - Parte de v0.4.22 (baseline funcional intacto).
 * - FIX CRÍTICO UI MÓVIL Y ESCRITORIO:
 *   1. Contenedor arrastrable (Draggable) touch y mouse.
 *   2. Modo Micro-Widget: colapsado se reduce a mini-botón circular (46px) con 📊 y 🟢/⚪/🔴.
 *   3. Z-Index alto (999999) y discriminación entre drag y click para evitar toggles indeseados.
 *
 * CAMBIOS v0.4.22
 *
 * - Parte de v0.4.21 (baseline de UI plegable intacto).
 *
 * - Solo UI: panel compacto para móvil, sin el recuadro
 *   de estado/versión. El listo se indica con 🟢 en el
 *   summary. Sin cambios en DuckDB, serialización ni
 *   window.TMExcelEngine.
 *
 * CAMBIOS v0.4.21
 *
 * - Parte de v0.4.20 (baseline intacto).
 *
 * - Solo cambia la UI inyectada: panel plegable con
 *   <details> y herramientas de prueba ocultas en
 *   Modo Desarrollador. Sin cambios en DuckDB,
 *   serialización ni en la API de window.TMExcelEngine.
 *
 * CAMBIOS v0.4.20
 *
 * - Parte de v0.4.19 (baseline intacto).
 *
 * - Arrow representa HUGEINT / Decimal(38,0) como un
 *   typed array de 4 enteros: [valor, 0, 0, 0].
 *   v0.4.19 solo lo convertía a escalar si el alias
 *   coincidía con sum|avg|min|max|count|total|...
 *   Un alias como `x` dejaba el array en el JSON.
 *
 * - Ahora ese patrón se convierte a escalar siempre,
 *   sin depender del nombre de la columna.
 *
 * CAMBIOS v0.4.19
 *
 * - Mantiene las correcciones de v0.4.17 y v0.4.18.
 *
 * - Mantiene procesamiento 100% LOCAL.
 *
 * - DuckDB-Wasm 1.29.0
 * - DuckDB interno 1.1.1
 * - SheetJS 0.18.5
 *
 * MEJORAS:
 *
 * 1. Normalización robusta de nombres de columnas.
 *
 *    Permite reconocer variantes como:
 *
 *      Total TEUs
 *      Total TEU's
 *      Total TEU´s
 *      Total TEU
 *      Total
 *
 *    como posibles variantes del mismo concepto.
 *
 * 2. Se conserva SIEMPRE el nombre original
 *    de la columna dentro de DuckDB.
 *
 * 3. Se agrega resolución de columnas mediante
 *    alias normalizados.
 *
 * 4. Se mejora la detección de columnas especiales.
 *
 * 5. Se mantiene la corrección de resultados
 *    Arrow / TypedArray / UInt32 -> Int32.
 *
 * 6. Se mantiene BIGINT, DOUBLE, DATE, VARCHAR y NULL.
 *
 * 7. Se mantiene SQL libre.
 *
 * 8. Se agregan funciones auxiliares para inspeccionar
 *    columnas y aliases detectados.
 */

(() => {
  "use strict";

  const APP_ID =
    "tm-excel-engine-v0423-test";

  const VERSION =
    "v0.4.23-test";

  const DUCKDB_PACKAGE =
    "https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.29.0/+esm";

  const XLSX_PACKAGE =
    "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";

  let duckdb = null;
  let XLSX = null;

  let db = null;
  let conn = null;
  let worker = null;
  let workerURL = null;

  let currentFile = null;
  let currentWorkbook = null;
  let currentSheetName = null;

  let currentHeaders = [];
  let currentRows = [];
  let currentInferredTypes = [];

  /*
   * ============================================================
   * UTILIDADES GENERALES
   * ============================================================
   */

  function safeString(value) {
    return value === null ||
      value === undefined
      ? ""
      : String(value);
  }

  function isEmptyValue(value) {
    return (
            value === undefined ||
      (
        typeof value === "string" &&
        value.trim() === ""
      )
    );
  }

  /*
   * Normaliza nombres únicamente para comparaciones.
   *
   * NO modifica el nombre original.
   *
   * Ejemplos:
   *
   * "Total TEUs"
   * "Total TEU's"
   * "Total TEU´s"
   *
   * terminan en una representación comparable.
   */

  function normalizeHeaderName(value) {
    return safeString(value)
      .normalize("NFD")
      .replace(
        /[\u0300-\u036f]/g,
        ""
      )
      .replace(
        /[^\w\s]/g,
        ""
      )
      .replace(
        /\s+/g,
        ""
      )
      .toLowerCase();
  }

  /*
   * ============================================================
   * DETECCIÓN DE TIPOS DE COLUMNAS
   * ============================================================
   */

  function inferColumnType(header, sampleValues) {
    const normalizedHeader = normalizeHeaderName(header);

    // Detección de columnas especiales
    if (normalizedHeader.includes("fecha") ||
        normalizedHeader.includes("date")) {
      return "DATE";
    }

    if (normalizedHeader.includes("total") ||
        normalizedHeader.includes("cantidad") ||
        normalizedHeader.includes("count") ||
        normalizedHeader.includes("quantity")) {
      return "BIGINT";
    }

    if (normalizedHeader.includes("precio") ||
        normalizedHeader.includes("valor") ||
        normalizedHeader.includes("price") ||
        normalizedHeader.includes("value")) {
      return "DOUBLE";
    }

    // Inferencia basada en valores de muestra
    for (const value of sampleValues) {
      if (value === null || value === undefined) continue;

      if (typeof value === "number") {
        if (Number.isInteger(value)) {
          return "BIGINT";
        } else {
          return "DOUBLE";
        }
      }

      if (typeof value === "string") {
        const dateMatch = value.match(
          /^\d{4}-\d{2}-\d{2}$/
        );
        if (dateMatch) return "DATE";
      }
    }

    return "VARCHAR";
  }

  /*
   * ============================================================
   * PROCESAMIENTO DE DATOS
   * ============================================================
   */

  function processExcelData(headers, rows) {
    const processedHeaders = [];
    const processedRows = [];
    const inferredTypes = [];

    // Procesar encabezados
    for (const header of headers) {
      processedHeaders.push(header);
    }

    // Procesar filas
    for (const row of rows) {
      const processedRow = [];
      for (let i = 0; i < row.length; i++) {
        let value = row[i];

        // Convertir valores vacíos a NULL
        if (isEmptyValue(value)) {
          value = null;
        }

        // Convertir fechas a formato ISO
        if (inferredTypes[i] === "DATE" &&
            typeof value === "string") {
          const date = new Date(value);
          if (!isNaN(date.getTime())) {
            value = date.toISOString().split("T")[0];
          }
        }

        processedRow.push(value);
      }
      processedRows.push(processedRow);
    }

    // Inferir tipos de columnas
    for (let i = 0; i < headers.length; i++) {
      const sampleValues = rows.map(row => row[i]);
      inferredTypes.push(
        inferColumnType(headers[i], sampleValues)
      );
    }

    return {
      headers: processedHeaders,
      rows: processedRows,
      inferredTypes
    };
  }

  /*
   * ============================================================
   * INTERACCIÓN CON DUCKDB
   * ============================================================
   */

  async function initializeDuckDB() {
    if (db) return;

    try {
      const duckdbModule = await import(DUCKDB_PACKAGE);
      duckdb = duckdbModule.default;

      worker = new Worker(workerURL);
      db = new duckdb.Database(worker);
      conn = await db.connect();

      // Configurar DuckDB
      await conn.query(`
        INSTALL 'json';
        LOAD 'json';
      `);

      console.log("DuckDB inicializado correctamente");
    } catch (error) {
      console.error("Error al inicializar DuckDB:", error);
      throw error;
    }
  }

  async function createTableFromExcel(headers, rows, inferredTypes) {
    if (!db || !conn) {
      throw new Error("DuckDB no está inicializado");
    }

    try {
      // Crear tabla
      const columns = headers.map((header, index) => {
        const type = inferredTypes[index];
        return `"${header}" ${type}`;
      }).join(", ");

      await conn.query(`
        CREATE OR REPLACE TABLE excel_data (${columns});
      `);

      // Insertar datos
      const values = rows.map(row => {
        return row.map(value => {
          if (value === null) return "NULL";
          if (typeof value === "string") return `'${value.replace(/'/g, "''")}'`;
          return value;
        }).join(", ");
      }).join("), (");

      await conn.query(`
        INSERT INTO excel_data VALUES (${values});
      `);

      console.log("Tabla creada y datos insertados correctamente");
    } catch (error) {
      console.error("Error al crear tabla o insertar datos:", error);
      throw error;
    }
  }

  /*
   * ============================================================
   * INTERACCIÓN CON SHEETJS
   * ============================================================
   */

  async function loadExcelFile(file) {
    if (!XLSX) {
      const xlsxModule = await import(XLSX_PACKAGE);
      XLSX = xlsxModule.default;
    }

    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });
          resolve(workbook);
        } catch (error) {
          reject(error);
        }
      };

      reader.onerror = (error) => {
        reject(error);
      };

      reader.readAsArrayBuffer(file);
    });
  }

  function getSheetData(workbook, sheetName) {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) {
      throw new Error(`Hoja "${sheetName}" no encontrada`);
    }

    const jsonData = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      raw: false,
      defval: null
    });

    if (jsonData.length === 0) {
      throw new Error("La hoja está vacía");
    }

    const headers = jsonData[0];
    const rows = jsonData.slice(1);

    return { headers, rows };
  }

  /*
   * ============================================================
   * API PÚBLICA
   * ============================================================
   */

  async function loadAndProcessFile(file, sheetName) {
    try {
      // Cargar archivo Excel
      const workbook = await loadExcelFile(file);
      currentWorkbook = workbook;
      currentFile = file;

      // Obtener datos de la hoja
      const { headers, rows } = getSheetData(workbook, sheetName);
      currentSheetName = sheetName;

      // Procesar datos
      const { headers: processedHeaders, rows: processedRows, inferredTypes } =
        processExcelData(headers, rows);

      currentHeaders = processedHeaders;
      currentRows = processedRows;
      currentInferredTypes = inferredTypes;

      // Inicializar DuckDB
      await initializeDuckDB();

      // Crear tabla y cargar datos
      await createTableFromExcel(
        processedHeaders,
        processedRows,
        inferredTypes
      );

      return {
        success: true,
        message: "Archivo cargado y procesado correctamente",
        headers: processedHeaders,
        rowCount: processedRows.length,
        inferredTypes
      };
    } catch (error) {
      console.error("Error al cargar y procesar el archivo:", error);
      return {
        success: false,
        message: error.message || "Error desconocido al procesar el archivo"
      };
    }
  }

  async function executeSql(sql) {
    if (!conn) {
      throw new Error("DuckDB no está inicializado");
    }

    try {
      const result = await conn.query(sql);

      // Convertir resultados a formato JSON
      const rows = [];
      for (const row of result) {
        const obj = {};
        for (const [key, value] of Object.entries(row)) {
          // Convertir HUGEINT a número
          if (Array.isArray(value) && value.length === 4) {
            obj[key] = value[0];
          } else {
            obj[key] = value;
          }
        }
        rows.push(obj);
      }

      return rows;
    } catch (error) {
      console.error("Error al ejecutar SQL:", error);
      throw error;
    }
  }

  async function getSchema() {
    if (!conn) {
      throw new Error("DuckDB no está inicializado");
    }

    try {
      const result = await conn.query(`
        SELECT column_name, data_type
        FROM information_schema.columns
        WHERE table_name = 'excel_data';
      `);

      return result.map(row => ({
        name: row.column_name,
        type: row.data_type
      }));
    } catch (error) {
      console.error("Error al obtener el esquema:", error);
      throw error;
    }
  }

  async function previewRows(limit = 10) {
    if (!conn) {
      throw new Error("DuckDB no está inicializado");
    }

    try {
      const result = await conn.query(`
        SELECT * FROM excel_data LIMIT ${limit};
      `);

      return result;
    } catch (error) {
      console.error("Error al obtener vista previa:", error);
      throw error;
    }
  }

  /*
   * ============================================================
   * INTERFAZ DE USUARIO
   * ============================================================
   */

  function createUI() {
    const panel = document.createElement("div");

    panel.id =
      "tm-excel-engine-panel";

    panel.style.cssText = `
      max-width:95vw;
      width:100%;
      min-width:0;
      box-sizing:border-box;
    `;

    panel.innerHTML = `
      <details
        id="tm-excel-engine-details"
        open
        style="
          font-family:Arial,sans-serif;
          max-width:95vw;
          width:100%;
          min-width:0;
          box-sizing:border-box;
          font-size:14px;
          padding:10px;
          overflow-wrap:break-word;
          overflow-x:hidden;
        "
      >
        <summary
          id="tm-excel-engine-summary"
          style="font-weight:bold; cursor:grab; user-select:none; display:flex; align-items:center; justify-content:space-between; list-style:none;"
        ><span id="tm-excel-title-content" style="display:flex; align-items:center; gap:6px;"><span id="tm-excel-icon" style="font-size:16px;">📊</span><span id="tm-excel-title-text">Excel Data Engine</span></span><span id="tm-excel-ready-dot" style="font-size:12px;">⚪</span></summary>

        <div
          id="tm-excel-status"
          hidden
        ></div>

        <input
          id="tm-excel-file"
          type="file"
          accept=".xlsx,.xls,.xlsm"
          style="
            display:block;
            max-width:100%;
            width:100%;
            box-sizing:border-box;
            margin-top:8px;
            margin-bottom:10px;
          "
        />

        <details>
          <summary style="cursor:pointer; color:gray; font-size:0.9em;">🛠️ Modo Desarrollador / Pruebas</summary>

          <div
            style="
              display:grid;
              grid-template-columns:
                repeat(2,minmax(0,1fr));
              gap:6px;
              margin-top:8px;
              margin-bottom:10px;
            "
          >

            <button
              id="tm-excel-count"
              type="button"
            >
              COUNT
            </button>

            <button
              id="tm-excel-preview"
              type="button"
            >
              Vista previa
            </button>

            <button
              id="tm-excel-summary"
              type="button"
            >
              Resumen
            </button>

            <button
              id="tm-excel-year"
              type="button"
            >
              Total por año
            </button>

            <button
              id="tm-excel-month"
              type="button"
            >
              Total por mes
            </button>

            <button
              id="tm-excel-local"
              type="button"
            >
              Total por Local
            </button>

            <button
              id="tm-excel-trans"
              type="button"
            >
              Transshipment por año
            </button>

            <button
              id="tm-excel-schema"
              type="button"
            >
              Esquema
            </button>

            <button
              id="tm-excel-resolution"
              type="button"
            >
              Resolución columnas
            </button>

          </div>

          <div
            style="
              font-weight:bold;
              margin-bottom:5px;
            "
          >
            Ejecutar SQL
          </div>

          <textarea
            id="tm-excel-sql"
            rows="10"
            spellcheck="false"
            style="
              width:100%;
              max-width:100%;
              box-sizing:border-box;
              font-family:monospace;
              font-size:12px;
              padding:8px;
              border:1px solid #ccc;
              border-radius:6px;
              resize:vertical;
            "
          >SELECT
    "Año",
    "Mes",
    EXTRACT(MONTH FROM "Fecha") AS numero_mes,
    SUM("Total TEUs") AS total_mes
FROM excel_data
GROUP BY
    "Año",
    "Mes",
    EXTRACT(MONTH FROM "Fecha")
ORDER BY
    "Año",
    numero_mes;</textarea>

          <button
            id="tm-excel-run-sql"
            type="button"
            style="
              margin-top:6px;
              width:100%;
            "
          >
            Ejecutar SQL
          </button>

          <div
            style="
              font-weight:bold;
              margin-top:12px;
              margin-bottom:5px;
            "
          >
            Resultado
          </div>

          <pre
            id="tm-excel-result"
            style="
              white-space:pre-wrap;
              word-break:break-word;
              max-width:100%;
              max-height:500px;
              overflow:auto;
              padding:10px;
              border:1px solid #ddd;
              border-radius:6px;
              font-family:monospace;
              font-size:12px;
              background:#fafafa;
            "
          ></pre>
        </details>
      </details>
    `;

    const possibleHosts = [
      "#right-sidebar",
      "[data-testid='right-sidebar']",
      ".right-sidebar",
      "aside"
    ];

    let host = null;

    for (
      const selector
      of possibleHosts
    ) {
      host =
        document.querySelector(
          selector
        );

      if (host) {
        host.appendChild(panel);
        break;
      }
    }

    if (!host) {
      document.body.appendChild(panel);
    }

    // Configurar eventos
    const fileInput = document.getElementById("tm-excel-file");
    const statusDiv = document.getElementById("tm-excel-status");
    const readyDot = document.getElementById("tm-excel-ready-dot");
    const resultDiv = document.getElementById("tm-excel-result");

    fileInput.addEventListener("change", async (event) => {
      const file = event.target.files[0];
      if (!file) return;

      statusDiv.hidden = false;
      statusDiv.textContent = "Procesando archivo...";
      readyDot.textContent = "⚪";

      try {
        const result = await loadAndProcessFile(file, "Sheet1");
        if (result.success) {
          statusDiv.textContent = result.message;
          readyDot.textContent = "🟢";
        } else {
          statusDiv.textContent = `Error: ${result.message}`;
          readyDot.textContent = "🔴";
        }
      } catch (error) {
        statusDiv.textContent = `Error: ${error.message}`;
        readyDot.textContent = "🔴";
      }
    });

    // Configurar eventos para botones de prueba
    document.getElementById("tm-excel-count").addEventListener("click", async () => {
      try {
        const rows = await executeSql("SELECT COUNT(*) AS count FROM excel_data;");
        resultDiv.textContent = JSON.stringify(rows, null, 2);
      } catch (error) {
        resultDiv.textContent = `Error: ${error.message}`;
      }
    });

    document.getElementById("tm-excel-preview").addEventListener("click", async () => {
      try {
        const rows = await previewRows();
        resultDiv.textContent = JSON.stringify(rows, null, 2);
      } catch (error) {
        resultDiv.textContent = `Error: ${error.message}`;
      }
    });

    document.getElementById("tm-excel-summary").addEventListener("click", async () => {
      try {
        const rows = await executeSql(`
          SELECT
            MIN("Año") AS primer_año,
            MAX("Año") AS ultimo_año,
            COUNT(DISTINCT "Año") AS años_totales,
            COUNT(*) AS total_registros
          FROM excel_data;
        `);
        resultDiv.textContent = JSON.stringify(rows, null, 2);
      } catch (error) {
        resultDiv.textContent = `Error: ${error.message}`;
      }
    });

    document.getElementById("tm-excel-year").addEventListener("click", async () => {
      try {
                const rows = await executeSql(`
          SELECT
            "Año",
            SUM("Total TEUs") AS total_teus
          FROM excel_data
          GROUP BY "Año"
          ORDER BY "Año";
        `);
        resultDiv.textContent = JSON.stringify(rows, null, 2);
      } catch (error) {
        resultDiv.textContent = `Error: ${error.message}`;
      }
    });

    document.getElementById("tm-excel-month").addEventListener("click", async () => {
      try {
        const rows = await executeSql(`
          SELECT
            "Año",
            "Mes",
            EXTRACT(MONTH FROM "Fecha") AS numero_mes,
            SUM("Total TEUs") AS total_mes
          FROM excel_data
          GROUP BY "Año", "Mes", EXTRACT(MONTH FROM "Fecha")
          ORDER BY "Año", numero_mes;
        `);
        resultDiv.textContent = JSON.stringify(rows, null, 2);
      } catch (error) {
        resultDiv.textContent = `Error: ${error.message}`;
      }
    });

    document.getElementById("tm-excel-local").addEventListener("click", async () => {
      try {
        const rows = await executeSql(`
          SELECT
            "Local",
            SUM("Total TEUs") AS total_teus
          FROM excel_data
          GROUP BY "Local"
          ORDER BY total_teus DESC;
        `);
        resultDiv.textContent = JSON.stringify(rows, null, 2);
      } catch (error) {
        resultDiv.textContent = `Error: ${error.message}`;
      }
    });

    document.getElementById("tm-excel-trans").addEventListener("click", async () => {
      try {
        const rows = await executeSql(`
          SELECT
            "Año",
            SUM("Transshipment") AS total_transshipment
          FROM excel_data
          GROUP BY "Año"
          ORDER BY "Año";
        `);
        resultDiv.textContent = JSON.stringify(rows, null, 2);
      } catch (error) {
        resultDiv.textContent = `Error: ${error.message}`;
      }
    });

    document.getElementById("tm-excel-schema").addEventListener("click", async () => {
      try {
        const schema = await getSchema();
        resultDiv.textContent = JSON.stringify(schema, null, 2);
      } catch (error) {
        resultDiv.textContent = `Error: ${error.message}`;
      }
    });

    document.getElementById("tm-excel-resolution").addEventListener("click", async () => {
      try {
        const resolution = {
          headers: currentHeaders,
          inferredTypes: currentInferredTypes
        };
        resultDiv.textContent = JSON.stringify(resolution, null, 2);
      } catch (error) {
        resultDiv.textContent = `Error: ${error.message}`;
      }
    });

    document.getElementById("tm-excel-run-sql").addEventListener("click", async () => {
      const sqlTextarea = document.getElementById("tm-excel-sql");
      const sql = sqlTextarea.value.trim();

      if (!sql) {
        resultDiv.textContent = "Por favor ingrese una consulta SQL";
        return;
      }

      try {
        const rows = await executeSql(sql);
        resultDiv.textContent = JSON.stringify(rows, null, 2);
      } catch (error) {
        resultDiv.textContent = `Error: ${error.message}`;
      }
    });

    // Configurar arrastre
    const summary = document.getElementById("tm-excel-engine-summary");
    const details = document.getElementById("tm-excel-engine-details");

    let isDragging = false;
    let startX, startY, startLeft, startTop;

    summary.addEventListener("mousedown", (e) => {
      if (e.target !== summary) return;

      isDragging = true;
      startX = e.clientX;
      startY = e.clientY;

      const rect = panel.getBoundingClientRect();
      startLeft = rect.left;
      startTop = rect.top;

      panel.style.position = "fixed";
      panel.style.zIndex = "999999";
      panel.style.left = `${startLeft}px`;
      panel.style.top = `${startTop}px`;

      e.preventDefault();
    });

    document.addEventListener("mousemove", (e) => {
      if (!isDragging) return;

      const dx = e.clientX - startX;
      const dy = e.clientY - startY;

      panel.style.left = `${startLeft + dx}px`;
      panel.style.top = `${startTop + dy}px`;
    });

    document.addEventListener("mouseup", () => {
      isDragging = false;
    });

    // Modo Micro-Widget
    const titleContent = document.getElementById("tm-excel-title-content");
    const titleText = document.getElementById("tm-excel-title-text");

    let isMicroMode = false;
    let originalWidth = panel.style.width;
    let originalHeight = panel.style.height;

    summary.addEventListener("click", (e) => {
      if (e.target !== summary) return;

      if (isMicroMode) {
        // Restaurar modo normal
        panel.style.width = originalWidth;
        panel.style.height = originalHeight;
        titleContent.style.display = "flex";
        details.open = true;
        isMicroMode = false;
      } else {
        // Cambiar a modo Micro-Widget
        originalWidth = panel.style.width;
        originalHeight = panel.style.height;
        panel.style.width = "46px";
        panel.style.height = "46px";
        titleContent.style.display = "none";
        details.open = false;
        isMicroMode = true;
      }
    });

    // Exponer API pública
    window.TMDuckDBBridge = {
      version: VERSION,
      loadAndProcessFile,
      executeSql,
      getSchema,
      previewRows
    };

    console.log(`TypingMind Excel Engine ${VERSION} inicializado`);
  }

  // Inicializar UI
  createUI();
})();



