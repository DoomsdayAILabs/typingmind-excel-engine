/*
 * TypingMind Excel Data Engine
 * v0.4.8.1 TEST
 *
 * PRUEBA MINIMA DE CARGA DE EXTENSION
 *
 * NO usa DuckDB.
 * NO carga Excel.
 * NO realiza consultas.
 *
 * Objetivo:
 * Confirmar que TypingMind ejecuta correctamente
 * el JavaScript de la Extension.
 */

(() => {
  "use strict";

  const APP_ID = "tm-excel-engine-v0481-test";

  console.log(`[${APP_ID}] INICIANDO EXTENSION`);

  function addStyles() {
    if (document.getElementById("tmxe-v0481-style")) {
      return;
    }

    const style = document.createElement("style");

    style.id = "tmxe-v0481-style";

    style.textContent = `
      #tmxe-v0481-button {
        position: fixed;
        right: 18px;
        bottom: 86px;
        z-index: 2147483000;

        border: 0;
        border-radius: 999px;

        padding: 11px 16px;

        background: #16a34a;
        color: white;

        font: 600 14px/1.1 system-ui, sans-serif;

        box-shadow: 0 6px 20px rgba(0,0,0,.25);

        cursor: pointer;
      }

      #tmxe-v0481-button:hover {
        filter: brightness(1.08);
      }

      #tmxe-v0481-overlay {
        position: fixed;
        inset: 0;

        z-index: 2147483001;

        background: rgba(0,0,0,.48);

        display: flex;
        align-items: center;
        justify-content: center;

        padding: 16px;
      }

      #tmxe-v0481-panel {
        width: min(600px, 94vw);

        background: Canvas;
        color: CanvasText;

        border: 1px solid rgba(127,127,127,.35);
        border-radius: 16px;

        padding: 22px;

        box-shadow: 0 20px 70px rgba(0,0,0,.35);

        font: 14px/1.5 system-ui, sans-serif;
      }

      #tmxe-v0481-panel h2 {
        margin: 0 0 10px;
        font-size: 21px;
      }

      #tmxe-v0481-ok {
        padding: 14px;

        border-radius: 10px;

        background: rgba(22,163,74,.12);

        border: 1px solid rgba(22,163,74,.30);

        margin: 15px 0;
      }

      #tmxe-v0481-close {
        border: 1px solid rgba(127,127,127,.45);

        border-radius: 9px;

        padding: 9px 14px;

        background: transparent;
        color: inherit;

        cursor: pointer;
      }
    `;

    document.head.appendChild(style);
  }

  function createButton() {
    if (document.getElementById("tmxe-v0481-button")) {
      console.log(
        `[${APP_ID}] BOTON YA EXISTE`
      );

      return;
    }

    const button = document.createElement("button");

    button.id = "tmxe-v0481-button";

    button.type = "button";

    button.textContent = "🟢 Excel Engine TEST";

    button.title =
      "Excel Engine v0.4.8.1 — prueba de carga";

    button.addEventListener(
      "click",
      openPanel
    );

    document.body.appendChild(button);

    console.log(
      `[${APP_ID}] BOTON CREADO CORRECTAMENTE`
    );
  }

  function openPanel() {
    if (
      document.getElementById(
        "tmxe-v0481-overlay"
      )
    ) {
      return;
    }

    const overlay =
      document.createElement("div");

    overlay.id =
      "tmxe-v0481-overlay";

    overlay.innerHTML = `
      <div
        id="tmxe-v0481-panel"
        role="dialog"
        aria-modal="true"
      >

        <h2>
          🟢 Excel Engine v0.4.8.1
        </h2>

        <div id="tmxe-v0481-ok">

          <strong>
            Extensión cargada correctamente.
          </strong>

          <br><br>

          TypingMind está ejecutando
          correctamente el JavaScript
          de la Extension.

        </div>

        <div style="
          margin-bottom:15px;
          opacity:.75;
        ">

          Esta versión es solamente una
          prueba de carga.

          <br>

          DuckDB y Excel todavía no se
          ejecutan en esta prueba.

        </div>

        <div style="
          padding:12px;
          border-radius:9px;
          background:rgba(127,127,127,.10);
          font-family:ui-monospace,
          SFMono-Regular,
          Menlo,
          monospace;
          font-size:12px;
        ">

          APP_ID:
          ${APP_ID}

          <br>

          STATUS:
          JAVASCRIPT EJECUTADO

        </div>

        <div style="
          margin-top:18px;
          display:flex;
          justify-content:flex-end;
        ">

          <button
            id="tmxe-v0481-close"
            type="button"
          >
            Cerrar
          </button>

        </div>

      </div>
    `;

    document.body.appendChild(overlay);

    document
      .getElementById(
        "tmxe-v0481-close"
      )
      .addEventListener(
        "click",
        () => overlay.remove()
      );

    overlay.addEventListener(
      "click",
      event => {
        if (event.target === overlay) {
          overlay.remove();
        }
      }
    );
  }

  function init() {
    console.log(
      `[${APP_ID}] INIT EJECUTADO`
    );

    addStyles();

    createButton();

    console.log(
      `[${APP_ID}] EXTENSION CARGADA CORRECTAMENTE`
    );
  }

  function start() {
    if (
      document.readyState === "loading"
    ) {
      document.addEventListener(
        "DOMContentLoaded",
        init,
        { once: true }
      );
    } else {
      init();
    }
  }

  start();

})();
