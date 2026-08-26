# Plugin TypingMind — Excel Engine (local)

Este plugin **no sustituye** al motor. El motor oficial sigue siendo:

`typingmind-excel-engine-v0.4.20-test.js`

TypingMind tiene dos piezas distintas:

1. **Extension** (userscript): inyecta la UI, carga el Excel y corre DuckDB-Wasm en la página.
2. **Plugin** (function calling): el LLM pide un SQL; el plugin solo envía esa consulta al motor y devuelve un JSON corto.

Los plugins corren dentro de un iframe aislado. Por eso hace falta un **puente** en la página principal.

## Archivos

| Archivo | Qué es | Dónde se instala |
|---|---|---|
| `../typingmind-excel-engine-v0.4.20-test.js` | Motor + UI | TypingMind → Settings → **Extensions** |
| `typingmind-excel-bridge.js` | Puente postMessage | TypingMind → Settings → **Extensions** (segundo script) |
| `plugin.json` | Metadatos + OpenAI Function Spec | TypingMind → **Plugins** → Create / Import |
| `implementation.js` | Código del plugin | Campo JavaScript del plugin |

## Flujo

1. El usuario carga el Excel en el panel del motor (Extension).
2. El usuario pregunta en el chat.
3. El LLM llama `query_excel_data` con un SQL.
4. El plugin (iframe) manda el SQL al puente.
5. El puente llama `window.TMExcelEngine.executeSql`.
6. El plugin devuelve al LLM un JSON corto (`procesamiento: LOCAL`, pocas filas).

El Excel **no** se envía al modelo.

## Instalación

1. Instala `typingmind-excel-engine-v0.4.20-test.js` como Extension.
2. Instala `typingmind-excel-bridge.js` como otra Extension.
3. Recarga TypingMind. Debe verse el panel del motor.
4. Crea un Plugin:
   - Implementation type: **JavaScript**
   - Output: **Give plugin output to the AI** (`respond_to_ai`)
   - Pega el contenido de `plugin.json` (spec) y `implementation.js` (código)
5. Activa el plugin en la conversación.
6. Carga un Excel en el panel del motor **antes** de preguntar.

## Límites

- El plugin recorta el resultado a 200 filas (configurable con `max_rows`, máximo 500).
- Solo se permiten consultas de lectura (`SELECT` / `WITH`).
- Si el motor no está cargado o no hay Excel, el plugin devuelve un error claro al LLM.
