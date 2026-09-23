# Plugin TypingMind — Excel Data Engine (Opción A)

El motor sigue siendo `typingmind-excel-engine-v1.0.js`, instalado como **Extensión**.
Este plugin (function calling) permite que la IA del chat consulte por sí misma los
datos cargados, escribiendo y ejecutando SQL localmente con DuckDB-WASM.

## Piezas

| Archivo | Qué es | Dónde se instala |
|---|---|---|
| `../typingmind-excel-engine-v1.0.js` | Motor + UI + puente postMessage | TypingMind → Settings → **Extensions** |
| `plugin.json` | Metadatos + especificación de la herramienta | TypingMind → **Plugins** → Create / Import |
| `implementation.js` | Código del plugin (iframe) | Campo JavaScript del plugin |

> Ya no hace falta un script de puente aparte: el listener `postMessage` está integrado
> en el propio motor (en la ventana principal).

## Flujo

1. El usuario carga el Excel en el panel del motor (Extensión).
2. El usuario hace una pregunta en lenguaje natural.
3. El LLM llama `query_excel_data` con un `sql_query` (solo SELECT/WITH).
4. El plugin (iframe) envía la consulta por `window.parent.postMessage`.
5. El motor ejecuta `executeQuery` en DuckDB y responde las filas.
6. El plugin devuelve a la IA un Markdown tabular (máx. 50 filas).

El Excel **no** se envía al modelo.

## Instalación

1. Instala `typingmind-excel-engine-v1.0.js` como Extension.
2. Recarga TypingMind; debe verse el panel del motor.
3. Crea un Plugin:
   - Implementation type: **JavaScript**
   - Output: **Give plugin output to the AI** (`respond_to_ai`)
   - Pega el contenido de `plugin.json` (spec) y `implementation.js` (código).
4. Activa el plugin en la conversación.
5. Carga un Excel en el panel del motor **antes** de preguntar.

## Límites

- Solo lectura (`SELECT` / `WITH` / `EXPLAIN`) contra la tabla `excel_data`.
- El plugin recorta el resultado a 50 filas.
- Si el motor no está cargado o no hay Excel, el plugin devuelve un error claro al LLM.
