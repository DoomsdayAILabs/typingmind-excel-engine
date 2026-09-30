# Product Context — TypingMind Excel Engine

> Por qué existe el producto, qué problema resuelve y cómo se espera que se comporte.
> Derivado de `README.md` (líneas 16‑88, 302‑351, 355‑399) y de la UI real del widget.

## Por qué existe

TypingMind es un cliente de chat: su contexto es texto. Las hojas de cálculo chocan con tres muros:

| Problema | Consecuencia sin el motor |
|---|---|
| **Privacidad** | Hojas con datos sensibles no se pueden pegar en un chat externo. |
| **Tamaño** | Decenas de miles de filas no caben en el contexto de un LLM. |
| **Precisión** | El modelo "adivina" sumas y promedios si no puede calcular. |

La solución: dar al navegador un motor SQL real (DuckDB‑WASM) y dar al modelo una **herramienta**
para consultarlo. El archivo no viaja; solo viaja el resultado de la consulta.

## Cómo debe funcionar (experiencia objetivo)

### Flujo manual (sin IA)
1. El usuario arrastra el archivo al widget flotante (o pulsa *Seleccionar Archivo*).
2. El motor responde con tabla, registros, esquema y preview de 10 filas.
3. El usuario escribe SQL en la consola del widget y pulsa **▶ Ejecutar Consulta**.
4. Ve hasta **100 filas** en pantalla; puede **⬇️ CSV** (descarga `tmee_resultados.csv`) o
   **💬 Enviar a TM** (pega la tabla Markdown en el input del chat, máx. 50 filas).

### Flujo IA autónoma (plugin)
1. El usuario carga el archivo y activa el plugin en la conversación.
2. Pregunta en lenguaje natural.
3. El LLM llama `query_excel_data({ sql_query })`.
4. El iframe del plugin reenvía la consulta al motor por `postMessage`.
5. El motor devuelve filas normalizadas; el plugin las convierte en Markdown (máx. 50 filas).
6. El LLM redacta la respuesta con el dato exacto.

## Objetivos de UX (v1.0)

- **Widget flotante no intrusivo**, abajo a la izquierda, arrastrable desde la cabecera.
- **Modo icono flotante de 50×50 px** al minimizar (pensado para móvil: FAB táctil con el emoji 📊).
- **Un solo gesto** en la cabecera: clic/tap = minimizar/maximizar, arrastrar = mover.
  El umbral que separa clic de arrastre es de **5 px**; el widget nunca se sale de la pantalla.
- **Retroalimentación constante** con un badge de estado: `Iniciando… / Motor Listo / Cargando… /
  Ejecutando SQL… / Consulta OK / Consulta IA OK / Enviado al chat` + variantes de error.
- **Inyección en el chat sin romper React:** se escribe en el `textarea` del chat y se dispara un
  evento `input` nativo para que el estado interno de TypingMind se sincronice.
- **Fallback digno:** si no se encuentra el input del chat, la tabla se copia al portapapeles.

## Principios de diseño asumidos por el producto

- **El LLM no ve el libro:** solo ve tablas dentro de DuckDB, por eso el spec del plugin le obliga a
  descubrir el esquema antes de consultar (ver `plugin-bridge-protocol.md`).
- **Respuestas acotadas a propósito:** 100 filas en pantalla, 50 al inyectar, 50 en el plugin —
  el límite es una decisión de producto (no saturar el contexto), no una limitación técnica.
- **Degradación tolerante en Excel:** una pestaña vacía o corrupta **no** aborta la carga del libro.
- **Errores legibles para el modelo:** el plugin devuelve texto ("timeout", "Error SQL del motor
  local: …") para que el propio LLM pueda explicar o reintentar.

## Solución de problemas que el producto ya anticipa

| Síntoma | Causa esperada | Manejo previsto |
|---|---|---|
| No aparece el widget | Extensión no guardada/recargada, bloqueador de scripts | Recargar TypingMind, revisar consola |
| "Fallo al iniciar: Failed to fetch" | `WORKER_PATH` inaccesible | Verificar URL de GitHub Pages / constante (línea 16) |
| "Error en Worker" | CDN bloqueado (red corporativa) | Permitir `cdn.jsdelivr.net` o autoalojar bundles |
| La IA dice que el motor no responde | Extensión ausente, plugin desactivado o sin archivo cargado | Cargar archivo + activar plugin (timeout del plugin: 30 s) |
| El modelo intenta escribir datos | Guardrail de solo lectura | Pedirle una consulta de lectura |

## Por confirmar

- No hay datos de uso reales ni métricas en el repo (no existe analítica ni telemetría).
- Público objetivo concreto (usuario individual vs. equipo) — no documentado en el repo.
