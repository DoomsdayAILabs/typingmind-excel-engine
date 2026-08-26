# TypingMind Excel Engine — Instrucciones para Codex

## Proyecto

Este repositorio contiene el proyecto TypingMind Excel Engine.

Su objetivo es permitir que TypingMind analice archivos Excel grandes mediante procesamiento LOCAL en el navegador, evitando enviar el contenido completo del Excel al modelo de IA.

Arquitectura principal:

Excel
→ JavaScript
→ DuckDB-Wasm
→ tabla `excel_data`
→ SQL
→ resultado pequeño
→ TypingMind / LLM

El procesamiento de los datos debe permanecer LOCAL.

---

## Versión oficial actual

La versión oficial y última versión validada actualmente es:

`typingmind-excel-engine-v0.4.19-test.js`

Esta versión es el BASELINE OFICIAL del proyecto.

### Regla fundamental

NO modificar directamente:

`typingmind-excel-engine-v0.4.19-test.js`

Las nuevas versiones deben partir de esta versión.

La siguiente versión, si se requiere una modificación, será:

`typingmind-excel-engine-v0.4.20-test.js`

No utilizar versiones anteriores como base de nuevas modificaciones.

---

## Reglas de desarrollo

1. No eliminar funcionalidades que ya funcionan.

2. No modificar código basándose únicamente en una sospecha cuando sea posible crear primero una prueba que permita aislar el problema.

3. Antes de modificar el motor:
   - formular la hipótesis;
   - diseñar una prueba;
   - ejecutar la prueba;
   - analizar el resultado;
   - identificar la capa donde ocurre el problema.

4. Cuando exista una versión funcional anterior, conservarla intacta como referencia.

5. Las modificaciones deben ser pequeñas y controladas.

6. Después de una modificación, ejecutar pruebas de regresión para comprobar que las funciones existentes continúan funcionando.

---

## Procesamiento LOCAL

El archivo Excel debe permanecer local en el navegador.

El motor debe realizar localmente:

- lectura del Excel;
- preparación de datos;
- creación de `excel_data`;
- ejecución SQL;
- agregaciones;
- filtros;
- joins;
- cálculos;
- conversiones;
- generación del resultado.

No enviar el contenido completo del Excel al LLM.

El LLM debe recibir solamente el resultado necesario de las consultas.

---

## Tabla principal

La tabla SQL principal utilizada por el motor es:

`excel_data`

Columnas principales conocidas:

- `Fecha`
- `Año`
- `Mes`
- `Fecha Combinada`
- `Local`
- `Transshipment`
- `Total TEUs`

La columna correcta actualmente es:

`Total TEUs`

No asumir que el nombre actual es `Total TEU's`.

---

## Estado actual

La mayoría de las pruebas SQL principales ya fueron validadas correctamente.

Entre ellas:

- SELECT
- WHERE
- ORDER BY
- GROUP BY
- HAVING
- DISTINCT
- IN
- BETWEEN
- SUM
- AVG
- MIN
- MAX
- COUNT
- JOIN
- CTE
- subconsultas
- funciones de ventana
- CASE
- COALESCE
- NULL
- NULLIF
- TRY_CAST
- CAST
- EXTRACT
- cálculos porcentuales
- FILTER con COUNT

Existe actualmente una anomalía pendiente relacionada con:

`SUM(...) FILTER (...)`

Esta anomalía todavía NO debe considerarse un error de DuckDB.

Primero debe determinarse si el valor incorrecto aparece en:

DuckDB
→ Arrow
→ JavaScript
→ serialización JSON

Antes de modificar el código.

---

## Diagnóstico de problemas

Cuando se encuentre un problema:

1. No modificar inmediatamente el código.
2. Intentar reproducir el problema.
3. Crear una prueba mínima.
4. Determinar en qué capa aparece el comportamiento incorrecto.
5. Documentar el resultado.
6. Solo después proponer una modificación.

---

## Regla de conservación

Las funcionalidades que ya fueron comprobadas como correctas deben considerarse protegidas.

Una nueva modificación no debe asumir que una funcionalidad existente puede romperse para solucionar otra.

Toda corrección debe comprobar posibles regresiones.

---

## Versionado

Baseline actual:

`v0.4.19-test`

Siguiente versión posible:

`v0.4.20-test`

No sobrescribir el baseline.

Cuando se cree una nueva versión, conservar el código completo y funcional.

---

## Estilo de trabajo

El propietario del proyecto no es programador profesional.

Las instrucciones y explicaciones deben ser claras, concretas y paso a paso.

Evitar cambios grandes cuando pueda realizarse una modificación pequeña y verificable.

Cuando una tarea pueda resolverse mediante una prueba antes de modificar código, realizar primero la prueba.
