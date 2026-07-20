# Monitor de series BCRA

El workflow `BCRA series monitor` se ejecuta diariamente a las 18:00 de Argentina y controla las series configuradas en `config/bcra-series-monitor.json` contra la API monetaria oficial `v4.0`.

## Controles

- vigencia y formato del endpoint
- catálogo completo con paginación
- existencia de cada identificador
- fecha máxima disponible
- observaciones nuevas frente a la última corrida exitosa
- duplicados por fecha
- valores nulos o no numéricos
- revisiones de valores ya publicados
- retrocesos de la fecha máxima
- periodicidad informada por el catálogo oficial

Cada ejecución genera un resumen visible en GitHub Actions y artefactos JSON/Markdown con retención de 90 días. El estado comparable se recupera del último artefacto exitoso.

## Estados

- `OK`: no se detectaron anomalías.
- `WARNING`: se detectaron duplicados, valores inválidos o revisiones históricas.
- `ERROR`: falló una consulta, desapareció una serie o retrocedió la fecha máxima.

El indicador `safeToRefresh` solo es verdadero cuando el estado general es `OK`.

## Compatibilidad v4.0

Los identificadores históricos `128`, `129` y `131` no existen en el catálogo `v4.0`. Para que el control sea operativo y no falle todos los días por series retiradas, la configuración monitorea las series públicas vigentes que alimentan `bcra_public_series.xlsx`:

- `1189`: tasa de depósitos a plazo fijo en pesos.
- `1190`: tasa de depósitos a plazo fijo en pesos de personas humanas.
- `1192`: tasa de depósitos a plazo fijo en pesos de otras personas jurídicas.

Las funciones productivas `cer.js` y `cer-ultimo.js` ya consumen correctamente la serie `30` en `v4.0` y no son modificadas por este monitor.
