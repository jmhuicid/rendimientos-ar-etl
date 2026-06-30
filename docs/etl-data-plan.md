# Rendimientos AR ETL - estado y plan de datos

Fecha de referencia: 2026-06-23.

## Objetivo propio

Este repositorio se usa como ETL/reporteria propio. No se empuja nada al repositorio original `arisbdar/rendimientos-ar`.

Quedan fuera del alcance de este ETL:

- Mundial
- Prode
- features sociales/foro/chat si no aportan datos financieros

## Prioridad de datos

1. Datos actuales para reporteria diaria.
2. Snapshots historicos propios, renovados todos los dias.
3. Historicos externos cuando existan y sean estables.

## Fuentes validadas

| Fuente | Uso | Estado |
| --- | --- | --- |
| ArgentinaDatos FCI | VCP, patrimonio y rendimiento diario estimado por categoria | Funciona. El 2026-06-23 devolvio datos al 2026-06-22. |
| ArgentinaDatos plazo fijo | Tasas bancarias actuales | Integrado al snapshot. |
| BCRA Estadisticas v4.0 | CER, UVA, tasas, dolares oficiales, inflacion, agregados | Integrado al snapshot. |
| BCRA REM | Expectativas macro, inflacion, tipo de cambio, TAMAR, actividad, fiscal, sector externo | Debe reportarse como fuente mensual oficial con PDF/XLSX. |
| Ministerio de Economia - Portal de Datos Economicos | Series macro oficiales | Debe mantenerse como fuente oficial de contraste y ampliacion. |
| Ministerio de Economia - Programacion Macroeconomica | Informes macro, actividad, precios, balanza de pagos, sector monetario/cambiario/financiero | Debe reportarse como fuente oficial. |
| Ministerio de Economia - Gasto Publico Consolidado | Informe y bases descargables de gasto publico | Debe reportarse cuando aplique para lectura fiscal/anual. |
| data912 | Notes, bonds y corporativos vivos | Integrado al snapshot. |
| Pellegrini FCI | Valores publicos de cuotaparte y pagina propia | Integrado como scrape defensivo; conservar raw excerpt. |
| Fonditos | Buena referencia de cobertura FCI y freshness | API de datos requiere autenticacion comercial; health publico OK. |

## Fonditos

El sitio declara rankings, comparador, TNA/TIR y datos oficiales diarios de CAFCI. Su health publico reporto:

- `last_sync`: 2026-06-22
- `funds_count`: 2293
- `is_complete`: true

Rutas detectadas en el bundle:

- `/funds`
- `/funds/returns`
- `/funds/history`
- `/funds/compare`

Pero las rutas de datos responden `Unauthorized` sin credenciales. Para uso propio, no conviene depender de scraping de esa API. Si se necesita esa cobertura completa, evaluar acceso comercial o usarlo solo como benchmark de freshness/cobertura.

## Corrida local

```bash
npm run etl:snapshot
npm run report:fci-mm
npm run report:argentina
npm run report:tamar-badlar-uvi
```

Salida:

```text
data/snapshots/YYYY-MM-DD/
  fci.json
  plazo_fijo.json
  bcra.json
  mercado.json
  pellegrini.json
  manifest.json
data/snapshots/latest.json
data/reports/YYYY-MM-DD/
  fci-money-market-benchmark.json
  fci-money-market-benchmark.csv
  fci-money-market-benchmark.md
  argentina-daily-report.json
  argentina-daily-report.md
  argentina-daily-key-metrics.csv
  analisis_tamar_badlar_uvi_titulos_YYYYMMDD.xlsx
  analisis_tamar_badlar_uvi_titulos_YYYYMMDD.md
```

El reporte `report:tamar-badlar-uvi` requiere Python con `pandas` y `openpyxl`:

```bash
pip install -r requirements.txt
```

Por defecto lee `data/input/series_historicas_YYYYMMDD.xlsx`, descargado desde series historicas BYMA, y cruza ese historico con BCRA Estadisticas v4.0 y data912.

## Benchmark Money Market

`npm run report:fci-mm` responde preguntas tipo:

> tenes una tabla mensual de rendimientos de fondos money market en Argentina, benchmark o Mercado Fondo?

La salida usa `mercadoDinero/ultimo` y `mercadoDinero/penultimo` de ArgentinaDatos. Con esos dos VCP calcula:

- rendimiento diario observado
- TNA aproximada
- TEM30 aproximada
- TEA aproximada

Esto no reemplaza una serie mensual historica cerrada de CAFCI. Para eso hay dos caminos: acumular snapshots diarios propios o contratar/autorizar una API con historico completo como Fonditos.

## Reporte diario Argentina

`npm run report:argentina` genera un informe de apertura profesional, pensado para lectura al inicio de la jornada laboral, con:

- dolar oficial, mayorista, blue, MEP, CCL y cripto
- brecha contra oficial
- riesgo pais
- reservas internacionales
- BADLAR, TAMAR y tasa de depositos a 30 dias
- inflacion mensual/interanual, CER, UVA e ICL
- top plazo fijo por TNA
- Mercado Fondo y benchmark FCI Money Market
- precios testigo de soberanos AL/GD
- analisis de series TAMAR/BADLAR/UVI-UVA, convirtiendo BADLAR/TAMAR de TNA a TEA y UVI/UVA desde indice a tasa efectiva anualizada por ventanas 30/90/365 dias
- fuentes oficiales y reportes a monitorear: BCRA, REM, Ministerio de Economia, Portal de Datos Economicos, Gasto Publico Consolidado

Cada metrica conserva fecha, fuente, brechas y variaciones cuando estan disponibles. La salida principal para lectura humana es `argentina-daily-report.md`; para BI o automatizacion usar `argentina-daily-report.json` y `argentina-daily-key-metrics.csv`.

Las fuentes oficiales no deben ser reemplazadas por agregadores cuando exista dato primario disponible. Los agregadores se usan como capa operativa cuando aportan frecuencia, historico o normalizacion, pero el informe debe mostrar las fuentes oficiales relevantes y sus reportes/Excel asociados.

## Siguiente mejora sugerida

- Agregar job diario (GitHub Actions o scheduler local) que ejecute `npm run etl:snapshot`.
- Versionar snapshots si el volumen es aceptable; si crece demasiado, mover historico a SQLite/Parquet/CSV comprimido.
- Crear reportes derivados en `data/reports/YYYY-MM-DD/`: top FCIs por TNA/TEA, money market vs plazo fijo, CER/UVA, curva soberanos/ONs.
- Profundizar el reporte TAMAR/BADLAR/UVI-UVA con spreads reales, ranking de letras/titulos por cobertura implicita y alertas cuando UVI anualizado supere tasas de mercado.
