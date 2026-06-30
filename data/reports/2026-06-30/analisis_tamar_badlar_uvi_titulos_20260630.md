# Analisis comparativo TAMAR vs BADLAR vs UVI/UVA

Fecha de generacion: 2026-06-30

## Lectura rapida

- **Corte / Fecha de generacion**: 2026-06-30. Reporte generado con datos disponibles al momento de corrida.
- **BYMA historico / Archivo fuente**: series_historicas_20260630.xlsx. Rango detectado 2021-07-01 a 2025-06-27. Si el archivo descargado hoy trae maximo 2025-06-27, el historico BYMA no esta actualizado a 2026-06-30.
- **BCRA / Series oficiales**: BADLAR, TAMAR, UVI_UVA. BADLAR y TAMAR se convierten de TNA a TEA. UVI/UVA se convierte de indice a TEA anualizada por ventanas moviles.
- **Tasas comparables / UVI/UVA TEA 30D**: 33.54002735823332. Tasa efectiva anualizada de la variacion observada del indice en los ultimos 30 dias calendario disponibles.
- **Tasas comparables / BADLAR TEA / TAMAR TEA**: 22.44% / 25.30%. Conversion de TNA a TEA con capitalizacion diaria para comparar en la misma escala que UVI/UVA anualizado.
- **Mercado vivo / Instrumentos data912 seleccionados**: 57. Snapshot vivo operativo para Letras y Titulos Publicos tomado de la misma fuente usada en rendimientos-ar-etl.
- **BCRA ultimo / BADLAR privados**: 20.25. % TNA al 2026-06-30; var. 30d -1.82%, var. 1y -36.72%.
- **BCRA ultimo / TAMAR privados**: 22.5625. % TNA al 2026-06-30; var. 30d 1.40%, var. 1y -32.90%.
- **BCRA ultimo / UVI/UVA**: 2016.85. indice al 2026-06-30; var. 30d 2.41%, var. 1y 32.88%.

## Que mide cada variable

- **BADLAR**: tasa promedio pagada por bancos privados por depositos a plazo fijo mayoristas. Desde el lado inversor, aproxima una referencia de tasa pasiva bancaria para pesos.
- **TAMAR**: tasa promedio para depositos a plazo fijo mayoristas de mayor monto. Suele mirar un segmento institucional y puede quedar por encima o por debajo de BADLAR segun liquidez bancaria.
- **UVI/UVA**: indice que ajusta por inflacion. Por eso su TEA anualizada suele verse mas estable que una tasa de mercado: no es una tasa ofrecida, sino la capitalizacion efectiva de la inflacion observada en el periodo.

Desde el inversor, la pregunta no es solo que serie sube o baja, sino si la tasa nominal compensa el ritmo del indice inflacionario. Cuando UVI/UVA anualizado supera BADLAR/TAMAR TEA, el deposito remunerado queda corriendo de atras a la inflacion del periodo observado.

## Ultimos datos BCRA

| serie | nombre | unidad | ultima_fecha | ultimo_valor | var_abs_dato_previo | var_pct_30d | var_pct_1y |
| --- | --- | --- | --- | --- | --- | --- | --- |
| BADLAR | BADLAR privados | % TNA | 2026-06-30 | 20.25 | 0.0 | -1.8181818181818188 | -36.71875 |
| TAMAR | TAMAR privados | % TNA | 2026-06-30 | 22.5625 | 0.0 | 1.404494382022481 | -32.899628252788105 |
| UVI_UVA | UVI/UVA | indice | 2026-06-30 | 2016.85 | 4.179999999999836 | 2.4057233672003075 | 32.8771996863944 |

## Tasas comparables TEA

| FECHA | BADLAR_TEA | TAMAR_TEA | UVI_TEA_30D | UVI_TEA_90D | UVI_TEA_365D |
| --- | --- | --- | --- | --- | --- |
| 2026-06-30 00:00:00 | 22.439133095218434 | 25.30183160949988 | 33.54002735823332 | 39.81167342658172 | 32.8771996863944 |

Metodologia: BADLAR/TAMAR se transforman desde TNA a TEA con capitalizacion diaria. UVI/UVA se transforma desde indice a tasa efectiva anualizada: `((indice_actual / indice_previo) ** (365 / dias) - 1) * 100`.

## Ejemplo inversor

Supuesto: capital inicial de $100.000.000, reinvirtiendo tasas diariamente con TNA/365 y ajustando UVI/UVA por variacion del indice.
Escenario 1: desde 2023 se compara BADLAR contra UVI/UVA, porque TAMAR no tiene dato disponible desde esa fecha.
Escenario 2: desde el inicio disponible de TAMAR se reinician los $100.000.000 y se comparan BADLAR, TAMAR y UVI/UVA en las mismas fechas.
Escenario 3: desde 2026-01-01 se vuelve a reiniciar el capital para mirar el ano corriente con las tres variables en paralelo.

| escenario | variable | metodo | fecha_inicio | fecha_final | capital_inicial | capital_final | ganancia_nominal | multiplicador | rendimiento_acumulado_pct | tea_equivalente_periodo_pct |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Desde 2023: BADLAR vs UVI/UVA | BADLAR privados | TNA reinvertida diaria | 2023-01-01 | 2026-06-30 | $100.000.000 | $728.296.326 | $628.296.326 | 7.28x | 628.30% | 76.47% |
| Desde 2023: BADLAR vs UVI/UVA | UVI/UVA | Ajuste por indice | 2023-01-01 | 2026-06-30 | $100.000.000 | $1.088.306.713 | $988.306.713 | 10.88x | 988.31% | 97.95% |
| Desde inicio TAMAR: mismas fechas | BADLAR privados | TNA reinvertida diaria | 2024-10-01 | 2026-06-30 | $100.000.000 | $176.853.661 | $76.853.661 | 1.77x | 76.85% | 38.64% |
| Desde inicio TAMAR: mismas fechas | TAMAR privados | TNA reinvertida diaria | 2024-10-01 | 2026-06-30 | $100.000.000 | $183.895.721 | $83.895.721 | 1.84x | 83.90% | 41.77% |
| Desde inicio TAMAR: mismas fechas | UVI/UVA | Ajuste por indice | 2024-10-01 | 2026-06-30 | $100.000.000 | $170.166.720 | $70.166.720 | 1.70x | 70.17% | 35.61% |
| Desde 2026: mismas fechas | BADLAR privados | TNA reinvertida diaria | 2026-01-01 | 2026-06-30 | $100.000.000 | $113.528.512 | $13.528.512 | 1.14x | 13.53% | 29.34% |
| Desde 2026: mismas fechas | TAMAR privados | TNA reinvertida diaria | 2026-01-01 | 2026-06-30 | $100.000.000 | $114.495.707 | $14.495.707 | 1.14x | 14.50% | 31.59% |
| Desde 2026: mismas fechas | UVI/UVA | Ajuste por indice | 2026-01-01 | 2026-06-30 | $100.000.000 | $118.003.101 | $18.003.101 | 1.18x | 18.00% | 39.89% |

## Escenario futuro a diciembre 2026

Para proyectar hacia adelante ya no usamos dato observado puro: se arma un estimativo. La idea es suponer un proceso de desinflacion donde la inflacion mensual converge gradualmente hacia 1,0% en diciembre 2026, y tomar como ancla de mercado la TAMAR esperada por REM para diciembre 2026 en torno a 22,1% TNA.

Si tasas e inflacion bajan juntas, la lectura del inversor cambia: ya no alcanza con mirar quien gana en el acumulado historico, sino si la tasa futura queda por encima o por debajo del sendero de inflacion esperado. En un escenario de inflacion bajando a 1% mensual, una TAMAR/BADLAR que no baje demasiado rapido puede volver a verse competitiva en terminos reales de corto plazo.

| escenario | variable | metodo | capital_inicial | capital_final_estimado | ganancia_nominal_estimada | multiplicador_estimado | rendimiento_estimado_pct | fecha_inicio | fecha_final | inflacion_mensual_dic_2026_supuesto | tamar_dic_2026_tna_rem |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Jul-Dic 2026 estimado | BADLAR privados | TNA proyectada, reinversion diaria | $100.000.000 | $110.373.020 | $10.373.020 | 1.10x | 10.37% | 2026-07-01 | 2026-12-31 | 1.00% | 22.10% |
| Jul-Dic 2026 estimado | TAMAR privados | REM mayo 2026: converge a 22,1% TNA en diciembre | $100.000.000 | $111.638.193 | $11.638.193 | 1.12x | 11.64% | 2026-07-01 | 2026-12-31 | 1.00% | 22.10% |
| Jul-Dic 2026 estimado | UVI/UVA | Inflacion mensual converge a 1,0% en diciembre | $100.000.000 | $110.654.734 | $10.654.734 | 1.11x | 10.65% | 2026-07-01 | 2026-12-31 | 1.00% | 22.10% |

Fuente/criterio: REM BCRA publicado en junio 2026 informa TAMAR proyectada de 22,1% TNA para diciembre 2026; el sendero de inflacion a 1,0% mensual en diciembre es un supuesto de desinflacion para sensibilizar el ejemplo.

## Correlacion de tendencias desde 2023

| serie_a | serie_b | correlacion_nivel | correlacion_cambio_30d | lectura |
| --- | --- | --- | --- | --- |
| BADLAR_TEA | TAMAR_TEA | 0.99 | 0.99 | Tendencias muy alineadas. |
| BADLAR_TEA | UVI_TEA_30D | 0.50 | 0.32 | Tendencias moderadamente alineadas. |
| BADLAR_TEA | UVI_TEA_90D | 0.39 | -0.07 | Tendencias moderadamente alineadas; los cambios de corto plazo no necesariamente se mueven juntos. |
| BADLAR_TEA | UVI_TEA_365D | 0.14 | -0.25 | Relacion debil en niveles; los cambios de corto plazo no necesariamente se mueven juntos. |
| TAMAR_TEA | UVI_TEA_30D | -0.28 | -0.01 | Relacion debil en niveles; los cambios de corto plazo no necesariamente se mueven juntos. |
| TAMAR_TEA | UVI_TEA_90D | -0.25 | -0.27 | Relacion debil en niveles. |
| TAMAR_TEA | UVI_TEA_365D | 0.13 | 0.02 | Relacion debil en niveles; los cambios de corto plazo no necesariamente se mueven juntos. |
| UVI_TEA_30D | UVI_TEA_90D | 0.79 | 0.26 | Tendencias muy alineadas. |
| UVI_TEA_30D | UVI_TEA_365D | 0.42 | 0.17 | Tendencias moderadamente alineadas; los cambios de corto plazo no necesariamente se mueven juntos. |
| UVI_TEA_365D | UVI_TEA_90D | 0.62 | 0.39 | Tendencias moderadamente alineadas. |

Lectura: la correlacion en niveles muestra si las curvas tienden a moverse en la misma direccion durante el periodo. La correlacion de cambios a 30 dias es mas exigente: pregunta si los giros de corto plazo tambien coinciden.

## Instrumentos publicos

- BYMA historico filtrado: 84 series de Letras/Titulos detectadas.
- data912 vivo: 201 instrumentos totales; 57 coinciden con la seleccion de rendimientos-ar-etl.

Nota: el pedido menciona UVI; la API usada por rendimientos-ar-etl expone el ID 31 como UVA. Se etiqueta UVI/UVA para mantener trazabilidad del pedido y de la fuente.
