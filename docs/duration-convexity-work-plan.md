# Plan de trabajo: Duration, convexidad y escenarios de tasas

Fecha de referencia: 2026-08-20.

## Objetivo

Incorporar a `rendimientos-ar-etl` una capa diaria de sensibilidad para letras y bonos que permita estimar el efecto de subas y bajas de tasas mediante una expansion de Taylor de segundo orden, comparar esa aproximacion con una revaluacion exacta y traducir el movimiento a resultado monetario por posicion.

El desarrollo debe servir tanto para posiciones compradas como para ventas en corto cubiertas con prestamo de valores. Complementa, pero no reemplaza, el analisis de liquidez, spread, volumen, profundidad y costo del prestamo.

## Modelo analitico

La variacion porcentual aproximada del precio se calculara como:

```text
Delta P / P ~= -DurationModificada * Delta y
               + 0.5 * Convexidad * (Delta y)^2
```

Donde `Delta y` se expresa en decimal: 1 punto basico es `0.0001` y 100 puntos basicos son `0.01`.

Para cada instrumento se calcularan, cuando corresponda:

- TIR o rendimiento de referencia.
- Duration Macaulay.
- Duration modificada.
- Convexidad.
- DV01 por VN 100, por VN 1.000.000 y por posicion informada.
- Cambio de precio de primer orden.
- Cambio de precio de segundo orden.
- Precio obtenido por revaluacion exacta.
- Error absoluto y relativo de Taylor frente a la revaluacion exacta.
- P&L bruto de una posicion comprada o vendida.

Los escenarios iniciales seran `-500`, `-200`, `-100`, `-50`, `-25`, `+25`, `+50`, `+100`, `+200` y `+500` puntos basicos. La configuracion debe permitir agregar escenarios sin modificar el motor.

## Tratamiento por familia

### Tasa fija, LECAP y BONCAP

- Construir los flujos contractuales completos, incluyendo capitalizacion y amortizaciones.
- Resolver la TIR consistente con precio, plazo de liquidacion y convencion del instrumento.
- Aplicar duration y convexidad sobre la tasa nominal correspondiente.
- Para instrumentos de pago unico, conservar la formula general basada en flujos para evitar una implementacion paralela.

### CER y familia TZX

- Calcular flujos reales y valuarlos con una TIR real.
- Aplicar los shocks en puntos basicos sobre la curva o TIR real, no sobre la inflacion.
- Separar el efecto tasa real del efecto indexacion CER.
- Incluir escenarios de CER/inflacion para convertir precios reales a pesos nominales cuando la informacion contractual lo requiera.
- Informar claramente fecha y rezago del CER utilizado.

### TAMAR

- Proyectar los cupones con una curva o sendero TAMAR configurable y el margen contractual.
- Calcular duration y convexidad efectivas mediante revaluaciones con shocks paralelos positivos y negativos.
- Mantener separado el shock de curva de la actualizacion de los cupones futuros.
- Documentar el supuesto de tasa futura utilizado en cada corrida.

### Posiciones vendidas con prestamo de valores

- Invertir el signo del P&L respecto de una posicion comprada.
- Incorporar, cuando se informe, comisiones, derechos de mercado y costo del prestamo.
- Mostrar movimiento minimo de precio necesario para cubrir los costos de ida y vuelta.
- No clasificar una especie como operable solamente por volatilidad: exigir metricas de liquidez y spread.

## Datos requeridos

La fuente normalizada por instrumento debe contener como minimo:

| Campo | Descripcion |
| --- | --- |
| `ticker` | Simbolo negociado. |
| `familia` | Fija, LECAP, BONCAP, CER/TZX o TAMAR. |
| `fecha_valuacion` | Fecha y hora de corte. |
| `fecha_liquidacion` | Fecha efectiva de liquidacion. |
| `precio` | Precio limpio o sucio, identificado explicitamente. |
| `moneda` | Moneda de cotizacion y pago. |
| `valor_residual` | Valor nominal pendiente. |
| `flujos` | Fechas, intereses, amortizaciones y ajustes. |
| `indice_base` | CER u otro indice contractual, cuando aplique. |
| `margen` | Spread contractual para tasa flotante. |
| `bid_ask` | Mejores puntas y cantidades. |
| `volumen_operaciones` | Volumen nominal, monto y cantidad de operaciones. |

Cada dato debe conservar fuente, fecha de observacion y estado de calidad. Una especie con flujos incompletos no debe recibir metricas definitivas.

## Entregables diarios

La corrida diaria agregara:

```text
data/reports/YYYY-MM-DD/
  duration-convexity-universe.json
  duration-convexity-universe.csv
  sensibilidad_duration_convexity_YYYYMMDD.xlsx
  duration-convexity-summary.md
```

El Excel contendra como minimo estas hojas:

- `Resumen`: precio, TIR, duration, convexidad, DV01 y alertas de calidad.
- `Escenarios_pb`: resultados para todos los shocks configurados.
- `Revaluacion_exacta`: comparacion entre Taylor de primer orden, segundo orden y precio exacto.
- `Posiciones`: P&L por nominal, lado comprado/vendido y costos informados.
- `Liquidez`: spread, volumen, operaciones y profundidad disponible.
- `Supuestos`: fecha de corte, CER, curva TAMAR, convenciones y fuentes.

El reporte Markdown destacara:

- Mayor DV01 y mayor convexidad por familia.
- Instrumentos con mayor error de aproximacion ante shocks grandes.
- Sensibilidad asimetrica ante bajas y subas de tasas.
- Candidatos para posiciones compradas y ventas cortas, condicionados por liquidez.
- Alertas por datos vencidos, flujos incompletos o mercados sin operaciones.

## Fases de implementacion

### Fase 1: Contrato de datos y motor comun

- Definir el esquema normalizado de instrumentos, flujos y posiciones.
- Implementar precio, TIR, duration, convexidad, DV01 y revaluacion exacta.
- Incorporar escenarios configurables en puntos basicos.
- Agregar pruebas unitarias con bonos cero cupon y bonos con cupon conocidos.

Criterio de salida: las formulas reproducen casos analiticos conocidos dentro de tolerancias documentadas.

### Fase 2: Tasa fija, LECAP y BONCAP

- Adaptar los instrumentos ya presentes en `public/config.json`.
- Integrar precios vigentes y convencion de liquidacion.
- Generar JSON, CSV, Excel y resumen Markdown.

Criterio de salida: todas las especies con datos completos muestran TIR, duration, convexidad, DV01 y escenarios exactos/aproximados.

### Fase 3: CER y familia TZX

- Normalizar flujos CER y coeficientes de indexacion.
- Calcular sensibilidad a tasa real y escenarios separados de inflacion/CER.
- Validar al menos TZXO6, TZXD6, TZXM7 y TZX28, sujetos a disponibilidad vigente.

Criterio de salida: el reporte distingue sin ambiguedad efecto CER, efecto tasa real y resultado nominal.

### Fase 4: TAMAR

- Incorporar curva o sendero de proyeccion TAMAR.
- Implementar duration y convexidad efectivas mediante shocks y revaluacion.
- Registrar los supuestos de cupones proyectados en la salida diaria.

Criterio de salida: los shocks simetricos pueden reproducirse y cada resultado identifica la curva base utilizada.

### Fase 5: Posiciones y operacion con prestamo

- Admitir nominal, lado, precio de entrada, comisiones y costo del prestamo.
- Calcular P&L comprado/vendido y punto de equilibrio de recompra.
- Combinar sensibilidad con spread, volumen y operaciones de mercado.

Criterio de salida: una venta corta muestra correctamente beneficio ante baja del precio, perdida ante suba y costos totales separados.

### Fase 6: Integracion diaria y observabilidad

- Incorporar el reporte al comando diario y al workflow manual/automatizado.
- Publicar artifacts con las restantes salidas de `rendimientos-ar-etl`.
- Agregar controles de frescura, cobertura y errores de valuacion al resumen de la corrida.

Criterio de salida: una corrida completa produce resultados reproducibles y falla de forma explicita cuando faltan datos esenciales.

## Validacion y tolerancias

- Comparar Taylor de primer y segundo orden contra revaluacion exacta en todos los escenarios.
- Verificar que Taylor de segundo orden no empeore materialmente el error para shocks pequenos en instrumentos convexos estandar.
- Probar signos para posiciones compradas y vendidas.
- Probar unidades: puntos basicos, decimales, porcentajes, VN y precio por VN 100.
- Probar fechas no laborables y convencion de liquidacion.
- Agregar casos de flujos amortizables, cero cupon, CER y tasa flotante.
- Marcar como no calculable cualquier instrumento cuyo calendario contractual no pueda validarse.

Las tolerancias numericas se fijaran por tipo de prueba. Como referencia inicial, precio y TIR deben converger con error relativo menor a `1e-8` en casos sinteticos, mientras que el error de Taylor se informa y no se oculta porque depende del tamano del shock.

## Decisiones pendientes

- Fuente contractual autoritativa y mecanismo de actualizacion de flujos de nuevas emisiones.
- Convencion de capitalizacion y day count por familia cuando no este explicitada en la fuente actual.
- Fuente de curva TAMAR futura: ultimo valor, curva plana, REM o sendero configurable.
- Persistencia de historial intradiario para estimar volatilidad, rango y profundidad, especialmente para ventas cortas.
- Formato de entrada de posiciones reales y costos por ALyC.

Estas decisiones no bloquean las fases 1 y 2. Deben resolverse antes de considerar productivas las metricas CER/TAMAR y la recomendacion operativa de ventas cortas.

## Definicion de terminado

El desarrollo se considera completo cuando:

- Las familias fija, CER y TAMAR usan convenciones y shocks apropiados.
- Cada metrica puede trazarse hasta precio, flujo, indice y supuesto fuente.
- El reporte diario presenta Taylor de primer y segundo orden junto con revaluacion exacta.
- DV01 y P&L se expresan por unidad estandar y por posicion.
- Las posiciones vendidas incluyen signo y costos correctos.
- El Excel, CSV, JSON y Markdown se generan en una corrida reproducible.
- Las pruebas automatizadas cubren formulas, unidades, signos, fechas y datos faltantes.
- La salida combina sensibilidad de tasas con liquidez antes de destacar candidatos operativos.
