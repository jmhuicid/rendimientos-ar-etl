const fs = require('fs');
const path = require('path');
const fetch = require('node-fetch');

const ROOT = path.resolve(__dirname, '..');
const REPORT_ROOT = path.join(ROOT, 'data', 'reports');

const OFFICIAL_SOURCES = [
  {
    institution: 'BCRA',
    name: 'Estadisticas e indicadores',
    cadence: 'diaria / segun serie',
    role: 'Reservas, tipo de cambio oficial/mayorista, tasas, agregados monetarios, inflacion, CER, UVA, ICL.',
    url: 'https://www.bcra.gob.ar/estadisticas-indicadores/',
    machine_url: 'https://api.bcra.gob.ar/estadisticas/v4.0',
    formats: ['API', 'web', 'series historicas'],
  },
  {
    institution: 'BCRA',
    name: 'Relevamiento de Expectativas de Mercado (REM)',
    cadence: 'mensual',
    role: 'Expectativas de inflacion, tipo de cambio, actividad, tasa TAMAR, sector externo y resultado fiscal.',
    url: 'https://www.bcra.gob.ar/relevamiento-expectativas-mercado-rem/',
    formats: ['PDF', 'XLSX', 'historico XLSX'],
  },
  {
    institution: 'BCRA',
    name: 'Mercado de cambios / sector externo',
    cadence: 'mensual',
    role: 'Lectura de flujos cambiarios, balance cambiario, cuenta corriente cambiaria e intervenciones cuando se publican.',
    url: 'https://www.bcra.gob.ar/estadisticas-indicadores/',
    formats: ['web', 'PDF/XLSX segun publicacion'],
  },
  {
    institution: 'Ministerio de Economia',
    name: 'Portal de Datos Economicos',
    cadence: 'segun serie',
    role: 'Series macroeconomicas oficiales para actividad, precios, sector externo, fiscal y monetario-financiero.',
    url: 'https://www.economia.gob.ar/datos',
    formats: ['web', 'descargas', 'series'],
  },
  {
    institution: 'Ministerio de Economia',
    name: 'Subsecretaria de Programacion Macroeconomica',
    cadence: 'segun publicacion',
    role: 'Informes macroeconomicos, actividad, empleo, pobreza, precios, balanza de pagos, sector monetario, cambiario y financiero.',
    url: 'https://www.argentina.gob.ar/economia/politicaeconomica/macroeconomica',
    formats: ['web', 'informes'],
  },
  {
    institution: 'Ministerio de Economia',
    name: 'Gasto Publico Consolidado',
    cadence: 'anual / actualizacion oficial',
    role: 'Base oficial e informe de gasto publico consolidado por finalidad, funcion y nivel de gobierno.',
    url: 'https://www.argentina.gob.ar/economia/politicaeconomica/macroeconomica/gastopublicoconsolidado',
    formats: ['informe', 'descargas de datos'],
  },
  {
    institution: 'Ministerio de Economia / ArgentinaDatos',
    name: 'ArgentinaDatos - finanzas',
    cadence: 'diaria / frecuente',
    role: 'Dolar historico, riesgo pais, plazo fijo y FCI como capa de datos publica reutilizable.',
    url: 'https://api.argentinadatos.com',
    formats: ['API'],
  },
];

const BCRA_VARIABLES = [
  { id: 1, key: 'reservas', label: 'Reservas internacionales', unit: 'MM USD', kind: 'level' },
  { id: 4, key: 'usd_minorista', label: 'Dolar minorista vendedor', unit: '$/USD', kind: 'fx' },
  { id: 5, key: 'usd_mayorista', label: 'Dolar mayorista referencia', unit: '$/USD', kind: 'fx' },
  { id: 7, key: 'badlar_tna', label: 'BADLAR privados', unit: '% TNA', kind: 'rate' },
  { id: 8, key: 'tm20_tna', label: 'TM20 privados', unit: '% TNA', kind: 'rate' },
  { id: 12, key: 'depositos_30d_tna', label: 'Depositos 30 dias', unit: '% TNA', kind: 'rate' },
  { id: 27, key: 'inflacion_mensual', label: 'Inflacion mensual', unit: '%', kind: 'pct' },
  { id: 28, key: 'inflacion_interanual', label: 'Inflacion interanual', unit: '%', kind: 'pct' },
  { id: 30, key: 'cer', label: 'CER', unit: 'indice', kind: 'level' },
  { id: 31, key: 'uva', label: 'UVA', unit: '$', kind: 'level' },
  { id: 40, key: 'icl', label: 'ICL alquileres', unit: 'indice', kind: 'level' },
  { id: 44, key: 'tamar_tna', label: 'TAMAR privados', unit: '% TNA', kind: 'rate' },
];

async function main() {
  const runDate = new Date().toISOString().slice(0, 10);
  const outDir = path.join(REPORT_ROOT, runDate);
  fs.mkdirSync(outDir, { recursive: true });

  const collected = await collectReport();
  const report = buildReport(runDate, collected);

  writeJSON(path.join(outDir, 'argentina-daily-report.json'), report);
  fs.writeFileSync(path.join(outDir, 'argentina-daily-report.md'), toMarkdown(report), 'utf8');
  fs.writeFileSync(path.join(outDir, 'argentina-daily-key-metrics.csv'), toCSV(report.key_metrics), 'utf8');

  console.log(toMarkdown(report));
}

async function collectReport() {
  const names = ['dolar', 'riesgo_pais', 'bcra', 'plazo_fijo', 'fci_mm', 'mercado'];
  const tasks = [
    collectDollar(),
    collectRiskCountry(),
    collectBcra(),
    collectFixedTerm(),
    collectMoneyMarketFunds(),
    collectMarketPrices(),
  ];

  const settled = await Promise.allSettled(tasks);
  const data = {};
  const warnings = [];
  settled.forEach((result, idx) => {
    const name = names[idx];
    if (result.status === 'fulfilled') {
      data[name] = result.value;
    } else {
      data[name] = null;
      warnings.push({ section: name, error: result.reason.message });
    }
  });
  return { data, warnings };
}

async function collectDollar() {
  const rows = await fetchJSON('https://dolarapi.com/v1/dolares');
  const wanted = ['oficial', 'mayorista', 'blue', 'bolsa', 'contadoconliqui', 'cripto'];
  const quotes = rows
    .filter((row) => wanted.includes(row.casa))
    .map((row) => ({
      key: row.casa,
      name: ascii(row.nombre || row.casa),
      buy: numberOrNull(row.compra),
      sell: numberOrNull(row.venta),
      updated_at: row.fechaActualizacion || null,
    }));

  const histories = await Promise.allSettled(
    quotes
      .filter((quote) => quote.key !== 'cripto')
      .map(async (quote) => [quote.key, await fetchDollarHistory(quote.key)])
  );

  const historyByKey = new Map();
  for (const result of histories) {
    if (result.status === 'fulfilled') historyByKey.set(result.value[0], result.value[1]);
  }

  const official = quotes.find((q) => q.key === 'oficial');
  for (const quote of quotes) {
    quote.gap_vs_official_sell = official?.sell && quote.sell
      ? round(((quote.sell / official.sell) - 1) * 100, 2)
      : null;
    const history = historyByKey.get(quote.key);
    if (history?.length) {
      addDollarChanges(quote, history);
    }
  }
  return {
    source: 'DolarAPI',
    source_url: 'https://dolarapi.com/v1/dolares + ArgentinaDatos historico',
    quotes,
  };
}

async function fetchDollarHistory(key) {
  const rows = await fetchJSON(`https://api.argentinadatos.com/v1/cotizaciones/dolares/${key}`);
  return rows
    .filter((row) => row.fecha && Number.isFinite(Number(row.venta)))
    .map((row) => ({ date: row.fecha, sell: Number(row.venta) }))
    .sort((a, b) => a.date.localeCompare(b.date));
}

function addDollarChanges(quote, history) {
  const latest = history[history.length - 1];
  const prev = history[history.length - 2] || null;
  const monthAgo = findOnOrBefore(history, addDays(latest.date, -30));
  const yearAgo = findOnOrBefore(history, addDays(latest.date, -365));

  quote.history_date = latest.date;
  quote.change_1d_pct = prev ? pctChange(latest.sell, prev.sell) : null;
  quote.change_30d_pct = monthAgo ? pctChange(latest.sell, monthAgo.sell) : null;
  quote.change_1y_pct = yearAgo ? pctChange(latest.sell, yearAgo.sell) : null;
}

async function collectRiskCountry() {
  const row = await fetchJSON('https://api.argentinadatos.com/v1/finanzas/indices/riesgo-pais/ultimo');
  return {
    source: 'ArgentinaDatos',
    source_url: 'https://api.argentinadatos.com/v1/finanzas/indices/riesgo-pais/ultimo',
    value: numberOrNull(row.valor),
    date: row.fecha || null,
  };
}

async function collectBcra() {
  const variables = await Promise.all(BCRA_VARIABLES.map(async (def) => {
    const url = `https://api.bcra.gob.ar/estadisticas/v4.0/Monetarias/${def.id}?limit=5`;
    const json = await fetchJSON(url);
    const detail = json.results?.[0]?.detalle || [];
    const latest = detail[0] || null;
    const previous = detail[1] || null;
    return {
      ...def,
      latest_date: latest?.fecha || null,
      latest_value: numberOrNull(latest?.valor),
      previous_date: previous?.fecha || null,
      previous_value: numberOrNull(previous?.valor),
      change_abs: latest && previous ? round(Number(latest.valor) - Number(previous.valor), 4) : null,
      history: detail,
    };
  }));
  return {
    source: 'BCRA Estadisticas Monetarias v4.0',
    source_url: 'https://api.bcra.gob.ar/estadisticas/v4.0',
    variables,
  };
}

async function collectFixedTerm() {
  const rows = await fetchJSON('https://api.argentinadatos.com/v1/finanzas/tasas/plazoFijo');
  const normalized = rows.map((row) => ({
    entidad: ascii(row.entidad),
    tna_clientes: normalizeRate(row.tnaClientes),
    tna_no_clientes: normalizeRate(row.tnaNoClientes),
  }));
  const ranked = normalized
    .map((row) => ({ ...row, best_tna: Math.max(row.tna_clientes || 0, row.tna_no_clientes || 0) || null }))
    .filter((row) => row.best_tna != null)
    .sort((a, b) => b.best_tna - a.best_tna);
  return {
    source: 'ArgentinaDatos / BCRA plazo fijo',
    source_url: 'https://api.argentinadatos.com/v1/finanzas/tasas/plazoFijo',
    count: normalized.length,
    top: ranked.slice(0, 10),
    rows: normalized,
  };
}

async function collectMoneyMarketFunds() {
  const [latest, previous] = await Promise.all([
    fetchJSON('https://api.argentinadatos.com/v1/finanzas/fci/mercadoDinero/ultimo'),
    fetchJSON('https://api.argentinadatos.com/v1/finanzas/fci/mercadoDinero/penultimo'),
  ]);
  const previousByFund = new Map(previous.filter(validFciRow).map((row) => [row.fondo, row]));
  const rows = latest.filter(validFciRow).map((row) => {
    const prev = previousByFund.get(row.fondo);
    const daily = prev ? calcDailyReturn(row, prev) : null;
    return {
      fondo: ascii(row.fondo),
      fecha: row.fecha,
      vcp: numberOrNull(row.vcp),
      patrimonio: numberOrNull(row.patrimonio),
      prev_fecha: prev?.fecha || null,
      tna: daily == null ? null : round(daily * 365 * 100, 2),
      tem30: daily == null ? null : round((Math.pow(1 + daily, 30) - 1) * 100, 2),
      tea: daily == null ? null : round((Math.pow(1 + daily, 365) - 1) * 100, 2),
    };
  });
  const liquid = rows
    .filter((row) => (row.patrimonio || 0) >= 1_000_000_000 && row.tna != null)
    .sort((a, b) => b.tna - a.tna);
  return {
    source: 'ArgentinaDatos FCI mercadoDinero',
    source_url: 'https://api.argentinadatos.com/v1/finanzas/fci/mercadoDinero',
    mercado_fondo: rows.find((row) => row.fondo === 'Mercado Fondo - Clase A') || null,
    top: liquid.slice(0, 10),
    count: rows.length,
  };
}

async function collectMarketPrices() {
  const [notes, bonds, corp] = await Promise.all([
    fetchJSON('https://data912.com/live/arg_notes'),
    fetchJSON('https://data912.com/live/arg_bonds'),
    fetchJSON('https://data912.com/live/arg_corp'),
  ]);
  const symbols = ['AL30', 'AL30D', 'GD30', 'GD30D', 'AL35', 'AL35D', 'GD35', 'GD35D'];
  const selected = (Array.isArray(bonds) ? bonds : [])
    .filter((row) => symbols.includes(row.symbol))
    .map((row) => ({
      symbol: row.symbol,
      price: numberOrNull(row.c),
      bid: numberOrNull(row.px_bid),
      ask: numberOrNull(row.px_ask),
      pct_change: numberOrNull(row.pct_change),
    }));
  return {
    source: 'data912',
    source_url: 'https://data912.com/live',
    selected_bonds: selected,
    counts: {
      notes: Array.isArray(notes) ? notes.length : 0,
      bonds: Array.isArray(bonds) ? bonds.length : 0,
      corp: Array.isArray(corp) ? corp.length : 0,
    },
  };
}

function buildReport(runDate, collected) {
  const data = collected.data;
  const bcraMap = new Map((data.bcra?.variables || []).map((item) => [item.key, item]));
  const dollarMap = new Map((data.dolar?.quotes || []).map((item) => [item.key, item]));

  const keyMetrics = [
    metric('Dolar oficial venta', dollarMap.get('oficial')?.sell, '$/USD', dollarMap.get('oficial')?.updated_at, 'DolarAPI'),
    metric('Dolar mayorista venta', dollarMap.get('mayorista')?.sell, '$/USD', dollarMap.get('mayorista')?.updated_at, 'DolarAPI'),
    metric('Dolar blue venta', dollarMap.get('blue')?.sell, '$/USD', dollarMap.get('blue')?.updated_at, 'DolarAPI'),
    metric('Dolar MEP venta', dollarMap.get('bolsa')?.sell, '$/USD', dollarMap.get('bolsa')?.updated_at, 'DolarAPI'),
    metric('Dolar CCL venta', dollarMap.get('contadoconliqui')?.sell, '$/USD', dollarMap.get('contadoconliqui')?.updated_at, 'DolarAPI'),
    metric('Brecha blue vs oficial', dollarMap.get('blue')?.gap_vs_official_sell, '%', dollarMap.get('blue')?.updated_at, 'DolarAPI'),
    metric('Riesgo pais', data.riesgo_pais?.value, 'pbs', data.riesgo_pais?.date, 'ArgentinaDatos'),
    metricFromBcra(bcraMap.get('reservas')),
    metricFromBcra(bcraMap.get('badlar_tna')),
    metricFromBcra(bcraMap.get('tamar_tna')),
    metricFromBcra(bcraMap.get('inflacion_mensual')),
    metricFromBcra(bcraMap.get('inflacion_interanual')),
    metricFromBcra(bcraMap.get('cer')),
    metricFromBcra(bcraMap.get('uva')),
    metric('Mejor plazo fijo TNA', data.plazo_fijo?.top?.[0]?.best_tna, '% TNA', runDate, 'ArgentinaDatos'),
    metric('Mercado Fondo TNA aprox', data.fci_mm?.mercado_fondo?.tna, '% TNA', data.fci_mm?.mercado_fondo?.fecha, 'ArgentinaDatos'),
    metric('Mercado Fondo TEM30 aprox', data.fci_mm?.mercado_fondo?.tem30, '%', data.fci_mm?.mercado_fondo?.fecha, 'ArgentinaDatos'),
  ].filter(Boolean);

  return {
    run_date: runDate,
    generated_at: new Date().toISOString(),
    title: 'Estado de Situacion Financiera Argentina',
    scope: 'Lectura de apertura para decision profesional: dolar, reservas, tasas, inflacion, CER/UVA, FCI money market, plazo fijo, riesgo pais y precios testigo de bonos.',
    caveats: [
      'Los datos pueden tener rezagos distintos por fuente.',
      'FCI Money Market usa ultimo y penultimo VCP; TEM30/TNA son aproximaciones de observacion diaria.',
      'El informe es informativo y no constituye recomendacion de inversion.',
    ],
    key_metrics: keyMetrics,
    official_sources: OFFICIAL_SOURCES,
    sections: data,
    warnings: collected.warnings,
  };
}

function toMarkdown(report) {
  const dollarMap = new Map((report.sections.dolar?.quotes || []).map((item) => [item.key, item]));
  const bcraMap = new Map((report.sections.bcra?.variables || []).map((item) => [item.key, item]));
  const blue = dollarMap.get('blue');
  const ccl = dollarMap.get('contadoconliqui');
  const mep = dollarMap.get('bolsa');
  const official = dollarMap.get('oficial');
  const reserves = bcraMap.get('reservas');
  const badlar = bcraMap.get('badlar_tna');
  const tamar = bcraMap.get('tamar_tna');
  const inflationMonthly = bcraMap.get('inflacion_mensual');
  const risk = report.sections.riesgo_pais;
  const mercadoFondo = report.sections.fci_mm?.mercado_fondo;
  const bestPf = report.sections.plazo_fijo?.top?.[0];

  const lines = [];
  lines.push('# Informe de apertura - Argentina');
  lines.push('');
  lines.push(`Fecha de lectura: ${report.run_date}`);
  lines.push('');
  lines.push('## Lectura ejecutiva');
  lines.push('');
  lines.push(`El arranque financiero queda dominado por un dolar blue en ${formatValue(blue?.sell, '$/USD')} y un CCL en ${formatValue(ccl?.sell, '$/USD')}. La brecha blue contra oficial se ubica en ${formatPct(blue?.gap_vs_official_sell)} y la del CCL en ${formatPct(ccl?.gap_vs_official_sell)}, niveles que sirven como termometro de presion cambiaria y demanda de cobertura.`);
  lines.push('');
  lines.push(`Las reservas internacionales informadas por BCRA estan en ${formatValue(reserves?.latest_value, 'MM USD')} al ${reserves?.latest_date || 'n/d'}, con variacion contra el dato previo de ${formatSignedNumber(reserves?.change_abs)} MM USD. En tasas, BADLAR marca ${formatPct(badlar?.latest_value)} y TAMAR ${formatPct(tamar?.latest_value)}, contra una inflacion mensual informada de ${formatPct(inflationMonthly?.latest_value)}.`);
  lines.push('');
  if (mercadoFondo || bestPf) {
    lines.push(`Para liquidez en pesos, Mercado Fondo muestra una TNA aproximada de ${formatPct(mercadoFondo?.tna)} y TEM30 de ${formatPct(mercadoFondo?.tem30)}, mientras el mejor plazo fijo relevado paga ${formatPct(bestPf?.best_tna)} TNA. La comparacion relevante al inicio del dia es tasa en pesos versus ritmo de crawling/variacion de dolares financieros.`);
    lines.push('');
  }
  if (risk?.value != null) {
    lines.push(`El riesgo pais en ${formatNumber(risk.value)} pbs al ${risk.date || 'n/d'} completa la lectura de credito soberano: si sube junto con CCL/MEP, el mercado suele estar pagando mas cobertura y menos duration argentina.`);
    lines.push('');
  }
  lines.push('## Tablero clave');
  lines.push('');
  lines.push('| Metrica | Valor | Fecha/Fuente |');
  lines.push('| --- | ---: | --- |');
  for (const item of report.key_metrics) {
    lines.push(`| ${item.name} | ${formatValue(item.value, item.unit)} | ${item.date || 'n/d'} / ${item.source} |`);
  }
  lines.push('');

  const dollar = report.sections.dolar;
  if (dollar?.quotes?.length) {
    lines.push('## Mercado cambiario');
    lines.push('');
    lines.push('| Tipo | Venta | 1D | 30D | 1Y | Brecha vs oficial | Actualizado |');
    lines.push('| --- | ---: | ---: | ---: | ---: | ---: | --- |');
    for (const row of dollar.quotes) {
      lines.push(`| ${row.name} | ${formatValue(row.sell, '$/USD')} | ${formatPct(row.change_1d_pct)} | ${formatPct(row.change_30d_pct)} | ${formatPct(row.change_1y_pct)} | ${formatPct(row.gap_vs_official_sell)} | ${row.updated_at || row.history_date || 'n/d'} |`);
    }
    lines.push('');
    lines.push(`Lectura: MEP en ${formatValue(mep?.sell, '$/USD')} y CCL en ${formatValue(ccl?.sell, '$/USD')} son las referencias de dolar financiero. La distancia CCL-MEP es ${formatPct(ccl?.sell && mep?.sell ? ((ccl.sell / mep.sell) - 1) * 100 : null)}, util para leer demanda de dolar cable frente a dolar local.`);
    lines.push('');
  }

  const bcra = report.sections.bcra;
  if (bcra?.variables?.length) {
    lines.push('## Liquidez, reservas e inflacion');
    lines.push('');
    lines.push('| Indicador | Ultimo | Var dato previo | Fecha |');
    lines.push('| --- | ---: | ---: | --- |');
    for (const key of ['reservas', 'badlar_tna', 'tamar_tna', 'depositos_30d_tna', 'inflacion_mensual', 'inflacion_interanual', 'cer', 'uva', 'icl']) {
      const row = bcra.variables.find((item) => item.key === key);
      if (!row) continue;
      lines.push(`| ${row.label} | ${formatValue(row.latest_value, row.unit)} | ${formatSigned(row.change_abs, row.unit)} | ${row.latest_date || 'n/d'} |`);
    }
    lines.push('');
    lines.push('Lectura: la curva corta en pesos se evalua contra inflacion mensual, crawling cambiario y brechas. Si tasas reales esperadas quedan comprimidas y los financieros aceleran, sube el incentivo a cobertura.');
    lines.push('');
  }

  const pf = report.sections.plazo_fijo;
  if (pf?.top?.length) {
    lines.push('## Tasas en pesos');
    lines.push('');
    lines.push('| Entidad | Mejor TNA | Clientes | No clientes |');
    lines.push('| --- | ---: | ---: | ---: |');
    for (const row of pf.top.slice(0, 8)) {
      lines.push(`| ${row.entidad} | ${formatPct(row.best_tna)} | ${formatPct(row.tna_clientes)} | ${formatPct(row.tna_no_clientes)} |`);
    }
    lines.push('');
  }

  const fci = report.sections.fci_mm;
  if (fci?.mercado_fondo || fci?.top?.length) {
    lines.push('## Liquidez remunerada');
    lines.push('');
    if (fci.mercado_fondo) {
      lines.push(`Mercado Fondo - Clase A al ${fci.mercado_fondo.fecha}: TNA aprox ${formatPct(fci.mercado_fondo.tna)}, TEM30 ${formatPct(fci.mercado_fondo.tem30)}, TEA ${formatPct(fci.mercado_fondo.tea)}.`);
      lines.push('');
    }
    lines.push('| Fondo | Fecha | Patrimonio | TNA aprox | TEM30 |');
    lines.push('| --- | --- | ---: | ---: | ---: |');
    for (const row of (fci.top || []).slice(0, 8)) {
      lines.push(`| ${escapeMd(row.fondo)} | ${row.fecha} | ${formatMoney(row.patrimonio)} | ${formatPct(row.tna)} | ${formatPct(row.tem30)} |`);
    }
    lines.push('');
  }

  const market = report.sections.mercado;
  if (market?.selected_bonds?.length) {
    lines.push('## Credito soberano y bonos testigo');
    lines.push('');
    lines.push('| Simbolo | Precio | Bid | Ask | Var % |');
    lines.push('| --- | ---: | ---: | ---: | ---: |');
    for (const row of market.selected_bonds) {
      lines.push(`| ${row.symbol} | ${formatNumber(row.price)} | ${formatNumber(row.bid)} | ${formatNumber(row.ask)} | ${formatPct(row.pct_change)} |`);
    }
    lines.push('');
  }

  lines.push('## Fuentes oficiales y reportes a monitorear');
  lines.push('');
  lines.push('| Institucion | Fuente | Frecuencia | Uso en la lectura | Formatos |');
  lines.push('| --- | --- | --- | --- | --- |');
  for (const source of report.official_sources) {
    lines.push(`| ${source.institution} | [${source.name}](${source.url}) | ${source.cadence} | ${source.role} | ${source.formats.join(', ')} |`);
  }
  lines.push('');
  lines.push('Lectura: los datos diarios salen principalmente de APIs/series; los Excel y PDF oficiales de BCRA y Economia aportan contexto mensual/anual y deben revisarse cuando se actualizan para recalibrar expectativas, fiscal, deuda, sector externo y actividad.');
  lines.push('');

  if (report.warnings.length) {
    lines.push('## Datos no disponibles');
    lines.push('');
    for (const warning of report.warnings) {
      lines.push(`- ${warning.section}: ${warning.error}`);
    }
    lines.push('');
  }

  lines.push('## Que mirar durante la rueda');
  lines.push('');
  lines.push('- Si CCL y MEP amplian brecha contra oficial: presion de cobertura y/o salida por dolar financiero.');
  lines.push('- Si reservas caen mientras el mayorista sube: tension entre flujo comercial, intervencion y expectativas.');
  lines.push('- Si tasas BADLAR/TAMAR quedan por debajo de inflacion esperada: peor premio por permanecer en pesos.');
  lines.push('- Si riesgo pais sube junto con bonos en baja: deterioro de credito soberano y mayor prima de duration.');
  lines.push('- Si FCI Money Market rinde menos que plazo fijo pero mantiene liquidez inmediata: la decision es rendimiento versus disponibilidad.');
  lines.push('');
  lines.push('## Notas metodologicas');
  lines.push('');
  for (const caveat of report.caveats) lines.push(`- ${caveat}`);
  lines.push('');
  return `${lines.join('\n')}\n`;
}

function toCSV(rows) {
  const headers = ['name', 'value', 'unit', 'date', 'source'];
  const lines = [headers.join(',')];
  for (const row of rows) lines.push(headers.map((h) => csvCell(row[h])).join(','));
  return `${lines.join('\n')}\n`;
}

async function fetchJSON(url) {
  const response = await fetch(url, {
    headers: {
      Accept: 'application/json',
      'User-Agent': 'rendimientos-ar-etl/1.0',
    },
  });
  if (!response.ok) throw new Error(`HTTP ${response.status} from ${url}`);
  return response.json();
}

function metric(name, value, unit, date, source) {
  if (value == null || Number.isNaN(value)) return null;
  return { name, value, unit, date: date || null, source };
}

function metricFromBcra(item) {
  if (!item) return null;
  return metric(item.label, item.latest_value, item.unit, item.latest_date, 'BCRA');
}

function validFciRow(row) {
  return row && row.fecha && Number.isFinite(Number(row.vcp));
}

function calcDailyReturn(row, prev) {
  const days = Math.abs(Math.round((new Date(row.fecha) - new Date(prev.fecha)) / 86400000));
  if (days <= 0 || !prev.vcp) return null;
  return (Number(row.vcp) - Number(prev.vcp)) / Number(prev.vcp) / days;
}

function normalizeRate(value) {
  const n = numberOrNull(value);
  if (n == null) return null;
  return n <= 1 ? round(n * 100, 4) : round(n, 4);
}

function findOnOrBefore(rows, targetDate) {
  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i].date <= targetDate) return rows[i];
  }
  return null;
}

function addDays(dateString, days) {
  const date = new Date(`${dateString}T00:00:00Z`);
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

function pctChange(current, previous) {
  if (!Number.isFinite(current) || !Number.isFinite(previous) || previous === 0) return null;
  return round(((current / previous) - 1) * 100, 2);
}

function numberOrNull(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function round(value, digits) {
  const factor = 10 ** digits;
  return Math.round(value * factor) / factor;
}

function formatValue(value, unit) {
  if (unit === '$/USD') return `$${formatNumber(value)}`;
  if (unit && unit.includes('%')) return formatPct(value);
  if (unit === 'pbs') return `${formatNumber(value)} pbs`;
  if (unit === 'MM USD') return `${formatNumber(value)} MM USD`;
  return `${formatNumber(value)} ${unit || ''}`.trim();
}

function formatNumber(value) {
  if (value == null || !Number.isFinite(Number(value))) return 'n/d';
  return Number(value).toLocaleString('es-AR', { maximumFractionDigits: 2 });
}

function formatPct(value) {
  if (value == null || !Number.isFinite(Number(value))) return 'n/d';
  return `${Number(value).toLocaleString('es-AR', { maximumFractionDigits: 2 })}%`;
}

function formatSigned(value, unit) {
  if (value == null || !Number.isFinite(Number(value))) return 'n/d';
  const sign = Number(value) > 0 ? '+' : '';
  if (unit && unit.includes('%')) return `${sign}${formatPct(value)}`;
  return `${sign}${formatNumber(value)} ${unit || ''}`.trim();
}

function formatSignedNumber(value) {
  if (value == null || !Number.isFinite(Number(value))) return 'n/d';
  const sign = Number(value) > 0 ? '+' : '';
  return `${sign}${formatNumber(value)}`;
}

function formatMoney(value) {
  if (value == null || !Number.isFinite(Number(value))) return 'n/d';
  const n = Number(value);
  if (n >= 1_000_000_000_000) return `$${(n / 1_000_000_000_000).toFixed(2)} B`;
  if (n >= 1_000_000_000) return `$${(n / 1_000_000_000).toFixed(1)} MM`;
  if (n >= 1_000_000) return `$${(n / 1_000_000).toFixed(1)} M`;
  return `$${n.toFixed(0)}`;
}

function ascii(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^\x20-\x7E]/g, '')
    .trim();
}

function escapeMd(value) {
  return String(value).replace(/\|/g, '\\|');
}

function csvCell(value) {
  if (value == null) return '';
  const text = String(value);
  return /[",\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function writeJSON(file, data) {
  fs.writeFileSync(file, `${JSON.stringify(data, null, 2)}\n`, 'utf8');
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
