'use strict';

const fs = require('fs');
const path = require('path');
const defaultFetch = (...args) => {
  if (typeof globalThis.fetch !== 'function') throw new Error('Node.js 18 o superior es requerido');
  return globalThis.fetch(...args);
};

const DEFAULT_BASE_URL = 'https://api.bcra.gob.ar/estadisticas/v4.0/monetarias';

function parseArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i += 1) {
    if (argv[i].startsWith('--')) args[argv[i].slice(2)] = argv[i + 1] || true;
  }
  return args;
}

async function fetchJson(url, params = {}, fetchImpl = defaultFetch) {
  const target = new URL(url);
  for (const [key, value] of Object.entries(params)) target.searchParams.set(key, String(value));
  let lastError;
  for (let attempt = 0; attempt < 4; attempt += 1) {
    try {
      const response = await fetchImpl(target.toString(), { headers: { Accept: 'application/json' } });
      if ([404, 410].includes(response.status)) throw new Error(`Endpoint retirado o inexistente: HTTP ${response.status} ${target}`);
      if (!response.ok) throw new Error(`BCRA API HTTP ${response.status} ${target}`);
      const payload = await response.json();
      if (!payload || typeof payload !== 'object' || (payload.status != null && payload.status !== 200)) {
        throw new Error(`Respuesta BCRA inválida: ${JSON.stringify(payload).slice(0, 300)}`);
      }
      return payload;
    } catch (error) {
      lastError = error;
      if (attempt === 3 || /HTTP (400|404|410)/.test(error.message)) break;
      await new Promise(resolve => setTimeout(resolve, 1000 * (2 ** attempt)));
    }
  }
  throw lastError;
}

async function fetchCatalog(baseUrl, fetchImpl = defaultFetch) {
  const catalog = new Map();
  let offset = 0;
  const limit = 1000;
  while (true) {
    const payload = await fetchJson(baseUrl, { offset, limit }, fetchImpl);
    if (!Array.isArray(payload.results)) throw new Error('El catálogo BCRA no contiene una lista en results');
    for (const item of payload.results) if (Number.isInteger(item.idVariable)) catalog.set(item.idVariable, item);
    const count = Number(payload.metadata?.resultset?.count ?? payload.results.length);
    if (payload.results.length === 0 || payload.results.length < limit || catalog.size >= offset + count) break;
    offset += payload.results.length;
  }
  return catalog;
}

async function fetchSeries(baseUrl, id, from, to, fetchImpl = defaultFetch) {
  const payload = await fetchJson(`${baseUrl}/${id}`, { desde: from, hasta: to, limit: 1000 }, fetchImpl);
  const detail = payload.results?.[0]?.detalle;
  if (!Array.isArray(detail)) throw new Error(`Serie ${id}: detalle ausente o inválido`);
  return detail.filter(row => row && typeof row === 'object');
}

function isValidNumber(value) {
  return typeof value === 'number' && Number.isFinite(value);
}

function analyzeSeries(id, rows, catalogItem, previous = null, frequency = 'D') {
  const values = {};
  let duplicates = 0;
  let invalidValues = 0;
  for (const row of rows) {
    if (typeof row.fecha !== 'string') {
      invalidValues += 1;
      continue;
    }
    if (Object.hasOwn(values, row.fecha)) duplicates += 1;
    values[row.fecha] = row.valor;
    if (!isValidNumber(row.valor)) invalidValues += 1;
  }
  const dates = Object.keys(values).sort();
  const previousValues = previous?.values || {};
  const revisedValues = Object.entries(values).filter(([date, value]) => Object.hasOwn(previousValues, date) && previousValues[date] !== value).length;
  const newObservations = Object.keys(values).filter(date => !Object.hasOwn(previousValues, date)).length;
  const latestDate = dates.at(-1) || null;
  let status = dates.length && !duplicates && !invalidValues && !revisedValues ? 'OK' : 'WARNING';
  if (latestDate && previous?.latestDate && latestDate < previous.latestDate) status = 'ERROR';
  return {
    idVariable: id,
    status,
    description: catalogItem.descripcion || null,
    category: catalogItem.categoria || null,
    frequency,
    latestDate,
    previousLatestDate: previous?.latestDate || null,
    observations: rows.length,
    newObservations,
    duplicates,
    invalidValues,
    revisedValues,
    values,
    error: null,
  };
}

function renderMarkdown(report) {
  const lines = [
    '# Control de series BCRA', '',
    `- Estado: **${report.status}**`,
    `- Fecha: ${report.generatedAtArgentina}`,
    `- Endpoint: \`${report.endpoint}\``,
    `- Seguro actualizar consumidores: **${report.safeToRefresh ? 'Sí' : 'No'}**`, '',
    '| ID | Estado | Frecuencia | Última fecha | Nuevas | Duplicados | Revisiones | Error |',
    '|---:|:---:|:---:|:---:|---:|---:|---:|:---|',
  ];
  for (const item of report.series) {
    lines.push(`| ${item.idVariable} | ${item.status} | ${item.frequency || '-'} | ${item.latestDate || '-'} | ${item.newObservations} | ${item.duplicates} | ${item.revisedValues} | ${item.error || ''} |`);
  }
  return `${lines.join('\n')}\n`;
}

async function runMonitor({ configPath, outputDir, previousStatePath, fetchImpl = defaultFetch, today = new Date() }) {
  const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
  const baseUrl = (config.baseUrl || DEFAULT_BASE_URL).replace(/\/$/, '');
  const previous = previousStatePath && fs.existsSync(previousStatePath)
    ? JSON.parse(fs.readFileSync(previousStatePath, 'utf8')).series || {}
    : {};
  const end = today.toISOString().slice(0, 10);
  const startDate = new Date(today);
  startDate.setUTCDate(startDate.getUTCDate() - Number(config.lookbackDays || 45));
  const start = startDate.toISOString().slice(0, 10);
  const catalog = await fetchCatalog(baseUrl, fetchImpl);
  const series = [];
  const stateSeries = {};
  for (const configured of config.series) {
    const id = Number(configured.id);
    try {
      if (!catalog.has(id)) throw new Error(`ID ${id} ausente del catálogo v4.0; requiere reemplazo funcional validado`);
      const rows = await fetchSeries(baseUrl, id, start, end, fetchImpl);
      const result = analyzeSeries(id, rows, catalog.get(id), previous[id], catalog.get(id).periodicidad || configured.frequency);
      stateSeries[id] = { latestDate: result.latestDate, description: result.description, category: result.category, values: result.values };
      series.push(result);
    } catch (error) {
      series.push({ idVariable: id, status: 'ERROR', frequency: configured.frequency || null, latestDate: null, previousLatestDate: previous[id]?.latestDate || null, observations: 0, newObservations: 0, duplicates: 0, invalidValues: 0, revisedValues: 0, values: {}, error: error.message });
    }
  }
  const status = series.some(item => item.status === 'ERROR') ? 'ERROR' : series.some(item => item.status === 'WARNING') ? 'WARNING' : 'OK';
  const generatedAtArgentina = new Intl.DateTimeFormat('sv-SE', { timeZone: 'America/Argentina/Cordoba', dateStyle: 'short', timeStyle: 'medium' }).format(today);
  const report = { generatedAtArgentina, endpoint: baseUrl, apiVersion: 'v4.0', status, safeToRefresh: status === 'OK', series };
  const state = { generatedAtArgentina, endpoint: baseUrl, series: stateSeries };
  fs.mkdirSync(outputDir, { recursive: true });
  fs.writeFileSync(path.join(outputDir, 'bcra-series-report.json'), JSON.stringify(report, null, 2));
  fs.writeFileSync(path.join(outputDir, 'bcra-series-report.md'), renderMarkdown(report));
  fs.writeFileSync(path.join(outputDir, 'bcra-series-state.json'), JSON.stringify(state, null, 2));
  return report;
}

if (require.main === module) {
  const args = parseArgs(process.argv.slice(2));
  runMonitor({
    configPath: args.config || 'config/bcra-series-monitor.json',
    outputDir: args.output || 'artifacts/bcra-series-monitor',
    previousStatePath: args.previous,
  }).then(report => {
    console.log(JSON.stringify({ status: report.status, safeToRefresh: report.safeToRefresh }));
    process.exitCode = report.status === 'ERROR' ? 1 : 0;
  }).catch(error => {
    console.error(error);
    process.exitCode = 1;
  });
}

module.exports = { analyzeSeries, fetchCatalog, fetchSeries, renderMarkdown, runMonitor };
