const fs = require('fs');
const path = require('path');
const fetch = require('node-fetch');

const ROOT = path.resolve(__dirname, '..');
const REPORT_ROOT = path.join(ROOT, 'data', 'reports');
const LATEST_SNAPSHOT = path.join(ROOT, 'data', 'snapshots', 'latest.json');
const SOURCE_URL = 'https://api.argentinadatos.com/v1/finanzas/fci/mercadoDinero';

async function main() {
  const rows = await loadRows();
  const validRows = rows
    .filter((row) => row.fecha && Number.isFinite(Number(row.vcp)) && Number.isFinite(Number(row.prev_vcp)))
    .map(normalizeRow)
    .filter((row) => row.daily_return != null);

  validRows.sort((a, b) => {
    const patrDiff = (b.patrimonio || 0) - (a.patrimonio || 0);
    if (a.fondo === 'Mercado Fondo - Clase A') return -1;
    if (b.fondo === 'Mercado Fondo - Clase A') return 1;
    return patrDiff;
  });

  const marketFund = validRows.find((row) => row.fondo === 'Mercado Fondo - Clase A') || null;
  const liquidRows = validRows.filter((row) => (row.patrimonio || 0) >= 1_000_000_000);
  const benchmarkRows = liquidRows
    .sort((a, b) => (b.tna || -Infinity) - (a.tna || -Infinity))
    .slice(0, 30);

  const runDate = new Date().toISOString().slice(0, 10);
  const outDir = path.join(REPORT_ROOT, runDate);
  fs.mkdirSync(outDir, { recursive: true });

  const csvRows = benchmarkRows.map((row) => ({
    fondo: row.fondo,
    fecha: row.fecha,
    vcp: row.vcp,
    patrimonio: row.patrimonio,
    tna_pct: row.tna,
    tem30_pct: row.tem30,
    tea_pct: row.tea,
  }));

  const report = {
    run_date: runDate,
    source: 'ArgentinaDatos FCI mercadoDinero ultimo/penultimo',
    source_url: SOURCE_URL,
    methodology: 'Rendimiento diario entre ultimo y penultimo VCP, nominalizado a 30 dias y anualizado como TNA/TEA. Es una aproximacion, no una serie mensual historica cerrada.',
    mercado_fondo: marketFund,
    benchmark_count: benchmarkRows.length,
    benchmark: benchmarkRows,
  };

  writeJSON(path.join(outDir, 'fci-money-market-benchmark.json'), report);
  fs.writeFileSync(path.join(outDir, 'fci-money-market-benchmark.csv'), toCSV(csvRows), 'utf8');
  fs.writeFileSync(path.join(outDir, 'fci-money-market-benchmark.md'), toMarkdown(report), 'utf8');

  console.log(toMarkdown(report));
}

async function loadRows() {
  if (fs.existsSync(LATEST_SNAPSHOT)) {
    const latest = JSON.parse(fs.readFileSync(LATEST_SNAPSHOT, 'utf8'));
    const fciFile = latest.files?.fci ? path.join(ROOT, latest.files.fci) : null;
    if (fciFile && fs.existsSync(fciFile)) {
      const snapshot = JSON.parse(fs.readFileSync(fciFile, 'utf8'));
      const rows = snapshot.categories?.mercado_dinero?.rows;
      if (Array.isArray(rows) && rows.length > 0) return rows;
    }
  }

  const [latest, previous] = await Promise.all([
    fetchJSON(`${SOURCE_URL}/ultimo`),
    fetchJSON(`${SOURCE_URL}/penultimo`),
  ]);
  const prevMap = new Map(previous.filter(validRawRow).map((row) => [row.fondo, row]));
  return latest.filter(validRawRow).map((row) => {
    const prev = prevMap.get(row.fondo);
    return {
      category: 'mercado_dinero',
      fondo: row.fondo,
      horizonte: row.horizonte || null,
      fecha: row.fecha,
      vcp: numberOrNull(row.vcp),
      ccp: numberOrNull(row.ccp),
      patrimonio: numberOrNull(row.patrimonio),
      prev_fecha: prev ? prev.fecha : null,
      prev_vcp: prev ? numberOrNull(prev.vcp) : null,
    };
  });
}

function normalizeRow(row) {
  const daily = row.daily_return != null ? Number(row.daily_return) : calcDailyReturn(row);
  return {
    ...row,
    vcp: numberOrNull(row.vcp),
    patrimonio: numberOrNull(row.patrimonio),
    daily_return: daily,
    tna: daily == null ? null : round(daily * 365 * 100, 2),
    tem30: daily == null ? null : round((Math.pow(1 + daily, 30) - 1) * 100, 2),
    tea: daily == null ? null : round((Math.pow(1 + daily, 365) - 1) * 100, 2),
  };
}

function calcDailyReturn(row) {
  const days = daysBetween(row.fecha, row.prev_fecha);
  if (days <= 0 || !row.prev_vcp) return null;
  return (Number(row.vcp) - Number(row.prev_vcp)) / Number(row.prev_vcp) / days;
}

function toMarkdown(report) {
  const lines = [];
  lines.push('# FCI Money Market - benchmark aproximado');
  lines.push('');
  lines.push(`Fecha corrida: ${report.run_date}`);
  lines.push('');
  lines.push(`Fuente: ${report.source}`);
  lines.push('');
  lines.push(`Metodologia: ${report.methodology}`);
  lines.push('');
  if (report.mercado_fondo) {
    const mf = report.mercado_fondo;
    lines.push('## Mercado Fondo');
    lines.push('');
    lines.push(`Mercado Fondo - Clase A al ${mf.fecha}: TNA ${fmt(mf.tna)}%, TEM30 ${fmt(mf.tem30)}%, TEA ${fmt(mf.tea)}%, patrimonio ${fmtMoney(mf.patrimonio)}.`);
    lines.push('');
  }
  lines.push('## Top Money Market liquidos');
  lines.push('');
  lines.push('| Fondo | Fecha | Patrimonio | TNA % | TEM30 % | TEA % |');
  lines.push('| --- | --- | ---: | ---: | ---: | ---: |');
  for (const row of report.benchmark) {
    lines.push(`| ${escapeMd(row.fondo)} | ${row.fecha} | ${fmtMoney(row.patrimonio)} | ${fmt(row.tna)} | ${fmt(row.tem30)} | ${fmt(row.tea)} |`);
  }
  lines.push('');
  return `${lines.join('\n')}\n`;
}

function toCSV(rows) {
  const headers = Object.keys(rows[0] || {
    fondo: '', fecha: '', vcp: '', patrimonio: '', tna_pct: '', tem30_pct: '', tea_pct: '',
  });
  const lines = [headers.join(',')];
  for (const row of rows) {
    lines.push(headers.map((h) => csvCell(row[h])).join(','));
  }
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

function validRawRow(row) {
  return row && row.fecha && Number.isFinite(Number(row.vcp));
}

function daysBetween(a, b) {
  if (!a || !b) return 0;
  return Math.abs(Math.round((new Date(a) - new Date(b)) / 86400000));
}

function numberOrNull(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function round(value, digits) {
  const factor = 10 ** digits;
  return Math.round(value * factor) / factor;
}

function fmt(value) {
  return value == null ? 'n/d' : value.toFixed(2);
}

function fmtMoney(value) {
  if (!Number.isFinite(value)) return 'n/d';
  if (value >= 1_000_000_000_000) return `$${(value / 1_000_000_000_000).toFixed(2)} B`;
  if (value >= 1_000_000_000) return `$${(value / 1_000_000_000).toFixed(1)} MM`;
  if (value >= 1_000_000) return `$${(value / 1_000_000).toFixed(1)} M`;
  return `$${value.toFixed(0)}`;
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
