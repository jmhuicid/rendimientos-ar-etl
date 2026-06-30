const fs = require('fs');
const path = require('path');
const fetch = require('node-fetch');

const ROOT = path.resolve(__dirname, '..');
const OUT_ROOT = path.join(ROOT, 'data', 'snapshots');

const FCI_CATEGORIES = [
  { key: 'mercadoDinero', label: 'mercado_dinero' },
  { key: 'rentaFija', label: 'renta_fija' },
  { key: 'rentaMixta', label: 'renta_mixta' },
  { key: 'rentaVariable', label: 'renta_variable' },
];

const BCRA_VARIABLES = [
  { id: 1, key: 'reservas' },
  { id: 4, key: 'usd_minorista' },
  { id: 5, key: 'usd_mayorista' },
  { id: 7, key: 'badlar_tna' },
  { id: 8, key: 'tm20' },
  { id: 12, key: 'depositos_30d_tna' },
  { id: 27, key: 'inflacion_mensual' },
  { id: 28, key: 'inflacion_interanual' },
  { id: 30, key: 'cer' },
  { id: 31, key: 'uva' },
  { id: 40, key: 'icl' },
  { id: 44, key: 'tamar_tna' },
];

async function main() {
  const startedAt = new Date();
  const runDate = startedAt.toISOString().slice(0, 10);
  const outDir = path.join(OUT_ROOT, runDate);
  fs.mkdirSync(outDir, { recursive: true });

  const sections = await collectAll();
  const manifest = {
    run_date: runDate,
    started_at: startedAt.toISOString(),
    finished_at: new Date().toISOString(),
    excluded_domains: ['mundial', 'prode'],
    files: {},
    counts: {},
    warnings: sections.warnings,
  };

  for (const [name, payload] of Object.entries(sections.data)) {
    const file = `${name}.json`;
    writeJSON(path.join(outDir, file), payload);
    manifest.files[name] = path.relative(ROOT, path.join(outDir, file)).replace(/\\/g, '/');
    manifest.counts[name] = countPayload(payload);
  }

  writeJSON(path.join(outDir, 'manifest.json'), manifest);
  writeJSON(path.join(OUT_ROOT, 'latest.json'), {
    ...manifest,
    snapshot_dir: path.relative(ROOT, outDir).replace(/\\/g, '/'),
  });

  console.log(JSON.stringify(manifest, null, 2));
}

async function collectAll() {
  const warnings = [];
  const settled = await Promise.allSettled([
    collectFci(),
    collectPlazoFijo(),
    collectBcra(),
    collectData912(),
    collectPellegrini(),
  ]);

  const names = ['fci', 'plazo_fijo', 'bcra', 'mercado', 'pellegrini'];
  const data = {};
  settled.forEach((result, idx) => {
    const name = names[idx];
    if (result.status === 'fulfilled') {
      data[name] = result.value;
    } else {
      warnings.push({ section: name, error: result.reason.message });
      data[name] = { error: result.reason.message, collected_at: new Date().toISOString() };
    }
  });
  return { data, warnings };
}

async function collectFci() {
  const categories = {};
  const all = [];

  await Promise.all(FCI_CATEGORIES.map(async (cat) => {
    const [latest, previous] = await Promise.all([
      fetchJSON(`https://api.argentinadatos.com/v1/finanzas/fci/${cat.key}/ultimo`),
      fetchJSON(`https://api.argentinadatos.com/v1/finanzas/fci/${cat.key}/penultimo`),
    ]);

    const previousByFund = new Map(previous.filter(validFciRow).map((row) => [row.fondo, row]));
    const rows = latest.filter(validFciRow).map((row) => {
      const prev = previousByFund.get(row.fondo);
      const daily = prev ? calcDailyReturn(row, prev) : null;
      return {
        category: cat.label,
        fondo: row.fondo,
        horizonte: row.horizonte || null,
        fecha: row.fecha,
        vcp: numberOrNull(row.vcp),
        ccp: numberOrNull(row.ccp),
        patrimonio: numberOrNull(row.patrimonio),
        prev_fecha: prev ? prev.fecha : null,
        prev_vcp: prev ? numberOrNull(prev.vcp) : null,
        daily_return: daily,
        tna: daily == null ? null : round(daily * 365 * 100, 4),
        tem_30d: daily == null ? null : round((Math.pow(1 + daily, 30) - 1) * 100, 4),
        tea: daily == null ? null : round((Math.pow(1 + daily, 365) - 1) * 100, 4),
      };
    });

    categories[cat.label] = {
      source_category: cat.key,
      latest_count: latest.length,
      valid_count: rows.length,
      rows,
    };
    all.push(...rows);
  }));

  all.sort((a, b) => (b.tna ?? -Infinity) - (a.tna ?? -Infinity));
  return {
    source: 'ArgentinaDatos',
    source_url: 'https://api.argentinadatos.com/v1/finanzas/fci',
    collected_at: new Date().toISOString(),
    total_valid: all.length,
    categories,
    ranking_by_tna: all,
  };
}

async function collectPlazoFijo() {
  const rows = await fetchJSON('https://api.argentinadatos.com/v1/finanzas/tasas/plazoFijo');
  return {
    source: 'ArgentinaDatos',
    source_url: 'https://api.argentinadatos.com/v1/finanzas/tasas/plazoFijo',
    collected_at: new Date().toISOString(),
    count: Array.isArray(rows) ? rows.length : null,
    rows,
  };
}

async function collectBcra() {
  const variables = await Promise.all(BCRA_VARIABLES.map(async (item) => {
    const url = `https://api.bcra.gob.ar/estadisticas/v4.0/Monetarias/${item.id}?limit=5`;
    const json = await fetchJSON(url);
    const detalle = json.results?.[0]?.detalle || [];
    return {
      ...item,
      latest: detalle[0] || null,
      previous: detalle[1] || null,
      history: detalle,
    };
  }));

  return {
    source: 'BCRA Estadisticas Monetarias v4.0',
    source_url: 'https://api.bcra.gob.ar/estadisticas/v4.0',
    collected_at: new Date().toISOString(),
    variables,
  };
}

async function collectData912() {
  const [notes, bonds, corp] = await Promise.all([
    fetchJSON('https://data912.com/live/arg_notes'),
    fetchJSON('https://data912.com/live/arg_bonds'),
    fetchJSON('https://data912.com/live/arg_corp'),
  ]);

  return {
    source: 'data912',
    collected_at: new Date().toISOString(),
    arg_notes: Array.isArray(notes) ? notes : [],
    arg_bonds: Array.isArray(bonds) ? bonds : [],
    arg_corp: Array.isArray(corp) ? corp : [],
  };
}

async function collectPellegrini() {
  const url = 'https://www.pellegrinifci.com.ar/valores-disponibles';
  const html = await fetchText(url);
  const text = htmlToText(html);
  const dateMatch = text.match(/VALOR DE CUOTAPARTE HISTORICO\s+Fecha\s+CONSULTAR\s+FONDO.*?AL\s+(\d{1,2}\/\d{1,2}\/\d{2})/i)
    || text.match(/Valor cuotaparte al\s+(\d{1,2}\/\d{1,2}\/\d{4})/i);

  return {
    source: 'Pellegrini FCI',
    source_url: url,
    collected_at: new Date().toISOString(),
    reported_date: dateMatch ? normalizeArDate(dateMatch[1]) : null,
    parsed_rows: parsePellegriniRows(text),
    raw_text_excerpt: text.slice(0, 12000),
  };
}

function parsePellegriniRows(text) {
  const rows = [];
  const regex = /([A-ZÁÉÍÓÚÑ0-9][A-ZÁÉÍÓÚÑa-záéíóúñ0-9 $.-]+?)\s+-\s+(Unica|Clase [A-Z])\s+Valor cuotaparte al\s+(\d{1,2}\/\d{1,2}\/\d{4})\s+\(cada mil\):\s+([\d.,]+)\s+Variacion %:\s+(-?[\d.,]+)/g;
  let match;
  while ((match = regex.exec(text)) !== null) {
    rows.push({
      fondo: match[1].trim(),
      clase: match[2],
      fecha: normalizeArDate(match[3]),
      vcp_cada_mil: parseArNumber(match[4]),
      variacion_pct: parseArNumber(match[5]),
    });
  }
  return rows;
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

async function fetchText(url) {
  const response = await fetch(url, {
    headers: {
      Accept: 'text/html,application/xhtml+xml',
      'User-Agent': 'rendimientos-ar-etl/1.0',
    },
  });
  if (!response.ok) throw new Error(`HTTP ${response.status} from ${url}`);
  return response.text();
}

function validFciRow(row) {
  return row && row.fecha && Number.isFinite(Number(row.vcp));
}

function calcDailyReturn(row, prev) {
  const days = daysBetween(row.fecha, prev.fecha);
  if (days <= 0 || !prev.vcp) return null;
  return (Number(row.vcp) - Number(prev.vcp)) / Number(prev.vcp) / days;
}

function daysBetween(a, b) {
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

function htmlToText(html) {
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/\s+/g, ' ')
    .trim();
}

function parseArNumber(value) {
  if (value == null) return null;
  const normalized = String(value).replace(/\./g, '').replace(',', '.');
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

function normalizeArDate(value) {
  const [d, m, y] = value.split('/');
  const year = y.length === 2 ? `20${y}` : y;
  return `${year}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
}

function countPayload(payload) {
  if (payload == null) return 0;
  if (Array.isArray(payload)) return payload.length;
  if (payload.total_valid != null) return payload.total_valid;
  if (payload.count != null) return payload.count;
  if (payload.variables) return payload.variables.length;
  if (payload.arg_notes) {
    return payload.arg_notes.length + payload.arg_bonds.length + payload.arg_corp.length;
  }
  if (payload.parsed_rows) return payload.parsed_rows.length;
  return Object.keys(payload).length;
}

function writeJSON(file, data) {
  fs.writeFileSync(file, `${JSON.stringify(data, null, 2)}\n`, 'utf8');
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
