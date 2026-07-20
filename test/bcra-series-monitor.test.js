'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { analyzeSeries, fetchCatalog } = require('../scripts/bcra-series-monitor');

test('detects new and revised observations', () => {
  const result = analyzeSeries(
    30,
    [{ fecha: '2026-07-16', valor: 10.5 }, { fecha: '2026-07-17', valor: 11 }],
    { descripcion: 'CER', categoria: 'Principales', periodicidad: 'D' },
    { latestDate: '2026-07-16', values: { '2026-07-16': 10 } },
    'D'
  );
  assert.equal(result.latestDate, '2026-07-17');
  assert.equal(result.newObservations, 1);
  assert.equal(result.revisedValues, 1);
  assert.equal(result.status, 'WARNING');
});

test('detects duplicate and invalid values', () => {
  const result = analyzeSeries(30, [
    { fecha: '2026-07-17', valor: 1 },
    { fecha: '2026-07-17', valor: null },
  ], {}, null, 'D');
  assert.equal(result.duplicates, 1);
  assert.equal(result.invalidValues, 1);
  assert.equal(result.status, 'WARNING');
});

test('paginates BCRA catalog', async () => {
  const calls = [];
  const fakeFetch = async url => {
    calls.push(url);
    const offset = Number(new URL(url).searchParams.get('offset'));
    const results = offset === 0
      ? Array.from({ length: 1000 }, (_, index) => ({ idVariable: index + 1 }))
      : [{ idVariable: 1001 }];
    return { ok: true, status: 200, json: async () => ({ status: 200, metadata: { resultset: { count: offset === 0 ? 1001 : 1 } }, results }) };
  };
  const catalog = await fetchCatalog('https://example.test/monetarias', fakeFetch);
  assert.equal(catalog.size, 1001);
  assert.equal(calls.length, 2);
});
