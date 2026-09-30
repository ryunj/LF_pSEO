import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import vm from 'node:vm';

const html = fs.readFileSync(new URL('../dashboard.html', import.meta.url), 'utf8');

function functionSource(name) {
  const start = html.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `${name} must exist`);
  const brace = html.indexOf('{', start);
  let depth = 0;
  for (let i = brace; i < html.length; i += 1) {
    if (html[i] === '{') depth += 1;
    if (html[i] === '}') depth -= 1;
    if (depth === 0) return html.slice(start, i + 1);
  }
  throw new Error(`Unclosed function ${name}`);
}

const context = {};
vm.createContext(context);
vm.runInContext(functionSource('metricFromName'), context);
vm.runInContext(functionSource('parsePerformanceFiles'), context);
vm.runInContext('const FILE_INPUT={perfFiles:[]};', context);
vm.runInContext(functionSource('keepPerformanceFile'), context);

test('separate PV and UV files stay separate and combine by date/code', () => {
  const perfFiles = [
    { sourceName: 'pSEO_PV.csv', rows: [['AF코드', '20260928'], ['PSBRD1', '7']] },
    { sourceName: 'pSEO_UV.csv', rows: [['AF코드', '20260928'], ['PSBRD1', '3']] },
  ];
  const codes = [{ c: 'PSBRD1' }];
  const result = context.parsePerformanceFiles(perfFiles, { PSBRD1: 'PSBRD1' }, codes, { PSBRD1: 0 }, {});
  assert.deepEqual(JSON.parse(JSON.stringify(result.rows)), [['2026-09-28', 0, 7, 3]]);
});

test('explicit PV and UV rows win over a misleading UV filename', () => {
  const perfFiles = [{
    sourceName: 'pSEO_UV.csv',
    rows: [['구분', 'AF코드', '20260928'], ['PV', 'PSBRD1', '7'], ['UV', 'PSBRD1', '3']],
  }];
  const codes = [{ c: 'PSBRD1' }];
  const result = context.parsePerformanceFiles(perfFiles, { PSBRD1: 'PSBRD1' }, codes, { PSBRD1: 0 }, {});
  assert.deepEqual(JSON.parse(JSON.stringify(result.rows)), [['2026-09-28', 0, 7, 3]]);
});

test('a mixed metric file replaces stale single-metric selections', () => {
  context.keepPerformanceFile({ sourceName: 'old_PV.csv', rows: [['AF코드', '20260921'], ['PSBRD1', '4']] });
  context.keepPerformanceFile({
    sourceName: 'new_UV.csv',
    rows: [['구분', 'AF코드', '20260928'], ['PV', 'PSBRD1', '7'], ['UV', 'PSBRD1', '3']],
  });
  assert.equal(vm.runInContext('FILE_INPUT.perfFiles.length', context), 1);
  assert.equal(vm.runInContext('FILE_INPUT.perfFiles[0].sourceName', context), 'new_UV.csv');
});
