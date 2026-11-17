const { test } = require('node:test');
const assert = require('node:assert/strict');
const { parseRange, makePlan, writeRequests, sheetRange } = require('../app/sync.cjs');

function grid(values, range) {
  return { range: parseRange(range), cells: values.map(row => row.map(value => typeof value === 'object' && value !== null ? value : { value })) };
}

test('matches reordered rows and appends a new zero quantity without writing notes', () => {
  const source = grid([['Branch', 'SKU', 'Stock'], ['A', '0001', 8], ['A', '0002', 0], ['A', '0003', null]], 'A1:C4');
  const target = grid([['SKU', 'Branch', 'Qty', 'Notes'], ['0003', 'A', 7, 'keep'], ['0001', 'B', 9, 'keep'], ['0001', 'A', 5, 'keep'], []], 'A1:D5');
  const plan = makePlan(source, target, {
    mode: 'key', sourceRange: source.range, targetRange: target.range, addMissing: true,
    mappings: [{ from: 0, to: 1, key: true }, { from: 1, to: 0, key: true }, { from: 2, to: 2 }],
  });
  assert.deepEqual(plan.changes.map(({ cell, value }) => [cell, value]), [['C4', 8], ['B5', 'A'], ['A5', '0002'], ['C5', 0]]);
  assert.equal(plan.skippedEmpty, 1);
  assert.equal(plan.added, 1);
  assert.equal(plan.rows, 2);
});

test('blocks formulas and preserves literal text in the write request', () => {
  const source = grid([['=1+2']], 'A1');
  const target = grid([[{ value: 1, formula: '=1' }]], 'B2');
  const definition = { mode: 'position', sourceRange: source.range, targetRange: target.range };
  assert.throws(() => makePlan(source, target, definition), /công thức/);
  const blank = grid([[null]], 'B2');
  const plan = makePlan(source, blank, definition);
  assert.deepEqual(writeRequests(91, plan.changes)[0].updateCells.rows[0].values[0], { userEnteredValue: { stringValue: '=1+2' } });
  assert.equal(sheetRange("Kho 'A'", parseRange('A1:B2')), "'Kho ''A'''!A1:B2");
});

test('blocks duplicate keys and overflow of the selected destination', () => {
  const source = grid([['Mã', 'Số'], ['A1', 2], ['A1', 3]], 'A1:B3');
  const destination = grid([['Mã', 'Số'], ['A9', 8]], 'A1:B2');
  const settings = { mode: 'key', sourceRange: source.range, targetRange: destination.range, mappings: [{ from: 0, to: 0, key: true }, { from: 1, to: 1 }], addMissing: true };
  assert.throws(() => makePlan(source, destination, settings), /trùng mã/);
  source.cells[2][0].value = 'A2';
  assert.throws(() => makePlan(source, destination, settings), /hết dòng trống/);
});

test('rewrites formatted Vietnamese amounts as numbers only in selected columns', () => {
  const source = grid([
    ['Mã', 'Lương', 'Mã phụ'],
    ['0012', '22.000.000 ₫', '0007'],
    ['0013', "'−1.500,25 VND", '0008'],
    ['0014', '0 ₫', '0009'],
  ], 'A6:C9');
  const target = grid([['Mã', 'Lương', 'Mã phụ'], ['0012', '22.000.000 ₫', '0070'], ['0013', -1500, '0008'], []], 'A1:C4');
  const definition = {
    mode: 'key', sourceRange: source.range, targetRange: target.range, addMissing: true,
    mappings: [{ from: 0, to: 0, key: true, asNumber: true }, { from: 1, to: 1, asNumber: true }, { from: 2, to: 2 }],
  };
  const plan = makePlan(source, target, definition);
  assert.deepEqual(plan.changes.map(({ cell, value }) => [cell, value]), [['B2', 22_000_000], ['C2', '0007'], ['B3', -1500.25], ['A4', '0014'], ['B4', 0], ['C4', '0009']]);
  assert.deepEqual(writeRequests(17, plan.changes)[0].updateCells.rows[0].values.map(value => value.userEnteredValue), [{ numberValue: 22_000_000 }, { stringValue: '0007' }]);
  source.cells[1][1].value = '22.00.0 ₫';
  assert.throws(() => makePlan(source, target, definition), /Excel B7 không phải số/);
});
