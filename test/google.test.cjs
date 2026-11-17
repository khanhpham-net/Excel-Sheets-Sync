const { test } = require('node:test');
const assert = require('node:assert/strict');
const { createGoogle, spreadsheetId } = require('../app/google.cjs');
const { parseRange, readCell } = require('../app/sync.cjs');

test('validates pasted Google Sheets links', () => {
  assert.equal(spreadsheetId('https://docs.google.com/spreadsheets/d/123456789_abc/edit#gid=0'), '123456789_abc');
  assert.throws(() => spreadsheetId('https://example.org/spreadsheets/d/123456789_abc/edit'));
});

test('reads a tab by ID and sends a value-only write through the signed-in account', async context => {
  const id = '123456789_abc';
  const book = { properties: { title: 'Test' }, sheets: [{ properties: { sheetId: 17, title: 'Kho', sheetType: 'GRID', gridProperties: { rowCount: 100, columnCount: 20 } } }] };
  const cell = { userEnteredValue: { formulaValue: '=3+4' }, effectiveValue: { numberValue: 7 }, formattedValue: '7' };
  const grid = { sheets: [{ properties: { sheetId: 17 }, data: [{ startRow: 2, startColumn: 1, rowData: [{ values: [cell, { effectiveValue: { stringValue: 'spill' } }] }] }] }] };
  const account = { refresh: 'test-token' };
  const client = { client_id: 'test-client', client_secret: 'test-secret' };
  const calls = [];
  context.mock.method(globalThis, 'fetch', async (input, options) => {
    const url = new URL(input);
    calls.push({ url, options });
    if (url.hostname === 'oauth2.googleapis.com') {
      assert.equal(options.body.get('refresh_token'), account.refresh);
      return Response.json({ access_token: 'token', expires_in: 3600 });
    }
    assert.equal(options.headers.Authorization, 'Bearer token');
    if (options.method === 'POST') return Response.json({});
    return Response.json(url.searchParams.has('ranges') ? grid : book);
  });
  const api = createGoogle(() => account, () => client);
  const opened = await api.open(id);
  const tab = opened.tabs[0];
  assert.deepEqual(tab, { id: 17, title: 'Kho', rows: 100, cols: 20 });
  const data = await api.grid(id, tab, parseRange('B3:C3'));
  assert.deepEqual({ value: readCell(data, 2, 1).value, formula: readCell(data, 2, 1).formula }, { value: 7, formula: '=3+4' });
  assert.equal(readCell(data, 2, 2).derived, true);
  assert.equal(calls[2].url.searchParams.get('ranges'), "'Kho'!B3:C3");
  assert.equal(calls[2].url.searchParams.get('fields'), 'sheets(properties(sheetId),merges,data(startRow,startColumn,rowData(values(userEnteredValue,effectiveValue,formattedValue))))');
  await api.write(id, tab, [{ row: 4, col: 2, value: '0012' }]);
  const request = JSON.parse(calls[3].options.body).requests[0].updateCells;
  assert.deepEqual(request.start, { sheetId: 17, rowIndex: 4, columnIndex: 2 });
  assert.equal(request.fields, 'userEnteredValue');
  assert.deepEqual(request.rows[0].values, [{ userEnteredValue: { stringValue: '0012' } }]);
});
