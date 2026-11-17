const { _electron: electron } = require('playwright-core');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const XLSX = require('xlsx');

async function smoke() {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'excel-sheets-sync-test-'));
  const fixture = path.join(directory, 'Source.xlsx');
  const book = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([['SKU', 'Qty'], ['VT-01', 41], ['VT-02', 12], ['VT-03', '22.000.000 ₫']]);
  XLSX.utils.book_append_sheet(book, sheet, 'Stock');
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['Order', 'Cases'], ['DH-01', 12]]), 'Orders');
  const tall = Array.from({ length: 1200 }, (_, i) => [`SKU-${i + 1}`, i]);
  tall[0][51] = 'Last column';
  const wideSheet = XLSX.utils.aoa_to_sheet(tall);
  wideSheet['!cols'] = [{ wpx: 150 }, { wpx: 90 }];
  wideSheet['!merges'] = [XLSX.utils.decode_range('AX1:AZ1')];
  wideSheet.AX1 = { t: 's', v: 'Merged range' };
  delete wideSheet.AZ1;
  XLSX.utils.book_append_sheet(book, wideSheet, 'Many rows');
  XLSX.writeFile(book, fixture, { cellStyles: true });
  const packaged = process.argv[2];
  const env = { ...process.env, ELECTRON_ENABLE_LOGGING: '1', EXCEL_SHEETS_TEST_PROFILE: directory };
  delete env.ELECTRON_RUN_AS_NODE;
  const app = await electron.launch({
    executablePath: packaged || require('electron'),
    args: packaged ? [] : [path.resolve(__dirname, '..')], env,
  });
  try {
    const window = await app.firstWindow();
    window.on('pageerror', error => { throw error; });
    await window.getByRole('button', { name: 'Đăng nhập Google' }).click();
    await window.getByRole('dialog', { name: 'Kết nối Google Sheets' }).getByText('Desktop app', { exact: false }).first().waitFor();
    await window.getByRole('button', { name: 'Đóng', exact: true }).click();
    await window.getByRole('button', { name: 'Mở dữ liệu mẫu' }).click();
    await window.getByRole('button', { name: 'Xem trước' }).click();
    await window.locator('#preview-panel').getByText('4 ô trên 3 dòng', { exact: false }).waitFor();
    await window.getByRole('button', { name: 'Đồng bộ', exact: true }).click();
    await window.locator('#status').getByText('Đã ghi 4 ô trên 3 dòng', { exact: false }).waitFor();
    assert.equal(await window.locator('.pane').nth(1).locator('[data-row="4"][data-col="1"]').textContent(), '0');
    assert.equal(await window.locator('.pane').nth(1).locator('[data-row="1"][data-col="2"]').textContent(), 'Kệ B');
    await window.evaluate(() => {
      const input = document.createElement('input');
      input.type = 'file';
      input.id = 'smoke-file';
      document.body.append(input);
    });
    await window.locator('#smoke-file').setInputFiles(fixture);
    await window.evaluate(() => {
      const transfer = new DataTransfer();
      transfer.items.add(document.getElementById('smoke-file').files[0]);
      document.querySelector('.pane').dispatchEvent(new DragEvent('drop', { bubbles: true, cancelable: true, dataTransfer: transfer }));
    });
    await window.getByRole('heading', { name: 'Source.xlsx' }).waitFor();
    const excelTabs = window.getByLabel('Tab Excel', { exact: true });
    assert.equal(await excelTabs.locator('option').count(), 3);
    await excelTabs.selectOption({ label: 'Many rows' });
    const excelGrid = window.getByRole('region', { name: 'Ô bảng tính' }).first();
    await excelGrid.getByText('SKU-1', { exact: true }).waitFor();
    const firstCell = excelGrid.locator('[data-row="0"][data-col="0"]');
    assert.ok((await firstCell.boundingBox()).width > 125);
    const start = await firstCell.boundingBox();
    const bounds = await excelGrid.boundingBox();
    await window.mouse.move(start.x + 15, start.y + 8);
    await window.mouse.down();
    await window.mouse.move(bounds.x + bounds.width + 5, bounds.y + bounds.height + 12, { steps: 5 });
    await window.waitForFunction(() => {
      const grid = document.querySelector('.pane .grid');
      return grid.scrollLeft > 540 && grid.scrollTop > 500 && /^A1:[A-Z]+[3-9]\d/.test(document.getElementById('source-range').value);
    });
    await window.mouse.up();
    const selected = await window.locator('#source-range').inputValue();
    assert.match(selected, /^A1:[A-Z]+\d+$/);
    await excelGrid.evaluate(grid => { grid.scrollTop = 0; grid.scrollLeft = 0; });
    await firstCell.waitFor();
    await excelGrid.hover();
    await window.keyboard.down('Shift');
    await window.mouse.wheel(0, 4100);
    await window.keyboard.up('Shift');
    await excelGrid.getByText('Merged range', { exact: true }).waitFor();
    assert.equal(await excelGrid.locator('[data-row="0"][data-col="49"]').first().getAttribute('class'), 'grid-cell merged');
    await excelGrid.evaluate(grid => { grid.scrollTop = 0; grid.scrollLeft = 0; });
    await firstCell.waitFor();
    await window.mouse.move(bounds.x + 270, bounds.y + 130);
    await window.keyboard.down('Space');
    await window.mouse.down();
    await window.mouse.move(bounds.x + 100, bounds.y + 80, { steps: 5 });
    await window.mouse.up();
    await window.keyboard.up('Space');
    assert.ok(await excelGrid.evaluate(grid => grid.scrollLeft >= 150 && grid.scrollTop >= 40));
    assert.equal(await window.locator('#source-range').inputValue(), selected);
    await excelGrid.evaluate(grid => { grid.scrollTop = 0; grid.scrollLeft = 0; });
    await excelGrid.hover();
    await window.mouse.wheel(0, 3900);
    await window.waitForFunction(() => document.querySelector('.pane .grid').scrollTop > 3000 && Number(document.querySelector('.pane .grid [data-row]')?.dataset.row) > 60);
    const jump = window.getByLabel('Tới dòng').first();
    await jump.fill('985');
    await jump.press('Enter');
    await excelGrid.getByText('SKU-985', { exact: true }).waitFor();
    const visible = await excelGrid.getByText('SKU-985', { exact: true }).evaluate(cell => {
      const container = cell.closest('.grid').getBoundingClientRect();
      const row = cell.getBoundingClientRect();
      return row.top >= container.top && row.bottom <= container.bottom;
    });
    assert.equal(visible, true);
    await excelTabs.selectOption({ label: 'Stock' });
    await excelGrid.getByText('VT-01', { exact: true }).waitFor();
    await window.getByRole('button', { name: 'Cột A Excel' }).dragTo(window.getByRole('button', { name: 'Cột A Google Sheets' }));
    await window.getByLabel('Dùng A → A làm khóa đối chiếu').check();
    await window.getByRole('button', { name: 'Cột B Excel' }).click();
    await window.getByRole('button', { name: 'Cột B Google Sheets' }).click();
    const writeNumber = window.getByRole('checkbox', { name: 'Ghi thành số B → B' });
    await writeNumber.check();
    await window.getByLabel('Tên cách ghép').fill('Kho của Lan');
    await window.getByRole('button', { name: 'Lưu cách ghép' }).click();
    await window.locator('#status').getByText('Đã lưu cách ghép', { exact: false }).waitFor();
    await window.getByLabel('Cách ghép đã lưu').selectOption({ label: 'Kho của Lan · Mẫu' });
    await window.getByRole('button', { name: 'Mở', exact: true }).click();
    await window.locator('#status').getByText('Đã mở cách ghép', { exact: false }).waitFor();
    assert.equal(await writeNumber.isChecked(), true);
    await window.getByRole('button', { name: 'Xem trước' }).click();
    await window.locator('#preview-panel').getByText('2 ô trên 2 dòng', { exact: false }).waitFor();
    await window.locator('#preview-panel').getByText('22.000.000 (số)', { exact: true }).waitFor();
    sheet.B2.v = 42;
    XLSX.writeFile(book, fixture, { cellStyles: true });
    await window.getByRole('button', { name: 'Đồng bộ', exact: true }).click();
    await window.locator('#status').getByText('đã đổi sau khi xem trước', { exact: false }).waitFor();
    await window.getByLabel('Cách đồng bộ').selectOption('position');
    await window.getByLabel('Vùng', { exact: true }).nth(0).fill('B2');
    await window.getByLabel('Vùng', { exact: true }).nth(1).fill('D2');
    await window.getByRole('button', { name: 'Xem trước' }).click();
    await window.locator('#status').getByText('Ô đích D2 chứa công thức', { exact: false }).waitFor();
    await window.getByRole('button', { name: 'Kéo vùng' }).dragTo(window.locator('.pane').nth(1).locator('[data-row="0"][data-col="4"]'));
    assert.equal(await window.getByLabel('Vùng', { exact: true }).nth(1).inputValue(), 'E1');
    await window.getByRole('button', { name: 'Xem trước' }).click();
    await window.locator('#preview-panel').getByText('1 ô trên 1 dòng', { exact: false }).waitFor();
    await window.getByRole('button', { name: 'Đồng bộ', exact: true }).click();
    await window.locator('#status').getByText('Đã ghi 1 ô trên 1 dòng', { exact: false }).waitFor();
    assert.equal(await window.locator('.pane').nth(1).locator('[data-row="0"][data-col="4"]').textContent(), '42');

    const client = path.join(directory, 'desktop-oauth-test.json');
    await fs.writeFile(client, JSON.stringify({ installed: { client_id: 'test.apps.googleusercontent.com', client_secret: 'fake-test-secret' } }));
    await app.evaluate(({ dialog, shell }, file) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [file] });
      const send = globalThis.fetch;
      globalThis.fetch = async (url, options) => {
        if (String(url).startsWith('http://127.0.0.1:')) return send(url, options);
        if (url === 'https://oauth2.googleapis.com/token') return Response.json({ access_token: 'fake-token', refresh_token: 'fake-refresh', scope: 'openid email https://www.googleapis.com/auth/spreadsheets' });
        if (url === 'https://openidconnect.googleapis.com/v1/userinfo') return Response.json({ sub: 'test-user', email: 'demo@example.com' });
        throw new Error('Unexpected network request in test.');
      };
      shell.openExternal = async link => { globalThis.pendingAuthTest = link; };
    }, client);
    await window.getByRole('button', { name: 'Đăng nhập Google' }).click();
    await window.getByRole('button', { name: 'Chọn JSON và đăng nhập' }).click();
    await window.getByRole('button', { name: 'Hủy đăng nhập' }).click();
    await window.locator('#status').getByText('Đã hủy đăng nhập Google.', { exact: true }).waitFor();
    await app.evaluate(({ shell }) => {
      shell.openExternal = async link => {
        const auth = new URL(link);
        const callback = new URL(auth.searchParams.get('redirect_uri'));
        callback.searchParams.set('state', auth.searchParams.get('state'));
        callback.searchParams.set('code', 'test-code');
        const response = await fetch(callback);
        if (!response.ok) throw new Error('OAuth callback failed.');
      };
    });
    await window.getByRole('button', { name: 'Đăng nhập Google' }).click();
    await window.locator('#account').getByText('demo@example.com', { exact: true }).waitFor();

    await app.evaluate(() => {
      const currentFetch = globalThis.fetch;
      let rejectGrid = true;
      globalThis.fetch = async (input, options) => {
        const url = new URL(input);
        if (url.hostname !== 'sheets.googleapis.com') return currentFetch(input, options);
        if (!url.searchParams.has('ranges')) return Response.json({ properties: { title: 'Test sheet' }, sheets: [{ properties: { sheetId: 0, title: 'Trang tính1', sheetType: 'GRID', gridProperties: { rowCount: 100, columnCount: 10 } } }] });
        if (rejectGrid) {
          rejectGrid = false;
          return Response.json({ error: { message: 'Request contains an invalid argument.' } }, { status: 400 });
        }
        const fields = url.searchParams.get('fields');
        if (!fields.endsWith('formattedValue,effectiveFormat))))')) throw new Error('Invalid field mask in styled grid request.');
        const cell = { userEnteredValue: { stringValue: 'Mã NV' }, formattedValue: 'Mã NV', effectiveFormat: { backgroundColor: { blue: 0.5 }, textFormat: { bold: true }, horizontalAlignment: 'CENTER' } };
        const title = { userEnteredValue: { stringValue: 'A large title flows across styled empty cells' }, effectiveFormat: { wrapStrategy: 'OVERFLOW_CELL', textFormat: { fontSize: 18, bold: true } } };
        const styledBlank = { effectiveFormat: { backgroundColor: { red: 1, green: 1, blue: 1 } } };
        const clipped = { ...title, effectiveFormat: { ...title.effectiveFormat, wrapStrategy: 'CLIP' } };
        const occupied = { userEnteredValue: { numberValue: 0 }, formattedValue: '0' };
        return Response.json({ sheets: [{ properties: { sheetId: 0 }, data: [{ columnMetadata: [{ pixelSize: 135 }], rowData: [{ values: [cell] }, { values: [title, styledBlank, styledBlank, occupied] }, { values: [clipped, styledBlank] }] }] }] });
      };
    });
    await window.getByLabel('Link Google Sheets').fill('https://docs.google.com/spreadsheets/d/example-sheet-123/edit');
    await window.getByRole('button', { name: 'Mở Sheets' }).click();
    const targetGrid = window.getByRole('region', { name: 'Ô bảng tính' }).nth(1);
    await targetGrid.getByText('Không tải được ô trong tab Trang tính1.', { exact: false }).waitFor();
    await targetGrid.getByRole('button', { name: 'Thử tải lại' }).click();
    await targetGrid.getByText('Mã NV', { exact: true }).waitFor();
    const targetHeader = targetGrid.locator('[data-row="0"][data-col="0"]');
    assert.equal(await targetHeader.evaluate(cell => getComputedStyle(cell).backgroundColor), 'rgb(0, 0, 128)');
    assert.equal(await targetHeader.evaluate(cell => getComputedStyle(cell).fontWeight), '700');
    assert.equal((await targetHeader.boundingBox()).width, 135);
    const overflow = targetGrid.locator('[data-row="1"][data-col="0"]');
    const text = await overflow.evaluate(cell => {
      const box = cell.querySelector('.cell-text').getBoundingClientRect();
      const blocked = cell.parentElement.querySelector('[data-row="1"][data-col="3"]').getBoundingClientRect();
      return { width: box.width, stopsBeforeContent: box.right <= blocked.left };
    });
    assert.ok(text.width > 300 && text.stopsBeforeContent);
    assert.equal(await targetGrid.locator('[data-row="2"][data-col="0"] .cell-text').evaluate(node => getComputedStyle(node).textOverflow), 'clip');

    const panes = window.locator('#panes > .pane');
    const splitWidth = (await panes.nth(1).boundingBox()).width;
    const excelRange = window.locator('#source-range');
    const sheetRange = window.locator('#target-range');
    await excelRange.fill('A2:B3');
    await sheetRange.fill('C4');
    await window.getByRole('button', { name: 'Phóng to Google Sheets', exact: true }).click();
    assert.ok((await panes.nth(1).boundingBox()).width > splitWidth * 1.8);
    await panes.nth(0).locator('.mini-open').click();
    assert.equal(await window.locator('#panes').getAttribute('data-focus'), 'source');
    await window.keyboard.press('Escape');
    assert.equal(await window.locator('#panes').getAttribute('data-focus'), null);
    await window.getByRole('button', { name: 'Phóng to Excel', exact: true }).click();
    await window.getByRole('button', { name: 'Trở về chia đôi từ Excel', exact: true }).click();
    assert.equal(await excelRange.inputValue(), 'A2:B3');
    assert.equal(await sheetRange.inputValue(), 'C4');
    console.log('Windows smoke passed: styled overflow, clip, focus/switch/restore, native scrolling, Sheets load/retry, mapping, and formula protection.');
  } finally {
    await app.close();
  }
}

smoke().catch(error => { console.error(error); process.exitCode = 1; });
