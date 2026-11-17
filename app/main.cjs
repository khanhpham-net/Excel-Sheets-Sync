const { app, BrowserWindow, dialog, ipcMain, safeStorage, shell } = require('electron');
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const XLSX = require('xlsx');
const { clientFromJson, signIn, spreadsheetId, createGoogle } = require('./google.cjs');
const { address, parseRange, readCell, makePlan } = require('./sync.cjs');
const { excelLayout, excelFormats } = require('./sheet-format.cjs');

app.setName('Excel Sheets Sync');
if (process.env.EXCEL_SHEETS_TEST_PROFILE) app.setPath('userData', process.env.EXCEL_SHEETS_TEST_PROFILE);
const MAX_FILE = 35 * 1024 * 1024;
let window;
let stateDir;
let client;
let account;
let source;
let target;
let preview;
let demoBook;
let demoSource;
let busy = false;
let loginAttempt;
const google = createGoogle(() => account, () => client);

function diskRead(name, fallback) {
  try { return JSON.parse(fs.readFileSync(path.join(stateDir, name), 'utf8')); }
  catch (error) { if (error.code === 'ENOENT') return fallback; throw error; }
}

function diskWrite(name, value) {
  fs.mkdirSync(stateDir, { recursive: true });
  const final = path.join(stateDir, name);
  const tmp = `${final}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(value, null, 2), 'utf8');
  fs.renameSync(tmp, final);
}

function restoreGoogle() {
  try {
    if (!safeStorage.isEncryptionAvailable()) return;
    const settings = diskRead('google-local.json', null);
    if (settings) ({ client, account } = JSON.parse(safeStorage.decryptString(Buffer.from(settings.encrypted, 'base64'))));
  } catch { client = undefined; account = undefined; }
}

function loadBundledOAuth() {
  if (client) return;
  const provisioned = path.join(app.isPackaged ? path.dirname(process.execPath) : app.getAppPath(), 'google-oauth.json');
  if (fs.existsSync(provisioned)) {
    try { client = clientFromJson(fs.readFileSync(provisioned, 'utf8')); }
    catch { /* Users can import a valid credential through the app. */ }
  }
}

function saveGoogle() {
  if (!safeStorage.isEncryptionAvailable()) return;
  diskWrite('google-local.json', { encrypted: safeStorage.encryptString(JSON.stringify({ client, account })).toString('base64') });
}

function info() {
  return { email: account?.email || '', configured: Boolean(client), profiles: profiles() };
}

function profileOwner() {
  return target?.demo ? 'demo' : account?.sub || '';
}

function profiles() {
  return diskRead('profiles.json', []).filter(item => item.owner === account?.sub || (item.owner === 'demo' && target?.demo));
}

function tabSize(sheet) {
  const end = XLSX.utils.decode_range(sheet['!ref'] || 'A1').e;
  return { rows: end.r + 1, cols: end.c + 1 };
}

function readWorkbook(file) {
  if (!/\.(xlsx|xlsm|xls|xlsb|csv|tsv)$/i.test(file)) throw new Error('Chọn file Excel, CSV hoặc TSV.');
  const before = fs.statSync(file);
  if (!before.isFile() || before.size > MAX_FILE) throw new Error('Bản thử nhận file tối đa 35 MB.');
  const data = fs.readFileSync(file);
  const after = fs.statSync(file);
  if (before.size !== after.size || before.mtimeMs !== after.mtimeMs) throw new Error('Excel đang lưu file. Hãy thử lại sau khi lưu xong.');
  const book = XLSX.read(data, { type: 'buffer', raw: true, cellFormula: true, cellNF: true, cellText: true, sheetStubs: true, cellStyles: true, bookFiles: true });
  const tabs = book.SheetNames.map(title => ({ id: title, title, ...tabSize(book.Sheets[title]) }));
  if (!tabs.length) throw new Error('File không có tab để đọc.');
  let formats = [];
  try { formats = excelFormats(book); } catch { /* Other workbook readers can still supply basic formatting. */ }
  return { book, file, title: path.basename(file), tabs, stamp: `${after.mtimeMs}/${after.size}`, formats, layouts: new Map() };
}

function sourceInfo() {
  return source && { title: source.title, tabs: source.tabs };
}

function demoCell(value) {
  return value == null ? {} : { value, display: String(value) };
}

function setupDemo() {
  const example = {
    'Tồn kho': [['SKU', 'Tồn cuối', 'Tên hàng'], ['VT-01', 24, 'Găng tay'], ['VT-02', 12, 'Khẩu trang'], ['VT-03', 8, 'Kính bảo hộ'], ['VT-04', 0, 'Băng keo']],
    'Đơn hàng': [['Mã đơn', 'Số kiện'], ['DH-01', 5], ['DH-02', 2]],
  };
  const book = XLSX.utils.book_new();
  for (const [name, rows] of Object.entries(example)) XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), name);
  demoSource = { book, title: 'ChiNhanhA.xlsx · Mẫu', tabs: book.SheetNames.map(title => ({ id: title, title, ...tabSize(book.Sheets[title]) })) };
  source = demoSource;
  demoBook = {
    0: [['Mã hàng', 'Số lượng', 'Ghi chú'], ['VT-02', 15, 'Kệ B'], ['VT-03', 8, 'Kệ C'], ['VT-01', 20, 'Kệ A']].map(row => row.map(demoCell)),
    1: [['Mã đơn', 'Kiện hàng'], ['DH-02', 2], ['DH-01', 3]].map(row => row.map(demoCell)),
  };
  demoBook[0][1][3] = { value: 30, display: '30', formula: '=B2*2' };
  target = { id: 'demo', title: 'Báo cáo công ty · Mẫu', tabs: [{ id: 0, title: 'Kho tổng', rows: 100, cols: 12 }, { id: 1, title: 'Đơn tổng', rows: 100, cols: 12 }], demo: true };
  preview = undefined;
  return { source: sourceInfo(), target, profiles: profiles() };
}

function selectedTab(book, id) {
  const tab = book?.tabs.find(item => item.id === id);
  if (!tab) throw new Error('Chọn một tab hợp lệ.');
  return tab;
}

function sourceGrid(tab, range, styled = false) {
  const sheet = source.book.Sheets[tab.id];
  const date1904 = source.book.Workbook?.WBProps?.date1904;
  let layout;
  if (styled) {
    source.layouts ||= new Map();
    if (!source.layouts.has(tab.id)) {
      let extracted;
      try { extracted = excelLayout(source.book, tab.id); }
      catch { extracted = excelLayout({ Sheets: source.book.Sheets }, tab.id); }
      source.layouts.set(tab.id, extracted);
    }
    layout = source.layouts.get(tab.id);
  }
  const cellAt = (row, col) => {
    const ref = address(row, col);
    const cell = sheet[ref];
    let value = cell?.v;
    if (cell?.t === 'n' && typeof value === 'number' && date1904 && cell.z && XLSX.SSF.is_date(cell.z) && value >= 1) value += 1462;
    if (cell?.t === 'n' && Number.isInteger(value) && /^0{2,}$/.test(cell.z || '') && /^0\d+$/.test(cell.w || '')) value = cell.w;
    let format;
    if (layout) {
      format = source.formats?.[layout.styles.get(ref) ?? 0];
      if (cell?.s?.patternType === 'solid' && /^[\da-f]{6}$/i.test(cell.s.fgColor?.rgb || '')) {
        format = { ...format, background: `#${cell.s.fgColor.rgb}` };
      }
    }
    return cell ? { value: cell.t === 'e' ? undefined : value, display: cell.w, formula: cell.f || (cell.F ? 'Công thức mảng' : undefined), error: cell.t === 'e', format } : { format };
  };
  const cells = [];
  for (let row = range.top; row <= range.bottom; row++) {
    const values = [];
    for (let col = range.left; col <= range.right; col++) {
      values.push(cellAt(row, col));
    }
    cells.push(values);
  }
  if (!layout) return { range, cells };
  const merges = layout.merges.filter(merge => merge.top <= range.bottom && merge.bottom >= range.top && merge.left <= range.right && merge.right >= range.left)
    .map(merge => ({ ...merge, anchor: cellAt(merge.top, merge.left) }));
  return { range, cells, merges, layout: { columnSize: layout.columnSize, rowSize: layout.rowSize, columns: layout.columns.slice(range.left, range.right + 1), rows: layout.rows.slice(range.top, range.bottom + 1), showGridlines: layout.showGridlines } };
}

function checkTarget(tab, range) {
  if (range.bottom >= tab.rows || range.right >= tab.cols) throw new Error('Vùng ghi nằm ngoài lưới đích. Thêm dòng hoặc cột trên Google Sheets rồi mở lại file.');
}

function targetGrid(tab, range, styled = false) {
  if (!target.demo) return google.grid(target.id, tab, range, styled);
  const rows = demoBook[tab.id];
  const cells = Array.from({ length: range.bottom - range.top + 1 }, (_, r) =>
    Array.from({ length: range.right - range.left + 1 }, (_, c) => rows[r + range.top]?.[c + range.left] || {}));
  return { range, cells };
}

function prepareDefinition(input) {
  const sourceTab = selectedTab(source, input.sourceTab);
  const targetTab = selectedTab(target, input.targetTab);
  const sourceRange = parseRange(input.sourceRange);
  const targetRange = parseRange(input.targetRange);
  const mode = input.mode;
  if (mode !== 'position' && mode !== 'key') throw new Error('Chọn cách đồng bộ.');
  if (sourceRange.bottom >= sourceTab.rows || sourceRange.right >= sourceTab.cols) throw new Error('Vùng nguồn vượt quá dữ liệu Excel. Hãy mở lại file hoặc chọn vùng nhỏ hơn.');
  if (mode === 'position' && targetRange.top === targetRange.bottom && targetRange.left === targetRange.right) {
    targetRange.bottom += sourceRange.bottom - sourceRange.top;
    targetRange.right += sourceRange.right - sourceRange.left;
  }
  checkTarget(targetTab, targetRange);
  const mappings = Array.isArray(input.mappings) ? input.mappings.map(({ from, to, key, asNumber }) => ({ from, to, key: key === true, asNumber: key !== true && asNumber === true })) : [];
  return { sourceTab, targetTab, definition: { mode, sourceRange, targetRange, mappings, addMissing: input.addMissing === true } };
}

function refreshSource() {
  if (!source?.file) return;
  const record = fs.statSync(source.file);
  if (`${record.mtimeMs}/${record.size}` !== source.stamp) source = readWorkbook(source.file);
}

async function createPlan(input) {
  refreshSource();
  const { sourceTab, targetTab, definition } = prepareDefinition(input);
  const from = sourceGrid(sourceTab, definition.sourceRange);
  const to = await targetGrid(targetTab, definition.targetRange);
  if (input.headers && JSON.stringify(input.headers) !== JSON.stringify(mappedHeaders(definition, from, to))) {
    throw new Error('Tiêu đề nguồn hoặc đích đã đổi từ lần lưu cách ghép. Kiểm tra và ghép lại các cột.');
  }
  const plan = makePlan(from, to, definition);
  return { plan, targetTab };
}

function fingerprint(plan) {
  return crypto.createHash('sha256').update(JSON.stringify(plan)).digest('hex');
}

function mappedHeaders(definition, sourceData, targetData) {
  return definition.mappings.map(({ from, to }) => [
    readCell(sourceData, definition.sourceRange.top, from).value ?? '',
    readCell(targetData, definition.targetRange.top, to).value ?? '',
  ]);
}

async function saveProfile(input) {
  if (!source || !target) throw new Error('Mở file nguồn và đích trước.');
  const owner = profileOwner();
  if (!owner) throw new Error('Đăng nhập Google để lưu cách ghép.');
  refreshSource();
  const { sourceTab, targetTab, definition } = prepareDefinition(input.definition);
  const toSave = { ...input.definition, headers: undefined };
  if (definition.mode === 'key') {
    const sourceData = sourceGrid(sourceTab, definition.sourceRange);
    const targetData = await targetGrid(targetTab, definition.targetRange);
    makePlan(sourceData, targetData, definition);
    toSave.headers = mappedHeaders(definition, sourceData, targetData);
  }
  const name = String(input.name || '').trim().slice(0, 100);
  if (!name) throw new Error('Nhập tên cách ghép.');
  const all = diskRead('profiles.json', []);
  const index = all.findIndex(item => item.owner === owner && item.name === name);
  const record = { owner, name, id: index < 0 ? crypto.randomUUID() : all[index].id, file: source.file, targetId: target.id, definition: toSave };
  if (index < 0) all.push(record);
  else all[index] = record;
  diskWrite('profiles.json', all);
  return profiles();
}

async function loadProfile(id) {
  const record = profiles().find(item => item.id === id);
  if (!record) throw new Error('Không tìm thấy cấu hình của tài khoản này.');
  const workbook = record.file ? readWorkbook(record.file) : undefined;
  if (record.targetId === 'demo') {
    if (!target?.demo) setupDemo();
    selectedTab(workbook || demoSource, record.definition.sourceTab);
    selectedTab(target, record.definition.targetTab);
    source = workbook || demoSource;
  } else {
    const opened = await google.open(record.targetId);
    selectedTab(workbook, record.definition.sourceTab);
    selectedTab(opened, record.definition.targetTab);
    target = opened;
    source = workbook;
  }
  preview = undefined;
  return { source: sourceInfo(), target, record };
}

const handlers = {
  info,
  demo: setupDemo,
  async chooseFile() {
    const { canceled, filePaths } = await dialog.showOpenDialog(window, { properties: ['openFile'], filters: [{ name: 'Excel và CSV', extensions: ['xlsx', 'xlsm', 'xls', 'xlsb', 'csv', 'tsv'] }] });
    return canceled ? null : handlers.openFile(filePaths[0]);
  },
  openFile(file) {
    source = readWorkbook(file);
    preview = undefined;
    return sourceInfo();
  },
  async configure() {
    const { canceled, filePaths } = await dialog.showOpenDialog(window, { properties: ['openFile'], filters: [{ name: 'Google OAuth Desktop JSON', extensions: ['json'] }] });
    if (canceled) return { ...info(), cancelled: true };
    if (fs.statSync(filePaths[0]).size > 64_000) throw new Error('Tệp OAuth JSON quá lớn.');
    client = clientFromJson(fs.readFileSync(filePaths[0], 'utf8'));
    account = undefined;
    if (!target?.demo) target = undefined;
    preview = undefined;
    saveGoogle();
    return info();
  },
  async login() {
    if (!client) throw new Error('Nhập cấu hình OAuth của app trước.');
    loginAttempt = new AbortController();
    let signedIn;
    try { signedIn = await signIn(client, loginAttempt.signal); }
    finally { loginAttempt = undefined; }
    if (!signedIn) return { ...info(), cancelled: true };
    account = signedIn;
    if (!target?.demo) target = undefined;
    preview = undefined;
    saveGoogle();
    return info();
  },
  cancelLogin() { loginAttempt?.abort(); },
  googleGuide() { return shell.openExternal('https://developers.google.com/workspace/guides/create-credentials#desktop-app'); },
  logout() {
    account = undefined;
    if (!target?.demo) target = undefined;
    preview = undefined;
    saveGoogle();
    return info();
  },
  async openSheet(link) {
    const opened = await google.open(spreadsheetId(link));
    if (source === demoSource) source = undefined;
    target = opened;
    preview = undefined;
    return { source: sourceInfo(), target, profiles: profiles() };
  },
  async grid({ side, id, range: value }) {
    const range = parseRange(value, 3200);
    if (side === 'source') return sourceGrid(selectedTab(source, id), range, true);
    const tab = selectedTab(target, id);
    checkTarget(tab, range);
    return targetGrid(tab, range, true);
  },
  async preview(input) {
    preview = undefined;
    const { plan, targetTab } = await createPlan(input);
    const id = crypto.randomUUID();
    preview = { id, input, digest: fingerprint(plan) };
    return { id, ...plan, changes: plan.changes.slice(0, 30), count: plan.changes.length, tab: targetTab.title };
  },
  async sync(id) {
    if (!preview || preview.id !== id) throw new Error('Bấm Xem trước để tạo lượt đồng bộ mới.');
    const record = preview;
    preview = undefined;
    const { plan, targetTab } = await createPlan(record.input);
    if (record.digest !== fingerprint(plan)) throw new Error('Excel hoặc Sheets đã đổi sau khi xem trước. Hãy xem trước lại.');
    if (plan.changes.length && target.demo) {
      const rows = demoBook[targetTab.id];
      for (const { row, col, value } of plan.changes) {
        rows[row] ||= [];
        rows[row][col] = demoCell(value);
      }
    } else if (plan.changes.length) {
      await google.write(target.id, targetTab, plan.changes);
    }
    const result = { cells: plan.changes.length, rows: plan.rows, demo: Boolean(target.demo) };
    try {
      fs.mkdirSync(stateDir, { recursive: true });
      fs.appendFileSync(path.join(stateDir, 'sync-history.jsonl'), JSON.stringify({ at: new Date().toISOString(), account: target.demo ? 'demo' : account.email, destination: target.id, tab: targetTab.title, ...result }) + '\n');
    } catch { result.logMissing = true; }
    return result;
  },
  saveProfile,
  loadProfile,
};

ipcMain.handle('sync:invoke', async (event, action, argument) => {
  if (event.sender !== window?.webContents || !Object.hasOwn(handlers, action)) return { error: 'Yêu cầu không hợp lệ.' };
  const exclusive = action !== 'grid' && action !== 'cancelLogin' && action !== 'googleGuide';
  if (exclusive && busy) return { error: 'App đang xử lý. Hãy chờ thao tác đang chạy.' };
  if (exclusive) busy = true;
  try { return { value: await handlers[action](argument) }; }
  catch (error) { return { error: error.message }; }
  finally { if (exclusive) busy = false; }
});

if (!app.requestSingleInstanceLock()) app.quit();
else app.whenReady().then(() => {
  stateDir = app.getPath('userData');
  restoreGoogle();
  loadBundledOAuth();
  window = new BrowserWindow({
    width: 1360, height: 900, minWidth: 760, minHeight: 560, autoHideMenuBar: true,
    backgroundColor: '#f2f4f5',
    webPreferences: { preload: path.join(__dirname, 'preload.cjs'), contextIsolation: true, nodeIntegration: false, sandbox: true },
  });
  window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  window.webContents.on('will-navigate', event => event.preventDefault());
  window.loadFile(path.join(__dirname, 'index.html'));
});

app.on('second-instance', () => { if (window?.isMinimized()) window.restore(); window?.focus(); });
app.on('window-all-closed', () => app.quit());
