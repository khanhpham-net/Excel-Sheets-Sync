const crypto = require('node:crypto');
const http = require('node:http');
const { shell } = require('electron');
const { sheetRange, writeRequests } = require('./sync.cjs');
const { googleFormat } = require('./sheet-format.cjs');

const SCOPES = 'openid email https://www.googleapis.com/auth/spreadsheets';
const TOKEN_URL = 'https://oauth2.googleapis.com/token';

async function request(url, options = {}) {
  const response = await fetch(url, { ...options, signal: AbortSignal.timeout(35_000) });
  const body = await response.json();
  if (!response.ok) {
    if (body.error === 'invalid_grant') throw new Error('Quyền Google đã hết hạn hoặc bị thu hồi. Hãy đăng xuất trong app rồi đăng nhập lại.');
    if (body.error === 'invalid_client') throw new Error('Google từ chối OAuth client. Hãy nhập lại JSON của Desktop app còn hoạt động.');
    if (response.status === 403) throw new Error('Google từ chối quyền truy cập. Hãy kiểm tra email, quyền Editor, vùng được bảo vệ hoặc cài đặt Google Cloud.');
    if (response.status === 404) throw new Error('Không tìm thấy Google Sheets hoặc tài khoản này chưa được chia sẻ file.');
    throw new Error(body.error?.message || body.error_description || `Google trả mã lỗi ${response.status}.`);
  }
  return body;
}

function token(params) {
  return request(TOKEN_URL, { method: 'POST', body: new URLSearchParams(params) });
}

function clientFromJson(text) {
  let client;
  try { client = JSON.parse(text.replace(/^\uFEFF/, ''))?.installed; }
  catch { throw new Error('Tệp cấu hình không phải JSON hợp lệ. Tải lại tệp từ Google Cloud.'); }
  if (typeof client?.client_id !== 'string' || !client.client_id.endsWith('.apps.googleusercontent.com') || typeof client.client_secret !== 'string' || !client.client_secret.trim()) {
    throw new Error('Chọn tệp JSON của OAuth client loại Desktop app do Google Cloud tải xuống.');
  }
  return { client_id: client.client_id, client_secret: client.client_secret };
}

async function signIn(client, signal) {
  const verifier = crypto.randomBytes(48).toString('base64url');
  const state = crypto.randomBytes(24).toString('base64url');
  let finish;
  const server = http.createServer((req, res) => {
    const result = new URL(req.url, 'http://127.0.0.1');
    if (result.pathname !== '/oauth') { res.writeHead(404).end(); return; }
    if (result.searchParams.get('state') !== state) { res.writeHead(400).end('Invalid state'); return; }
    const code = result.searchParams.get('code');
    res.writeHead(code ? 200 : 400, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end(code ? 'Đã nhận phản hồi Google. Quay lại app để kiểm tra kết nối.' : 'Đăng nhập đã bị từ chối. Quay lại app để thử lại.');
    finish?.({ code });
  });
  await new Promise((resolve, reject) => server.once('error', reject).listen(0, '127.0.0.1', resolve));
  const redirect_uri = `http://127.0.0.1:${server.address().port}/oauth`;
  const params = new URLSearchParams({
    client_id: client.client_id, redirect_uri, scope: SCOPES, state,
    response_type: 'code', access_type: 'offline', prompt: 'consent select_account',
    code_challenge: crypto.createHash('sha256').update(verifier).digest('base64url'), code_challenge_method: 'S256',
  });
  let timer;
  const cancel = () => finish?.({ cancelled: true });
  try {
    const pending = new Promise(resolve => {
      finish = resolve;
      timer = setTimeout(() => resolve({ timedOut: true }), 180_000);
    });
    signal?.addEventListener('abort', cancel, { once: true });
    if (signal?.aborted) cancel();
    else await shell.openExternal(`https://accounts.google.com/o/oauth2/v2/auth?${params}`);
    const { code, cancelled, timedOut } = await pending;
    if (cancelled) return null;
    if (timedOut) throw new Error('Đã hết thời gian đăng nhập. Kiểm tra trang Google trong trình duyệt rồi thử lại.');
    if (!code) throw new Error('Chưa cấp quyền Google. Hãy thử đăng nhập lại và đồng ý quyền Google Sheets.');
    const response = await token({ ...client, code, redirect_uri, code_verifier: verifier, grant_type: 'authorization_code' });
    if (signal?.aborted) return null;
    if (!response.scope?.includes('https://www.googleapis.com/auth/spreadsheets')) throw new Error('Hãy cấp quyền Google Sheets để kết nối.');
    const profile = await request('https://openidconnect.googleapis.com/v1/userinfo', {
      headers: { Authorization: `Bearer ${response.access_token}` },
    });
    if (signal?.aborted) return null;
    if (!profile.sub || !profile.email || !response.refresh_token) throw new Error('Google chưa cấp đủ thông tin đăng nhập.');
    return { sub: profile.sub, email: profile.email, refresh: response.refresh_token };
  } finally {
    clearTimeout(timer);
    signal?.removeEventListener('abort', cancel);
    server.close();
  }
}

function spreadsheetId(link) {
  let url;
  try { url = new URL(String(link).trim()); } catch { throw new Error('Dán link Google Sheets đầy đủ.'); }
  const id = url.pathname.match(/^\/spreadsheets\/(?:u\/\d+\/)?d\/([\w-]{12,})(?:\/|$)/)?.[1];
  if (url.origin !== 'https://docs.google.com' || !id || id === 'e') throw new Error('Dùng link Google Sheets dạng docs.google.com/spreadsheets/d/...');
  return id;
}

function createGoogle(getAccount, getClient) {
  let cached;
  async function call(id, params, body) {
    const account = getAccount();
    if (!account || !getClient()) throw new Error('Đăng nhập Google trước.');
    if (!cached || cached.account !== account || cached.expires < Date.now()) {
      const response = await token({ ...getClient(), refresh_token: account.refresh, grant_type: 'refresh_token' });
      cached = { account, expires: Date.now() + Math.max(0, response.expires_in - 60) * 1000, access: response.access_token };
    }
    if (getAccount() !== account) throw new Error('Tài khoản đã đổi. Hãy mở lại bảng.');
    const headers = { Authorization: `Bearer ${cached.access}` };
    const endpoint = `https://sheets.googleapis.com/v4/spreadsheets/${id}`;
    if (body) {
      return request(`${endpoint}:batchUpdate`, { method: 'POST', headers: { ...headers, 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
    }
    return request(`${endpoint}?${new URLSearchParams(params)}`, { headers });
  }
  return {
    async open(id) {
      const book = await call(id, { fields: 'properties(title),sheets(properties(sheetId,title,sheetType,gridProperties(rowCount,columnCount)))' });
      const tabs = (book.sheets || []).filter(sheet => sheet.properties.sheetType === 'GRID')
        .map(({ properties: p }) => ({ id: p.sheetId, title: p.title, rows: p.gridProperties.rowCount, cols: p.gridProperties.columnCount }));
      if (!tabs.length) throw new Error('File chưa có tab dạng bảng để đọc.');
      return { id, title: book.properties.title, tabs };
    },
    async grid(id, tab, range, styled = false) {
      const fields = styled
        ? 'properties(spreadsheetTheme(themeColors)),sheets(properties(sheetId,gridProperties(hideGridlines)),merges,data(startRow,startColumn,rowMetadata(pixelSize,hiddenByUser,hiddenByFilter),columnMetadata(pixelSize,hiddenByUser),rowData(values(userEnteredValue,effectiveValue,formattedValue,effectiveFormat))))'
        : 'sheets(properties(sheetId),merges,data(startRow,startColumn,rowData(values(userEnteredValue,effectiveValue,formattedValue))))';
      const book = await call(id, {
        ranges: sheetRange(tab.title, range),
        fields,
      });
      const sheet = book.sheets?.find(sheet => sheet.properties.sheetId === tab.id);
      if (!sheet) throw new Error('Tab Google Sheets đã đổi. Hãy mở lại file.');
      const data = sheet.data?.[0];
      const theme = Object.fromEntries((book.properties?.spreadsheetTheme?.themeColors || []).map(color => [color.colorType, color.color?.rgbColor]));
      const cells = Array.from({ length: range.bottom - range.top + 1 }, () => []);
      (data?.rowData || []).forEach((row, i) => {
        const out = cells[(data.startRow || 0) + i - range.top];
        if (!out) return;
        (row.values || []).forEach((cell, j) => {
          const entered = cell.userEnteredValue || {};
          const effective = cell.effectiveValue || {};
          const index = (data.startColumn || 0) + j - range.left;
          if (index < 0 || index > range.right - range.left) return;
          out[index] = {
            value: entered.numberValue ?? entered.stringValue ?? entered.boolValue ?? effective.numberValue ?? effective.stringValue ?? effective.boolValue,
            display: cell.formattedValue,
            formula: entered.formulaValue,
            error: Boolean(effective.errorValue),
            derived: !cell.userEnteredValue && Boolean(cell.effectiveValue),
            format: styled ? googleFormat(cell.effectiveFormat, theme) : undefined,
          };
        });
      });
      const merges = (sheet.merges || []).map(merge => ({
        top: merge.startRowIndex || 0, left: merge.startColumnIndex || 0,
        bottom: merge.endRowIndex - 1, right: merge.endColumnIndex - 1,
      }));
      if (!styled) return { range, cells, merges };
      const pixels = dimensions => (dimensions || []).map(item => item.hiddenByUser || item.hiddenByFilter ? 0 : item.pixelSize);
      return { range, cells, merges, layout: { rows: pixels(data?.rowMetadata), columns: pixels(data?.columnMetadata), showGridlines: !sheet.properties.gridProperties?.hideGridlines } };
    },
    async write(id, tab, changes) {
      const requests = writeRequests(tab.id, changes);
      const body = { requests };
      if (Buffer.byteLength(JSON.stringify(body)) > 2_000_000) throw new Error('Lượt ghi quá lớn. Hãy chia vùng thành nhiều lượt.');
      await call(id, null, body);
    },
  };
}

module.exports = { clientFromJson, signIn, spreadsheetId, createGoogle };
