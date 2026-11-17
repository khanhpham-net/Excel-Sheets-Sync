const MAX_RANGE_CELLS = 30_000;

function columnName(index) {
  let label = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    label = String.fromCharCode(65 + (n - 1) % 26) + label;
  }
  return label;
}

function address(row, column) {
  return `${columnName(column)}${row + 1}`;
}

function parseRange(text, maxCells = MAX_RANGE_CELLS) {
  const match = String(text ?? '').trim().toUpperCase().match(/^\$?([A-Z]{1,3})\$?([1-9]\d{0,6})(?::\$?([A-Z]{1,3})\$?([1-9]\d{0,6}))?$/);
  if (!match) throw new Error('Nhập ô hoặc vùng hợp lệ, ví dụ A1 hoặc A1:D20.');
  const [, c1, r1, c2 = c1, r2 = r1] = match;
  const index = label => [...label].reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0) - 1;
  const top = Number(r1) - 1;
  const left = index(c1);
  const bottom = Number(r2) - 1;
  const right = index(c2);
  if (top > bottom || left > right) throw new Error('Chọn vùng từ góc trên trái đến góc dưới phải.');
  if ((bottom - top + 1) * (right - left + 1) > maxCells) {
    throw new Error(`Vùng chọn vượt ${maxCells.toLocaleString('vi-VN')} ô. Hãy chia thành vùng nhỏ hơn.`);
  }
  return { top, left, bottom, right };
}

function rangeLabel({ top, left, bottom, right }) {
  const first = address(top, left);
  return top === bottom && left === right ? first : `${first}:${address(bottom, right)}`;
}

function sheetRange(title, range) {
  return `'${title.replaceAll("'", "''")}'!${rangeLabel(range)}`;
}

function readCell(grid, row, col) {
  return grid.cells[row - grid.range.top]?.[col - grid.range.left] || {};
}

function contains(outer, inner) {
  return outer.top <= inner.top && outer.left <= inner.left && outer.bottom >= inner.bottom && outer.right >= inner.right;
}

function keyPart(cell) {
  if (cell.error || (cell.formula && cell.value == null)) throw new Error('Có lỗi công thức hoặc thiếu giá trị đã tính trong cột mã.');
  return String(cell.value ?? '').trim();
}

function matchedRows(grid, range, columns, label) {
  const index = new Map();
  for (let row = range.top + 1; row <= range.bottom; row++) {
    const parts = columns.map(col => keyPart(readCell(grid, row, col)));
    if (parts.every(part => part === '')) {
      const occupied = Array.from({ length: range.right - range.left + 1 }, (_, i) => readCell(grid, row, range.left + i))
        .some(cell => (cell.value != null && cell.value !== '') || cell.formula || cell.error);
      if (occupied) throw new Error(`${label} có dữ liệu nhưng thiếu mã ở dòng ${row + 1}. Thu hẹp vùng nếu đó là dòng tổng.`);
      continue;
    }
    if (parts.some(part => !part)) throw new Error(`${label} có mã ghép thiếu thành phần ở dòng ${row + 1}.`);
    const key = JSON.stringify(parts);
    if (index.has(key)) throw new Error(`${label} trùng mã ${parts.join(' / ')} ở dòng ${index.get(key) + 1} và ${row + 1}.`);
    index.set(key, row);
  }
  return index;
}

function vietnameseNumber(value, sourceAddress) {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'string') {
    const text = value.trim().replace(/^'\s*/, '').replace(/\s*(?:₫|đ|VND)\s*$/iu, '')
      .replaceAll(/[\u00a0\u202f]/g, ' ').replace(/^−/, '-');
    const valid = /^[+-]?(?:\d+|\d{1,3}(?:\.\d{3})+|\d{1,3}(?: \d{3})+)(?:,\d+)?$/;
    if (valid.test(text)) {
      const normalized = text.replaceAll(/[. ]/g, '').replace(',', '.');
      const number = Number(normalized);
      const integer = Number(normalized.split('.')[0]);
      if (Number.isFinite(number) && Number.isSafeInteger(integer)) return number;
    }
  }
  throw new Error(`Ô Excel ${sourceAddress} không phải số theo định dạng Việt Nam. Kiểm tra giá trị hoặc bỏ chọn Ghi thành số cho cột này.`);
}

function makePlan(source, target, definition) {
  const { sourceRange, targetRange, mode } = definition;
  if (!contains(source.range, sourceRange) || !contains(target.range, targetRange)) throw new Error('Vùng xem trước không đầy đủ.');
  const changes = [];
  const affectedRows = new Set();
  let skippedEmpty = 0;
  let missing = 0;
  let added = 0;

  function mapCell(fromRow, fromCol, toRow, toCol, asNumber = false) {
    const from = readCell(source, fromRow, fromCol);
    const to = readCell(target, toRow, toCol);
    if (from.error || (from.formula && from.value == null)) {
      throw new Error(`Ô nguồn ${address(fromRow, fromCol)} có lỗi hoặc công thức chưa có kết quả đã lưu. Hãy tính và lưu lại trong Excel.`);
    }
    if (from.value == null || from.value === '') {
      skippedEmpty++;
      return;
    }
    if (to.formula || to.derived || (target.merges || []).some(merge => contains(merge, { top: toRow, bottom: toRow, left: toCol, right: toCol }))) {
      throw new Error(`Ô đích ${address(toRow, toCol)} chứa công thức, kết quả sinh bởi công thức hoặc thuộc ô gộp. Hãy chọn vùng nhập liệu khác.`);
    }
    const value = asNumber ? vietnameseNumber(from.value, address(fromRow, fromCol)) : from.value;
    if (value === to.value) return;
    const display = asNumber ? value.toLocaleString('vi-VN', { maximumFractionDigits: 15 }) : from.display ?? value;
    changes.push({ row: toRow, col: toCol, cell: address(toRow, toCol), before: to.display ?? to.value ?? '', beforeValue: to.value, value, display, asNumber });
    affectedRows.add(toRow);
  }

  if (mode === 'position') {
    if (sourceRange.bottom - sourceRange.top !== targetRange.bottom - targetRange.top || sourceRange.right - sourceRange.left !== targetRange.right - targetRange.left) {
      throw new Error('Vùng Excel và vùng ghi phải có cùng kích thước. Chọn một ô đích để app tự tính vùng ghi.');
    }
    for (let r = sourceRange.top; r <= sourceRange.bottom; r++) {
      for (let c = sourceRange.left; c <= sourceRange.right; c++) {
        mapCell(r, c, targetRange.top + r - sourceRange.top, targetRange.left + c - sourceRange.left);
      }
    }
  } else {
    const mappings = definition.mappings || [];
    const keys = mappings.filter(mapping => mapping.key);
    if (!keys.length) throw new Error('Đánh dấu ít nhất một cặp cột làm khóa đối chiếu.');
    if (new Set(mappings.map(mapping => mapping.to)).size !== mappings.length) throw new Error('Mỗi cột đích chỉ nhận từ một cột nguồn.');
    if (mappings.some(({ from, to }) => !Number.isInteger(from) || !Number.isInteger(to) || from < sourceRange.left || from > sourceRange.right || to < targetRange.left || to > targetRange.right)) {
      throw new Error('Cặp cột phải nằm trong vùng dữ liệu đã chọn.');
    }
    const sourceRows = matchedRows(source, sourceRange, keys.map(key => key.from), 'Excel');
    const targetRows = matchedRows(target, targetRange, keys.map(key => key.to), 'Google Sheets');
    let nextRow = Math.max(targetRange.top, ...targetRows.values()) + 1;
    for (const [key, fromRow] of sourceRows) {
      const existingRow = targetRows.get(key);
      if (existingRow === undefined && !definition.addMissing) {
        missing++;
        continue;
      }
      if (existingRow === undefined && nextRow > targetRange.bottom) {
        throw new Error('Vùng đích hết dòng trống cho mã mới. Hãy mở rộng vùng đích xuống dưới.');
      }
      const toRow = existingRow ?? nextRow++;
      for (const { from, to, key: isKey, asNumber } of mappings) {
        if (!isKey || existingRow === undefined) mapCell(fromRow, from, toRow, to, !isKey && asNumber);
      }
      if (existingRow === undefined) added++;
    }
  }
  return { changes, rows: affectedRows.size, skippedEmpty, missing, added };
}

function googleValue(value) {
  if (typeof value === 'number') return { numberValue: value };
  if (typeof value === 'boolean') return { boolValue: value };
  return { stringValue: String(value) };
}

function writeRequests(sheetId, changes) {
  const sorted = [...changes].sort((a, b) => a.row - b.row || a.col - b.col);
  const requests = [];
  let block;
  for (const change of sorted) {
    const position = block?.start;
    if (!position || position.rowIndex !== change.row || position.columnIndex + block.rows[0].values.length !== change.col) {
      block = {
        start: { sheetId, rowIndex: change.row, columnIndex: change.col },
        rows: [{ values: [] }],
        fields: 'userEnteredValue',
      };
      requests.push({ updateCells: block });
    }
    block.rows[0].values.push({ userEnteredValue: googleValue(change.value) });
  }
  return requests;
}

module.exports = { MAX_RANGE_CELLS, address, columnName, parseRange, rangeLabel, sheetRange, readCell, makePlan, writeRequests };
