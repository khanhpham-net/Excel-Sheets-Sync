const $ = id => document.getElementById(id);
const invoke = (method, value) => window.desktop.invoke(method, value);
const mode = $('mode');
const googleDialog = $('google-dialog');
const panes = {};
const COLUMN_HEADER = 26;
const ROW_HEADER = 46;
const GRID_READ_LIMIT = 3200;
let mappings = [];
let pickedColumn;
let preview;
let savedHeaders;
let account = {};
let busy = false;
let focusedPane;

function columnName(index) {
  let result = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) result = String.fromCharCode(65 + (n - 1) % 26) + result;
  return result;
}

const address = (row, col) => `${columnName(col)}${row + 1}`;

function selection(from, to = from) {
  const first = address(Math.min(from.row, to.row), Math.min(from.col, to.col));
  const last = address(Math.max(from.row, to.row), Math.max(from.col, to.col));
  return first === last ? first : `${first}:${last}`;
}

function element(tag, text, className) {
  const node = document.createElement(tag);
  if (text !== undefined) node.textContent = text;
  if (className) node.className = className;
  return node;
}

function button(text, action) {
  const node = element('button', text);
  node.type = 'button';
  node.addEventListener('click', action);
  return node;
}

function showStatus(text, error = false) {
  $('status').textContent = text;
  $('status').classList.toggle('error', error);
}

function controls() {
  $('preview').disabled = busy || !panes.source.book || !panes.target.book;
  $('write').disabled = busy || !preview?.count;
  $('save').disabled = busy || !panes.source.book || !panes.target.book;
  $('load').disabled = busy || !$('profile').value;
  $('login').disabled = busy;
}

async function run(action) {
  if (busy) return;
  busy = true;
  controls();
  document.body.setAttribute('aria-busy', 'true');
  $('workspace').inert = true;
  document.querySelector('.top-controls').inert = true;
  try { await action(); }
  catch (error) { showStatus(error.message, true); }
  finally {
    busy = false;
    $('workspace').inert = false;
    document.querySelector('.top-controls').inert = false;
    document.body.setAttribute('aria-busy', 'false');
    controls();
  }
}

function invalidate() {
  preview = undefined;
  $('preview-panel').hidden = true;
  controls();
}

function definition() {
  return {
    sourceTab: panes.source.tab()?.id, targetTab: panes.target.tab()?.id,
    sourceRange: panes.source.range.value, targetRange: panes.target.range.value,
    mode: mode.value, mappings, addMissing: $('add-missing').checked, headers: mode.value === 'key' ? savedHeaders : undefined,
  };
}

function setProfiles(profiles) {
  const select = $('profile');
  const previous = select.value;
  select.replaceChildren(new Option('Chọn cách ghép', ''), ...profiles.map(record => new Option(record.targetId === 'demo' ? `${record.name} · Mẫu` : record.name, record.id)));
  if (profiles.some(record => record.id === previous)) select.value = previous;
  controls();
}

function setAccount(info) {
  account = info;
  $('account').textContent = info.email || 'Chưa đăng nhập';
  $('logout').hidden = !info.email;
  $('configure').hidden = Boolean(info.email);
  $('login').hidden = Boolean(info.email);
  setProfiles(info.profiles);
}

function showPairs() {
  const root = $('pairs');
  root.replaceChildren();
  $('mapping-help').textContent = pickedColumn === undefined
    ? 'Hai vùng gồm hàng tiêu đề. Kéo chữ cột sang đích, hoặc bấm chữ cột ở hai bên.'
    : `Đã chọn cột Excel ${columnName(pickedColumn)}. Bấm chữ cột bên Sheets để ghép.`;
  for (const [index, pair] of mappings.entries()) {
    const label = `${columnName(pair.from)} → ${columnName(pair.to)}`;
    const item = element('div', undefined, 'pair');
    const keyLabel = element('label');
    const key = element('input');
    key.type = 'checkbox';
    key.checked = pair.key;
    key.setAttribute('aria-label', `Dùng ${label} làm khóa đối chiếu`);
    keyLabel.append(key, ' Khóa');
    const numberLabel = element('label');
    const number = element('input');
    number.type = 'checkbox';
    number.checked = Boolean(pair.asNumber && !pair.key);
    number.setAttribute('aria-label', `Ghi thành số ${label}`);
    numberLabel.title = 'Ví dụ: 22.000.000 ₫ được ghi thành số để tính tổng';
    numberLabel.hidden = pair.key;
    numberLabel.append(number, ' Ghi thành số');
    number.addEventListener('change', () => { pair.asNumber = number.checked; invalidate(); });
    key.addEventListener('change', () => {
      pair.key = key.checked;
      if (key.checked) { pair.asNumber = false; number.checked = false; }
      numberLabel.hidden = key.checked;
      invalidate();
    });
    const remove = button('×', () => { mappings.splice(index, 1); invalidate(); showPairs(); });
    remove.addEventListener('click', () => { savedHeaders = undefined; });
    remove.setAttribute('aria-label', `Xóa cặp ${label}`);
    item.append(element('strong', label), keyLabel, numberLabel, remove);
    root.append(item);
  }
}

function pairColumns(from, to) {
  if (!Number.isInteger(from) || !Number.isInteger(to)) return;
  const index = mappings.findIndex(pair => pair.to === to);
  const newPair = { from, to, key: false };
  if (index < 0) mappings.push(newPair);
  else if (mappings[index].from !== from) mappings[index] = newPair;
  savedHeaders = undefined;
  pickedColumn = undefined;
  invalidate();
  showPairs();
}

function clearPairs() {
  pickedColumn = undefined;
  mappings = [];
  savedHeaders = undefined;
  invalidate();
  showPairs();
}

function modeChanged() {
  $('mapping-detail').hidden = mode.value !== 'key';
  $('position-help').hidden = mode.value !== 'position';
  $('add-option').hidden = mode.value !== 'key';
  panes.source.rangeLink.hidden = mode.value !== 'position' || !panes.source.book;
  invalidate();
}

function dropOn(node, mime, callback) {
  node.addEventListener('dragover', event => {
    if (!event.dataTransfer.types.includes(mime)) return;
    event.preventDefault();
    node.classList.add('drop-hover');
  });
  node.addEventListener('dragleave', () => node.classList.remove('drop-hover'));
  node.addEventListener('drop', event => {
    if (!event.dataTransfer.types.includes(mime)) return;
    event.preventDefault();
    event.stopPropagation();
    node.classList.remove('drop-hover');
    callback(event.dataTransfer);
  });
}

function gridAxis(count, initialSize, minimum, maximum) {
  const offsets = new Float64Array(count + 1);
  const sizes = new Map();
  let defaultSize = initialSize;
  function rebuild() {
    for (let i = 0; i < count; i++) offsets[i + 1] = offsets[i] + (sizes.get(i) ?? defaultSize);
  }
  rebuild();
  return {
    offsets,
    size: i => offsets[i + 1] - offsets[i],
    index(pixel) {
      let low = 0;
      let high = count;
      while (low < high) {
        const mid = Math.floor((low + high) / 2);
        if (offsets[mid + 1] <= pixel) low = mid + 1;
        else high = mid;
      }
      return Math.min(low, count - 1);
    },
    update(start, values = [], preferred = defaultSize) {
      const standard = Math.round(Math.min(maximum, Math.max(minimum, preferred)));
      let changed = standard !== defaultSize;
      defaultSize = standard;
      values.forEach((value, i) => {
        if (!Number.isFinite(value) || value < 0 || i + start >= count) return;
        const size = value === 0 ? 0 : Math.round(Math.min(maximum, Math.max(minimum, value)));
        if (sizes.get(start + i) !== size) { sizes.set(start + i, size); changed = true; }
      });
      if (changed) rebuild();
      return changed;
    },
  };
}

function createPane(side) {
  const node = $('pane-template').content.firstElementChild.cloneNode(true);
  const get = selector => node.querySelector(selector);
  const pane = { side, node, select: get('.tab'), range: get('.range'), rangeLink: get('.range-link') };
  const isSource = side === 'source';
  const grid = get('.grid');
  const entry = get('.pane-entry');
  const jump = get('.jump-row');
  const gridContent = element('div', undefined, 'grid-content');
  const focus = get('.pane-focus');
  const miniOpen = get('.mini-open');
  const paneName = isSource ? 'Excel' : 'Google Sheets';
  focus.addEventListener('click', () => setPaneFocus(focusedPane === side ? undefined : side));
  miniOpen.addEventListener('click', () => {
    setPaneFocus(side);
    focus.focus({ preventScroll: true });
  });
  miniOpen.setAttribute('aria-label', `Mở lớn bản xem ${paneName}`);
  pane.setFocus = sideToFocus => {
    const expanded = sideToFocus === side;
    const miniature = Boolean(sideToFocus && !expanded);
    node.classList.toggle('is-focused', expanded);
    node.classList.toggle('is-mini', miniature);
    focus.textContent = expanded ? 'Chia đôi' : 'Phóng to';
    focus.setAttribute('aria-label', expanded ? `Trở về chia đôi từ ${paneName}` : `Phóng to ${paneName}`);
    focus.setAttribute('aria-pressed', String(expanded));
    grid.inert = miniature;
    miniOpen.hidden = !miniature;
  };
  pane.setFocus(focusedPane);
  const layers = {
    cells: element('div', undefined, 'cell-layer'),
    columns: element('div', undefined, 'column-layer'),
    rows: element('div', undefined, 'row-layer'),
    corner: element('div', undefined, 'grid-corner'),
  };
  let axes;
  let axesTab;
  let view;
  let generation = 0;
  let loading;
  let gridError;
  get('.pane-label').textContent = isSource ? 'EXCEL NGUỒN' : 'GOOGLE SHEETS ĐÍCH';
  pane.select.setAttribute('aria-label', isSource ? 'Tab Excel' : 'Tab Google Sheets');
  pane.range.id = `${side}-range`;
  get('.range-label').htmlFor = pane.range.id;
  pane.rangeLink.hidden = true;
  pane.tab = () => pane.book?.tabs[Number(pane.select.value)];

  pane.setBook = (book, tabId, range) => {
    pane.book = book;
    get('.filename').textContent = book?.title || (isSource ? 'Chưa mở Excel' : 'Chưa mở Google Sheets');
    const tabs = book?.tabs || [];
    pane.select.replaceChildren(...tabs.map((tab, index) => new Option(tab.title, index)));
    const index = tabs.findIndex(tab => tab.id === tabId);
    if (index >= 0) pane.select.value = index;
    pane.select.disabled = !book;
    pane.range.disabled = !book;
    pane.range.value = range ?? pane.defaultRange();
    pane.rangeLink.hidden = !isSource || mode.value !== 'position' || !pane.book;
    if (book && !isSource && book.id !== 'demo') pane.link.value = `https://docs.google.com/spreadsheets/d/${book.id}/edit`;
    jump.value = '';
    grid.scrollTop = 0;
    grid.scrollLeft = 0;
  };

  pane.defaultRange = () => {
    const tab = pane.tab();
    if (!tab) return '';
    return `A1:${columnName(Math.min(8, tab.cols) - 1)}${isSource ? Math.min(tab.rows, 12) : Math.min(tab.rows, 20)}`;
  };

  function positionChrome() {
    layers.columns.style.top = `${grid.scrollTop}px`;
    layers.rows.style.left = `${grid.scrollLeft}px`;
    layers.corner.style.top = `${grid.scrollTop}px`;
    layers.corner.style.left = `${grid.scrollLeft}px`;
    if (gridError) {
      gridError.node.style.top = `${grid.scrollTop + COLUMN_HEADER}px`;
      gridError.node.style.left = `${grid.scrollLeft + ROW_HEADER}px`;
      gridError.node.style.width = `${Math.max(130, grid.clientWidth - ROW_HEADER)}px`;
    }
  }

  function sizeContent() {
    gridContent.style.width = `${axes.cols.offsets.at(-1) + ROW_HEADER}px`;
    gridContent.style.height = `${axes.rows.offsets.at(-1) + COLUMN_HEADER}px`;
  }

  async function renderVisible() {
    const tab = pane.tab();
    if (!tab || !axes || !gridContent.isConnected) return;
    const { rows, cols } = axes;
    positionChrome();
    const top = rows.index(grid.scrollTop);
    const bottom = rows.index(grid.scrollTop + Math.max(0, grid.clientHeight - COLUMN_HEADER - 1));
    const left = cols.index(grid.scrollLeft);
    const right = cols.index(grid.scrollLeft + Math.max(0, grid.clientWidth - ROW_HEADER - 1));
    const current = selection({ row: top, col: left }, { row: bottom, col: right });
    get('.window-label').textContent = `${current} · ${tab.rows} dòng × ${tab.cols} cột`;
    if (view && top >= view.top && bottom <= view.bottom && left >= view.left && right <= view.right) {
      gridError?.node.remove();
      gridError = undefined;
      return;
    }
    if (loading === generation) return;
    const maxColumns = Math.max(1, Math.floor(GRID_READ_LIMIT / (bottom - top + 1)));
    const extraColumns = Math.max(0, Math.min(22, maxColumns - (right - left + 1)));
    const startCol = Math.max(0, left - Math.min(16, extraColumns));
    const endCol = Math.min(tab.cols - 1, right + extraColumns - (left - startCol));
    const rowBudget = Math.floor(GRID_READ_LIMIT / (endCol - startCol + 1));
    const spare = Math.max(0, Math.min(32, rowBudget - (bottom - top + 1)));
    const first = Math.max(0, top - Math.floor(spare / 2));
    const last = Math.min(tab.rows - 1, bottom + Math.ceil(spare / 2));
    const region = selection({ row: first, col: startCol }, { row: last, col: endCol });
    if (gridError) {
      if (gridError.region === region) return;
      gridError.node.remove();
      gridError = undefined;
    }
    const request = generation;
    loading = request;
    let data;
    try { data = await invoke('grid', { side, id: tab.id, range: region }); }
    catch (error) {
      if (request !== generation) return;
      const notice = element('div', undefined, 'grid-error');
      notice.append(element('p', `Không tải được ô trong tab ${tab.title}. ${error.message}`), button('Thử tải lại', () => {
        notice.remove();
        gridError = undefined;
        run(renderVisible);
      }));
      gridContent.append(notice);
      gridError = { node: notice, region };
      positionChrome();
      throw error;
    }
    finally { if (loading === request) loading = undefined; }
    if (request !== generation) return;
    const rowAnchor = rows.index(grid.scrollTop);
    const colAnchor = cols.index(grid.scrollLeft);
    const y = grid.scrollTop - rows.offsets[rowAnchor];
    const x = grid.scrollLeft - cols.offsets[colAnchor];
    const heightChanged = rows.update(data.range.top, data.layout?.rows, data.layout?.rowSize);
    const widthChanged = cols.update(data.range.left, data.layout?.columns, data.layout?.columnSize);
    if (heightChanged || widthChanged) {
      sizeContent();
      if (heightChanged) grid.scrollTop = rows.offsets[rowAnchor] + Math.min(y, Math.max(0, rows.size(rowAnchor) - 1));
      if (widthChanged) grid.scrollLeft = cols.offsets[colAnchor] + Math.min(x, Math.max(0, cols.size(colAnchor) - 1));
    }
    view = data.range;
    drawGrid(layers, axes, pane, data, isSource);
    await renderVisible();
  }

  pane.draw = async () => {
    generation++;
    view = undefined;
    gridError = undefined;
    for (const layer of Object.values(layers)) layer.replaceChildren();
    gridContent.replaceChildren(...Object.values(layers));
    const tab = pane.tab();
    grid.classList.toggle('loaded', Boolean(tab));
    jump.disabled = !tab;
    if (!tab) {
      grid.replaceChildren(element('p', isSource ? 'Kéo file Excel vào đây hoặc bấm Chọn file.' : 'Dán link sau khi đăng nhập Google.'));
      get('.window-label').textContent = '';
      axes = undefined;
      return;
    }
    jump.max = tab.rows;
    if (!axes || axesTab !== tab) {
      axes = { rows: gridAxis(tab.rows, isSource ? 20 : 21, 16, 1000), cols: gridAxis(tab.cols, isSource ? 85 : 100, 26, 1600) };
      axesTab = tab;
    }
    sizeContent();
    if (!gridContent.isConnected) grid.replaceChildren(gridContent);
    await renderVisible();
  };

  let scrollFrame;
  grid.addEventListener('scroll', () => {
    if (scrollFrame) return;
    scrollFrame = requestAnimationFrame(() => {
      scrollFrame = undefined;
      renderVisible().catch(error => showStatus(error.message, true));
    });
  });
  new ResizeObserver(() => renderVisible().catch(error => showStatus(error.message, true))).observe(grid);
  jump.addEventListener('keydown', event => {
    if (event.key !== 'Enter' || !jump.checkValidity() || !jump.value) return;
    event.preventDefault();
    grid.scrollTop = axes.rows.offsets[jump.valueAsNumber - 1];
    grid.focus({ preventScroll: true });
  });

  pane.select.addEventListener('change', () => run(async () => {
    grid.scrollTop = 0;
    grid.scrollLeft = 0;
    jump.value = '';
    pane.range.value = pane.defaultRange();
    clearPairs();
    await pane.draw();
  }));
  pane.range.addEventListener('input', invalidate);
  pane.range.addEventListener('input', () => markSelection(layers.cells, pane.range.value, axes));
  connectGridPointer(grid, () => axes, (from, to = from) => {
    const value = selection(from, to);
    if (pane.range.value === value) return;
    pane.range.value = value;
    markSelection(layers.cells, value, axes);
    invalidate();
  }, isSource);
  pane.setRange = cell => {
    pane.range.value = address(cell.row, cell.col);
    markSelection(layers.cells, pane.range.value, axes);
    invalidate();
  };

  if (isSource) {
    entry.append(button('Chọn file Excel', () => run(async () => loadExcel(await invoke('chooseFile')))), element('span', 'hoặc kéo vào đây'));
    dropOn(node, 'Files', files => run(async () => {
      if (files.files.length !== 1) throw new Error('Kéo một file Excel mỗi lần.');
      await loadExcel(await invoke('openFile', window.desktop.filePath(files.files[0])));
    }));
    pane.rangeLink.addEventListener('dragstart', event => event.dataTransfer.setData('application/x-sync-range', pane.range.value));
  } else {
    pane.link = element('input');
    pane.link.placeholder = 'Dán link Google Sheets';
    pane.link.setAttribute('aria-label', 'Link Google Sheets');
    const open = () => run(async () => {
      if (!account.email) {
        showStatus('Đăng nhập Google rồi bấm Mở Sheets lại.');
        if (!account.configured) showGoogleSetup();
        else await logIn();
        return;
      }
      const result = await invoke('openSheet', pane.link.value);
      if (!result.source) { panes.source.setBook(undefined); await panes.source.draw(); }
      pane.setBook(result.target);
      clearPairs();
      setProfiles(result.profiles);
      await pane.draw();
      showStatus('Chọn tab và vùng đích.');
    });
    entry.append(pane.link, button('Mở Sheets', open));
    pane.link.addEventListener('keydown', event => { if (event.key === 'Enter') open(); });
  }
  $('panes').append(node);
  return pane;
}

function setPaneFocus(side) {
  focusedPane = side;
  if (side) $('panes').dataset.focus = side;
  else delete $('panes').dataset.focus;
  for (const pane of Object.values(panes)) pane.setFocus(side);
}

function markSelection(grid, value, axes) {
  const match = value.trim().toUpperCase().replaceAll('$', '').match(/^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/);
  const index = label => [...label].reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0) - 1;
  const bounds = match && { top: Number(match[2]) - 1, bottom: Number(match[4] || match[2]) - 1, left: index(match[1]), right: index(match[3] || match[1]) };
  const existing = grid.querySelector('.range-highlight');
  if (!bounds || !axes) { existing?.remove(); return; }
  const left = axes.cols.offsets[bounds.left];
  const top = axes.rows.offsets[bounds.top];
  const right = axes.cols.offsets[Math.min(bounds.right + 1, axes.cols.offsets.length - 1)];
  const bottom = axes.rows.offsets[Math.min(bounds.bottom + 1, axes.rows.offsets.length - 1)];
  if (right > left && bottom > top) grid.append(place(existing || element('div', undefined, 'range-highlight'), left, top, right - left, bottom - top));
  else existing?.remove();
}

function connectGridPointer(grid, getAxes, select, isSource) {
  let interaction;
  let point;
  let frame;
  let spaceHeld = false;

  const isEditable = event => event.target instanceof Element && event.target.closest('input, select, button, textarea, [contenteditable]');
  window.addEventListener('keydown', event => {
    if (event.code !== 'Space' || isEditable(event)) return;
    spaceHeld = true;
    if (grid.matches(':hover')) {
      event.preventDefault();
      grid.classList.add('pan-ready');
    }
  });
  window.addEventListener('keyup', event => {
    if (event.code !== 'Space') return;
    spaceHeld = false;
    grid.classList.remove('pan-ready');
  });

  function viewport() {
    const rect = grid.getBoundingClientRect();
    return { left: rect.left + ROW_HEADER, top: rect.top + COLUMN_HEADER, right: rect.left + grid.clientWidth, bottom: rect.top + grid.clientHeight };
  }

  function updateSelection(bounds) {
    const axes = getAxes();
    if (!axes || !interaction?.anchor) return;
    const x = Math.max(0, Math.min(bounds.right - bounds.left - 1, point.x - bounds.left));
    const y = Math.max(0, Math.min(bounds.bottom - bounds.top - 1, point.y - bounds.top));
    select(interaction.anchor, { row: axes.rows.index(grid.scrollTop + y), col: axes.cols.index(grid.scrollLeft + x) });
  }

  function step() {
    frame = undefined;
    if (!point || (interaction && !interaction.anchor)) return;
    const bounds = viewport();
    const speed = (value, start, end) => {
      const distance = value < start + 22 ? value - start - 22 : value > end - 22 ? value - end + 22 : 0;
      return Math.max(-44, Math.min(44, Math.round(distance * 0.47)));
    };
    const dx = speed(point.x, bounds.left, bounds.right);
    const dy = speed(point.y, bounds.top, bounds.bottom);
    if (dx) grid.scrollLeft += dx;
    if (dy) grid.scrollTop += dy;
    updateSelection(bounds);
    if (dx || dy) frame = requestAnimationFrame(step);
  }

  function moveTo(event) {
    point = { x: event.clientX, y: event.clientY };
    if (!frame) frame = requestAnimationFrame(step);
  }

  function stop() {
    if (frame) cancelAnimationFrame(frame);
    frame = point = interaction = undefined;
    grid.classList.remove('panning');
  }

  grid.addEventListener('pointerdown', event => {
    if (busy || !getAxes()) return;
    const cell = event.target.closest('.grid-cell');
    const pan = event.button === 1 || (event.button === 0 && spaceHeld && !isEditable(event));
    if (!pan && (event.button !== 0 || !cell)) return;
    event.preventDefault();
    grid.focus({ preventScroll: true });
    grid.setPointerCapture(event.pointerId);
    interaction = pan
      ? { pan, x: event.clientX, y: event.clientY, left: grid.scrollLeft, top: grid.scrollTop }
      : { anchor: { row: Number(cell.dataset.row), col: Number(cell.dataset.col) } };
    if (pan) grid.classList.add('panning');
    else select(interaction.anchor);
  });
  grid.addEventListener('pointermove', event => {
    if (!interaction) return;
    if (interaction.pan) {
      grid.scrollLeft = interaction.left + interaction.x - event.clientX;
      grid.scrollTop = interaction.top + interaction.y - event.clientY;
    } else moveTo(event);
  });
  grid.addEventListener('pointerup', event => {
    if (interaction?.anchor && point) {
      point = { x: event.clientX, y: event.clientY };
      updateSelection(viewport());
    }
    stop();
  });
  grid.addEventListener('lostpointercapture', stop);
  grid.addEventListener('pointercancel', stop);
  window.addEventListener('blur', () => {
    spaceHeld = false;
    grid.classList.remove('pan-ready');
    stop();
  });
  grid.addEventListener('auxclick', event => { if (event.button === 1) event.preventDefault(); });
  grid.addEventListener('wheel', event => {
    if (!event.shiftKey || event.deltaX || !event.deltaY) return;
    event.preventDefault();
    const scale = event.deltaMode === 1 ? 20 : event.deltaMode === 2 ? grid.clientWidth : 1;
    grid.scrollLeft += event.deltaY * scale;
  }, { passive: false });

  if (!isSource) {
    let exitTimer;
    grid.addEventListener('dragover', event => {
      if (!event.dataTransfer.types.some(type => type === 'application/x-sync-range' || type === 'application/x-sync-column')) return;
      clearTimeout(exitTimer);
      event.preventDefault();
      moveTo(event);
    });
    grid.addEventListener('dragleave', () => {
      clearTimeout(exitTimer);
      exitTimer = setTimeout(stop, 70);
    });
    window.addEventListener('drop', stop, true);
    window.addEventListener('dragend', stop, true);
  }
}

function place(node, x, y, width, height) {
  Object.assign(node.style, { left: x + 'px', top: y + 'px', width: width + 'px', height: height + 'px' });
  return node;
}

function drawGrid(layers, axes, pane, data, isSource) {
  const { rows, cols } = axes;
  const { range, cells, merges = [], layout } = data;
  const columnLabels = document.createDocumentFragment();
  const rowLabels = document.createDocumentFragment();
  const body = document.createDocumentFragment();
  layers.cells.classList.toggle('no-gridlines', layout?.showGridlines === false);
  const cellAt = (row, col) => cells[row - range.top]?.[col - range.left] || {};
  const mergeCells = new Set();

  for (const merge of merges) {
    const top = Math.max(range.top, merge.top);
    const bottom = Math.min(range.bottom, merge.bottom);
    const left = Math.max(range.left, merge.left);
    const right = Math.min(range.right, merge.right);
    for (let row = top; row <= bottom; row++) {
      for (let col = left; col <= right; col++) mergeCells.add(row * cols.offsets.length + col);
    }
  }

  function drawCell(row, col, item, lastRow = row, lastCol = col, merged = false) {
    const width = cols.offsets[lastCol + 1] - cols.offsets[col];
    const height = rows.offsets[lastRow + 1] - rows.offsets[row];
    if (width === 0 || height === 0) return;
    const cell = place(element('div', undefined, 'grid-cell'), cols.offsets[col], rows.offsets[row], width, height);
    cell.dataset.row = row;
    cell.dataset.col = col;
    const label = item.display ?? item.value ?? '';
    const format = item.format || {};
    const text = element('span', label, 'cell-text');
    const horizontal = ['left', 'center', 'right'].includes(format.horizontal) ? format.horizontal : typeof item.value === 'number' ? 'right' : 'left';
    cell.title = item.formula || String(label);
    cell.style.textAlign = horizontal;
    if (merged) cell.classList.add('merged');
    if (format.background) cell.style.backgroundColor = format.background;
    if (format.color) cell.style.color = format.color;
    if (format.font) cell.style.fontFamily = '"' + format.font.replace(/["\\\r\n]/g, '') + '", Arial, sans-serif';
    if (Number.isFinite(format.size) && format.size > 0) cell.style.fontSize = format.size + 'pt';
    if (format.bold) cell.style.fontWeight = '700';
    if (format.italic) cell.style.fontStyle = 'italic';
    if (format.underline || format.strike) text.style.textDecoration = [format.underline && 'underline', format.strike && 'line-through'].filter(Boolean).join(' ');
    if (format.vertical === 'top') cell.style.alignItems = 'flex-start';
    if (format.vertical === 'bottom') cell.style.alignItems = 'flex-end';
    if (format.wrap) text.classList.add('wrap');
    if (format.overflow === false && !format.wrap) text.classList.add('clip');
    for (const [side, border] of Object.entries(format.borders || {})) {
      if (border) cell.style.setProperty('border-' + side, border);
    }
    if (!merged && !format.wrap && format.overflow !== false && label && typeof item.value === 'string') {
      const isEmpty = otherCol => {
        if (!cols.size(otherCol)) return true;
        const next = cellAt(row, otherCol);
        return (next.value == null || next.value === '') && !next.display && !next.formula && !next.error && !next.derived && !mergeCells.has(row * cols.offsets.length + otherCol);
      };
      let first = col;
      let last = col;
      if (horizontal !== 'left') while (first > range.left && isEmpty(first - 1)) first--;
      if (horizontal !== 'right') while (last < range.right && isEmpty(last + 1)) last++;
      if (first < col || last > col) {
        const offset = cols.offsets[col] - cols.offsets[first];
        text.style.left = `${3 - offset}px`;
        text.style.width = `${Math.max(0, cols.offsets[last + 1] - cols.offsets[first] - 6)}px`;
        const ink = element('span', label, 'cell-overflow-text');
        const origin = horizontal === 'right' ? width - 6 : horizontal === 'center' ? (width - 6) / 2 : 0;
        ink.style.left = `${offset + origin}px`;
        if (horizontal !== 'left') ink.style.transform = `translateX(${horizontal === 'center' ? -50 : -100}%)`;
        text.replaceChildren(ink);
        cell.classList.add('text-overflow');
      }
    }
    cell.append(text);
    if (!isSource) dropOn(cell, 'application/x-sync-range', transfer => {
      if (busy || mode.value !== 'position') return;
      panes.source.range.value = transfer.getData('application/x-sync-range');
      pane.setRange({ row, col });
    });
    body.append(cell);
  }

  for (let col = range.left; col <= range.right; col++) {
    if (!cols.size(col)) continue;
    const label = button(columnName(col), () => {
      if (busy || mode.value !== 'key') return;
      if (isSource) { pickedColumn = col; showPairs(); }
      else pairColumns(pickedColumn, col);
    });
    label.className = 'column-label';
    place(label, cols.offsets[col] + ROW_HEADER, 0, cols.size(col), COLUMN_HEADER);
    label.setAttribute('aria-label', `Cột ${columnName(col)} ${isSource ? 'Excel' : 'Google Sheets'}`);
    if (isSource) {
      label.draggable = true;
      label.addEventListener('dragstart', event => {
        if (busy || mode.value !== 'key') event.preventDefault();
        else event.dataTransfer.setData('application/x-sync-column', String(col));
      });
    } else dropOn(label, 'application/x-sync-column', transfer => {
      if (!busy && mode.value === 'key') pairColumns(Number(transfer.getData('application/x-sync-column')), col);
    });
    columnLabels.append(label);
  }

  for (let row = range.top; row <= range.bottom; row++) {
    if (!rows.size(row)) continue;
    rowLabels.append(place(element('div', String(row + 1), 'row-label'), 0, rows.offsets[row] + COLUMN_HEADER, ROW_HEADER, rows.size(row)));
    for (let col = range.left; col <= range.right; col++) {
      if (!mergeCells.has(row * cols.offsets.length + col)) drawCell(row, col, cellAt(row, col));
    }
  }
  for (const merge of merges) {
    drawCell(merge.top, merge.left, merge.anchor || cellAt(merge.top, merge.left), merge.bottom, merge.right, true);
  }
  layers.columns.replaceChildren(columnLabels);
  layers.rows.replaceChildren(rowLabels);
  layers.cells.replaceChildren(body);
  markSelection(layers.cells, pane.range.value, axes);
}
async function loadExcel(book) {
  if (!book) return;
  panes.source.setBook(book);
  clearPairs();
  await panes.source.draw();
  showStatus('Đã mở Excel. Chọn tab và vùng dữ liệu muốn gửi.');
}

async function redraw() {
  await panes.source.draw();
  await panes.target.draw();
}

function showPreview(model) {
  const root = $('preview-panel');
  const rows = element('table');
  const headings = rows.createTHead().insertRow();
  for (const title of ['Ô đích', 'Giá trị hiện tại', 'Giá trị mới']) headings.append(element('th', title));
  const body = rows.createTBody();
  for (const change of model.changes) {
    const line = body.insertRow();
    const current = change.asNumber && typeof change.beforeValue === 'string' && change.beforeValue !== '' ? `${change.before} (chữ)` : change.before;
    const next = change.asNumber ? `${change.display} (số)` : change.display;
    for (const value of [change.cell, current, next]) line.insertCell().textContent = value ?? '';
  }
  root.replaceChildren(element('strong', `${model.tab} · ${model.count} ô trên ${model.rows} dòng`));
  if (model.added || model.missing || model.skippedEmpty) root.append(element('p', `${model.added} mã mới · ${model.missing} mã bỏ qua · ${model.skippedEmpty} ô nguồn trống bỏ qua`));
  if (model.count) root.append(rows);
  if (model.count > model.changes.length) root.append(element('p', `Hiện ${model.changes.length} thay đổi đầu tiên.`));
  root.hidden = false;
}

panes.source = createPane('source');
panes.target = createPane('target');
window.addEventListener('keydown', event => {
  if (event.key !== 'Escape' || !focusedPane || googleDialog.open || busy) return;
  const control = panes[focusedPane].node.querySelector('.pane-focus');
  setPaneFocus(undefined);
  control.focus({ preventScroll: true });
});
mode.addEventListener('change', modeChanged);
$('add-missing').addEventListener('change', invalidate);
$('profile').addEventListener('change', controls);

$('demo').onclick = () => run(async () => {
  const sample = await invoke('demo');
  panes.source.setBook(sample.source, 'Tồn kho', 'A1:C5');
  panes.target.setBook(sample.target, 0, 'A1:C12');
  panes.target.link.value = '';
  mode.value = 'key';
  $('add-missing').checked = true;
  mappings = [{ from: 0, to: 0, key: true }, { from: 1, to: 1, key: false }];
  savedHeaders = undefined;
  pickedColumn = undefined;
  modeChanged();
  showPairs();
  setProfiles(sample.profiles);
  await redraw();
  showStatus('Đang dùng dữ liệu mẫu. Bấm Xem trước để kiểm tra.');
});

async function updateAccount(info) {
  const before = account.email;
  setAccount(info);
  if (before !== account.email && panes.target.book?.id !== 'demo') {
    panes.target.setBook(undefined);
    clearPairs();
    await panes.target.draw();
  }
  showStatus(account.email ? `Đã đăng nhập ${account.email}. Dán link bảng đích.` : 'Đã đăng xuất Google.');
}

function showGoogleSetup(error = '') {
  $('google-setup').hidden = false;
  $('google-pending').hidden = true;
  $('google-setup-error').textContent = error;
  $('google-setup-error').hidden = !error;
  if (!googleDialog.open) googleDialog.showModal();
  $('google-import').focus();
}

async function logIn() {
  $('google-setup').hidden = true;
  $('google-pending').hidden = false;
  if (!googleDialog.open) googleDialog.showModal();
  $('google-cancel').focus();
  let info;
  try { info = await invoke('login'); }
  finally { googleDialog.close(); }
  if (info.cancelled) { showStatus('Đã hủy đăng nhập Google.'); return; }
  await updateAccount(info);
}

$('login').onclick = () => account.configured ? run(logIn) : showGoogleSetup();
$('configure').onclick = () => showGoogleSetup();
$('logout').onclick = () => run(async () => updateAccount(await invoke('logout')));
$('google-close').onclick = () => googleDialog.close();
$('google-guide').onclick = async () => {
  try { await invoke('googleGuide'); }
  catch (error) { showGoogleSetup(error.message); }
};
$('google-import').onclick = () => run(async () => {
  googleDialog.close();
  let info;
  try { info = await invoke('configure'); }
  catch (error) { showGoogleSetup(error.message); return; }
  if (info.cancelled) { showGoogleSetup(); return; }
  setAccount(info);
  await logIn();
});
$('google-cancel').onclick = () => { invoke('cancelLogin').catch(error => showStatus(error.message, true)); };
googleDialog.addEventListener('cancel', event => {
  if ($('google-pending').hidden) return;
  event.preventDefault();
  $('google-cancel').click();
});

$('preview').onclick = () => run(async () => {
  invalidate();
  preview = await invoke('preview', definition());
  showPreview(preview);
  showStatus(preview.count ? 'Kiểm tra cột giá trị mới, rồi bấm Đồng bộ.' : 'Không có ô nào cần cập nhật.');
});

$('write').onclick = () => run(async () => {
  const pending = preview.id;
  invalidate();
  const result = await invoke('sync', pending);
  showStatus(`Đã ghi ${result.cells} ô trên ${result.rows} dòng${result.demo ? ' trong dữ liệu mẫu' : ''}.${result.logMissing ? ' Chưa lưu được lịch sử trên máy.' : ''}`);
  try { await panes.target.draw(); }
  catch { $('status').textContent += ' Không tải lại được bảng xem. Hãy mở lại tab.'; }
});

$('save').onclick = () => run(async () => {
  const name = $('profile-name').value;
  const records = await invoke('saveProfile', { name, definition: definition() });
  const saved = records.find(record => record.name === name.trim() && record.targetId === panes.target.book.id);
  savedHeaders = saved?.definition.headers;
  setProfiles(records);
  if (saved) $('profile').value = saved.id;
  showStatus(`Đã lưu cách ghép “${name.trim()}”.`);
});

$('load').onclick = () => run(async () => {
  const { source, target, record } = await invoke('loadProfile', $('profile').value);
  const data = record.definition;
  mode.value = data.mode;
  $('add-missing').checked = data.addMissing;
  $('profile-name').value = record.name;
  panes.source.setBook(source, data.sourceTab, data.sourceRange);
  panes.target.setBook(target, data.targetTab, data.targetRange);
  mappings = data.mappings;
  savedHeaders = data.headers;
  pickedColumn = undefined;
  modeChanged();
  showPairs();
  await redraw();
  showStatus(`Đã mở cách ghép “${record.name}”.`);
});

window.addEventListener('dragover', event => event.preventDefault());
window.addEventListener('drop', event => event.preventDefault());
for (const pane of Object.values(panes)) pane.setBook(undefined);
run(async () => { setAccount(await invoke('info')); await redraw(); });
