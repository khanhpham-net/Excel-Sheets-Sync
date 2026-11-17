const { posix } = require('node:path');
const { SaxesParser } = require('saxes');

function scanXml(book, file, open, close) {
  const entry = book.files?.[file];
  if (!entry) return;
  const parser = new SaxesParser();
  parser.on('opentag', tag => open(tag.name.split(':').at(-1), tag.attributes));
  if (close) parser.on('closetag', tag => close(tag.name.split(':').at(-1)));
  parser.write(Buffer.from(entry.content).toString('utf8')).close();
}

function excelLayout(book, title) {
  const sheet = book.Sheets[title];
  const layout = {
    columns: (sheet['!cols'] || []).map(col => col?.hidden ? 0 : col?.wpx ?? (col?.wch == null ? null : col.wch * 7 + 5)),
    rows: (sheet['!rows'] || []).map(row => row?.hidden ? 0 : row?.hpx ?? (row?.hpt == null ? null : row.hpt * 4 / 3)),
    merges: (sheet['!merges'] || []).map(({ s, e }) => ({ top: s.r, left: s.c, bottom: e.r, right: e.c })),
    showGridlines: true, styles: new Map(),
  };
  let id;
  let part;
  scanXml(book, 'xl/workbook.xml', (name, attributes) => {
    if (name === 'sheet' && attributes.name === title) id = attributes['r:id'];
  });
  if (!id) return layout;
  scanXml(book, 'xl/_rels/workbook.xml.rels', (name, attributes) => {
    if (name === 'Relationship' && attributes.Id === id && attributes.Type?.endsWith('/worksheet')) {
      part = attributes.Target.startsWith('/') ? attributes.Target.slice(1) : posix.join('xl', attributes.Target);
    }
  });
  if (!part) return layout;
  scanXml(book, part, (name, attributes) => {
    if (name === 'c' && attributes.r && attributes.s) layout.styles.set(attributes.r, Number(attributes.s));
    if (name === 'sheetView') layout.showGridlines = attributes.showGridLines !== '0' && attributes.showGridLines !== 'false';
    if (name === 'sheetFormatPr' && Number(attributes.defaultRowHeight) > 0) layout.rowSize = Number(attributes.defaultRowHeight) * 4 / 3;
    if (name === 'sheetFormatPr' && Number(attributes.defaultColWidth) > 0) layout.columnSize = Number(attributes.defaultColWidth) * 7 + 5;
  });
  return layout;
}

function excelColor(color, theme) {
  const rgb = color?.rgb || theme?.[color?.theme]?.rgb;
  if (!/^[\dA-F]{6,8}$/i.test(rgb || '')) return undefined;
  const base = rgb.slice(-6);
  if (!color.tint || color.rgb) return `#${base}`;
  const tint = Number(color.tint);
  return `#${base.match(/../g).map(hex => {
    const component = parseInt(hex, 16);
    return Math.round(tint < 0 ? component * (1 + tint) : component + (255 - component) * tint).toString(16).padStart(2, '0');
  }).join('')}`;
}

const sides = ['top', 'right', 'bottom', 'left'];

function borderLine(style, color) {
  if (!style || ['none', 'style_unspecified'].includes(style)) return undefined;
  const line = style === 'double' ? '3px double' : style.includes('dash') ? '1px dashed' : style.includes('dot') ? '1px dotted' : `${style.includes('thick') ? 3 : style.includes('medium') ? 2 : 1}px solid`;
  return `${line} ${color || '#273138'}`;
}

function excelFormats(book) {
  const borders = [];
  let readingBorders = false;
  let border;
  let side;
  const theme = book.Themes?.themeElements?.clrScheme;
  scanXml(book, 'xl/styles.xml', (name, attrs) => {
    if (name === 'borders') readingBorders = true;
    if (!readingBorders) return;
    if (name === 'border') { border = {}; borders.push(border); }
    if (border && sides.includes(name)) { side = name; border[side] = borderLine(attrs.style?.toLowerCase()); }
    if (side && name === 'color' && border[side]) {
      const color = excelColor(attrs, theme);
      if (color) border[side] = border[side].replace(/#[\da-f]+$/i, color);
    }
  }, name => {
    if (name === 'borders') readingBorders = false;
    if (name === side) side = undefined;
    if (name === 'border') border = undefined;
  });

  return (book.Styles?.CellXf || []).map(xf => {
    const font = book.Styles.Fonts?.[xf.fontId] || {};
    const fill = book.Styles.Fills?.[xf.fillId];
    const align = xf.alignment || {};
    return {
      background: fill?.patternType === 'solid' ? excelColor(fill.fgColor, theme) : undefined,
      color: excelColor(font.color, theme), font: font.name, size: font.sz,
      bold: Boolean(font.bold), italic: Boolean(font.italic), underline: Boolean(font.underline), strike: Boolean(font.strike),
      horizontal: align.horizontal === 'centerContinuous' ? 'center' : align.horizontal,
      vertical: align.vertical, wrap: align.wrapText === '1' || align.wrapText === true,
      borders: borders[xf.borderId],
    };
  });
}

function googleColor(style, backup, theme) {
  const value = style?.rgbColor || theme[style?.themeColor] || backup;
  if (!value) return undefined;
  const rgb = ['red', 'green', 'blue'].map(channel => Math.round(Math.max(0, Math.min(1, value[channel] || 0)) * 255));
  return `rgb(${rgb.join(',')})`;
}

function googleFormat(format, theme) {
  if (!format) return undefined;
  const text = format.textFormat || {};
  return {
    background: googleColor(format.backgroundColorStyle, format.backgroundColor, theme),
    color: googleColor(text.foregroundColorStyle, text.foregroundColor, theme),
    font: text.fontFamily, size: text.fontSize,
    bold: text.bold, italic: text.italic, underline: text.underline, strike: text.strikethrough,
    horizontal: format.horizontalAlignment?.toLowerCase(), vertical: format.verticalAlignment?.toLowerCase(),
    wrap: format.wrapStrategy === 'WRAP',
    overflow: !['CLIP', 'WRAP', 'LEGACY_WRAP'].includes(format.wrapStrategy),
    borders: format.borders && Object.fromEntries(sides.map(side => [side, borderLine(format.borders[side]?.style?.toLowerCase(), googleColor(format.borders[side]?.colorStyle, format.borders[side]?.color, theme))])),
  };
}

module.exports = { excelLayout, excelFormats, googleFormat };
