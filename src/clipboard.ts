import { mergeParagraphStyle, nodeId, normalize, paragraph, plainText } from './model';
import type { Block, DocModel, Inline, ListStyle, Paragraph, ParagraphStyle, Table, TableCell, TextFormat, TableBorders, CellMargins } from './model';
import { safeHref } from './ooxml';

export type PasteMode = 'source' | 'merge' | 'text';
type Declarations = Map<string, string>;
const IGNORED = new Set(['SCRIPT', 'STYLE', 'META', 'LINK', 'IFRAME', 'OBJECT', 'EMBED', 'HEAD', 'TITLE', 'NOSCRIPT', 'SVG']);
const BLOCK = /^(P|H[1-6]|DIV|BLOCKQUOTE|PRE|ADDRESS|SECTION|ARTICLE|FIGCAPTION|FIGURE)$/;
const MAX_HTML = 32 * 1024 * 1024;

function declarations(css: string): Declarations {
  return new Map(css.split(';').flatMap(pair => {
    const colon = pair.indexOf(':');
    return colon < 0 ? [] : [[pair.slice(0, colon).trim().toLowerCase(), pair.slice(colon + 1).trim().replace(/\s*!important$/, '')]];
  }));
}

function cssLength(value?: string) {
  if (!value) return;
  const length = /^(-?\d*\.?\d+)(px|pt|in|cm|mm)?$/i.exec(value.trim());
  if (!length) return;
  const number = Number(length[1]);
  const factors: Record<string, number> = { px: 1, pt: 4 / 3, in: 96, cm: 96 / 2.54, mm: 96 / 25.4 };
  return Math.max(-2000, Math.min(4000, number * (factors[length[2]?.toLowerCase()] ?? 1)));
}

function color(document: Document, value?: string) {
  if (!value || /^(auto|inherit|transparent|currentcolor)$/i.test(value)) return;
  const probe = document.createElement('span');
  probe.style.color = value;
  const rgb = /^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*1(?:\.0+)?)?\)$/.exec(probe.style.color);
  if (rgb) return '#' + rgb.slice(1).map(channel => Number(channel).toString(16).padStart(2, '0')).join('');
  if (!probe.style.color) return;
  const canvas = document.createElement('canvas').getContext('2d');
  if (!canvas) return;
  canvas.fillStyle = '#010203';
  canvas.fillStyle = probe.style.color;
  if (/^#[0-9a-f]{6}$/i.test(canvas.fillStyle)) return canvas.fillStyle;
}

function formatOf(node: Element, css: Declarations, mode: PasteMode): TextFormat {
  const values: TextFormat = {};
  const weight = css.get('font-weight');
  const italic = css.get('font-style');
  const decoration = css.get('text-decoration-line') ?? css.get('text-decoration');
  const vertical = css.get('vertical-align');
  if (weight) values.bold = weight === 'bold' || Number(weight) >= 600;
  else if (/^(B|STRONG)$/.test(node.tagName)) values.bold = true;
  if (italic) values.italic = italic === 'italic' || italic === 'oblique';
  else if (/^(I|EM)$/.test(node.tagName)) values.italic = true;
  if (decoration) { values.underline = /underline/.test(decoration); values.strike = /line-through/.test(decoration); }
  else if (node.tagName === 'U') values.underline = true;
  else if (/^(S|STRIKE|DEL)$/.test(node.tagName)) values.strike = true;
  if (vertical === 'super' || node.tagName === 'SUP') values.vertical = 'superscript';
  if (vertical === 'sub' || node.tagName === 'SUB') values.vertical = 'subscript';
  if (mode === 'merge') return values;
  const font = (css.get('font-family') ?? node.getAttribute('face'))?.split(',')[0].replace(/['"]/g, '').trim();
  const size = cssLength(css.get('font-size'));
  const foreground = color(node.ownerDocument, css.get('color') ?? node.getAttribute('color') ?? undefined);
  const background = color(node.ownerDocument, css.get('background-color') ?? css.get('background') ?? (node.tagName === 'MARK' ? 'yellow' : undefined));
  if (font && font.length <= 80) values.font = font;
  if (size && size >= 4 / 3 && size <= 533) values.size = Math.round(size * 1.5) / 2;
  if (foreground) values.color = foreground;
  if (background) values.highlight = background;
  return values;
}

function paragraphProps(css: Declarations): ParagraphStyle {
  const style: ParagraphStyle = {};
  const align = css.get('text-align');
  if (align === 'left' || align === 'center' || align === 'right' || align === 'justify') style.alignment = align;
  for (const [property, key] of [['margin-top', 'before'], ['margin-bottom', 'after'], ['margin-left', 'indent'],
    ['margin-right', 'rightIndent'], ['text-indent', 'firstLine']] as const) {
    const value = cssLength(css.get(property));
    if (value !== undefined) style[key] = value;
  }
  const line = css.get('line-height');
  const lineHeight = line?.endsWith('%') ? Number.parseFloat(line) / 100 : Number(line);
  if (lineHeight >= .8 && lineHeight <= 6) style.lineHeight = lineHeight;
  else {
    const value = cssLength(line);
    if (value && value > 0) style.lineSpacing = { rule: css.get('mso-line-height-rule') === 'at-least' ? 'atLeast' : 'exact', value };
  }
  if (css.get('direction') === 'rtl' || css.get('direction') === 'ltr') style.direction = css.get('direction') as 'rtl' | 'ltr';
  if (css.get('page-break-after') === 'avoid' || css.get('break-after') === 'avoid') style.keepNext = true;
  if (css.get('page-break-inside') === 'avoid' || css.get('break-inside') === 'avoid') style.keepLines = true;
  return style;
}

function boxFormatting(node: Element, css: Declarations) {
  const probe = node.ownerDocument.createElement('span');
  for (const [name, value] of css) probe.style.setProperty(name, value);
  const borders: TableBorders = {}, margins: CellMargins = {};
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    const borderStyle = probe.style.getPropertyValue(`border-${side}-style`);
    const size = cssLength(probe.style.getPropertyValue(`border-${side}-width`));
    if (borderStyle) borders[side] = { style: borderStyle === 'none' || borderStyle === 'hidden' ? 'none' : borderStyle === 'double' ? 'double'
      : borderStyle === 'dotted' ? 'dotted' : borderStyle === 'dashed' ? 'dashed' : 'solid', size: (size ?? 1) * .75,
    color: color(node.ownerDocument, probe.style.getPropertyValue(`border-${side}-color`)) ?? '#000000' };
    const padding = cssLength(probe.style.getPropertyValue(`padding-${side}`));
    if (padding !== undefined) margins[side] = Math.max(0, padding);
  }
  return { borders: Object.keys(borders).length ? borders : undefined, margins: Object.keys(margins).length ? margins : undefined };
}

function inlineImage(node: Element): Inline | undefined {
  const source = node.getAttribute('src') ?? '';
  const match = /^data:(image\/(?:png|jpeg));base64,([\w+/=\s]+)$/i.exec(source);
  if (!match || match[2].length > 16 * 1024 * 1024) return;
  try {
    const bytes = Uint8Array.from(atob(match[2]), char => char.charCodeAt(0));
    const mime = match[1].toLowerCase();
    const isPng = mime === 'image/png' && [137, 80, 78, 71, 13, 10, 26, 10].every((value, index) => bytes[index] === value);
    const isJpeg = mime === 'image/jpeg' && bytes[0] === 255 && bytes[1] === 216 && bytes[2] === 255;
    if (!isPng && !isJpeg) return;
    const dimensions = declarations(node.getAttribute('style') ?? '');
    const view = new DataView(bytes.buffer);
    const naturalWidth = isPng && bytes.length >= 24 ? view.getUint32(16) : 0;
    const naturalHeight = isPng && bytes.length >= 24 ? view.getUint32(20) : 0;
    const width = cssLength(node.getAttribute('width') ?? dimensions.get('width')) || naturalWidth || 160;
    const height = cssLength(node.getAttribute('height') ?? dimensions.get('height')) || naturalHeight || 100;
    const scale = Math.min(1, 2000 / width, 2000 / height);
    return { kind: 'image', id: nodeId('image'), blob: new Blob([bytes], { type: mime }),
      width: Math.max(1, width * scale), height: Math.max(1, height * scale), alt: node.getAttribute('alt') ?? 'Pasted picture' };
  } catch {
    return;
  }
}

export function readClipboard(html: string, model: DocModel, mode: PasteMode, warnings: string[] = []): Block[] {
  if (!html || mode === 'text') return [];
  const limitNotice = 'The pasted content exceeds rich-text limits. Rich formatting was skipped to avoid truncating it.';
  if (html.length > MAX_HTML) { warnings.push(limitNotice); return []; }
  const template = document.createElement('template');
  template.innerHTML = html;
  const cssRules: { tag?: string; name?: string; specificity: number; rule: Declarations }[] = [];
  const listRules = new Map<string, Declarations>();
  for (const sheet of template.content.querySelectorAll('style')) {
    for (const match of (sheet.textContent ?? '').matchAll(/([^{}]+)\{([^{}]*)\}/g)) {
      const rule = declarations(match[2]);
      const selector = match[1].replace(/<!--[\s\S]*?-->|\/\*[\s\S]*?\*\//g, '').replace(/^<!--/, '').trim();
      const wordList = /@list\s+(l\d+):level(\d+)$/i.exec(selector);
      if (wordList) listRules.set(wordList[1] + ':' + wordList[2], rule);
      for (const part of selector.split(',')) {
        const simple = /^([a-z][a-z0-9]*)?(?:\.([\w-]+))?$/i.exec(part.trim());
        if (simple && (simple[1] || simple[2])) cssRules.push({ tag: simple[1]?.toLowerCase(), name: simple[2],
          specificity: (simple[1] ? 1 : 0) + (simple[2] ? 10 : 0), rule });
        if (cssRules.length > 2000) { warnings.push(limitNotice); return []; }
      }
    }
  }
  const fragment = /<!--\s*StartFragment\s*-->([\s\S]*?)<!--\s*EndFragment\s*-->/i.exec(html)?.[1];
  if (fragment) template.innerHTML = fragment;
  const styleCache = new Map<string, Declarations>();
  const styleOf = (node: Element) => {
    const tag = node.tagName.toLowerCase();
    const key = tag + ':' + node.className;
    let rules = styleCache.get(key);
    if (!rules) {
      rules = new Map(cssRules.filter(rule => (!rule.tag || rule.tag === tag) && (!rule.name || node.classList.contains(rule.name)))
        .sort((a, b) => a.specificity - b.specificity).flatMap(entry => [...entry.rule]));
      styleCache.set(key, rules);
    }
    return new Map([...rules, ...declarations(node.getAttribute('style') ?? '')]);
  };
  const headingId = (node: Element) => {
    const level = /^(?:H|MsoHeading)([1-9])$/i.exec(node.tagName)?.[1] ??
      Array.from(node.classList).map(name => /^MsoHeading([1-9])$/i.exec(name)?.[1]).find(Boolean);
    return level ? model.styles.find(style => style.heading === Number(level))?.id ?? 'Heading' + level : model.defaultStyleId;
  };
  const wordLists = new Map<string, string>();
  function wordList(css: Declarations, label: string): ListStyle | undefined {
    const entry = /(?:^|\s)(l\d+)\s+level(\d+)\s+(lfo\d+)/i.exec(css.get('mso-list') ?? '');
    if (!entry) return;
    const rule = listRules.get(entry[1] + ':' + entry[2]);
    const kind = rule?.get('mso-level-number-format') === 'bullet' || (!rule && label.trim() !== '' && !/^\s*\d/.test(label))
      ? 'bullet' : 'number';
    const key = entry[3] + ':' + kind;
    if (!wordLists.has(key)) wordLists.set(key, nodeId('list'));
    const value = Number.parseInt(rule?.get('mso-level-start-at') ?? label, 10);
    const formats: Record<string, ListStyle['numberFormat']> = { 'alpha-lower': 'lowerLetter', 'alpha-upper': 'upperLetter',
      'roman-lower': 'lowerRoman', 'roman-upper': 'upperRoman', decimal: 'decimal' };
    return { id: wordLists.get(key)!, kind, level: Math.max(0, Math.min(8, Number(entry[2]) - 1)),
      numberFormat: formats[rule?.get('mso-level-number-format') ?? 'decimal'],
      start: Number.isFinite(value) ? value : 1 };
  }
  let visited = 0;
  const limit = new Error('Clipboard complexity limit');
  function visit() { if (++visited > 50_000) throw limit; }
  function inline(nodes: Node[], format: TextFormat, link?: string, preserve = false, depth = 0): Inline[] {
    if (depth > 48) throw limit;
    const output: Inline[] = [];
    for (const node of nodes) {
      visit();
      if (node.nodeType === Node.TEXT_NODE) {
        const text = preserve ? node.textContent ?? '' : (node.textContent ?? '').replace(/[ \r\n\t\f]+/g, ' ');
        if (text) output.push({ kind: 'text', text, format, link: link ? { href: link } : undefined });
        continue;
      }
      if (!(node instanceof Element) || IGNORED.has(node.tagName)) continue;
      const css = styleOf(node);
      if (/^ignore$/i.test(css.get('mso-list') ?? '') || css.get('display') === 'none' || node.hasAttribute('hidden')) continue;
      if (node.tagName === 'IMG') {
        const image = inlineImage(node);
        if (image) output.push(image);
        else {
          warnings.push('Some pictures could not be pasted. Open the DOCX to preserve embedded pictures.');
          if (node.getAttribute('alt')) output.push({ kind: 'text', text: node.getAttribute('alt')!, format });
        }
        continue;
      }
      if (node.tagName === 'BR') {
        output.push({ kind: 'text', text: '\n', format });
        continue;
      }
      const tabs = Number(css.get('mso-tab-count'));
      if (tabs > 0 && tabs <= 16) {
        output.push({ kind: 'text', text: '\t'.repeat(tabs), format });
        continue;
      }
      output.push(...inline(Array.from(node.childNodes), { ...format, ...formatOf(node, css, mode) },
        node.tagName === 'A' ? safeHref(node.getAttribute('href') ?? '') : link,
        preserve || /^pre/.test(css.get('white-space') ?? ''), depth + 1));
    }
    return output;
  }
  function makeParagraph(node: Element, nodes = Array.from(node.childNodes), list?: ListStyle,
    inherited: TextFormat = {}, inheritedStyle: ParagraphStyle = {}): Paragraph {
    visit();
    const css = styleOf(node);
    const runs = normalize(inline(nodes, { ...inherited, ...formatOf(node, css, mode) }, undefined, node.tagName === 'PRE'));
    const visible = runs.some(run => run.kind !== 'text' || run.text.trim());
    const item = paragraph(headingId(node), visible ? runs : []);
    if (mode === 'source') item.style = { ...mergeParagraphStyle(inheritedStyle, paragraphProps(css)),
      ...(/^(rtl|ltr)$/.test(node.getAttribute('dir') ?? '') ? { direction: node.getAttribute('dir') as 'rtl' | 'ltr' } : {}) };
    const marker = Array.from(node.querySelectorAll('span')).find(span => /^ignore$/i.test(styleOf(span).get('mso-list') ?? ''));
    item.list = list ?? wordList(css, marker?.textContent ?? '');
    return item;
  }
  function walk(container: Element | DocumentFragment, depth = 0, listDepth = 0,
    inherited: TextFormat = {}, inheritedStyle: ParagraphStyle = {}): Block[] {
    if (depth > 28) throw limit;
    const blocks: Block[] = [];
    let buffer: Node[] = [];
    function flush() {
      if (buffer.length) {
        const runs = normalize(inline(buffer, inherited));
        if (runs.some(run => run.kind !== 'text' || run.text.trim())) {
          const item = paragraph(model.defaultStyleId, runs);
          if (mode === 'source') item.style = { ...inheritedStyle };
          blocks.push(item);
        }
      }
      buffer = [];
    }
    for (const node of Array.from(container.childNodes)) {
      visit();
      if (!(node instanceof Element)) { buffer.push(node); continue; }
      if (IGNORED.has(node.tagName)) continue;
      const css = styleOf(node);
      if (css.get('display') === 'none' || node.hasAttribute('hidden')) continue;
      const format = { ...inherited, ...formatOf(node, css, mode) };
      const properties = paragraphProps(css);
      const style = mergeParagraphStyle(inheritedStyle, {
        ...(properties.alignment === undefined ? {} : { alignment: properties.alignment }),
        ...(properties.lineHeight === undefined ? {} : { lineHeight: properties.lineHeight }),
        ...(properties.lineSpacing === undefined ? {} : { lineSpacing: properties.lineSpacing }),
        ...(properties.direction === undefined ? {} : { direction: properties.direction }) });
      if (node.tagName === 'TABLE') {
        flush();
        const rows: Table['rows'] = [];
        const tableBox = mode === 'source' ? boxFormatting(node, css) : {};
        const rawWidth = css.get('width') ?? node.getAttribute('width') ?? '';
        const width = mode === 'source' ? rawWidth.endsWith('%') ? { unit: 'percent' as const, value: Math.max(1, Math.min(100, Number.parseFloat(rawWidth) || 100)) }
          : cssLength(rawWidth) ? { unit: 'px' as const, value: Math.max(1, cssLength(rawWidth)!) } : undefined : undefined;
        const tablePixels = width?.unit === 'px' ? width.value : (model.page.width - model.page.left - model.page.right) * (width?.value ?? 100) / 100;
        const covered: number[] = [];
        const tableRows = Array.from((node as HTMLTableElement).rows);
        if (tableRows.length > 200) throw limit;
        const preserveSpans = tableRows.every(row => Array.from(row.cells).every(cell => cell.rowSpan === 1));
        const merged = preserveSpans && tableRows.some(row => Array.from(row.cells).some(cell => cell.colSpan > 1));
        for (const [rowIndex, row] of tableRows.entries()) {
          const cells: TableCell[] = [];
          const empty = (): TableCell => ({ id: nodeId('cell'), paragraphs: [paragraph(model.defaultStyleId)] });
          const skip = () => {
            while ((covered[cells.length] ?? 0) > 0) {
              if (cells.length >= 20) throw limit;
              covered[cells.length]--;
              cells.push(empty());
            }
          };
          for (const cell of Array.from(row.cells)) {
            skip();
            const occupied = cells.reduce((sum, cell) => sum + (cell.colSpan ?? 1), 0);
            if (occupied >= 20 || occupied + cell.colSpan > 20) throw limit;
            if (!preserveSpans && (cell.colSpan > 1 || cell.rowSpan !== 1)) warnings.push('Vertically merged table cells were converted to regular cells.');
            const start = cells.length;
            const cellCss = styleOf(cell);
            const cellFormat = { ...format, ...formatOf(row, styleOf(row), mode), ...formatOf(cell, cellCss, mode) };
            const cellStyle = paragraphProps(cellCss);
            const cellBlocks = walk(cell, depth + 1, 0, cellFormat, { ...style,
              ...(cellStyle.alignment ? { alignment: cellStyle.alignment } : {}) });
            if (cellBlocks.some(item => item.kind === 'table')) warnings.push('Nested tables were flattened into paragraphs.');
            const paragraphs = cellBlocks.flatMap(item => item.kind === 'paragraph' ? [item] :
              item.kind === 'table' ? item.rows.flatMap(row => row.cells.flatMap(value => value.paragraphs)) : []);
            const css = styleOf(cell);
            const cellWidth = css.get('width') ?? cell.getAttribute('width') ?? '';
            const vertical = css.get('vertical-align') ?? cell.getAttribute('valign');
            cells.push({ id: nodeId('cell'), paragraphs: paragraphs.length ? paragraphs : [paragraph(model.defaultStyleId)],
              colSpan: preserveSpans && cell.colSpan > 1 ? cell.colSpan : undefined,
              ...(mode === 'source' ? boxFormatting(cell, css) : {}),
              width: mode === 'source' ? cellWidth.endsWith('%') ? Math.max(1, Number.parseFloat(cellWidth) * tablePixels / 100) : cssLength(cellWidth) : undefined,
              vertical: mode === 'source' ? vertical === 'middle' ? 'center' : vertical === 'top' || vertical === 'bottom' ? vertical : undefined : undefined,
              fill: mode === 'source' ? color(cell.ownerDocument, css.get('background-color') ?? css.get('background')) : undefined });
            if (!preserveSpans) for (let i = 1; i < Math.min(cell.colSpan, 20 - start); i++) cells.push(empty());
            for (let i = start; i < cells.length; i++) covered[i] = cell.rowSpan === 0 ? tableRows.length - rowIndex - 1 : Math.max(0, cell.rowSpan - 1);
          }
          skip();
          const rowCss = styleOf(row);
          if (cells.length) rows.push({ id: nodeId('row'), cells, height: mode === 'source' ? cssLength(rowCss.get('height') ?? row.getAttribute('height') ?? undefined) : undefined });
        }
        if (rows.length) {
          const count = (row: Table['rows'][number]) => row.cells.reduce((sum, cell) => sum + (cell.colSpan ?? 1), 0);
          const columns = Math.max(...rows.map(count));
          for (const row of rows) while (count(row) < columns) row.cells.push({ id: nodeId('cell'), paragraphs: [paragraph(model.defaultStyleId)] });
          const alignment = node.getAttribute('align') ?? css.get('text-align');
          blocks.push({ kind: 'table', id: nodeId('table'), rows, width, ...tableBox,
            structuralLock: merged || undefined,
            columnWidths: merged ? rows[0].cells.flatMap(cell => Array.from({ length: cell.colSpan ?? 1 }, () => (cell.width ?? cell.colSpan ?? 1) / (cell.colSpan ?? 1))) : undefined,
            align: mode === 'source' && (alignment === 'center' || alignment === 'right' || alignment === 'left') ? alignment : undefined });
        }
      } else if (node.tagName === 'OL' || node.tagName === 'UL') {
        flush();
        const list: ListStyle = { id: nodeId('list'), level: Math.min(listDepth, 8), kind: node.tagName === 'OL' ? 'number' : 'bullet',
          numberFormat: ({ a: 'lowerLetter', A: 'upperLetter', i: 'lowerRoman', I: 'upperRoman' } as const)[node.getAttribute('type') as 'a' | 'A' | 'i' | 'I'] ?? 'decimal',
          start: Math.max(1, Number.parseInt(node.getAttribute('start') ?? '1', 10) || 1) };
        for (const item of Array.from(node.children).filter(child => child.tagName === 'LI')) {
          const nested = Array.from(item.children).filter(child => /^(OL|UL)$/.test(child.tagName));
          blocks.push(makeParagraph(item, Array.from(item.childNodes).filter(child => !nested.includes(child as Element)), list, format, style));
          nested.forEach(child => {
            const wrapper = document.createDocumentFragment();
            wrapper.append(child.cloneNode(true));
            blocks.push(...walk(wrapper, depth + 1, listDepth + 1, format, style));
          });
        }
      } else if (BLOCK.test(node.tagName)) {
        flush();
        if (Array.from(node.children).some(child => BLOCK.test(child.tagName) || /^(TABLE|OL|UL)$/.test(child.tagName))) {
          blocks.push(...walk(node, depth + 1, listDepth, format, style));
        } else blocks.push(makeParagraph(node, undefined, undefined, inherited, inheritedStyle));
      } else buffer.push(node);
    }
    flush();
    return blocks;
  }
  try {
    return walk(template.content);
  } catch (cause) {
    if (cause !== limit) throw cause;
    warnings.push(limitNotice);
    return [];
  }
}

export function clipboardText(blocks: Block[]) {
  return blocks.map(block => block.kind === 'paragraph' ? plainText(block.inlines) : block.kind === 'table'
    ? block.rows.map(row => row.cells.map(cell => cell.paragraphs.map(item => plainText(item.inlines)).join(' ') + '\t'.repeat((cell.colSpan ?? 1) - 1)).join('\t')).join('\n') : '').join('\n');
}
