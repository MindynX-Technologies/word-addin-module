import {
  DEFAULT_CELL_MARGINS, DEFAULT_FORMAT, DEFAULT_TABLE_BORDER, lengthOf, orderedSelection,
  mergeParagraphStyle, paragraphs, plainText, sliceInlines, tableColumnWeights, wholeTableSelection,
} from './model';
import type { Block, DocModel, ImageRun, Inline, ModelSelection, Paragraph, Table, TableBorder, TableBorderSide, TextFormat } from './model';
import { safeHref } from './ooxml';
import { clipboardText } from './clipboard';
import { DocxEditorError } from './types';

const imageSources = new WeakMap<Blob, string>();
const pendingImages = new WeakMap<Blob, Promise<string>>();
function prepareImage(blob: Blob) {
  let pending = pendingImages.get(blob);
  if (!pending) {
    pending = new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => { const source = String(reader.result); imageSources.set(blob, source); resolve(source); };
      reader.onerror = () => { pendingImages.delete(blob); reject(reader.error); };
      reader.readAsDataURL(blob);
    });
    pendingImages.set(blob, pending);
  }
  return pending;
}

export function prepareClipboardImages(model: DocModel) {
  for (const item of paragraphs(model)) for (const run of item.inlines) {
    if (run.kind === 'image') void prepareImage(run.blob).catch(() => {});
  }
}

function selectedImages(model: DocModel, selection: ModelSelection) {
  const range = orderedSelection(model, selection);
  const inlines = range?.selected.flatMap(item => sliceInlines(item.inlines,
    item.id === range.start.id ? range.start.offset : 0,
    item.id === range.end.id ? range.end.offset : lengthOf(item.inlines))) ?? [];
  return { images: inlines.filter((run): run is ImageRun => run.kind === 'image'),
    pictureOnly: inlines.length === 1 && inlines[0].kind === 'image' };
}

export function writeImageClipboard(model: DocModel, selection: ModelSelection) {
  const { images, pictureOnly } = selectedImages(model, selection);
  if (!navigator.clipboard?.write || typeof ClipboardItem === 'undefined') {
    return Promise.reject(new DocxEditorError('CLIPBOARD_FAILED', 'Your browser cannot write pictures to the clipboard. Use a supported browser over HTTPS, or download the document.'));
  }
  const payload = Promise.all(images.map(image => prepareImage(image.blob))).then(() => {
    const content = new Map<string, string>();
    writeClipboard({ setData: (type, value) => { content.set(type, value); } }, model, selection);
    return content;
  });
  const data: Record<string, Promise<Blob>> = {
    'text/html': payload.then(value => new Blob([value.get('text/html') ?? ''], { type: 'text/html' })),
    'text/plain': payload.then(value => new Blob([value.get('text/plain') ?? ''], { type: 'text/plain' })),
  };
  if (pictureOnly) data['image/png'] = images[0].blob.type === 'image/png' ? Promise.resolve(images[0].blob) : pngImage(images[0].blob);
  // Invoke write before yielding, preserving Safari's transient user activation.
  return navigator.clipboard.write([new ClipboardItem(data)]);
}

async function pngImage(blob: Blob) {
  const bitmap = await createImageBitmap(blob);
  try {
    const canvas = document.createElement('canvas');
    canvas.width = bitmap.width; canvas.height = bitmap.height;
    const context = canvas.getContext('2d');
    if (!context) throw new Error('Picture encoding is unavailable.');
    context.drawImage(bitmap, 0, 0);
    return await new Promise<Blob>((resolve, reject) => canvas.toBlob(value => value ? resolve(value) : reject(new Error('Picture encoding failed.')), 'image/png'));
  } finally { bitmap.close(); }
}

export function selectionHasImages(model: DocModel, selection: ModelSelection) {
  const range = orderedSelection(model, selection);
  return !!range?.selected.some(item => sliceInlines(item.inlines,
    item.id === range.start.id ? range.start.offset : 0,
    item.id === range.end.id ? range.end.offset : lengthOf(item.inlines)).some(run => run.kind === 'image'));
}

function applyTextStyle(element: HTMLElement, format: TextFormat) {
  Object.assign(element.style, {
    fontFamily: format.font ? `"${format.font.replace(/["\\\r\n]/g, '')}"` : '',
    fontSize: format.size ? `${format.size}pt` : '',
    fontWeight: format.bold ? 'bold' : 'normal',
    fontStyle: format.italic ? 'italic' : 'normal',
    textDecoration: [format.underline && 'underline', format.strike && 'line-through'].filter(Boolean).join(' '),
    color: format.color ?? '',
    backgroundColor: format.highlight ?? '',
    verticalAlign: format.vertical === 'superscript' ? 'super' : format.vertical === 'subscript' ? 'sub' : '',
  });
}

function appendInline(parent: HTMLElement, inline: Inline, base: TextFormat) {
  if (inline.kind === 'image') {
    const source = imageSources.get(inline.blob);
    if (!source) throw new DocxEditorError('CLIPBOARD_FAILED', 'The picture is still preparing. Try copying again.');
    const image = document.createElement('img');
    image.src = source; image.width = Math.round(inline.width); image.height = Math.round(inline.height); image.alt = inline.alt;
    parent.append(image);
    return;
  }
  if (inline.kind === 'pageBreak') {
    parent.append(document.createElement('br'));
    return;
  }
  const span = document.createElement('span');
  applyTextStyle(span, { ...base, ...inline.format });
  span.style.whiteSpace = 'pre-wrap';
  plainText([inline]).split('\n').forEach((line, index) => {
    if (index) span.append(document.createElement('br'));
    span.append(line);
  });
  const href = inline.kind === 'text' && inline.link && safeHref(inline.link.href);
  if (href) {
    const link = document.createElement('a');
    link.href = href;
    if (!inline.format.color) span.style.color = '';
    if (inline.format.underline === undefined && !base.underline) {
      span.style.textDecoration = (inline.format.strike ?? base.strike) ? 'underline line-through' : 'underline';
    }
    link.append(span);
    parent.append(link);
  } else parent.append(span);
}

function borderCss(border?: TableBorder) {
  return border && border.style !== 'none' ? `${border.size}pt ${border.style} ${border.color}` : 'none';
}

export function writeClipboard(data: Pick<DataTransfer, 'setData'>, model: DocModel, selection: ModelSelection) {
  const range = orderedSelection(model, selection);
  if (!range) return;

  const fragment = document.createElement('div');
  const selected = new Map(range.selected.map(item => [item.id, sliceInlines(item.inlines,
    item.id === range.start.id ? range.start.offset : 0,
    item.id === range.end.id ? range.end.offset : lengthOf(item.inlines))]));
  const wholeTable = wholeTableSelection(model, selection);
  const styles = new Map(model.styles.map(style => [style.id, style]));
  const listCounts = new Map<string, number>();

  function appendParagraph(parent: HTMLElement, item: Paragraph, listItem = false) {
    const named = styles.get(item.styleId) ?? styles.get(model.defaultStyleId);
    const format = { ...DEFAULT_FORMAT, ...named?.format };
    const style = mergeParagraphStyle(named?.paragraph, item.style);
    const node = document.createElement(listItem ? 'li' : named?.heading ? `h${Math.min(6, named.heading)}` : 'p');
    applyTextStyle(node, format);
    Object.assign(node.style, {
      textAlign: style.alignment ?? '',
      marginTop: `${style.before ?? 0}px`,
      marginBottom: `${style.after ?? 4}px`,
      marginLeft: style.indent === undefined ? '' : `${style.indent}px`,
      marginRight: style.rightIndent === undefined ? '' : `${style.rightIndent}px`,
      textIndent: style.firstLine === undefined ? '' : `${style.firstLine}px`,
      lineHeight: style.lineHeight === undefined && style.lineSpacing ? `${style.lineSpacing.value}px` : String(style.lineHeight ?? 1.32),
      direction: style.direction ?? '',
    });
    if (style.lineSpacing?.rule === 'atLeast') node.setAttribute('style', node.getAttribute('style') + ';mso-line-height-rule:at-least');
    const inlines = selected.get(item.id);
    if (inlines?.length) for (const inline of inlines) appendInline(node, inline, format);
    else node.append(document.createElement('br'));
    parent.append(node);
    return node;
  }

  function appendBlocks(parent: HTMLElement, blocks: readonly Block[]) {
    const lists: { level: number; id: string; node: HTMLOListElement | HTMLUListElement; last?: HTMLElement }[] = [];
    for (const block of blocks) {
      if (block.kind !== 'paragraph') {
        lists.length = 0;
        if (block.kind === 'table') appendTable(parent, block);
        continue;
      }
      const list = block.list;
      let number = 1;
      if (list?.kind === 'number') {
        const key = `${list.id}:${list.level}`;
        number = (listCounts.get(key) ?? ((list.start ?? 1) - 1)) + 1;
        listCounts.set(key, number);
        for (let level = list.level + 1; level < 9; level++) listCounts.delete(`${list.id}:${level}`);
      }
      if (!selected.has(block.id)) { lists.length = 0; continue; }
      if (!list) {
        lists.length = 0;
        appendParagraph(parent, block);
        continue;
      }
      const tag = list.kind === 'number' ? 'ol' : 'ul';
      while (lists.length && lists.at(-1)!.level > list.level) lists.pop();
      let current = lists.at(-1);
      if (current?.level === list.level && (current.id !== list.id || current.node.localName !== tag)) {
        lists.pop();
        current = lists.at(-1);
      }
      if (current?.level !== list.level) {
        const node = document.createElement(tag);
        if (list.kind === 'number' && number !== 1) node.setAttribute('start', String(number));
        const type = list.numberFormat && ({ lowerLetter: 'a', upperLetter: 'A', lowerRoman: 'i', upperRoman: 'I', decimal: '1' } as const)[list.numberFormat];
        if (type) node.setAttribute('type', type);
        (current?.last ?? parent).append(node);
        current = { level: list.level, id: list.id, node };
        lists.push(current);
      }
      current.last = appendParagraph(current.node, block, true);
    }
  }

  function appendTable(parent: HTMLElement, table: Table) {
    const rows = table.rows.map(row => row.cells.map(cell => cell.paragraphs.some(item => selected.has(item.id))));
    const first = rows.findIndex(row => row.some(Boolean));
    if (first < 0) return;
    let last = rows.length - 1;
    while (!rows[last].some(Boolean)) last--;
    if (wholeTable !== table && first === last && rows[first].filter(Boolean).length === 1) {
      appendBlocks(parent, table.rows[first].cells[rows[first].indexOf(true)].paragraphs);
      return;
    }
    const node = document.createElement('table');
    node.style.borderCollapse = 'collapse';
    node.style.width = table.width?.unit === 'px' ? `${table.width.value}px` : `${table.width?.value ?? 100}%`;
    if (table.align) node.setAttribute('align', table.align);
    const body = node.createTBody();
    const weights = tableColumnWeights(table);
    const total = weights.reduce((sum, weight) => sum + weight, 0);
    const border = (side: TableBorderSide) => table.borders ? table.borders[side] : DEFAULT_TABLE_BORDER;
    for (let index = first; index <= last; index++) {
      const row = body.insertRow();
      let gridColumn = 0;
      for (const cell of table.rows[index].cells) {
        const column = gridColumn;
        gridColumn += cell.colSpan ?? 1;
        const td = row.insertCell();
        td.colSpan = cell.colSpan ?? 1;
        const margins = { ...DEFAULT_CELL_MARGINS, ...table.margins, ...cell.margins };
        Object.assign(td.style, {
          width: `${weights.slice(column, gridColumn).reduce((sum, width) => sum + width, 0) / total * 100}%`,
          padding: `${margins.top}px ${margins.right}px ${margins.bottom}px ${margins.left}px`,
          backgroundColor: cell.fill ?? '',
          verticalAlign: cell.vertical === 'center' ? 'middle' : cell.vertical ?? 'top',
          borderTop: borderCss(cell.borders?.top ?? border(index === 0 ? 'top' : 'insideH')),
          borderBottom: borderCss(cell.borders?.bottom ?? border(index === rows.length - 1 ? 'bottom' : 'insideH')),
          borderLeft: borderCss(cell.borders?.left ?? border(column === 0 ? 'left' : 'insideV')),
          borderRight: borderCss(cell.borders?.right ?? border(gridColumn === weights.length ? 'right' : 'insideV')),
        });
        appendBlocks(td, cell.paragraphs);
      }
    }
    parent.append(node);
  }

  appendBlocks(fragment, range.blocks);
  data.setData('text/html', fragment.innerHTML);
  const textBlocks = range.blocks.flatMap((block): Block[] => {
    if (block.kind === 'paragraph') return selected.has(block.id) ? [{ ...block, inlines: selected.get(block.id)! }] : [];
    if (block.kind !== 'table') return [];
    const rows = block.rows.filter(row => row.cells.some(cell => cell.paragraphs.some(item => selected.has(item.id))))
      .map(row => ({ ...row, cells: row.cells.map(cell => ({ ...cell,
        paragraphs: cell.paragraphs.filter(item => selected.has(item.id)).map(item => ({ ...item, inlines: selected.get(item.id)! })),
      })) }));
    if (wholeTable !== block && rows.length === 1 && rows[0].cells.filter(cell => cell.paragraphs.length).length === 1) {
      return rows[0].cells.flatMap(cell => cell.paragraphs);
    }
    return rows.length ? [{ ...block, rows }] : [];
  });
  data.setData('text/plain', clipboardText(textBlocks));
}
