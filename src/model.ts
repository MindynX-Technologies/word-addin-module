export type Alignment = 'left' | 'center' | 'right' | 'justify';
export type TextFormat = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  highlight?: string;
  vertical?: 'baseline' | 'superscript' | 'subscript';
  font?: string;
  size?: number;
  color?: string;
};
export type Link = { href: string; relationship?: string; anchor?: string };
export type Revision = { id: number; kind: 'ins' | 'del'; author: string; date: string };
export type TextRun = { kind: 'text'; text: string; format: TextFormat; link?: Link; source?: string; revision?: Revision };
export type CommentAnchor = { id: number; from: number; to: number };
export type DocComment = { id: number; author: string; date: string; text: string };
export type ImageRun = {
  kind: 'image';
  id: string;
  blob: Blob;
  width: number;
  height: number;
  alt: string;
};
export type BreakRun = { kind: 'pageBreak' };
export type PageField = { kind: 'pageField'; field: 'PAGE' | 'NUMPAGES'; format: TextFormat; source?: string };
export type Inline = TextRun | ImageRun | BreakRun | PageField;
export type ParagraphStyle = {
  alignment?: Alignment;
  before?: number;
  after?: number;
  lineHeight?: number;
  lineSpacing?: { rule: 'exact' | 'atLeast'; value: number };
  direction?: 'ltr' | 'rtl';
  keepNext?: boolean;
  keepLines?: boolean;
  widowControl?: boolean;
  indent?: number;
  rightIndent?: number;
  firstLine?: number;
  pageBreakBefore?: boolean;
};
export type ListStyle = { kind: 'bullet' | 'number'; id: string; level: number; start?: number; pattern?: string;
  numberFormat?: 'decimal' | 'lowerLetter' | 'upperLetter' | 'lowerRoman' | 'upperRoman' };

export function formatListNumber(value: number, format: ListStyle['numberFormat'] = 'decimal') {
  if (value < 1 || !Number.isFinite(value) || format === 'decimal') return String(value);
  if (format === 'lowerLetter' || format === 'upperLetter') {
    let text = '', number = value;
    while (number > 0) { number--; text = String.fromCharCode(65 + number % 26) + text; number = Math.floor(number / 26); }
    return format === 'lowerLetter' ? text.toLowerCase() : text;
  }
  if (value > 3999) return String(value);
  let text = '', number = value;
  for (const [digit, symbol] of [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']] as const) {
    while (number >= digit) { text += symbol; number -= digit; }
  }
  return format === 'lowerRoman' ? text.toLowerCase() : text;
}
export type Paragraph = {
  kind: 'paragraph';
  id: string;
  inlines: Inline[];
  styleId: string;
  style: ParagraphStyle;
  list?: ListStyle;
  comments?: CommentAnchor[];
  locked?: string;
};
export type CellAlignment = 'top' | 'center' | 'bottom';
export type TableBorderSide = 'top' | 'left' | 'bottom' | 'right' | 'insideH' | 'insideV';
export const TABLE_BORDER_SIDES: TableBorderSide[] = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];
export type TableBorder = { style: 'none' | 'solid' | 'dashed' | 'dotted' | 'double'; color: string; size: number };
export type TableBorders = Partial<Record<TableBorderSide, TableBorder>>;
export type CellMargins = Partial<Record<'top' | 'right' | 'bottom' | 'left', number>>;
export type TableCell = { id: string; paragraphs: Paragraph[]; colSpan?: number; width?: number; fill?: string; vertical?: CellAlignment; margins?: CellMargins; borders?: TableBorders };
export type TableRow = { id: string; cells: TableCell[]; header?: boolean; height?: number; exactHeight?: boolean };
export type Table = {
  kind: 'table'; id: string; rows: TableRow[];
  width?: { value: number; unit: 'percent' | 'px' };
  align?: 'left' | 'center' | 'right';
  margins?: CellMargins;
  borders?: TableBorders;
  columnWidths?: number[];
  structuralLock?: boolean;
};
export const DEFAULT_TABLE_BORDER: TableBorder = { style: 'solid', color: '#cbd3df', size: 0.75 };
export const DEFAULT_CELL_MARGINS = { top: 5, right: 8, bottom: 5, left: 8 };
export type TableFormatting = {
  table?: Partial<Pick<Table, 'width' | 'align' | 'margins' | 'borders'>>;
  cells?: { scope: 'cell' | 'row' | 'column' | 'table'; fill?: string; vertical?: CellAlignment };
  rowHeight?: number;
  repeatHeader?: boolean;
  columnPercent?: number | 'equal';
};
export type OpaqueBlock = { kind: 'opaque'; id: string; label: string };
export type Block = Paragraph | Table | OpaqueBlock;
export type PageSetup = { width: number; height: number; top: number; right: number; bottom: number; left: number; header: number; footer: number };
export const REGION_KINDS = ['header', 'footer'] as const;
export const REGION_VARIANTS = ['default', 'first', 'even'] as const;
export type RegionKind = typeof REGION_KINDS[number];
export type RegionVariant = typeof REGION_VARIANTS[number];
export type RegionKey = `${RegionKind}:${RegionVariant}`;
export type PageRegion = { blocks: Block[]; locked?: boolean };
export type PageRegions = {
  first: boolean;
  even: boolean;
  items: Partial<Record<RegionKey, PageRegion>>;
};
export type NamedStyle = {
  id: string;
  name: string;
  type: 'paragraph' | 'character';
  paragraph: ParagraphStyle;
  format: TextFormat;
  heading?: number;
  basedOn?: string;
  local?: { paragraph: ParagraphStyle; format: TextFormat; outline?: number };
};
export type DocModel = {
  blocks: Block[];
  page: PageSetup;
  styles: NamedStyle[];
  defaultStyleId: string;
  regions: PageRegions;
  comments: DocComment[];
  tracking: boolean;
  defaultTabStop?: number;
};
export type Position = { id: string; offset: number; affinity?: 'before' | 'after' };
export type ModelSelection = { anchor: Position; focus: Position; tableId?: string };
export type Transaction = { model: DocModel; selection: ModelSelection; typing?: boolean };

let sequence = 0;
export const nodeId = (prefix: string) => `${prefix}-${++sequence}`;

export const DEFAULT_PAGE: PageSetup = { width: 816, height: 1056, top: 96, right: 96, bottom: 96, left: 96, header: 48, footer: 48 };
export const DEFAULT_FORMAT: TextFormat = { font: 'Arial', size: 11, color: '#24272d' };
export const DEFAULT_STYLES: NamedStyle[] = [
  { id: 'Normal', name: 'Normal', type: 'paragraph', paragraph: { after: 9, lineHeight: 1.35 }, format: DEFAULT_FORMAT },
  { id: 'Title', name: 'Title', type: 'paragraph', paragraph: { after: 14, lineHeight: 1.14 }, format: { font: 'Georgia', size: 34, color: '#17263b' } },
  { id: 'Heading1', name: 'Heading 1', type: 'paragraph', paragraph: { before: 17, after: 6, lineHeight: 1.22 }, format: { font: 'Arial', bold: true, size: 17, color: '#17263b' }, heading: 1 },
  { id: 'Heading2', name: 'Heading 2', type: 'paragraph', paragraph: { before: 12, after: 5, lineHeight: 1.25 }, format: { font: 'Arial', bold: true, size: 13, color: '#244467' }, heading: 2 },
  ...Array.from({ length: 7 }, (_, index): NamedStyle => {
    const level = index + 3;
    return { id: 'Heading' + level, name: 'Heading ' + level, type: 'paragraph', heading: level,
      paragraph: { before: 10, after: 5, lineHeight: 1.22 },
      format: { ...DEFAULT_FORMAT, size: level < 5 ? 12 : 11, bold: true, color: '#244467' } };
  }),
];

export function paragraph(styleId: string, inlines: Inline[] = []): Paragraph {
  return { kind: 'paragraph', id: nodeId('p'), styleId, style: {}, inlines };
}

export function mergeParagraphStyle(base: ParagraphStyle = {}, local: ParagraphStyle = {}): ParagraphStyle {
  return { ...base, ...local, ...(local.lineSpacing ? { lineHeight: undefined } : local.lineHeight !== undefined ? { lineSpacing: undefined } : {}) };
}

export function blankModel(): DocModel {
  return { blocks: [paragraph('Normal')], page: DEFAULT_PAGE, styles: DEFAULT_STYLES, defaultStyleId: 'Normal', regions: { first: false, even: false, items: {} }, comments: [], tracking: false };
}

export function regionForPage(regions: PageRegions, kind: RegionKind, page: number): RegionKey {
  return `${kind}:${regions.first && page === 1 ? 'first' : regions.even && page % 2 === 0 ? 'even' : 'default'}`;
}

export function blockParagraphs(blocks: Block[]) {
  return blocks.flatMap(block => {
    if (block.kind === 'paragraph') return [block];
    if (block.kind === 'table') return block.rows.flatMap(row => row.cells.flatMap(cell => cell.paragraphs));
    return [];
  });
}

export function paragraphs(model: DocModel) {
  return blockParagraphs([...model.blocks, ...Object.values(model.regions.items).flatMap(region => region.blocks)]);
}

const wordCounts = new WeakMap<Paragraph, number>();
const wordCounter = new Intl.Segmenter(undefined, { granularity: 'word' });
export function paragraphWordCount(item: Paragraph) {
  let count = wordCounts.get(item);
  if (count === undefined) {
    count = Array.from(wordCounter.segment(plainText(item.inlines))).filter(segment => segment.isWordLike).length;
    wordCounts.set(item, count);
  }
  return count;
}

export function paragraphScope(model: DocModel, id: string) {
  for (const [key, region] of Object.entries(model.regions.items)) {
    if (blockParagraphs(region.blocks).some(item => item.id === id)) return { key: key as RegionKey, blocks: region.blocks };
  }
  return { key: undefined, blocks: model.blocks };
}

function inScope(model: DocModel, key: RegionKey | undefined, blocks: Block[]): DocModel {
  return key ? { ...model, regions: { ...model.regions, items: {
    ...model.regions.items, [key]: { ...model.regions.items[key], blocks },
  } } } : { ...model, blocks };
}

export function editableStart(model: DocModel): ModelSelection | null {
  const items = blockParagraphs(model.blocks);
  const first = items.find(item => !item.locked) ?? items[0];
  return first ? caret({ id: first.id, offset: 0 }) : null;
}

export function caret(position: Position): ModelSelection {
  return { anchor: position, focus: position };
}

export function collapsed(selection: ModelSelection) {
  return !selection.tableId && selection.anchor.id === selection.focus.id && selection.anchor.offset === selection.focus.offset;
}

export function inlineLength(inline: Inline) {
  return inline.kind === 'text' ? inline.text.length : 1;
}

export function lengthOf(inlines: Inline[]) {
  return inlines.reduce((length, inline) => length + inlineLength(inline), 0);
}

export function sliceInlines(inlines: Inline[], from: number, to = Infinity) {
  let offset = 0;
  const slice: Inline[] = [];
  for (const inline of inlines) {
    const end = offset + inlineLength(inline);
    if (from < end && to > offset) {
      if (inline.kind !== 'text' || (from <= offset && to >= end)) slice.push(inline);
      else slice.push({ ...inline, text: inline.text.slice(Math.max(from - offset, 0), to - offset) });
    }
    offset = end;
    if (offset >= to) break;
  }
  return slice;
}

function sameFormat(a: TextFormat, b: TextFormat) {
  return a.bold === b.bold && a.italic === b.italic && a.underline === b.underline &&
    a.strike === b.strike && a.highlight === b.highlight && a.vertical === b.vertical && a.font === b.font && a.size === b.size && a.color === b.color;
}

export function normalize(inlines: Inline[]) {
  const result: Inline[] = [];
  for (const inline of inlines) {
    if (inline.kind === 'text' && !inline.text) continue;
    const previous = result.at(-1);
    if (inline.kind === 'text' && previous?.kind === 'text' && inline.source === previous.source &&
      inline.link === previous.link && inline.revision === previous.revision && sameFormat(inline.format, previous.format)) {
      result[result.length - 1] = { ...previous, text: previous.text + inline.text };
    } else {
      result.push(inline);
    }
  }
  return result;
}

export function orderedSelection(model: DocModel, selection: ModelSelection) {
  const scope = paragraphScope(model, selection.anchor.id);
  if (scope.key && model.regions.items[scope.key]?.locked) return null;
  const items = blockParagraphs(scope.blocks);
  const anchor = items.findIndex(item => item.id === selection.anchor.id);
  const focus = items.findIndex(item => item.id === selection.focus.id);
  if (anchor < 0 || focus < 0) return null;
  const forward = anchor < focus || (anchor === focus && selection.anchor.offset <= selection.focus.offset);
  const [start, end] = forward ? [selection.anchor, selection.focus] : [selection.focus, selection.anchor];
  const selected = items.slice(Math.min(anchor, focus), Math.max(anchor, focus) + 1);
  return { start, end, selected, ...scope };
}

function hasProtectedRange(blocks: Block[], selected: Paragraph[]) {
  if (selected.some(item => item.locked)) return true;
  if (selected.length < 2) return false;
  const indices = selected.map(item => blocks.findIndex(block =>
    block.id === item.id || (block.kind === 'table' && block.rows.some(row =>
      row.cells.some(cell => cell.paragraphs.some(candidate => candidate.id === item.id))))));
  return blocks.slice(indices[0], indices.at(-1)! + 1).some(block => block.kind === 'opaque');
}

function mapParagraphs(blocks: Block[], updates: Map<string, Paragraph>) {
  return blocks.map(block => {
    if (block.kind === 'paragraph') return updates.get(block.id) ?? block;
    if (block.kind !== 'table') return block;
    let changed = false;
    const rows = block.rows.map(row => {
      const cells = row.cells.map(cell => {
        const next = cell.paragraphs.map(item => updates.get(item.id) ?? item);
        if (next.every((item, index) => item === cell.paragraphs[index])) return cell;
        changed = true;
        return { ...cell, paragraphs: next };
      });
      return cells.every((cell, index) => cell === row.cells[index]) ? row : { ...row, cells };
    });
    return changed ? { ...block, rows } : block;
  });
}


export function withParagraphUpdates(model: DocModel, updates: Map<string, Paragraph>): DocModel {
  const items = { ...model.regions.items };
  for (const [key, region] of Object.entries(items)) {
    const blocks = mapParagraphs(region.blocks, updates);
    if (blocks.some((block, index) => block !== region.blocks[index])) items[key as RegionKey] = { ...region, blocks };
  }
  return { ...model, blocks: mapParagraphs(model.blocks, updates), regions: { ...model.regions, items } };
}

export function updateParagraphs(model: DocModel, selection: ModelSelection, update: (item: Paragraph) => Paragraph): Transaction | null {
  const range = orderedSelection(model, selection);
  if (!range || hasProtectedRange(range.blocks, range.selected)) return null;
  return { model: inScope(model, range.key, mapParagraphs(range.blocks, new Map(range.selected.map(item => [item.id, update(item)])))), selection };
}

export function formatRange(model: DocModel, selection: ModelSelection, update: (run: TextRun) => TextRun): Transaction | null {
  if (collapsed(selection)) return null;
  const range = orderedSelection(model, selection);
  if (!range || hasProtectedRange(range.blocks, range.selected)) return null;
  return updateParagraphs(model, selection, item => {
    const from = item.id === range.start.id ? range.start.offset : 0;
    const to = item.id === range.end.id ? range.end.offset : lengthOf(item.inlines);
    const selected = sliceInlines(item.inlines, from, to);
    if (!selected.some(run => run.kind === 'text')) return item;
    return {
      ...item,
      inlines: normalize([
        ...sliceInlines(item.inlines, 0, from),
        ...selected.map(run => run.kind === 'text' ? update(run) : run),
        ...sliceInlines(item.inlines, to),
      ]),
    };
  });
}

function spliceParagraphs(blocks: Block[], selected: Paragraph[], replacement: Paragraph[]): Block[] | null {
  const first = blocks.findIndex(block => block.id === selected[0].id);
  const last = blocks.findIndex(block => block.id === selected.at(-1)!.id);
  if (first >= 0 && last >= first && blocks.slice(first, last + 1).every(block => block.kind !== 'opaque')) {
    return [...blocks.slice(0, first), ...replacement, ...blocks.slice(last + 1)];
  }
  for (const block of blocks) {
    if (block.kind !== 'table') continue;
    for (const row of block.rows) {
      for (const cell of row.cells) {
        const from = cell.paragraphs.indexOf(selected[0]);
        const to = cell.paragraphs.indexOf(selected.at(-1)!);
        if (from < 0 || to < from) continue;
        return blocks.map(item => item !== block ? item : {
          ...block,
          rows: block.rows.map(itemRow => itemRow !== row ? itemRow : {
            ...row,
            cells: row.cells.map(itemCell => itemCell !== cell ? itemCell : {
              ...cell,
              paragraphs: [...cell.paragraphs.slice(0, from), ...replacement, ...cell.paragraphs.slice(to + 1)],
            }),
          }),
        });
      }
    }
  }
  return null;
}

export function revisionAtBoundary(inlines: Inline[], offset: number) {
  if (!offset) return false;
  const before = sliceInlines(inlines, offset - 1, offset)[0];
  const after = sliceInlines(inlines, offset, offset + 1)[0];
  return before?.kind === 'text' && after?.kind === 'text' && !!before.revision && before.revision.id === after.revision?.id;
}

export function replaceRange(
  model: DocModel, selection: ModelSelection, lines: Inline[][], typing = false,
): Transaction | null {
  const range = orderedSelection(model, selection);
  if (!range || hasProtectedRange(range.blocks, range.selected)) return null;
  const first = range.selected[0];
  const last = range.selected.at(-1)!;
  if ((first !== last || lines.length > 1) && range.selected.some(item => item.comments?.length ||
    item.inlines.some(run => run.kind === 'text' && run.revision))) return null;
  const affected = sliceInlines(first.inlines, range.start.offset, first === last ? range.end.offset : Infinity);
  if (affected.some(run => run.kind === 'text' && run.revision) || revisionAtBoundary(first.inlines, range.start.offset)) return null;
  const prefix = sliceInlines(first.inlines, 0, range.start.offset);
  const suffix = sliceInlines(last.inlines, range.end.offset);
  const replacement = lines.map((line, index) => {
    const base = index === 0 ? first : {
      ...first, id: nodeId('p'),
      style: { ...first.style, pageBreakBefore: undefined },
      styleId: index === lines.length - 1 && range.start.offset === lengthOf(first.inlines) &&
        model.styles.some(style => style.id === first.styleId && (style.heading || style.id === 'Title'))
        ? model.defaultStyleId : first.styleId,
    };
    return {
      ...base,
      comments: base.comments?.map(mark => {
        const delta = lengthOf(lines[0]) - (range.end.offset - range.start.offset);
        const map = (offset: number, end: boolean) => offset < range.start.offset ||
          (offset === range.start.offset && !end) ? offset : offset >= range.end.offset
          ? offset + delta : range.start.offset + (end ? lengthOf(lines[0]) : 0);
        return { ...mark, from: map(mark.from, false), to: map(mark.to, true) };
      }),
      inlines: normalize([
        ...(index === 0 ? prefix : []), ...line, ...(index === lines.length - 1 ? suffix : []),
      ]),
    };
  });
  const updated = spliceParagraphs(range.blocks, range.selected, replacement);
  if (!updated) return null;
  return {
    model: inScope(model, range.key, updated), typing,
    selection: caret({
      id: replacement.at(-1)!.id,
      offset: lengthOf(lines.at(-1)!) + (lines.length === 1 ? lengthOf(prefix) : 0),
    }),
  };
}

export function insertTable(model: DocModel, selection: ModelSelection, rows: number, columns: number): Transaction | null {
  const range = orderedSelection(model, selection);
  if (!range || hasProtectedRange(range.blocks, range.selected)) return null;
  const index = range.blocks.findIndex(block => block.id === range.end.id);
  if (index < 0) return null;
  const table: Table = {
    kind: 'table', id: nodeId('table'),
    rows: Array.from({ length: rows }, () => ({
      id: nodeId('row'),
      cells: Array.from({ length: columns }, () => ({
        id: nodeId('cell'), paragraphs: [paragraph(model.defaultStyleId)],
      })),
    })),
  };
  const blocks = [...range.blocks];
  blocks.splice(index + 1, 0, table);
  if (index + 2 === blocks.length) blocks.push(paragraph(model.defaultStyleId));
  return { model: inScope(model, range.key, blocks), selection: caret({ id: table.rows[0].cells[0].paragraphs[0].id, offset: 0 }) };
}

export function plainText(inlines: Inline[]) {
  return inlines.map(inline => inline.kind === 'text' ? inline.text : inline.kind === 'image' ? inline.alt : inline.kind === 'pageField' ? inline.field : '\n').join('');
}

export function activeFormat(model: DocModel, selection: ModelSelection) {
  const item = paragraphs(model).find(candidate => candidate.id === selection.focus.id);
  const named = model.styles.find(style => style.id === item?.styleId) ??
    model.styles.find(style => style.id === model.defaultStyleId);
  const offset = selection.focus.offset;
  const atStart = offset === 0;
  const run = item && sliceInlines(item.inlines, atStart ? 0 : offset - 1, atStart ? 1 : offset)[0];
  return { item, format: { ...DEFAULT_FORMAT, ...named?.format, ...(run?.kind === 'text' ? run.format : {}) },
    direct: run?.kind === 'text' ? run.format : {}, source: run?.kind === 'text' ? run.source : undefined,
    link: offset > 0 && offset < lengthOf(item?.inlines ?? []) && run?.kind === 'text' ? run.link : undefined };
}

const graphemes = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
const words = new Intl.Segmenter(undefined, { granularity: 'word' });

export function deletionRange(model: DocModel, selection: ModelSelection, backward: boolean, byWord: boolean) {
  if (!collapsed(selection)) return selection;
  const items = blockParagraphs(paragraphScope(model, selection.focus.id).blocks);
  const index = items.findIndex(item => item.id === selection.focus.id);
  const item = items[index];
  if (!item || item.locked) return null;
  const { offset } = selection.focus;
  const length = lengthOf(item.inlines);
  if ((backward && offset === 0) || (!backward && offset === length)) {
    const next = items[index + (backward ? -1 : 1)];
    return next ? { anchor: selection.anchor, focus: { id: next.id, offset: backward ? lengthOf(next.inlines) : 0 } } : null;
  }
  const text = item.inlines.map(run => run.kind === 'text' ? run.text : '\uFFFC').join('');
  const segmenter = byWord ? words : graphemes;
  const segment = segmenter.segment(text).containing(backward ? offset - 1 : offset);
  if (!segment) return null;
  return {
    anchor: selection.anchor,
    focus: { id: item.id, offset: backward ? segment.index : segment.index + segment.segment.length },
  };
}

export function updateNamedStyle(styles: NamedStyle[], edited: NamedStyle): NamedStyle[] {
  const changed = new Map([[edited.id, edited]]);
  const original = new Map(styles.map(style => [style.id, style]));
  function resolve(style: NamedStyle, seen = new Set<string>()): NamedStyle {
    const cached = changed.get(style.id);
    if (cached) return cached;
    if (!style.basedOn || !style.local || seen.has(style.id)) return style;
    const previousBase = original.get(style.basedOn);
    const base = previousBase && resolve(previousBase, new Set([...seen, style.id]));
    if (!base || base === previousBase) return style;
    const next = {
      ...style, paragraph: mergeParagraphStyle(base.paragraph, style.local.paragraph),
      format: { ...base.format, ...style.local.format },
      heading: style.local.outline === undefined ? base.heading : style.local.outline < 9 ? style.local.outline + 1 : undefined,
    };
    changed.set(style.id, next);
    return next;
  }
  return styles.some(style => style.id === edited.id) ? styles.map(style => resolve(style)) : [...styles, edited];
}

export function tableAt(model: DocModel, selection: ModelSelection) {
  const scope = paragraphScope(model, selection.focus.id);
  for (const table of scope.blocks) {
    if (table.kind !== 'table') continue;
    for (const [rowIndex, row] of table.rows.entries()) {
      const columnIndex = row.cells.findIndex(cell => cell.paragraphs.some(item => item.id === selection.focus.id));
      if (columnIndex >= 0) return { ...scope, table, rowIndex, columnIndex };
    }
  }
}

export function selectWholeTable(table: Table): ModelSelection {
  const first = table.rows[0].cells[0].paragraphs[0];
  const last = table.rows.at(-1)!.cells.at(-1)!.paragraphs.at(-1)!;
  return { tableId: table.id, anchor: { id: first.id, offset: 0 }, focus: { id: last.id, offset: lengthOf(last.inlines) } };
}

export function wholeTableSelection(model: DocModel, selection: ModelSelection) {
  const table = tableAt(model, selection)?.table;
  if (!table) return;
  if (selection.tableId === table.id) return table;
  const { anchor, focus } = selectWholeTable(table);
  // A selection within one paragraph is ordinary text, even if its cell fills the table.
  if (anchor.id === focus.id) return;
  const matches = (a: Position, b: Position) => a.id === b.id && a.offset === b.offset;
  if ((matches(selection.anchor, anchor) && matches(selection.focus, focus)) ||
    (matches(selection.anchor, focus) && matches(selection.focus, anchor))) return table;
}

export type TableAction = 'row-before' | 'row-after' | 'column-before' | 'column-after' | 'delete-row' | 'delete-column' | 'delete-table';

export function editTable(model: DocModel, selection: ModelSelection, action: TableAction): Transaction | null {
  const context = tableAt(model, selection);
  if (!context || context.key && model.regions.items[context.key]?.locked) return null;
  const { table, rowIndex, columnIndex, blocks, key } = context;
  if (table.structuralLock && action !== 'delete-table') return null;
  const allCells = table.rows.flatMap(row => row.cells);
  const removed = action === 'delete-table' || (action === 'delete-row' && table.rows.length === 1) ||
    (action === 'delete-column' && table.rows[0].cells.length === 1);
  const affectedCells = removed ? allCells : action === 'delete-row' ? table.rows[rowIndex].cells :
    action === 'delete-column' ? table.rows.map(row => row.cells[columnIndex]) : [];
  if (affectedCells.some(cell => cell.paragraphs.some(item => item.locked || item.comments?.length ||
    item.inlines.some(run => run.kind === 'text' && run.revision)))) return null;

  if (removed) {
    const index = blocks.indexOf(table);
    const adjacent = blocks[index + 1];
    const empty = adjacent?.kind === 'paragraph' && !adjacent.locked ? adjacent : paragraph(model.defaultStyleId);
    const replacement = adjacent === empty ? [] : [empty];
    return { model: inScope(model, key, [...blocks.slice(0, index), ...replacement, ...blocks.slice(index + 1)]),
      selection: caret({ id: empty.id, offset: 0 }) };
  }

  const newCell = (width?: number): TableCell => ({ id: nodeId('cell'), width, paragraphs: [paragraph(model.defaultStyleId)] });
  let rows = table.rows;
  let row = rowIndex;
  let column = columnIndex;
  if (action === 'row-before' || action === 'row-after') {
    row += action === 'row-after' ? 1 : 0;
    rows = [...rows.slice(0, row), {
      id: nodeId('row'), cells: table.rows[rowIndex].cells.map(cell => newCell(cell.width)),
    }, ...rows.slice(row)];
  } else if (action === 'column-before' || action === 'column-after') {
    if (table.rows[0].cells.length >= 20) return null;
    column += action === 'column-after' ? 1 : 0;
    rows = rows.map(item => ({ ...item, cells: [
      ...item.cells.slice(0, column), newCell(item.cells[columnIndex].width), ...item.cells.slice(column),
    ] }));
  } else if (action === 'delete-row') {
    rows = rows.filter((_, index) => index !== rowIndex);
    row = Math.min(row, rows.length - 1);
  } else {
    rows = rows.map(item => ({ ...item, cells: item.cells.filter((_, index) => index !== columnIndex) }));
    column = Math.min(column, rows[0].cells.length - 1);
  }

  return { model: inScope(model, key, blocks.map(block => block === table ? { ...table, rows } : block)),
    selection: caret({ id: rows[row].cells[column].paragraphs[0].id, offset: 0 }) };
}

export function tableHeaderCount(table: Table) {
  const firstBody = table.rows.findIndex(row => !row.header);
  return firstBody < 0 ? table.rows.length : firstBody;
}

export function tableColumnWeights(table: Table) {
  if (table.columnWidths) return table.columnWidths;
  return table.rows[0].cells.map(cell => cell.width ?? 1);
}

export function formatTable(model: DocModel, selection: ModelSelection, change: TableFormatting): Transaction | null {
  const context = tableAt(model, selection);
  if (!context || context.key && model.regions.items[context.key]?.locked) return null;
  const { table, rowIndex, columnIndex, blocks, key } = context;
  if (table.structuralLock && (change.table || change.columnPercent !== undefined || change.rowHeight !== undefined ||
    change.repeatHeader !== undefined || change.cells?.scope === 'column')) return null;
  let widths: number[] | undefined;
  if (change.columnPercent === 'equal') widths = tableColumnWeights(table).map(() => 1);
  else if (change.columnPercent !== undefined) {
    const previous = tableColumnWeights(table);
    const remaining = previous.reduce((sum, value, index) => index === columnIndex ? sum : sum + value, 0);
    const percent = Math.max(2, Math.min(98, change.columnPercent));
    if (remaining) widths = previous.map((value, index) => index === columnIndex ? percent : value / remaining * (100 - percent));
  }
  const cells = change.cells;
  const replaceBorders = change.table?.borders !== undefined;
  const headerCount = change.repeatHeader ? Math.max(1, tableHeaderCount(table)) : 0;
  function updateRow(row: TableRow, index: number) {
    const formatCells = cells && (cells.scope === 'table' || cells.scope === 'column' || index === rowIndex);
    const height = change.rowHeight !== undefined && index === rowIndex ? change.rowHeight || undefined : row.height;
    const exactHeight = change.rowHeight !== undefined && index === rowIndex ? undefined : row.exactHeight;
    const header = change.repeatHeader === undefined ? row.header : index < headerCount;
    const updated = widths || formatCells || replaceBorders ? row.cells.map((cell, column) => {
      const cellWidth = widths?.[column] ?? cell.width;
      const inScope = formatCells && (cells.scope === 'table' || cells.scope === 'row' || column === columnIndex);
      const cellFormat = inScope && !cell.paragraphs.some(paragraph => paragraph.locked) ? cells : undefined;
      const fill = cellFormat && 'fill' in cellFormat ? cellFormat.fill : cell.fill;
      const vertical = cellFormat?.vertical ?? cell.vertical;
      const borders = replaceBorders ? undefined : cell.borders;
      return cellWidth === cell.width && fill === cell.fill && vertical === cell.vertical && borders === cell.borders
        ? cell : { ...cell, width: cellWidth, fill, vertical, borders };
    }) : row.cells;
    return height === row.height && exactHeight === row.exactHeight && !!header === !!row.header &&
      updated.every((cell, column) => cell === row.cells[column]) ? row : { ...row, cells: updated, height, exactHeight, header };
  }
  let rows = table.rows;
  if (widths || replaceBorders || change.repeatHeader !== undefined || cells?.scope === 'table' || cells?.scope === 'column') {
    const updated = table.rows.map(updateRow);
    if (updated.some((row, index) => row !== table.rows[index])) rows = updated;
  } else if (cells || change.rowHeight !== undefined) {
    const updated = updateRow(table.rows[rowIndex], rowIndex);
    if (updated !== table.rows[rowIndex]) {
      rows = table.rows.slice();
      rows[rowIndex] = updated;
    }
  }
  const properties = change.table ?? {};
  const tableChanged = (Object.keys(properties) as Array<keyof typeof properties>).some(property => properties[property] !== table[property]);
  if (!tableChanged && rows === table.rows) return null;
  const next = { ...table, ...properties, rows };
  return { model: inScope(model, key, blocks.map(block => block === table ? next : block)), selection };
}

export function shadeCell(model: DocModel, selection: ModelSelection, fill?: string): Transaction | null {
  return formatTable(model, selection, { cells: { scope: wholeTableSelection(model, selection) ? 'table' : 'cell', fill } });
}

export function imageAt(model: DocModel, id: string) {
  for (const item of paragraphs(model)) {
    let offset = 0;
    for (const inline of item.inlines) {
      if (inline.kind === 'image' && inline.id === id) {
        return { image: inline, selection: caret({ id: item.id, offset }), paragraph: item };
      }
      offset += inlineLength(inline);
    }
  }
}

export function pasteBlocks(model: DocModel, selection: ModelSelection, incoming: Block[]): Transaction | null {
  const range = orderedSelection(model, selection);
  if (!range || !incoming.length || hasProtectedRange(range.blocks, range.selected)) return null;
  const first = range.selected[0];
  const last = range.selected.at(-1)!;
  // A single normal paragraph copied from Word is usually a text fragment, even
  // when its HTML wrapper carries paragraph spacing and alignment.
  if (first === last && first.inlines.length && incoming.length === 1 && incoming[0].kind === 'paragraph' &&
    !incoming[0].list && incoming[0].styleId === model.defaultStyleId && !incoming[0].style.pageBreakBefore) {
    return replaceRange(model, selection, [incoming[0].inlines]);
  }
  const hasStructure = incoming.length > 1 || incoming[0].kind !== 'paragraph' ||
    incoming[0].list || incoming[0].styleId !== model.defaultStyleId || Object.keys(incoming[0].style).length > 0;
  if (!hasStructure && incoming[0].kind === 'paragraph') return replaceRange(model, selection, [incoming[0].inlines]);
  if (range.selected.some(item => item.comments?.length || item.inlines.some(run => run.kind === 'text' && run.revision))) return null;
  const prefix = sliceInlines(first.inlines, 0, range.start.offset);
  const suffix = sliceInlines(last.inlines, range.end.offset);
  const leading = { ...first, inlines: prefix };
  const trailing = { ...last, id: first === last ? nodeId('p') : last.id, inlines: suffix };
  const replacement = [...(prefix.length ? [leading] : []), ...incoming, ...(suffix.length ? [trailing] : [])];
  let end = replacement.at(-1)!;
  if (end.kind !== 'paragraph') {
    end = paragraph(model.defaultStyleId);
    replacement.push(end);
  }
  const from = range.blocks.indexOf(first);
  const to = range.blocks.indexOf(last);
  const target = caret({ id: end.id, offset: end === trailing ? 0 : lengthOf(end.inlines) });
  if (from >= 0 && to >= from && range.blocks.slice(from, to + 1).every(block => block.kind === 'paragraph')) {
    return { model: inScope(model, range.key, [...range.blocks.slice(0, from), ...replacement, ...range.blocks.slice(to + 1)]),
      selection: target };
  }
  if (replacement.every((block): block is Paragraph => block.kind === 'paragraph')) {
    const blocks = spliceParagraphs(range.blocks, range.selected, replacement);
    if (blocks) return { model: inScope(model, range.key, blocks), selection: target };
  }
  return null;
}
