import { Move } from 'lucide-react';
import { createContext, memo, useContext, useEffect, useState } from 'react';
import type { CSSProperties, RefObject } from 'react';
import { DEFAULT_FORMAT, DEFAULT_TABLE_BORDER, DEFAULT_CELL_MARGINS, TABLE_BORDER_SIDES, blockParagraphs, formatListNumber, inlineLength, lengthOf, mergeParagraphStyle, sliceInlines, tableColumnWeights, tableHeaderCount } from './model';
import type { Block, DocModel, ImageRun, Inline, ModelSelection, Paragraph, Position, Table, TableBorder, TableBorderSide, TextFormat } from './model';

type ParagraphFragment = { kind: 'paragraph'; paragraph: Paragraph; from: number; to: number; label?: string };
type TableFragment = { kind: 'table'; table: Table; from: number; to: number; labels: Map<string, string>; repeated?: number; measuring?: boolean };
type OpaqueFragment = { kind: 'opaque'; id: string; label: string };
export type Fragment = ParagraphFragment | TableFragment | OpaqueFragment;
export type Page = Fragment[];
export const PageNumbers = createContext({ page: 1, total: 1 });
export const SelectedPicture = createContext<string | undefined>(undefined);

type Measure = {
  key: string;
  fragment: Fragment;
  force?: boolean;
  explicitBreak?: boolean;
  keepNext?: boolean;
  keepLines?: boolean;
  widowControl?: boolean;
};
const measureKey = (id: string, offset: number) => id + ':' + offset;
// Leave room for both strokes and their gap after the browser snaps the border to pixels.
const DOUBLE_BORDER_MIN_PX = 5;
const DOUBLE_BORDER_MIN_PT = .25;
const extraDoubleBorderPoints = (size: number) => Math.max(0, size - DOUBLE_BORDER_MIN_PT) / 2;

function borderWidth(border: TableBorder | undefined, scale: number) {
  if (!border || border.style === 'none') return 0;
  const pixels = border.size * 4 / 3;
  return border.style === 'double'
    ? Math.max(pixels, DOUBLE_BORDER_MIN_PX / scale + extraDoubleBorderPoints(border.size) * 4 / 3) : pixels;
}

export function buildMeasures(model: DocModel, blocks: Block[] = model.blocks, pageBreaks = true): Measure[] {
  const labels = new Map<string, string>();
  const listCounts = new Map<string, number>();
  const listFormats = new Map(blockParagraphs(blocks).flatMap(item => item.list ? [[item.list.id + ':' + item.list.level, item.list.numberFormat] as const] : []));
  function labelFor(item: Paragraph) {
    const list = item.list;
    if (!list) return;
    if (list.kind === 'bullet') return '•';
    const key = list.id + ':' + list.level;
    listCounts.set(key, (listCounts.get(key) ?? ((list.start ?? 1) - 1)) + 1);
    for (let level = list.level + 1; level < 9; level++) listCounts.delete(list.id + ':' + level);
    return (list.pattern ?? `%${list.level + 1}.`).replace(/%([1-9])/g, (_, level: string) => {
      const key = list.id + ':' + (Number(level) - 1);
      return formatListNumber(listCounts.get(key) ?? 1, listFormats.get(key));
    });
  }
  for (const item of blockParagraphs(blocks)) {
    const label = labelFor(item);
    if (label) labels.set(item.id, label);
  }
  const measures: Measure[] = [];
  for (const block of blocks) {
    if (block.kind === 'opaque') {
      measures.push({ key: block.id, fragment: block });
    } else if (block.kind === 'table') {
      block.rows.forEach((_, row) => measures.push({
        key: measureKey(block.id, row),
        fragment: { kind: 'table', table: block, from: row, to: row + 1, labels, measuring: true },
      }));
    } else {
      const label = labels.get(block.id);
      let from = 0;
      let offset = 0;
      let force = block.style.pageBreakBefore ?? model.styles.find(style => style.id === block.styleId)?.paragraph.pageBreakBefore;
      let explicitBreak = false;
      const style = mergeParagraphStyle(model.styles.find(style => style.id === block.styleId)?.paragraph, block.style);
      const append = () => {
        measures.push({
          key: measureKey(block.id, from),
          fragment: { kind: 'paragraph', paragraph: block, from, to: offset, label: from === 0 ? label : undefined },
          force, explicitBreak, keepNext: style.keepNext, keepLines: style.keepLines, widowControl: style.widowControl,
        });
      };
      for (const inline of block.inlines) {
        if (pageBreaks && inline.kind === 'pageBreak') {
          // A page-break-only paragraph has no visible line on the preceding page.
          if (offset !== from || from !== 0 || force) append();
          from = offset + 1;
          force = true;
          explicitBreak = true;
        }
        offset += inlineLength(inline);
      }
      append();
    }
  }
  return measures;
}

function textPoint(block: HTMLElement, offset: number, towardEnd: boolean): [Node, number] | undefined {
  const inlines = Array.from(block.querySelectorAll<HTMLElement>('[data-inline-start]'));
  const matching = inlines.filter(inline => Number(inline.dataset.inlineStart) <= offset && offset <= Number(inline.dataset.inlineEnd));
  const inline = towardEnd ? matching.at(-1) : matching[0];
  if (!inline) {
    if (!inlines.length) return [block, 0];
    const index = Number(inlines[0].dataset.inlineStart) > offset ? 0 : block.childNodes.length;
    return [block, index];
  }
  const from = Number(inline.dataset.inlineStart);
  if (inline.dataset.atom !== undefined) {
    const host = inline.parentNode!;
    const index = Array.prototype.indexOf.call(host.childNodes, inline) as number;
    return [host, index + (offset > from ? 1 : 0)];
  }
  return inline.firstChild ? [inline.firstChild, offset - from] : [inline, 0];
}

function rangeBounds(block: HTMLElement, from: number, to: number) {
  const start = textPoint(block, from, true);
  const end = textPoint(block, to, false);
  if (!start || !end) return;
  const range = block.ownerDocument.createRange();
  range.setStart(...start);
  range.setEnd(...end);
  const rect = range.getBoundingClientRect();
  return rect.height ? rect : undefined;
}

type TableMeasurement = { row: Table['rows'][number]; styles: DocModel['styles']; margins: Table['margins'];
  borderWidths: string; widths: string; doubleCellBorderScale?: number; height: number };
export type TableMeasurements = Map<string, TableMeasurement>;

export function paginate(measures: Measure[], element: HTMLElement, capacity: number, styles: DocModel['styles'],
  tableMeasurements: TableMeasurements = new Map(), scale = 1) {
  const measured = new Map(Array.from(element.querySelectorAll<HTMLElement>('[data-measure]')).map(node => [node.dataset.measure, node]));
  const pages: Page[] = [[]];
  let space = capacity;
  const liveRows = new Set<string>();
  const tableMetrics = new Map<Table, { borderWidths: string; widths: string }>();
  const headerHeights = new Map<string, { rows: number; height: number }>();
  const current = () => pages.at(-1)!;
  const nextPage = () => { pages.push([]); space = capacity; };
  function place(fragment: Fragment, height: number) {
    const page = current();
    const previous = page.at(-1);
    if (fragment.kind === 'table' && previous?.kind === 'table' &&
      previous.table === fragment.table && previous.to === fragment.from) {
      page[page.length - 1] = { ...previous, to: fragment.to };
    } else {
      page.push(fragment);
    }
    space -= height;
  }
  for (let measureIndex = 0; measureIndex < measures.length; measureIndex++) {
    const measure = measures[measureIndex];
    const fragment = measure.fragment;
    const measurement = measured.get(measure.key);
    if (measure.explicitBreak || (measure.force && current().length)) nextPage();
    if (measure.keepNext && current().length) {
      let combined = measurement?.getBoundingClientRect().height || 24;
      let next = measureIndex + 1;
      while (next < measures.length && !measures[next].force && !measures[next].explicitBreak) {
        combined += measured.get(measures[next].key)?.getBoundingClientRect().height || 24;
        if (!measures[next++].keepNext || combined > capacity) break;
      }
      if (combined <= capacity + .5 && combined > space + .5) nextPage();
    }
    if (fragment.kind === 'table') {
      const { table, from } = fragment;
      let metrics = tableMetrics.get(table);
      if (!metrics) {
        metrics = {
          widths: JSON.stringify([table.width, tableColumnWeights(table), element.clientWidth]),
          borderWidths: JSON.stringify(TABLE_BORDER_SIDES.map(side => {
            const border = table.borders ? table.borders[side] : DEFAULT_TABLE_BORDER;
            return borderWidth(border, scale);
          })),
        };
        tableMetrics.set(table, metrics);
      }
      const row = table.rows[from];
      const cacheKey = row.id + (from === 0 ? ':first' : '') + (from === table.rows.length - 1 ? ':last' : '');
      liveRows.add(cacheKey);
      const previous = tableMeasurements.get(cacheKey);
      const doubleCellBorderScale = row.cells.some(cell =>
        TABLE_BORDER_SIDES.some(side => cell.borders?.[side]?.style === 'double')) ? scale : undefined;
      // Shading, alignment and the repeat flag leave the measured row dimensions alone.
      const reused = previous && previous.styles === styles && previous.margins === table.margins &&
        previous.widths === metrics.widths && previous.borderWidths === metrics.borderWidths &&
        previous.doubleCellBorderScale === doubleCellBorderScale &&
        (previous.row === row || previous.row.height === row.height && previous.row.cells.length === row.cells.length &&
          row.cells.every((cell, index) => cell.paragraphs === previous.row.cells[index].paragraphs &&
            cell.borders === previous.row.cells[index].borders && cell.margins === previous.row.cells[index].margins));
      const height = reused ? previous.height : measurement?.getBoundingClientRect().height || 24;
      if (previous?.row !== row || !reused) tableMeasurements.set(cacheKey, { row, styles, margins: table.margins, ...metrics, doubleCellBorderScale, height });
      let header = headerHeights.get(table.id);
      if (!header) {
        header = { rows: tableHeaderCount(fragment.table), height: 0 };
        headerHeights.set(fragment.table.id, header);
      }
      if (fragment.from < header.rows) header.height += height;
      let repeated: number | undefined;
      if (height > space + 0.5 && current().length) {
        nextPage();
        if (header.rows && fragment.from >= header.rows && header.height + height <= capacity + 0.5) {
          repeated = header.rows;
          space -= header.height;
        }
      }
      place({ ...fragment, measuring: false, repeated }, height);
      continue;
    }
    const box = measurement?.getBoundingClientRect();
    const height = box?.height || 24;
    if (measure.keepLines && height <= capacity + .5 && height > space + .5 && current().length) nextPage();
    if (fragment.kind !== 'paragraph' || !box || !measurement || fragment.paragraph.locked || fragment.from === fragment.to) {
      if (height > space + 0.5 && current().length) nextPage();
      place(fragment, height);
      continue;
    }
    if (height <= space + 0.5) {
      place(fragment, height);
      continue;
    }
    const block = measurement.querySelector<HTMLElement>('[data-paragraph]')!;
    const inlines = sliceInlines(fragment.paragraph.inlines, fragment.from, fragment.to);
    const text = inlines.map(inline => inline.kind === 'text' ? inline.text : '\uFFFC').join('');
    const boundaries = [fragment.from];
    for (const part of new Intl.Segmenter().segment(text)) boundaries.push(fragment.from + part.index + part.segment.length);
    let start = 0;
    let top = box.top;
    while (start < boundaries.length - 1) {
      if (box.bottom - top <= space + 0.5) {
        place({ ...fragment, from: boundaries[start], label: start === 0 ? fragment.label : undefined }, box.bottom - top);
        break;
      }
      let low = start + 1, high = boundaries.length - 1, best = start;
      while (low <= high) {
        const middle = Math.floor((low + high) / 2);
        const bottom = rangeBounds(block, boundaries[start], boundaries[middle])?.bottom ?? box.bottom;
        if (bottom - top + 2 <= space) { best = middle; low = middle + 1; }
        else high = middle - 1;
      }
      if (best === start && current().length) {
        nextPage();
        continue;
      }
      if (best === start) best = start + 1;
      if (measure.widowControl !== false && best < boundaries.length - 1) {
        const lineHeight = Number.parseFloat(getComputedStyle(block).lineHeight) || 18;
        let bounds = rangeBounds(block, boundaries[start], boundaries[best]);
        if (bounds && bounds.height < lineHeight * 1.5 && current().length) { nextPage(); continue; }
        if (bounds && box.bottom - bounds.bottom < lineHeight * 1.5 && bounds.height > lineHeight * 2) {
          const lastLine = rangeBounds(block, boundaries[best - 1], boundaries[best])?.top;
          while (best > start + 1 && lastLine !== undefined &&
            (rangeBounds(block, boundaries[best - 1], boundaries[best])?.top ?? 0) >= lastLine - .5) best--;
          bounds = rangeBounds(block, boundaries[start], boundaries[best]);
        }
      }
      const bottom = rangeBounds(block, boundaries[start], boundaries[best])?.bottom ?? box.bottom;
      place({ ...fragment, from: boundaries[start], to: boundaries[best], label: start === 0 ? fragment.label : undefined }, bottom - top);
      start = best;
      top = bottom;
      if (start < boundaries.length - 1) nextPage();
    }
  }
  return pages;
}

function Image({ image }: { image: ImageRun }) {
  const [url, setUrl] = useState('');
  useEffect(() => {
    const next = URL.createObjectURL(image.blob);
    setUrl(next);
    return () => URL.revokeObjectURL(next);
  }, [image.blob]);
  return <img src={url || undefined} alt={image.alt} draggable={false} width={image.width} height={image.height} />;
}

function fontFamily(font: string | undefined) {
  return font ? `"${font.replace(/["\\\n\r]/g, '')}", ${/cambria|georgia|times/i.test(font) ? 'Georgia, serif' : 'Arial, sans-serif'}` : undefined;
}

function formatCss(format: TextFormat): CSSProperties {
  return {
    fontFamily: fontFamily(format.font),
    fontWeight: format.bold === undefined ? undefined : format.bold ? 700 : 400,
    fontStyle: format.italic === undefined ? undefined : format.italic ? 'italic' : 'normal',
    textDecoration: format.underline === undefined && format.strike === undefined ? undefined :
      [format.underline && 'underline', format.strike && 'line-through'].filter(Boolean).join(' ') || 'none',
    backgroundColor: format.highlight,
    verticalAlign: format.vertical === 'superscript' ? 'super' : format.vertical === 'subscript' ? 'sub' : format.vertical,
    fontSize: format.size ? format.size + 'pt' : format.vertical === 'superscript' || format.vertical === 'subscript' ? '.75em' : undefined,
    color: format.color,
  };
}

function PageNumber({ field }: { field: 'PAGE' | 'NUMPAGES' }) {
  const { page, total } = useContext(PageNumbers);
  return field === 'PAGE' ? page : total;
}

function renderInline(inline: Inline, offset: number, comment?: boolean, selectedImage?: string) {
  const range = { 'data-inline-start': offset, 'data-inline-end': offset + inlineLength(inline) };
  if (inline.kind === 'image') {
    return <span key={offset} {...range} data-atom="" data-image={inline.id} data-selected={inline.id === selectedImage} className="rwe-image" contentEditable={false}
      style={{ width: inline.width, aspectRatio: inline.width / inline.height }}>
      <Image image={inline} />
    </span>;
  }
  if (inline.kind === 'pageBreak') return <span key={offset} {...range} data-atom="" contentEditable={false} className="rwe-cell-break">↵</span>;
  if (inline.kind === 'pageField') return <span key={offset} {...range} data-atom="" contentEditable={false} style={formatCss(inline.format)}
    title={inline.field === 'PAGE' ? 'Page number' : 'Total pages'}><PageNumber field={inline.field} /></span>;
  return <span key={offset} {...range} className={[inline.link && 'rwe-link', inline.revision && 'rwe-review-' + inline.revision.kind,
      comment && 'rwe-comment-range'].filter(Boolean).join(' ') || undefined} style={formatCss(inline.format)}
    data-href={inline.link && !inline.link.href.startsWith('#') ? inline.link.href : undefined}>{inline.text}</span>;
}

export function RenderParagraph({ model, fragment }: { model: DocModel; fragment: ParagraphFragment }) {
  const { paragraph, from, to, label } = fragment;
  const selectedImage = useContext(SelectedPicture);
  const named = model.styles.find(style => style.id === paragraph.styleId) ?? model.styles.find(style => style.id === model.defaultStyleId);
  const format = { ...DEFAULT_FORMAT, ...named?.format };
  const style = mergeParagraphStyle(named?.paragraph, paragraph.style);
  let offset = from;
  const marks = paragraph.comments ?? [];
  const points = Array.from(new Set(marks.flatMap(mark => [mark.from, mark.to]).filter(point => point > from && point < to))).sort((a, b) => a - b);
  const inlines = [...points, to].flatMap((point, index) => sliceInlines(paragraph.inlines, index ? points[index - 1] : from, point));
  const indent = paragraph.list ? Math.max(style.indent ?? 0, 34 + paragraph.list.level * 24) : style.indent ?? 0;
  const Tag = named?.heading && named.heading <= 6 ? ('h' + named.heading) as 'h1' : 'p';
  return <Tag className="rwe-paragraph" data-paragraph={paragraph.id} data-from={from} data-to={to}
    data-locked={paragraph.locked ? '' : undefined} contentEditable={paragraph.locked ? false : undefined}
    title={paragraph.locked} dir={style.direction ?? 'auto'}
    role={named?.heading && named.heading > 6 ? 'heading' : undefined}
    aria-level={named?.heading && named.heading > 6 ? named.heading : undefined} style={{
      fontFamily: fontFamily(format.font), fontSize: (format.size ?? 11) + 'pt', color: format.color,
      fontWeight: format.bold ? 700 : 400, fontStyle: format.italic ? 'italic' : 'normal',
      textDecoration: format.underline ? 'underline' : undefined,
      lineHeight: style.lineHeight !== undefined ? Math.max(.8, Math.min(6, style.lineHeight)) : style.lineSpacing
        ? style.lineSpacing.rule === 'exact' ? `${style.lineSpacing.value}px` : `max(${style.lineSpacing.value}px, 1.2em)` : 1.32,
      tabSize: `${model.defaultTabStop ?? 48}px`,
      textAlign: style.alignment, paddingTop: from === 0 ? style.before : 0,
      paddingBottom: to === lengthOf(paragraph.inlines) ? style.after ?? 4 : 0,
      paddingInlineStart: Math.max(0, indent), paddingInlineEnd: Math.max(0, style.rightIndent ?? 0),
      textIndent: from === 0 && !paragraph.list ? style.firstLine : undefined,
    }}>
    {label && <span className="rwe-list-label" style={{ insetInlineStart: indent - 29 }} contentEditable={false}>{label}</span>}
    {inlines.map(inline => {
      const node = renderInline(inline, offset, marks.some(mark => mark.from <= offset && mark.to > offset), selectedImage);
      offset += inlineLength(inline);
      return node;
    })}
    {inlines.length === 0 && !paragraph.locked && <br />}
    {paragraph.locked && <span className="rwe-locked-label" contentEditable={false}>{paragraph.locked}</span>}
  </Tag>;
}

export const RenderFragment = memo(function RenderFragment({ model, fragment, selectTable, selectedTable }: {
  model: DocModel;
  fragment: Fragment;
  selectTable?: (table: Table) => void;
  selectedTable?: string;
}) {
  if (fragment.kind === 'paragraph') return <RenderParagraph model={model} fragment={fragment} />;
  if (fragment.kind === 'opaque') return <div className="rwe-protected" contentEditable={false}>{fragment.label}</div>;
  const { table, from, to, labels, repeated = 0, measuring } = fragment;
  const widths = tableColumnWeights(table);
  const total = widths.reduce((sum, width) => sum + width, 0);
  const margins = { ...DEFAULT_CELL_MARGINS, ...table.margins };
  const borderCss = (value: TableBorder | undefined) => {
    if (!value || value.style === 'none') return '0';
    const width = value.style === 'double'
      ? `max(${value.size}pt, calc(${DOUBLE_BORDER_MIN_PX}px / var(--rwe-page-scale, 1) + ${extraDoubleBorderPoints(value.size)}pt))`
      : `${value.size}pt`;
    return `${width} ${value.style} ${value.color}`;
  };
  const border = (side: TableBorderSide) => borderCss(table.borders ? table.borders[side] : DEFAULT_TABLE_BORDER);
  const renderRow = (index: number, duplicate = false) => {
    const row = table.rows[index];
    const first = duplicate ? index === 0 : index === from && !repeated;
    const bottom = index === table.rows.length - 1 ? 'bottom' : 'insideH';
    return <tr key={(duplicate ? 'repeat-' : '') + row.id} style={{ height: row.height }}
      data-table-row={index} data-repeat-row={duplicate ? '' : undefined}
      contentEditable={duplicate ? false : undefined} aria-hidden={duplicate || undefined}>
      {row.cells.map((cell, column) => {
        const padding = { ...margins, ...cell.margins };
        return <td key={cell.id} colSpan={cell.colSpan} style={{
          backgroundColor: cell.fill,
          padding: `${padding.top}px ${padding.right}px ${padding.bottom}px ${padding.left}px`,
          verticalAlign: cell.vertical === 'center' ? 'middle' : cell.vertical ?? 'top',
          borderLeft: cell.borders?.left ? borderCss(cell.borders.left) : border(column === 0 ? 'left' : 'insideV'),
          borderRight: cell.borders?.right ? borderCss(cell.borders.right) : column === row.cells.length - 1 ? border('right') : undefined,
          borderTop: cell.borders?.top ? borderCss(cell.borders.top) : (measuring ? index === 0 : first) ? border(index === 0 ? 'top' : 'insideH') : '0',
          borderBottom: cell.borders?.bottom ? borderCss(cell.borders.bottom) : border(bottom),
        }}>
          {cell.paragraphs.map(item => <RenderParagraph key={item.id} model={model} fragment={{
            kind: 'paragraph', paragraph: item, from: 0, to: lengthOf(item.inlines), label: labels.get(item.id),
          }} />)}
        </td>;
      })}
    </tr>;
  };
  const grid = <table className="rwe-table" aria-label="Document table" data-table-id={table.id}
    style={{ width: table.width?.unit === 'px' ? table.width.value : (table.width?.value ?? 100) + '%', maxWidth: '100%',
      marginLeft: table.align === 'center' || table.align === 'right' ? 'auto' : undefined,
      marginRight: table.align === 'center' ? 'auto' : undefined }}>
    <colgroup>{widths.map((width, index) => <col key={index} style={{ width: width / total * 100 + '%' }} />)}</colgroup>
    <tbody>
      {Array.from({ length: repeated }, (_, index) => renderRow(index, true))}
      {Array.from({ length: to - from }, (_, index) => renderRow(from + index))}
    </tbody>
  </table>;
  if (!selectTable) return grid;
  const selected = selectedTable === table.id;
  return <div className="rwe-table-frame" data-table-frame={table.id} data-table-selected={selected}>
    <button type="button" className="rwe-table-handle" aria-label="Select entire table" aria-pressed={selected}
      title="Select entire table" tabIndex={-1} contentEditable={false}
      onMouseDown={event => event.preventDefault()} onClick={() => selectTable(table)}><Move size={13} /></button>
    {grid}
  </div>;
}, (previous, next) => sameFragment(previous.fragment, next.fragment) && previous.model.styles === next.model.styles &&
  previous.model.defaultStyleId === next.model.defaultStyleId && previous.selectTable === next.selectTable &&
  previous.model.defaultTabStop === next.model.defaultTabStop &&
  (next.fragment.kind !== 'table' ||
    (previous.selectedTable === next.fragment.table.id) === (next.selectedTable === next.fragment.table.id)));

function sameFragment(a: Fragment, b: Fragment) {
  if (a === b) return true;
  if (a.kind === 'paragraph' && b.kind === 'paragraph') return a.paragraph === b.paragraph && a.from === b.from && a.to === b.to && a.label === b.label;
  if (a.kind === 'opaque' && b.kind === 'opaque') return a.id === b.id && a.label === b.label;
  return false;
}

export function useLayoutWidth(ref: RefObject<HTMLElement | null>) {
  const [width, setWidth] = useState(0);
  useEffect(() => {
    if (!ref.current) return;
    const node = ref.current;
    const observer = new ResizeObserver(() => setWidth(node.clientWidth));
    observer.observe(node);
    return () => observer.disconnect();
  }, [ref]);
  return width;
}

function domPosition(root: HTMLElement, node: Node | null, offset: number): Position | undefined {
  if (!node || !root.contains(node)) return;
  const element = node.nodeType === Node.ELEMENT_NODE ? node as HTMLElement : node.parentElement;
  const block = element?.closest<HTMLElement>('[data-paragraph]');
  if (element?.closest('[data-region-editing="false"], [data-repeat-row]')) return;
  if (!block) {
    const boundary = node.childNodes[offset] ?? node.childNodes[offset - 1];
    const boundaryElement = boundary?.nodeType === Node.ELEMENT_NODE ? boundary as HTMLElement : element;
    const candidates = boundaryElement?.querySelectorAll<HTMLElement>('[data-paragraph]:not([data-repeat-row] *)');
    const after = !node.childNodes[offset];
    const nearest = after ? candidates?.[candidates.length - 1] : candidates?.[0];
    return nearest ? { id: nearest.dataset.paragraph!, offset: Number(after ? nearest.dataset.to : nearest.dataset.from), affinity: after ? 'before' : 'after' } : undefined;
  }
  const run = element?.closest<HTMLElement>('[data-inline-start]');
  const from = Number(block.dataset.from);
  const to = Number(block.dataset.to);
  if (run) {
    const start = Number(run.dataset.inlineStart);
    const end = Number(run.dataset.inlineEnd);
    const position = run.dataset.atom !== undefined ? (offset ? end : start) : node.nodeType === Node.TEXT_NODE ? start + offset : offset ? end : start;
    return { id: block.dataset.paragraph!, offset: Math.max(start, Math.min(end, position)), affinity: position === to ? 'before' : 'after' };
  }
  const next = node.childNodes[offset];
  const previous = node.childNodes[offset - 1];
  const adjacent = (next ?? previous) as HTMLElement | undefined;
  const value = adjacent?.dataset?.[next ? 'inlineStart' : 'inlineEnd'];
  return { id: block.dataset.paragraph!, offset: value === undefined ? (offset ? to : from) : Number(value),
    affinity: offset ? 'before' : 'after' };
}

export function browserSelection(root: HTMLElement, range?: StaticRange): ModelSelection | null {
  const selection = root.ownerDocument.getSelection();
  const from = range?.startContainer ?? selection?.anchorNode ?? null;
  const to = range?.endContainer ?? selection?.focusNode ?? null;
  const start = range?.startOffset ?? selection?.anchorOffset ?? 0;
  const end = range?.endOffset ?? selection?.focusOffset ?? 0;
  if (from && from === to && root.contains(from) && from instanceof Element && start !== end &&
    !from.closest('[data-region-editing="false"]')) {
    const low = Math.min(start, end);
    const high = Math.max(start, end);
    const container = from.matches('table[data-table-id]') && low === 0 && high === from.childNodes.length ? from :
      high === low + 1 ? from.childNodes[low] : undefined;
    const table = container instanceof Element && (container.matches('table[data-table-id]') ? container :
      container.matches('[data-table-frame]') ? container.querySelector('table[data-table-id]') : null);
    if (table) {
      const items = table.querySelectorAll<HTMLElement>('[data-paragraph]:not([data-repeat-row] *)');
      const first = items[0], last = items[items.length - 1];
      if (first && last) {
        const leading = { id: first.dataset.paragraph!, offset: Number(first.dataset.from) };
        const trailing = { id: last.dataset.paragraph!, offset: Number(last.dataset.to) };
        // A page can show just one fragment of a table. The editor expands the selection to every row.
        return { tableId: table.getAttribute('data-table-id')!, anchor: start < end ? leading : trailing,
          focus: start < end ? trailing : leading };
      }
    }
  }
  const anchor = domPosition(root, from, start);
  const focus = domPosition(root, to, end);
  return anchor && focus ? { anchor, focus } : null;
}

export function restoreSelection(root: HTMLElement, selection: ModelSelection) {
  function locate(position: Position) {
    const blocks = Array.from(root.querySelectorAll<HTMLElement>('[data-paragraph="' + position.id + '"]:not([data-repeat-row] *):not([data-region-editing="false"] *)'));
    const matching = blocks.filter(block => Number(block.dataset.from) <= position.offset && position.offset <= Number(block.dataset.to));
    const block = (position.affinity === 'before' ? matching[0] : matching.at(-1)) ?? blocks.at(-1);
    return block && textPoint(block, position.offset, position.affinity !== 'before');
  }
  const anchor = locate(selection.anchor);
  const focus = locate(selection.focus);
  if (!anchor || !focus) return;
  const native = root.ownerDocument.getSelection();
  native?.setBaseAndExtent(anchor[0], anchor[1], focus[0], focus[1]);
}
