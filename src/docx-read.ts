import { DEFAULT_FORMAT, DEFAULT_PAGE, REGION_KINDS, REGION_VARIANTS, TABLE_BORDER_SIDES, mergeParagraphStyle, nodeId, paragraph } from './model';
import type { Block, DocModel, ImageRun, Inline, Link, ListStyle, NamedStyle, PageSetup, Paragraph, ParagraphStyle, Table, TextFormat, PageRegion, RegionKey, PageField, CommentAnchor, DocComment, Revision, TableBorder, TableBorders, CellMargins } from './model';
import {
  A, DOCUMENT_TYPE, TEMPLATE_TYPE, PIC, R, W, WP, TYPES, attr, child, children, containsOnly, invalid, partXml,
  relationshipPath, relationships, resolvePart, safeHref, unzip,
} from './ooxml';
import type { Parts, Relationship } from './ooxml';
import { DocxEditorError } from './types';

export interface DocxSource {
  file: File;
  parts: Parts;
  mainPath: string;
  main: XMLDocument;
  relsPath: string;
  rels?: XMLDocument;
  types: XMLDocument;
  stylesPath?: string;
  numberingPath?: string;
  nodes: Map<string, Element>;
  originalModel: DocModel;
  template?: boolean;
  settingsPath?: string;
  commentsPath?: string;
  regionParts: Partial<Record<RegionKey, { path: string; id: string; xml: XMLDocument }>>;
}
export interface ImportedDocx {
  source: DocxSource;
  model: DocModel;
  readOnlyReason?: string;
  warnings: string[];
}

function numeric(value: string | undefined) {
  if (!value?.trim()) return undefined;
  const number = Number(value);
  return Number.isFinite(number) ? number : undefined;
}
const twips = (value: number | undefined) => value === undefined ? undefined : value / 15;
const valueOf = (parent: Element | undefined, name: string) => attr(child(parent, name));
const onOff = (node: Element | undefined) => node ? !['0', 'false', 'off', 'none'].includes(attr(node) ?? '') : undefined;
const hex = (color?: string) => color && /^[0-9a-f]{6}$/i.test(color) ? '#' + color.toLowerCase() : undefined;

function present<T extends object>(item: T): T {
  return Object.fromEntries(Object.entries(item).filter(([, value]) => value !== undefined)) as T;
}

const WORD_HIGHLIGHTS: Record<string, string> = {
  yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff',
  red: '#ff0000', darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080',
  darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080', lightGray: '#c0c0c0', black: '#000000', white: '#ffffff',
};

function readFormat(node: Element | undefined, fonts: Record<string, string>): TextFormat {
  const faces = child(node, 'rFonts');
  const size = numeric(valueOf(node, 'sz') ?? valueOf(node, 'szCs'));
  const vertical = valueOf(node, 'vertAlign');
  return present({
    strike: onOff(child(node, 'strike')),
    highlight: hex(attr(child(node, 'shd'), 'fill')) ?? WORD_HIGHLIGHTS[valueOf(node, 'highlight') ?? ''],
    vertical: vertical === 'superscript' || vertical === 'subscript' || vertical === 'baseline' ? vertical : undefined,
    bold: onOff(child(node, 'b')) ?? onOff(child(node, 'bCs')),
    italic: onOff(child(node, 'i')) ?? onOff(child(node, 'iCs')),
    underline: onOff(child(node, 'u')),
    font: fonts[attr(faces, 'asciiTheme') ?? ''] ?? attr(faces, 'ascii') ?? attr(faces, 'hAnsi') ?? attr(faces, 'cs'),
    size: size && size >= 2 && size <= 800 ? size / 2 : undefined,
    color: hex(valueOf(node, 'color')),
  });
}

function readParagraphStyle(node: Element | undefined): ParagraphStyle {
  const alignment = valueOf(node, 'jc');
  const spacing = child(node, 'spacing');
  const indent = child(node, 'ind');
  const line = numeric(attr(spacing, 'line'));
  const rule = attr(spacing, 'lineRule');
  const first = numeric(attr(indent, 'firstLine'));
  const hanging = numeric(attr(indent, 'hanging'));
  return present({
    alignment: alignment === 'both' ? 'justify' : alignment === 'start' ? 'left' : alignment === 'end' ? 'right' :
      alignment === 'left' || alignment === 'right' || alignment === 'center' ? alignment : undefined,
    before: twips(numeric(attr(spacing, 'before'))), after: twips(numeric(attr(spacing, 'after'))),
    lineHeight: line && (!attr(spacing, 'lineRule') || attr(spacing, 'lineRule') === 'auto') ? line / 240 : undefined,
    lineSpacing: line && (rule === 'exact' || rule === 'atLeast') ? { rule, value: line / 15 } : undefined,
    direction: child(node, 'bidi') ? onOff(child(node, 'bidi')) ? 'rtl' : 'ltr' : undefined,
    keepNext: onOff(child(node, 'keepNext')), keepLines: onOff(child(node, 'keepLines')), widowControl: onOff(child(node, 'widowControl')),
    indent: twips(numeric(attr(indent, 'start') ?? attr(indent, 'left'))),
    rightIndent: twips(numeric(attr(indent, 'end') ?? attr(indent, 'right'))),
    firstLine: twips(first ?? (hanging === undefined ? undefined : -hanging)),
    pageBreakBefore: onOff(child(node, 'pageBreakBefore')),
  });
}

function readPage(node: Element | undefined): PageSetup {
  const size = child(node, 'pgSz');
  const margin = child(node, 'pgMar');
  const bounded = (element: Element | undefined, name: string, fallback: number, min: number, max: number) => {
    const value = twips(numeric(attr(element, name))) ?? fallback;
    return Math.min(max, Math.max(min, value));
  };
  const width = bounded(size, 'w', DEFAULT_PAGE.width, 190, 3000);
  const height = bounded(size, 'h', DEFAULT_PAGE.height, 190, 3000);
  return {
    width, height,
    top: bounded(margin, 'top', DEFAULT_PAGE.top, 0, height * .4),
    bottom: bounded(margin, 'bottom', DEFAULT_PAGE.bottom, 0, height * .4),
    left: bounded(margin, 'left', DEFAULT_PAGE.left, 0, width * .4),
    right: bounded(margin, 'right', DEFAULT_PAGE.right, 0, width * .4),
    header: bounded(margin, 'header', DEFAULT_PAGE.header, 0, height * .4),
    footer: bounded(margin, 'footer', DEFAULT_PAGE.footer, 0, height * .4),
  };
}

function themeFonts(theme?: XMLDocument) {
  const scheme = theme?.getElementsByTagNameNS(A, 'fontScheme')[0];
  const major = child(child(scheme, 'majorFont', A), 'latin', A)?.getAttribute('typeface') ?? 'Cambria';
  const minor = child(child(scheme, 'minorFont', A), 'latin', A)?.getAttribute('typeface') ?? 'Calibri';
  return { majorAscii: major, majorHAnsi: major, minorAscii: minor, minorHAnsi: minor };
}

function readStyles(xml: XMLDocument | undefined, fonts: Record<string, string>) {
  const root = xml?.documentElement;
  const defaults = child(root, 'docDefaults');
  const format = { ...DEFAULT_FORMAT, ...readFormat(child(child(defaults, 'rPrDefault'), 'rPr'), fonts) };
  const paragraph = readParagraphStyle(child(child(defaults, 'pPrDefault'), 'pPr'));
  const elements = new Map(children(root, 'style').flatMap(node => {
    const id = attr(node, 'styleId');
    return id ? [[id, node]] as const : [];
  }));
  const styles = new Map<string, NamedStyle>();
  const defaultStyleId = Array.from(elements).find(([, node]) => attr(node, 'type') === 'paragraph' &&
    ['1', 'true', 'on'].includes(attr(node, 'default') ?? ''))?.[0] ?? 'Normal';

  function resolve(id: string, chain = new Set<string>()): NamedStyle | undefined {
    if (styles.has(id)) return styles.get(id);
    const node = elements.get(id);
    const type = attr(node, 'type');
    if (!node || chain.has(id) || chain.size >= 64 || (type !== 'paragraph' && type !== 'character')) return;
    const baseId = valueOf(node, 'basedOn');
    const base = baseId ? resolve(baseId, new Set([...chain, id])) : undefined;
    const name = valueOf(node, 'name') ?? id;
    const props = child(node, 'pPr');
    const outline = numeric(valueOf(props, 'outlineLvl'));
    const heading = outline === undefined ? base?.heading : outline < 9 ? outline + 1 : undefined;
    const local = { paragraph: readParagraphStyle(props), format: readFormat(child(node, 'rPr'), fonts), outline };
    const style: NamedStyle = {
      id, name, type,
      paragraph: mergeParagraphStyle(base?.paragraph ?? (type === 'paragraph' ? paragraph : {}), local.paragraph),
      format: { ...(base?.format ?? (type === 'paragraph' ? format : {})), ...local.format },
      heading, basedOn: base?.id, local,
    };
    styles.set(id, style);
    return style;
  }
  for (const id of elements.keys()) resolve(id);
  if (!styles.has(defaultStyleId)) styles.set(defaultStyleId, { id: defaultStyleId, name: 'Normal', type: 'paragraph', paragraph, format });
  return { styles, defaultStyleId, elements };
}

function numbering(xml?: XMLDocument) {
  const root = xml?.documentElement;
  const abstract = new Map(children(root, 'abstractNum').map(node => [attr(node, 'abstractNumId'), node]));
  const nums = new Map(children(root, 'num').map(node => [attr(node, 'numId'), node]));
  return (id: string, level: number): ListStyle | null | undefined => {
    const num = nums.get(id);
    const override = children(num, 'lvlOverride').find(node => Number(attr(node, 'ilvl')) === level);
    const definition = child(override, 'lvl') ?? children(abstract.get(valueOf(num, 'abstractNumId')), 'lvl')
      .find(node => Number(attr(node, 'ilvl')) === level);
    const format = valueOf(definition, 'numFmt');
    if (format === 'none') return null;
    if (!['bullet', 'decimal', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman'].includes(format ?? '')) return;
    return {
      id, level, kind: format === 'bullet' ? 'bullet' : 'number',
      numberFormat: format === 'bullet' ? undefined : format as ListStyle['numberFormat'],
      start: numeric(valueOf(override, 'startOverride') ?? valueOf(definition, 'start')) ?? 1,
      pattern: format === 'decimal' ? valueOf(definition, 'lvlText') ?? `%${level + 1}.` : undefined,
    };
  };
}

function inheritedNumbering(styleId: string, elements: Map<string, Element>) {
  const seen = new Set<string>();
  let id: string | undefined = styleId;
  while (id && !seen.has(id)) {
    seen.add(id);
    const node = elements.get(id);
    const num = child(child(node, 'pPr'), 'numPr');
    if (num) return num;
    id = valueOf(node, 'basedOn');
  }
}

function revisionPresent(body: Element, ignoreTextRevisions = false) {
  const names = ['ins', 'del', 'moveFrom', 'moveTo', 'pPrChange', 'rPrChange', 'tblPrChange', 'trPrChange',
    'tcPrChange', 'sectPrChange', 'tblGridChange', 'cellIns', 'cellDel', 'cellMerge', 'numberingChange',
    'moveFromRangeStart', 'moveToRangeStart', 'customXmlInsRangeStart', 'customXmlDelRangeStart'];
  return names.some(name => !(ignoreTextRevisions && (name === 'ins' || name === 'del')) && body.getElementsByTagNameNS(W, name).length > 0);
}

export async function importDocx(file: File, templates = false): Promise<ImportedDocx> {
  const parts = await unzip(file, templates);
  const types = await partXml(parts, '[Content_Types].xml');
  const root = relationships(await partXml(parts, '_rels/.rels'));
  const mainRel = Array.from(root.values()).find(rel => rel.type.endsWith('/officeDocument') && !rel.external);
  if (!mainRel) invalid('This archive has no Word document.');
  const mainPath = resolvePart('', mainRel.target);
  const main = await partXml(parts, mainPath);
  const contentType = Array.from(types.documentElement.children).find(node => node.getAttribute('PartName') === '/' + mainPath)
    ?.getAttribute('ContentType');
  const template = contentType === TEMPLATE_TYPE;
  if (types.documentElement.namespaceURI !== TYPES || (contentType !== DOCUMENT_TYPE && !(templates && template))) {
    invalid(templates ? 'This archive is not an unencrypted Word document or template.' : 'This archive is not a .docx document.');
  }
  if (main.documentElement.localName !== 'document') invalid('The package does not contain a Word document.');
  if (main.documentElement.namespaceURI !== W) {
    throw new DocxEditorError('UNSUPPORTED_FILE', 'This version of WordprocessingML is unsupported. Save a standard Word Document (.docx) in Word.');
  }
  const body = child(main.documentElement, 'body');
  if (!body) invalid('The Word document has no body.');
  const relsPath = relationshipPath(mainPath);
  const rels = parts.has(relsPath) ? await partXml(parts, relsPath) : undefined;
  const relMap = rels ? relationships(rels) : new Map<string, Relationship>();
  function related(type: string) {
    const rel = Array.from(relMap.values()).find(rel => rel.type === R + '/' + type && !rel.external);
    return rel ? resolvePart(mainPath, rel.target) : undefined;
  }
  const stylesPath = related('styles');
  const numberingPath = related('numbering');
  const themePath = related('theme');
  const settingsPath = related('settings');
  const commentsPath = related('comments');
  const [styleXml, numberXml, themeXml, settings, commentXml] = await Promise.all(
    [stylesPath, numberingPath, themePath, settingsPath, commentsPath].map(path => path ? partXml(parts, path) : undefined),
  );
  const fonts = themeFonts(themeXml);
  const { styles, defaultStyleId, elements } = readStyles(styleXml, fonts);
  const findList = numbering(numberXml);
  const nodes = new Map<string, Element>();
  function remember(prefix: string, node: Element) {
    const id = nodeId(prefix);
    nodes.set(id, node);
    return id;
  }
  const warnings = new Set<string>();
  const sections = body.getElementsByTagNameNS(W, 'sectPr');
  const section = child(body, 'sectPr');
  let readOnlyReason: string | undefined;
  if (sections.length > (section ? 1 : 0) || Number(attr(child(section, 'cols'), 'num') ?? 1) > 1) {
    readOnlyReason = 'This document uses multiple sections or columns. Download the original to edit it in Word.';
  } else if (revisionPresent(body, true)) {
    readOnlyReason = 'This document contains unsupported tracked changes.';
  } else if (['permStart', 'permEnd'].some(name =>
    body.getElementsByTagNameNS(W, name).length) || child(settings?.documentElement, 'documentProtection') ||
    Array.from(root.values()).some(rel => rel.type.endsWith('/digital-signature/origin'))) {
    readOnlyReason = 'This document contains comments, editing restrictions, or digital signatures.';
  }
  const revisionIds = new Set<number>();
  let unsupportedReview = false;
  const reviewPart = children(commentXml?.documentElement, 'comment');
  const comments: DocComment[] = [];
  for (const entry of reviewPart) {
    const id = numeric(attr(entry, 'id'));
    if (id === undefined || comments.some(comment => comment.id === id)) { unsupportedReview = true; continue; }
    comments.push({ id, author: attr(entry, 'author') ?? 'Author', date: attr(entry, 'date') ?? '',
      text: children(entry, 'p').map(item => Array.from(item.getElementsByTagNameNS(W, 't')).map(node => node.textContent).join('')).join('\n') });
  }
  const modernComments = Array.from(relMap.values()).some(rel => /commentsExtended|commentsIds|commentsExtensible/.test(rel.type));
  const mainContext = { path: mainPath, relations: relMap };
  type PartContext = typeof mainContext;
  function readImage(run: Element, drawing: Element, context: PartContext): ImageRun | undefined {
    const inline = child(drawing, 'inline', WP);
    const graphic = child(inline, 'graphic', A);
    const data = child(graphic, 'graphicData', A);
    const picture = child(data, 'pic', PIC);
    const blip = child(child(picture, 'blipFill', PIC), 'blip', A);
    if (!blip || attr(blip, 'link', R)) return;
    const relation = context.relations.get(attr(blip, 'embed', R) ?? '');
    if (!relation || relation.external || relation.type !== R + '/image') return;
    const path = resolvePart(context.path, relation.target);
    const mime = /\.png$/i.test(path) ? 'image/png' : /\.jpe?g$/i.test(path) ? 'image/jpeg' : undefined;
    const blob = parts.get(path);
    const extent = child(inline, 'extent', WP);
    const width = Number(extent?.getAttribute('cx')) / 9525;
    const height = Number(extent?.getAttribute('cy')) / 9525;
    if (!blob || !mime || width <= 0 || height <= 0 || !Number.isFinite(width + height)) return;
    return {
      kind: 'image', id: remember('image', run), blob: blob.slice(0, blob.size, mime), width, height,
      alt: child(inline, 'docPr', WP)?.getAttribute('descr') ?? 'Picture',
    };
  }

  function readRun(run: Element, context: PartContext, link?: Link, revision?: Revision): { items: Inline[]; unsupported: boolean } {
    const props = child(run, 'rPr');
    const named = styles.get(valueOf(props, 'rStyle') ?? '');
    const format = { ...named?.format, ...readFormat(props, fonts) };
    const source = remember('run', run);
    const items: Inline[] = [];
    let text = '';
    let unsupported = false;
    function flush() {
      if (text) items.push({ kind: 'text', text, format, source, link, revision });
      text = '';
    }
    for (const node of Array.from(run.children)) {
      if (node.namespaceURI !== W) { unsupported = true; continue; }
      switch (node.localName) {
        case 'rPr':
        case 'lastRenderedPageBreak': break;
        case 'delText':
        case 't': text += node.textContent ?? ''; break;
        case 'tab': text += '\t'; break;
        case 'cr': text += '\n'; break;
        case 'noBreakHyphen': text += '\u2011'; break;
        case 'softHyphen': text += '\u00ad'; break;
        case 'br':
          if (attr(node, 'type') === 'page') {
            flush();
            items.push({ kind: 'pageBreak' });
          } else if (!attr(node, 'type') || attr(node, 'type') === 'textWrapping') text += '\n';
          else unsupported = true;
          break;
        case 'drawing': {
          flush();
          const image = readImage(run, node, context);
          if (image && !link && children(run, 'drawing').length === 1 &&
            containsOnly(run, ['rPr', 'drawing', 'lastRenderedPageBreak'])) items.push(image);
          else unsupported = true;
          break;
        }
        default: unsupported = true;
      }
    }
    flush();
    return { items, unsupported };
  }

  function readField(code: string, field: Element, simple: boolean): PageField | undefined {
    const name = code.trim().replace(/\s+\\\*\s+(?:MERGEFORMAT|CHARFORMAT)\s*$/i, '').toUpperCase();
    if (name !== 'PAGE' && name !== 'NUMPAGES') return;
    const run = simple ? child(field, 'r') : field;
    return { kind: 'pageField', field: name, format: readFormat(child(run, 'rPr'), fonts),
      source: simple ? remember('field', field) : undefined };
  }

  function readParagraph(node: Element, context: PartContext): Paragraph {
    const props = child(node, 'pPr');
    const styleId = valueOf(props, 'pStyle') ?? defaultStyleId;
    const number = child(props, 'numPr') ?? inheritedNumbering(styleId, elements);
    const numId = valueOf(number, 'numId');
    const numbering = numId && numId !== '0' ? findList(numId, numeric(valueOf(number, 'ilvl')) ?? 0) : undefined;
    const list = numbering ?? undefined;
    let locked = numId && numId !== '0' && numbering !== null && !list ? 'Unsupported list numbering' : undefined;
    if (child(props, 'sectPr') || child(props, 'framePr')) locked = 'Protected paragraph layout';
    const inlines: Inline[] = [];
    const anchors = new Map<number, Partial<CommentAnchor>>();
    const currentOffset = () => inlines.reduce((total, run) => total + (run.kind === 'text' ? run.text.length : 1), 0);
    function appendRuns(container: Element, link?: Link) {
      for (const run of children(container, 'r')) {
        const result = readRun(run, context, link);
        inlines.push(...result.items);
        if (result.unsupported) locked ??= 'Contains unsupported content';
      }
    }
    const childrenOfParagraph = Array.from(node.children);
    for (let index = 0; index < childrenOfParagraph.length; index++) {
      const inline = childrenOfParagraph[index];
      if (inline.namespaceURI !== W) { locked ??= 'Contains unsupported content'; continue; }
      switch (inline.localName) {
        case 'pPr':
        case 'proofErr': break;
        case 'r': {
          if (attr(child(inline, 'fldChar'), 'fldCharType') === 'begin') {
            let next = index + 1;
            let code = '';
            let display = inline;
            for (; next < childrenOfParagraph.length; next++) {
              const candidate = childrenOfParagraph[next];
              if (candidate.namespaceURI !== W || candidate.localName !== 'r') break;
              const kind = attr(child(candidate, 'fldChar'), 'fldCharType');
              if (kind === 'begin') break;
              if (kind === 'end') {
                const field = readField(code, display, false);
                if (field) { inlines.push(field); index = next; }
                break;
              }
              code += child(candidate, 'instrText')?.textContent ?? '';
              if (child(candidate, 't')) display = candidate;
            }
            if (index === next) break;
          }
          if (context === mainContext && child(inline, 'commentReference') && containsOnly(inline, ['rPr', 'commentReference'])) break;
          const result = readRun(inline, context);
          inlines.push(...result.items);
          if (result.unsupported) locked ??= 'Contains unsupported content';
          break;
        }
        case 'hyperlink': {
          const rel = context.relations.get(attr(inline, 'id', R) ?? '');
          const anchor = attr(inline, 'anchor');
          const href = rel?.external && rel.type === R + '/hyperlink' ? safeHref(rel.target) : undefined;
          const link = href ? { href, relationship: rel!.id, anchor } : anchor && !rel ? { href: '#' + anchor, anchor } : undefined;
          if (!link || !containsOnly(inline, ['r'])) locked ??= 'Contains an unsupported link';
          appendRuns(inline, link);
          break;
        }
        case 'fldSimple': {
          const field = readField(attr(inline, 'instr') ?? '', inline, true);
          if (field) inlines.push(field);
          else locked ??= 'Contains an unsupported field';
          break;
        }
        case 'commentRangeStart':
        case 'commentRangeEnd': {
          const id = numeric(attr(inline, 'id'));
          if (context !== mainContext || id === undefined) { locked ??= 'Contains protected review markers'; break; }
          const boundary = inline.localName === 'commentRangeStart' ? 'from' : 'to';
          const mark = anchors.get(id) ?? { id };
          if (mark[boundary] !== undefined) unsupportedReview = true;
          mark[boundary] = currentOffset();
          anchors.set(id, mark);
          break;
        }
        case 'ins':
        case 'del': {
          const id = numeric(attr(inline, 'id'));
          const author = attr(inline, 'author');
          if (context !== mainContext || id === undefined || revisionIds.has(id) || !author || !containsOnly(inline, ['r'])) {
            unsupportedReview = true;
            locked ??= 'Contains unsupported tracked changes';
            break;
          }
          revisionIds.add(id);
          const revision: Revision = { id, kind: inline.localName, author, date: attr(inline, 'date') ?? '' };
          for (const run of children(inline, 'r')) {
            if (!containsOnly(run, ['rPr', 't', 'delText', 'tab', 'br', 'cr', 'noBreakHyphen', 'softHyphen', 'lastRenderedPageBreak'])) unsupportedReview = true;
            const parsed = readRun(run, context, undefined, revision);
            if (parsed.unsupported || parsed.items.some(item => item.kind !== 'text')) unsupportedReview = true;
            inlines.push(...parsed.items);
          }
          break;
        }
        case 'moveTo': appendRuns(inline); break;
        case 'bookmarkStart':
        case 'bookmarkEnd': locked ??= 'Contains a bookmark'; break;
        default: locked ??= 'Contains unsupported content';
      }
    }
    const marks: CommentAnchor[] = [];
    for (const [id, mark] of anchors) {
      if (mark.from === undefined || mark.to === undefined || mark.from > mark.to ||
        !comments.some(comment => comment.id === id) || locked || inlines.some(run => run.kind === 'text' && run.revision)) unsupportedReview = true;
      else marks.push({ id, from: mark.from, to: mark.to });
    }
    return { kind: 'paragraph', id: remember('p', node), inlines, styleId, style: readParagraphStyle(props), list, locked,
      comments: marks.length ? marks : undefined };
  }

  function readTableBorder(node: Element | undefined): TableBorder | undefined {
    if (!node) return;
    const kind = attr(node);
    const style = kind === 'nil' || kind === 'none' ? 'none' : kind === 'dotted' ? 'dotted' :
      kind === 'dashed' || kind === 'dashSmallGap' ? 'dashed' : kind === 'double' ? 'double' : 'solid';
    return { style, color: hex(attr(node, 'color')) ?? '#000000', size: (numeric(attr(node, 'sz')) ?? 4) / 8 };
  }

  function readTableBorders(node: Element | undefined): TableBorders | undefined {
    return node && Object.fromEntries(TABLE_BORDER_SIDES.flatMap(side => {
      const border = readTableBorder(child(node, side) ?? (side === 'left' ? child(node, 'start') : side === 'right' ? child(node, 'end') : undefined));
      return border ? [[side, border]] : [];
    }));
  }

  function readCellMargins(node: Element | undefined): CellMargins | undefined {
    return node && Object.fromEntries(['top', 'right', 'bottom', 'left'].flatMap(side => {
      const margin = child(node, side) ?? (side === 'left' ? child(node, 'start') : side === 'right' ? child(node, 'end') : undefined);
      const value = twips(numeric(attr(margin, 'w')));
      return value === undefined ? [] : [[side, Math.max(0, value)]];
    }));
  }

  function readTable(node: Element, context: PartContext): Table | undefined {
    if (!containsOnly(node, ['tblPr', 'tblGrid', 'tr']) || child(child(node, 'tblPr'), 'tblpPr')) return;
    const rows = children(node, 'tr');
    const spanOf = (cell: Element) => numeric(valueOf(child(cell, 'tcPr'), 'gridSpan')) ?? 1;
    const columns = children(rows[0], 'tc').reduce((sum, cell) => sum + spanOf(cell), 0);
    if (!rows.length || !columns || columns > 20 || rows.some(row => !containsOnly(row, ['trPr', 'tblPrEx', 'tc']) ||
      children(row, 'tc').reduce((sum, cell) => sum + spanOf(cell), 0) !== columns || child(child(row, 'trPr'), 'gridBefore') || child(child(row, 'trPr'), 'gridAfter'))) return;
    const cells = rows.flatMap(row => children(row, 'tc'));
    if (cells.some(cell => !containsOnly(cell, ['tcPr', 'p']) || !children(cell, 'p').length ||
      !Number.isInteger(spanOf(cell)) || spanOf(cell) < 1 || ['vMerge', 'hMerge'].some(name => child(child(cell, 'tcPr'), name)))) return;
    const grid = children(child(node, 'tblGrid'), 'gridCol');
    const properties = child(node, 'tblPr');
    const width = child(properties, 'tblW');
    const widthValue = numeric(attr(width, 'w'));
    const widthType = attr(width, 'type');
    const justification = attr(child(properties, 'jc'));
    const borders = child(properties, 'tblBorders');
    const margins = child(properties, 'tblCellMar');
    return {
      kind: 'table', id: remember('table', node),
      structuralLock: cells.some(cell => spanOf(cell) > 1) || undefined,
      columnWidths: cells.some(cell => spanOf(cell) > 1) ? Array.from({ length: columns }, (_, index) => twips(numeric(attr(grid[index], 'w'))) ?? 1) : undefined,
      width: widthValue && widthValue > 0 && (widthType === 'pct' || widthType === 'dxa') ? {
        value: widthValue / (widthType === 'pct' ? 50 : 15), unit: widthType === 'pct' ? 'percent' : 'px',
      } : undefined,
      align: justification === 'center' ? 'center' : justification === 'right' || justification === 'end' ? 'right' :
        justification === 'left' || justification === 'start' ? 'left' : undefined,
      borders: readTableBorders(borders),
      margins: readCellMargins(margins),
      rows: rows.map(row => {
        const rowProperties = child(row, 'trPr');
        const height = child(rowProperties, 'trHeight');
        return {
          id: remember('row', row), header: onOff(child(rowProperties, 'tblHeader')),
          height: attr(height, 'hRule') === 'auto' ? undefined : twips(numeric(attr(height, 'val'))),
          exactHeight: attr(height, 'hRule') === 'exact',
          cells: children(row, 'tc').map((cell, index) => {
            const props = child(cell, 'tcPr');
            const width = child(props, 'tcW');
            const vertical = attr(child(props, 'vAlign'));
            return {
              vertical: vertical === 'top' || vertical === 'center' || vertical === 'bottom' ? vertical : undefined,
              borders: readTableBorders(child(props, 'tcBorders')),
              margins: readCellMargins(child(props, 'tcMar')),
              id: remember('cell', cell),
              colSpan: spanOf(cell) > 1 ? spanOf(cell) : undefined,
              paragraphs: children(cell, 'p').map(item => readParagraph(item, context)),
              width: Math.max(1, twips(numeric(attr(grid[index], 'w') ??
                (attr(width, 'type') === 'dxa' ? attr(width, 'w') : undefined))) ?? 1),
              fill: hex(attr(child(props, 'shd'), 'fill')),
            };
          }),
        };
      }),
    };
  }

  function readBlocks(container: Element, context: PartContext, skip?: Element): Block[] {
    return Array.from(container.children).filter(node => node !== skip).map(node => {
      const block = node.namespaceURI === W && node.localName === 'p' ? readParagraph(node, context) :
        node.namespaceURI === W && node.localName === 'tbl' ? readTable(node, context) : undefined;
      return block ?? { kind: 'opaque', id: remember('opaque', node),
        label: node.localName === 'tbl' ? 'Protected table' : 'Protected document content' };
    });
  }

  const regionParts: DocxSource['regionParts'] = {};
  const items: Partial<Record<RegionKey, PageRegion>> = {};
  const usedRegionPaths = new Set<string>();
  for (const kind of REGION_KINDS) {
    for (const reference of children(section, kind + 'Reference')) {
      const variant = REGION_VARIANTS.find(value => value === attr(reference, 'type'));
      const id = attr(reference, 'id', R);
      const relationship = relMap.get(id ?? '');
      if (!variant || !relationship || relationship.external || relationship.type !== R + '/' + kind) {
        warnings.add('Some header or footer references could not be displayed and are preserved.');
        continue;
      }
      const key: RegionKey = `${kind}:${variant}`;
      const path = resolvePart(mainPath, relationship.target);
      const xml = await partXml(parts, path);
      if (xml.documentElement.namespaceURI !== W || xml.documentElement.localName !== (kind === 'header' ? 'hdr' : 'ftr')) {
        warnings.add('An unsupported header or footer was preserved.');
        continue;
      }
      const relationshipFile = relationshipPath(path);
      const relations = parts.has(relationshipFile) ? relationships(await partXml(parts, relationshipFile)) : new Map();
      const locked = usedRegionPaths.has(path) || revisionPresent(xml.documentElement) ||
        ['commentRangeStart', 'commentRangeEnd', 'commentReference', 'permStart', 'permEnd'].some(name => xml.getElementsByTagNameNS(W, name).length > 0);
      if (usedRegionPaths.has(path)) {
        for (const [existingKey, existing] of Object.entries(regionParts)) {
          if (existing.path === path) items[existingKey as RegionKey]!.locked = true;
        }
      }
      if (locked) warnings.add('Shared or reviewed headers and footers are protected.');
      usedRegionPaths.add(path);
      regionParts[key] = { path, id: relationship.id, xml };
      const content = readBlocks(xml.documentElement, { path, relations });
      items[key] = { blocks: content.length ? content : [paragraph(defaultStyleId)], locked };
    }
  }
  const regions = { first: !!onOff(child(section, 'titlePg')), even: !!onOff(child(settings?.documentElement, 'evenAndOddHeaders')), items };
  const blocks = readBlocks(body, mainContext, section);
  if (!blocks.length) blocks.push(paragraph(defaultStyleId));
  const anyEditable = blocks.some(block => block.kind === 'paragraph' ? !block.locked :
    block.kind === 'table' && block.rows.some(row => row.cells.some(cell => cell.paragraphs.some(item => !item.locked))));
  if (!anyEditable) {
    readOnlyReason ??= 'The document has no editable paragraphs.';
  }
  if (blocks.some(block => block.kind === 'opaque' || (block.kind === 'paragraph' && block.locked) ||
    (block.kind === 'table' && block.rows.some(row => row.cells.some(cell => cell.paragraphs.some(item => item.locked)))))) {
    warnings.add('Protected content stays in the DOCX and cannot be edited here.');
  }
  const marked = blocks.flatMap(block => block.kind === 'paragraph' ? [block] :
    block.kind === 'table' ? block.rows.flatMap(row => row.cells.flatMap(cell => cell.paragraphs)) : []).flatMap(item => item.comments ?? []);
  if (unsupportedReview || modernComments || comments.some(comment => marked.filter(mark => mark.id === comment.id).length !== 1) ||
    body.getElementsByTagNameNS(W, 'ins').length + body.getElementsByTagNameNS(W, 'del').length !== revisionIds.size) {
    readOnlyReason ??= 'This document contains comments or tracked changes that need Word to edit.';
  }
  const tracking = !!onOff(child(settings?.documentElement, 'trackRevisions'));
  const model: DocModel = { blocks, page: readPage(section), styles: Array.from(styles.values()), defaultStyleId, regions, comments, tracking,
    defaultTabStop: twips(numeric(valueOf(settings?.documentElement, 'defaultTabStop'))) ?? 48 };
  return { model, readOnlyReason, warnings: Array.from(warnings),
    source: { file, parts, mainPath, main, relsPath, rels, types, stylesPath, numberingPath, settingsPath, commentsPath, regionParts, nodes, originalModel: model, template } };
}
