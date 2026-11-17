import { REGION_KINDS, REGION_VARIANTS, TABLE_BORDER_SIDES, DEFAULT_TABLE_BORDER, tableColumnWeights, lengthOf, paragraphs, sliceInlines } from './model';
import type { Block, DocModel, ImageRun, Link, ListStyle, NamedStyle, Paragraph, ParagraphStyle, Table, TextFormat, TextRun, RegionKey, PageField, Revision } from './model';
import type { DocxSource } from './docx-read';
import { A, DOCUMENT_TYPE, MIME, PIC, R, RELS, TYPES, W, WP, add, attr, child, childProperty, children, element, parseXml, partXml, putXml, relationshipPath, zip } from './ooxml';
import type { Parts } from './ooxml';

const P_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
  'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct',
  'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
  'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
  'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];
const R_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike',
  'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color',
  'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText',
  'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange'];

const TABLE_ORDER = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize',
  'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout',
  'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription', 'tblPrChange'];
const ROW_ORDER = ['cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore', 'wAfter', 'cantSplit',
  'trHeight', 'tblHeader', 'tblCellSpacing', 'jc', 'hidden', 'ins', 'del', 'trPrChange'];
const CELL_ORDER = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap',
  'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark', 'headers', 'cellIns', 'cellDel', 'cellMerge', 'tcPrChange'];

const twips = (px: number) => Math.round(px * 15);
const emu = (px: number) => Math.round(px * 9525);
const halfPoints = (points: number) => Math.round(points * 2);

function blankXml(root: string, namespace: string, prefix = '') {
  const name = prefix ? prefix + ':' + root : root;
  return parseXml(`<${name} xmlns${prefix ? ':' + prefix : ''}="${namespace}"/>`);
}

function sameLink(a?: Link, b?: Link) {
  return a?.href === b?.href && a?.relationship === b?.relationship && a?.anchor === b?.anchor;
}

function writeFormat(node: Element, format: TextFormat, previous: TextFormat = {}) {
  function set(key: keyof TextFormat, name: string, attributes?: Record<string, string | number>) {
    if (format[key] !== previous[key]) childProperty(node, name, R_ORDER, attributes);
  }
  set('font', 'rFonts', format.font ? { ascii: format.font, hAnsi: format.font, eastAsia: format.font, cs: format.font } : undefined);
  set('bold', 'b', format.bold === undefined ? undefined : { val: format.bold ? 1 : 0 });
  set('bold', 'bCs', format.bold === undefined ? undefined : { val: format.bold ? 1 : 0 });
  set('italic', 'i', format.italic === undefined ? undefined : { val: format.italic ? 1 : 0 });
  set('italic', 'iCs', format.italic === undefined ? undefined : { val: format.italic ? 1 : 0 });
  set('underline', 'u', format.underline === undefined ? undefined : { val: format.underline ? 'single' : 'none' });
  set('strike', 'strike', format.strike === undefined ? undefined : { val: format.strike ? 1 : 0 });
  set('vertical', 'vertAlign', format.vertical ? { val: format.vertical } : undefined);
  if (format.highlight !== previous.highlight) {
    child(node, 'highlight')?.remove();
    childProperty(node, 'shd', R_ORDER, format.highlight ? { val: 'clear', color: 'auto', fill: format.highlight.slice(1).toUpperCase() } : undefined);
  }
  set('color', 'color', format.color ? { val: format.color.slice(1).toUpperCase() } : undefined);
  set('size', 'sz', format.size ? { val: halfPoints(format.size) } : undefined);
  set('size', 'szCs', format.size ? { val: halfPoints(format.size) } : undefined);
}

function writeParagraphFormat(node: Element, style: ParagraphStyle, previous: ParagraphStyle = {}) {
  function updateAttributes(name: string, changes: Record<string, string | number | undefined>) {
    const existing = child(node, name);
    const property = existing ? existing.cloneNode(true) as Element : element(node, name);
    for (const [key, value] of Object.entries(changes)) {
      if (value === undefined) property.removeAttributeNS(W, key);
      else property.setAttributeNS(W, 'w:' + key, String(value));
    }
    if (existing) existing.replaceWith(property);
    else {
      childProperty(node, name, P_ORDER, {});
      child(node, name)!.replaceWith(property);
    }
  }
  if (style.alignment !== previous.alignment) {
    childProperty(node, 'jc', P_ORDER, style.alignment ? { val: style.alignment === 'justify' ? 'both' : style.alignment } : undefined);
  }
  if (style.pageBreakBefore !== previous.pageBreakBefore) {
    childProperty(node, 'pageBreakBefore', P_ORDER, style.pageBreakBefore === undefined ? undefined : { val: style.pageBreakBefore ? 1 : 0 });
  }
  for (const key of ['keepNext', 'keepLines', 'widowControl'] as const) {
    if (style[key] !== previous[key]) childProperty(node, key, P_ORDER, style[key] === undefined ? undefined : { val: style[key] ? 1 : 0 });
  }
  if (style.direction !== previous.direction) childProperty(node, 'bidi', P_ORDER,
    style.direction === undefined ? undefined : { val: style.direction === 'rtl' ? 1 : 0 });
  if (style.before !== previous.before || style.after !== previous.after || style.lineHeight !== previous.lineHeight) {
    const values: Record<string, number | string | undefined> = {};
    for (const key of ['before', 'after'] as const) {
      if (style[key] !== previous[key]) Object.assign(values, {
        [key]: style[key] === undefined ? undefined : twips(style[key]), [key + 'Lines']: undefined, [key + 'Autospacing']: undefined,
      });
    }
    if (style.lineHeight !== previous.lineHeight) Object.assign(values, {
      line: style.lineHeight === undefined ? undefined : Math.round(style.lineHeight * 240),
      lineRule: style.lineHeight === undefined ? undefined : 'auto',
    });
    updateAttributes('spacing', values);
  }
  if (style.lineHeight === undefined && style.lineSpacing !== previous.lineSpacing) {
    updateAttributes('spacing', { line: style.lineSpacing ? twips(style.lineSpacing.value) : undefined, lineRule: style.lineSpacing?.rule });
  }
  if (style.indent !== previous.indent || style.rightIndent !== previous.rightIndent || style.firstLine !== previous.firstLine) {
    const values: Record<string, number | undefined> = {};
    for (const [key, side, logical] of [['indent', 'left', 'start'], ['rightIndent', 'right', 'end']] as const) {
      if (style[key] !== previous[key]) Object.assign(values, {
        [side]: style[key] === undefined ? undefined : twips(style[key]),
        [logical]: undefined, [side + 'Chars']: undefined, [logical + 'Chars']: undefined,
      });
    }
    if (style.firstLine !== previous.firstLine) Object.assign(values, {
      firstLine: undefined, hanging: undefined, firstLineChars: undefined, hangingChars: undefined,
      ...(style.firstLine === undefined ? {} : { [style.firstLine < 0 ? 'hanging' : 'firstLine']: twips(Math.abs(style.firstLine)) }),
    });
    updateAttributes('ind', values);
  }
}

function styleNode(root: Element, style: NamedStyle, defaultStyleId: string) {
  const attributes: Record<string, string> = { type: 'paragraph', styleId: style.id };
  if (style.id === defaultStyleId) attributes.default = '1';
  const node = add(root, 'style', attributes);
  add(node, 'name', { val: style.name });
  if (style.id !== defaultStyleId) add(node, 'basedOn', { val: style.basedOn ?? defaultStyleId });
  add(node, 'qFormat');
  const paragraph = add(node, 'pPr');
  writeParagraphFormat(paragraph, style.local?.paragraph ?? style.paragraph);
  const outline = style.local?.outline ?? (style.heading ? style.heading - 1 : undefined);
  if (outline !== undefined) add(paragraph, 'outlineLvl', { val: outline });
  writeFormat(add(node, 'rPr'), style.local?.format ?? style.format);
}

export async function exportDocx(model: DocModel, name: string, source?: DocxSource): Promise<File> {
  if (source?.originalModel === model && !source.template) return new File([source.file], name, { type: MIME });
  const parts: Parts = new Map(source?.parts);
  const mainPath = source?.mainPath ?? 'word/document.xml';
  const relsPath = source?.relsPath ?? 'word/_rels/document.xml.rels';
  const directory = mainPath.slice(0, mainPath.lastIndexOf('/') + 1);
  const main = source ? source.main.cloneNode(true) as XMLDocument : blankXml('document', W, 'w');
  const types = source ? source.types.cloneNode(true) as XMLDocument : blankXml('Types', TYPES);
  const rels = source?.rels ? source.rels.cloneNode(true) as XMLDocument : blankXml('Relationships', RELS);
  let typesChanged = !source || !!source.template;
  if (source?.template) {
    Array.from(types.documentElement.children).find(node => node.getAttribute('PartName') === '/' + mainPath)
      ?.setAttribute('ContentType', DOCUMENT_TYPE);
  }
  if (source?.originalModel === model) {
    putXml(parts, '[Content_Types].xml', types);
    return zip(parts, name);
  }
  let relsChanged = !source;
  const oldBlocks = new Map(source && [...source.originalModel.blocks,
    ...Object.values(source.originalModel.regions.items).flatMap(item => item.blocks)].map(block => [block.id, block]));
  const oldParagraphs = new Map(source && paragraphs(source.originalModel).map(item => [item.id, item]));
  const oldImages = new Map(Array.from(oldParagraphs.values()).flatMap(item => item.inlines.flatMap(inline =>
    inline.kind === 'image' ? [[inline.id, inline] as const] : [])));
  const oldRuns = new Map(Array.from(oldParagraphs.values()).flatMap(item => item.inlines.flatMap(inline =>
    inline.kind === 'text' && inline.source ? [[inline.source, inline] as const] : [])));
  const oldNode = (id: string) => source?.nodes.get(id);
  const clone = (node: Element) => main.importNode(node, true);
  const content = (path: string, type: string) => {
    if (Array.from(types.documentElement.children).some(node => node.getAttribute('PartName') === '/' + path)) return;
    types.documentElement.append(element(types, 'Override', { PartName: '/' + path, ContentType: type }, TYPES, ''));
    typesChanged = true;
  };
  const extension = (ext: string, type: string) => {
    if (Array.from(types.documentElement.children).some(node => node.getAttribute('Extension')?.toLowerCase() === ext)) return;
    types.documentElement.insertBefore(element(types, 'Default', { Extension: ext, ContentType: type }, TYPES, ''), types.documentElement.firstChild);
    typesChanged = true;
  };
  function makeRelationships(xml: XMLDocument, changed: () => void) {
    const used = new Set(Array.from(xml.documentElement.children).map(node => node.getAttribute('Id')));
    let index = 0;
    return (type: string, target: string, external = false) => {
      const existing = Array.from(xml.documentElement.children).find(node => node.getAttribute('Type') === R + '/' + type &&
        node.getAttribute('Target') === target && (node.getAttribute('TargetMode') === 'External') === external);
      if (existing) return existing.getAttribute('Id')!;
      extension('rels', 'application/vnd.openxmlformats-package.relationships+xml');
      let id: string;
      do { id = 'rId' + ++index; } while (used.has(id));
      used.add(id);
      changed();
      xml.documentElement.append(element(xml, 'Relationship', {
        Id: id, Type: R + '/' + type, Target: target, ...(external ? { TargetMode: 'External' } : {}),
      }, RELS, ''));
      return id;
    };
  }
  const relationship = makeRelationships(rels, () => { relsChanged = true; });
  let inlineRelationship = relationship;
  let inlineDirectory = directory;
  function newPartName(stem: string, extension: string, partDirectory = directory) {
    let suffix = '';
    let index = 0;
    const used = new Set(Array.from(parts.keys(), path => path.toLowerCase()));
    while (used.has((partDirectory + stem + suffix + '.' + extension).toLowerCase())) suffix = '-' + ++index;
    return stem + suffix + '.' + extension;
  }

  const body = child(main.documentElement, 'body') ?? add(main.documentElement, 'body');
  const section = child(body, 'sectPr') ?? element(body, 'sectPr');
  const pageChanged = model.page !== source?.originalModel.page;
  if (!source || !section.children.length || pageChanged) {
    child(section, 'pgSz')?.remove();
    const previousMargins = child(section, 'pgMar');
    const gutter = attr(previousMargins, 'gutter') ?? 0;
    previousMargins?.remove();
    const following = Array.from(section.children).find(node => !['headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type'].includes(node.localName));
    const size = element(section, 'pgSz', { w: twips(model.page.width), h: twips(model.page.height),
      ...(model.page.width > model.page.height ? { orient: 'landscape' } : {}) });
    section.insertBefore(size, following ?? null);
    const margins = element(section, 'pgMar', { top: twips(model.page.top), right: twips(model.page.right),
      bottom: twips(model.page.bottom), left: twips(model.page.left), header: twips(model.page.header), footer: twips(model.page.footer), gutter });
    section.insertBefore(margins, following ?? null);
  }
  const usedStyles = new Set(paragraphs(model).map(item => item.styleId).concat(model.defaultStyleId));
  const styles = source?.stylesPath ? await partXml(parts, source.stylesPath) : blankXml('styles', W, 'w');
  const defined = new Map(children(styles.documentElement, 'style').map(item => [attr(item, 'styleId'), item]));
  const originalStyles = new Map(source?.originalModel.styles.map(style => [style.id, style]));
  let stylesChanged = false;
  for (const style of model.styles) {
    if (style.type !== 'paragraph') continue;
    const previous = originalStyles.get(style.id);
    const existing = defined.get(style.id);
    if (!existing) {
      if (!usedStyles.has(style.id) && !style.local) continue;
      styleNode(styles.documentElement, style, model.defaultStyleId);
      stylesChanged = true;
      continue;
    }
    if (!previous || previous.local === style.local || !style.local) continue;
    if (style.name !== previous.name) {
      const styleName = child(existing, 'name');
      if (styleName) styleName.setAttributeNS(W, 'w:val', style.name);
      else existing.prepend(element(existing, 'name', { val: style.name }));
    }
    const props = child(existing, 'pPr') ?? element(existing, 'pPr');
    const format = child(existing, 'rPr') ?? element(existing, 'rPr');
    writeParagraphFormat(props, style.local.paragraph, previous.local?.paragraph);
    if (style.local.outline !== previous.local?.outline) {
      childProperty(props, 'outlineLvl', P_ORDER, style.local.outline === undefined ? undefined : { val: style.local.outline });
    }
    writeFormat(format, style.local.format, previous.local?.format);
    if (!props.parentNode && props.children.length) existing.insertBefore(props, child(existing, 'rPr') ?? null);
    if (!format.parentNode && format.children.length) existing.append(format);
    stylesChanged = true;
  }
  if (stylesChanged) {
    const target = source?.stylesPath ?? directory + newPartName('styles', 'xml');
    if (!source?.stylesPath) relationship('styles', target.slice(directory.length));
    content(target, 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml');
    putXml(parts, target, styles);
  }

  const listIds = new Map<string, string>();
  const newLists = new Map<string, Map<number, ListStyle>>();
  for (const item of paragraphs(model)) {
    const list = item.list;
    if (!list || /^\d+$/.test(list.id)) continue;
    const levels = newLists.get(list.id) ?? new Map();
    if (!levels.has(list.level)) levels.set(list.level, list);
    newLists.set(list.id, levels);
  }
  if (newLists.size) {
    const numbering = source?.numberingPath ? await partXml(parts, source.numberingPath) : blankXml('numbering', W, 'w');
    const root = numbering.documentElement;
    let absId = Math.max(-1, ...children(root, 'abstractNum').map(node => Number(attr(node, 'abstractNumId')) || 0));
    let numId = Math.max(0, ...children(root, 'num').map(node => Number(attr(node, 'numId')) || 0));
    for (const [id, levels] of newLists) {
      const abstract = element(root, 'abstractNum', { abstractNumId: ++absId });
      const maximum = Math.max(...levels.keys());
      add(abstract, 'multiLevelType', { val: maximum > 0 ? 'multilevel' : 'singleLevel' });
      const fallback = levels.values().next().value!;
      for (let index = 0; index <= maximum; index++) {
        const list = levels.get(index) ?? fallback;
        const level = add(abstract, 'lvl', { ilvl: index });
        add(level, 'start', { val: levels.get(index)?.start ?? 1 });
        add(level, 'numFmt', { val: list.kind === 'number' ? list.numberFormat ?? 'decimal' : 'bullet' });
        add(level, 'lvlText', { val: list.kind === 'number' ? list.pattern ?? '%' + (index + 1) + '.' : '•' });
        add(level, 'lvlJc', { val: 'left' });
        add(add(level, 'pPr'), 'ind', { left: 540 + 360 * index, hanging: 270 });
      }
      root.insertBefore(abstract, child(root, 'num') ?? child(root, 'numIdMacAtCleanup') ?? null);
      const number = element(root, 'num', { numId: ++numId });
      add(number, 'abstractNumId', { val: absId });
      root.insertBefore(number, child(root, 'numIdMacAtCleanup') ?? null);
      listIds.set(id, String(numId));
    }
    const target = source?.numberingPath ?? directory + newPartName('numbering', 'xml');
    if (!source?.numberingPath) relationship('numbering', target.slice(directory.length));
    content(target, 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml');
    putXml(parts, target, numbering);
  }

  function writeNumbering(props: Element, list?: ListStyle) {
    const old = child(props, 'numPr');
    childProperty(props, 'numPr', P_ORDER, {});
    const node = child(props, 'numPr')!;
    if (old && list && attr(child(old, 'numId')) === list.id) {
      node.replaceWith(old);
      return;
    }
    if (list) add(node, 'ilvl', { val: list.level });
    add(node, 'numId', { val: list ? listIds.get(list.id) ?? list.id : '0' });
  }

  let drawingId = Math.max(0, ...Array.from(main.getElementsByTagNameNS(WP, 'docPr')).map(node => Number(node.getAttribute('id')) || 0));
  const imageRelationships = new Map<string, string>();
  function picture(parent: Element, image: ImageRun) {
    const original = oldNode(image.id);
    if (original) {
      const node = clone(original);
      const previous = oldImages.get(image.id);
      if (previous && previous !== image) {
        const inline = child(child(node, 'drawing'), 'inline', WP);
        child(inline, 'docPr', WP)?.setAttribute('descr', image.alt);
        const graphics = child(child(inline, 'graphic', A), 'graphicData', A);
        const transform = child(child(child(graphics, 'pic', PIC), 'spPr', PIC), 'xfrm', A);
        for (const extent of [child(inline, 'extent', WP), child(transform, 'ext', A)]) {
          extent?.setAttribute('cx', String(emu(image.width)));
          extent?.setAttribute('cy', String(emu(image.height)));
        }
      }
      parent.append(node);
      return;
    }
    const ext = image.blob.type === 'image/png' ? 'png' : 'jpeg';
    let id = imageRelationships.get(image.id);
    if (!id) {
      const path = newPartName('media/image', ext, inlineDirectory);
      parts.set(inlineDirectory + path, image.blob);
      extension(ext, image.blob.type);
      content(inlineDirectory + path, image.blob.type);
      id = inlineRelationship('image', path);
      imageRelationships.set(image.id, id);
    }
    const drawing = add(add(parent, 'r'), 'drawing');
    const wp = (parent: Element, name: string, attrs?: Record<string, string | number>) =>
      parent.appendChild(element(parent, name, attrs, WP, 'wp'));
    const a = (parent: Element, name: string, attrs?: Record<string, string | number>) =>
      parent.appendChild(element(parent, name, attrs, A, 'a'));
    const pic = (parent: Element, name: string, attrs?: Record<string, string | number>) =>
      parent.appendChild(element(parent, name, attrs, PIC, 'pic'));
    const cx = emu(image.width), cy = emu(image.height);
    const inline = wp(drawing, 'inline', { distT: 0, distB: 0, distL: 0, distR: 0 });
    wp(inline, 'extent', { cx, cy });
    wp(inline, 'docPr', { id: ++drawingId, name: 'Picture ' + drawingId, descr: image.alt });
    a(wp(inline, 'cNvGraphicFramePr'), 'graphicFrameLocks', { noChangeAspect: 1 });
    const data = a(a(inline, 'graphic'), 'graphicData', { uri: PIC });
    const node = pic(data, 'pic');
    const nonVisual = pic(node, 'nvPicPr');
    pic(nonVisual, 'cNvPr', { id: 0, name: 'Picture' });
    pic(nonVisual, 'cNvPicPr');
    const fill = pic(node, 'blipFill');
    a(fill, 'blip').setAttributeNS(R, 'r:embed', id);
    a(a(fill, 'stretch'), 'fillRect');
    const shape = pic(node, 'spPr');
    const transform = a(shape, 'xfrm');
    a(transform, 'off', { x: 0, y: 0 });
    a(transform, 'ext', { cx, cy });
    a(a(shape, 'prstGeom', { prst: 'rect' }), 'avLst');
  }

  function textRun(parent: Element, run: TextRun) {
    const original = run.source ? oldNode(run.source) : undefined;
    const previous = run.source ? oldRuns.get(run.source) : undefined;
    if (original && previous === run && children(original, 'br').every(node => attr(node, 'type') !== 'page') &&
      !child(original, 'drawing')) {
      parent.append(clone(original));
      return;
    }
    const node = original ? clone(original) : element(parent, 'r');
    const properties = child(node, 'rPr') ?? element(node, 'rPr');
    writeFormat(properties, run.format, previous?.format);
    node.replaceChildren(...(properties.children.length || properties.attributes.length ? [properties] : []));
    for (const text of run.text.split(/([\t\n])/)) {
      if (!text) continue;
      if (text === '\t') add(node, 'tab');
      else if (text === '\n') add(node, 'br');
      else {
        const content = add(node, run.revision?.kind === 'del' ? 'delText' : 't');
        if (/^\s|\s$/.test(text)) content.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
        content.textContent = text;
      }
    }
    parent.append(node);
  }

  function writePageField(parent: Element, field: PageField) {
    const original = field.source && oldNode(field.source);
    if (original) { parent.append(clone(original)); return; }
    const node = add(parent, 'fldSimple', { instr: field.field, dirty: 1 });
    textRun(node, { kind: 'text', text: '1', format: field.format });
  }

  function writeParagraph(item: Paragraph): Element {
    const original = oldNode(item.id);
    const previous = oldParagraphs.get(item.id);
    if (original && (item === previous || item.locked)) return clone(original);
    const node = original ? clone(original) : element(main, 'p');
    const props = child(node, 'pPr') ?? element(node, 'pPr');
    if (item.styleId !== previous?.styleId) {
      childProperty(props, 'pStyle', P_ORDER, item.styleId !== model.defaultStyleId || previous ? { val: item.styleId } : undefined);
    }
    if (item.list !== previous?.list || (item.styleId !== previous?.styleId && item.list)) writeNumbering(props, item.list);
    writeParagraphFormat(props, item.style, previous?.style);
    if (original && item.inlines === previous?.inlines && item.comments === previous?.comments) {
      if (!props.parentNode && props.children.length) node.prepend(props);
      return node;
    }
    node.replaceChildren(...(props.children.length ? [props] : []));
    let lastLink: Link | undefined;
    let linkNode: Element | undefined;
    let lastRevision: Revision | undefined;
    let revisionNode: Element | undefined;
    function resetGroups() { lastLink = undefined; lastRevision = undefined; }
    function writeInlines(inlines: Paragraph['inlines']) {
      for (const inline of inlines) {
        if (inline.kind === 'pageField') { writePageField(node, inline); resetGroups(); continue; }
        if (inline.kind === 'image') { picture(node, inline); resetGroups(); continue; }
        if (inline.kind === 'pageBreak') {
          add(add(node, 'r'), 'br', { type: 'page' });
          resetGroups();
          continue;
        }
        if (inline.revision && lastRevision?.id !== inline.revision.id) {
          const { kind, id, author, date } = inline.revision;
          revisionNode = add(node, kind, { id, author, ...(date ? { date } : {}) });
        }
        if (inline.link && !sameLink(lastLink, inline.link)) {
          const from = inline.source ? oldNode(inline.source)?.parentElement : undefined;
          linkNode = from?.namespaceURI === W && from.localName === 'hyperlink' && sameLink(oldRuns.get(inline.source!)?.link, inline.link)
            ? main.importNode(from, false) : element(node, 'hyperlink');
          if (inline.link.anchor) linkNode.setAttributeNS(W, 'w:anchor', inline.link.anchor);
          if (!inline.link.href.startsWith('#')) {
            const id = inline.link.relationship ?? inlineRelationship('hyperlink', inline.link.href, true);
            linkNode.setAttributeNS(R, 'r:id', id);
          }
          node.append(linkNode);
        }
        textRun(inline.revision && revisionNode ? revisionNode : inline.link && linkNode ? linkNode : node, inline);
        lastLink = inline.link;
        lastRevision = inline.revision;
      }
    }
    if (!item.comments?.length) writeInlines(item.inlines);
    else {
      const points = Array.from(new Set(item.comments.flatMap(mark => [mark.from, mark.to]))).sort((a, b) => a - b);
      let start = 0;
      for (const point of points) {
        writeInlines(sliceInlines(item.inlines, start, point));
        resetGroups();
        for (const mark of item.comments.filter(mark => mark.from === point)) add(node, 'commentRangeStart', { id: mark.id });
        for (const mark of item.comments.filter(mark => mark.to === point)) {
          add(node, 'commentRangeEnd', { id: mark.id });
          add(add(node, 'r'), 'commentReference', { id: mark.id });
        }
        start = point;
      }
      writeInlines(sliceInlines(item.inlines, start, lengthOf(item.inlines)));
    }
    return node;
  }

  function writeTable(table: Table) {
    const original = oldNode(table.id);
    const node = original ? main.importNode(original, false) : element(main, 'tbl');
    const oldBlock = oldBlocks.get(table.id);
    const previous = oldBlock?.kind === 'table' ? oldBlock : undefined;
    const bordersChanged = previous?.borders !== table.borders;
    for (const item of Array.from(original?.children ?? [])) {
      if (item.localName !== 'tr' || item.namespaceURI !== W) node.append(clone(item));
    }
    const props = child(node, 'tblPr') ?? node.insertBefore(element(node, 'tblPr'), node.firstChild);
    if (!previous || previous.width !== table.width) {
      const width = table.width;
      childProperty(props, 'tblW', TABLE_ORDER, width
        ? { w: Math.round(width.value * (width.unit === 'percent' ? 50 : 15)), type: width.unit === 'percent' ? 'pct' : 'dxa' }
        : { w: 0, type: 'auto' });
    }
    if (previous?.align !== table.align) {
      childProperty(props, 'jc', TABLE_ORDER, table.align ? { val: table.align } : undefined);
    }
    if (!previous || bordersChanged) {
      const existing = child(props, 'tblBorders');
      if (!table.borders && previous) existing?.remove();
      else {
        childProperty(props, 'tblBorders', TABLE_ORDER, {});
        const borders = child(props, 'tblBorders')!;
        for (const side of TABLE_BORDER_SIDES) {
          const border = table.borders?.[side];
          if (border) add(borders, side, { val: border.style === 'solid' ? 'single' : border.style === 'none' ? 'nil' : border.style,
            sz: Math.round(border.size * 8), color: border.color.slice(1).toUpperCase() });
          else if (!table.borders && !previous) add(borders, side, { val: 'single', sz: Math.round(DEFAULT_TABLE_BORDER.size * 8),
            color: DEFAULT_TABLE_BORDER.color.slice(1).toUpperCase() });
        }
      }
    }
    if (previous?.margins !== table.margins) {
      childProperty(props, 'tblCellMar', TABLE_ORDER, table.margins ? {} : undefined);
      const margins = child(props, 'tblCellMar');
      if (margins) for (const side of ['top', 'left', 'bottom', 'right'] as const) {
        const value = table.margins?.[side];
        if (value !== undefined) add(margins, side, { w: twips(value), type: 'dxa' });
      }
    }

    const weights = tableColumnWeights(table);
    const oldWeights = previous && tableColumnWeights(previous);
    const gridChanged = !oldWeights || previous?.width !== table.width || weights.length !== oldWeights.length ||
      weights.some((width, index) => width !== oldWeights[index]);
    const available = model.page.width - model.page.left - model.page.right;
    const tableWidth = table.width?.unit === 'px' ? table.width.value : available * (table.width?.value ?? 100) / 100;
    const sum = weights.reduce((total, weight) => total + weight, 0);
    const widths = weights.map(weight => Math.max(1, twips(tableWidth * weight / sum)));
    if (gridChanged) {
      const grid = child(node, 'tblGrid') ?? add(node, 'tblGrid');
      for (const column of children(grid, 'gridCol')) column.remove();
      for (const width of widths) add(grid, 'gridCol', { w: width });
      childProperty(props, 'tblLayout', TABLE_ORDER, { type: 'fixed' });
    }

    const oldRows = new Map(previous?.rows.map(row => [row.id, row]));
    const oldCells = new Map(previous?.rows.flatMap(row => row.cells.map(cell => [cell.id, cell])));
    for (const row of table.rows) {
      const oldRow = oldNode(row.id);
      const previousRow = oldRows.get(row.id);
      const replaceRowBorders = bordersChanged && !!child(child(oldRow, 'tblPrEx'), 'tblBorders');
      if (oldRow && row === previousRow && !gridChanged && !replaceRowBorders) { node.append(clone(oldRow)); continue; }
      const rowNode = oldRow ? main.importNode(oldRow, false) : element(node, 'tr');
      for (const item of Array.from(oldRow?.children ?? [])) {
        if (item.localName !== 'tc' || item.namespaceURI !== W) rowNode.append(clone(item));
      }
      if (replaceRowBorders) child(child(rowNode, 'tblPrEx'), 'tblBorders')?.remove();
      if (row.header !== previousRow?.header || row.height !== previousRow?.height || row.exactHeight !== previousRow?.exactHeight) {
        const rowProps = child(rowNode, 'trPr') ?? element(rowNode, 'trPr');
        if (row.header !== previousRow?.header) childProperty(rowProps, 'tblHeader', ROW_ORDER, row.header ? {} : undefined);
        if (row.height !== previousRow?.height || row.exactHeight !== previousRow?.exactHeight) {
          childProperty(rowProps, 'trHeight', ROW_ORDER, row.height === undefined ? undefined :
            { val: twips(row.height), hRule: row.exactHeight ? 'exact' : 'atLeast' });
        }
        if (!rowProps.parentNode && rowProps.children.length) rowNode.append(rowProps);
      }
      let gridColumn = 0;
      for (const cell of row.cells) {
        const column = gridColumn;
        gridColumn += cell.colSpan ?? 1;
        const oldCell = oldNode(cell.id);
        const previousCell = oldCells.get(cell.id);
        if (oldCell && cell === previousCell && !gridChanged) { rowNode.append(clone(oldCell)); continue; }
        const cellNode = oldCell ? main.importNode(oldCell, false) : element(rowNode, 'tc');
        const originalProps = child(oldCell, 'tcPr');
        const cellProps = originalProps ? clone(originalProps) : element(cellNode, 'tcPr');
        if (previousCell?.borders && !cell.borders) {
          const borders = child(cellProps, 'tcBorders');
          for (const side of [...TABLE_BORDER_SIDES, 'start', 'end']) child(borders, side)?.remove();
          if (borders && !borders.children.length) borders.remove();
        }
        if (cell.borders && previousCell?.borders !== cell.borders) {
          childProperty(cellProps, 'tcBorders', CELL_ORDER, {});
          const borders = child(cellProps, 'tcBorders')!;
          for (const side of TABLE_BORDER_SIDES) {
            const border = cell.borders[side];
            if (border) add(borders, side, { val: border.style === 'solid' ? 'single' : border.style === 'none' ? 'nil' : border.style,
              sz: Math.round(border.size * 8), color: border.color.slice(1).toUpperCase() });
          }
        }
        if (cell.margins !== previousCell?.margins) {
          childProperty(cellProps, 'tcMar', CELL_ORDER, cell.margins ? {} : undefined);
          const margins = child(cellProps, 'tcMar');
          if (margins) for (const side of ['top', 'left', 'bottom', 'right'] as const) {
            if (cell.margins?.[side] !== undefined) add(margins, side, { w: twips(cell.margins[side]), type: 'dxa' });
          }
        }
        if (cell.colSpan !== previousCell?.colSpan) childProperty(cellProps, 'gridSpan', CELL_ORDER, cell.colSpan ? { val: cell.colSpan } : undefined);
        if (gridChanged || !oldCell) childProperty(cellProps, 'tcW', CELL_ORDER,
          { w: widths.slice(column, column + (cell.colSpan ?? 1)).reduce((sum, width) => sum + width, 0), type: 'dxa' });
        if (cell.fill !== previousCell?.fill) {
          childProperty(cellProps, 'shd', CELL_ORDER,
            cell.fill ? { val: 'clear', color: 'auto', fill: cell.fill.slice(1).toUpperCase() } : undefined);
        }
        if (cell.vertical !== previousCell?.vertical) {
          childProperty(cellProps, 'vAlign', CELL_ORDER, cell.vertical ? { val: cell.vertical } : undefined);
        }
        if (cellProps.children.length) cellNode.append(cellProps);
        for (const paragraph of cell.paragraphs) cellNode.append(writeParagraph(paragraph));
        rowNode.append(cellNode);
      }
      node.append(rowNode);
    }
    return node;
  }

  function writeBlock(block: Block) {
    const original = oldNode(block.id);
    if (original && (block === oldBlocks.get(block.id) || block.kind === 'opaque')) return clone(original);
    if (block.kind === 'paragraph') return writeParagraph(block);
    if (block.kind === 'table') return writeTable(block);
    throw new Error('Protected content lost its source XML.');
  }

  const oldRegions = source?.originalModel.regions;
  for (const kind of REGION_KINDS) {
    for (const variant of REGION_VARIANTS) {
      const key: RegionKey = `${kind}:${variant}`;
      const region = model.regions.items[key];
      if (region === oldRegions?.items[key]) continue;
      const original = source?.regionParts[key];
      const oldReference = children(section, kind + 'Reference').find(node => attr(node, 'type') === variant);
      if (!region) { oldReference?.remove(); continue; }
      const xml = original ? original.xml.cloneNode(true) as XMLDocument : blankXml(kind === 'header' ? 'hdr' : 'ftr', W, 'w');
      const target = original?.path ?? directory + newPartName(kind, 'xml');
      const relationPath = relationshipPath(target);
      const regionRels = parts.has(relationPath) ? await partXml(parts, relationPath) : blankXml('Relationships', RELS);
      let regionRelsChanged = false;
      inlineRelationship = makeRelationships(regionRels, () => { regionRelsChanged = true; });
      inlineDirectory = target.slice(0, target.lastIndexOf('/') + 1);
      xml.documentElement.replaceChildren();
      for (const block of region.blocks) xml.documentElement.append(writeBlock(block));
      inlineRelationship = relationship;
      inlineDirectory = directory;
      if (regionRelsChanged) putXml(parts, relationPath, regionRels);
      putXml(parts, target, xml);
      if (!original) {
        content(target, `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`);
        const id = relationship(kind, target.slice(directory.length));
        const reference = element(section, kind + 'Reference', { type: variant });
        reference.setAttributeNS(R, 'r:id', id);
        const next = Array.from(section.children).find(node => node.localName !== 'headerReference' &&
          (kind === 'header' || node.localName !== 'footerReference'));
        if (oldReference) oldReference.replaceWith(reference);
        else section.insertBefore(reference, next ?? null);
      }
    }
  }

  if (model.regions.first !== (oldRegions?.first ?? false)) {
    childProperty(section, 'titlePg', ['headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz',
      'pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote',
      'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange'],
    model.regions.first ? {} : undefined);
  }
  const evenChanged = model.regions.even !== (oldRegions?.even ?? false);
  const trackingChanged = model.tracking !== (source?.originalModel.tracking ?? false);
  if (evenChanged || trackingChanged) {
    const target = source?.settingsPath ?? directory + newPartName('settings', 'xml');
    const settings = source?.settingsPath ? await partXml(parts, target) : blankXml('settings', W, 'w');
    const root = settings.documentElement;
    const existing = child(root, 'evenAndOddHeaders');
    if (evenChanged && !model.regions.even) existing?.remove();
    else if (evenChanged && existing) existing.removeAttributeNS(W, 'val');
    else if (evenChanged) {
      const preceding = new Set(['writeProtection', 'view', 'zoom', 'removePersonalInformation', 'removeDateAndTime',
        'doNotDisplayPageBoundaries', 'displayBackgroundShape', 'printPostScriptOverText', 'printFractionalCharacterWidth',
        'printFormsData', 'embedTrueTypeFonts', 'embedSystemFonts', 'saveSubsetFonts', 'saveFormsData', 'mirrorMargins',
        'alignBordersAndEdges', 'bordersDoNotSurroundHeader', 'bordersDoNotSurroundFooter', 'gutterAtTop', 'hideSpellingErrors',
        'hideGrammaticalErrors', 'activeWritingStyle', 'proofState', 'formsDesign', 'attachedTemplate', 'linkStyles',
        'stylePaneFormatFilter', 'stylePaneSortMethod', 'documentType', 'mailMerge', 'revisionView', 'trackRevisions',
        'doNotTrackMoves', 'doNotTrackFormatting', 'documentProtection', 'autoFormatOverride', 'styleLockTheme',
        'styleLockQFSet', 'defaultTabStop', 'autoHyphenation', 'consecutiveHyphenLimit', 'hyphenationZone', 'doNotHyphenateCaps',
        'showEnvelope', 'summaryLength', 'clickAndTypeStyle', 'defaultTableStyle']);
      root.insertBefore(element(root, 'evenAndOddHeaders'), Array.from(root.children).find(node =>
        node.namespaceURI === W && !preceding.has(node.localName)) ?? null);
    }
    if (trackingChanged) {
      const track = child(root, 'trackRevisions');
      if (!model.tracking) track?.remove();
      else if (track) track.removeAttributeNS(W, 'val');
      else {
        const preceding = new Set(['writeProtection', 'view', 'zoom', 'removePersonalInformation', 'removeDateAndTime',
          'doNotDisplayPageBoundaries', 'displayBackgroundShape', 'printPostScriptOverText', 'printFractionalCharacterWidth',
          'printFormsData', 'embedTrueTypeFonts', 'embedSystemFonts', 'saveSubsetFonts', 'saveFormsData', 'mirrorMargins',
          'alignBordersAndEdges', 'bordersDoNotSurroundHeader', 'bordersDoNotSurroundFooter', 'gutterAtTop', 'hideSpellingErrors',
          'hideGrammaticalErrors', 'activeWritingStyle', 'proofState', 'formsDesign', 'attachedTemplate', 'linkStyles',
          'stylePaneFormatFilter', 'stylePaneSortMethod', 'documentType', 'mailMerge', 'revisionView']);
        root.insertBefore(element(root, 'trackRevisions'), Array.from(root.children).find(node =>
          node.namespaceURI === W && !preceding.has(node.localName)) ?? null);
      }
    }
    if (!source?.settingsPath) relationship('settings', target.slice(directory.length));
    content(target, 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
    putXml(parts, target, settings);
  }

  if (model.comments !== source?.originalModel.comments && (model.comments.length || source?.commentsPath)) {
    const target = source?.commentsPath ?? directory + newPartName('comments', 'xml');
    const xml = source?.commentsPath ? await partXml(parts, target) : blankXml('comments', W, 'w');
    const root = xml.documentElement;
    const existing = new Map(children(root, 'comment').map(node => [Number(attr(node, 'id')), node]));
    const kept = new Set(model.comments.map(comment => comment.id));
    for (const [id, node] of existing) if (!kept.has(id)) node.remove();
    for (const { id, author, date, text } of model.comments) {
      if (existing.has(id)) continue;
      const node = add(root, 'comment', { id, author, ...(date ? { date } : {}) });
      for (const line of text.split('\n')) {
        const p = add(node, 'p');
        textRun(p, { kind: 'text', text: line, format: {} });
      }
    }
    if (!source?.commentsPath) relationship('comments', target.slice(directory.length));
    content(target, 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml');
    putXml(parts, target, xml);
  }

  const bodyChanged = !source || model.blocks !== source.originalModel.blocks;
  const sectionChanged = pageChanged || model.regions !== oldRegions && (model.regions.first !== oldRegions?.first ||
    REGION_KINDS.some(kind => REGION_VARIANTS.some(variant => {
      const key: RegionKey = `${kind}:${variant}`;
      return !!model.regions.items[key] !== !!oldRegions?.items[key];
    })));
  if (bodyChanged) {
    body.replaceChildren();
    for (const block of model.blocks) body.append(writeBlock(block));
    body.append(section);
  } else if (sectionChanged && !section.parentNode) {
    body.append(section);
  }
  if (bodyChanged || sectionChanged) putXml(parts, mainPath, main);
  if (relsChanged) putXml(parts, relsPath, rels);
  if (!source) {
    const root = blankXml('Relationships', RELS);
    root.documentElement.append(element(root, 'Relationship', { Id: 'rId1', Type: R + '/officeDocument', Target: mainPath }, RELS, ''));
    putXml(parts, '_rels/.rels', root);
    extension('rels', 'application/vnd.openxmlformats-package.relationships+xml');
    extension('xml', 'application/xml');
    content(mainPath, DOCUMENT_TYPE);
  }
  if (typesChanged) putXml(parts, '[Content_Types].xml', types);
  return zip(parts, name);
}
