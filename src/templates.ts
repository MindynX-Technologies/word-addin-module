import { importDocx } from './docx-read';
import { exportDocx } from './docx-write';
import { blockParagraphs, lengthOf, normalize, sliceInlines, withParagraphUpdates } from './model';
import type { DocModel, Inline, Paragraph } from './model';
import { DocxEditorError } from './types.js';

const FIELD = /\{\{\s*([a-z][a-z0-9_.-]{0,79})\s*\}\}/gi;

export interface TemplateField {
  name: string;
  occurrences: number;
}

export type TemplateValues = Readonly<Record<string, string>>;

function matches(item: Paragraph) {
  const text = item.inlines.map(run => run.kind === 'text' ? run.text : '\uFFFC').join('');
  return Array.from(text.matchAll(FIELD));
}

function candidates(model: DocModel) {
  const blocks = [...model.blocks, ...Object.values(model.regions.items).filter(region => !region.locked).flatMap(region => region.blocks)];
  return blockParagraphs(blocks).filter(item => !item.locked && !item.comments?.length &&
    !item.inlines.some(inline => inline.kind === 'text' && inline.revision));

}

export function fieldsIn(model: DocModel): TemplateField[] {
  const fields = new Map<string, number>();
  for (const paragraph of candidates(model)) {
    for (const [, name] of matches(paragraph)) fields.set(name, (fields.get(name) ?? 0) + 1);
  }
  return Array.from(fields, ([name, occurrences]) => ({ name, occurrences }));
}

export function fillTemplate(model: DocModel, values: TemplateValues) {
  const fields = fieldsIn(model);
  const missing = fields.filter(field => !Object.hasOwn(values, field.name) || typeof values[field.name] !== 'string');
  if (missing.length) throw new DocxEditorError('TEMPLATE_FIELDS', 'Fill every template field before creating the document.');
  const updates = new Map<string, Paragraph>();
  for (const item of candidates(model)) {
    const placeholders = matches(item);
    if (!placeholders.length) continue;
    const inlines: Inline[] = [];
    let offset = 0;
    for (const match of placeholders) {
      inlines.push(...sliceInlines(item.inlines, offset, match.index));
      const [original, name] = match;
      offset = match.index + original.length;
      const first = sliceInlines(item.inlines, match.index, match.index + 1)[0];
      const replacement = values[name].replace(/\r\n?/g, '\n').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');
      if (first?.kind === 'text' && replacement) inlines.push({ ...first, text: replacement });
    }
    inlines.push(...sliceInlines(item.inlines, offset, lengthOf(item.inlines)));
    updates.set(item.id, { ...item, inlines: normalize(inlines) });
  }
  return updates.size ? withParagraphUpdates(model, updates) : model;
}

export async function prepareTemplate(file: File) {
  const imported = await importDocx(file, true);
  if (imported.readOnlyReason) throw new DocxEditorError('TEMPLATE_UNSUPPORTED', imported.readOnlyReason);
  return imported;
}

export function templateFileName(name: string) {
  const base = name.replace(/\.do[ct]x$/i, '').trim().replace(/[\\/:*?"<>|\u0000-\u001f]/g, '-').slice(0, 120);
  return (base || 'Untitled') + '.docx';
}

export async function readTemplateFields(file: File): Promise<TemplateField[]> {
  return fieldsIn((await prepareTemplate(file)).model);
}

export async function createDocumentFromTemplate(file: File, values: TemplateValues, name = file.name): Promise<File> {
  const template = await prepareTemplate(file);
  const model = fillTemplate(template.model, values);
  return exportDocx(model, templateFileName(name), template.source);
}
