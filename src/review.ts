import { caret, collapsed, lengthOf, normalize, orderedSelection, paragraphScope, paragraphs, revisionAtBoundary, sliceInlines, updateParagraphs } from './model';
import type { DocComment, DocModel, Inline, ModelSelection, Revision, TextRun, Transaction } from './model';

export function revisions(model: DocModel) {
  return paragraphs(model).flatMap(item => {
    const groups = new Map<number, { revision: Revision; id: string; text: string; offset: number }>();
    let offset = 0;
    for (const run of item.inlines) {
      if (run.kind === 'text' && run.revision) {
        const group = groups.get(run.revision.id);
        if (group) group.text += run.text;
        else groups.set(run.revision.id, { revision: run.revision, id: item.id, text: run.text, offset });
      }
      offset += run.kind === 'text' ? run.text.length : 1;
    }
    return Array.from(groups.values());
  });
}

export function trackEdit(model: DocModel, selection: ModelSelection, inlines: Inline[], author: string, typing = false): Transaction | null {
  const range = orderedSelection(model, selection);
  if (!range || range.key || range.selected.length !== 1) return null;
  const item = range.selected[0];
  if (item.locked || item.comments?.length || revisionAtBoundary(item.inlines, range.start.offset)) return null;
  const old = sliceInlines(item.inlines, range.start.offset, range.end.offset);
  if ([...old, ...inlines].some(run => run.kind !== 'text' || run.revision || run.link)) return null;
  let id = 0;
  for (const change of revisions(model)) id = Math.max(id, change.revision.id);
  const date = new Date().toISOString();
  const before = sliceInlines(item.inlines, 0, range.start.offset);
  const after = sliceInlines(item.inlines, range.end.offset);
  const last = before.at(-1);
  const next = (kind: Revision['kind']): Revision => ({ id: ++id, kind, author, date });
  const deletion = next('del');
  const insertion = !old.length && typing && last?.kind === 'text' && last.revision?.kind === 'ins' &&
    last.revision.author === author ? last.revision : next('ins');
  const mark = (runs: Inline[], revision: Revision) => (runs as TextRun[]).map(run => ({ ...run, revision }));
  const transaction = updateParagraphs(model, selection, value => ({
    ...value, inlines: normalize([...before, ...mark(old, deletion), ...mark(inlines, insertion), ...after]),
  }));
  return transaction && { ...transaction, typing, selection: caret({
    id: item.id, offset: range.start.offset + lengthOf(old) + lengthOf(inlines),
  }) };
}

export function resolveRevision(model: DocModel, id: number, accept: boolean): Transaction | null {
  const group = revisions(model).find(item => item.revision.id === id);
  if (!group) return null;
  const target = caret({ id: group.id, offset: group.offset });
  const keep = accept === (group.revision.kind === 'ins');
  return updateParagraphs(model, target, item => ({
    ...item,
    inlines: normalize(item.inlines.flatMap(run => {
      if (run.kind !== 'text' || run.revision?.id !== id) return [run];
      return keep ? [{ ...run, revision: undefined }] : [];
    })),
  }));
}

export function createComment(model: DocModel, selection: ModelSelection, text: string, author: string): Transaction | null {
  const range = orderedSelection(model, selection);
  if (!range || collapsed(selection) || range.key || range.selected.length !== 1 ||
    range.selected[0].inlines.some(run => run.kind !== 'text' || run.revision)) return null;
  const item = range.selected[0];
  const id = Math.max(-1, ...model.comments.map(comment => comment.id)) + 1;
  const comment: DocComment = { id, author, date: new Date().toISOString(), text };
  return updateParagraphs({ ...model, comments: [...model.comments, comment] }, selection, value => ({
    ...value, comments: [...item.comments ?? [], { id, from: range.start.offset, to: range.end.offset }],
  }));
}

export function deleteComment(model: DocModel, id: number): Transaction | null {
  const item = paragraphs(model).find(value => value.comments?.some(mark => mark.id === id));
  if (!item || paragraphScope(model, item.id).key) return null;
  return updateParagraphs({ ...model, comments: model.comments.filter(value => value.id !== id) }, caret({ id: item.id, offset: 0 }), value => ({
    ...value, comments: value.comments?.filter(mark => mark.id !== id),
  }));
}
