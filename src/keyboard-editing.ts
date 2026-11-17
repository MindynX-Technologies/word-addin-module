import { caret, lengthOf, orderedSelection, paragraphScope, blockParagraphs, replaceRange, tableAt, editTable, sliceInlines } from './model';
import type { DocModel, Inline, ModelSelection, Position } from './model';
import type { Shortcut } from './shortcuts';

export function documentBoundary(model: DocModel, selection: ModelSelection, end: boolean): Position | undefined {
  const items = blockParagraphs(paragraphScope(model, selection.focus.id).blocks);
  const item = end ? items.at(-1) : items[0];
  return item && { id: item.id, offset: end ? lengthOf(item.inlines) : 0 };
}

export function paragraphBoundary(model: DocModel, selection: ModelSelection, forward: boolean): Position | undefined {
  const items = blockParagraphs(paragraphScope(model, selection.focus.id).blocks);
  const index = items.findIndex(item => item.id === selection.focus.id);
  const item = items[index];
  if (!item) return;
  if (!forward && selection.focus.offset > 0) return { id: item.id, offset: 0 };
  const next = items[index + (forward ? 1 : -1)];
  if (next) return { id: next.id, offset: 0 };
  return { id: item.id, offset: forward ? lengthOf(item.inlines) : 0 };
}

export function tableShortcut(model: DocModel, selection: ModelSelection, command: Shortcut) {
  const context = tableAt(model, selection);
  if (!context) return;
  const { table, rowIndex, columnIndex } = context;
  const cells = table.rows.flatMap(row => row.cells);
  const rowStart = table.rows.slice(0, rowIndex).reduce((sum, row) => sum + row.cells.length, 0);
  const logicalColumn = table.rows[rowIndex].cells.slice(0, columnIndex).reduce((sum, cell) => sum + (cell.colSpan ?? 1), 0);
  const columnCell = (row: typeof table.rows[number]) => {
    let column = 0;
    return row.cells.find(cell => { const end = column + (cell.colSpan ?? 1); const contains = column <= logicalColumn && logicalColumn < end; column = end; return contains; });
  };
  let index = rowStart + columnIndex;
  switch (command) {
    case 'next-cell': index++; break;
    case 'prev-cell': index--; break;
    case 'row-start': index = rowStart; break;
    case 'row-end': index = rowStart + table.rows[rowIndex].cells.length - 1; break;
    case 'column-start': index = cells.indexOf(columnCell(table.rows[0])!); break;
    case 'column-end': index = cells.indexOf(columnCell(table.rows.at(-1)!)!); break;
    default: return;
  }
  const cell = cells[index];
  if (cell) {
    const first = cell.paragraphs[0];
    const last = cell.paragraphs.at(-1)!;
    return { selection: { anchor: { id: first.id, offset: 0 }, focus: { id: last.id, offset: lengthOf(last.inlines) } } };
  }
  if (command === 'next-cell') return { transaction: editTable(model, selection, 'row-after') };
}

export function changeCase(model: DocModel, selection: ModelSelection) {
  const range = orderedSelection(model, selection);
  if (!range || range.selected.length !== 1 || range.start.offset === range.end.offset) return null;
  const inlines = sliceInlines(range.selected[0].inlines, range.start.offset, range.end.offset);
  if (inlines.some(inline => inline.kind !== 'text' || inline.revision)) return null;
  const text = inlines.map(inline => inline.kind === 'text' ? inline.text : '').join('');
  if (text.toLowerCase() === text.toUpperCase()) return null;
  const title = (value: string) => value.toLowerCase().replace(/(^|[^\p{L}\p{M}\p{N}])(\p{L})/gu, (_, before: string, letter: string) => before + letter.toUpperCase());
  const transform = text === text.toUpperCase() ? (value: string) => value.toLowerCase() :
    text === title(text) ? (value: string) => value.toUpperCase() : title;
  const lines: Inline[] = inlines.map(inline => inline.kind === 'text' ? { ...inline, text: transform(inline.text) } : inline);
  const changed = replaceRange(model, selection, [lines]);
  return changed && { ...changed, selection: { anchor: range.start, focus: changed.selection.focus } };
}

export function textMatches(model: DocModel, query: string): ModelSelection[] {
  if (!query.trim()) return [];
  const pattern = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'giu');
  return blockParagraphs(model.blocks).flatMap(item => {
    const text = item.inlines.map(run => run.kind === 'text' ? run.text : '\uFFFC').join('');
    return Array.from(text.matchAll(pattern), match => ({
      anchor: { id: item.id, offset: match.index },
      focus: { id: item.id, offset: match.index + match[0].length },
    }));
  });
}

export function replaceMatches(model: DocModel, matches: ModelSelection[], text: string) {
  let current = model;
  let count = 0;
  let selection: ModelSelection | undefined;
  for (const target of [...matches].reverse()) {
    const scope = orderedSelection(current, target);
    if (!scope || scope.selected[0].locked) continue;
    const first = sliceInlines(scope.selected[0].inlines, target.anchor.offset, target.focus.offset)[0];
    if (first?.kind !== 'text') continue;
    const changed = replaceRange(current, target, [text ? [{ ...first, text }] : []]);
    if (!changed) continue;
    current = changed.model;
    selection = caret(changed.selection.focus);
    count++;
  }
  return { count, transaction: selection ? { model: current, selection } : null };
}
