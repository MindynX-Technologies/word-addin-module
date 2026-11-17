import type { Tab } from './ribbon';

const groups = {
  'Files and search': [
    ['new', 'New document', 'Primary+N'], ['open', 'Open a DOCX', 'Primary+O'],
    ['save', 'Download the document', 'Primary+S', 'Alt+Shift+F2'], ['print', 'Print the document', 'Primary+P'],
    ['find', 'Find text', 'Primary+F'], ['replace', 'Find and replace', 'Primary+H', 'Control+H'],
    ['goto', 'Go to a page', 'Primary+G'], ['find-next', 'Find next match', 'F3'], ['find-previous', 'Find previous match', 'Control+Shift+G'],
  ],
  'Editing and clipboard': [
    ['undo', 'Undo', 'Primary+Z'], ['redo', 'Redo', 'Primary+Y', 'Primary+Shift+Z'],
    ['all', 'Select document or active header/footer', 'Primary+A'],
    ['cut', 'Cut selection', 'Primary+X'], ['copy', 'Copy selection', 'Primary+C'],
    ['paste', 'Paste', 'Primary+V'], ['paste-text', 'Paste as plain text', 'Primary+Shift+V'],
    ['copy-format', 'Copy character formatting', 'Primary+Alt+C'],
    ['paste-format', 'Apply copied character formatting', 'Primary+Alt+V'],
    ['backspace', 'Delete previous character', 'Backspace'], ['delete', 'Delete next character', 'Delete'],
    ['word-backspace', 'Delete previous word', 'Control+Backspace', 'Alt+Backspace'],
    ['word-delete', 'Delete next word', 'Control+Delete', 'Alt+Delete'],
  ],
  'Text formatting': [
    ['bold', 'Bold', 'Primary+B'], ['italic', 'Italic', 'Primary+I'], ['underline', 'Underline', 'Primary+U'],
    ['clear-text', 'Clear direct character formatting', 'Control+Space'],
    ['grow', 'Increase to the next font size', 'Primary+Shift+Period'],
    ['shrink', 'Decrease to the previous font size', 'Primary+Shift+Comma'],
    ['plus-point', 'Increase font size by one point', 'Primary+BracketRight'],
    ['minus-point', 'Decrease font size by one point', 'Primary+BracketLeft'],
    ['super', 'Superscript', 'Primary+Shift+Equal'], ['sub', 'Subscript', 'Primary+Shift+Minus', 'Primary+Equal'],
    ['font', 'Focus the font picker', 'Primary+Shift+F', 'Primary+D'],
    ['font-size', 'Focus the font size picker', 'Primary+Shift+P'],
    ['case', 'Cycle selected text case', 'Shift+F3'],
  ],
  'Paragraphs and styles': [
    ['left', 'Align left', 'Primary+L'], ['center', 'Center', 'Primary+E'], ['right', 'Align right', 'Primary+R'],
    ['justify', 'Justify', 'Primary+J'], ['indent', 'Increase paragraph indent', 'Primary+M', 'Control+Shift+M'],
    ['outdent', 'Decrease paragraph indent', 'Primary+Shift+M'],
    ['hang', 'Increase hanging indent', 'Primary+T'], ['unhang', 'Decrease hanging indent', 'Primary+Shift+T'],
    ['clear-paragraph', 'Clear direct paragraph formatting', 'Control+Q'],
    ['single', 'Single line spacing', 'Primary+1'], ['double', 'Double line spacing', 'Primary+2'],
    ['one-half', '1.5 line spacing', 'Primary+5'], ['normal', 'Apply the default paragraph style', 'Primary+Shift+N'],
    ['heading-1', 'Apply Heading 1', 'Primary+Alt+1'], ['heading-2', 'Apply Heading 2', 'Primary+Alt+2'],
    ['heading-3', 'Apply Heading 3', 'Primary+Alt+3'], ['bullets', 'Apply or remove bullets', 'Primary+Shift+L', 'Primary+Period'],
    ['styles', 'Focus paragraph styles', 'Primary+Shift+S', 'Primary+Alt+Shift+S'],
  ],
  'Insert and review': [
    ['link', 'Insert a hyperlink', 'Primary+K'], ['paragraph', 'Insert a paragraph', 'Enter'],
    ['line-break', 'Insert a line break', 'Shift+Enter'], ['page-break', 'Insert a page break', 'Primary+Enter'],
    ['page-field', 'Insert a page number in a header or footer', 'Alt+Shift+P'],
    ['nbsp', 'Insert a nonbreaking space', 'Primary+Shift+Space', 'Alt+Space'],
    ['nbhyphen', 'Insert a nonbreaking hyphen', 'Alt+Shift+Minus'], ['soft-hyphen', 'Insert an optional hyphen', 'Primary+Alt+Minus'],
    ['emdash', 'Insert an em dash', 'Primary+Alt+NumpadSubtract'], ['endash', 'Insert an en dash', 'Primary+NumpadSubtract'],
    ['ellipsis', 'Insert an ellipsis', 'Primary+Alt+Period'], ['registered', 'Insert ®', 'Primary+Alt+R'],
    ['trademark', 'Insert ™', 'Primary+Alt+T'],
    ['comment', 'Add a comment to selected text', 'Primary+Alt+M', 'Meta+Alt+A'],
    ['post-comment', 'Post a comment', 'Primary+Enter'],
    ['track', 'Toggle Track Changes', 'Primary+Shift+E'], ['close-pane', 'Close the review pane', 'Alt+Shift+C'],
  ],
  'Navigation and tables': [
    ['move', 'Move the cursor', 'ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'],
    ['extend', 'Extend selection with the arrow keys', 'Shift+ArrowLeft', 'Shift+ArrowRight', 'Shift+ArrowUp', 'Shift+ArrowDown'],
    ['line-start', 'Start of line', 'Home'], ['line-end', 'End of line', 'End'],
    ['word-move', 'Move by word', 'Control+ArrowLeft', 'Control+ArrowRight', 'Alt+ArrowLeft', 'Alt+ArrowRight'],
    ['prev-para', 'Previous paragraph boundary', 'Control+ArrowUp', 'Alt+ArrowUp'],
    ['next-para', 'Next paragraph boundary', 'Control+ArrowDown', 'Alt+ArrowDown'],
    ['select-prev-para', 'Select to the previous paragraph boundary', 'Control+Shift+ArrowUp', 'Alt+Shift+ArrowUp'],
    ['select-next-para', 'Select to the next paragraph boundary', 'Control+Shift+ArrowDown', 'Alt+Shift+ArrowDown'],
    ['doc-start', 'Start of document or active header/footer', 'Primary+Home', 'Meta+ArrowUp'],
    ['doc-end', 'End of document or active header/footer', 'Primary+End', 'Meta+ArrowDown'],
    ['select-start', 'Select to the start of the document', 'Primary+Shift+Home', 'Meta+Shift+ArrowUp'],
    ['select-end', 'Select to the end of the document', 'Primary+Shift+End', 'Meta+Shift+ArrowDown'],
    ['prev-page', 'Previous page', 'Primary+PageUp'], ['next-page', 'Next page', 'Primary+PageDown'],
    ['next-cell', 'Next table cell, or append a row', 'Tab'], ['prev-cell', 'Previous table cell', 'Shift+Tab'],
    ['row-start', 'First cell in the row', 'Alt+Home'], ['row-end', 'Last cell in the row', 'Alt+End'],
    ['column-start', 'First cell in the column', 'Alt+PageUp'], ['column-end', 'Last cell in the column', 'Alt+PageDown'],
    ['cell-tab', 'Insert a tab in a table cell', 'Control+Tab'],
    ['select-table', 'Select the entire table', 'Alt+Numpad5'],
  ],
  'Ribbon and help': [
    ['keys', 'Show ribbon KeyTips', 'Alt+Period', 'Control+Period', 'F10'],
    ['landmark-next', 'Focus the next editor region', 'Control+F6'],
    ['landmark-prev', 'Focus the previous editor region', 'Control+Shift+F6'],
    ['ribbon', 'Expand or collapse the ribbon', 'Control+F1'],
    ['shortcuts', 'Search keyboard shortcuts', 'Primary+Alt+Slash'],
    ['escape', 'Dismiss a dialog, KeyTips or active header/footer', 'Escape'],
    ['tab-keys', 'Select a ribbon tab with KeyTips', 'F', 'H', 'N', 'P', 'R', 'W'],
  ],
} as const;

type Entry = (typeof groups)[keyof typeof groups][number];
export type Shortcut = Entry[0];
const native = new Set<Shortcut>(['cut', 'copy', 'paste', 'paste-text', 'backspace', 'delete', 'word-backspace', 'word-delete',
  'paragraph', 'line-break', 'move', 'extend', 'line-start', 'line-end', 'word-move', 'post-comment', 'escape', 'tab-keys']);
const macOnly = new Set(['Meta+Alt+A', 'Alt+Space', 'Control+Period', 'Meta+ArrowUp', 'Meta+ArrowDown',
  'Meta+Shift+ArrowUp', 'Meta+Shift+ArrowDown', 'Control+Shift+M', 'Alt+ArrowUp', 'Alt+ArrowDown', 'Alt+Shift+ArrowUp',
  'Alt+Shift+ArrowDown', 'Alt+ArrowLeft', 'Alt+ArrowRight', 'Alt+Backspace', 'Alt+Delete']);
const pcOnly = new Set(['Alt+Period', 'Primary+M', 'Control+ArrowUp', 'Control+ArrowDown', 'Control+Shift+ArrowUp',
  'Control+Shift+ArrowDown', 'Control+ArrowLeft', 'Control+ArrowRight', 'Control+Backspace', 'Control+Delete']);

function platformKeys(entry: Entry, mac: boolean) {
  return entry.slice(2).filter(key => (mac || !macOnly.has(key)) && (!mac || !pcOnly.has(key)) &&
    (key !== 'Control+H' || mac)).map(key => key.replace('Primary', mac ? 'Meta' : 'Control'));
}

function entries() {
  return Object.entries(groups).map(([group, entries]) => ({ group, entries }));
}

export function shortcutMap(mac: boolean) {
  const map = new Map<string, Shortcut>();
  for (const { entries: group } of entries()) {
    for (const entry of group) {
      if (native.has(entry[0])) continue;
      for (const keys of platformKeys(entry, mac)) map.set(keys, entry[0]);
    }
  }
  return map;
}

export function keyChord(event: KeyboardEvent) {
  const printable = event.key.length === 1;
  const modified = event.ctrlKey || event.altKey || event.metaKey;
  const key = event.code === 'NumpadSubtract' || event.code === 'Numpad5' ? event.code
    : event.key === ' ' ? 'Space'
    : modified && (printable || event.key === 'Dead') ? event.code.replace(/^Key|^Digit/, '')
    : printable ? event.key.toUpperCase() : event.key;
  return [event.ctrlKey && 'Control', event.metaKey && 'Meta', event.altKey && 'Alt', event.shiftKey && 'Shift', key]
    .filter(Boolean).join('+');
}

const labels: Record<string, string> = {
  Control: 'Ctrl', Meta: '⌘', ArrowLeft: '←', ArrowRight: '→', ArrowUp: '↑', ArrowDown: '↓',
  Period: '.', Comma: ',', BracketLeft: '[', BracketRight: ']', Equal: '=', Minus: '-', NumpadSubtract: 'Num −', Numpad5: 'Num 5', Slash: '/',
};

export function shortcutGroups(mac: boolean) {
  return entries().map(({ group, entries: items }) => ({ group, entries: items.map(entry => ({
    id: entry[0], label: entry[1], keys: [...new Set(platformKeys(entry, mac))]
      .map(chord => chord.split('+').map(part => part === 'Alt' && mac ? 'Option' : labels[part] ?? part).join('+')),
  })) }));
}

export const tabTips: Record<Tab, string> = { File: 'F', Home: 'H', Insert: 'N', Layout: 'P', Review: 'R', View: 'W', 'Format Table': 'J' };

export const commandTips: Record<Tab, Record<string, string>> = {
  File: { N: 'New document', O: 'Open DOCX', A: 'Download DOCX', T: 'Document templates', P: 'Print document' },
  Home: {
    V: 'Paste from clipboard', B: 'Bold', I: 'Italic', U: 'Underline', 'FC': 'Text color',
    'FF': 'Font family', 'FS': 'Font size', 'ST': 'Paragraph style',
    'AL': 'Align left', 'AC': 'Align center', 'AR': 'Align right', 'AJ': 'Align justify',
    'UL': 'Bulleted list', 'NL': 'Numbered list', 'FD': 'Find in document', 'RE': 'Find and replace',
  },
  Insert: { T: 'Insert table', P: 'Insert picture', F: 'Insert template field', I: 'Insert link', B: 'Insert page break', H: 'Edit header', O: 'Edit footer' },
  Layout: { M: 'Page setup', B: 'Insert page break', H: 'Edit header', O: 'Edit footer' },
  Review: { C: 'New comment', T: 'Track changes', K: 'Review pane' },
  View: { O: 'Toggle outline', Q: 'Zoom', G: 'Go to page', A: 'Actual size', F: 'Fit to width', K: 'Keyboard shortcuts' },
  'Format Table': { S: 'Select table', P: 'Format table', D: 'Delete table', R: 'Table action', C: 'Cell shading' },
};
