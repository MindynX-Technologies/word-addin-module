import {
  forwardRef, useCallback, useEffect, useId, useImperativeHandle, useLayoutEffect, useMemo, useRef, useState,
} from 'react';
import type { CSSProperties, KeyboardEvent as ReactKeyboardEvent } from 'react';
import { AlignLeft, ChevronRight, Download, FileText, FolderOpen, Minus, PanelLeft, Plus, X } from 'lucide-react';
import { importDocx } from './docx-read';
import type { DocxSource, ImportedDocx } from './docx-read';
import { exportDocx as writeDocx } from './docx-write';
import { browserSelection, buildMeasures, paginate, PageNumbers, RenderFragment, SelectedPicture, restoreSelection, useLayoutWidth } from './layout';
import type { Fragment, Page, TableMeasurements } from './layout';
import {
  activeFormat, blankModel, blockParagraphs, caret, collapsed, DEFAULT_STYLES, deletionRange, editableStart, formatRange, formatTable,
  editTable, imageAt, insertTable, nodeId, selectWholeTable, wholeTableSelection, orderedSelection, paragraph, paragraphScope, paragraphWordCount, pasteBlocks, plainText, regionForPage, replaceRange, shadeCell, tableAt, updateNamedStyle, updateParagraphs,
} from './model';
import type { DocModel, ImageRun, Inline, ModelSelection, NamedStyle, RegionKey, RegionKind, Table, TableAction, TextFormat, Transaction } from './model';
import { readStyleFields, StyleFields } from './style-dialog';
import { clipboardText, readClipboard } from './clipboard';
import type { PasteMode } from './clipboard';
import { prepareClipboardImages, selectionHasImages, writeClipboard, writeImageClipboard } from './clipboard-write';
import { createComment, deleteComment, resolveRevision, revisions, trackEdit } from './review';
import { safeHref } from './ooxml';
import { Ribbon } from './ribbon';
import type { DialogName, Tab } from './ribbon';
import { keyChord, shortcutMap, commandTips, tabTips } from './shortcuts';
import type { Shortcut } from './shortcuts';
import { TemplateDialog } from './template-dialog';
import { readTableFormatting, TableFormatFields } from './table-format';
import { ShortcutGuide } from './shortcut-guide';
import { changeCase, documentBoundary, paragraphBoundary, replaceMatches, tableShortcut, textMatches } from './keyboard-editing';
import { DocxEditorError } from './types.js';
import type { DocxEditorHandle, DocxEditorProps } from './types.js';
import { useDraftRecovery } from './recovery';

type Session = {
  model: DocModel;
  saved: DocModel;
  fileName: string;
  source?: DocxSource;
  readOnlyReason?: string;
  warnings: string[];
  legacyOriginal?: File;
};
type History = { model: DocModel; selection: ModelSelection | null };
type DocumentDialog = DialogName | 'discard' | 'picture-size' | 'table-delete' | 'table-format';
type NativeEdit = { range: ModelSelection; restore: () => void; initialText: string; read: () => string };
const TABS: Tab[] = ['File', 'Home', 'Insert', 'Layout', 'Review', 'View', 'Format Table'];
const useBrowserLayoutEffect = typeof window === 'undefined' ? useEffect : useLayoutEffect;

function newSession(): Session {
  const model = blankModel();
  return { model, saved: model, fileName: 'Untitled.docx', warnings: [] };
}

function errorOf(cause: unknown, code: 'OPEN_FAILED' | 'EXPORT_FAILED', message: string) {
  return cause instanceof DocxEditorError ? cause : new DocxEditorError(code, message, { cause });
}

function downloadFile(file: File, ownerDocument: Document) {
  const url = URL.createObjectURL(file);
  const link = ownerDocument.createElement('a');
  link.href = url;
  link.download = file.name;
  ownerDocument.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

function snapshotSelection(root: HTMLElement, range: ModelSelection): NativeEdit | undefined {
  const anchor = root.ownerDocument.getSelection()?.anchorNode;
  const element = anchor?.nodeType === Node.ELEMENT_NODE ? anchor as Element : anchor?.parentElement;
  const paragraph = element?.closest<HTMLElement>('[data-paragraph]') ??
    root.querySelector<HTMLElement>('[data-paragraph="' + range.anchor.id + '"]');
  if (!paragraph || !root.contains(paragraph)) return;
  function snapshot(node: Node): () => void {
    if (node.nodeType === Node.TEXT_NODE) {
      const value = node.nodeValue;
      return () => { if (node.nodeValue !== value) node.nodeValue = value; };
    }
    const children = Array.from(node.childNodes);
    const restore = children.map(snapshot);
    return () => {
      if (children.length !== node.childNodes.length || children.some((child, index) => node.childNodes[index] !== child)) {
        (node as Element).replaceChildren(...children);
      }
      restore.forEach(reset => reset());
    };
  }
  const read = () => paragraph.textContent ?? '';
  return { range, restore: snapshot(root), initialText: read(), read };
}

function insertedText(previous: string, current: string) {
  let start = 0;
  let end = 0;
  while (start < previous.length && start < current.length && previous[start] === current[start]) start++;
  while (end < previous.length - start && end < current.length - start &&
    previous[previous.length - end - 1] === current[current.length - end - 1]) end++;
  return current.slice(start, current.length - end);
}

export const DocxEditor = forwardRef<DocxEditorHandle, DocxEditorProps>(function DocxEditor(
  { initialDocument, readOnly = false, className, onChange, onError, convertDocument, recoveryKey, reviewAuthor = 'Editor' }, ref,
) {
  const [session, setSession] = useState(newSession);
  const sessionRef = useRef(session);
  const [tab, setTab] = useState<Tab>('Home');
  const [pasteMode, setPasteMode] = useState<PasteMode>('source');
  const [pasteNotice, setPasteNotice] = useState('');
  const [showReview, setShowReview] = useState(false);
  const [outline, setOutline] = useState(true);
  const [zoom, setZoom] = useState<number | null>(null);
  const [pageNumber, setPageNumber] = useState(1);
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [error, setError] = useState<DocxEditorError>();
  const [keyTips, setKeyTips] = useState<{ tab?: Tab; prefix: string }>();
  const [ribbonCollapsed, setRibbonCollapsed] = useState(false);
  const [dialog, setDialog] = useState<DocumentDialog>();
  const [regionEdit, setRegionEdit] = useState<{ key: RegionKey; page: number } | null>(null);
  const regionRef = useRef(regionEdit);
  const [selectedImage, setSelectedImage] = useState<string>();
  const findCursor = useRef({ query: '', index: -1 });
  const lastFindQuery = useRef('');
  const copiedFormat = useRef<TextFormat | undefined>(undefined);
  const plainPaste = useRef(false);
  const mac = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/i.test(navigator.platform);
  const shortcuts = useMemo(() => shortcutMap(mac), [mac]);
  const pendingTableAction = useRef<TableAction | undefined>(undefined);
  const styleForDialog = useRef<NamedStyle | undefined>(undefined);
  const [selection, setSelection] = useState<ModelSelection | null>(() => editableStart(session.model));
  const selectionRef = useRef(selection);
  const [typingFormat, setTypingFormat] = useState<TextFormat>({});
  const typingRef = useRef<TextFormat>({});
  const history = useRef<{ undo: History[]; redo: History[]; lastTyping?: number }>({ undo: [], redo: [] });
  const openSequence = useRef(0);
  const initialRef = useRef<File | undefined>(undefined);
  const editingRoot = useRef<HTMLDivElement>(null);
  const viewport = useRef<HTMLDivElement>(null);
  const measurement = useRef<HTMLDivElement>(null);
  const tableMeasurements = useRef<TableMeasurements>(new Map());
  const fileInput = useRef<HTMLInputElement>(null);
  const approvedFileSelection = useRef<DocModel | null>(null);
  const imageInput = useRef<HTMLInputElement>(null);
  const imageSelection = useRef<{ selection: ModelSelection; model: DocModel } | null>(null);
  const dialogRef = useRef<HTMLDialogElement>(null);
  const selectedForDialog = useRef<ModelSelection | null>(null);
  const discard = useRef<(() => void) | null>(null);
  const pendingSelection = useRef<{ selection: ModelSelection | null; focus: boolean } | null>(null);
  const nativeEdit = useRef<NativeEdit | undefined>(undefined);
  const composing = useRef(false);
  const callbacks = useRef({ onChange, onError });
  callbacks.current = { onChange, onError };
  const ribbonId = useId();

  const { model } = session;
  const editorRoot = useRef<HTMLElement>(null);
  const editorWidth = useLayoutWidth(editorRoot);
  const viewportWidth = useLayoutWidth(viewport);
  const canShowOutline = editorWidth === 0 || editorWidth >= 625;
  const phone = viewportWidth > 0 && typeof window !== 'undefined' &&
    window.matchMedia('(pointer: coarse) and (max-width: 760px)').matches;
  const disabled = readOnly || !!session.readOnlyReason || phone || loading;
  const disabledRef = useRef(disabled);
  disabledRef.current = disabled;
  const dirty = session.model !== session.saved;
  const availableWidth = viewportWidth ? Math.max(160, viewportWidth - (viewportWidth < 620 ? 24 : 56)) : model.page.width;
  const scale = zoom === null ? Math.min(1, availableWidth / model.page.width) : zoom / 100;
  const measures = useMemo(() => buildMeasures(model), [model]);
  const regionFragments = useMemo(() => Object.fromEntries(Object.entries(model.regions.items).map(([key, region]) =>
    [key, buildMeasures(model, region.blocks, false).map(measure => measure.fragment)])), [model]);
  const [layout, setLayout] = useState<{ model: DocModel; pages: Page[]; top: number; bottom: number } | null>(null);
  const pages = layout?.model === model ? layout.pages : [measures.map(measure => measure.fragment)];
  const stats = useMemo(() => {
    const items = blockParagraphs(model.blocks);
    return {
      items,
      words: items.reduce((sum, item) => sum + paragraphWordCount(item), 0),
      headings: items.flatMap(item => {
        const style = model.styles.find(style => style.id === item.styleId);
        const text = plainText(item.inlines).trim();
        return style?.heading && text ? [{ id: item.id, text, level: style.heading }] : [];
      }),
    };
  }, [model]);
  const active = selection ? activeFormat(model, selection) : undefined;
  const activeTable = selection && tableAt(model, selection);
  const selectedTable = selection && wholeTableSelection(model, selection);
  const reviewChanges = useMemo(() => revisions(model), [model]);
  const activeImage = selectedImage ? imageAt(model, selectedImage) : undefined;
  const tableContextId = activeTable && !disabled && !activeImage ? activeTable.table.id : undefined;
  const availableTabs = phone ? TABS.filter(value => value === 'File' || value === 'View')
    : TABS.filter(value => value !== 'Format Table' || !!tableContextId);
  const visibleTab = availableTabs.includes(tab) ? tab : phone ? 'View' : 'Home';

  useEffect(() => {
    if (tableContextId) { setTab('Format Table'); setRibbonCollapsed(false); }
    else setTab(current => current === 'Format Table' ? 'Home' : current);
    setKeyTips(undefined);
  }, [tableContextId]);

  const report = useCallback((problem: DocxEditorError) => {
    if (problem.code !== 'OPEN_SUPERSEDED') {
      setError(problem);
      callbacks.current.onError?.(problem);
    }
  }, []);
  const recovery = useDraftRecovery(recoveryKey, session, report);
  const author = reviewAuthor.trim() || 'Editor';
  useEffect(() => prepareClipboardImages(model), [model]);

  const updateSession = useCallback((next: Session) => {
    sessionRef.current = next;
    setSession(next);
  }, []);

  const updateSelection = useCallback((next: ModelSelection | null) => {
    selectionRef.current = next;
    setSelection(next);
  }, []);

  const resetTyping = useCallback(() => {
    if (!Object.keys(typingRef.current).length) return;
    typingRef.current = {};
    setTypingFormat({});
  }, []);

  function updateRegion(next: typeof regionEdit) {
    regionRef.current = next;
    setRegionEdit(next);
    setSelectedImage(undefined);
  }

  function currentSelection(native = editingRoot.current && browserSelection(editingRoot.current)) {
    const saved = selectionRef.current;
    if (!native || (saved?.tableId && sameSelectionEndpoints(saved, native))) return saved;
    if (native.tableId) {
      const table = wholeTableSelection(sessionRef.current.model, native);
      if (table) return selectWholeTable(table);
    }
    return native;
  }

  function focusSelection(next: ModelSelection | null) {
    if (!next || !editingRoot.current) return;
    editingRoot.current.focus({ preventScroll: true });
    restoreSelection(editingRoot.current, next);
    updateSelection(next);
  }

  function apply(transaction: Transaction | null, previous: ModelSelection | null = selectionRef.current, focus = true) {
    if (!transaction || disabledRef.current) return false;
    const current = sessionRef.current;
    const undo = history.current;
    const now = performance.now();
    if (!transaction.typing || undo.lastTyping === undefined || now - undo.lastTyping > 850) {
      undo.undo.push({ model: current.model, selection: previous });
      if (undo.undo.length > 120) undo.undo.shift();
    }
    undo.lastTyping = transaction.typing ? now : undefined;
    undo.redo = [];
    pendingSelection.current = { selection: transaction.selection, focus };
    updateSelection(transaction.selection);
    updateSession({ ...current, model: transaction.model });
    return true;
  }

  function travel(direction: 'undo' | 'redo') {
    if (disabledRef.current) return;
    const stack = history.current;
    const target = stack[direction].pop();
    if (!target) return;
    const current = sessionRef.current;
    stack[direction === 'undo' ? 'redo' : 'undo'].push({ model: current.model, selection: selectionRef.current });
    stack.lastTyping = undefined;
    resetTyping();
    const key = target.selection && paragraphScope(target.model, target.selection.focus.id).key;
    if (key) updateRegion({ key, page: regionRef.current?.page ?? pageNumber });
    else updateRegion(null);
    pendingSelection.current = { selection: target.selection, focus: true };
    updateSelection(target.selection);
    updateSession({ ...current, model: target.model });
  }

  const resetDocument = useCallback((next: Session) => {
    nativeEdit.current?.restore();
    nativeEdit.current = undefined;
    composing.current = false;
    history.current = { undo: [], redo: [] };
    regionRef.current = null;
    setRegionEdit(null);
    setSelectedImage(undefined);
    setPasteNotice('');
    setShowReview(false);
    findCursor.current = { query: '', index: -1 };
    lastFindQuery.current = '';
    setKeyTips(undefined);
    const start = editableStart(next.model);
    resetTyping();
    pendingSelection.current = null;
    updateSelection(start);
    updateSession(next);
    setError(undefined);
    setTab('Home');
    setPageNumber(1);
    if (viewport.current) viewport.current.scrollTop = 0;
  }, [resetTyping, updateSelection, updateSession]);

  const open = useCallback(async (file: File, options?: { discardChanges?: boolean }) => {
    const starting = sessionRef.current;
    const id = ++openSequence.current;
    try {
      if (!options?.discardChanges && starting.model !== starting.saved) {
        throw new DocxEditorError('UNSAVED_CHANGES', 'Save or discard the current changes before opening another document.');
      }
      setLoading(true);
      setError(undefined);
      let importedFile = file;
      if (/\.doc$/i.test(file.name) && convertDocument) importedFile = await convertDocument(file, 'docx');
      if (id !== openSequence.current) throw new DocxEditorError('OPEN_SUPERSEDED', 'Another document was opened first.');
      const incoming = await importDocx(importedFile);
      if (id !== openSequence.current) throw new DocxEditorError('OPEN_SUPERSEDED', 'Another document was opened first.');
      if (sessionRef.current.model !== starting.model) {
        throw new DocxEditorError('UNSAVED_CHANGES', 'The current document changed while the file was opening.');
      }
      resetDocument({
        model: incoming.model, saved: incoming.model, fileName: importedFile.name,
        legacyOriginal: importedFile !== file ? file : undefined,
        source: incoming.source, warnings: incoming.warnings, readOnlyReason: incoming.readOnlyReason,
      });
    } catch (cause) {
      const problem = id !== openSequence.current
        ? new DocxEditorError('OPEN_SUPERSEDED', 'Another open request replaced this one.')
        : errorOf(cause, 'OPEN_FAILED', 'The document could not be opened.');
      report(problem);
      throw problem;
    } finally {
      if (id === openSequence.current) setLoading(false);
    }
  }, [report, resetDocument, convertDocument]);

  const exportCurrent = useCallback(async () => {
    const snapshot = sessionRef.current;
    try {
      return await writeDocx(snapshot.model, snapshot.fileName, snapshot.source);
    } catch (cause) {
      const problem = errorOf(cause, 'EXPORT_FAILED', 'The document could not be exported.');
      report(problem);
      throw problem;
    }
  }, [report]);

  const exportLegacy = useCallback(async () => {
    const snapshot = sessionRef.current;
    try {
      if (snapshot.legacyOriginal && snapshot.model === snapshot.source?.originalModel) return snapshot.legacyOriginal;
      if (!convertDocument) throw new DocxEditorError('UNSUPPORTED_FILE', 'Configure a Word converter to export legacy DOC files.');
      return await convertDocument(await exportCurrent(), 'doc');
    } catch (cause) {
      const problem = cause instanceof DocxEditorError ? cause : new DocxEditorError('CONVERSION_FAILED', 'The document could not be converted to DOC.', { cause });
      report(problem); throw problem;
    }
  }, [convertDocument, exportCurrent, report]);

  useImperativeHandle(ref, () => ({ open, exportDocx: exportCurrent, exportDoc: exportLegacy, flushRecovery: recovery.flush }),
    [open, exportCurrent, exportLegacy, recovery.flush]);

  useEffect(() => {
    if (initialDocument && !initialRef.current) {
      initialRef.current = initialDocument;
      void open(initialDocument).catch(() => {});
    }
  }, [initialDocument, open]);

  useEffect(() => {
    callbacks.current.onChange?.({ fileName: session.fileName, dirty, readOnly: readOnly || !!session.readOnlyReason || phone });
  }, [session.fileName, session.readOnlyReason, dirty, readOnly, phone]);

  useEffect(() => {
    if (!dirty) return;
    const owner = editorRoot.current?.ownerDocument.defaultView;
    const guard = (event: BeforeUnloadEvent) => { event.preventDefault(); event.returnValue = ''; };
    owner?.addEventListener('beforeunload', guard);
    return () => owner?.removeEventListener('beforeunload', guard);
  }, [dirty]);

  useBrowserLayoutEffect(() => {
    if (!measurement.current || composing.current) return;
    const measure = () => {
      if (!measurement.current || composing.current) return;
      const box = measurement.current;
      const regionHeight = (kind: RegionKind) => Math.max(0, ...Array.from(box.querySelectorAll<HTMLElement>('[data-measure-region^="' + kind + ':"]'))
        .filter(node => !node.dataset.measureRegion?.endsWith(':first') || model.regions.first)
        .filter(node => !node.dataset.measureRegion?.endsWith(':even') || model.regions.even)
        .map(node => node.getBoundingClientRect().height));
      const header = regionHeight('header'), footer = regionHeight('footer');
      const top = Math.max(model.page.top, header ? model.page.header + header + 6 : 0);
      const bottom = Math.max(model.page.bottom, footer ? model.page.footer + footer + 6 : 0);
      setLayout({ model, top, bottom,
        pages: paginate(measures, box, Math.max(60, model.page.height - top - bottom), model.styles, tableMeasurements.current, scale),
      });
    };
    measure();
    const fonts = measurement.current.ownerDocument.fonts;
    const fontsLoaded = () => {
      tableMeasurements.current.clear();
      measure();
    };
    fonts?.addEventListener('loadingdone', fontsLoaded);
    return () => fonts?.removeEventListener('loadingdone', fontsLoaded);
  }, [model, measures, regionFragments, scale]);

  useBrowserLayoutEffect(() => {
    if (layout?.model !== model || !pendingSelection.current || !editingRoot.current || composing.current) return;
    const next = pendingSelection.current;
    pendingSelection.current = null;
    if (next.selection) {
      if (next.focus) editingRoot.current.focus({ preventScroll: true });
      restoreSelection(editingRoot.current, next.selection);
      selectionRef.current = next.selection.tableId ? next.selection : browserSelection(editingRoot.current) ?? next.selection;
      const native = editingRoot.current.ownerDocument.getSelection();
      const focus = native?.focusNode?.parentElement?.closest('[data-page]');
      if (focus) focusCaretIntoView(native!, viewport.current!);
    }
  }, [layout, model]);

  function trackSelection() {
    if (composing.current || pendingSelection.current || !editingRoot.current) return;
    const native = browserSelection(editingRoot.current);
    if (!native) return;
    const next = currentSelection(native);
    const previous = selectionRef.current;
    if (!next || (previous && previous.tableId === next.tableId && sameSelectionEndpoints(previous, next))) return;
    history.current.lastTyping = undefined;
    resetTyping();
    updateSelection(next);
  }

  useEffect(() => {
    const owner = editingRoot.current?.ownerDocument;
    owner?.addEventListener('selectionchange', trackSelection);
    return () => owner?.removeEventListener('selectionchange', trackSelection);
  });

  function textTransaction(range: ModelSelection, text: string, typing = false) {
    const document = sessionRef.current.model;
    const value = activeFormat(document, range);
    const format = { ...value.direct, ...typingRef.current };
    const lines = text.replace(/\r\n?/g, '\n').split('\n').map((line): Inline[] =>
      line ? [{ kind: 'text', text: line, format, link: value.link, source: value.source }] : []);
    if (document.tracking) return lines.length === 1 ? trackEdit(document, range, lines[0], author, typing) : null;
    return replaceRange(document, range, lines, typing);
  }

  function insertText(range: ModelSelection, text: string, typing = false) {
    const inserted = apply(textTransaction(range, text, typing), range);
    if (!inserted && sessionRef.current.model.tracking) {
      report(new DocxEditorError('PROTECTED_CONTENT', 'Track Changes supports text edits within one paragraph. Turn it off to edit other content.'));
    }
    return inserted;
  }

  async function pasteContent(range: ModelSelection, html: string, text: string, mode: PasteMode, images: Blob[] = []) {
    const current = sessionRef.current.model;
    const warnings: string[] = [];
    const blocks = readClipboard(html, current, mode, warnings);
    const plain = text || clipboardText(blocks);
    if (mode === 'merge') {
      const destination = { ...activeFormat(current, range).format, ...typingRef.current };
      for (const item of blockParagraphs(blocks)) {
        item.inlines = item.inlines.map(run => run.kind === 'text' ? { ...run, format: { ...destination, ...run.format } } : run);
      }
    }
    const textFallback = !!html && !blocks.length && !!plain && warnings.length > 0;
    if (!current.tracking && mode !== 'text' && !textFallback) {
      const existingPicture = blockParagraphs(blocks).some(item => item.inlines.some(run => run.kind === 'image'));
      if (!existingPicture && images.length) {
        for (const image of images.slice(0, 8)) {
          try {
            const inline = await decodePicture(image, current, image instanceof File && image.name ? image.name : 'Pasted picture');
            if (sessionRef.current.model !== current) return;
            blocks.push(paragraph(current.defaultStyleId, [inline]));
          } catch {
            warnings.push('Some clipboard pictures could not be decoded. Open the DOCX to preserve embedded pictures.');
          }
        }
      }
      if (sessionRef.current.model !== current) return;
      if (blocks.length) {
        const usedStyles = new Set(blockParagraphs(blocks).map(item => item.styleId));
        const headings = DEFAULT_STYLES.filter(style => usedStyles.has(style.id) && !current.styles.some(value => value.id === style.id));
        const updated = headings.length ? { ...current, styles: [...current.styles, ...headings] } : current;
        if (apply(pasteBlocks(updated, range, blocks), range)) {
          setPasteNotice([...new Set(warnings)].join(' '));
          return;
        }
        warnings.push('Rich formatting could not be inserted at this selection. Plain text was used instead.');
      }
    }
    if (plain && !insertText(range, plain)) warnings.push('Paste could not change this selection. Choose editable text outside protected content.');
    if (!plain && (html || images.length)) warnings.push('No usable content could be pasted. Try copying plain text or opening the DOCX.');
    setPasteNotice([...new Set(warnings)].join(' '));
  }

  async function pasteFromClipboard() {
    const range = currentSelection();
    const snapshot = sessionRef.current.model;
    if (!range || disabled) return;
    try {
      let html = '', text = '';
      const images: Blob[] = [];
      const clipboard = navigator.clipboard;
      if (clipboard.read) {
        for (const item of await clipboard.read()) {
          if (!html && item.types.includes('text/html')) html = await (await item.getType('text/html')).text();
          if (!text && item.types.includes('text/plain')) text = await (await item.getType('text/plain')).text();
          const image = item.types.find(type => type === 'image/png' || type === 'image/jpeg');
          if (image) images.push(await item.getType(image));
        }
      } else text = await clipboard.readText();
      if (snapshot === sessionRef.current.model) await pasteContent(range, html, text, pasteMode, images);
    } catch {
      setPasteNotice('Use Ctrl+V or Command+V to paste from your clipboard.');
      focusSelection(range);
    }
  }

  function transferPictures(range: ModelSelection, cut: boolean, data: DataTransfer) {
    const snapshot = sessionRef.current.model;
    const table = wholeTableSelection(snapshot, range);
    const transaction = cut && !table ? snapshot.tracking ? trackEdit(snapshot, range, [], author) : replaceRange(snapshot, range, [[]]) : null;
    if (cut && !table && !transaction) {
      report(new DocxEditorError('PROTECTED_CONTENT', 'This selection cannot be cut while its content or review marks are protected.'));
      return;
    }
    let written: Promise<void>;
    try {
      if (typeof navigator.clipboard?.write === 'function' && typeof ClipboardItem !== 'undefined') written = writeImageClipboard(snapshot, range);
      else { writeClipboard(data, snapshot, range); written = Promise.resolve(); }
    } catch (cause) { written = Promise.reject(cause); }
    void written.then(() => {
      const selected = currentSelection();
      if (cut) {
        if (sessionRef.current.model !== snapshot || !selected || !sameSelectionEndpoints(selected, range) || disabledRef.current) {
          setPasteNotice('Copied with pictures. The document changed during copying, so the original selection was kept.');
          return;
        }
        if (table) deleteSelectedTable(range); else apply(transaction, range);
      }
      setPasteNotice('Copied with embedded pictures. Picture-only selections also include a PNG for Word.');
    }).catch(cause => report(new DocxEditorError('CLIPBOARD_FAILED', 'The clipboard could not receive the pictures. Your document was kept unchanged. Check clipboard permission and try again.', { cause })));
  }

  function findText(query: string, direction = 1) {
    const matches = textMatches(sessionRef.current.model, query);
    if (!matches.length) return false;
    lastFindQuery.current = query;
    if (findCursor.current.query !== query) findCursor.current = { query, index: direction > 0 ? -1 : 0 };
    const index = (findCursor.current.index + direction + matches.length) % matches.length;
    const found = matches[index];
    findCursor.current.index = index;
    gotoHeading(found.anchor.id);
    if (editingRoot.current) restoreSelection(editingRoot.current, found);
    updateSelection(found);
    return true;
  }

  function setTextFormat(change: TextFormat, target = currentSelection()) {
    if (disabledRef.current || !target) return;
    const document = sessionRef.current.model;
    if (collapsed(target)) {
      if (activeFormat(document, target).item?.locked) return;
      typingRef.current = { ...typingRef.current, ...change };
      setTypingFormat(typingRef.current);
      focusSelection(target);
    } else {
      apply(formatRange(document, target, run => ({ ...run, format: { ...run.format, ...change } })), target);
    }
  }

  function clearFormatting() {
    const target = currentSelection();
    if (!target || disabled) return;
    if (collapsed(target)) {
      const named = model.styles.find(style => style.id === activeFormat(model, target).item?.styleId);
      setTextFormat({
        bold: false, italic: false, underline: false, strike: false, vertical: 'baseline', highlight: undefined,
        ...named?.format,
      }, target);
    } else apply(formatRange(model, target, run => ({ ...run, format: {}, source: undefined })), target);
  }

  function withParagraphs(update: Parameters<typeof updateParagraphs>[2]) {
    const target = currentSelection();
    if (target) apply(updateParagraphs(sessionRef.current.model, target, update), target);
  }

  function insertInline(inline: Inline, target = currentSelection()) {
    if (target) apply(replaceRange(sessionRef.current.model, target, [[inline]]), target);
  }

  function handleInput(event: InputEvent) {
    const root = editingRoot.current;
    if (!root) return;
    const browserRange = browserSelection(root, event.getTargetRanges?.()[0]);
    const selected = currentSelection();
    const range = selected && !collapsed(selected) ? selected : browserRange ?? selected;
    if (disabledRef.current || !range) {
      if (event.cancelable) event.preventDefault();
      return;
    }
    if (composing.current || event.isComposing) return;
    if (!event.cancelable) {
      nativeEdit.current = snapshotSelection(root, range);
      return;
    }
    event.preventDefault();
    const document = sessionRef.current.model;
    if (event.inputType === 'insertText' || event.inputType === 'insertReplacementText') {
      insertText(range, event.data ?? event.dataTransfer?.getData('text/plain') ?? '', event.inputType === 'insertText');
    } else if (event.inputType === 'insertParagraph') {
      const item = activeFormat(document, range).item;
      if (item?.list && !item.inlines.length && collapsed(range)) withParagraphs(value => ({ ...value, list: undefined }));
      else insertText(range, '\n');
    } else if (event.inputType === 'insertLineBreak') {
      const format = activeFormat(document, range);
      insertInline({ kind: 'text', text: '\n', format: { ...format.direct, ...typingRef.current } }, range);
    } else if (event.inputType.startsWith('delete')) {
      if (wholeTableSelection(document, range)) {
        deleteSelectedTable(range);
        return;
      }
      const target = deletionRange(document, range, event.inputType.endsWith('Backward'), event.inputType.startsWith('deleteWord'));
      if (target) {
        if (document.tracking) apply(trackEdit(document, target, [], author), range);
        else apply(replaceRange(document, target, [[]]), range);
      }
    } else if (event.inputType === 'historyUndo') travel('undo');
    else if (event.inputType === 'historyRedo') travel('redo');
    else if (event.inputType === 'formatBold' || event.inputType === 'formatItalic' || event.inputType === 'formatUnderline') {
      const key = event.inputType.slice(6).toLowerCase() as 'bold' | 'italic' | 'underline';
      setTextFormat({ [key]: !{ ...activeFormat(document, range).format, ...typingRef.current }[key] }, range);
    }
  }

  function reconcileNative(text: string | null, edit = nativeEdit.current) {
    nativeEdit.current = undefined;
    if (!edit || !editingRoot.current) return;
    const value = text ?? insertedText(edit.initialText, edit.read());
    edit.restore();
    const committed = value && apply(textTransaction(edit.range, value), edit.range);
    if (!committed) restoreSelection(editingRoot.current, edit.range);
  }

  useEffect(() => {
    const root = editingRoot.current;
    root?.addEventListener('beforeinput', handleInput as EventListener);
    return () => root?.removeEventListener('beforeinput', handleInput as EventListener);
  });

  function requestReplacement(action: () => void) {
    const current = sessionRef.current;
    if (current.model === current.saved) action();
    else {
      discard.current = action;
      setDialog('discard');
    }
  }

  function chooseDocument() {
    requestReplacement(() => {
      approvedFileSelection.current = sessionRef.current.model;
      fileInput.current?.click();
    });
  }

  function closeDialog() {
    dialogRef.current?.close();
    setDialog(undefined);
  }

  useEffect(() => {
    const element = dialogRef.current;
    if (dialog && element && !element.open) element.showModal();
  }, [dialog]);

  async function download(format: 'doc' | 'docx' = 'docx') {
    if (exporting || loading) return;
    const snapshot = sessionRef.current;
    setExporting(true);
    try {
      const file = await (format === 'doc' ? exportLegacy() : exportCurrent());
      downloadFile(file, editingRoot.current?.ownerDocument ?? document);
      const current = sessionRef.current;
      if (current.model === snapshot.model) {
        updateSession({ ...current, saved: snapshot.model });
      }
    } catch {
      // exportCurrent reports the failure.
    } finally {
      setExporting(false);
    }
  }

  async function attachImage(file: File, target: ModelSelection, snapshot: DocModel) {
    try {
      const image = await decodePicture(file, snapshot, file.name.replace(/\.[^.]+$/, ''));
      if (sessionRef.current.model === snapshot) insertInline(image, target);
    } catch (cause) {
      report(cause instanceof DocxEditorError ? cause : new DocxEditorError('INVALID_IMAGE', 'This image could not be opened.', { cause }));
    }
  }

  function setList(kind: 'bullet' | 'number') {
    const target = currentSelection();
    if (!target) return;
    const document = sessionRef.current.model;
    const range = orderedSelection(document, target);
    const previous = range?.selected[0];
    if (!previous) return;
    const items = blockParagraphs(range.blocks);
    const sibling = items[items.indexOf(previous) - 1];
    const inherited = sibling?.list?.kind === kind ? sibling.list : undefined;
    const all = range.selected.every(item => item.list?.kind === kind);
    const list = all ? undefined : inherited ?? { kind, id: nodeId('list'), level: 0 };
    apply(updateParagraphs(document, target, item => ({ ...item, list })), target);
  }

  function scrollStatus() {
    const pane = viewport.current;
    if (!pane) return;
    const top = pane.getBoundingClientRect().top + 40;
    const sheets = Array.from(pane.querySelectorAll('[data-page]'));
    const index = sheets.findIndex(sheet => sheet.getBoundingClientRect().bottom > top);
    if (index >= 0) setPageNumber(index + 1);
  }

  function gotoHeading(id: string) {
    const heading = editingRoot.current?.querySelector<HTMLElement>('[data-paragraph="' + id + '"]');
    if (!heading || !viewport.current) return;
    const pane = viewport.current;
    pane.scrollTop += heading.getBoundingClientRect().top - pane.getBoundingClientRect().top - 30;
    if (!disabled) focusSelection(caret({ id, offset: 0 }));
  }

  function editPageRegion(kind: RegionKind, requestedPage?: number) {
    if (disabled) return;
    const current = sessionRef.current.model;
    const native = editingRoot.current?.ownerDocument.getSelection();
    const selectionPage = native?.focusNode?.parentElement?.closest<HTMLElement>('[data-page]')?.dataset.page;
    const page = Math.max(1, Math.min(requestedPage ?? (selectionPage ? Number(selectionPage) : pageNumber), pages.length));
    const key = regionForPage(current.regions, kind, page);
    const existing = current.regions.items[key];
    if (existing?.locked) {
      report(new DocxEditorError('PROTECTED_CONTENT', 'This header or footer contains protected content.'));
      return;
    }
    const first = existing ? blockParagraphs(existing.blocks).find(item => !item.locked) : { ...paragraph(current.defaultStyleId), style: { before: 0, after: 0 } };
    if (!first) {
      report(new DocxEditorError('PROTECTED_CONTENT', 'This header or footer contains protected content.'));
      return;
    }
    const target = caret({ id: first.id, offset: 0 });
    updateRegion({ key, page });
    resetTyping();
    setTab('Home');
    if (!existing) {
      apply({ model: { ...current, regions: { ...current.regions, items: {
        ...current.regions.items, [key]: { blocks: [first] },
      } } }, selection: target });
    } else {
      updateSelection(target);
      pendingSelection.current = { selection: target, focus: true };
    }
  }

  useBrowserLayoutEffect(() => {
    if (!regionEdit || !editingRoot.current || layout?.model !== model || !pendingSelection.current) return;
    const pending = pendingSelection.current;
    pendingSelection.current = null;
    if (pending.selection) focusSelection(pending.selection);
  }, [regionEdit, model, layout]);

  function closePageRegion() {
    updateRegion(null);
    const start = editableStart(sessionRef.current.model);
    updateSelection(start);
    editingRoot.current?.ownerDocument.getSelection()?.removeAllRanges();
  }

  function changeRegionSetting(key: 'first' | 'even', value: boolean) {
    const current = sessionRef.current.model;
    const target = editableStart(current);
    if (!target) return;
    closePageRegion();
    apply({ model: { ...current, regions: { ...current.regions, [key]: value } }, selection: target }, target, false);
  }

  function removePageRegion() {
    if (!regionEdit) return;
    const current = sessionRef.current.model;
    const target = editableStart(current);
    if (!target || current.regions.items[regionEdit.key]?.locked) return;
    const items = { ...current.regions.items };
    delete items[regionEdit.key];
    closePageRegion();
    apply({ model: { ...current, regions: { ...current.regions, items } }, selection: target }, target, false);
  }

  function addPageField(field: 'PAGE' | 'NUMPAGES') {
    const target = currentSelection();
    if (!target || !regionEdit || paragraphScope(model, target.focus.id).key !== regionEdit.key) return;
    insertInline({ kind: 'pageField', field, format: { ...activeFormat(model, target).direct, ...typingRef.current } }, target);
  }

  const chooseTable = useCallback((table: Table) => {
    if (disabledRef.current || !editingRoot.current) return;
    resetTyping();
    history.current.lastTyping = undefined;
    setSelectedImage(undefined);
    const selection = selectWholeTable(table);
    editingRoot.current.focus({ preventScroll: true });
    restoreSelection(editingRoot.current, selection);
    updateSelection(selection);
  }, [resetTyping, updateSelection]);

  function deleteSelectedTable(range: ModelSelection) {
    const document = sessionRef.current.model;
    const table = wholeTableSelection(document, range);
    if (!table || disabledRef.current) return false;
    const target = selectWholeTable(table);
    if (apply(editTable(document, target, 'delete-table'), target)) return true;
    report(new DocxEditorError('PROTECTED_CONTENT', 'This table contains protected content, comments, or tracked changes.'));
    return false;
  }

  function performTableAction(action: TableAction) {
    const target = currentSelection();
    if (!target || disabled) return;
    if (action.startsWith('delete')) {
      selectedForDialog.current = target;
      pendingTableAction.current = action;
      setDialog('table-delete');
    } else if (!apply(editTable(sessionRef.current.model, target, action), target)) {
      report(new DocxEditorError('PROTECTED_CONTENT', 'This table cannot be expanded further.'));
    }
  }

  function updatePicture(width: number, height: number, alt: string) {
    if (!activeImage || activeImage.paragraph.locked || width < 1 || height < 1 || width > 4000 || height > 4000) return;
    const { image, selection: start } = activeImage;
    const target = { anchor: start.anchor, focus: { ...start.focus, offset: start.focus.offset + 1 } };
    insertInline({ ...image, width, height, alt }, target);
  }

  function fit() {
    setZoom(null);
  }

  function createFromTemplate(model: DocModel, template: ImportedDocx, fileName: string) {
    setDialog(undefined);
    requestReplacement(() => {
      const saved = blankModel();
      resetDocument({ model, saved, fileName, source: template.source, warnings: template.warnings });
    });
  }

  function openCommandDialog(name: DocumentDialog) {
    setKeyTips(undefined);
    const target = currentSelection();
    selectedForDialog.current = target;
    styleForDialog.current = styles.find(style => style.id === (target ? activeFormat(model, target).item?.styleId : model.defaultStyleId));
    setDialog(name);
  }

  function applyStyle(styleId: string) {
    const target = currentSelection();
    const document = sessionRef.current.model;
    if (!target) return;
    const style = styles.find(item => item.id === styleId);
    const updated = style && !document.styles.some(item => item.id === styleId)
      ? { ...document, styles: [...document.styles, style] } : document;
    apply(updateParagraphs(updated, target, item => ({ ...item, styleId })), target);
  }

  function alignParagraph(alignment: 'left' | 'center' | 'right' | 'justify') {
    withParagraphs(item => ({ ...item, style: { ...item.style, alignment } }));
  }

  function changeIndent(direction: number, hanging = false) {
    const key = hanging ? 'firstLine' : 'indent';
    withParagraphs(item => {
      const original = item.style[key] ?? styles.find(style => style.id === item.styleId)?.paragraph[key] ?? 0;
      const next = hanging ? Math.min(0, original - direction * 24) : Math.max(0, original + direction * 24);
      return { ...item, style: { ...item.style, [key]: next } };
    });
  }

  function toggleTracking() {
    const target = currentSelection();
    if (target) apply({ model: { ...model, tracking: !model.tracking }, selection: target }, target);
    setShowReview(true);
  }

  function moveSelection(position: ModelSelection['focus'], anchor = position) {
    const next = { anchor, focus: position };
    focusSelection(next);
    const native = editingRoot.current?.ownerDocument.getSelection();
    if (native && viewport.current) focusCaretIntoView(native, viewport.current);
  }

  function printDocument() {
    const editor = editorRoot.current;
    const document = editor?.ownerDocument;
    const view = document?.defaultView;
    if (!editor || !document || !view || !editingRoot.current || document.querySelector('.rwe-print-document')) return;
    const print = document.createElement('div');
    print.className = editor.className + ' rwe-print-document';
    const canvas = editingRoot.current.cloneNode(true) as HTMLElement;
    canvas.contentEditable = 'false';
    canvas.removeAttribute('role');
    for (const region of canvas.querySelectorAll('[data-region-editing]')) region.setAttribute('data-region-editing', 'false');
    print.append(canvas);
    const sheet = document.createElement('style');
    sheet.textContent = `@page { size: ${model.page.width}px ${model.page.height}px; margin: 0; }`;
    document.head.append(sheet);
    document.body.append(print);
    const cleanup = () => {
      print.remove();
      sheet.remove();
    };
    view.addEventListener('afterprint', cleanup, { once: true });
    view.print();
  }

  function goToPage(page: number) {
    const root = editingRoot.current;
    const pane = viewport.current;
    const sheet = root?.querySelector<HTMLElement>('[data-page="' + page + '"]');
    if (!root || !sheet || !pane) return;
    const item = sheet.querySelector<HTMLElement>('[data-paragraph]:not([data-page-region] *, [data-repeat-row] *)');
    if (item?.dataset.paragraph) moveSelection({ id: item.dataset.paragraph, offset: Number(item.dataset.from ?? 0) });
    pane.scrollTop += sheet.getBoundingClientRect().top - pane.getBoundingClientRect().top - 12;
    setPageNumber(page);
  }

  function ribbonControl(label: string) {
    return Array.from(editorRoot.current?.querySelectorAll<HTMLElement>('button, select, input') ?? [])
      .find(control => control.getAttribute('aria-label') === label && !control.matches(':disabled') &&
        !control.closest('[hidden]'));
  }

  function focusRibbonControl(label: string) {
    setTab('Home');
    setRibbonCollapsed(false);
    requestAnimationFrame(() => {
      const control = ribbonControl(label);
      if (control?.getClientRects().length) control.focus();
      else {
        const overflow = ribbonControl('More formatting');
        if (overflow?.getAttribute('aria-expanded') === 'false') overflow.click();
        requestAnimationFrame(() => ribbonControl(label)?.focus());
      }
    });
  }

  function handleShortcut(event: ReactKeyboardEvent<HTMLElement>) {
    if (event.defaultPrevented || event.nativeEvent.isComposing || composing.current ||
      event.nativeEvent.getModifierState('AltGraph')) return;
    const target = event.target as HTMLElement;
    const chord = keyChord(event.nativeEvent);
    const onCanvas = !!editingRoot.current?.contains(target);
    const primary = mac ? event.metaKey : event.ctrlKey;
    if (dialog) {
      if (dialog === 'comment' && primary && event.key === 'Enter') {
        event.preventDefault();
        target.closest('form')?.requestSubmit();
      }
      return;
    }

    if (keyTips && chord === 'Escape') {
      event.preventDefault();
      setKeyTips(undefined);
      return;
    }
    if (keyTips && /^[A-Z]$/.test(chord)) {
      event.preventDefault();
      if (!keyTips.tab) {
        const next = availableTabs.find(value => tabTips[value] === chord);
        if (next) { setTab(next); setKeyTips({ tab: next, prefix: '' }); }
        return;
      }
      const prefix = keyTips.prefix + chord;
      const commands = commandTips[keyTips.tab];
      const label = commands[prefix];
      if (label) {
        const control = ribbonControl(label);
        setKeyTips(undefined);
        const activate = (target: HTMLElement | undefined) => {
          if (target?.tagName === 'BUTTON') target.click();
          else target?.focus();
        };
        if (control?.getClientRects().length) activate(control);
        else {
          const overflow = ribbonControl('More formatting');
          if (overflow?.getAttribute('aria-expanded') === 'false') overflow.click();
          requestAnimationFrame(() => activate(ribbonControl(label)));
        }
      } else if (Object.keys(commands).some(tip => tip.startsWith(prefix))) setKeyTips({ ...keyTips, prefix });
      else setKeyTips({ ...keyTips, prefix: '' });
      return;
    }

    const command = shortcuts.get(chord);
    if (command === 'keys') {
      event.preventDefault();
      setRibbonCollapsed(false);
      setKeyTips(keyTips ? undefined : { prefix: '' });
      return;
    }
    if (command === 'ribbon') {
      event.preventDefault();
      setKeyTips(undefined);
      setRibbonCollapsed(value => !value);
      return;
    }
    if (command === 'landmark-next' || command === 'landmark-prev') {
      event.preventDefault();
      const root = editorRoot.current;
      const stops = [root?.querySelector<HTMLElement>('.rwe-title-actions button:not(:disabled)'),
        root?.querySelector<HTMLElement>('[role="tab"][aria-selected="true"]'), editingRoot.current,
        showReview ? root?.querySelector<HTMLElement>('.rwe-review-pane header button') : null,
        root?.querySelector<HTMLElement>('.rwe-zoom input')].filter((node): node is HTMLElement => !!node);
      const index = stops.indexOf(target);
      const direction = command === 'landmark-next' ? 1 : -1;
      const next = stops[((index < 0 && direction < 0 ? 0 : index) + direction + stops.length) % stops.length];
      if (next === editingRoot.current) focusSelection(currentSelection());
      else next?.focus();
      return;
    }
    if (command === 'shortcuts') {
      event.preventDefault();
      openCommandDialog('shortcuts');
      return;
    }
    if (!onCanvas) return;
    if (chord === 'Escape') {
      const range = currentSelection();
      if (range && wholeTableSelection(sessionRef.current.model, range)) {
        event.preventDefault();
        focusSelection(caret(range.anchor));
      } else if (regionRef.current) { event.preventDefault(); closePageRegion(); }
      else if (showReview) { event.preventDefault(); setShowReview(false); }
      return;
    }

    if (primary && !event.altKey && event.code === 'KeyV') {
      plainPaste.current = event.shiftKey;
      if (event.shiftKey) setTimeout(() => { plainPaste.current = false; }, 0);
      return;
    }
    const deleting = chord === 'Delete' || chord === 'Backspace';
    if (!command && (disabled || !deleting)) return;
    const range = currentSelection();
    const document = sessionRef.current.model;
    if (!disabled && range && deleting && wholeTableSelection(document, range)) {
      event.preventDefault();
      deleteSelectedTable(range);
      return;
    }
    if (!command) return;
    const run = (action: () => void) => { event.preventDefault(); action(); };

    switch (command) {
      case 'new': return run(() => requestReplacement(() => resetDocument(newSession())));
      case 'open': return run(chooseDocument);
      case 'save': return run(() => void download());
      case 'print': return run(printDocument);
      case 'find': return run(() => openCommandDialog('find'));
      case 'replace': return disabled ? undefined : run(() => openCommandDialog('replace'));
      case 'goto': return run(() => openCommandDialog('goto'));
      case 'find-next':
      case 'find-previous':
        return run(() => {
          if (lastFindQuery.current) findText(lastFindQuery.current, command === 'find-next' ? 1 : -1);
          else openCommandDialog('find');
        });
      case 'close-pane': return run(() => setShowReview(false));
      case 'prev-page':
      case 'next-page': return run(() => goToPage(Math.max(1, Math.min(pages.length, pageNumber + (command === 'next-page' ? 1 : -1)))));
      case 'all':
      case 'doc-start':
      case 'doc-end':
      case 'select-start':
      case 'select-end':
        if (!range) return;
        return run(() => {
          const end = command === 'all' || command.endsWith('end');
          const position = documentBoundary(document, range, end);
          const anchor = command === 'all' ? documentBoundary(document, range, false) : command.startsWith('select') ? range.anchor : position;
          if (position && anchor) moveSelection(position, anchor);
        });
      case 'prev-para':
      case 'next-para':
      case 'select-prev-para':
      case 'select-next-para':
        if (!range) return;
        return run(() => {
          const position = paragraphBoundary(document, range, command.includes('next'));
          if (position) moveSelection(position, command.startsWith('select') ? range.anchor : position);
        });
      case 'next-cell':
      case 'prev-cell':
      case 'row-start':
      case 'row-end':
      case 'column-start':
      case 'column-end': {
        const result = range && tableShortcut(document, range, command);
        if (!result) return;
        if (result.selection) return run(() => moveSelection(result.selection.focus, result.selection.anchor));
        if (!disabled && result.transaction) return run(() => apply(result.transaction));
        return;
      }
    }
    if (disabled || !range) return;
    const currentFormat = { ...activeFormat(document, range).format, ...typingRef.current };
    switch (command) {
      case 'undo': return run(() => travel('undo'));
      case 'redo': return run(() => travel('redo'));
      case 'bold':
      case 'italic':
      case 'underline': return run(() => setTextFormat({ [command]: !currentFormat[command] }, range));
      case 'clear-text': return run(clearFormatting);
      case 'copy-format': return run(() => { copiedFormat.current = currentFormat; });
      case 'paste-format': return run(() => { if (copiedFormat.current) setTextFormat(copiedFormat.current, range); });
      case 'grow':
      case 'shrink':
      case 'plus-point':
      case 'minus-point': return run(() => {
        const current = currentFormat.size ?? 11;
        const sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72, 80, 96, 120];
        const size = command === 'grow' ? sizes.find(value => value > current) ?? current + 10 :
          command === 'shrink' ? sizes.reverse().find(value => value < current) ?? current - 1 :
          current + (command === 'plus-point' ? 1 : -1);
        setTextFormat({ size: Math.max(1, Math.min(400, size)) }, range);
      });
      case 'super':
      case 'sub': return run(() => {
        const vertical = command === 'super' ? 'superscript' : 'subscript';
        setTextFormat({ vertical: currentFormat.vertical === vertical ? 'baseline' : vertical }, range);
      });
      case 'font': return run(() => focusRibbonControl('Font family'));
      case 'font-size': return run(() => focusRibbonControl('Font size'));
      case 'styles': return run(() => focusRibbonControl('Paragraph style'));
      case 'case': return run(() => apply(changeCase(document, range), range));
      case 'left':
      case 'right':
      case 'center':
      case 'justify': return run(() => alignParagraph(command));
      case 'indent':
      case 'outdent':
      case 'hang':
      case 'unhang': return run(() => changeIndent(command === 'indent' || command === 'hang' ? 1 : -1,
        command === 'hang' || command === 'unhang'));
      case 'clear-paragraph': return run(() => withParagraphs(item => ({ ...item, style: {} })));
      case 'single':
      case 'double':
      case 'one-half': return run(() => withParagraphs(item => ({ ...item, style: { ...item.style,
        lineHeight: command === 'single' ? 1 : command === 'double' ? 2 : 1.5 } })));
      case 'normal': return run(() => applyStyle(document.defaultStyleId));
      case 'heading-1':
      case 'heading-2':
      case 'heading-3': return run(() => {
        const level = Number(command.at(-1));
        applyStyle(styles.find(style => style.heading === level)?.id ?? 'Heading' + level);
      });
      case 'bullets': return run(() => setList('bullet'));
      case 'link': return run(() => openCommandDialog('link'));
      case 'comment': if (!collapsed(range)) return run(() => openCommandDialog('comment')); return;
      case 'track': return run(toggleTracking);
      case 'page-break': return run(() => { if (!regionRef.current) insertInline({ kind: 'pageBreak' }); });
      case 'page-field': return run(() => addPageField('PAGE'));
      case 'select-table': {
        const table = tableAt(document, range)?.table;
        if (table) return run(() => chooseTable(table));
        return;
      }
      case 'cell-tab': if (!tableAt(document, range)) return; break;
    }
    const symbols: Partial<Record<Shortcut, string>> = {
      nbsp: '\u00a0', nbhyphen: '‑', 'soft-hyphen': '\u00ad', emdash: '—', endash: '–', ellipsis: '…',
      registered: '®', trademark: '™', 'cell-tab': '\t',
    };
    const symbol = symbols[command];
    if (symbol) run(() => insertText(range, symbol));
  }

  const styles = model.styles.filter(style => style.type === 'paragraph');
  for (const style of DEFAULT_STYLES) {
    if (style.id !== 'Normal' && !styles.some(candidate => candidate.id === style.id)) styles.push(style);
  }
  const displayZoom = Math.round(scale * 100);
  const readOnlyMode = readOnly || !!session.readOnlyReason || phone;
  const label = session.fileName.replace(/\.docx$/i, '');
  const format = { ...active?.format, ...typingFormat };

  return <section ref={editorRoot} onKeyDown={handleShortcut} onPointerDown={() => {
    if (keyTips) setKeyTips(undefined);
  }} className={'rwe-editor' + (className ? ' ' + className : '')} aria-label="Word document editor">
    <input ref={fileInput} hidden type="file" accept={convertDocument ? '.doc,.docx,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document' : '.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document'}
      onChange={event => {
        const file = event.currentTarget.files?.[0];
        event.currentTarget.value = '';
        if (!file) return;
        const original = sessionRef.current.model;
        const selectedForOpen = approvedFileSelection.current;
        approvedFileSelection.current = null;
        const load = () => { void open(file, { discardChanges: true }).catch(() => {}); };
        if (selectedForOpen === original) load();
        else requestReplacement(load);
      }} />
    <input ref={imageInput} hidden type="file" accept="image/png,image/jpeg" onChange={event => {
      const file = event.currentTarget.files?.[0];
      event.currentTarget.value = '';
      const pending = imageSelection.current;
      imageSelection.current = null;
      if (file && pending && pending.model === sessionRef.current.model) void attachImage(file, pending.selection, pending.model);
    }} />
    <header className="rwe-titlebar">
      <div className="rwe-document-icon"><FileText size={17} aria-hidden="true" /></div>
      <div className="rwe-title-text">
        <span className="rwe-document-title" title={session.fileName}>{label}</span>
        <span className="rwe-local-state"><span data-dirty={dirty} />{loading ? 'Opening…' : exporting ? 'Preparing download…' :
          readOnlyMode ? 'Read only' : dirty ? 'Unsaved changes' : 'Local document'}</span>
      </div>
      <div className="rwe-title-actions">
        <button type="button" className="rwe-open" disabled={loading} aria-label="Open document" title="Open document"
          onClick={chooseDocument}><FolderOpen size={16} /> <span>Open</span></button>
        <button type="button" className="rwe-primary" disabled={loading || exporting} onClick={() => void download()}>
          <Download size={15} /> <span>Download</span></button>
      </div>
    </header>
    <div className="rwe-tabs" role="tablist" aria-label="Editor ribbon" onKeyDown={event => {
      const index = availableTabs.indexOf(visibleTab);
      const size = availableTabs.length;
      const next = event.key === 'ArrowRight' ? availableTabs[(index + 1) % size] :
        event.key === 'ArrowLeft' ? availableTabs[(index + size - 1) % size] : undefined;
      if (next) {
        event.preventDefault();
        setTab(next);
        event.currentTarget.querySelectorAll('button')[availableTabs.indexOf(next)]?.focus();
      }
    }}>
      {availableTabs.map(value => <button key={value} type="button" role="tab" aria-selected={value === visibleTab}
        aria-controls={ribbonId} tabIndex={value === visibleTab ? 0 : -1} onClick={() => { setKeyTips(undefined); setTab(value); setRibbonCollapsed(false); }}>{value}
        {keyTips && !keyTips.tab && <span className="rwe-keytip" aria-hidden="true">{tabTips[value]}</span>}</button>)}
    </div>
    <div id={ribbonId} className="rwe-ribbon-container" hidden={ribbonCollapsed}>
      {keyTips?.tab && <div className="rwe-keytip-commands" aria-label={keyTips.tab + ' access keys'}>
        {Object.entries(commandTips[keyTips.tab]).filter(([tip]) => tip.startsWith(keyTips.prefix)).map(([tip, label]) =>
          <span key={tip}><kbd>{tip}</kbd> {label}</span>)}
      </div>}
      <Ribbon tab={visibleTab} busy={loading || exporting} disabled={disabled} styles={styles}
        styleId={active?.item?.styleId ?? model.defaultStyleId} format={format} list={active?.item?.list?.kind}
        alignment={active?.item?.style.alignment ?? styles.find(style => style.id === active?.item?.styleId)?.paragraph.alignment}
        canUndo={history.current.undo.length > 0} canRedo={history.current.redo.length > 0}
        outline={outline} zoom={displayZoom} canShowOutline={canShowOutline} undo={() => travel('undo')} redo={() => travel('redo')}
        paste={() => void pasteFromClipboard()} pasteMode={pasteMode} setPasteMode={setPasteMode}
        changeIndent={changeIndent} clearFormatting={clearFormatting}
        tracking={model.tracking} canComment={!!selection && !collapsed(selection)} showReview={showReview}
        toggleReview={() => setShowReview(value => !value)} toggleTracking={toggleTracking}
        formatText={setTextFormat} setStyle={applyStyle} align={alignParagraph}
        toggleList={setList} onNew={() => requestReplacement(() => resetDocument(newSession()))}
        open={chooseDocument} download={() => void download()} downloadDoc={convertDocument ? () => void download('doc') : undefined} print={printDocument}
        dialog={openCommandDialog} pageRegion={editPageRegion} inPageRegion={!!regionEdit}
        image={() => {
          const selected = currentSelection();
          if (!selected) return;
          imageSelection.current = { selection: selected, model: sessionRef.current.model };
          imageInput.current?.click();
        }}
        pageBreak={() => { if (!regionEdit) insertInline({ kind: 'pageBreak' }); }} toggleOutline={() => setOutline(value => !value)} fit={fit} setZoom={setZoom}>
        {activeTable && !disabled && !activeImage && <div className="rwe-contextbar rwe-table-tools" role="group" aria-label="Table tools">
          <strong>{selectedTable ? 'Table selected' : 'Table'}</strong>
          <button type="button" aria-pressed={!!selectedTable} onMouseDown={event => event.preventDefault()}
            onClick={() => chooseTable(activeTable.table)}>Select table</button>
          <button type="button" disabled={activeTable.table.structuralLock} onMouseDown={event => event.preventDefault()} onClick={() => openCommandDialog('table-format')}>Format table</button>
          <button type="button" onMouseDown={event => event.preventDefault()} onClick={() => performTableAction('delete-table')}>Delete table</button>
          <select aria-label="Table action" disabled={activeTable.table.structuralLock} value="" onChange={event => {
            const action = event.target.value as TableAction;
            if (action) performTableAction(action);
          }}>
            <option value="">Rows and columns</option>
            <option value="row-before">Insert row above</option>
            <option value="row-after">Insert row below</option>
            <option value="column-before">Insert column left</option>
            <option value="column-after">Insert column right</option>
            <option value="delete-row">Delete row</option>
            <option value="delete-column">Delete column</option>
            <option value="delete-table">Delete table</option>
          </select>
          <label>Cell shading<input type="color" value={activeTable.table.rows[activeTable.rowIndex].cells[activeTable.columnIndex].fill ?? '#ffffff'}
            onChange={event => {
              const target = currentSelection();
              if (target) apply(shadeCell(sessionRef.current.model, target, event.target.value), target);
            }} /></label>
          <button type="button" onClick={() => {
            const target = currentSelection();
            if (target) apply(shadeCell(sessionRef.current.model, target), target);
          }}>Clear shading</button>
        </div>}
      </Ribbon>
    </div>
    {recovery.offer && <div className="rwe-notice" role="status">A local recovery copy of {recovery.offer.file.name} is available.
      <button type="button" disabled={loading} onClick={() => requestReplacement(() => {
        void open(recovery.offer!.file, { discardChanges: true }).then(() => {
          updateSession({ ...sessionRef.current, saved: blankModel() }); recovery.restored();
        }).catch(() => {});
      })}>Restore draft</button>
      <button type="button" onClick={() => void recovery.discard().catch(() => {})}>Discard recovery copy</button>
    </div>}
    {recoveryKey && !recovery.offer && recovery.status !== 'idle' && <div className="rwe-recovery-state" role="status">
      {recovery.status === 'saved' ? 'Recovery copy saved on this device' : recovery.status === 'saving' ? 'Saving recovery copy…' : 'Recovery unavailable — download your document'}
    </div>}
    {pasteNotice && <div className="rwe-notice" role="status">{pasteNotice}
      <button type="button" aria-label="Dismiss paste tip" onClick={() => setPasteNotice('')}><X size={15} /></button></div>}
    {model.tracking && !disabled && <div className="rwe-notice" role="status">Track Changes is on for text insertions and deletions.</div>}
    {regionEdit && <div className="rwe-contextbar" role="group" aria-label="Header and footer tools">
      <strong>{regionEdit.key.startsWith('header') ? 'Header' : 'Footer'} · Page {regionEdit.page}</strong>
      <button type="button" onMouseDown={event => event.preventDefault()} onClick={() => addPageField('PAGE')}>Page number</button>
      <button type="button" onMouseDown={event => event.preventDefault()} onClick={() => addPageField('NUMPAGES')}>Total pages</button>
      <label><input type="checkbox" checked={model.regions.first} onChange={event => changeRegionSetting('first', event.target.checked)} /> Different first page</label>
      <label><input type="checkbox" checked={model.regions.even} onChange={event => changeRegionSetting('even', event.target.checked)} /> Different odd and even pages</label>
      <button type="button" onClick={removePageRegion}>Remove {regionEdit.key.startsWith('header') ? 'header' : 'footer'}</button>
      <button type="button" onClick={closePageRegion}>Close header and footer</button>
    </div>}
    {activeImage && !disabled && <div className="rwe-contextbar" role="group" aria-label="Picture tools">
      <strong>Picture</strong>
      <button type="button" onClick={() => setDialog('picture-size')}>Size and alternative text</button>
      <button type="button" onClick={() => {
        const start = activeImage.selection;
        const target = { anchor: start.anchor, focus: { ...start.focus, offset: start.focus.offset + 1 } };
        apply(replaceRange(model, target, [[]]), start);
        setSelectedImage(undefined);
      }}>Remove picture</button>
    </div>}
    {error && <div className="rwe-notice rwe-error" role="alert">{error.message}
      <button type="button" aria-label="Dismiss error" onClick={() => setError(undefined)}><X size={16} /></button></div>}
    {(session.readOnlyReason || session.warnings.length > 0) && <div className="rwe-notice" role="status">
      {[session.readOnlyReason, ...session.warnings].filter(Boolean).join(' ')}
    </div>}
    <div className="rwe-workspace" data-outline={outline}>
      {outline && <aside className="rwe-outline" aria-label="Document outline">
        <div className="rwe-outline-heading"><PanelLeft size={15} /> <span>Outline</span></div>
        <nav aria-label="Headings">
          {stats.headings.length ? stats.headings.map(heading => <button type="button" key={heading.id}
            style={{ paddingInlineStart: 9 + Math.min(heading.level - 1, 4) * 12 }}
            className={selection?.focus.id === heading.id ? 'rwe-current-heading' : undefined}
            onClick={() => gotoHeading(heading.id)}><ChevronRight size={12} /><span>{heading.text}</span></button>)
            : <p>Headings will appear here.</p>}
        </nav>
      </aside>}
      <div className="rwe-viewport" ref={viewport} onScroll={scrollStatus}>
        <div className="rwe-pages" data-narrow={viewportWidth < 620} ref={editingRoot} contentEditable={!disabled} suppressContentEditableWarning
          style={{ '--rwe-page-scale': scale } as CSSProperties}
          role="textbox" aria-label="Document content" aria-multiline="true" aria-readonly={disabled} spellCheck
          onPointerDown={event => {
            if (event.ctrlKey || event.metaKey) return;
            const previous = selectionRef.current;
            if (previous?.tableId) updateSelection({ anchor: previous.anchor, focus: previous.focus });
            resetTyping();
            history.current.lastTyping = undefined;
          }}
          onClick={event => {
            const target = event.target as HTMLElement;
            const link = target.closest<HTMLElement>('[data-href]');
            if (regionRef.current && !target.closest('[data-page-region]')) {
              updateRegion(null);
            }
            const picture = target.closest<HTMLElement>('[data-image]');
            if (picture && !disabled && !picture.closest('[data-region-editing="false"]') && !picture.closest('[data-locked]')) {
              const selected = imageAt(model, picture.dataset.image!);
              if (selected) {
                focusSelection(selected.selection);
                setSelectedImage(selected.image.id);
              }
            } else {
              setSelectedImage(undefined);
            }
            if (link?.dataset.href && (event.ctrlKey || event.metaKey || disabled)) window.open(link.dataset.href, '_blank', 'noopener,noreferrer');
          }}
          onInput={event => {
            if (!composing.current && nativeEdit.current) reconcileNative((event.nativeEvent as InputEvent).data);
          }}
          onCompositionStart={() => {
            const range = currentSelection();
            if (range && editingRoot.current && !disabled) {
              nativeEdit.current = snapshotSelection(editingRoot.current, range);
              composing.current = !!nativeEdit.current;
            }
          }}
          onCompositionEnd={event => {
            if (!composing.current) return;
            const text = event.data;
            queueMicrotask(() => {
              composing.current = false;
              reconcileNative(text);
            });
          }}
          onCopy={event => {
            const range = currentSelection();
            if (range && !collapsed(range)) {
              event.preventDefault();
              if (selectionHasImages(sessionRef.current.model, range)) {
                transferPictures(range, false, event.clipboardData); return;
              }
              writeClipboard(event.clipboardData, sessionRef.current.model, range);
            }
          }}
          onCut={event => {
            event.preventDefault();
            const range = currentSelection();
            if (!range || disabled || collapsed(range)) return;
            const model = sessionRef.current.model;
            if (selectionHasImages(model, range)) {
              transferPictures(range, true, event.clipboardData);
              return;
            }
            if (wholeTableSelection(model, range)) {
              writeClipboard(event.clipboardData, model, range);
              deleteSelectedTable(range);
              return;
            }
            const transaction = model.tracking ? trackEdit(model, range, [], author) : replaceRange(model, range, [[]]);
            if (transaction) {
              writeClipboard(event.clipboardData, model, range);
              apply(transaction, range);
            }
          }}
          onPaste={event => {
            event.preventDefault();
            setPasteNotice('');
            const range = currentSelection();
            const mode = plainPaste.current ? 'text' : pasteMode;
            plainPaste.current = false;
            if (range && !disabled) void pasteContent(range, event.clipboardData.getData('text/html'),
              event.clipboardData.getData('text/plain'), mode,
              Array.from(event.clipboardData.files).filter(file => file.type === 'image/png' || file.type === 'image/jpeg'));
          }}
          onDrop={event => event.preventDefault()}
          onKeyDown={event => {
            if (selectedImage && !event.nativeEvent.isComposing) setSelectedImage(undefined);
          }}>
          <SelectedPicture.Provider value={selectedImage}>{pages.map((page, index) => <div className="rwe-sheet-slot" key={index}
            style={{ width: model.page.width * scale, minHeight: model.page.height * scale }}>
            <PageNumbers.Provider value={{ page: index + 1, total: pages.length }}>
              <div className="rwe-sheet" data-page={index + 1} style={{ width: model.page.width, minHeight: model.page.height, zoom: scale,
                padding: [layout?.top ?? model.page.top, model.page.right, layout?.bottom ?? model.page.bottom, model.page.left].map(value => value + 'px').join(' ') }}>
                {(['header', 'footer'] as const).map(kind => {
                  const key = regionForPage(model.regions, kind, index + 1);
                  const fragments = regionFragments[key];
                  const editing = !disabled && regionEdit?.key === key && regionEdit.page === index + 1;
                  const edge = kind === 'header' ? { top: model.page.header } : { bottom: model.page.footer };
                  return fragments?.length ? <div key={kind} className="rwe-page-region" data-page-region={kind}
                    data-region-editing={editing} contentEditable={editing ? undefined : false}
                    aria-label={kind === 'header' ? 'Page header' : 'Page footer'}
                    style={{ left: model.page.left, right: model.page.right, ...edge }}
                    onDoubleClick={() => { if (!editing) editPageRegion(kind, index + 1); }}>
                    {fragments.map(fragment => <RenderFragment model={model} fragment={fragment} key={fragmentKey(fragment)}
                      selectTable={editing ? chooseTable : undefined} selectedTable={selectedTable?.id} />)}
                  </div> : null;
                })}
                {page.map(fragment => <RenderFragment model={model} fragment={fragment} key={fragmentKey(fragment)}
                  selectTable={!disabled ? chooseTable : undefined} selectedTable={selectedTable?.id} />)}
              </div>
            </PageNumbers.Provider>
          </div>)}</SelectedPicture.Provider>
        </div>
        <div ref={measurement} className="rwe-measure" aria-hidden="true"
          style={{ width: model.page.width - model.page.left - model.page.right, '--rwe-page-scale': scale } as CSSProperties}>
          {measures.map(measure => <div key={measure.key} data-measure={measure.key}><RenderFragment model={model} fragment={measure.fragment} /></div>)}
          {Object.entries(regionFragments).map(([key, fragments]) => <div key={key} data-measure-region={key}>
            {fragments.map(fragment => <RenderFragment model={model} fragment={fragment} key={fragmentKey(fragment)} />)}
          </div>)}
        </div>
      </div>
      {showReview && <aside className="rwe-review-pane" aria-label="Review pane">
        <header><strong>Review</strong><button type="button" aria-label="Close review pane" onClick={() => setShowReview(false)}><X size={15} /></button></header>
        <h2>Comments</h2>
        {model.comments.length ? model.comments.map(comment => <article key={comment.id}>
          <div><strong>{comment.author}</strong></div>
          <p>{comment.text}</p>
          <div className="rwe-review-actions">
            <button type="button" onClick={() => {
              const item = stats.items.find(item => item.comments?.some(mark => mark.id === comment.id));
              if (item) gotoHeading(item.id);
            }}>Go to comment</button>
            <button type="button" disabled={disabled} onClick={() => apply(deleteComment(model, comment.id))}>Delete comment</button>
          </div>
        </article>) : <p>No comments</p>}
        <h2>Tracked changes</h2>
        {reviewChanges.length ? reviewChanges.map(change => <article key={change.revision.id}>
          <div><strong>{change.revision.author}</strong> · {change.revision.kind === 'ins' ? 'Insertion' : 'Deletion'}</div>
          <p>{change.text}</p>
          <div className="rwe-review-actions">
            <button type="button" disabled={disabled} onClick={() => apply(resolveRevision(model, change.revision.id, true))}>Accept</button>
            <button type="button" disabled={disabled} onClick={() => apply(resolveRevision(model, change.revision.id, false))}>Reject</button>
          </div>
        </article>) : <p>No tracked changes</p>}
      </aside>}
    </div>
    <footer className="rwe-status">
      <div className="rwe-status-info">
        <button type="button" aria-label="Show or hide outline" aria-pressed={outline} title="Document outline"
          disabled={!canShowOutline} onClick={() => setOutline(value => !value)}><AlignLeft size={14} /></button>
        <span>Page {Math.min(pageNumber, pages.length)} of {pages.length}</span>
        <span className="rwe-status-divider" />
        <span>{stats.words} {stats.words === 1 ? 'word' : 'words'}</span>
      </div>
      <div className="rwe-zoom">
        <button type="button" aria-label="Zoom out" onClick={() => setZoom(Math.max(25, displayZoom - 10))}><Minus size={15} /></button>
        <input aria-label="Zoom" type="range" min={25} max={150} value={Math.max(25, displayZoom)} onChange={event => setZoom(Number(event.target.value))} />
        <button type="button" aria-label="Zoom in" onClick={() => setZoom(Math.min(150, displayZoom + 10))}><Plus size={15} /></button>
        <output>{displayZoom}%</output>
      </div>
    </footer>
    {dialog && <dialog ref={dialogRef} className="rwe-dialog" data-shortcut-guide={dialog === 'shortcuts'}
      data-template-dialog={dialog === 'templates'}
      onClose={event => { if (!event.currentTarget.open) setDialog(undefined); }} aria-labelledby={ribbonId + '-dialog'}>
      {dialog === 'templates' ? <>
        <div className="rwe-template-heading"><h2 id={ribbonId + '-dialog'}>Document templates</h2>
          <button type="button" aria-label="Close templates" onClick={closeDialog}><X size={16} /></button></div>
        <TemplateDialog fileName={session.fileName} exportDocument={exportCurrent} create={createFromTemplate} allowSave={!loading && !session.readOnlyReason} />
      </> : <form onSubmit={event => {
        event.preventDefault();
        const data = new FormData(event.currentTarget);
        const target = selectedForDialog.current;
        if (dialog === 'discard') {
          closeDialog();
          const action = discard.current;
          discard.current = null;
          action?.();
        } else if (dialog === 'link' && target) {
          const href = safeHref(String(data.get('url') ?? '').trim());
          if (!href) {
            const input = event.currentTarget.elements.namedItem('url') as HTMLInputElement;
            input.setCustomValidity('Enter an http, https, or mailto link.');
            input.reportValidity();
            return;
          }
          closeDialog();
          const document = sessionRef.current.model;
          if (collapsed(target)) {
            const text = String(data.get('text') ?? '').trim() || href;
            const run = { kind: 'text' as const, text, format: activeFormat(document, target).direct, link: { href } };
            insertInline(run, target);
          } else {
            apply(formatRange(document, target, run => ({ ...run, link: { href } })), target);
          }
        } else if ((dialog === 'modify-style' || dialog === 'new-style') && target && styleForDialog.current) {
          const base = styleForDialog.current;
          const create = dialog === 'new-style';
          const style = readStyleFields(data, base, create ? nodeId('CustomStyle') : base.id, create);
          if (!style.name || !style.format.font) return;
          const current = sessionRef.current.model;
          const updated = { ...current, styles: updateNamedStyle(current.styles, style) };
          closeDialog();
          if (create) apply(updateParagraphs(updated, target, item => ({ ...item, styleId: style.id })), target);
          else apply({ model: updated, selection: target }, target);
        } else if (dialog === 'page-setup' && target) {
          const px = (name: string) => Number(data.get(name)) * 96;
          const page = { width: px('page-width'), height: px('page-height'), left: px('page-left'), right: px('page-right'),
            top: px('page-top'), bottom: px('page-bottom'), header: px('page-header'), footer: px('page-footer') };
          const size = event.currentTarget.elements.namedItem('page-width') as HTMLInputElement;
          if (page.left + page.right >= page.width - 48 || page.top + page.bottom >= page.height - 48 ||
            page.header > page.height / 2 || page.footer > page.height / 2) {
            size.setCustomValidity('Leave at least half an inch for the body and place headers and footers inside the page.');
            size.reportValidity();
            return;
          }
          closeDialog();
          apply({ model: { ...model, page }, selection: target }, target);
        } else if (dialog === 'template-field' && target) {
          const name = String(data.get('field') ?? '').trim();
          if (!/^[a-z][a-z0-9_.-]{0,79}$/i.test(name)) return;
          closeDialog();
          insertText(target, '{{' + name + '}}');
        } else if (dialog === 'shortcuts') {
          closeDialog();
        } else if (dialog === 'goto') {
          closeDialog();
          goToPage(Number(data.get('page')));
        } else if (dialog === 'replace') {
          const query = String(data.get('query') ?? '');
          const input = event.currentTarget.elements.namedItem('query') as HTMLInputElement;
          const matches = textMatches(model, query);
          const all = (event.nativeEvent as SubmitEvent).submitter?.getAttribute('value') === 'all';
          const result = model.tracking ? { transaction: null } :
            replaceMatches(model, all ? matches : matches.slice(0, 1), String(data.get('replacement') ?? ''));
          if (result.transaction) {
            closeDialog();
            apply(result.transaction);
          } else {
            input.setCustomValidity(model.tracking ? 'Turn off Track Changes to use Replace.' : 'No editable matches found.');
            input.reportValidity();
          }
        } else if (dialog === 'find') {
          const input = event.currentTarget.elements.namedItem('query') as HTMLInputElement;
          if (!findText(input.value)) {
            input.setCustomValidity('No matching text found.');
            input.reportValidity();
          } else {
            closeDialog();
            const found = selectionRef.current;
            if (found && !disabled) focusSelection(found);
          }
        } else if (dialog === 'comment' && target) {
          const text = String(data.get('comment') ?? '').trim();
          const author = String(data.get('author') ?? '').trim();
          if (!text || !author) return;
          closeDialog();
          if (apply(createComment(model, target, text, author), target)) setShowReview(true);
          else report(new DocxEditorError('PROTECTED_CONTENT', 'Select text within a single body paragraph to add a comment.'));
        } else if (dialog === 'picture-size' && activeImage) {
          const width = Number(data.get('picture-width'));
          const height = data.has('picture-lock') ? width / activeImage.image.width * activeImage.image.height : Number(data.get('picture-height'));
          if (height < 1 || height > 4000) {
            const field = event.currentTarget.elements.namedItem('picture-width') as HTMLInputElement;
            field.setCustomValidity('The proportional height must be between 1 and 4,000 pixels.');
            field.reportValidity();
            return;
          }
          closeDialog();
          updatePicture(width, height, String(data.get('picture-alt') ?? ''));
        } else if (dialog === 'table-format' && target) {
          const document = sessionRef.current.model;
          const context = tableAt(document, target);
          closeDialog();
          if (context) {
            const width = document.page.width - document.page.left - document.page.right;
            const settings = readTableFormatting(data, context.table, context.rowIndex, context.columnIndex, width);
            if (!apply(formatTable(document, target, settings), target)) focusSelection(target);
          }
        } else if (dialog === 'table-delete' && target && pendingTableAction.current) {
          closeDialog();
          if (!apply(editTable(sessionRef.current.model, target, pendingTableAction.current), target)) {
            report(new DocxEditorError('PROTECTED_CONTENT', 'This table contains protected content.'));
          }
        } else if (dialog === 'table' && target) {
          closeDialog();
          const transaction = insertTable(sessionRef.current.model, target, Number(data.get('rows')), Number(data.get('columns')));
          if (!apply(transaction, target)) report(new DocxEditorError('PROTECTED_CONTENT', 'Place the cursor in a paragraph outside a table.'));
        }
      }}>
        <h2 id={ribbonId + '-dialog'}>{{ discard: 'Discard unsaved changes?', link: 'Insert link', table: 'Insert table',
          'modify-style': 'Modify style', 'new-style': 'New style', 'picture-size': 'Picture size and alternative text',
          'table-delete': pendingTableAction.current === 'delete-table' ? 'Delete entire table?' : 'Delete table content?', comment: 'New comment', 'page-setup': 'Page setup', find: 'Find in document', shortcuts: 'Keyboard shortcuts',
          replace: 'Find and replace', goto: 'Go to page', 'template-field': 'Insert template field', 'table-format': 'Format table' }[dialog]}</h2>
        {dialog === 'discard' && <p>Your changes will be lost when you replace this document.</p>}
        {dialog === 'link' && <>
          {(!selectedForDialog.current || collapsed(selectedForDialog.current)) && <label>Display text<input name="text" autoFocus /></label>}
          <label>Address<input name="url" type="url" placeholder="https://example.com" required
            onChange={event => event.currentTarget.setCustomValidity('')} /></label>
        </>}
        {(dialog === 'new-style' || dialog === 'modify-style') && styleForDialog.current &&
          <StyleFields style={dialog === 'new-style' ? { ...styleForDialog.current, name: '' } : styleForDialog.current} />}
        {dialog === 'page-setup' && <div className="rwe-dialog-grid" onChange={event => {
          const input = event.currentTarget.querySelector<HTMLInputElement>('[name="page-width"]');
          input?.setCustomValidity('');
        }}>
          {([['width', 'Width', 2], ['height', 'Height', 2], ['top', 'Top margin', 0], ['bottom', 'Bottom margin', 0],
            ['left', 'Left margin', 0], ['right', 'Right margin', 0], ['header', 'Header from top', 0], ['footer', 'Footer from bottom', 0]] as const)
            .map(([key, label, min]) => <label key={key}>{label} (in)<input type="number" min={min} max="25" step="any"
              name={'page-' + key} defaultValue={Number((model.page[key] / 96).toFixed(3))} required /></label>)}
        </div>}
        {dialog === 'shortcuts' && <ShortcutGuide mac={mac} />}
        {dialog === 'template-field' && <label>Field name<input name="field" autoFocus required maxLength={80}
          pattern={'[a-zA-Z][a-zA-Z0-9_.\\-]{0,79}'} placeholder="client_name" /></label>}
        {dialog === 'goto' && <label>Page number<input name="page" type="number" min={1} max={pages.length} required autoFocus /></label>}
        {(dialog === 'find' || dialog === 'replace') && <label>Search text<input name="query" required autoFocus
          defaultValue={lastFindQuery.current} onChange={event => event.target.setCustomValidity('')} /></label>}
        {dialog === 'replace' && <label>Replace with<input name="replacement" /></label>}
        {dialog === 'comment' && <>
          <label>Author<input name="author" defaultValue={author} maxLength={80} required /></label>
          <label>Comment<textarea name="comment" rows={4} maxLength={2000} required autoFocus /></label>
        </>}
        {dialog === 'table-format' && activeTable && <TableFormatFields table={activeTable.table}
          row={activeTable.rowIndex} column={activeTable.columnIndex} whole={!!selectedTable}
          availableWidth={model.page.width - model.page.left - model.page.right} />}
        {dialog === 'table-delete' && <p>{pendingTableAction.current === 'delete-table' ?
          'The table and all its cells will be removed.' : 'Text in the selected rows or columns will be removed.'}</p>}
        {dialog === 'picture-size' && activeImage && <>
          <div className="rwe-dialog-grid">
            <label>Width (px)<input onChange={event => event.target.setCustomValidity('')} type="number" name="picture-width" min="1" max="4000" step="any" defaultValue={activeImage.image.width} required /></label>
            <label>Height (px)<input type="number" name="picture-height" min="1" max="4000" step="any" defaultValue={activeImage.image.height} required /></label>
          </div>
          <div className="rwe-style-toggles"><label><input name="picture-lock" type="checkbox" defaultChecked onChange={event =>
            (event.target.form?.elements.namedItem('picture-width') as HTMLInputElement | undefined)?.setCustomValidity('')} /> Keep proportions</label></div>
          <label>Alternative text<input name="picture-alt" defaultValue={activeImage.image.alt} maxLength={512} /></label>
        </>}
        {dialog === 'table' && <div className="rwe-dialog-grid">
          <label>Rows<input name="rows" type="number" min={1} max={20} defaultValue={3} autoFocus required /></label>
          <label>Columns<input name="columns" type="number" min={1} max={8} defaultValue={3} required /></label>
        </div>}
        <div className="rwe-dialog-actions"><button type="button" onClick={closeDialog}>{dialog === 'shortcuts' ? 'Close' : 'Cancel'}</button>
          {dialog === 'replace' && <button type="submit" value="all" className="rwe-primary">Replace all</button>}
          {dialog !== 'shortcuts' &&
          <button type="submit" className="rwe-primary">{dialog === 'discard' ? 'Discard changes' : dialog === 'new-style' || dialog === 'modify-style' ? 'Save style' : dialog === 'picture-size' ? 'Apply' : dialog === 'table-delete' ? 'Delete' : dialog === 'comment' ? 'Add comment' : dialog === 'page-setup' || dialog === 'table-format' ? 'Apply' : dialog === 'find' ? 'Find' : dialog === 'replace' ? 'Replace' : dialog === 'goto' ? 'Go' : 'Insert'}</button>}</div>
      </form>}
    </dialog>}
  </section>;
});

async function decodePicture(file: Blob, snapshot: DocModel, alt: string): Promise<ImageRun> {
  const name = file instanceof File ? file.name : '';
  const mime = /\.png$/i.test(name) ? 'image/png' : /\.jpe?g$/i.test(name) ? 'image/jpeg' : file.type;
  const header = new Uint8Array(await file.slice(0, 8).arrayBuffer());
  const png = mime === 'image/png' && [137, 80, 78, 71, 13, 10, 26, 10].every((value, index) => header[index] === value);
  const jpeg = mime === 'image/jpeg' && header[0] === 255 && header[1] === 216 && header[2] === 255;
  if ((!png && !jpeg) || file.size > 12 * 1024 * 1024) {
    throw new DocxEditorError('INVALID_IMAGE', 'Choose a PNG or JPEG smaller than 12 MB.');
  }
  const blob = file.slice(0, file.size, mime);
  const bitmap = await createImageBitmap(blob);
  const factor = Math.min(1, (snapshot.page.width - snapshot.page.left - snapshot.page.right) / bitmap.width,
    (snapshot.page.height - snapshot.page.top - snapshot.page.bottom - 30) / bitmap.height);
  const image: ImageRun = { kind: 'image', id: nodeId('image'), blob,
    width: Math.max(1, bitmap.width * factor), height: Math.max(1, bitmap.height * factor), alt };
  bitmap.close();
  return image;
}

function sameSelectionEndpoints(a: ModelSelection, b: ModelSelection) {
  return a.anchor.id === b.anchor.id && a.anchor.offset === b.anchor.offset &&
    a.focus.id === b.focus.id && a.focus.offset === b.focus.offset;
}

function fragmentKey(fragment: Fragment) {
  if (fragment.kind === 'paragraph') return fragment.paragraph.id + ':' + fragment.from;
  if (fragment.kind === 'table') return fragment.table.id + ':' + fragment.from;
  return fragment.id;
}

function focusCaretIntoView(selection: Selection, viewport: HTMLDivElement) {
  if (!selection.focusNode) return;
  const range = viewport.ownerDocument.createRange();
  range.setStart(selection.focusNode, selection.focusOffset);
  range.collapse(true);
  const bounds = range.getBoundingClientRect();
  if (!bounds.height) return;
  const pane = viewport.getBoundingClientRect();
  if (bounds.bottom > pane.bottom - 22) viewport.scrollTop += bounds.bottom - pane.bottom + 22;
  else if (bounds.top < pane.top + 16) viewport.scrollTop -= pane.top + 16 - bounds.top;
}
