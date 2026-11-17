import {
  ALargeSmall, ClipboardPaste, Eraser, Highlighter, IndentDecrease, IndentIncrease, Keyboard, Files, Braces, Printer, Replace, Search, Strikethrough, Subscript, Superscript, AlignCenter, AlignJustify, AlignLeft, AlignRight, Bold, ChevronDown, Columns3, Download, FilePlus2,
  FolderOpen, Heading, ImagePlus, Italic, Link2, List, ListOrdered, PanelBottom, PanelLeft, PanelTop, Redo2, Scan, SeparatorHorizontal,
  MessageSquarePlus, FilePenLine, MessagesSquare, SlidersHorizontal, Table2, Underline, Undo2,
} from 'lucide-react';
import { useState } from 'react';
import type { ReactNode } from 'react';
import type { Alignment, NamedStyle, TextFormat, RegionKind } from './model';
import type { PasteMode } from './clipboard';
import { FONTS } from './style-dialog';
import { TextColor } from './text-color';


export type Tab = 'File' | 'Home' | 'Insert' | 'Layout' | 'Review' | 'View' | 'Format Table';
export type DialogName = 'link' | 'table' | 'modify-style' | 'new-style' | 'comment' | 'page-setup' | 'find' | 'shortcuts' | 'templates' | 'template-field' | 'replace' | 'goto';
type RibbonProps = {
  tab: Tab;
  busy: boolean;
  disabled: boolean;
  styles: NamedStyle[];
  styleId?: string;
  format: TextFormat;
  alignment?: Alignment;
  list?: 'bullet' | 'number';
  canUndo: boolean;
  canRedo: boolean;
  outline: boolean;
  canShowOutline: boolean;
  pageRegion: (kind: RegionKind) => void;
  inPageRegion: boolean;
  tracking: boolean;
  canComment: boolean;
  showReview: boolean;
  toggleTracking: () => void;
  toggleReview: () => void;
  paste: () => void;
  pasteMode: PasteMode;
  setPasteMode: (mode: PasteMode) => void;
  changeIndent: (direction: number) => void;
  clearFormatting: () => void;
  zoom: number;
  formatText: (format: TextFormat) => void;
  setStyle: (styleId: string) => void;
  align: (alignment: Alignment) => void;
  toggleList: (kind: 'bullet' | 'number') => void;
  undo: () => void;
  redo: () => void;
  onNew: () => void;
  open: () => void;
  download: () => void;
  downloadDoc?: () => void;
  print: () => void;
  dialog: (dialog: DialogName) => void;
  image: () => void;
  pageBreak: () => void;
  toggleOutline: () => void;
  fit: () => void;
  setZoom: (zoom: number) => void;
  children?: ReactNode;
};

function Group({ label, children, secondary }: { label: string; children: ReactNode; secondary?: boolean }) {
  return <div className={'rwe-group' + (secondary ? ' rwe-secondary-group' : '')} role="group" aria-label={label}>{children}</div>;
}

export function Tool({ label, children, disabled, active, onClick, text }: {
  label: string; children: ReactNode; disabled?: boolean; active?: boolean; onClick: () => void; text?: boolean;
}) {
  return <button className={'rwe-tool' + (text ? ' rwe-tool-text' : '')} type="button" title={label} aria-label={label}
    aria-pressed={active} disabled={disabled} onMouseDown={event => event.preventDefault()} onClick={onClick}>{children}</button>;
}

export function Ribbon(props: RibbonProps) {
  const [overflow, setOverflow] = useState(false);
  const {
    tab, busy, disabled, format, styles, styleId, alignment, list, canUndo, canRedo, outline, canShowOutline, zoom,
    formatText, setStyle, align, toggleList, undo, redo, onNew, open, download, print, dialog, image, pageBreak,
    toggleOutline, fit, setZoom, pageRegion, inPageRegion, tracking, canComment, showReview, toggleTracking, toggleReview, paste, pasteMode, setPasteMode, changeIndent, clearFormatting,
  } = props;
  const fonts = FONTS;
  const sizes = [8, 9, 10, 11, 12, 14, 16, 17, 18, 20, 22, 24, 28, 32, 34, 36, 48, 60, 72];
  return <div className="rwe-ribbon" role="tabpanel" aria-label={tab + ' commands'}>
    {tab === 'File' && <Group label="Document">
      <Tool label="New document" onClick={onNew} disabled={busy} text><FilePlus2 /> New document</Tool>
      <Tool label="Document templates" onClick={() => dialog('templates')} disabled={busy || disabled} text><Files /> Templates</Tool>
      <Tool label="Open DOCX" onClick={open} disabled={busy} text><FolderOpen /> Open</Tool>
      <Tool label="Download DOCX" onClick={download} disabled={busy} text><Download /> Download DOCX</Tool>
      {props.downloadDoc && <Tool label="Download DOC" onClick={props.downloadDoc} disabled={busy} text><Download /> Download DOC</Tool>}
      <Tool label="Print document" onClick={print} disabled={busy} text><Printer /> Print</Tool>
    </Group>}
    {tab === 'Home' && <>
      <Group label="Clipboard">
        <Tool label="Paste from clipboard" onClick={paste} disabled={disabled}><ClipboardPaste /></Tool>
        <label className="rwe-select rwe-paste-select" title="Paste formatting">
          <select aria-label="Paste formatting" value={pasteMode} onChange={event => setPasteMode(event.target.value as PasteMode)} disabled={disabled}>
            <option value="source">Keep source formatting</option>
            <option value="merge">Merge formatting</option>
            <option value="text">Keep text only</option>
          </select><ChevronDown aria-hidden="true" />
        </label>
      </Group>
      <Group label="History">
        <Tool label="Undo" onClick={undo} disabled={disabled || !canUndo}><Undo2 /></Tool>
        <Tool label="Redo" onClick={redo} disabled={disabled || !canRedo}><Redo2 /></Tool>
      </Group>
      <Group label="Text formatting">
        <label className="rwe-select rwe-font-select">
          <select aria-label="Font family" disabled={disabled} value={format.font ?? 'Arial'} onChange={event => formatText({ font: event.target.value })}>
            {Array.from(new Set([...(format.font ? [format.font] : []), ...fonts])).map(font => <option key={font}>{font}</option>)}
          </select><ChevronDown aria-hidden="true" />
        </label>
        <label className="rwe-select rwe-size-select">
          <select aria-label="Font size" disabled={disabled} value={format.size ?? 11} onChange={event => formatText({ size: Number(event.target.value) })}>
            {Array.from(new Set([...(format.size ? [format.size] : []), ...sizes])).sort((a, b) => a - b).map(size => <option key={size}>{size}</option>)}
          </select><ChevronDown aria-hidden="true" />
        </label>
        <Tool label="Bold" active={!!format.bold} disabled={disabled} onClick={() => formatText({ bold: !format.bold })}><Bold /></Tool>
        <Tool label="Italic" active={!!format.italic} disabled={disabled} onClick={() => formatText({ italic: !format.italic })}><Italic /></Tool>
        <Tool label="Underline" active={!!format.underline} disabled={disabled} onClick={() => formatText({ underline: !format.underline })}><Underline /></Tool>
        <TextColor value={format.color} disabled={disabled} onChange={color => formatText({ color })} />
      </Group>
      <div className={'rwe-ribbon-secondary' + (overflow ? ' rwe-show-more' : '')}>
        <Group label="More text formatting" secondary>
          <Tool label="Strikethrough" active={!!format.strike} disabled={disabled} onClick={() => formatText({ strike: !format.strike })}><Strikethrough /></Tool>
          <Tool label="Superscript" active={format.vertical === 'superscript'} disabled={disabled} onClick={() => formatText({ vertical: format.vertical === 'superscript' ? 'baseline' : 'superscript' })}><Superscript /></Tool>
          <Tool label="Subscript" active={format.vertical === 'subscript'} disabled={disabled} onClick={() => formatText({ vertical: format.vertical === 'subscript' ? 'baseline' : 'subscript' })}><Subscript /></Tool>
          <Tool label="Text highlight" active={!!format.highlight} disabled={disabled} onClick={() => formatText({ highlight: format.highlight ? undefined : '#fff19c' })}><Highlighter /></Tool>
          <Tool label="Clear formatting" disabled={disabled} onClick={clearFormatting}><Eraser /></Tool>
        </Group>
        <Group label="Paragraph" secondary>
          <Tool label="Bulleted list" active={list === 'bullet'} disabled={disabled} onClick={() => toggleList('bullet')}><List /></Tool>
          <Tool label="Numbered list" active={list === 'number'} disabled={disabled} onClick={() => toggleList('number')}><ListOrdered /></Tool>
          <span className="rwe-group-divider" />
          {([['left', AlignLeft], ['center', AlignCenter], ['right', AlignRight], ['justify', AlignJustify]] as const)
            .map(([value, Icon]) => <Tool key={value} label={'Align ' + value} active={(alignment ?? 'left') === value}
              disabled={disabled} onClick={() => align(value)}><Icon /></Tool>)}
          <Tool label="Decrease indent" disabled={disabled} onClick={() => changeIndent(-1)}><IndentDecrease /></Tool>
          <Tool label="Increase indent" disabled={disabled} onClick={() => changeIndent(1)}><IndentIncrease /></Tool>
        </Group>
        <Group label="Styles" secondary>
          <label className="rwe-select rwe-style-select">
            <select aria-label="Paragraph style" disabled={disabled} value={styleId ?? ''} onChange={event => setStyle(event.target.value)}>
              {styles.map(style => <option key={style.id} value={style.id}>{style.name}</option>)}
            </select><ChevronDown aria-hidden="true" />
          </label>
          <Tool label="Modify style" disabled={disabled} onClick={() => dialog('modify-style')}><SlidersHorizontal /></Tool>
          <Tool label="New style" disabled={disabled} onClick={() => dialog('new-style')}><Heading /></Tool>
        </Group>
        <Group label="Editing" secondary>
          <Tool label="Find in document" onClick={() => dialog('find')}><Search /></Tool>
          <Tool label="Find and replace" disabled={disabled} onClick={() => dialog('replace')}><Replace /></Tool>
        </Group>
      </div>
      <button className="rwe-tool rwe-overflow" type="button" aria-label="More formatting" aria-expanded={overflow}
        onMouseDown={event => event.preventDefault()} onClick={() => setOverflow(value => !value)}>···</button>
    </>}
    {tab === 'Insert' && <Group label="Insert">
      <Tool label="Insert link" disabled={disabled} onClick={() => dialog('link')} text><Link2 /> Link</Tool>
      <Tool label="Insert picture" disabled={disabled} onClick={image} text><ImagePlus /> Picture</Tool>
      <Tool label="Insert table" disabled={disabled} onClick={() => dialog('table')} text><Table2 /> Table</Tool>
      <Tool label="Insert page break" disabled={disabled || inPageRegion} onClick={pageBreak} text><SeparatorHorizontal /> Page break</Tool>
      <Tool label="Insert template field" disabled={disabled || tracking} onClick={() => dialog('template-field')} text><Braces /> Field</Tool>
      <span className="rwe-group-divider" />
      <Tool label="Edit header" disabled={disabled} onClick={() => pageRegion('header')} text><PanelTop /> Header</Tool>
      <Tool label="Edit footer" disabled={disabled} onClick={() => pageRegion('footer')} text><PanelBottom /> Footer</Tool>
    </Group>}
    {tab === 'Layout' && <Group label="Page setup">
      <Tool label="Page setup" disabled={disabled} onClick={() => dialog('page-setup')} text><ALargeSmall /> Page setup</Tool>
      <Tool label="Insert page break" disabled={disabled || inPageRegion} onClick={pageBreak} text><SeparatorHorizontal /> Page break</Tool>
      <Tool label="Edit header" disabled={disabled} onClick={() => pageRegion('header')} text><PanelTop /> Header</Tool>
      <Tool label="Edit footer" disabled={disabled} onClick={() => pageRegion('footer')} text><PanelBottom /> Footer</Tool>
    </Group>}
    {tab === 'Review' && <Group label="Review">
      <Tool label="New comment" disabled={disabled || !canComment} onClick={() => dialog('comment')} text><MessageSquarePlus /> New comment</Tool>
      <Tool label="Track changes" disabled={disabled} active={tracking} onClick={toggleTracking} text><FilePenLine /> Track changes</Tool>
      <Tool label="Review pane" active={showReview} onClick={toggleReview} text><MessagesSquare /> Review pane</Tool>
    </Group>}
    {tab === 'View' && <Group label="View">
      <Tool label="Keyboard shortcuts" onClick={() => dialog('shortcuts')} text><Keyboard /> Shortcuts</Tool>
      <Tool label="Toggle outline" disabled={!canShowOutline} onClick={toggleOutline} active={outline} text><PanelLeft /> Outline</Tool>
      <Tool label="Fit to width" onClick={fit} text><Scan /> Fit to width</Tool>
      <Tool label="Actual size" onClick={() => setZoom(100)} active={zoom === 100} text><Columns3 /> Actual size</Tool>
      <Tool label="Go to page" onClick={() => dialog('goto')} text>Go to page</Tool>
    </Group>}
    {tab === 'Format Table' && props.children}
  </div>;
}
