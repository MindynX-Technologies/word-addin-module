# React Word Editor

A React 18 and 19 component for opening, editing, and downloading Word documents. DOCX editing runs in the browser; legacy DOC import/export uses an optional host-controlled converter. The package includes a Word-style ribbon, a paginated canvas, and its own WordprocessingML reader, editor, and writer. ZIP compression is handled by zip.js.

## Try the demo

```sh
npm ci
npm run dev
```

Open the URL printed by Vite. The demo starts with an editable sample document. Add `?blank` to start with a blank document.

The demo enables local draft recovery. To open and download legacy `.doc` files as well, start its local conversion companion:

```sh
docker compose up --build -d word-converter
npm run dev
```

The companion listens on `127.0.0.1:4175`; Vite proxies `/word-converter` to it. It uses LibreOffice to perform actual DOC ↔ DOCX conversion. Stop it with `docker compose stop word-converter`. Conversion sends the chosen document to this local service; it does not use a third-party conversion website. DOCX editing requires no conversion service.

## Embed the editor

Install the tarball produced by `npm pack`.

```sh
npm install /path/to/react-word-editor-0.4.2.tgz
```

```tsx
import { useRef, useState } from 'react';
import { DocxEditor } from 'react-word-editor';
import type { DocxEditorHandle, DocxEditorState } from 'react-word-editor';
import 'react-word-editor/style.css';

export function DocumentPanel({ initialFile }: { initialFile?: File }) {
  const editor = useRef<DocxEditorHandle>(null);
  const [state, setState] = useState<DocxEditorState>();
  const [error, setError] = useState('');

  async function keepCopy() {
    try {
      const file = await editor.current?.exportDocx();
      if (file) console.info('Exported', file.name, file.size, 'bytes');
    } catch {
      // The editor calls onError and shows the failure in the ribbon.
    }
  }

  return (
    <section>
      <button type="button" onClick={keepCopy}>Export to host</button>
      {state?.dirty && <span>Unsaved changes</span>}
      {error && <p role="alert">{error}</p>}
      <div style={{ height: '75vh' }}>
        <DocxEditor
          ref={editor}
          initialDocument={initialFile}
          onChange={setState}
          onError={problem => setError(problem.message)}
        />
      </div>
    </section>
  );
}
```

Give the component a parent with a definite height. The stylesheet is scoped to the editor. The package exports ESM and TypeScript declarations. In frameworks with server components, mount the editor in a client component.

### API

| Interface | Behavior |
| --- | --- |
| `initialDocument?: File` | Opens the first supplied file. Without a file the editor starts blank. Use the ref to replace a document after mounting. |
| `readOnly?: boolean` | Disables document mutations. Opening and downloading files remain available. |
| `className?: string` | Adds a class to the editor root. |
| `onChange?: (state: DocxEditorState) => void` | Reports changes to `fileName`, `dirty`, and `readOnly`. |
| `onError?: (error: DocxEditorError) => void` | Reports operation failures with a typed `code` and a display message. |
| `convertDocument?: DocumentConverter` | Enables DOC import and DOC export through a host-controlled `(file, targetFormat) => Promise<File>` converter. Without it the component remains DOCX-only. |
| `recoveryKey?: string` | Enables IndexedDB recovery for this stable host-document key. A saved draft is offered for explicit restore/discard on the next mount. Use different keys for different documents/editors. |
| `reviewAuthor?: string` | Author for new tracked changes and the initial author in the comment form. Defaults to `Editor`. |
| `open(file, { discardChanges? }): Promise<void>` | Replaces the document after a successful import. Rejects with `UNSAVED_CHANGES` if the document is dirty and discard was not authorized. A newer open supersedes a pending one. |
| `exportDocx(): Promise<File>` | Exports a snapshot for the host. It leaves the current dirty state intact. The built-in Download command clears the dirty state if the exported snapshot is still current. |
| `exportDoc(): Promise<File>` | Exports legacy DOC using the configured converter. An untouched imported DOC returns its original bytes. It leaves dirty state intact. File → Download DOC saves through the UI. |
| `flushRecovery(): Promise<void>` | Flushes pending edits to the local recovery store when `recoveryKey` is set. Rejects if persistence fails. |

Configure conversion and recovery in a host:

```tsx
import { createDocumentConverter, DocxEditor } from 'react-word-editor';

const convertDocument = createDocumentConverter('/word-converter');
// Proxy this path to the companion in your server, with your app's authentication.
<DocxEditor convertDocument={convertDocument} recoveryKey="document-123" reviewAuthor="Alex Reviewer" />;
```

The supplied converter accepts binary DOC and standard DOCX with a 32 MiB input/output limit, bounded ZIP expansion, two concurrent conversions, a 45-second conversion timeout, temporary per-request profiles, disabled macros/link updates, and automatic temporary-file cleanup. The container runs as a non-root user with a read-only filesystem, temporary storage, and CPU/memory/process limits. It is bound to loopback for local use. A deployed host must place it behind the host's authenticated same-origin proxy; the Vite proxy is a development configuration, not a production endpoint. Conversion can change formatting that the older DOC format cannot represent.

Recovery saves a DOCX snapshot 750 ms after editing pauses and also attempts a flush when the page becomes hidden or the editor unmounts. Successful downloads and undoing all changes remove the recovery copy. Existing recovery copies are not silently restored or overwritten. Recovery is scoped to the current browser/site and key; clearing site data removes it. The browser can terminate a page before its last asynchronous write completes, so a small amount of the most recent typing can still be lost. Hosts can await `flushRecovery()` before unmounting or navigating. Storage failures are shown and reported as `RECOVERY_FAILED`.

## Editing and review

- **Home** includes paste, undo and redo, fonts, color, highlighting, strikethrough, superscript, subscript, alignment, lists, indentation, find, and a paragraph style gallery. Use **Modify style** to update every paragraph using that style, or **New style** to create one. Styles can set font, size, emphasis, color, alignment, spacing, indentation, and any of nine heading levels.
- **Insert** adds links, pictures, tables, page breaks, headers, and footers. Double-click an existing header or footer to edit it. The header and footer controls support page numbers, total pages, and different content for the first, odd, and even pages.
- **Layout** controls page width, height, margins, and header and footer distances. Selecting a table cell shows row, column, and shading controls. Use **Select table** or the handle at the table’s upper-left corner, then press **Delete** or **Backspace** to remove the entire table. **Delete table** is also available in Table tools. Selecting an inline picture shows size and alternative text controls.
- **Review** supports single-paragraph comments and tracked text insertions and deletions, with accept and reject controls. The DOCX carries comments and changes for file-based review. Tracking mode applies to text edits within a paragraph.

## Table formatting

Click a cell to show the contextual **Format Table** tab immediately to the right of **View**. Its ribbon contains table selection, row/column actions, deletion, and cell shading. Choose **Format table** there to set the table width and alignment, choose an outside, horizontal, full-grid, or borderless style, and adjust border color and weight. Set the current column’s width, distribute columns evenly, choose the current row’s minimum height, and adjust table-wide cell padding.

Border widths accept decimal points and are saved to the nearest eighth of a point. Applying table borders replaces imported cell and row border overrides along the grid. Thin double borders use extra display width on the canvas so both strokes are visible.

Apply shading and vertical alignment to the current cell, row, column, or whole table. Enable **Repeat header row on each page** for long tables. Formatting changes export as native DOCX properties and each **Apply** is a single undo step. Horizontally merged cells support text editing and shading, preserve their native grid spans, and protect row/column geometry changes. Vertically merged and nested DOCX tables remain protected.

## Word keyboard shortcuts

Open **View → Shortcuts** for the searchable, platform-specific list. It covers editing and formatting, paragraph styles, clipboard actions, document and table navigation, comments, Track Changes, Find and Replace, page navigation, and printing. In the editor, press **Alt+.** on Windows/Linux or **Ctrl+.** on Mac to show ribbon KeyTips. Use **Ctrl+F6** to move between editor regions.

Some browser and operating system combinations reserve keys such as Ctrl+N or Cmd+T before they reach the page. Ribbon KeyTips provide keyboard access to file and formatting commands in those cases. Desktop-only Word commands require the corresponding editor feature. [Microsoft's Word shortcut reference](https://support.microsoft.com/en-us/accessibility/word/keyboard-shortcuts-in-word)

## Document templates

1. Design a document in the editor, then use **Insert → Field** to add placeholders such as `{{client_name}}`. You can also type the placeholders directly.
2. Open **File → Templates** to save the current document or import an existing `.docx` or unencrypted `.dotx`.
3. Select the template, fill its fields, and choose **Create document**. Repeated fields take the same value. The resulting DOCX is independent of the template and retains its layout, styles, tables, and headers and footers.

The template library is stored in IndexedDB for the current browser and site. Fields support text in editable body paragraphs, table cells, and headers and footers. Field names begin with a letter and can include letters, numbers, underscores, periods, and hyphens. Matching is case-sensitive. Formatting for inserted text follows the start of the placeholder. A value can be an empty string. Fields in locked or reviewed content are excluded. Conditional sections, repeating rows, and Word mail merge fields require separate support.

Host apps can also read the fields and generate a file directly. The functions run in the browser and accept a Word `File`.

```ts
import { createDocumentFromTemplate, readTemplateFields } from 'react-word-editor';

const fields = await readTemplateFields(templateFile);
const file = await createDocumentFromTemplate(templateFile, { client_name: 'North & West' }, 'Client brief');
```

## Pasting from Word

Paste with Ctrl+V or Command+V. **Keep source formatting** is the default. The clipboard menu also offers **Merge formatting** and **Keep text only**. Browser support for the ribbon's Paste button depends on clipboard permissions.

When Word supplies HTML or image data through the browser clipboard, the editor imports supported heading levels, inline formatting, links, paragraph spacing, decimal/letter/Roman lists, tables, and PNG/JPEG pictures. Table widths, cell widths, borders, padding, shading, and vertical alignment are retained where represented by supported CSS. Horizontal merges remain merged; vertical merges and nested clipboard tables are simplified with a notice. The editor falls back to plain text when formatted content cannot be inserted at the current selection. Images must be embedded in the HTML or provided as image data. The browser never fetches URLs from pasted markup.

Source formatting includes supported text properties inherited from HTML wrappers and simple tag/class stylesheet rules. **Merge formatting** uses the destination font, size, and color while retaining pasted emphasis. A single normal paragraph pasted into existing text is inserted into that sentence; its paragraph formatting does not split the sentence. Headings, lists, and multiple paragraphs retain their block structure.

Rich paste accepts tables up to 200 rows and 20 columns and applies HTML size, stylesheet, node, and nesting limits. Exceeding a limit rejects the entire rich fragment and uses the clipboard's plain text when available, with a visible notice. It does not insert a truncated table. Notices also identify flattened merged cells, nested tables, and unavailable pictures. If no usable fallback exists, the editor leaves the document unchanged and explains the failure.

Copy and Cut put formatted HTML and a plain-text fallback on the clipboard. Table cells use tabs and rows use newlines in the plain-text representation. Pictures are embedded into clipboard HTML at their selected positions. Picture-only selections also supply a native PNG representation for applications such as Word. Cuts containing pictures remove the selection only after a successful clipboard write and only if the document and selection are still current. Permission/encoding failures keep the document unchanged and report `CLIPBOARD_FAILED`. Receiving applications can sanitize HTML data-image URLs; DOCX download remains the reliable fallback for mixed text/picture transfers into applications that discard them. Choose **Keep Source Formatting** in Word. Word for the web can apply its own font and size to hyperlinks.

## DOCX compatibility

- The editor supports single-section documents with everyday text, lists, tables, and inline images. Page numbers and total pages use native Word fields. Pagination can differ from Word because of font substitution and browser line wrapping. Oversized table rows stay together and can extend the page.
- Unsupported blocks such as floating objects, advanced fields, and complex tables are protected. Multiple sections or columns, structural revisions, modern threaded comments, editing restrictions, and digital signatures open read-only. Shapes and live co-authoring are outside this release.
- Imports support standard Transitional WordprocessingML `.docx` packages up to 32 MiB compressed, 128 MiB expanded, 48 MiB per part, and 4,096 archive entries. Legacy `.doc` files are converted when `convertDocument` is configured. Password-protected, macro-enabled, and Strict WordprocessingML files require conversion to a standard unencrypted `.docx` first.
- Unchanged imports export with the original bytes. Editing changes the affected WordprocessingML and preserves unrelated package parts. Exact/minimum line spacing, paragraph direction, default tab intervals, keep-with-next/keep-lines, and widow/orphan pagination have browser display support; unrelated imported paragraph XML attributes remain preserved. DOCX stays in the browser except when the host explicitly invokes its converter. Drafts persist locally when `recoveryKey` is configured. A browser navigation guard is registered while changes are unsaved.
- Desktop editing targets Chrome, Edge, Firefox, and Safari. The test configuration supports Chromium, Firefox, and WebKit. OS clipboard permission automation runs in Chromium; other engines exercise editing/import/export and synthetic clipboard events. Touch phones use fit-to-width reading mode.

### Word editing assessment and remaining gaps

This editor supports browser-local DOCX editing and conversion-backed legacy DOC workflows, with safeguards for unsupported content. It does not provide complete Microsoft Word fidelity.

| Area | Current behavior and practical consequence |
| --- | --- |
| Legacy `.doc` / `.dot` | Binary DOC import/export is available through the supplied converter and enabled in the demo. A host must configure `convertDocument` and run its service. Legacy DOT templates are not supported. Renaming extensions is not conversion. |
| Complex DOCX | Multiple sections/columns, structural revisions, restrictions, signatures, and modern comments can force read-only mode. Complex tables, floating objects, bookmarks, footnote references, equations, and unsupported fields can protect affected content or appear as placeholders. Preservation in the package does not mean faithful display or editability. |
| Clipboard fidelity | Clipboard HTML does not carry a complete DOCX package. Comments, tracked-change metadata, fields, section layout, internal bookmark links, and some custom numbering are not transferred faithfully. Horizontal merges and supported table CSS are retained; vertical merges and nested tables are simplified with a notice. |
| Pictures | Embedded PNG/JPEG import, HTML copy/cut, and picture-only PNG clipboard output are supported. File/remote URLs, VML/EMF/WMF drawings, and floating objects remain unsupported. Some receiving Word versions discard data-image HTML in mixed content; DOCX export avoids this restriction. |
| Layout and printing | Exact/minimum line spacing, paragraph direction, keep rules, default tab intervals, and widow/orphan handling are supported. Font substitution, custom per-paragraph tab stops, complex-script shaping differences, table styles, section layouts, and oversized rows can still differ from Word. |
| Review | Review authors are configurable. Comments span one paragraph; tracked text editing is limited to one paragraph. Structural changes and richer review workflows remain protected and need Word. |
| Browser clipboard | Programmatic Paste requires a secure context and browser permission/user interaction. Embedded hosts may also need clipboard permissions in their iframe policy. Keyboard paste is the fallback; receiving Word applications may reinterpret the supplied HTML. |
| Unsaved work | Per-document local recovery is available through `recoveryKey` and enabled in the demo. Writes are asynchronous and can be interrupted by sudden termination. Hosts should await `flushRecovery()` for controlled navigation and retain downloaded or server-saved copies. |
| Scale | Word counts are cached per unchanged paragraph and unchanged paragraph fragments avoid rerendering. Table measurements are reused. Very large documents still require main-thread model construction/pagination; page virtualization and worker-based editing are not implemented. |

The automated tests cover undo/redo, archive XML assertions, export/reimport, recovery, and browser clipboard handling. Optional integration tests exercise real LibreOffice DOC conversion and reopen edited DOC exports. Firefox/WebKit runs test their browser engines; they do not certify the native Windows/macOS Word clipboard or the Safari application. Generated documents have not been opened in Microsoft Word in this environment.

Before a release that promises Word interoperability, validate a representative document corpus in Word for Windows, Word for Mac, and Word for the web. For each target browser, check both copy directions with partial sentences, nested/restarted lists, shaded/merged tables, inline/floating pictures, multilingual text, comments, and tracked changes. Export edited documents, reopen them in Word, check for repair prompts, and compare text, formatting, links, page layout, and review metadata. Include converted legacy DOC files, rejected encrypted/Strict files, large documents, and recovery after interrupted work.

Format and platform references: [Microsoft Word file formats](https://learn.microsoft.com/en-us/office/compatibility/office-file-format-reference), [legacy DOC specification](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/), [browser Clipboard API requirements](https://developer.mozilla.org/en-US/docs/Web/API/Clipboard_API).

## Development

```sh
npm run typecheck
npx playwright install chromium
npm test
PLAYWRIGHT_ALL_BROWSERS=1 npm test
WORD_CONVERTER_TESTS=1 npm test
npm run build:demo
npm pack
```

Set `PLAYWRIGHT_CHANNEL=chrome` to run the tests with an installed Google Chrome. The package is built during `npm pack`. The demo builds to `demo-dist/`.

Run `npm run test:precommit` to type-check the project and run the full Chromium suite locally. In a Git checkout, run `git config --local core.hooksPath .githooks` once to run the same checks before each commit. A failing check blocks the commit.

Install Firefox/WebKit with `npx playwright install firefox webkit` before the browser matrix. Playwright may also require its documented OS dependencies. Start the conversion companion before setting `WORD_CONVERTER_TESTS=1`.
