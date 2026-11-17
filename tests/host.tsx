import { createRoot } from 'react-dom/client';
import { createDocumentConverter, DocxEditor } from '../src';
import '../src/style.css';
import type { DocxEditorHandle, DocxEditorState } from '../src';

declare global {
  interface Window {
    editor: DocxEditorHandle | null;
    editorState: DocxEditorState;
    editorErrors: string[];
  }
}

window.editorErrors = [];
const params = new URLSearchParams(location.search);
createRoot(document.getElementById('root')!).render(<DocxEditor
  ref={value => { window.editor = value; }}
  readOnly={new URLSearchParams(location.search).has('readonly')}
  recoveryKey={params.get('recovery') ?? undefined}
  reviewAuthor={params.get('author') ?? undefined}
  convertDocument={params.has('convert') ? createDocumentConverter('/word-converter') : undefined}
  onChange={state => { window.editorState = state; }}
  onError={error => { window.editorErrors.push(error.code); }}
/>);
