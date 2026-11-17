import { createRoot } from 'react-dom/client';
import { useEffect, useRef } from 'react';
import { createDocumentConverter, DocxEditor } from '../src';
import '../src/style.css';
import type { DocxEditorHandle } from '../src';
import sample from './Editorial Plan.docx?url';
import './demo.css';
const convertDocument = createDocumentConverter('/word-converter');

function Demo() {
  const editor = useRef<DocxEditorHandle>(null);

  useEffect(() => {
    if (new URLSearchParams(location.search).has('blank')) return;
    const abort = new AbortController();
    async function openSample() {
      const response = await fetch(sample, { signal: abort.signal });
      if (!response.ok || abort.signal.aborted) return;
      const blob = await response.blob();
      if (abort.signal.aborted) return;
      await editor.current?.open(new File([blob], 'Editorial Plan.docx', { type: blob.type }));
    }
    void openSample().catch(() => {});
    return () => abort.abort();
  }, []);

  return <main className="demo"><DocxEditor ref={editor} convertDocument={convertDocument} recoveryKey="demo-document" /></main>;
}

createRoot(document.getElementById('root')!).render(<Demo />);
