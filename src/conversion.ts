import { DocxEditorError } from './types';
import type { DocumentConverter } from './types';
import { MIME } from './ooxml';

/** Connects to a host-controlled converter. DOCX editing itself stays browser-local. */
export function createDocumentConverter(endpoint: string): DocumentConverter {
  return async (file, format) => {
    const input = /\.docx$/i.test(file.name) ? 'docx' : /\.doc$/i.test(file.name) ? 'doc' : undefined;
    if (!input) throw new DocxEditorError('UNSUPPORTED_FILE', 'Choose a DOC or DOCX document.');
    if (file.size > 32 * 1024 * 1024) throw new DocxEditorError('FILE_TOO_LARGE', 'Choose a document smaller than 32 MiB.');
    if (input === format) return file;
    const url = new URL(endpoint.replace(/\/$/, '') + '/convert', document.baseURI);
    url.searchParams.set('input', input);
    url.searchParams.set('format', format);
    const abort = new AbortController();
    const timer = setTimeout(() => abort.abort(), 60_000);
    try {
      const response = await fetch(url, { method: 'POST', body: file, signal: abort.signal,
        headers: { 'Content-Type': input === 'doc' ? 'application/msword' : MIME }, credentials: 'same-origin' });
      if (!response.ok) {
        const problem = await response.json().catch(() => ({})) as { message?: string };
        throw new DocxEditorError('CONVERSION_FAILED', problem.message ?? 'Word conversion failed.');
      }
      const blob = await response.blob();
      if (!blob.size || blob.size > 32 * 1024 * 1024) throw new DocxEditorError('CONVERSION_FAILED', 'The converter returned an empty or oversized document.');
      const signature = new Uint8Array(await blob.slice(0, 8).arrayBuffer());
      const valid = format === 'doc' ? [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].every((value, index) => signature[index] === value)
        : signature[0] === 0x50 && signature[1] === 0x4b;
      if (!valid) throw new DocxEditorError('CONVERSION_FAILED', 'The converter did not return the requested Word format.');
      return new File([blob], file.name.replace(/\.docx?$/i, '.' + format), { type: format === 'doc' ? 'application/msword' : MIME });
    } catch (cause) {
      if (cause instanceof DocxEditorError) throw cause;
      throw new DocxEditorError('CONVERSION_FAILED', abort.signal.aborted ? 'Word conversion timed out.' : 'The Word conversion service is unavailable.', { cause });
    } finally { clearTimeout(timer); }
  };
}
