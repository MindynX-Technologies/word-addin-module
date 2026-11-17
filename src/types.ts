export type DocxEditorErrorCode =
  | 'INVALID_FILE' | 'FILE_TOO_LARGE' | 'UNSUPPORTED_FILE' | 'UNSAVED_CHANGES'
  | 'OPEN_SUPERSEDED' | 'OPEN_FAILED' | 'EXPORT_FAILED' | 'INVALID_IMAGE'
  | 'INVALID_LINK' | 'PROTECTED_CONTENT' | 'TEMPLATE_FIELDS' | 'TEMPLATE_UNSUPPORTED'
  | 'CONVERSION_FAILED' | 'RECOVERY_FAILED' | 'CLIPBOARD_FAILED';

export class DocxEditorError extends Error {
  readonly code: DocxEditorErrorCode;

  constructor(code: DocxEditorErrorCode, message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = 'DocxEditorError';
    this.code = code;
  }
}

export interface DocxEditorState {
  fileName: string;
  dirty: boolean;
  readOnly: boolean;
}

export interface DocxEditorProps {
  initialDocument?: File;
  readOnly?: boolean;
  className?: string;
  onChange?: (state: DocxEditorState) => void;
  onError?: (error: DocxEditorError) => void;
  convertDocument?: DocumentConverter;
  /** Stable host-document identifier. Enables local recovery without mixing different documents. */
  recoveryKey?: string;
  reviewAuthor?: string;
}

export type DocumentFormat = 'doc' | 'docx';
export type DocumentConverter = (file: File, format: DocumentFormat) => Promise<File>;

export interface DocxEditorHandle {
  open(file: File, options?: { discardChanges?: boolean }): Promise<void>;
  exportDocx(): Promise<File>;
  exportDoc(): Promise<File>;
  flushRecovery(): Promise<void>;
}
