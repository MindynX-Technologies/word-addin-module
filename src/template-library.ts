import { loadFile, storeFile } from './file-storage';
import type { StoredFile } from './file-storage';

export interface StoredTemplate {
  id: string;
  name: string;
  file: File;
  updated: number;
}

function openLibrary() {
  return new Promise<IDBDatabase>((resolve, reject) => {
    const request = indexedDB.open('react-word-editor-templates', 1);
    request.onupgradeneeded = () => request.result.createObjectStore('templates', { keyPath: 'id' });
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function operation<T>(mode: IDBTransactionMode, use: (store: IDBObjectStore) => IDBRequest<T>) {
  const db = await openLibrary();
  try {
    return await new Promise<T>((resolve, reject) => {
      const transaction = db.transaction('templates', mode);
      const request = use(transaction.objectStore('templates'));
      transaction.oncomplete = () => resolve(request.result);
      transaction.onabort = () => reject(transaction.error ?? request.error);
      transaction.onerror = () => reject(transaction.error ?? request.error);
    });
  } finally {
    db.close();
  }
}

export async function listTemplates() {
  const records = await operation('readonly', store => store.getAll());
  return (records as (Omit<StoredTemplate, 'file'> & { file: File | StoredFile })[])
    .map(record => ({ ...record, file: loadFile(record.file) })).sort((a, b) => b.updated - a.updated);
}

export async function saveTemplate(name: string, file: File) {
  const template: StoredTemplate = { id: globalThis.crypto.randomUUID(), name, file, updated: Date.now() };
  const fileData = await storeFile(file);
  await operation('readwrite', store => store.put({ ...template, file: fileData }));
  return template;
}

export function removeTemplate(id: string) {
  return operation('readwrite', store => store.delete(id));
}
