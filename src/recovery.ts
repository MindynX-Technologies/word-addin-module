import { useCallback, useEffect, useRef, useState } from 'react';
import { exportDocx } from './docx-write';
import type { DocxSource } from './docx-read';
import type { DocModel } from './model';
import { DocxEditorError } from './types';
import { loadFile, storeFile } from './file-storage';
import type { StoredFile } from './file-storage';

export interface RecoveryRecord { key: string; file: File; updated: number }
type Snapshot = { model: DocModel; saved: DocModel; fileName: string; source?: DocxSource };

async function draftOperation<T>(mode: IDBTransactionMode, operation: (store: IDBObjectStore) => IDBRequest<T>) {
  const db = await new Promise<IDBDatabase>((resolve, reject) => {
    const request = indexedDB.open('react-word-editor-recovery', 1);
    request.onupgradeneeded = () => request.result.createObjectStore('drafts', { keyPath: 'key' });
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
  try {
    return await new Promise<T>((resolve, reject) => {
      const transaction = db.transaction('drafts', mode);
      const request = operation(transaction.objectStore('drafts'));
      transaction.oncomplete = () => resolve(request.result);
      transaction.onabort = () => reject(transaction.error ?? request.error);
      transaction.onerror = () => reject(transaction.error ?? request.error);
    });
  } finally { db.close(); }
}

export function useDraftRecovery(key: string | undefined, snapshot: Snapshot, report: (error: DocxEditorError) => void) {
  const [offer, setOffer] = useState<RecoveryRecord>();
  const [loaded, setLoaded] = useState(false);
  const [status, setStatus] = useState<'idle' | 'saving' | 'saved' | 'failed'>('idle');
  const latest = useRef({ key, snapshot, offer, loaded, report });
  latest.current = { key, snapshot, offer, loaded, report };
  const queue = useRef<Promise<void>>(Promise.resolve());
  const epoch = useRef(0);
  const wasDirty = useRef(false);
  const mounted = useRef(true);
  const failed = useCallback((cause: unknown) => {
    const error = new DocxEditorError('RECOVERY_FAILED', 'The local recovery copy could not be saved. Download your document to keep it.', { cause });
    if (mounted.current) { setStatus('failed'); latest.current.report(error); }
    throw error;
  }, []);

  const flush = useCallback((): Promise<void> => {
    const { key, snapshot, offer, loaded } = latest.current;
    if (!key || !loaded || offer || snapshot.model === snapshot.saved) return queue.current;
    const version = epoch.current;
    if (mounted.current) setStatus('saving');
    const task = queue.current.catch(() => {}).then(async () => {
      if (version !== epoch.current || key !== latest.current.key) return;
      const file = await exportDocx(snapshot.model, snapshot.fileName, snapshot.source);
      if (version !== epoch.current || key !== latest.current.key) return;
      const data = await storeFile(file);
      if (version !== epoch.current) return;
      await draftOperation('readwrite', store => store.put({ key, file: data, updated: Date.now() }));
      if (mounted.current && latest.current.snapshot.model === snapshot.model) setStatus('saved');
    }).catch(failed);
    queue.current = task;
    return task;
  }, [failed]);

  const discard = useCallback(async () => {
    const key = latest.current.key;
    epoch.current++;
    const task = queue.current.catch(() => {}).then(async () => {
      if (key) await draftOperation('readwrite', store => store.delete(key));
      if (mounted.current && key === latest.current.key) { setOffer(undefined); setStatus('idle'); }
    }).catch(failed);
    queue.current = task;
    await task;
  }, [failed]);

  useEffect(() => {
    let cancelled = false;
    epoch.current++;
    wasDirty.current = false;
    setLoaded(false); setOffer(undefined); setStatus('idle');
    if (!key) return;
    void draftOperation('readonly', store => store.get(key)).then(record => {
      if (!cancelled) {
        const stored = record as (Omit<RecoveryRecord, 'file'> & { file: File | StoredFile }) | undefined;
        setOffer(stored ? { ...stored, file: loadFile(stored.file) } : undefined); setLoaded(true);
      }
    }).catch(cause => { if (!cancelled) { setLoaded(true); try { failed(cause); } catch { /* already reported */ } } });
    return () => { cancelled = true; };
  }, [key, failed]);

  useEffect(() => {
    if (!key || !loaded || offer) return;
    if (snapshot.model === snapshot.saved) {
      if (wasDirty.current) void discard().catch(() => {});
      wasDirty.current = false;
      return;
    }
    wasDirty.current = true;
    const timer = setTimeout(() => { void flush().catch(() => {}); }, 750);
    return () => clearTimeout(timer);
  }, [key, loaded, offer, snapshot.model, snapshot.saved, flush, discard]);

  useEffect(() => {
    mounted.current = true;
    const save = () => { void flush().catch(() => {}); };
    const visibility = () => { if (document.visibilityState === 'hidden') save(); };
    window.addEventListener('pagehide', save);
    document.addEventListener('visibilitychange', visibility);
    return () => {
      mounted.current = false; save();
      window.removeEventListener('pagehide', save);
      document.removeEventListener('visibilitychange', visibility);
    };
  }, [flush]);

  return { offer: offer?.key === key ? offer : undefined, status, flush, discard, restored: () => setOffer(undefined) };
}
